"""
Paso 1 del flujo de montos SSI: leer del libro de Google las dos listas de CUI y guardarlas en un archivo.

  recien_insertados = hoja PIM-diario, columna A desde A3 hacia abajo
  todos_cuis        = hoja Proyección26, columna A (título "CUI" en A2) desde A3,
                      solo las filas cuya Etapa (columna F, título "Etapa" en F2) NO es "PC"
                      (la Etapa vacía también entra)

Controles (aviso por correo):
  - Si un título no está donde debe (Proyección26 A2 "CUI" y F2 "Etapa"; PIM-diario C1 "CUI"
    y A2 vacía), se avisa y se DETIENE sin escribir nada: alguien pudo insertar o borrar filas
    o mover columnas, y entonces las listas no empezarían donde se cree.
  - Cada valor leído debe ser un número de 7 cifras. Los que no lo son se dejan fuera, se
    anotan en el archivo y se avisa por correo; el resto sigue normal.
  - Cualquier otro error (credenciales, hoja inexistente, sin conexión) también se avisa.

Lee todo en una sola consulta y después trabaja con la copia en memoria (no toca la hoja).

Paso 2: descarga del SSI del MEF (servicio traeDetInvSSI, el mismo de la pantalla del buscador)
  - todos_cuis: monto de inversión = COSTO TOTAL DE LA INVERSIÓN ACTUALIZADO (a+b+c+d)
                = MTO_F8_C + MTO_LAUDO + MTO_CartFza (misma fórmula que la página del SSI)
  - recien_insertados: además UEI (DES_UNIDAD_UEI) y tipo de inversión (TIPO_FORMATO)
  - Se consulta la unión de las dos listas, una vez por CUI (si un insertado no está en
    todos_cuis, igual se consulta y se avisa).
  - Monto vacío -> 0. CUI que el MEF no encuentra o que no responde -> monto null (no 0, para
    no pisar un dato bueno con un cero cuando el servicio falla) y aviso por correo.
  - Resultado en CARPETA_MONTOS, un JSON y un CSV por corrida con fecha y hora en el nombre.
  - Se usa curl.exe de Windows: el servidor del MEF no envía su certificado intermedio y
    Python lo rechaza; curl lo completa con el almacén de Windows sin desactivar la verificación.
  - El SSI bloquea a los servidores de Google: por eso esto corre en la PC y no en Apps Script.

Paso 3: escritura en PIM-diario (una lectura y una escritura, solo columnas O, P y Q)
  - Tabla de N a Q: una fila por cada CUI de la columna N (CUI_proyeccion26_noPC); cada fila se
    ubica por su CUI, no por posición.
  - O = monto de inversión de ese CUI. Si el SSI no dio monto (no existe o no respondió), se
    deja el valor que ya tenía la celda.
  - P (UEI) y Q (TIPO_INVERSION) solo en las filas de los recién insertados; las demás vacías.
  - Si el título de O, P o Q no coincide, esa columna no se escribe y las otras sí (con aviso).
    Si N1 no dice CUI_proyeccion26_noPC no se escribe nada (con aviso).
  - CUIS_SIN_AVISO (3000882 y 3999999): no generan aviso aunque el SSI no los encuentre.

Comparación con la corrida anterior (último monto_ssi_*.json de CARPETA_MONTOS)
  - Montos que cambiaron -> Excel variacion_montos_<fecha>.xlsx (CUI, nombre, monto anterior,
    monto actual, % de variación con signo; resaltadas las que pasan el umbral).
  - El correo alerta solo las variaciones mayores a UMBRAL_VARIACION_PCT (en más o en menos).
  - Monto que llega en 0 cuando antes tenía valor: se toma como error del SSI, se mantiene el
    anterior (hoja y archivo, campo monto_final) y se avisa SOLO la primera vez; si sigue llegando
    en 0 en las corridas siguientes, se sigue manteniendo el anterior sin volver a avisar (queda
    anotado en el campo observacion). Si el SSI no da dato, también se mantiene el anterior.

Todos los avisos de una corrida van en un solo correo.

Uso:
  python cuis_desde_sheets.py                 corrida normal
  python cuis_desde_sheets.py --probar-correo envía un correo de prueba y termina

Códigos de salida: 0 bien (aunque haya avisos enviados por correo), 2 detenido por títulos
o estructura, 3 error inesperado, 4 no se pudo enviar el correo de aviso, 5 el SSI no
respondió para ningún CUI.
"""

import csv
import json
import os
import smtplib
import subprocess
import sys
import time
import traceback
import unicodedata
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timedelta
from email.message import EmailMessage
from pathlib import Path

import gspread

# ---------------- configuración ----------------
ID_LIBRO = '1u4JHmEyIn6WusJWZLF6DQPDyhGxgRF5cEEp4z_oE8iw'
HOJA_PIM = 'PIM-diario'
HOJA_PROY = 'Proyección26'
CORREO_AVISO = 'nataliac1598@gmail.com'      # puedes poner varios separados por coma; el primero envía
CARPETA_BASE = Path(__file__).resolve().parent.parent           # ...\daily-new-inv
CREDENCIALES = CARPETA_BASE / 'credenciales' / 'credenciales.json'
CARPETA_SALIDA = CARPETA_BASE / 'nata_apuntes' / 'test_responses'
CARPETA_MONTOS = Path(r'C:\Users\HP\Documents\drives_exclusivos_descarga\monto_ssi')

# SSI del MEF
SSI_SERVICIO = 'https://ofi5.mef.gob.pe/invierteWS/Ssi/traeDetInvSSI'
SSI_HILOS = 12          # consultas simultáneas
SSI_INTENTOS = 3        # por CUI, con pausa creciente
SSI_TIEMPO_MAX = 40     # segundos por consulta
CUIS_SIN_AVISO = {3000882, 3999999}   # no existen en el SSI y ya se sabe: no avisar por correo

# Comparación con la corrida anterior
UMBRAL_VARIACION_PCT = 10   # el correo alerta solo variaciones mayores a este % (en aumento o disminución)

# Paso 3: escritura en PIM-diario. Tabla de N a Q, una fila por CUI de la columna N.
PIM_COL_LLAVE, PIM_TIT_LLAVE = 'N', 'CUI_proyeccion26_noPC'
PIM_COL_MONTO, PIM_TIT_MONTO = 'O', 'COSTO_TOTAL_INVERSION_ACTUALIZADO (a+b+c+d)'
PIM_COL_UEI, PIM_TIT_UEI = 'P', 'UEI'
PIM_COL_TIPO, PIM_TIT_TIPO = 'Q', 'TIPO_INVERSION'

# Dónde debe estar cada cosa (fila y columna de la hoja, empezando en 1)
PROY_FILA_TITULOS = 2
PROY_COL_CUI, PROY_TIT_CUI = 1, 'CUI'        # A2
PROY_COL_ETAPA, PROY_TIT_ETAPA = 6, 'Etapa'  # F2
PROY_FILA_INI = 3
PIM_CELDA_TIT_CUI, PIM_TIT_CUI = 'C1', 'CUI' # ancla: si alguien inserta una fila arriba, C1 deja de decir CUI
PIM_FILA_INI = 3                              # A3 hacia abajo: insertados (A1 = fecha, A2 vacía)
ETAPA_EXCLUIDA = 'PC'


class Detener(Exception):
    """Problema de estructura: se avisa y no se escribe nada."""


# ---------------- utilidades ----------------
def norm(t):
    """Texto sin espacios sobrantes, sin tildes y en mayúsculas, para comparar títulos."""
    t = unicodedata.normalize('NFKD', str(t if t is not None else '')).strip()
    return ''.join(c for c in t if not unicodedata.combining(c)).upper()


def col_letra(c):
    s = ''
    while c > 0:
        c, r = divmod(c - 1, 26)
        s = chr(65 + r) + s
    return s


def como_cui(v):
    """Devuelve el CUI como entero si v es un número de 7 cifras; None si no lo es."""
    if isinstance(v, bool):
        return None
    if isinstance(v, int):
        return v if 1_000_000 <= v <= 9_999_999 else None
    if isinstance(v, float):
        return int(v) if v.is_integer() and 1_000_000 <= v <= 9_999_999 else None
    s = str(v).strip()
    if s.endswith('.0'):
        s = s[:-2]
    return int(s) if s.isdigit() and len(s) == 7 else None


def vacio(v):
    return v is None or str(v).strip() == ''


def celda(fila, i):
    return fila[i] if i < len(fila) else ''


def fecha_serial(v):
    """Fecha y hora de la hoja (número de serie) a texto; si no es número, el texto tal cual."""
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return (datetime(1899, 12, 30) + timedelta(days=float(v))).strftime('%Y-%m-%d %H:%M:%S')
    return str(v)


def clave_gmail():
    clave = os.environ.get('GMAIL_APP_PASSWORD')
    if not clave and sys.platform == 'win32':
        # setx guarda la variable para las ventanas NUEVAS; si esta se abrió antes, se lee del registro
        import winreg
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, 'Environment') as k:
                clave = winreg.QueryValueEx(k, 'GMAIL_APP_PASSWORD')[0]
        except OSError:
            clave = None
    return (clave or '').replace(' ', '')


def enviar_correo(asunto, cuerpo):
    destinos = [c.strip() for c in CORREO_AVISO.split(',') if c.strip()]
    msg = EmailMessage()
    msg['Subject'] = '[CUI PIM-diario] ' + asunto
    msg['From'] = destinos[0]
    msg['To'] = ', '.join(destinos)
    msg.set_content(cuerpo)
    try:
        clave = clave_gmail()
        if not clave:
            raise RuntimeError('no está definida la variable GMAIL_APP_PASSWORD')
        with smtplib.SMTP_SSL('smtp.gmail.com', 465, timeout=60) as s:
            s.login(destinos[0], clave)
            s.send_message(msg)
        print('Correo enviado:', asunto)
        return True
    except Exception as e:
        # que el aviso no se pierda: se deja en un archivo junto a las salidas
        CARPETA_SALIDA.mkdir(parents=True, exist_ok=True)
        f = CARPETA_SALIDA / ('AVISO_NO_ENVIADO_' + datetime.now().strftime('%Y-%m-%d_%H%M%S') + '.txt')
        f.write_text(asunto + '\n\n' + cuerpo + '\n\nNo se pudo enviar por correo: ' + str(e), encoding='utf-8')
        print('NO se pudo enviar el correo (' + str(e) + '). Aviso guardado en', f)
        return False


def enlace(gid=None):
    return 'https://docs.google.com/spreadsheets/d/' + ID_LIBRO + '/edit' + (f'#gid={gid}' if gid else '')


# ---------------- lectura ----------------
def buscar_hoja(libro, nombre):
    """Busca la pestaña por nombre exacto; si no, ignorando tildes y mayúsculas (Proyeccion26 = Proyección26)."""
    hojas = libro.worksheets()
    for w in hojas:
        if w.title == nombre:
            return w
    for w in hojas:
        if norm(w.title) == norm(nombre):
            return w
    raise Detener(f'No existe la hoja "{nombre}" en el libro "{libro.title}".\n'
                  f'Hojas parecidas: {", ".join(w.title for w in hojas if norm(nombre)[:6] in norm(w.title)) or "ninguna"}.\n'
                  f'Libro: {enlace()}')


def leer():
    gc = gspread.service_account(filename=str(CREDENCIALES))
    libro = gc.open_by_key(ID_LIBRO)
    hp, hr = buscar_hoja(libro, HOJA_PIM), buscar_hoja(libro, HOJA_PROY)
    ult_proy = col_letra(PROY_COL_ETAPA)
    rangos = [
        f"'{hp.title}'!A1:C2",                                   # A1 fecha, A2 vacía, C1 título CUI
        f"'{hp.title}'!A{PIM_FILA_INI}:A",                        # insertados
        f"'{hr.title}'!A{PROY_FILA_TITULOS}:{ult_proy}{PROY_FILA_TITULOS}",  # títulos de Proyección26
        f"'{hr.title}'!A{PROY_FILA_INI}:{ult_proy}",              # datos de Proyección26
    ]
    r = libro.values_batch_get(rangos, params={'valueRenderOption': 'UNFORMATTED_VALUE'})
    vr = [x.get('values', []) for x in r['valueRanges']]
    return libro, hp, hr, vr


# ---------------- validación ----------------
def validar_y_armar(hp, hr, vr):
    cab_pim, col_a_pim, tit_proy, datos_proy = vr
    errores = []

    # Proyección26: títulos en la fila 2
    tits = tit_proy[0] if tit_proy else []
    for col, esperado in ((PROY_COL_CUI, PROY_TIT_CUI), (PROY_COL_ETAPA, PROY_TIT_ETAPA)):
        hallado = celda(tits, col - 1)
        if norm(hallado) != norm(esperado):
            errores.append(f'Hoja "{hr.title}", celda {col_letra(col)}{PROY_FILA_TITULOS}: dice '
                           f'"{hallado if not vacio(hallado) else "(vacía)"}" y debería decir "{esperado}".')

    # PIM-diario: C1 = "CUI" (fila de títulos de la descarga) y A2 vacía (la lista empieza en A3)
    fila1 = cab_pim[0] if cab_pim else []
    fila2 = cab_pim[1] if len(cab_pim) > 1 else []
    c1 = celda(fila1, 2)
    if norm(c1) != norm(PIM_TIT_CUI):
        errores.append(f'Hoja "{hp.title}", celda {PIM_CELDA_TIT_CUI}: dice "{c1 if not vacio(c1) else "(vacía)"}" '
                       f'y debería decir "{PIM_TIT_CUI}". Puede que se haya insertado o borrado una fila arriba, '
                       f'así que la lista de insertados ya no empezaría en A{PIM_FILA_INI}.')
    a2 = celda(fila2, 0)
    if not vacio(a2):
        errores.append(f'Hoja "{hp.title}", celda A2: debería estar vacía y dice "{a2}". '
                       f'La lista de insertados debe empezar en A{PIM_FILA_INI}; si empieza antes, '
                       f'puede que se haya borrado una fila arriba.')

    if errores:
        raise Detener('\n'.join('- ' + e for e in errores))

    observaciones = []

    # recién insertados
    insertados, vistos, repetidos = [], set(), []
    for i, fila in enumerate(col_a_pim):
        v = celda(fila, 0)
        if vacio(v):
            continue
        cui = como_cui(v)
        ref = f'A{PIM_FILA_INI + i}'
        if cui is None:
            observaciones.append({'hoja': hp.title, 'celda': ref, 'valor': v,
                                  'motivo': 'no es un número de 7 cifras; se dejó fuera'})
        elif cui in vistos:
            repetidos.append({'hoja': hp.title, 'celda': ref, 'cui': cui})
        else:
            vistos.add(cui)
            insertados.append(cui)

    # todos menos PC
    todos, vistos_t, n_pc, n_filas = [], set(), 0, 0
    for i, fila in enumerate(datos_proy):
        v = celda(fila, PROY_COL_CUI - 1)
        etapa = celda(fila, PROY_COL_ETAPA - 1)
        if vacio(v):
            if not vacio(etapa):
                observaciones.append({'hoja': hr.title, 'celda': f'A{PROY_FILA_INI + i}', 'valor': '',
                                      'motivo': f'fila con Etapa "{etapa}" pero sin CUI; se dejó fuera'})
            continue
        n_filas += 1
        cui = como_cui(v)
        ref = f'{col_letra(PROY_COL_CUI)}{PROY_FILA_INI + i}'
        if cui is None:
            observaciones.append({'hoja': hr.title, 'celda': ref, 'valor': v,
                                  'motivo': 'no es un número de 7 cifras; se dejó fuera'})
            continue
        if norm(etapa) == ETAPA_EXCLUIDA:
            n_pc += 1
            continue
        if cui in vistos_t:
            repetidos.append({'hoja': hr.title, 'celda': ref, 'cui': cui})
            continue
        vistos_t.add(cui)
        todos.append(cui)

    fecha_pim = fecha_serial(celda(fila1, 0))
    resumen = {
        'recien_insertados': len(insertados),
        'todos_cuis': len(todos),
        'filas_con_cui_en_proyeccion26': n_filas,
        'excluidos_por_etapa_PC': n_pc,
        'valores_invalidos': len(observaciones),
        'repetidos_omitidos': len(repetidos),
    }
    return insertados, todos, observaciones, repetidos, resumen, fecha_pim


# ---------------- paso 2: SSI del MEF ----------------
def num(x):
    try:
        n = float(x)
        return n if n == n else 0.0      # NaN -> 0
    except (TypeError, ValueError):
        return 0.0


def consultar_ssi(cui):
    """('ok', ficha) | ('no_existe', None) | ('falla', 'motivo')."""
    motivo = ''
    for intento in range(1, SSI_INTENTOS + 1):
        try:
            p = subprocess.run(
                ['curl.exe', '-s', '-m', str(SSI_TIEMPO_MAX), '-X', 'POST',
                 '-H', 'X-Requested-With: XMLHttpRequest', '-H', 'Referer: https://ofi5.mef.gob.pe/ssi/',
                 '-d', f'id={cui}', '-w', '\n%{http_code}', SSI_SERVICIO],
                capture_output=True, timeout=SSI_TIEMPO_MAX + 15)
            cuerpo, _, codigo = p.stdout.rpartition(b'\n')
            codigo = codigo.decode('ascii', 'replace').strip()
            if p.returncode != 0:
                motivo = f'sin conexión con el servicio (curl {p.returncode})'
            elif codigo != '200':
                motivo = f'el servicio respondió con código {codigo}'
            else:
                try:
                    texto = cuerpo.decode('utf-8')
                except UnicodeDecodeError:
                    texto = cuerpo.decode('latin-1')     # el MEF a veces manda Latin-1
                try:
                    datos = json.loads(texto)
                except ValueError:
                    motivo = 'respuesta que no es JSON (posible página de error)'
                else:
                    if isinstance(datos, list):
                        return ('ok', datos[0]) if datos else ('no_existe', None)
                    if isinstance(datos, dict) and datos:
                        return 'ok', datos
                    return 'no_existe', None
        except subprocess.TimeoutExpired:
            motivo = 'el servicio no respondió a tiempo'
        if intento < SSI_INTENTOS:
            time.sleep(2 * intento)
    return 'falla', motivo


def descargar_ssi(insertados, todos):
    set_ins = set(insertados)
    union = list(dict.fromkeys(list(todos) + list(insertados)))
    t0 = time.time()
    with ThreadPoolExecutor(SSI_HILOS) as ex:
        respuestas = dict(zip(union, ex.map(consultar_ssi, union)))
    segundos = round(time.time() - t0, 1)

    filas, no_existe, fallas, no_registrados = [], [], [], []
    for cui in union:
        estado, dato = respuestas[cui]
        fila = {'cui': cui, 'recien_insertado': cui in set_ins, 'nombre_inversion': '', 'estado_ssi': '',
                'monto_inversion': None, 'resultado': 'ok'}
        if cui in set_ins:
            fila.update({'uei': None, 'tipo_inversion': None})
        if estado == 'ok':
            fila['nombre_inversion'] = ' '.join(str(dato.get('NOMBRE_INVERSION') or '').split())
            fila['estado_ssi'] = (dato.get('ESTADO') or '').strip()
            fila['monto_inversion'] = round(num(dato.get('MTO_F8_C')) + num(dato.get('MTO_LAUDO'))
                                            + num(dato.get('MTO_CartFza')), 2)   # vacío -> 0
            if cui in set_ins:
                fila['uei'] = (dato.get('DES_UNIDAD_UEI') or '').strip()
                fila['tipo_inversion'] = (dato.get('TIPO_FORMATO') or '').strip()
            if fila['estado_ssi'].upper().startswith('NO REGISTRADO'):
                no_registrados.append(cui)
        elif estado == 'no_existe':
            fila['resultado'] = 'el SSI no encuentra este CUI'
            no_existe.append(cui)
        else:
            fila['resultado'] = 'sin respuesta del servicio: ' + dato
            fallas.append((cui, dato))
        filas.append(fila)

    resumen = {
        'cuis_consultados': len(union),
        'con_dato': sum(1 for f in filas if f['resultado'] == 'ok'),
        'no_encontrados_en_ssi': len(no_existe),
        'sin_respuesta_del_servicio': len(fallas),
        'no_registrados_en_invierte_monto_0': len(no_registrados),
        'insertados_fuera_de_todos_cuis': sorted(set_ins - set(todos)),
        'segundos': segundos,
    }
    return filas, resumen, no_existe, fallas


def corrida_anterior():
    """Ruta y datos {cui: fila} de la última corrida guardada en CARPETA_MONTOS (o None)."""
    for p in sorted(CARPETA_MONTOS.glob('monto_ssi_*.json'), reverse=True):
        try:
            d = json.loads(p.read_text(encoding='utf-8'))
            return p, {int(f['cui']): f for f in d['datos']}
        except Exception as e:
            print('No se pudo leer la corrida anterior', p.name, '->', e, '(se prueba con la previa)')
    return None, {}


def comparar_con_anterior(filas):
    """
    Fija en cada fila 'monto_final' (el que se guarda y se escribe en la hoja) y compara con la corrida
    anterior (su 'monto_final' si lo tiene; si no, su 'monto_inversion').
      - Monto que pasa de un valor > 0 a 0: dato sospechoso del SSI -> se queda el valor anterior y se avisa.
      - Monto null (no encontrado o sin respuesta): se queda el valor anterior si existe.
      - Cualquier otro cambio: va a la lista de variaciones.
    Devuelve (ruta_anterior, variaciones, ceros_sospechosos).
    """
    ruta, ant = corrida_anterior()
    variaciones, ceros = [], []
    for f in filas:
        previo = ant.get(f['cui'])
        m_ant = None
        if previo:
            m_ant = previo.get('monto_final', previo.get('monto_inversion'))
        m_act = f['monto_inversion']
        f['monto_final'], f['observacion'] = m_act, ''
        if m_act is None:
            if m_ant is not None:
                f['monto_final'] = m_ant
                f['observacion'] = 'sin dato del SSI: se mantuvo el monto anterior'
            continue
        if m_ant is None:
            continue
        if m_ant > 0 and m_act == 0:
            f['monto_final'] = m_ant
            if previo.get('monto_inversion') == 0:
                # ya llegó en 0 la corrida anterior y entonces se avisó: se mantiene sin volver a avisar
                f['observacion'] = 'el SSI sigue trayendo 0: se mantuvo el anterior (ya avisado antes, sin alerta)'
                continue
            f['observacion'] = 'el SSI trajo 0 y antes tenía monto: se mantuvo el anterior'
            ceros.append({'cui': f['cui'], 'nombre_inversion': f.get('nombre_inversion') or previo.get('nombre_inversion', ''),
                          'monto_anterior': m_ant, 'estado_ssi': f.get('estado_ssi', '')})
            continue
        if abs(m_act - m_ant) >= 0.005:
            pct = None if m_ant == 0 else round((m_act - m_ant) / m_ant * 100, 2)
            variaciones.append({'cui': f['cui'], 'nombre_inversion': f.get('nombre_inversion') or previo.get('nombre_inversion', ''),
                                'monto_anterior': m_ant, 'monto_actual': m_act, 'variacion_pct': pct})
    variaciones.sort(key=lambda v: -(abs(v['variacion_pct']) if v['variacion_pct'] is not None else float('inf')))
    return ruta, variaciones, ceros


def excel_variaciones(variaciones, ceros, sello, ruta_anterior):
    """Excel con las variaciones (y los ceros sospechosos en otra hoja). Devuelve la ruta."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    wb = Workbook()
    ws = wb.active
    ws.title = 'Variaciones'
    cab = ['CUI', 'Nombre de la inversión', 'Monto anterior (S/)', 'Monto actual (S/)', '% de variación']
    ws.append(cab)
    for v in variaciones:
        ws.append([v['cui'], v['nombre_inversion'], v['monto_anterior'], v['monto_actual'],
                   None if v['variacion_pct'] is None else v['variacion_pct'] / 100])
    alerta = PatternFill('solid', fgColor='FCE4D6')
    for fila in ws.iter_rows(min_row=2):
        fila[0].number_format = '0'
        fila[2].number_format = fila[3].number_format = '#,##0.00'
        fila[4].number_format = '0.00%'
        pct = fila[4].value
        if pct is None:
            fila[4].value = 'antes 0'
        if pct is None or abs(pct) * 100 > UMBRAL_VARIACION_PCT:
            for c in fila:
                c.fill = alerta
    hojas = [ws]
    if ceros:
        w2 = wb.create_sheet('Ceros sospechosos')
        w2.append(['CUI', 'Nombre de la inversión', 'Monto anterior que se mantuvo (S/)', 'Estado en el SSI'])
        for z in ceros:
            w2.append([z['cui'], z['nombre_inversion'], z['monto_anterior'], z['estado_ssi']])
        for fila in w2.iter_rows(min_row=2):
            fila[0].number_format = '0'
            fila[2].number_format = '#,##0.00'
        hojas.append(w2)
    for h in hojas:
        for c in h[1]:
            c.font = Font(bold=True, color='FFFFFF')
            c.fill = PatternFill('solid', fgColor='1F3864')
            c.alignment = Alignment(wrap_text=True, vertical='center')
        for col, ancho in zip('ABCDE', (11, 80, 20, 20, 15)):
            h.column_dimensions[col].width = ancho
        h.freeze_panes = 'A2'
    nota = wb.create_sheet('Nota')
    nota.append(['Comparación de la corrida', sello])
    nota.append(['contra la corrida anterior', ruta_anterior.name if ruta_anterior else '(ninguna)'])
    nota.append(['Umbral de alerta (%)', UMBRAL_VARIACION_PCT])
    nota.append(['Filas resaltadas', f'variación mayor a {UMBRAL_VARIACION_PCT} % o monto anterior en 0'])
    nota.column_dimensions['A'].width = 30
    nota.column_dimensions['B'].width = 50
    CARPETA_MONTOS.mkdir(parents=True, exist_ok=True)
    ruta = CARPETA_MONTOS / f'variacion_montos_{sello}.xlsx'
    wb.save(ruta)
    return ruta


def guardar_montos(filas, resumen, ahora, sello, archivo_cuis):
    CARPETA_MONTOS.mkdir(parents=True, exist_ok=True)
    base = CARPETA_MONTOS / f'monto_ssi_{sello}'
    salida = {
        'generado': ahora.strftime('%Y-%m-%d %H:%M:%S'),
        'fuente': SSI_SERVICIO,
        'monto_inversion': 'COSTO TOTAL DE LA INVERSIÓN ACTUALIZADO (S/) (a+b+c+d), tal como lo devolvió el SSI',
        'monto_final': 'el que se escribe en la hoja: igual al descargado, salvo 0 sospechoso o sin dato '
                       '(se mantiene el de la corrida anterior)',
        'lista_de_cuis': str(archivo_cuis),
        'resumen': resumen,
        'datos': filas,
    }
    Path(str(base) + '.json').write_text(json.dumps(salida, ensure_ascii=False, indent=2), encoding='utf-8')
    with open(str(base) + '.csv', 'w', newline='', encoding='utf-8-sig') as f:
        w = csv.writer(f, delimiter=';')
        w.writerow(['fecha_hora', 'cui', 'nombre_inversion', 'recien_insertado', 'estado_ssi',
                    'monto_descargado', 'monto_final', 'uei', 'tipo_inversion', 'resultado', 'observacion'])
        txt = lambda m: '' if m is None else f'{m:.2f}'.replace('.', ',')
        for x in filas:
            w.writerow([salida['generado'], x['cui'], x.get('nombre_inversion', ''),
                        'SI' if x['recien_insertado'] else '', x['estado_ssi'],
                        txt(x['monto_inversion']), txt(x.get('monto_final', x['monto_inversion'])),
                        x.get('uei') or '', x.get('tipo_inversion') or '', x['resultado'],
                        x.get('observacion', '')])
    return Path(str(base) + '.json'), Path(str(base) + '.csv')


# ---------------- paso 3: escribir en PIM-diario ----------------
def escribir_pim(libro, hp, filas, insertados):
    """
    Escribe en una sola operación:
      O = monto de cada CUI de la columna N (si el monto vino vacío por falla del SSI, se deja
          el valor que ya tenía la celda);
      P y Q = UEI y tipo solo en las filas de los recién insertados (las demás quedan vacías).
    Una columna cuyo título no coincide no se toca; las otras sí. Devuelve (avisos, resumen).
    """
    t = hp.title
    lect = libro.values_batch_get(
        [f"'{t}'!{PIM_COL_LLAVE}1:{PIM_COL_TIPO}1", f"'{t}'!{PIM_COL_LLAVE}2:{PIM_COL_LLAVE}",
         f"'{t}'!{PIM_COL_MONTO}2:{PIM_COL_TIPO}"],
        params={'valueRenderOption': 'UNFORMATTED_VALUE'})['valueRanges']
    tits = (lect[0].get('values') or [[]])[0]
    col_n = lect[1].get('values', [])
    previo = lect[2].get('values', [])
    avisos = []

    # títulos
    llave = celda(tits, 0)
    if norm(llave) != norm(PIM_TIT_LLAVE):
        avisos.append(f'No se escribió nada en "{t}": la celda {PIM_COL_LLAVE}1 dice '
                      f'"{llave if not vacio(llave) else "(vacía)"}" y debería decir "{PIM_TIT_LLAVE}". '
                      'Sin esa columna no se sabe a qué CUI corresponde cada fila.')
        return avisos, {'escrito': False}
    columnas_ok = {}
    for i, (col, esperado) in enumerate(((PIM_COL_MONTO, PIM_TIT_MONTO), (PIM_COL_UEI, PIM_TIT_UEI),
                                         (PIM_COL_TIPO, PIM_TIT_TIPO)), start=1):
        hallado = celda(tits, i)
        columnas_ok[col] = norm(hallado) == norm(esperado)
        if not columnas_ok[col]:
            avisos.append(f'No se escribió la columna {col} de "{t}": {col}1 dice '
                          f'"{hallado if not vacio(hallado) else "(vacía)"}" y debería decir "{esperado}". '
                          'Las demás columnas sí se escribieron.')

    por_cui = {f['cui']: f for f in filas}
    set_ins = set(insertados)
    n = len(col_n)
    alto = max(n, len(previo))            # también se limpian filas sobrantes de corridas anteriores
    o_vals, p_vals, q_vals = [], [], []
    sin_dato, puestos_ins = [], set()
    for i in range(alto):
        anterior_o = celda(previo[i], 0) if i < len(previo) else ''
        if i >= n:
            o_vals.append(['']); p_vals.append(['']); q_vals.append([''])
            continue
        v = celda(col_n[i], 0)
        cui = como_cui(v)
        f = por_cui.get(cui) if cui is not None else None
        monto = None if f is None else f.get('monto_final', f['monto_inversion'])
        if monto is None:
            o_vals.append([anterior_o])           # se conserva lo que había
            if not vacio(v) and cui not in CUIS_SIN_AVISO:
                sin_dato.append(f'{PIM_COL_LLAVE}{i + 2} = {v}')
        else:
            o_vals.append([monto])
        if cui in set_ins and f is not None and f.get('uei') is not None:
            p_vals.append([f['uei']]); q_vals.append([f['tipo_inversion']])
            puestos_ins.add(cui)
        else:
            p_vals.append(['']); q_vals.append([''])

    if sin_dato:
        avisos.append(f'{len(sin_dato)} fila(s) de la columna {PIM_COL_LLAVE} sin monto nuevo (CUI no descargado o '
                      f'sin respuesta del SSI); se dejó en {PIM_COL_MONTO} el valor anterior:\n'
                      + '\n'.join('- ' + s for s in sin_dato))
    faltan_ins = [c for c in insertados if c not in puestos_ins]
    if faltan_ins and (columnas_ok[PIM_COL_UEI] or columnas_ok[PIM_COL_TIPO]):
        avisos.append(f'{len(faltan_ins)} recién insertado(s) sin UEI ni tipo en "{t}" porque no están en la '
                      f'columna {PIM_COL_LLAVE} o el SSI no los encontró:\n' + '\n'.join(f'- {c}' for c in faltan_ins))

    datos = []
    if alto:
        fin = alto + 1
        for col, vals in ((PIM_COL_MONTO, o_vals), (PIM_COL_UEI, p_vals), (PIM_COL_TIPO, q_vals)):
            if columnas_ok[col]:
                datos.append({'range': f"'{t}'!{col}2:{col}{fin}", 'values': vals})
    if datos:
        libro.values_batch_update({'valueInputOption': 'RAW', 'data': datos})
        if columnas_ok[PIM_COL_MONTO]:
            hp.format(f'{PIM_COL_MONTO}2:{PIM_COL_MONTO}{alto + 1}',
                      {'numberFormat': {'type': 'NUMBER', 'pattern': '#,##0.00'}})
    resumen = {'escrito': bool(datos), 'filas_columna_N': n,
               'columnas_escritas': [c for c in (PIM_COL_MONTO, PIM_COL_UEI, PIM_COL_TIPO) if columnas_ok[c]],
               'celdas_con_monto': sum(1 for x in o_vals[:n] if isinstance(x[0], (int, float))),
               'insertados_con_uei_tipo': len(puestos_ins)}
    return avisos, resumen


# ---------------- principal ----------------
def main():
    ahora = datetime.now()
    sello = ahora.strftime('%Y-%m-%d_%H%M%S')

    if '--probar-correo' in sys.argv:
        ok = enviar_correo('Correo de prueba', 'Prueba del aviso del programa cuis_desde_sheets.py.\n'
                           'Si llegó este correo, los avisos funcionan.\n\n' + ahora.strftime('%d/%m/%Y %H:%M:%S'))
        return 0 if ok else 4

    hp = hr = None
    try:
        libro, hp, hr, vr = leer()
        insertados, todos, obs, repetidos, resumen, fecha_pim = validar_y_armar(hp, hr, vr)
    except Detener as e:
        cuerpo = ('El programa se DETUVO y no generó el archivo de CUI.\n\n'
                  'Problemas encontrados:\n' + str(e) + '\n\n'
                  'Revisar la hoja y volver a correr el programa.\n'
                  f'Libro: {enlace(hp.id if hp else None)}\n'
                  f'Hora: {ahora:%d/%m/%Y %H:%M:%S}')
        print(cuerpo)
        return 2 if enviar_correo('Detenido: los títulos de la hoja no coinciden', cuerpo) else 4
    except Exception as e:
        cuerpo = ('Error inesperado; no se generó el archivo de CUI.\n\n'
                  f'Detalle: {type(e).__name__}: {e}\n\n{traceback.format_exc()}\n'
                  f'Libro: {enlace()}\nHora: {ahora:%d/%m/%Y %H:%M:%S}')
        print(cuerpo)
        return 3 if enviar_correo('Error al leer el libro de Google', cuerpo) else 4

    salida = {
        'generado': ahora.strftime('%Y-%m-%d %H:%M:%S'),
        'libro': {'id': ID_LIBRO, 'titulo': libro.title},
        'fecha_descarga_pim_diario_A1': fecha_pim,
        'resumen': resumen,
        'recien_insertados': insertados,
        'todos_cuis': todos,
        'observaciones': obs,
        'repetidos_omitidos': repetidos,
    }
    CARPETA_SALIDA.mkdir(parents=True, exist_ok=True)
    archivo = CARPETA_SALIDA / f'cuis_{sello}.json'
    archivo.write_text(json.dumps(salida, ensure_ascii=False, indent=2), encoding='utf-8')
    print('Archivo:', archivo)
    print('Resumen:', json.dumps(resumen, ensure_ascii=False))

    # ---- paso 2: SSI ----
    avisos, asuntos, codigo = [], [], 0
    f_var = None
    try:
        filas, res_ssi, no_existe, fallas = descargar_ssi(insertados, todos)

        # ---- comparación con la corrida anterior (antes de guardar, para fijar el monto final) ----
        ruta_ant, variaciones, ceros = comparar_con_anterior(filas)
        grandes = [v for v in variaciones
                   if v['variacion_pct'] is None or abs(v['variacion_pct']) > UMBRAL_VARIACION_PCT]
        res_ssi['comparacion'] = {
            'corrida_anterior': ruta_ant.name if ruta_ant else None,
            'montos_que_cambiaron': len(variaciones),
            f'cambios_mayores_a_{UMBRAL_VARIACION_PCT}_pct': len(grandes),
            'ceros_sospechosos_se_mantuvo_el_anterior': len(ceros),
        }
        f_var = excel_variaciones(variaciones, ceros, sello, ruta_ant) if (variaciones or ceros) else None
        if f_var:
            print('Variaciones:', f_var)
        mon = lambda x: f'{x:,.2f}'
        if grandes:
            avisos.append(
                f'{len(grandes)} monto(s) con variación mayor a {UMBRAL_VARIACION_PCT} % respecto de la corrida anterior '
                f'({ruta_ant.name}). En total cambiaron {len(variaciones)} monto(s); el detalle está en el Excel.\n'
                + '\n'.join(f'- {v["cui"]} {v["nombre_inversion"][:70]}: {mon(v["monto_anterior"])} -> '
                            f'{mon(v["monto_actual"])} ('
                            + ('antes 0' if v['variacion_pct'] is None else f'{v["variacion_pct"]:+.2f} %') + ')'
                            for v in grandes))
            asuntos.append(f'{len(grandes)} monto(s) variaron más de {UMBRAL_VARIACION_PCT} %')
        if ceros:
            avisos.append(
                f'{len(ceros)} CUI llegaron con monto 0 desde el SSI cuando en la corrida anterior tenían monto. '
                'Se tomó como dato erróneo: se MANTUVO el monto anterior (en la hoja y en el archivo):\n'
                + '\n'.join(f'- {z["cui"]} {z["nombre_inversion"][:70]}: se mantuvo {mon(z["monto_anterior"])}'
                            for z in ceros))
            asuntos.append(f'{len(ceros)} monto(s) en 0 sospechosos')

        f_json, f_csv = guardar_montos(filas, res_ssi, ahora, sello, archivo)
        print('Montos:', f_json)
        print('        ', f_csv)
        print('Resumen SSI:', json.dumps(res_ssi, ensure_ascii=False))
    except Exception as e:
        filas, res_ssi, no_existe, fallas = None, None, [], []
        avisos.append('ERROR en la descarga del SSI; no se guardaron montos.\n'
                      f'Detalle: {type(e).__name__}: {e}\n{traceback.format_exc()}')
        asuntos.append('error en la descarga del SSI')
        codigo = 3

    todo_caido = bool(res_ssi) and bool(fallas) and len(fallas) == res_ssi['cuis_consultados']
    # CUI que ya se sabe que no existen en el SSI: no se avisan (siguen anotados en los archivos)
    no_existe = [c for c in no_existe if c not in CUIS_SIN_AVISO]
    if not todo_caido:
        fallas = [(c, m) for c, m in fallas if c not in CUIS_SIN_AVISO]

    if res_ssi is not None:
        if todo_caido:
            avisos.append(f'El servicio del SSI no respondió para NINGUNO de los {len(fallas)} CUI '
                          f'(motivo del primero: {fallas[0][1]}). Probablemente el MEF está caído o '
                          'no hay internet. Los montos quedaron vacíos (null), no en cero.')
            asuntos.append('el SSI no responde')
            codigo = 5
        elif fallas:
            avisos.append(f'{len(fallas)} CUI sin respuesta del servicio del SSI tras {SSI_INTENTOS} intentos '
                          '(monto vacío, no cero):\n' + '\n'.join(f'- {c}: {m}' for c, m in fallas))
            asuntos.append(f'{len(fallas)} CUI sin respuesta del SSI')
        if no_existe:
            avisos.append(f'{len(no_existe)} CUI que el SSI no encuentra (el servicio responde, pero sin ficha; '
                          'revisar si el código está bien escrito en la hoja):\n'
                          + '\n'.join(f'- {c}{" (recién insertado)" if c in insertados else ""}' for c in no_existe))
            asuntos.append(f'{len(no_existe)} CUI no encontrados en el SSI')
        fuera = res_ssi['insertados_fuera_de_todos_cuis']
        if fuera:
            avisos.append(f'{len(fuera)} CUI recién insertados (PIM-diario, columna A) que NO están en la lista '
                          f'de {HOJA_PROY} sin PC. Se descargaron igual. Puede que tengan Etapa "PC" o que su '
                          'celda en Proyección26 tenga un error:\n' + '\n'.join(f'- {c}' for c in fuera))
            asuntos.append('insertados fuera de la lista')

        # ---- paso 3: escribir en PIM-diario ----
        if todo_caido:
            avisos.append(f'No se escribió nada en "{hp.title}" porque el SSI no respondió; '
                          'la hoja se quedó con los datos de la corrida anterior.')
        else:
            try:
                av_hoja, res_hoja = escribir_pim(libro, hp, filas, insertados)
                print('Hoja:', json.dumps(res_hoja, ensure_ascii=False))
                if av_hoja:
                    avisos.extend(av_hoja)
                    asuntos.append('problemas al escribir en ' + hp.title)
            except Exception as e:
                avisos.append(f'ERROR al escribir en "{hp.title}"; los montos sí quedaron guardados en los archivos.\n'
                              f'Detalle: {type(e).__name__}: {e}\n{traceback.format_exc()}')
                asuntos.append('error al escribir en ' + hp.title)
                codigo = 3

    if obs:
        lineas = '\n'.join(f'- Hoja "{o["hoja"]}", celda {o["celda"]}: "{o["valor"]}" -> {o["motivo"]}' for o in obs)
        avisos.insert(0, f'{len(obs)} celda(s) que no son CUI válidos (número de 7 cifras). Se dejaron fuera; '
                         'el resto se procesó normalmente:\n' + lineas)
        asuntos.insert(0, f'{len(obs)} valor(es) que no son CUI')

    if avisos:
        cuerpo = ('\n\n'.join(f'{i}) {a}' for i, a in enumerate(avisos, 1)) + '\n\n'
                  f'Lista de CUI: {archivo}\n'
                  + (f'Montos: {f_json}\n' if res_ssi is not None else '')
                  + (f'Excel de variaciones: {f_var}\n' if f_var else '')
                  + f'Libro: {enlace()}\nHora: {ahora:%d/%m/%Y %H:%M:%S}')
        if not enviar_correo('; '.join(asuntos), cuerpo):
            return 4
    return codigo


if __name__ == '__main__':
    sys.exit(main())
