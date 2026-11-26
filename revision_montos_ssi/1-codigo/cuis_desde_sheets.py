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
  - UEI (DES_UNIDAD_UEI) y tipo de inversión (TIPO_FORMATO) de TODOS los CUI: ya vienen en la
    misma respuesta, así que no cuesta ninguna consulta extra
  - Se consulta la unión de las dos listas, una vez por CUI (si un insertado no está en
    todos_cuis, igual se consulta y se avisa).
  - Monto vacío -> 0. CUI que el MEF no encuentra o que no responde -> monto null (no 0, para
    no pisar un dato bueno con un cero cuando el servicio falla) y aviso por correo.
  - Resultado en CARPETA_MONTOS, un JSON y un CSV por corrida con fecha y hora en el nombre.
  - Se usa curl.exe de Windows: el servidor del MEF no envía su certificado intermedio y
    Python lo rechaza; curl lo completa con el almacén de Windows sin desactivar la verificación.
  - El SSI bloquea a los servidores de Google: por eso esto corre en la PC y no en Apps Script.

Paso 3: escritura en Proyección26, ubicando cada fila por su CUI de la columna A
  - I  "(A) Monto de inversión" = COSTO TOTAL DE LA INVERSIÓN ACTUALIZADO. Si el SSI no dio
       monto, la celda NO se toca (nunca se pisa un dato bueno con 0 ni con vacío).
  - D  "UEI" = la sigla de Proyección26 equivalente al texto del SSI (GRI, GRDE, GRDIS,
       GRRNyGA, PET, DIRESA, DRAT, DRTYC). Si no hay equivalencia, la celda no se toca y se avisa.
  - F  "Etapa" = "PC" si el SSI dice PROYECTO PROCOMPITE, "Desactivado" si está desactivado.
       Las demás etapas todavía no se saben derivar del SSI: esas celdas no se tocan (solo se
       avisa de los recién insertados). OJO: marcar "PC" saca a ese CUI de las corridas siguientes.
  - Solo se escriben las celdas que CAMBIAN, en tramos de filas consecutivas. Una celda que no
    corresponde actualizar no entra en el envío, así no se pisa ni se congela ninguna fórmula.
  - Justo antes de escribir se vuelve a leer la columna A y los títulos: si las filas se
    corrieron o una columna se movió durante la corrida, no se escribe nada y se avisa.
  - PIM-diario ya no se escribe; sus columnas O, P y Q se vacían (LIMPIAR_PIM_OPQ) para que no
    queden montos viejos engañando a las fórmulas.
  - CUIS_SIN_AVISO (3000882 y 3999999): no generan aviso aunque el SSI no los encuentre.

Comparación con la corrida anterior (último monto_ssi_*.json de CARPETA_MONTOS)
  - Montos que cambiaron -> Excel variacion_montos_<fecha>.xlsx (CUI, nombre, monto anterior,
    monto actual, % de variación con signo; resaltadas las que pasan el umbral).
  - El correo alerta solo las variaciones mayores a UMBRAL_VARIACION_PCT (en más o en menos).
  - La base es la última corrida de un día ANTERIOR (las del mismo día no sirven de base). Si no
    es la del día hábil previo se avisa; sábado y domingo no cuentan como falta (la PC está apagada).
  - Monto que llega en 0 cuando antes tenía valor: se toma como error del SSI, se mantiene el
    anterior (hoja y archivo, campo monto_final) y se avisa SOLO la primera vez; si sigue llegando
    en 0 en las corridas siguientes, se sigue manteniendo el anterior sin volver a avisar (queda
    anotado en el campo observacion). Si el SSI no da dato, también se mantiene el anterior.
  - Si los ceros nuevos llegan a MAX_CEROS_SOSPECHOSOS, se considera falla de la página del MEF:
    NO se escribe nada en la hoja y se avisa (un 0 o dos son un cambio real; veinte son basura).
  - Si CARPETA_MONTOS no está disponible, los archivos se guardan en CARPETA_RESPALDO y se avisa.

Solo corre una instancia a la vez (ARCHIVO_BLOQUEO): si la tarea programada y una prueba manual
coinciden, la segunda no hace nada.

Todos los avisos de una corrida van en un solo correo.

Uso:
  python cuis_desde_sheets.py                 corrida normal
  python cuis_desde_sheets.py --probar-correo envía un correo de prueba y termina

Códigos de salida: 0 bien (aunque haya avisos enviados por correo), 2 detenido por títulos
o estructura, 3 error inesperado, 4 no se pudo enviar el correo de aviso, 5 el SSI no
respondió para ningún CUI, 6 detenido por ceros masivos (data basura del SSI).
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
CARPETA_RESPALDO = Path(r'C:\Users\HP\Downloads')   # si CARPETA_MONTOS no está disponible (se avisa)
ARCHIVO_BLOQUEO = CARPETA_BASE / 'nata_apuntes' / 'corrida_en_curso.lock'
BLOQUEO_MAX_MIN = 30     # un .lock más viejo que esto se considera abandonado y se reusa

# SSI del MEF
SSI_SERVICIO = 'https://ofi5.mef.gob.pe/invierteWS/Ssi/traeDetInvSSI'
SSI_HILOS = 12          # consultas simultáneas
SSI_INTENTOS = 3        # por CUI, con pausa creciente
SSI_TIEMPO_MAX = 40     # segundos por consulta
CUIS_SIN_AVISO = {3000882, 3999999}   # no existen en el SSI y ya se sabe: no avisar por correo

# Comparación con la corrida anterior
UMBRAL_VARIACION_PCT = 10   # el correo alerta solo variaciones mayores a este % (en aumento o disminución)
MAX_CEROS_SOSPECHOSOS = 5   # desde esta cantidad de ceros nuevos se considera falla del SSI: no se escribe nada

# ---- Paso 3: escritura en Proyección26 (ubicada por el CUI de la columna A) ----
# Dónde debe estar cada cosa (fila y columna de la hoja, empezando en 1)
PROY_FILA_TITULOS = 2
PROY_FILA_INI = 3
PROY_COL_CUI, PROY_TIT_CUI = 1, 'CUI'                              # A2  (se lee)
PROY_COL_UEI, PROY_TIT_UEI = 4, 'UEI'                              # D2  (se escribe)
PROY_COL_ETAPA, PROY_TIT_ETAPA = 6, 'Etapa'                        # F2  (se lee y se escribe)
PROY_COL_MONTO, PROY_TIT_MONTO = 9, '(A) Monto de inversión'       # I2  (se escribe)
PROY_COL_ULT = 9             # hasta dónde se lee la hoja (A:I)
ETAPA_EXCLUIDA = 'PC'

# PIM-diario: solo se LEE (lista de recién insertados y anclas de estructura).
PIM_CELDA_TIT_CUI, PIM_TIT_CUI = 'C1', 'CUI' # ancla: si alguien inserta una fila arriba, C1 deja de decir CUI
PIM_FILA_INI = 3                              # A3 hacia abajo: insertados (A1 = fecha, A2 vacía)
# Las columnas O, P y Q de PIM-diario ya no se usan (el monto ahora va a Proyección26 I).
# Con esto en True se vacían una vez por corrida para que no queden montos viejos engañando a las fórmulas.
LIMPIAR_PIM_OPQ = True
PIM_RANGO_MUERTO = 'O2:Q'

# ---- Equivalencias SSI -> Proyección26 ----
# UEI: del texto largo del SSI (DES_UNIDAD_UEI) a la sigla de Proyección26.
# Se busca por palabras distintivas, sin tildes y en mayúsculas. EL ORDEN IMPORTA:
# el PET se evalúa primero porque su nombre también contiene "INFRAESTRUCTURA".
UEI_EQUIVALENCIAS = [
    ('PET',      (('PROYECTO ESPECIAL',), ('AFIANZAMIENTO',), ('RECURSOS HIDRICOS',), ('PET',))),
    ('GRI',      (('INFRAESTRUCTURA',),)),
    ('GRDE',     (('DESARROLLO ECONOMICO',),)),
    ('GRDIS',    (('DESARROLLO SOCIAL',), ('DESARROLLO E INCLUSION',), ('INCLUSION SOCIAL',))),
    ('GRRNyGA',  (('RECURSOS NATURALES',),)),
    ('DIRESA',   (('SALUD',),)),
    ('DRAT',     (('AGRICULTURA',),)),
    ('DRTYC',    (('TRANSPORTES',),)),
]

# Etapa: solo estos dos casos se saben derivar del SSI. Derivar las demás etiquetas
# (ET, EF, L...) queda pendiente, así que esas filas NO se tocan.
ETAPA_TIPO_PROCOMPITE = 'PROYECTO PROCOMPITE'   # TIPO_FORMATO -> Etapa "PC"
ETAPA_VALOR_PROCOMPITE = 'PC'
ETAPA_VALOR_DESACTIVADO = 'Desactivado'         # ESTADO que empieza con "DESACTIVADO"


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


def carpeta_montos():
    """CARPETA_MONTOS si se puede escribir; si no, CARPETA_RESPALDO. Devuelve (carpeta, aviso|None)."""
    for carpeta, es_respaldo in ((CARPETA_MONTOS, False), (CARPETA_RESPALDO, True)):
        try:
            carpeta.mkdir(parents=True, exist_ok=True)
            prueba = carpeta / '.escritura_ok'
            prueba.write_text('ok', encoding='utf-8')
            prueba.unlink()
        except Exception as e:      # no solo OSError: una ruta inválida de Windows lanza ValueError
            motivo = f'{type(e).__name__}: {e}'
            continue
        if not es_respaldo:
            return carpeta, None
        return carpeta, (f'La carpeta de montos no está disponible ({CARPETA_MONTOS}), así que los archivos de '
                         f'esta corrida se guardaron en {carpeta}. Motivo: {motivo}. '
                         'OJO: la comparación con la corrida anterior busca en la carpeta normal, así que '
                         'conviene mover estos archivos ahí cuando vuelva a estar disponible.')
    raise Detener(f'No se pudo escribir ni en {CARPETA_MONTOS} ni en {CARPETA_RESPALDO}: {motivo}')


def tomar_bloqueo():
    """Evita dos corridas a la vez. Devuelve True si se tomó el bloqueo."""
    try:
        ARCHIVO_BLOQUEO.parent.mkdir(parents=True, exist_ok=True)
        if ARCHIVO_BLOQUEO.exists():
            edad_min = (time.time() - ARCHIVO_BLOQUEO.stat().st_mtime) / 60
            if edad_min < BLOQUEO_MAX_MIN:
                print(f'Ya hay una corrida en curso desde hace {edad_min:.0f} min '
                      f'({ARCHIVO_BLOQUEO}). Esta corrida no hace nada.')
                return False
            print(f'Se encontró un bloqueo abandonado de hace {edad_min:.0f} min; se reusa.')
        ARCHIVO_BLOQUEO.write_text(f'{os.getpid()} {datetime.now():%Y-%m-%d %H:%M:%S}', encoding='utf-8')
        return True
    except OSError as e:
        print('No se pudo usar el archivo de bloqueo (' + str(e) + '); se sigue sin bloqueo.')
        return True


def soltar_bloqueo():
    try:
        ARCHIVO_BLOQUEO.unlink(missing_ok=True)
    except OSError:
        pass


def codigo_uei(texto):
    """Sigla de Proyección26 para el texto de UEI del SSI; None si no coincide con ninguna."""
    t = norm(texto)
    if not t:
        return None
    for sigla, grupos in UEI_EQUIVALENCIAS:
        for palabras in grupos:
            if all(p in t for p in palabras):
                return sigla
    return None


def etapa_desde_ssi(fila):
    """Etapa que corresponde según el SSI: "PC", "Desactivado" o None (no se sabe derivar)."""
    if norm(fila.get('tipo_inversion')) == norm(ETAPA_TIPO_PROCOMPITE):
        return ETAPA_VALOR_PROCOMPITE
    if norm(fila.get('estado_ssi')).startswith('DESACTIVADO'):
        return ETAPA_VALOR_DESACTIVADO
    return None


def tramos(por_fila):
    """{fila: valor} -> [(primera_fila, [valores consecutivos]), ...]; las filas ausentes NO se escriben."""
    salida = []
    for f in sorted(por_fila):
        if salida and f == salida[-1][0] + len(salida[-1][1]):
            salida[-1][1].append(por_fila[f])
        else:
            salida.append((f, [por_fila[f]]))
    return salida


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
    ult_proy = col_letra(PROY_COL_ULT)
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
    for col, esperado in ((PROY_COL_CUI, PROY_TIT_CUI), (PROY_COL_UEI, PROY_TIT_UEI),
                          (PROY_COL_ETAPA, PROY_TIT_ETAPA), (PROY_COL_MONTO, PROY_TIT_MONTO)):
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
    filas_proy = {}      # cui -> {fila, uei, etapa, monto} tal como está hoy en la hoja
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
        filas_proy[cui] = {'fila': PROY_FILA_INI + i,
                           'uei': celda(fila, PROY_COL_UEI - 1),
                           'etapa': etapa,
                           'monto': celda(fila, PROY_COL_MONTO - 1)}

    fecha_pim = fecha_serial(celda(fila1, 0))
    resumen = {
        'recien_insertados': len(insertados),
        'todos_cuis': len(todos),
        'filas_con_cui_en_proyeccion26': n_filas,
        'excluidos_por_etapa_PC': n_pc,
        'valores_invalidos': len(observaciones),
        'repetidos_omitidos': len(repetidos),
    }
    return insertados, todos, observaciones, repetidos, resumen, fecha_pim, filas_proy


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
                'monto_inversion': None, 'uei': None, 'tipo_inversion': None, 'resultado': 'ok'}
        if estado == 'ok':
            fila['nombre_inversion'] = ' '.join(str(dato.get('NOMBRE_INVERSION') or '').split())
            fila['estado_ssi'] = (dato.get('ESTADO') or '').strip()
            fila['monto_inversion'] = round(num(dato.get('MTO_F8_C')) + num(dato.get('MTO_LAUDO'))
                                            + num(dato.get('MTO_CartFza')), 2)   # vacío -> 0
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


def fecha_de_archivo(p):
    """Fecha del nombre monto_ssi_AAAA-MM-DD_HHMMSS.json; None si no se puede leer."""
    try:
        return datetime.strptime(p.name[len('monto_ssi_'):][:10], '%Y-%m-%d').date()
    except ValueError:
        return None


def dias_habiles_sin_corrida(desde, hoy):
    """Días de lunes a viernes entre dos fechas (sin incluirlas) en los que faltó la corrida."""
    faltan, d = [], desde + timedelta(days=1)
    while d < hoy:
        if d.weekday() < 5:          # 5 = sábado, 6 = domingo: la PC está apagada
            faltan.append(d)
        d += timedelta(days=1)
    return faltan


def corrida_anterior(hoy=None):
    """
    Base de comparación: la última corrida de un día ANTERIOR a hoy. Las corridas del mismo día
    no sirven de base (una prueba de la mañana no debe ser la referencia de la tarde).
    Si la base no es del día hábil previo, se avisa; sábado y domingo no cuentan como falta.
    Devuelve (ruta, {cui: fila}, avisos).
    """
    hoy = hoy or datetime.now().date()
    avisos = []
    try:
        candidatos = sorted(((f, p) for p in CARPETA_MONTOS.glob('monto_ssi_*.json')
                             for f in [fecha_de_archivo(p)] if f is not None and f < hoy), reverse=True)
    except OSError as e:
        return None, {}, [f'No se pudo mirar la carpeta de montos {CARPETA_MONTOS} ({e}). '
                          'Esta corrida va SIN comparación con la anterior, así que un 0 erróneo del SSI '
                          'podría entrar a la hoja sin que nada lo detecte.']
    for f, p in candidatos:
        try:
            d = json.loads(p.read_text(encoding='utf-8'))
            datos = {int(x['cui']): x for x in d['datos']}
        except Exception as e:
            avisos.append(f'El archivo de la corrida del {f} ({p.name}) está dañado y no se pudo leer '
                          f'({type(e).__name__}: {e}). Se buscó una corrida más antigua para comparar.')
            continue
        faltan = dias_habiles_sin_corrida(f, hoy)
        if faltan:
            avisos.append(f'La comparación se hizo contra la corrida del {f} ({p.name}), no contra el día '
                          f'hábil previo. No hay corrida de: '
                          + ', '.join(d.strftime('%A %d/%m') for d in faltan) + '. '
                          'Los porcentajes de variación son acumulados de varios días, no de un día.')
        return p, datos, avisos
    avisos.append('No se encontró NINGUNA corrida anterior en ' + str(CARPETA_MONTOS) + '. Esta corrida va '
                  'sin comparación: un 0 erróneo del SSI entraría a la hoja sin que nada lo detecte, y no se '
                  'puede saber qué montos cambiaron.')
    return None, {}, avisos


def comparar_con_anterior(filas, hoy=None):
    """
    Fija en cada fila 'monto_final' (el que se guarda y se escribe en la hoja) y compara con la corrida
    anterior (su 'monto_final' si lo tiene; si no, su 'monto_inversion').
      - Monto que pasa de un valor > 0 a 0: dato sospechoso del SSI -> se queda el valor anterior y se avisa.
      - Monto null (no encontrado o sin respuesta): se queda el valor anterior si existe.
      - Cualquier otro cambio: va a la lista de variaciones.
    Devuelve (ruta_anterior, variaciones, ceros_sospechosos, avisos_de_la_base).
    """
    ruta, ant, avisos_base = corrida_anterior(hoy)
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
    return ruta, variaciones, ceros, avisos_base


def excel_variaciones(variaciones, ceros, sello, ruta_anterior, carpeta):
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
    carpeta.mkdir(parents=True, exist_ok=True)
    ruta = carpeta / f'variacion_montos_{sello}.xlsx'
    wb.save(ruta)
    return ruta


def guardar_montos(filas, resumen, ahora, sello, archivo_cuis, carpeta):
    carpeta.mkdir(parents=True, exist_ok=True)
    base = carpeta / f'monto_ssi_{sello}'
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


# ---------------- paso 3: escribir en Proyección26 ----------------
def igual_monto(actual, nuevo):
    """True si la celda ya tiene ese monto (para no reescribirla sin necesidad)."""
    try:
        return abs(float(actual) - float(nuevo)) < 0.005
    except (TypeError, ValueError):
        return False


def limpiar_pim_opq(hp):
    """
    Vacía O, P y Q de PIM-diario: ya no se usan (el monto ahora va a Proyección26 I) y dejar
    montos viejos ahí engañaría a cualquier fórmula que los lea.
    Se usa batch_clear y no una escritura de '': escribir '' deja un texto vacío y entonces
    ESBLANCO() devuelve FALSO (comprobado en la hoja), lo que rompe CONTAR.BLANCO y CONTARA.
    """
    hp.batch_clear([PIM_RANGO_MUERTO])


def escribir_proyeccion(libro, hr, filas, filas_proy, insertados):
    """
    Escribe en Proyección26 ubicando cada fila por su CUI (columna A), no por posición:
      D  UEI    = sigla equivalente a la UEI del SSI; si no hay equivalencia NO se toca la celda
                  y se avisa (así no se pisa con vacío lo que haya escrito a mano).
      F  Etapa  = "PC" si el SSI dice PROYECTO PROCOMPITE, "Desactivado" si está desactivado.
                  Las demás etapas todavía no se saben derivar: esas celdas no se tocan.
      I  Monto  = COSTO TOTAL DE LA INVERSIÓN ACTUALIZADO. Si no hay monto nuevo no se toca.
    Solo se escriben las celdas que cambian, en tramos de filas consecutivas: una celda que no
    corresponde actualizar no entra en el envío, así no se pisa ni se congela nada.
    Devuelve (avisos, resumen).
    """
    t = hr.title
    set_ins = set(insertados)
    por_fila = {PROY_COL_UEI: {}, PROY_COL_ETAPA: {}, PROY_COL_MONTO: {}}
    sin_uei, etapa_rara, etapa_cambios, marcados_pc, sin_monto = [], [], [], [], []
    avisos = []

    for f in filas:
        ref = filas_proy.get(f['cui'])
        if ref is None:
            continue                      # CUI sin fila en Proyección26; se avisa aparte
        fila = ref['fila']

        monto = f.get('monto_final', f['monto_inversion'])
        if monto is None:
            if f['cui'] not in CUIS_SIN_AVISO:
                sin_monto.append(f['cui'])
        elif not igual_monto(ref['monto'], monto):
            por_fila[PROY_COL_MONTO][fila] = monto

        if f['resultado'] != 'ok':
            continue                      # sin ficha del SSI no hay UEI ni etapa que escribir

        sigla = codigo_uei(f.get('uei'))
        if sigla is None:
            if not vacio(f.get('uei')):
                sin_uei.append((f['cui'], str(f.get('uei'))[:60]))
        elif norm(ref['uei']) != norm(sigla):
            por_fila[PROY_COL_UEI][fila] = sigla

        etapa = etapa_desde_ssi(f)
        if etapa is None:
            if f['cui'] in set_ins:       # solo interesan los recién insertados
                etapa_rara.append((f['cui'], f.get('estado_ssi'), f.get('tipo_inversion')))
        elif norm(ref['etapa']) != norm(etapa):
            por_fila[PROY_COL_ETAPA][fila] = etapa
            etapa_cambios.append((f['cui'], str(ref['etapa']), etapa))
            if etapa == ETAPA_VALOR_PROCOMPITE:
                marcados_pc.append(f['cui'])

    if sin_monto:
        avisos.append(f'{len(sin_monto)} CUI sin monto nuevo (el SSI no los encontró o no respondió). '
                      f'Su celda de {col_letra(PROY_COL_MONTO)} quedó como estaba, no se pisó con 0:\n'
                      + '\n'.join(f'- {c}' for c in sin_monto))
    if sin_uei:
        avisos.append(f'{len(sin_uei)} UEI del SSI que no coinciden con ninguna sigla de "{t}". '
                      f'Esas celdas de {col_letra(PROY_COL_UEI)} NO se tocaron (hay que completarlas a mano '
                      'o agregar la equivalencia en UEI_EQUIVALENCIAS):\n'
                      + '\n'.join(f'- {c}: "{u}"' for c, u in sin_uei))
    if etapa_rara:
        avisos.append(f'{len(etapa_rara)} CUI recién insertados cuya Etapa no se puede deducir del SSI '
                      '(solo se saben deducir PROCOMPITE y desactivado). Su celda de '
                      f'{col_letra(PROY_COL_ETAPA)} quedó como estaba; hay que poner la etapa a mano:\n'
                      + '\n'.join(f'- {c}: estado "{e}", tipo "{ti}"' for c, e, ti in etapa_rara))
    if etapa_cambios:
        avisos.append(f'{len(etapa_cambios)} Etapa(s) actualizadas en "{t}" según el SSI:\n'
                      + '\n'.join(f'- {c}: "{v if v else "(vacía)"}" -> "{nv}"' for c, v, nv in etapa_cambios))
    if marcados_pc:
        avisos.append(f'OJO: {len(marcados_pc)} CUI quedaron marcados como "{ETAPA_VALOR_PROCOMPITE}" '
                      f'(PROCOMPITE). Desde la próxima corrida ya NO se consultan al SSI ni aparecen en '
                      f'{HOJA_PIM}, porque la Etapa "{ETAPA_EXCLUIDA}" los excluye:\n'
                      + '\n'.join(f'- {c}' for c in marcados_pc))

    datos = []
    for col, valores in por_fila.items():
        letra = col_letra(col)
        for ini, vals in tramos(valores):
            datos.append({'range': f"'{t}'!{letra}{ini}:{letra}{ini + len(vals) - 1}",
                          'values': [[v] for v in vals]})
    resumen = {'escrito': False, 'filas_en_proyeccion26': len(filas_proy),
               'montos_actualizados': len(por_fila[PROY_COL_MONTO]),
               'uei_actualizadas': len(por_fila[PROY_COL_UEI]),
               'etapas_actualizadas': len(por_fila[PROY_COL_ETAPA]),
               'uei_sin_equivalencia': len(sin_uei), 'rangos_enviados': len(datos)}
    if not datos:
        print('No había nada que actualizar en', t)
        return avisos, resumen

    # Antes de escribir: comprobar que la hoja sigue igual que cuando se leyó. El Apps Script
    # inserta filas en Proyección26, así que entre la lectura y esta escritura las filas pueden
    # haberse corrido; si eso pasara, cada monto caería en el CUI equivocado.
    filas_a_escribir = set()
    for valores in por_fila.values():
        filas_a_escribir |= set(valores)
    ult = max(filas_a_escribir)
    letra_cui, letra_ult = col_letra(PROY_COL_CUI), col_letra(PROY_COL_ULT)
    chk = libro.values_batch_get(
        [f"'{t}'!{letra_cui}{PROY_FILA_INI}:{letra_cui}{ult}",
         f"'{t}'!{letra_cui}{PROY_FILA_TITULOS}:{letra_ult}{PROY_FILA_TITULOS}"],
        params={'valueRenderOption': 'UNFORMATTED_VALUE'})['valueRanges']
    col_a = chk[0].get('values', [])
    tits = (chk[1].get('values') or [[]])[0]

    for col, esperado in ((PROY_COL_CUI, PROY_TIT_CUI), (PROY_COL_UEI, PROY_TIT_UEI),
                          (PROY_COL_ETAPA, PROY_TIT_ETAPA), (PROY_COL_MONTO, PROY_TIT_MONTO)):
        hallado = celda(tits, col - 1)
        if norm(hallado) != norm(esperado):
            avisos.append(f'NO se escribió nada en "{t}": justo antes de escribir, la celda '
                          f'{col_letra(col)}{PROY_FILA_TITULOS} dice "{hallado if not vacio(hallado) else "(vacía)"}" '
                          f'y debería decir "{esperado}". Alguien movió una columna mientras corría el programa.')
            resumen['motivo'] = 'los títulos cambiaron durante la corrida'
            return avisos, resumen

    por_numero = {r['fila']: cui for cui, r in filas_proy.items()}
    movidas = []
    for fila in sorted(filas_a_escribir):
        i = fila - PROY_FILA_INI
        actual = como_cui(celda(col_a[i], 0)) if 0 <= i < len(col_a) else None
        if actual != por_numero.get(fila):
            movidas.append(f'- fila {fila}: se esperaba el CUI {por_numero.get(fila)} y ahora hay '
                           f'{actual if actual is not None else "(vacío)"}')
    if movidas:
        avisos.append(f'NO se escribió nada en "{t}": las filas se corrieron entre la lectura y la '
                      'escritura (alguien insertó o borró filas, o corrió el guion de Google al mismo '
                      'tiempo). Escribir habría puesto cada monto en el CUI equivocado. '
                      'Volver a correr el programa.\n' + '\n'.join(movidas[:10]))
        resumen['motivo'] = 'las filas se corrieron durante la corrida'
        return avisos, resumen

    libro.values_batch_update({'valueInputOption': 'RAW', 'data': datos})
    resumen['escrito'] = True
    return avisos, resumen


# ---------------- principal ----------------
def main():
    if not tomar_bloqueo():
        return 0
    try:
        return corrida()
    finally:
        soltar_bloqueo()


def corrida():
    ahora = datetime.now()
    sello = ahora.strftime('%Y-%m-%d_%H%M%S')

    if '--probar-correo' in sys.argv:
        ok = enviar_correo('Correo de prueba', 'Prueba del aviso del programa cuis_desde_sheets.py.\n'
                           'Si llegó este correo, los avisos funcionan.\n\n' + ahora.strftime('%d/%m/%Y %H:%M:%S'))
        return 0 if ok else 4

    hp = hr = None
    try:
        libro, hp, hr, vr = leer()
        insertados, todos, obs, repetidos, resumen, fecha_pim, filas_proy = validar_y_armar(hp, hr, vr)
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
    f_var = f_json = None
    try:
        carpeta, aviso_carpeta = carpeta_montos()
        if aviso_carpeta:
            avisos.append(aviso_carpeta)
            asuntos.append('carpeta de montos no disponible')

        filas, res_ssi, no_existe, fallas = descargar_ssi(insertados, todos)

        # ---- comparación con la corrida anterior (antes de guardar, para fijar el monto final) ----
        ruta_ant, variaciones, ceros, avisos_base = comparar_con_anterior(filas)
        if avisos_base:
            avisos.extend(avisos_base)
            asuntos.append('problema con la base de comparación')
        grandes = [v for v in variaciones
                   if v['variacion_pct'] is None or abs(v['variacion_pct']) > UMBRAL_VARIACION_PCT]
        basura = len(ceros) >= MAX_CEROS_SOSPECHOSOS
        res_ssi['comparacion'] = {
            'corrida_anterior': ruta_ant.name if ruta_ant else None,
            'montos_que_cambiaron': len(variaciones),
            f'cambios_mayores_a_{UMBRAL_VARIACION_PCT}_pct': len(grandes),
            'ceros_sospechosos_se_mantuvo_el_anterior': len(ceros),
            'detenido_por_ceros_masivos': basura,
        }
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

        # el Excel es un informe aparte: si falla, no debe tumbar la corrida ni perder los montos
        if variaciones or ceros:
            try:
                f_var = excel_variaciones(variaciones, ceros, sello, ruta_ant, carpeta)
                print('Variaciones:', f_var)
            except Exception as e:
                avisos.append('No se pudo generar el Excel de variaciones (los montos sí se guardaron y la '
                              f'hoja sí se actualizó).\nDetalle: {type(e).__name__}: {e}')
                asuntos.append('falló el Excel de variaciones')

        f_json, f_csv = guardar_montos(filas, res_ssi, ahora, sello, archivo, carpeta)
        print('Montos:', f_json)
        print('        ', f_csv)
        print('Resumen SSI:', json.dumps(res_ssi, ensure_ascii=False))
    except Exception as e:
        filas, res_ssi, no_existe, fallas, basura = None, None, [], [], False
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
                          f'de {HOJA_PROY} sin PC. Se descargaron igual, pero no tienen fila donde escribir. '
                          'Puede que tengan Etapa "PC" o que su celda en Proyección26 tenga un error:\n'
                          + '\n'.join(f'- {c}' for c in fuera))
            asuntos.append('insertados fuera de la lista')

        # ---- paso 3: escribir en Proyección26 ----
        if todo_caido:
            avisos.append(f'No se escribió nada en "{hr.title}" porque el SSI no respondió; '
                          'la hoja se quedó con los datos de la corrida anterior.')
        elif basura:
            avisos.append(f'NO se escribió nada en "{hr.title}": {len(ceros)} CUI que ayer tenían monto '
                          f'llegaron hoy en 0, y el límite es {MAX_CEROS_SOSPECHOSOS}. Tantos ceros a la vez '
                          'no es un cambio real sino una falla de la página del MEF, así que no se copió esa '
                          'data basura. La hoja quedó con los montos de la corrida anterior. '
                          'Revisar el SSI y volver a correr el programa.')
            asuntos.append(f'DETENIDO: {len(ceros)} montos en 0 (falla del SSI)')
            codigo = 6
        else:
            try:
                av_hoja, res_hoja = escribir_proyeccion(libro, hr, filas, filas_proy, insertados)
                print('Hoja:', json.dumps(res_hoja, ensure_ascii=False))
                if av_hoja:
                    avisos.extend(av_hoja)
                    asuntos.append('avisos al escribir en ' + hr.title)
            except Exception as e:
                avisos.append(f'ERROR al escribir en "{hr.title}"; los montos sí quedaron guardados en los archivos.\n'
                              f'Detalle: {type(e).__name__}: {e}\n{traceback.format_exc()}')
                asuntos.append('error al escribir en ' + hr.title)
                codigo = 3

        if LIMPIAR_PIM_OPQ:
            try:
                limpiar_pim_opq(hp)
            except Exception as e:
                avisos.append(f'No se pudieron vaciar las columnas {PIM_RANGO_MUERTO} de "{hp.title}" '
                              f'(ya no se usan; si quedan montos viejos ahí pueden engañar a una fórmula).\n'
                              f'Detalle: {type(e).__name__}: {e}')
                asuntos.append('no se pudo limpiar ' + hp.title)

    if obs:
        lineas = '\n'.join(f'- Hoja "{o["hoja"]}", celda {o["celda"]}: "{o["valor"]}" -> {o["motivo"]}' for o in obs)
        avisos.insert(0, f'{len(obs)} celda(s) que no son CUI válidos (número de 7 cifras). Se dejaron fuera; '
                         'el resto se procesó normalmente:\n' + lineas)
        asuntos.insert(0, f'{len(obs)} valor(es) que no son CUI')

    if avisos:
        cuerpo = ('\n\n'.join(f'{i}) {a}' for i, a in enumerate(avisos, 1)) + '\n\n'
                  f'Lista de CUI: {archivo}\n'
                  + (f'Montos: {f_json}\n' if f_json else '')
                  + (f'Excel de variaciones: {f_var}\n' if f_var else '')
                  + f'Libro: {enlace()}\nHora: {ahora:%d/%m/%Y %H:%M:%S}')
        if not enviar_correo('; '.join(asuntos), cuerpo):
            return 4
    return codigo


if __name__ == '__main__':
    sys.exit(main())
