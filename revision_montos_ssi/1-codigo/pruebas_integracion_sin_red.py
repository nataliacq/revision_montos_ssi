"""
Corrida completa simulada: leer -> SSI -> comparar -> guardar -> escribir Proyección26.
No usa red, ni libro de Google, ni correo: todo está simulado en memoria y en una carpeta temporal.

    python pruebas_integracion_sin_red.py

Complementa a pruebas_sin_red.py, que prueba cada función por separado: esta comprueba el cableado
de main() de punta a punta (bloqueo, freno por ceros masivos, limpieza de PIM-diario, correo).
"""
import copy
import json
import tempfile
from datetime import date, timedelta
from pathlib import Path

import cuis_desde_sheets as m

correos = []
m.enviar_correo = lambda a, c: correos.append((a, c)) or True


class Hoja:
    def __init__(s, t, gid):
        s.title, s.id, s.limpiados = t, gid, []

    def format(s, *a, **k):
        pass

    def batch_clear(s, r):
        s.limpiados += list(r)


HP, HR = Hoja('PIM-diario', 1), Hoja('Proyección26', 2)


class Libro:
    title = 'Copia de Drives exclusivos 19.11'

    def __init__(s):
        s.escrito = None
        s.correr_filas = False      # True simula que alguien insertó una fila mientras corría

    def values_batch_get(s, rangos, params=None):
        # la comprobación previa a escribir: columna A desde la fila 3, y los títulos
        cuis = ([[9999999]] if s.correr_filas else []) + [[c] for c in CUIS]
        return {'valueRanges': [{'values': cuis}, {'values': [TITULOS]}]}

    def values_batch_update(s, cuerpo):
        s.escrito = {d['range'].split('!')[1]: [x[0] for x in d['values']] for d in cuerpo['data']}


CUIS = [2064096, 2410530, 2094012, 2158997, 3000882]
TITULOS = ['CUI', 'Inversión', 'ID', 'UEI', 'Modalidad', 'Etapa', '', '', '(A) Monto de inversión']
# Proyección26 A3:I7  (CUI, Inversión, ID, UEI, Modalidad, Etapa, -, -, monto)
DATOS_PROY = [
    [2064096, 'x', '', '',    '', 'ET', '', '', ''],      # fila vieja, sin UEI y sin monto
    [2410530, 'x', '', 'GRI', '', 'ET', '', '', 7.0],     # fila vieja, ya completa
    [2094012, 'x', '', '',    '', '',   '', '', 80.0],    # RECIÉN INSERTADO
    [2158997, 'x', '', '',    '', 'ET', '', '', 10.0],    # fila vieja, el SSI dice PROCOMPITE
    [3000882, 'x', '', '',    '', 'ET', '', '', 50.0],    # no existe en el SSI (CUIS_SIN_AVISO)
]
CAB_PIM = [[46349.3, '', 'CUI'], ['', '', 'x']]
# El histórico acumulado de PIM-diario: CUI en A y su fecha de inserción en B. No se borra; el
# archivo insertados_historico.csv es lo que dice cuál ya se procesó.
COL_A_PIM = [[2064096, '2026-11-02'],    # ya procesado antes (está en el CSV de abajo)
             [2094012, '2026-11-25']]    # este es el nuevo de esta corrida

GRI = 'GERENCIA REGIONAL DE INFRAESTRUCTURA'
PET = 'PROYECTO ESPECIAL AFIANZAMIENTO Y AMPLIACION DE LOS RECURSOS HIDRICOS DE TACNA'
FICHAS_BASE = {
    2064096: {'NOMBRE_INVERSION': 'MEJORAMIENTO VIAL', 'ESTADO': 'ACTIVO', 'MTO_F8_C': 100.5,
              'DES_UNIDAD_UEI': GRI, 'TIPO_FORMATO': 'PROYECTO'},
    2410530: {'NOMBRE_INVERSION': 'COLEGIO', 'ESTADO': 'ACTIVO', 'MTO_F8_C': 7.0,
              'DES_UNIDAD_UEI': GRI, 'TIPO_FORMATO': 'PROYECTO'},
    2094012: {'NOMBRE_INVERSION': 'PARQUE', 'ESTADO': 'DESACTIVADO TEMPORAL', 'MTO_F8_C': 80.0,
              'DES_UNIDAD_UEI': 'MUNICIPALIDAD DISTRITAL DE POCOLLAY', 'TIPO_FORMATO': 'PROYECTO'},
    2158997: {'NOMBRE_INVERSION': 'RIEGO', 'ESTADO': 'ACTIVO', 'MTO_F8_C': 900.0,
              'DES_UNIDAD_UEI': PET, 'TIPO_FORMATO': 'PROYECTO PROCOMPITE'},
}


def responder(fichas, caidos=()):
    def f(cui):
        if cui in caidos:
            return 'falla', 'el servicio no respondió a tiempo'
        return ('ok', fichas[cui]) if cui in fichas else ('no_existe', None)
    return f


def titulo(x):
    print('\n' + '=' * 78 + '\n' + x + '\n' + '=' * 78)


def ultimo_correo():
    return correos[-1] if correos else ('(ninguno)', '')


with tempfile.TemporaryDirectory() as tmp:
    t = Path(tmp)
    m.CARPETA_SALIDA, m.CARPETA_MONTOS, m.CARPETA_RESPALDO = t / 'cuis', t / 'montos', t / 'bajados'
    m.ARCHIVO_BLOQUEO = t / 'corrida.lock'
    m.CARPETA_HISTORICO = t / 'historico'
    m.ARCHIVO_HISTORICO = m.CARPETA_HISTORICO / 'insertados_historico.csv'
    m.CARPETA_HISTORICO.mkdir(parents=True)
    m.ARCHIVO_HISTORICO.write_text(
        'fecha_corrida;cui;fecha_insercion;fila_proyeccion26;estado_ssi;tipo_inversion;uei_ssi;'
        'uei_escrita;etapa_escrita;monto;resultado;falta_completar_a_mano\n'
        '2026-11-02 06:45:00;2064096;2026-11-02;3;ACTIVO;PROYECTO;;GRI;;100,50;ok;\n',
        encoding='utf-8-sig')
    lib = Libro()
    m.leer = lambda: (lib, HP, HR, (CAB_PIM, COL_A_PIM, [TITULOS], DATOS_PROY))
    m.consultar_ssi = responder(FICHAS_BASE)

    titulo('1) corrida normal (sin corrida anterior)')
    codigo = m.corrida()
    print('código de salida:', codigo)
    print('enviado a Proyección26:', json.dumps(lib.escrito, ensure_ascii=False))
    # El monto va a todas las filas que cambian: 2064096 (fila 3) y 2158997 (fila 6).
    # 2410530 y 2094012 ya tenían su monto; 3000882 no tiene dato y no se toca.
    assert lib.escrito['I3:I3'] == [100.5], lib.escrito
    assert lib.escrito['I6:I6'] == [900.0], lib.escrito
    # D y F solo en la fila del recién insertado (2094012, fila 5): su UEI no tiene equivalencia,
    # así que solo recibe la Etapa.
    assert lib.escrito['F5:F5'] == ['Desactivado'], lib.escrito
    assert not any(r.startswith('D') for r in lib.escrito), 'la UEI no tenía equivalencia'
    # la fila vieja 3 tiene UEI vacía y el SSI da GRI, pero NO se escribe: no es recién insertada
    assert not any(r.startswith('D3') or r.startswith('F3') for r in lib.escrito), lib.escrito
    # 2158997 es PROCOMPITE pero es fila vieja: no se le escribe la Etapa, se reporta
    assert not any(r.startswith('F6') for r in lib.escrito), lib.escrito
    asunto, cuerpo = ultimo_correo()
    assert 'reporta como PROCOMPITE' in cuerpo and '2158997' in cuerpo, cuerpo
    assert 'POCOLLAY' in cuerpo, cuerpo
    print('PIM-diario limpiado:', HP.limpiados)
    # la columna A NO se borra: solo se vacía O2:Q, que ya no se usa
    assert HP.limpiados == ['O2:Q'], HP.limpiados
    # 2064096 ya estaba en el CSV: no se reprocesa. Solo se anota el nuevo.
    hist = m.ARCHIVO_HISTORICO.read_text(encoding='utf-8-sig').splitlines()
    assert len(hist) == 3, hist          # encabezado + el ya procesado + el nuevo de hoy
    assert hist[2].split(';')[1] == '2094012' and hist[2].split(';')[2] == '2026-11-25', hist[2]
    assert 'UEI' in hist[2].split(';')[-1], hist[2]   # su UEI no tuvo equivalencia
    print('histórico (línea nueva):', hist[2][:120])
    # 2064096 es fila vieja ya procesada: no recibe D ni F aunque su UEI del SSI sí tenga sigla
    assert not any(r.startswith('D3') or r.startswith('F3') for r in lib.escrito), lib.escrito
    print('correo ->', asunto)
    for linea in cuerpo.split('\n')[:10]:
        print('   ', linea[:110])

    titulo('2) bloqueo: ya hay una corrida en curso')
    m.ARCHIVO_BLOQUEO.write_text('otra', encoding='utf-8')
    lib.escrito = None
    assert m.main() == 0 and lib.escrito is None
    print('código de salida: 0, no hizo nada')
    m.ARCHIVO_BLOQUEO.unlink()

    # la corrida 1 quedó guardada con la fecha de hoy; se renombra a ayer para que sirva de base
    viejo = next((t / 'montos').glob('monto_ssi_*.json'))
    ayer = (date.today() - timedelta(days=1)).strftime('%Y-%m-%d')
    base = viejo.parent / f'monto_ssi_{ayer}_064500.json'
    viejo.rename(base)

    titulo('3) un recién insertado se queda sin ficha del SSI (se procesa una sola vez)')
    # 2094012 ya quedó anotado en la corrida 1, así que para tener un insertado nuevo se agrega
    # otro al histórico de la hoja, como lo haría el guion de Google
    COL_A_PIM.append([2410530, '2026-11-26'])
    lib.escrito = None
    m.consultar_ssi = responder(FICHAS_BASE, caidos={2410530})
    codigo = m.corrida()
    asunto, cuerpo = ultimo_correo()
    print('código de salida:', codigo)
    print('enviado a Proyección26:', lib.escrito)
    assert not any(r.startswith('D4') or r.startswith('F4') for r in (lib.escrito or {})), lib.escrito
    assert 'SIN UEI ni Etapa' in cuerpo and '2410530' in cuerpo, cuerpo
    # queda anotado como procesado, con la causa, así que no se vuelve a intentar
    hist = m.ARCHIVO_HISTORICO.read_text(encoding='utf-8-sig').splitlines()
    assert hist[-1].split(';')[1] == '2410530' and 'no dio su ficha' in hist[-1], hist[-1]
    aviso = [b for b in cuerpo.split('\n\n') if 'SIN UEI ni Etapa' in b][0]
    print('correo ->', asunto)
    print('   ', aviso.replace('\n', ' | ')[:230])

    titulo('4) las filas se corrieron: no se escribe y el CUI nuevo no se anota')
    # un insertado nuevo, para comprobar que una escritura abortada NO lo marca como procesado
    COL_A_PIM.append([3000882, '2026-11-26'])
    lineas_antes = len(m.ARCHIVO_HISTORICO.read_text(encoding='utf-8-sig').splitlines())
    lib.correr_filas = True
    lib.escrito = None
    HP.limpiados = []
    codigo = m.corrida()
    asunto, cuerpo = ultimo_correo()
    print('código de salida:', codigo)
    assert lib.escrito is None, 'no se debe escribir nada'
    assert 'NO se procesaron' in cuerpo, cuerpo
    # no se anotó nada: el CUI sigue pendiente y se procesa en la próxima corrida
    assert len(m.ARCHIVO_HISTORICO.read_text(encoding='utf-8-sig').splitlines()) == lineas_antes
    print('correo ->', asunto)
    print('   ', [b for b in cuerpo.split('\n\n') if 'NO se procesaron' in b][0][:200])
    lib.correr_filas = False

    titulo('5) el SSI trae 0 en todo: freno por ceros masivos')
    ceros = copy.deepcopy(FICHAS_BASE)
    for f in ceros.values():
        f['MTO_F8_C'] = 0
    m.consultar_ssi = responder(ceros)
    m.MAX_CEROS_SOSPECHOSOS = 3
    lib.escrito = None
    HP.limpiados = []
    for p in (t / 'montos').glob('monto_ssi_*.json'):
        if p != base:
            p.unlink()                   # dejar solo la base de ayer
    codigo = m.corrida()
    asunto, cuerpo = ultimo_correo()
    print('código de salida:', codigo, '(6 = detenido por ceros masivos)')
    assert codigo == 6, codigo
    assert lib.escrito is None, 'no se debe escribir nada en la hoja'
    assert 'DETENIDO' in asunto, asunto
    assert len(m.ARCHIVO_HISTORICO.read_text(encoding='utf-8-sig').splitlines()) == lineas_antes, \
        'tampoco se anota cuando se frena por ceros masivos'
    print('enviado a Proyección26:', lib.escrito, '<- nada')
    print('correo ->', asunto)

print('\nTODAS LAS PRUEBAS DE INTEGRACIÓN PASARON')
