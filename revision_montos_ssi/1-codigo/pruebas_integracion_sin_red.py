"""
Corrida completa simulada: leer -> SSI -> comparar -> guardar -> escribir Proyección26.
No usa red, ni libro de Google, ni correo: todo está simulado en memoria y en una carpeta temporal.

    python pruebas_integracion_sin_red.py

Complementa a pruebas_sin_red.py, que prueba cada función por separado: esta comprueba el cableado
de main() de punta a punta (bloqueo, freno por ceros masivos, limpieza de PIM-diario, correo).
"""
import json
import tempfile
from pathlib import Path

import cuis_desde_sheets as m

correos = []
m.enviar_correo = lambda a, c: correos.append((a, c)) or True

class Hoja:
    def __init__(s, t, gid): s.title, s.id, s.limpiado = t, gid, None
    def format(s, *a, **k): pass
    def batch_clear(s, r): s.limpiado = r

HP, HR = Hoja('PIM-diario', 1), Hoja('Proyección26', 2)

class Libro:
    title = 'Copia de Drives exclusivos 19.11'
    def __init__(s): s.escrito = None
    def values_batch_get(s, rangos, params=None):
        # la comprobación previa a escribir: columna A desde la fila 3, y los títulos
        return {'valueRanges': [{'values': [[c] for c in CUIS]},
                                {'values': [TITULOS]}]}
    def values_batch_update(s, cuerpo):
        s.escrito = {d['range'].split('!')[1]: [x[0] for x in d['values']] for d in cuerpo['data']}

CUIS = [2064096, 2410530, 2094012, 2158997, 3000882]
TITULOS = ['CUI', 'Inversión', 'ID', 'UEI', 'Modalidad', 'Etapa', '', '', '(A) Monto de inversión']
# Proyección26 A3:I7  (CUI, Inv, ID, UEI, Modalidad, Etapa, -, -, monto)
DATOS_PROY = [
    [2064096, 'x', '', '',    '', 'ET', '', '', ''],
    [2410530, 'x', '', 'GRI', '', 'ET', '', '', 7.0],
    [2094012, 'x', '', '',    '', 'EF', '', '', 80.0],
    [2158997, 'x', '', '',    '', 'ET', '', '', 10.0],
    [3000882, 'x', '', '',    '', 'ET', '', '', 50.0],
]
CAB_PIM = [[46349.3, '', 'CUI'], ['', '', 'x']]
COL_A_PIM = [[2094012]]          # un recién insertado

FICHAS = {
    2064096: {'NOMBRE_INVERSION': 'MEJORAMIENTO VIAL', 'ESTADO': 'ACTIVO', 'MTO_F8_C': 100.5,
              'DES_UNIDAD_UEI': 'GERENCIA REGIONAL DE INFRAESTRUCTURA', 'TIPO_FORMATO': 'PROYECTO'},
    2410530: {'NOMBRE_INVERSION': 'COLEGIO', 'ESTADO': 'ACTIVO', 'MTO_F8_C': 7.0,
              'DES_UNIDAD_UEI': 'GERENCIA REGIONAL DE INFRAESTRUCTURA', 'TIPO_FORMATO': 'PROYECTO'},
    2094012: {'NOMBRE_INVERSION': 'PARQUE', 'ESTADO': 'DESACTIVADO TEMPORAL', 'MTO_F8_C': 80.0,
              'DES_UNIDAD_UEI': 'MUNICIPALIDAD DISTRITAL DE POCOLLAY', 'TIPO_FORMATO': 'PROYECTO'},
    2158997: {'NOMBRE_INVERSION': 'RIEGO', 'ESTADO': 'ACTIVO', 'MTO_F8_C': 900.0,
              'DES_UNIDAD_UEI': 'PROYECTO ESPECIAL AFIANZAMIENTO Y AMPLIACION DE LOS RECURSOS HIDRICOS DE TACNA',
              'TIPO_FORMATO': 'PROYECTO PROCOMPITE'},
}
m.consultar_ssi = lambda cui: ('ok', FICHAS[cui]) if cui in FICHAS else ('no_existe', None)

with tempfile.TemporaryDirectory() as tmp:
    t = Path(tmp)
    m.CARPETA_SALIDA, m.CARPETA_MONTOS, m.CARPETA_RESPALDO = t/'cuis', t/'montos', t/'bajados'
    m.ARCHIVO_BLOQUEO = t / 'corrida.lock'
    lib = Libro()
    m.leer = lambda: (lib, HP, HR, (CAB_PIM, COL_A_PIM, [TITULOS], DATOS_PROY))

    print('=== corrida 1 (sin corrida anterior) ===')
    codigo = m.main()
    print('código de salida:', codigo)
    print('enviado a Proyección26:', json.dumps(lib.escrito, ensure_ascii=False))
    print('PIM-diario limpiado:', HP.limpiado)
    print('archivos:', sorted(p.name for p in (t/'montos').glob('*')))
    print('\ncorreo ->', correos[-1][0])
    for linea in correos[-1][1].split('\n')[:14]:
        print('   ', linea)

    print('\n=== corrida 2 (bloqueo: ya hay una en curso) ===')
    m.ARCHIVO_BLOQUEO.write_text('otra', encoding='utf-8')
    print('código de salida:', m.main(), '(no hizo nada)')
    m.ARCHIVO_BLOQUEO.unlink()

    print('\n=== corrida 3 (el SSI trae 0 en todos: freno por basura) ===')
    for f in FICHAS.values():
        f['MTO_F8_C'] = 0
    # la base es la corrida 1, pero del mismo día: se renombra a ayer para que sirva de base
    from datetime import date, timedelta
    viejo = next((t/'montos').glob('monto_ssi_*.json'))
    ayer = (date.today() - timedelta(days=1)).strftime('%Y-%m-%d')
    viejo.rename(viejo.parent / f'monto_ssi_{ayer}_064500.json')
    lib.escrito = None
    m.MAX_CEROS_SOSPECHOSOS = 3
    codigo = m.main()
    print('código de salida:', codigo, '(6 = detenido por ceros masivos)')
    print('enviado a Proyección26:', lib.escrito, '<- no se escribió nada')
    print('correo ->', correos[-1][0])
