"""
Pruebas de cuis_desde_sheets.py SIN red, SIN libro de Google y SIN correo.
Simulan las lecturas de la hoja y las corridas anteriores en una carpeta temporal.

    python pruebas_sin_red.py      -> imprime cada caso y termina con "TODAS LAS PRUEBAS PASARON"

No cubren (necesitan red o credenciales): la consulta real al SSI (consultar_ssi), la lectura real
del libro (leer) y el envío de correo (enviar_correo).
"""
import json
import tempfile
from pathlib import Path

import cuis_desde_sheets as m

m.enviar_correo = lambda asunto, cuerpo: print('   (correo simulado)', asunto) or True


class Hoja:
    def __init__(self, titulo, gid=1):
        self.title, self.id = titulo, gid

    def format(self, *a, **k):
        pass


HP, HR = Hoja('PIM-diario', 37269580), Hoja('Proyección26', 1133143073)
OK_CAB = [[46349.3, '', 'CUI'], ['', '', 'x']]
OK_TIT = [['CUI', 'Inversión', 'ID', 'UEI', 'Modalidad', 'Etapa']]


def caso(nombre):
    print('==', nombre)


# ---------- paso 1: validar_y_armar ----------
def prueba_listas():
    datos = [[2000270, '', '', '', '', 'ET'], ['2000351', '', '', '', '', 'PC'], [2001621.0, '', '', '', '', ''],
             ['20016', '', '', '', '', 'ET'], ['abc', '', '', '', '', 'EF'], [], ['', '', '', '', '', 'L'],
             [2000270, '', '', '', '', 'EF'], [' 2064096 ', '', '', '', '', 'pc ']]
    col_a = [[2000270], ['2000351 '], ['hola'], [], [12345678]]

    caso('listas: valores raros, PC con espacios, repetidos, vacíos')
    ins, todos, obs, rep, res, _ = m.validar_y_armar(HP, HR, (OK_CAB, col_a, OK_TIT, datos))
    assert ins == [2000270, 2000351], ins
    assert todos == [2000270, 2001621], todos
    assert [o['celda'] for o in obs] == ['A5', 'A7', 'A6', 'A7', 'A9'], obs
    assert [(r['celda'], r['cui']) for r in rep] == [('A10', 2000270)], rep

    caso('listas: títulos con otra mayúscula y espacios se aceptan')
    m.validar_y_armar(HP, HR, (OK_CAB, col_a, [[' cui ', '', '', '', '', 'ETAPA ']], datos))

    for nombre, vr in [
        ('fila insertada arriba en Proyección26', (OK_CAB, col_a, [[]], datos)),
        ('columna Etapa movida', (OK_CAB, col_a, [['CUI', '', '', '', 'Etapa', 'Modalidad']], datos)),
        ('fila insertada arriba en PIM-diario', ([[], [46349.3, '', 'CUI']], col_a, OK_TIT, datos)),
        ('A2 de PIM-diario con un CUI', ([[46349.3, '', 'CUI'], [2000270]], col_a, OK_TIT, datos)),
    ]:
        caso('se detiene: ' + nombre)
        try:
            m.validar_y_armar(HP, HR, vr)
        except m.Detener as e:
            print('   ' + str(e).replace('\n', '\n   '))
        else:
            raise AssertionError('debió detenerse: ' + nombre)


# ---------- paso 3: escribir_pim ----------
class Libro:
    def __init__(self, tits, n, prev):
        self.r = [{'values': [tits]}, {'values': n}, {'values': prev}]
        self.escrito = None

    def values_batch_get(self, rangos, params=None):
        return {'valueRanges': self.r}

    def values_batch_update(self, cuerpo):
        self.escrito = {d['range'].split('!')[1]: [x[0] for x in d['values']] for d in cuerpo['data']}


def prueba_escritura():
    ok = ['CUI_proyeccion26_noPC', 'COSTO_TOTAL_INVERSION_ACTUALIZADO (a+b+c+d)', 'UEI', 'TIPO_INVERSION']
    filas = [{'cui': 2064096, 'monto_inversion': 100.5},
             {'cui': 2410530, 'monto_inversion': 7.0, 'uei': 'UEI X', 'tipo_inversion': 'PROYECTO'},
             {'cui': 3999999, 'monto_inversion': None, 'uei': None, 'tipo_inversion': None},
             {'cui': 2094012, 'monto_inversion': None},
             {'cui': 2158997, 'monto_inversion': 0.0, 'monto_final': 999.0}]   # cero sospechoso ya resuelto
    n = [[2064096], [2410530], [3999999], [2094012], [2158997]]
    prev = [[1], [2], [3, '', ''], [555.5], [999.0], ['viejo', 'a', 'b'], ['x']]
    insertados = [2410530, 2173274]

    caso('escritura: títulos correctos')
    lib = Libro(ok, n, prev)
    av, res = m.escribir_pim(lib, HP, filas, insertados)
    assert lib.escrito['O2:O8'] == [100.5, 7.0, 3, 555.5, 999.0, '', ''], lib.escrito
    assert lib.escrito['P2:P8'] == ['', 'UEI X', '', '', '', '', '']
    assert lib.escrito['Q2:Q8'] == ['', 'PROYECTO', '', '', '', '', '']
    assert any('N5 = 2094012' in a for a in av) and not any('3999999' in a for a in av), av
    assert any('2173274' in a for a in av), av

    caso('escritura: P1 cambiado -> no escribe P, sí O y Q')
    lib = Libro(ok[:2] + ['U.E.I.'] + ok[3:], n, prev)
    av, res = m.escribir_pim(lib, HP, filas, insertados)
    assert set(lib.escrito) == {'O2:O8', 'Q2:Q8'}, lib.escrito
    assert any('columna P' in a for a in av)

    caso('escritura: N1 cambiado -> no escribe nada')
    lib = Libro(['CUI'] + ok[1:], n, prev)
    av, res = m.escribir_pim(lib, HP, filas, insertados)
    assert lib.escrito is None and res == {'escrito': False}


# ---------- comparación con la corrida anterior ----------
def prueba_comparacion():
    with tempfile.TemporaryDirectory() as tmp:
        m.CARPETA_MONTOS = Path(tmp)

        def guardar(nombre, datos):
            (Path(tmp) / nombre).write_text(json.dumps({'datos': datos}), encoding='utf-8')

        caso('comparación: sin corrida anterior')
        filas = [{'cui': 1, 'monto_inversion': 10.0}]
        ruta, var, ceros = m.comparar_con_anterior(filas)
        assert ruta is None and var == [] and ceros == [] and filas[0]['monto_final'] == 10.0

        caso('comparación día 1: alzas, bajas, cero sospechoso, sin dato, antes 0, igual, nuevo')
        guardar('monto_ssi_2026-11-24_064500.json',
                [{'cui': 1, 'monto_inversion': 100.0}, {'cui': 2, 'monto_inversion': 200.0},
                 {'cui': 3, 'monto_inversion': 300.0}, {'cui': 4, 'monto_inversion': 400.0},
                 {'cui': 5, 'monto_inversion': 0.0}, {'cui': 6, 'monto_inversion': 600.0},
                 {'cui': 8, 'monto_inversion': None}])
        hoy = [{'cui': 1, 'monto_inversion': 105.0, 'nombre_inversion': 'SUBE 5'},
               {'cui': 2, 'monto_inversion': 170.0, 'nombre_inversion': 'BAJA 15'},
               {'cui': 3, 'monto_inversion': 0.0, 'nombre_inversion': 'CAE A CERO', 'estado_ssi': 'ACTIVO'},
               {'cui': 4, 'monto_inversion': None, 'nombre_inversion': ''},
               {'cui': 5, 'monto_inversion': 50.0, 'nombre_inversion': 'ANTES CERO'},
               {'cui': 6, 'monto_inversion': 600.0, 'nombre_inversion': 'IGUAL'},
               {'cui': 7, 'monto_inversion': 70.0, 'nombre_inversion': 'NUEVO'},
               {'cui': 8, 'monto_inversion': 80.0, 'nombre_inversion': 'ANTES SIN DATO'}]
        ruta, var, ceros = m.comparar_con_anterior(hoy)
        finales = {f['cui']: f['monto_final'] for f in hoy}
        assert finales == {1: 105.0, 2: 170.0, 3: 300.0, 4: 400.0, 5: 50.0, 6: 600.0, 7: 70.0, 8: 80.0}, finales
        assert [(v['cui'], v['variacion_pct']) for v in var] == [(5, None), (2, -15.0), (1, 5.0)], var
        assert [z['cui'] for z in ceros] == [3], ceros
        x = m.excel_variaciones(var, ceros, 'prueba', ruta)
        from openpyxl import load_workbook
        wb = load_workbook(x)
        assert wb.sheetnames == ['Variaciones', 'Ceros sospechosos', 'Nota']
        assert wb['Variaciones']['E3'].value == -0.15 and wb['Variaciones']['E3'].number_format == '0.00%'
        print('   Excel', x.name, 'con hojas', wb.sheetnames)

        caso('comparación día 2: el CUI 3 sigue en 0 -> se mantiene 300 SIN alerta')
        guardar('monto_ssi_2026-11-25_064500.json', hoy)
        dia2 = [{'cui': 3, 'monto_inversion': 0.0, 'nombre_inversion': 'CAE A CERO'}]
        ruta, var, ceros = m.comparar_con_anterior(dia2)
        assert ruta.name == 'monto_ssi_2026-11-25_064500.json'
        assert dia2[0]['monto_final'] == 300.0 and ceros == [] and var == [], (dia2, ceros, var)
        print('   observación:', dia2[0]['observacion'])

        caso('comparación día 3: el CUI 3 vuelve con 310 -> variación contra 300')
        guardar('monto_ssi_2026-11-26_064500.json', dia2)
        dia3 = [{'cui': 3, 'monto_inversion': 310.0, 'nombre_inversion': 'CAE A CERO'}]
        ruta, var, ceros = m.comparar_con_anterior(dia3)
        assert [(v['cui'], v['monto_anterior'], v['variacion_pct']) for v in var] == [(3, 300.0, 3.33)], var

        caso('comparación: un cero nuevo en otro CUI sí alerta aunque otro siga en 0')
        guardar('monto_ssi_2026-11-27_064500.json', dia2 + [{'cui': 9, 'monto_inversion': 90.0}])
        dia4 = [{'cui': 3, 'monto_inversion': 0.0}, {'cui': 9, 'monto_inversion': 0.0}]
        ruta, var, ceros = m.comparar_con_anterior(dia4)
        assert [z['cui'] for z in ceros] == [9] and {f['cui']: f['monto_final'] for f in dia4} == {3: 300.0, 9: 90.0}


if __name__ == '__main__':
    prueba_listas()
    prueba_escritura()
    prueba_comparacion()
    print('\nTODAS LAS PRUEBAS PASARON')
