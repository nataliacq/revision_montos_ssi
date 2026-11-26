"""
Pruebas de cuis_desde_sheets.py SIN red, SIN libro de Google y SIN correo.
Simulan las lecturas de la hoja y las corridas anteriores en una carpeta temporal.

    python pruebas_sin_red.py      -> imprime cada caso y termina con "TODAS LAS PRUEBAS PASARON"

No cubren (necesitan red o credenciales): la consulta real al SSI (consultar_ssi), la lectura real
del libro (leer) y el envío de correo (enviar_correo).
"""
import json
import tempfile
from datetime import date, timedelta
from pathlib import Path

import cuis_desde_sheets as m

m.enviar_correo = lambda asunto, cuerpo: print('   (correo simulado)', asunto) or True


class Hoja:
    def __init__(self, titulo, gid=1):
        self.title, self.id = titulo, gid
        self.limpiado = None

    def format(self, *a, **k):
        pass

    def batch_clear(self, rangos):
        self.limpiado = rangos


HP, HR = Hoja('PIM-diario', 37269580), Hoja('Proyección26', 1133143073)
OK_CAB = [[46349.3, '', 'CUI'], ['', '', 'x']]
# Títulos de Proyección26 (fila 2), de A a I: el programa exige CUI, UEI, Etapa y el monto
OK_TIT = [['CUI', 'Inversión', 'ID', 'UEI', 'Modalidad', 'Etapa', 'g', 'h', '(A) Monto de inversión']]


def caso(nombre):
    print('==', nombre)


# ---------- paso 1: validar_y_armar ----------
def prueba_listas():
    datos = [[2000270, '', '', 'GRI', '', 'ET', '', '', 100.0],
             ['2000351', '', '', '', '', 'PC'], [2001621.0, '', '', '', '', ''],
             ['20016', '', '', '', '', 'ET'], ['abc', '', '', '', '', 'EF'], [], ['', '', '', '', '', 'L'],
             [2000270, '', '', '', '', 'EF'], [' 2064096 ', '', '', '', '', 'pc ']]
    col_a = [[2000270], ['2000351 '], ['hola'], [], [12345678]]

    caso('listas: valores raros, PC con espacios, repetidos, vacíos')
    ins, todos, obs, rep, res, _, filas_proy = m.validar_y_armar(HP, HR, (OK_CAB, col_a, OK_TIT, datos))
    assert ins == [2000270, 2000351], ins
    assert todos == [2000270, 2001621], todos
    assert [o['celda'] for o in obs] == ['A5', 'A7', 'A6', 'A7', 'A9'], obs
    assert [(r['celda'], r['cui']) for r in rep] == [('A10', 2000270)], rep

    caso('listas: se guarda la fila y el contenido actual de D, F e I de cada CUI')
    assert filas_proy[2000270] == {'fila': 3, 'uei': 'GRI', 'etapa': 'ET', 'monto': 100.0}, filas_proy[2000270]
    assert filas_proy[2001621] == {'fila': 5, 'uei': '', 'etapa': '', 'monto': ''}, filas_proy[2001621]
    assert 2000351 not in filas_proy, 'los PC no se escriben'

    caso('listas: títulos con otra mayúscula y espacios se aceptan')
    m.validar_y_armar(HP, HR, (OK_CAB, col_a,
                               [[' cui ', 'Inversión', '', 'uei ', '', 'ETAPA ', '', '',
                                 '(A) MONTO DE INVERSION']], datos))

    for nombre, vr in [
        ('fila insertada arriba en Proyección26', (OK_CAB, col_a, [[]], datos)),
        ('columna Etapa movida', (OK_CAB, col_a, [['CUI', '', '', 'UEI', 'Etapa', 'Modalidad']], datos)),
        ('falta el título de UEI en D2', (OK_CAB, col_a,
                                         [['CUI', 'Inversión', 'ID', '', '', 'Etapa', '', '',
                                           '(A) Monto de inversión']], datos)),
        ('falta el título del monto en I2', (OK_CAB, col_a,
                                            [['CUI', 'Inversión', 'ID', 'UEI', '', 'Etapa']], datos)),
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


# ---------- equivalencias SSI -> Proyección26 ----------
def prueba_equivalencias():
    caso('UEI: cada nombre largo del SSI a su sigla')
    pares = [
        ('GOBIERNO REGIONAL DE TACNA - GERENCIA REGIONAL DE INFRAESTRUCTURA', 'GRI'),
        ('Gerencia Regional de Desarrollo Económico', 'GRDE'),
        ('GERENCIA REGIONAL DE DESARROLLO SOCIAL', 'GRDIS'),
        ('Gerencia Regional de Desarrollo e Inclusion Social', 'GRDIS'),
        ('GERENCIA REGIONAL DE RECURSOS NATURALES Y GESTION DEL MEDIO AMBIENTE', 'GRRNyGA'),
        ('PROYECTO ESPECIAL AFIANZAMIENTO Y AMPLIACION DE LOS RECURSOS HIDRICOS DE TACNA', 'PET'),
        ('gerencia de infraestructura del proyecto especial afianzamiento', 'PET'),
        ('DIRECCION REGIONAL DE SALUD', 'DIRESA'),
        ('Dirección Regional de Agricultura', 'DRAT'),
        ('DIRECCION REGIONAL DE TRANSPORTES Y COMUNICACIONES', 'DRTYC'),
    ]
    for texto, sigla in pares:
        got = m.codigo_uei(texto)
        assert got == sigla, f'{texto!r} -> {got!r}, se esperaba {sigla!r}'
        print(f'   {sigla:8} <- {texto[:58]}')

    caso('UEI: sin equivalencia devuelve None (la celda no se toca y se avisa)')
    for texto in ['MUNICIPALIDAD DISTRITAL DE POCOLLAY', '', None, 'otra cosa']:
        assert m.codigo_uei(texto) is None, texto

    caso('Etapa: solo PROCOMPITE y desactivado se saben deducir')
    f = lambda est, tipo: {'estado_ssi': est, 'tipo_inversion': tipo}
    assert m.etapa_desde_ssi(f('ACTIVO', 'PROYECTO PROCOMPITE')) == 'PC'
    assert m.etapa_desde_ssi(f('DESACTIVADO TEMPORAL', 'PROYECTO')) == 'Desactivado'
    assert m.etapa_desde_ssi(f('DESACTIVADO PERMANENTE', 'IOARR')) == 'Desactivado'
    assert m.etapa_desde_ssi(f('ACTIVO', 'PROYECTO')) is None
    assert m.etapa_desde_ssi(f('CERRADO', 'IOARR')) is None
    assert m.etapa_desde_ssi(f('NO REGISTRADO EN EL INVIERTE', 'PROYECTO GENERICO')) is None

    caso('tramos: solo filas consecutivas, las ausentes no se escriben')
    assert m.tramos({3: 'a', 4: 'b', 9: 'c'}) == [(3, ['a', 'b']), (9, ['c'])]
    assert m.tramos({}) == []


# ---------- paso 3: escribir_proyeccion ----------
class Libro:
    """Devuelve la columna A y los títulos que ve la comprobación previa a la escritura."""

    def __init__(self, col_a, tits=None):
        self.col_a, self.tits = col_a, tits or OK_TIT[0]
        self.escrito = None

    def values_batch_get(self, rangos, params=None):
        return {'valueRanges': [{'values': self.col_a}, {'values': [self.tits]}]}

    def values_batch_update(self, cuerpo):
        assert cuerpo['valueInputOption'] == 'RAW'
        self.escrito = {d['range'].split('!')[1]: [x[0] for x in d['values']] for d in cuerpo['data']}


def prueba_escritura():
    # filas 3 a 7 de Proyección26
    filas_proy = {
        2064096: {'fila': 3, 'uei': '', 'etapa': 'ET', 'monto': ''},          # todo nuevo
        2410530: {'fila': 4, 'uei': 'GRI', 'etapa': 'ET', 'monto': 7.0},      # ya está igual
        3999999: {'fila': 5, 'uei': '', 'etapa': '', 'monto': 500.0},         # el SSI no lo encuentra
        2094012: {'fila': 6, 'uei': '', 'etapa': 'EF', 'monto': 80.0},        # UEI sin equivalencia
        2158997: {'fila': 7, 'uei': '', 'etapa': 'ET', 'monto': 10.0},        # PROCOMPITE
    }
    col_a = [[2064096], [2410530], [3999999], [2094012], [2158997]]
    filas = [
        {'cui': 2064096, 'resultado': 'ok', 'monto_inversion': 100.5, 'monto_final': 100.5,
         'uei': 'GERENCIA REGIONAL DE INFRAESTRUCTURA', 'tipo_inversion': 'PROYECTO', 'estado_ssi': 'ACTIVO'},
        {'cui': 2410530, 'resultado': 'ok', 'monto_inversion': 7.0, 'monto_final': 7.0,
         'uei': 'GERENCIA REGIONAL DE INFRAESTRUCTURA', 'tipo_inversion': 'PROYECTO', 'estado_ssi': 'ACTIVO'},
        {'cui': 3999999, 'resultado': 'el SSI no encuentra este CUI', 'monto_inversion': None,
         'monto_final': None, 'uei': None, 'tipo_inversion': None, 'estado_ssi': ''},
        {'cui': 2094012, 'resultado': 'ok', 'monto_inversion': 80.0, 'monto_final': 80.0,
         'uei': 'MUNICIPALIDAD DISTRITAL DE POCOLLAY', 'tipo_inversion': 'PROYECTO', 'estado_ssi': 'ACTIVO'},
        {'cui': 2158997, 'resultado': 'ok', 'monto_inversion': 0.0, 'monto_final': 999.0,
         'uei': 'DIRECCION REGIONAL DE SALUD', 'tipo_inversion': 'PROYECTO PROCOMPITE', 'estado_ssi': 'ACTIVO'},
    ]

    caso('escritura: solo se envían las celdas que cambian, en tramos')
    lib = Libro(col_a)
    av, res = m.escribir_proyeccion(lib, HR, filas, filas_proy, insertados=[2094012])
    print('   enviado:', lib.escrito)
    # monto: 2064096 (vacío -> 100.5) y 2158997 (10 -> 999). 2410530 ya tenía 7.0: no se reescribe.
    assert lib.escrito['I3:I3'] == [100.5], lib.escrito
    assert lib.escrito['I7:I7'] == [999.0], lib.escrito
    # la fila 5 (monto null) y la 4 (sin cambio) NO aparecen en ningún rango
    assert not any('I4' in r or 'I5' in r for r in lib.escrito), lib.escrito
    # UEI: 2064096 -> GRI, 2158997 -> DIRESA. 2410530 ya decía GRI. 2094012 sin equivalencia.
    assert lib.escrito['D3:D3'] == ['GRI'] and lib.escrito['D7:D7'] == ['DIRESA'], lib.escrito
    assert not any('D4' in r or 'D6' in r for r in lib.escrito), lib.escrito
    # Etapa: solo el PROCOMPITE de la fila 7
    assert lib.escrito['F7:F7'] == ['PC'], lib.escrito
    assert res['escrito'] and res['montos_actualizados'] == 2 and res['uei_sin_equivalencia'] == 1, res

    caso('escritura: avisos de UEI sin equivalencia, etapa rara y marcado PC')
    assert any('POCOLLAY' in a for a in av), av
    assert any('2158997' in a and 'PC' in a for a in av), av
    assert any('3999999' not in a for a in av)              # está en CUIS_SIN_AVISO
    for a in av:
        print('   -', a.split('\n')[0][:105])

    caso('escritura: un CUI recién insertado sin etapa deducible sí se reporta')
    assert any('2094012' in a and 'Etapa' in a for a in av), av

    caso('escritura: las filas se corrieron -> no se escribe nada')
    lib = Libro([[9999999]] + col_a)      # alguien insertó una fila arriba
    av2, res2 = m.escribir_proyeccion(lib, HR, filas, filas_proy, insertados=[])
    assert lib.escrito is None and not res2['escrito'], res2
    assert any('se corrieron' in a for a in av2), av2
    print('   ', [a.split('\n')[0][:95] for a in av2 if 'corrieron' in a][0])

    caso('escritura: una columna se movió durante la corrida -> no se escribe nada')
    lib = Libro(col_a, tits=['CUI', 'Inversión', 'ID', 'Modalidad', 'UEI', 'Etapa', '', '',
                             '(A) Monto de inversión'])
    av3, res3 = m.escribir_proyeccion(lib, HR, filas, filas_proy, insertados=[])
    assert lib.escrito is None and not res3['escrito'], res3
    assert any('D2' in a for a in av3), av3

    caso('escritura: nada que actualizar -> no se llama a la hoja')
    lib = Libro(col_a)
    sin_cambios = {2410530: filas_proy[2410530]}
    av4, res4 = m.escribir_proyeccion(lib, HR, [filas[1]], sin_cambios, insertados=[])
    assert lib.escrito is None and not res4['escrito'] and res4['rangos_enviados'] == 0, res4

    caso('limpieza: PIM-diario O:Q se vacía con batch_clear (no con "")')
    m.limpiar_pim_opq(HP)
    assert HP.limpiado == ['O2:Q'], HP.limpiado


# ---------- comparación con la corrida anterior ----------
def prueba_comparacion():
    with tempfile.TemporaryDirectory() as tmp:
        m.CARPETA_MONTOS = Path(tmp)
        lunes = date(2026, 11, 23)        # 23/11/2026 es lunes
        viernes = lunes - timedelta(days=3)

        def guardar(f, datos, hora='064500'):
            (Path(tmp) / f'monto_ssi_{f:%Y-%m-%d}_{hora}.json').write_text(
                json.dumps({'datos': datos}), encoding='utf-8')

        caso('comparación: sin corrida anterior -> se avisa')
        filas = [{'cui': 1, 'monto_inversion': 10.0}]
        ruta, var, ceros, av = m.comparar_con_anterior(filas, hoy=lunes)
        assert ruta is None and var == [] and ceros == [] and filas[0]['monto_final'] == 10.0
        assert av and 'NINGUNA corrida anterior' in av[0], av

        caso('comparación día 1: alzas, bajas, cero sospechoso, sin dato, antes 0, igual, nuevo')
        guardar(lunes - timedelta(days=1),
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
        ruta, var, ceros, av = m.comparar_con_anterior(hoy, hoy=lunes)
        finales = {f['cui']: f['monto_final'] for f in hoy}
        assert finales == {1: 105.0, 2: 170.0, 3: 300.0, 4: 400.0, 5: 50.0, 6: 600.0, 7: 70.0, 8: 80.0}, finales
        assert [(v['cui'], v['variacion_pct']) for v in var] == [(5, None), (2, -15.0), (1, 5.0)], var
        assert [z['cui'] for z in ceros] == [3], ceros
        x = m.excel_variaciones(var, ceros, 'prueba', ruta, Path(tmp))
        from openpyxl import load_workbook
        wb = load_workbook(x)
        assert wb.sheetnames == ['Variaciones', 'Ceros sospechosos', 'Nota']
        assert wb['Variaciones']['E3'].value == -0.15 and wb['Variaciones']['E3'].number_format == '0.00%'
        print('   Excel', x.name, 'con hojas', wb.sheetnames)

        caso('comparación día 2: el CUI 3 sigue en 0 -> se mantiene 300 SIN alerta')
        guardar(lunes, hoy)
        dia2 = [{'cui': 3, 'monto_inversion': 0.0, 'nombre_inversion': 'CAE A CERO'}]
        ruta, var, ceros, av = m.comparar_con_anterior(dia2, hoy=lunes + timedelta(days=1))
        assert ruta.name == f'monto_ssi_{lunes:%Y-%m-%d}_064500.json', ruta.name
        assert dia2[0]['monto_final'] == 300.0 and ceros == [] and var == [], (dia2, ceros, var)
        print('   observación:', dia2[0]['observacion'])

        caso('comparación día 3: el CUI 3 vuelve con 310 -> variación contra 300')
        guardar(lunes + timedelta(days=1), dia2)
        dia3 = [{'cui': 3, 'monto_inversion': 310.0, 'nombre_inversion': 'CAE A CERO'}]
        ruta, var, ceros, av = m.comparar_con_anterior(dia3, hoy=lunes + timedelta(days=2))
        assert [(v['cui'], v['monto_anterior'], v['variacion_pct']) for v in var] == [(3, 300.0, 3.33)], var

        caso('comparación: un cero nuevo en otro CUI sí alerta aunque otro siga en 0')
        guardar(lunes + timedelta(days=2), dia2 + [{'cui': 9, 'monto_inversion': 90.0}])
        dia4 = [{'cui': 3, 'monto_inversion': 0.0}, {'cui': 9, 'monto_inversion': 0.0}]
        ruta, var, ceros, av = m.comparar_con_anterior(dia4, hoy=lunes + timedelta(days=3))
        assert [z['cui'] for z in ceros] == [9] and {f['cui']: f['monto_final'] for f in dia4} == {3: 300.0, 9: 90.0}

    caso('base: una prueba del MISMO día no sirve de base')
    with tempfile.TemporaryDirectory() as tmp:
        m.CARPETA_MONTOS = Path(tmp)
        (Path(tmp) / f'monto_ssi_{lunes:%Y-%m-%d}_120000.json').write_text(
            json.dumps({'datos': [{'cui': 1, 'monto_inversion': 1.0}]}), encoding='utf-8')
        ruta, ant, av = m.corrida_anterior(hoy=lunes)
        assert ruta is None and ant == {}, (ruta, ant)
        print('   ', av[0][:95])

    caso('base: lunes contra viernes NO avisa (sábado y domingo la PC está apagada)')
    with tempfile.TemporaryDirectory() as tmp:
        m.CARPETA_MONTOS = Path(tmp)
        (Path(tmp) / f'monto_ssi_{viernes:%Y-%m-%d}_064500.json').write_text(
            json.dumps({'datos': [{'cui': 1, 'monto_inversion': 1.0}]}), encoding='utf-8')
        ruta, ant, av = m.corrida_anterior(hoy=lunes)
        assert ruta is not None and av == [], av

        caso('base: martes contra viernes SÍ avisa (falta el lunes)')
        ruta, ant, av = m.corrida_anterior(hoy=lunes + timedelta(days=1))
        assert len(av) == 1 and 'No hay corrida de' in av[0], av
        print('   ', av[0][:120])

    caso('base: archivo dañado -> usa uno más viejo y avisa')
    with tempfile.TemporaryDirectory() as tmp:
        m.CARPETA_MONTOS = Path(tmp)
        (Path(tmp) / 'monto_ssi_2026-01-10_064500.json').write_text(
            json.dumps({'datos': [{'cui': 1, 'monto_inversion': 100.0}]}), encoding='utf-8')
        (Path(tmp) / f'monto_ssi_{lunes - timedelta(days=1):%Y-%m-%d}_064500.json').write_text(
            '{"datos": [{"cui": 1, "mon', encoding='utf-8')
        ruta, var, ceros, av = m.comparar_con_anterior([{'cui': 1, 'monto_inversion': 180.0}], hoy=lunes)
        assert ruta.name == 'monto_ssi_2026-01-10_064500.json', ruta.name
        assert any('dañado' in a for a in av) and any('No hay corrida de' in a for a in av), av
        print('   ', [a[:95] for a in av])


def prueba_ceros_masivos():
    with tempfile.TemporaryDirectory() as tmp:
        m.CARPETA_MONTOS = Path(tmp)
        lunes = date(2026, 11, 23)
        n = m.MAX_CEROS_SOSPECHOSOS
        (Path(tmp) / f'monto_ssi_{lunes - timedelta(days=1):%Y-%m-%d}_064500.json').write_text(
            json.dumps({'datos': [{'cui': c, 'monto_inversion': 100.0 * c} for c in range(1, n + 2)]}),
            encoding='utf-8')

        caso(f'ceros masivos: {n - 1} ceros NO frenan la corrida (solo alertan)')
        pocos = [{'cui': c, 'monto_inversion': 0.0} for c in range(1, n)]
        _, _, ceros, _ = m.comparar_con_anterior(pocos, hoy=lunes)
        assert len(ceros) == n - 1 and len(ceros) < m.MAX_CEROS_SOSPECHOSOS, len(ceros)
        assert [f['monto_final'] for f in pocos] == [100.0 * c for c in range(1, n)]

        caso(f'ceros masivos: {n} ceros SÍ frenan (umbral MAX_CEROS_SOSPECHOSOS = {n})')
        muchos = [{'cui': c, 'monto_inversion': 0.0} for c in range(1, n + 1)]
        _, _, ceros, _ = m.comparar_con_anterior(muchos, hoy=lunes)
        assert len(ceros) >= m.MAX_CEROS_SOSPECHOSOS, len(ceros)
        print(f'   {len(ceros)} ceros >= {m.MAX_CEROS_SOSPECHOSOS}: no se escribe nada en la hoja (código 6)')


def prueba_bloqueo():
    with tempfile.TemporaryDirectory() as tmp:
        m.ARCHIVO_BLOQUEO = Path(tmp) / 'corrida.lock'
        caso('bloqueo: la primera corrida lo toma, la segunda no hace nada')
        assert m.tomar_bloqueo() is True
        assert m.tomar_bloqueo() is False
        m.soltar_bloqueo()
        assert m.tomar_bloqueo() is True
        m.soltar_bloqueo()

        caso('bloqueo: uno abandonado (más viejo que BLOQUEO_MAX_MIN) se reusa')
        m.ARCHIVO_BLOQUEO.write_text('viejo', encoding='utf-8')
        import os
        viejo = __import__('time').time() - (m.BLOQUEO_MAX_MIN + 5) * 60
        os.utime(m.ARCHIVO_BLOQUEO, (viejo, viejo))
        assert m.tomar_bloqueo() is True
        m.soltar_bloqueo()


def prueba_carpeta_respaldo():
    with tempfile.TemporaryDirectory() as tmp:
        caso('carpeta: si la normal no se puede usar, se guarda en la de respaldo y se avisa')
        (Path(tmp) / 'ocupado').write_text('soy un archivo, no una carpeta', encoding='utf-8')
        m.CARPETA_MONTOS = Path(tmp) / 'ocupado' / 'monto_ssi'   # el padre es un archivo: no se puede crear
        m.CARPETA_RESPALDO = Path(tmp) / 'respaldo'
        carpeta, aviso = m.carpeta_montos()
        assert carpeta == m.CARPETA_RESPALDO and aviso and 'respaldo' in str(carpeta), (carpeta, aviso)
        print('   ', aviso[:130])

        caso('carpeta: si la normal sirve, no hay aviso')
        m.CARPETA_MONTOS = Path(tmp) / 'normal'
        carpeta, aviso = m.carpeta_montos()
        assert carpeta == m.CARPETA_MONTOS and aviso is None


if __name__ == '__main__':
    prueba_listas()
    prueba_equivalencias()
    prueba_escritura()
    prueba_comparacion()
    prueba_ceros_masivos()
    prueba_bloqueo()
    prueba_carpeta_respaldo()
    print('\nTODAS LAS PRUEBAS PASARON')
