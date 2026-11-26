/**
 * Pruebas de la cola de CUI insertados de apps_script_vigente_descarga_diaria.gs, con una hoja
 * simulada. No usa Google: comprueba la parte delicada de escribirListasPIM_ (dónde empieza a
 * agregar, que no duplique, que no borre la cola y que sí reemplace la columna N).
 *
 *     node pruebas_cola_gs.js
 */
const fs = require('fs');
const path = require('path');
const src = fs.readFileSync(path.join(__dirname, 'apps_script_vigente_descarga_diaria.gs'), 'utf8');

// --- hoja simulada: una matriz de celdas, 1-indexada ---
function Hoja(filas, cols) {
  const c = Array.from({length: filas}, () => Array(cols).fill(''));
  return {
    celdas: c,
    getMaxRows: () => c.length,
    getMaxColumns: () => cols,
    insertRowsAfter: (desp, n) => { for (let i=0;i<n;i++) c.push(Array(cols).fill('')); },
    insertColumnsAfter: () => {},
    getRange(f, col, nf = 1, nc = 1) {
      return {
        getValues: () => Array.from({length: nf}, (_, i) =>
          Array.from({length: nc}, (_, j) => c[f-1+i][col-1+j])),
        setValues(v) { v.forEach((r, i) => r.forEach((x, j) => { c[f-1+i][col-1+j] = x; })); return this; },
        setValue(v) { c[f-1][col-1] = v; return this; },
        setNumberFormat() { return this; },
        setFontWeight() { return this; },
        clearContent() { for (let i=0;i<nf;i++) for (let j=0;j<nc;j++) c[f-1+i][col-1+j] = ''; return this; },
      };
    },
  };
}

let hoja;
global.SpreadsheetApp = { getActive: () => ({ getSheetByName: () => hoja, getUrl: () => 'x' }), flush: () => {} };
global.Utilities = { formatDate: () => '2026-11-26' };
// las constantes y funciones reales del .gs, puestas en el ambito global
for (const nom of ['FILA_INI_INSERTADOS','COL_FECHA_INS','TITULO_FECHA_INS','COL_NOPC',
                   'TITULO_NOPC','NOMBRE_HOJA','ZONA']) {
  const linea = src.match(new RegExp('^const ' + nom + ' = (.*?);', 'm'));
  if (!linea) throw new Error('no se encontro la constante ' + nom);
  eval('globalThis.' + nom + ' = ' + linea[1] + ';');
}
eval('globalThis.esCui_ = ' + src.match(/^function esCui_\s*\(([\s\S]*?)\n/m)[0]
     .replace(/^function esCui_/, 'function'));
eval('globalThis.escribirListasPIM_ = ' +
     src.match(/function escribirListasPIM_[\s\S]*?\n}/)[0].replace(/^function escribirListasPIM_/, 'function'));

const colA = () => hoja.celdas.slice(2).map(r => r[0]).filter(x => x !== '');
const colB = () => hoja.celdas.slice(2).map(r => r[1]).filter(x => x !== '');
const ok = (a, b, q) => { if (JSON.stringify(a) !== JSON.stringify(b)) throw new Error(q + ': ' + JSON.stringify(a) + ' != ' + JSON.stringify(b)); };

console.log('== cola vacía: se agregan los nuevos desde A3, con fecha en B');
hoja = Hoja(300, 17);
escribirListasPIM_([2064096, 2410530], [1111111, 2222222]);
ok(colA(), [2064096, 2410530], 'A');
ok(colB(), ['2026-11-26', '2026-11-26'], 'B');
ok(hoja.celdas[1][1], 'fecha_insercion', 'título B2');
ok(hoja.celdas[1][0], '', 'A2 sigue vacía (ancla de Python)');
ok(hoja.celdas[0][13], 'CUI_proyeccion26_noPC', 'título N1');

console.log('== segunda corrida, Python NO consumió: se AGREGAN debajo, sin borrar');
Utilities.formatDate = () => '2026-11-27';
escribirListasPIM_([3000333], [1111111]);
ok(colA(), [2064096, 2410530, 3000333], 'A acumulada');
ok(colB(), ['2026-11-26', '2026-11-26', '2026-11-27'], 'cada uno con SU fecha');

console.log('== un CUI que ya está en la cola no se duplica');
escribirListasPIM_([2410530, 3000333], [1111111]);
ok(colA(), [2064096, 2410530, 3000333], 'sin duplicados');

console.log('== sin CUI nuevos: la cola queda intacta (no se borra)');
escribirListasPIM_([], [1111111]);
ok(colA(), [2064096, 2410530, 3000333], 'intacta');

console.log('== después de que Python borra A3:B, la cola arranca de nuevo en A3');
hoja.getRange(3, 1, hoja.getMaxRows() - 2, 2).clearContent();
Utilities.formatDate = () => '2026-11-28';
escribirListasPIM_([2555555], [1111111]);
ok(colA(), [2555555], 'A3 otra vez');
ok(colB(), ['2026-11-28'], 'B3');
ok(hoja.celdas[2][0], 2555555, 'está en la fila 3');

console.log('== la columna N sí se reemplaza entera, y O:Q se vacía');
hoja.celdas[2][14] = 'monto viejo'; hoja.celdas[3][16] = 'tipo viejo';
escribirListasPIM_([], [9999999]);
ok(hoja.celdas.slice(1).map(r => r[13]).filter(x => x !== ''), [9999999], 'N reemplazada');
ok(hoja.celdas[2][14], '', 'O vaciada');
ok(hoja.celdas[3][16], '', 'Q vaciada');

console.log('== sin lista de no-PC lanza error y no toca nada');
const antes = JSON.stringify(hoja.celdas);
try { escribirListasPIM_([2777777], []); throw new Error('debió lanzar'); }
catch (e) { if (!/vino vacía/.test(e.message)) throw e; }
ok(JSON.stringify(hoja.celdas), antes, 'hoja sin cambios');

console.log('\nTODAS LAS PRUEBAS DE LA COLA (.gs) PASARON');
