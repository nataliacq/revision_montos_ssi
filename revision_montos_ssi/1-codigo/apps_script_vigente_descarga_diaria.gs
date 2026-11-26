/**
 * Descarga diaria de Consulta Amigable (MEF) – GORE Tacna, solo proyectos
 * Nivel de gobierno R · Departamento 23 TACNA · Producto/Proyecto
 *
 * Se pega en: el libro PIM-diario › Extensiones › Apps Script.
 * Cada corrida descarga los datos del día y REEMPLAZA lo que haya en la hoja.
 * No guarda histórico ni toca otros libros.
 *
 * Uso:
 *   1) descargaDiaria   -> correr una vez a mano para probar (y autorizar permisos)
 *   2) instalarHorario  -> correr una vez; deja la descarga diaria cerca de las 6:20 (hora de Lima)
 *   3) probarProyeccion -> agrega a Proyección26 los CUI que falten y pega en PIM-diario:
 *        · columna A (desde A3): los CUI insertados en esta corrida
 *        · columna N (desde N2): los CUI de Proyección26 cuya Etapa (col. F) no es "PC"
 */

const ANIO = 2026;  // cambiar cada 1 de enero
// El correo de avisos NO va escrito aquí: se lee de las propiedades del proyecto de Apps Script
// (Configuración del proyecto > Propiedades de la secuencia de comandos), con el nombre CORREO_USUARIO.
const PROP_CORREO = 'CORREO_USUARIO';
const NOMBRE_HOJA = 'PIM-diario';          // donde se pega dato del dia
const MIN_FILAS = 150;           // si llegan menos, se considera descarga fallida y no se toca la hoja
const CELDA_EJECUCION = 'A1';    // aquí se escribe la fecha y hora de la última descarga buena

// ================== Proyección26: insertar CUIs nuevos ==================
const HOJA_PROY = 'Proyección26';
const COL_FORM_INI = 13;  // M
const COL_FORM_FIN = 27;  // AA
const COL_ETAPA = 6;      // F en Proyección26 (título esperado: "Etapa")
// Los títulos de Proyección26 están en la fila 2, EXACTAMENTE ahí. Si se movieron (alguien
// insertó o borró una fila arriba) no se sabe dónde empieza la lista: se detiene y se reporta,
// igual que hace cuis_desde_sheets.py. Antes se buscaban en las primeras 30 filas, así que este
// guion seguía trabajando y rearmaba la columna N mientras el de Python se detenía.
const PROY_FILA_TITULOS = 2;
const PROY_TITULOS = [[1, 'CUI'], [2, 'Inversión'], [COL_ETAPA, 'Etapa']];
// Títulos esperados de M a AA (deben coincidir al 100%)
const TITULOS_M_AA = [
  'PIA', 'PIM', 'Por Cert.\n(PIM-Cert)', 'Certificación',
  'CA-Cert\n(Por Comp).\nEn pedido cot', 'Compromiso Anual',
  'Atención de Compromiso Mensual', 'CA-Dev-\n(Por Devengar)\nCon Orden',
  'Dev-Cert\nEntre pedidos y órdenes', 'Libre PIM, pedidos y órdenes',
  '(C) Devengado 2026', 'PIM-Dev\n(Saldo por devengar)',
  'Por Girar\n(Dev-Gir)', 'Girado', 'Avance %\n(Dev/P)'
];

// Dónde se pegan las listas de CUI en PIM-diario
const FILA_INI_INSERTADOS = 3;   // columna A, desde A3: CUI insertados en Proyección26
const COL_NOPC = 14;             // columna N, desde N2: CUI de Proyección26 con Etapa ≠ "PC"
const TITULO_NOPC = 'CUI_proyeccion26_noPC';

// no cambian
const ZONA = 'America/Lima';
const UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0 Safari/537.36';

// Columnas A y B no se tocan con la descarga (A1 = fecha y hora; A3 hacia abajo = CUI insertados).
// Los datos van de C a L.
const ENCABEZADOS = ['', '', 'CUI', 'Inversión', 'PIA', 'PIM', 'Certificación',
  'Compromiso Anual', 'Atención de Compromiso Mensual', 'Devengado', 'Girado', 'Avance %'];

function urlMEF_(ANIO) {
  return 'https://apps5.mineco.gob.pe/transparencia/Navegador/Navegar_7.aspx' +
    '?_tgt=xls&_uhc=yes&0=&1=R&21=23&31=&y=' + ANIO + '&ap=Proyecto&cpage=1&psize=400';
}

/** Tarea diaria. */
function descargaDiaria() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(60 * 1000)) return;
  try {
    // Momento de la corrida en hora de Lima, como valor de fecha y hora de la hoja
    const momento = Utilities.formatDate(new Date(), ZONA, 'yyyy-MM-dd HH:mm:ss');
    // Primero se descarga; si falla, la hoja se queda con la última descarga buena
    const datos = descargarMEF_(ANIO);
    escribirHoja_(datos, momento);
    // Las listas de CUI se rearman en la MISMA corrida. Antes esto solo pasaba al correr
    // probarProyeccion a mano, así que la columna N se quedaba congelada y los CUI nuevos de
    // Proyección26 nunca entraban: el disparador diario solo llenaba C:L y A1.
    agregarCuisNuevos_(datos.map(f => [f[0], f[1]]));
  } catch (e) {
    console.error(e);
    avisar_('Falló la descarga diaria del MEF',
      'La hoja no se actualizó (se quedó con la última descarga buena).' +
      '\n\nDetalle: ' + e.message);
  } finally {
    lock.releaseLock();
  }
}

/** Crea (una sola vez) el disparador diario cerca de las 6:20, hora de Lima. */
function instalarHorario() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'descargaDiaria')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('descargaDiaria')
    .timeBased().everyDays(1).atHour(6).nearMinute(20).inTimezone(ZONA)
    .create();
  console.log('Horario instalado: todos los días cerca de las 6:20 (' + ZONA + ').');
}

// ---------- auxiliares ----------

/**
 * Deja en la hoja solo la descarga del día (columnas C a L).
 * Borra contenido (no filas), así las fórmulas o IMPORTRANGE que apunten a esta hoja no se rompen.
 * No toca las columnas A y B (salvo A1 con la fecha) ni la columna N.
 */
function escribirHoja_(datos, fechaHora) {
  const hoja = SpreadsheetApp.getActive().getSheetByName(NOMBRE_HOJA);
  const n = datos.length;
  const ancho = ENCABEZADOS.length - 2;   // de C a L

  if (!hoja) throw new Error('No existe la hoja "' + NOMBRE_HOJA + '". Revisar el nombre de la pestaña.');

  // medir la descarga anterior ANTES de escribir
  const ultima = Math.max(hoja.getLastRow(), 1);

  // que haya filas suficientes (solo se agregan al final, nunca se borran)
  const faltan = n + 1 - hoja.getMaxRows();
  if (faltan > 0) hoja.insertRowsAfter(hoja.getMaxRows(), faltan);

  // escribir primero, sin borrar antes
  hoja.getRange(1, 3, 1, ancho).setValues([ENCABEZADOS.slice(2)]).setFontWeight('bold');
  hoja.getRange(2, 3, n, ancho).setValues(datos);

  // después borrar solo lo que sobró de la descarga anterior (de C a L)
  const sobrantes = ultima - (n + 1);
  if (sobrantes > 0) hoja.getRange(n + 2, 3, sobrantes, ancho).clearContent();

  hoja.getRange(2, 3, n, 1).setNumberFormat('0');            // CUI
  hoja.getRange(2, 4, n, 1).setNumberFormat('@');            // Inversión (texto)
  hoja.getRange(2, 5, n, 7).setNumberFormat('#,##0');        // PIA a Girado
  hoja.getRange(2, 12, n, 1).setNumberFormat('0.0');         // Avance %

  // fecha y hora de esta descarga, en una sola celda
  hoja.getRange(CELDA_EJECUCION).setValue(fechaHora)
    .setNumberFormat('dd/mm/yyyy hh:mm:ss').setFontWeight('bold');
  SpreadsheetApp.flush();
}

function descargarMEF_(ANIO) {
  let ultimoError = '';
  for (let intento = 1; intento <= 3; intento++) {
    try {
      const r = UrlFetchApp.fetch(urlMEF_(ANIO), { muteHttpExceptions: true, headers: { 'User-Agent': UA } });
      if (r.getResponseCode() === 200) {
        const filas = extraerFilas_(r.getContentText('windows-1252'));
        if (filas.length >= MIN_FILAS) return filas;
        ultimoError = 'la respuesta trajo solo ' + filas.length + ' filas de proyectos';
      } else {
        ultimoError = 'código de respuesta ' + r.getResponseCode();
      }
    } catch (e) {
      ultimoError = e.message;
    }
    if (intento < 3) Utilities.sleep(30 * 1000);
  }
  throw new Error('No se pudo descargar del MEF (' + ultimoError + ').');
}

/** Devuelve [CUI, nombre, PIA, PIM, Cert, CA, ACM, Dev, Gir, Avance] por proyecto, en el orden del MEF. */
function extraerFilas_(html) {
  const filas = [];
  const trs = html.match(/<tr[^>]*>[\s\S]*?<\/tr>/gi) || [];
  trs.forEach(tr => {
    const celdas = (tr.match(/<td[^>]*>[\s\S]*?<\/td>/gi) || []).map(limpiar_);
    if (celdas.length < 9) return;
    const m = celdas[0].match(/^(\d{7})[:.]\s*(.*)$/);
    if (!m) return;
    filas.push([Number(m[1]), m[2]].concat(celdas.slice(1, 9).map(numero_)));
  });
  return filas;
}

function limpiar_(td) {
  return td.replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/gi, ' ').replace(/&amp;/gi, '&').replace(/&quot;/gi, '"')
    .replace(/&lt;/gi, '<').replace(/&gt;/gi, '>').replace(/&#39;/g, "'")
    .replace(/&#(\d+);/g, (_, n) => String.fromCharCode(Number(n)))
    .replace(/[\s ]+/g, ' ')
    .trim();
}

function numero_(txt) {
  const n = parseFloat(String(txt).replace(/,/g, '').replace(/\s/g, ''));
  return isNaN(n) ? '' : n;
}

// ================== Proyección26: insertar CUIs nuevos ==================

/** Para probar a mano: descarga el MEF, agrega a Proyección26 los CUI que falten y pega las listas en PIM-diario. */
function probarProyeccion() {
  const proyectos = descargarMEF_(ANIO);             // [CUI, nombre, PIA, PIM, ...] desde la web
  agregarCuisNuevos_(proyectos.map(f => [f[0], f[1]])); // solo CUI y nombre
}

/**
 * proyectos = [[CUI, nombre], ...] según la descarga del MEF
 * 1) Inserta en Proyección26 los CUI que falten.
 * 2) Pega en PIM-diario:
 *    - columna A, desde A3: los CUI insertados en esta corrida
 *    - columna N, desde N2 (título en N1): los CUI de Proyección26 con Etapa distinta de "PC"
 */
function agregarCuisNuevos_(proyectos) {
  const insertados = [];   // se va llenando a medida que se insertan filas

  const hoja = SpreadsheetApp.getActive().getSheetByName(HOJA_PROY);
  if (!hoja) {
    avisar_('Proyección26: no existe la hoja',
      'No se encontró la hoja "' + HOJA_PROY + '". NO se tocaron las listas de CUI de "' +
      NOMBRE_HOJA + '" (columnas A y N): se quedaron como estaban. Revisar el nombre de la pestaña.');
    return;
  }

  // Los títulos tienen que estar exactamente en la fila PROY_FILA_TITULOS. Si no, se para aquí
  // SIN tocar PIM-diario: vaciar la columna N dejaría la hoja sin la lista que todo lo demás usa.
  const faltas = verificarTitulosProy_(hoja);
  if (faltas.length) {
    avisar_('Proyección26: los títulos no están donde deben',
      'Los títulos deben estar en la fila ' + PROY_FILA_TITULOS + ' de "' + HOJA_PROY + '":\n\n' +
      faltas.join('\n') +
      '\n\nPuede que se haya insertado o borrado una fila arriba, o que se haya movido una columna. ' +
      'No se insertó ningún CUI y NO se tocaron las listas de "' + NOMBRE_HOJA + '" (columnas A y N), ' +
      'que se quedaron como estaban. Corregir los títulos y volver a correr.');
    return;
  }

  try {
    insertarCuisNuevos_(hoja, proyectos, insertados);
  } catch (e) {
    console.error(e);
    avisar_('Proyección26: error al insertar CUIs nuevos',
      'La descarga a PIM-diario sí se hizo, pero falló la actualización de Proyección26.\n\nDetalle: ' + e.message +
      (insertados.length ? '\n\nAlcanzaron a insertarse estos CUI:\n' + insertados.join('\n') : ''));
  }

  let noPC = null;
  try {
    noPC = cuisNoPC_(hoja);
  } catch (e) {
    console.error(e);
    noPC = null;
  }
  if (noPC === null || !noPC.length) {
    avisar_('Proyección26: no se pudo armar la lista de CUI que no son PC',
      'NO se tocaron las listas de "' + NOMBRE_HOJA + '" (columnas A y N): se quedaron como estaban. ' +
      'Vaciar la columna N habría borrado la lista con la que trabaja el programa de montos.' +
      (insertados.length ? '\n\nCUI insertados en Proyección26:\n' + insertados.join('\n') : ''));
    return;
  }

  try {
    escribirListasPIM_(insertados, noPC);
  } catch (e) {
    console.error(e);
    avisar_('PIM-diario: no se pudieron pegar las listas de CUI',
      'Falló al escribir en "' + NOMBRE_HOJA + '" (columna A desde A' + FILA_INI_INSERTADOS +
      ' y columna N).\n\nDetalle: ' + e.message +
      (insertados.length ? '\n\nCUI insertados en Proyección26:\n' + insertados.join('\n') : ''));
  }
}


/** Títulos que faltan o están mal en la fila PROY_FILA_TITULOS de Proyección26 (lista vacía = todo bien). */
function verificarTitulosProy_(hoja) {
  const ultCol = Math.max(hoja.getLastColumn(), COL_ETAPA);
  if (hoja.getLastRow() < PROY_FILA_TITULOS) {
    return ['la hoja tiene menos de ' + PROY_FILA_TITULOS + ' filas'];
  }
  const fila = hoja.getRange(PROY_FILA_TITULOS, 1, 1, ultCol).getValues()[0];
  return PROY_TITULOS
    .filter(([c, t]) => normCmp_(fila[c - 1]) !== normCmp_(t))
    .map(([c, t]) => colLetra_(c) + PROY_FILA_TITULOS + ': se esperaba "' + t +
                     '" y dice "' + (normCmp_(fila[c - 1]) || '(vacía)') + '"');
}

/** Inserta en Proyección26 los CUI que falten. Cada CUI insertado se agrega al arreglo `insertados`. */
function insertarCuisNuevos_(hoja, proyectos, insertados) {
  const ultFila = hoja.getLastRow();
  const ultCol = Math.max(hoja.getLastColumn(), COL_FORM_FIN);
  const valores = hoja.getRange(1, 1, ultFila, ultCol).getValues();

  // 1) Los títulos ya se verificaron en agregarCuisNuevos_: están en la fila PROY_FILA_TITULOS
  const filaTit = PROY_FILA_TITULOS;

  // 2) Ubicar el primer y el último CUI (pueden haber filas vacías o con otro texto antes/después)
  const existentes = [];   // [{fila, cui}]
  for (let i = filaTit; i < valores.length; i++) {
    if (esCui_(valores[i][0])) existentes.push({ fila: i + 1, cui: Number(valores[i][0]) });
  }
  if (!existentes.length) {
    avisar_('Proyección26: la columna A no tiene CUIs',
      'Debajo del título "CUI" (fila ' + filaTit + ') no hay códigos de 7 cifras en la columna A. ' +
      'No se insertó nada.');
    return;
  }
  const primera = existentes[0].fila;
  const ultima = existentes[existentes.length - 1].fila;

  // 3) CUIs nuevos (en el MEF pero no en Proyección26)
  const ya = new Set(existentes.map(e => e.cui));
  const vistos = new Set();
  const nuevos = proyectos
    .filter(p => esCui_(p[0]))
    .map(p => [Number(p[0]), String(p[1] || '').trim()])
    .filter(p => !ya.has(p[0]) && !vistos.has(p[0]) && vistos.add(p[0]));
  if (!nuevos.length) return;

  // 4) ¿Los títulos de M a AA coinciden al 100%?
  const titActuales = valores[filaTit - 1].slice(COL_FORM_INI - 1, COL_FORM_FIN).map(norm_);
  const difs = [];
  TITULOS_M_AA.forEach((t, k) => {
    if (titActuales[k] !== norm_(t)) {
      difs.push(colLetra_(COL_FORM_INI + k) + ': se esperaba "' + t + '" y dice "' + titActuales[k] + '"');
    }
  });
  const copiarFormulas = difs.length === 0;

  // 5) Insertar de abajo hacia arriba (así las posiciones no se corren)
  nuevos.sort((a, b) => b[0] - a[0]);
  const alFinal = [];
  const enDuro = new Map();      // celdas de referencia con valor en duro
  const quedaronVacias = [];     // celdas nuevas sin fórmula (ni arriba ni abajo tenían)

  nuevos.forEach(([cui, nombre]) => {
    const sig = existentes.find(e => e.cui > cui);
    let fila;
    if (sig) {
      hoja.insertRowBefore(sig.fila);
      fila = sig.fila;
    } else {
      hoja.insertRowAfter(ultima);
      fila = ultima + 1;
      alFinal.push(cui);
    }
    // Formato igual a la fila vecina
    const ref = fila > primera ? fila - 1 : fila + 1;
    hoja.getRange(ref, 1, 1, ultCol)
      .copyTo(hoja.getRange(fila, 1, 1, ultCol), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);

    // CUI y nombre (con el CUI adelante, como el resto de la hoja; la columna J lo usa)
    hoja.getRange(fila, 1, 1, 2).setValues([[cui, cui + ': ' + nombre]]);
    insertados.push(cui);

    // Fórmulas M:AA de la fila de arriba, o de abajo si arriba no hay
    if (copiarFormulas) {
      const n = COL_FORM_FIN - COL_FORM_INI + 1;
      const leer = r => {
        if (r < 1 || r > hoja.getMaxRows()) return null;
        const cuiRef = hoja.getRange(r, 1).getValue();
        if (!esCui_(cuiRef)) return null;
        const rg = hoja.getRange(r, COL_FORM_INI, 1, n);
        return { cui: Number(cuiRef), f: rg.getFormulasR1C1()[0], v: rg.getValues()[0] };
      };
      const arriba = leer(fila - 1);
      const abajo = leer(fila + 1);

      [arriba, abajo].forEach(ref => {
        if (!ref) return;
        ref.f.forEach((fx, k) => {
          const v = ref.v[k];
          if (!fx && v !== '' && v !== null) {
            const col = colLetra_(COL_FORM_INI + k);
            enDuro.set(ref.cui + '|' + col,
              'CUI ' + ref.cui + ', columna ' + col + ' (' + TITULOS_M_AA[k].replace(/\n/g, ' ') +
              '): tiene el valor "' + v + '" en vez de fórmula');
          }
        });
      });

      for (let k = 0; k < n; k++) {
        const fx = (arriba && arriba.f[k]) || (abajo && abajo.f[k]) || '';
        if (fx) hoja.getRange(fila, COL_FORM_INI + k).setFormulaR1C1(fx);
        else quedaronVacias.push('CUI ' + cui + ', columna ' + colLetra_(COL_FORM_INI + k));
      }
    }
  });
  SpreadsheetApp.flush();

  if (enDuro.size || quedaronVacias.length) {
    avisar_('Proyección26: valores en duro en vez de fórmulas',
      'Se insertaron los CUIs nuevos, pero en las filas vecinas que se usaron de modelo ' +
      'hay celdas de M a AA con valores fijos en lugar de fórmulas.\n\n' +
      (enDuro.size ? 'Celdas con valor en duro:\n' + [...enDuro.values()].join('\n') + '\n\n' : '') +
      (quedaronVacias.length ? 'Celdas nuevas que quedaron vacías (no se encontró fórmula ni arriba ni abajo):\n' +
        quedaronVacias.join('\n') : ''));
  }

  if (!copiarFormulas) {
    avisar_('Proyección26: títulos de M a AA cambiaron',
      'Se insertaron los CUIs nuevos (con CUI e Inversión), pero NO se copiaron las fórmulas de M a AA ' +
      'porque los títulos no coinciden:\n\n' + difs.join('\n') +
      '\n\nCUIs insertados sin fórmulas:\n' + listar_(nuevos));
  }
  if (alFinal.length) {
    avisar_('Proyección26: CUI agregado después del último',
      'Estos CUIs eran mayores que todos los existentes y quedaron al final de la lista. ' +
      'Revisar que los totales (SUMA/SUBTOTAL) los incluyan:\n' + alFinal.join('\n'));
  }
}

/**
 * Lee Proyección26 (ya con los CUI nuevos) y devuelve los CUI cuya Etapa (columna F)
 * es distinta de "PC", incluidas las celdas vacías.
 * Devuelve null si no se puede calcular (y en ese caso avisa por correo).
 */
function cuisNoPC_(hoja) {
  const ultFila = hoja.getLastRow();
  if (ultFila < 1) return null;
  const valores = hoja.getRange(1, 1, ultFila, COL_ETAPA).getValues();

  const filaTit = PROY_FILA_TITULOS;   // ya verificado en agregarCuisNuevos_

  const lista = [];
  const vistos = new Set();
  for (let i = filaTit; i < valores.length; i++) {
    if (!esCui_(valores[i][0])) continue;
    const cui = Number(valores[i][0]);
    const etapa = String(valores[i][COL_ETAPA - 1]).trim().toUpperCase();
    if (etapa !== 'PC' && !vistos.has(cui)) {
      vistos.add(cui);
      lista.push(cui);
    }
  }
  return lista;
}

/**
 * Pega en PIM-diario:
 *   - columna A desde A3: CUI insertados en esta corrida (de menor a mayor)
 *   - columna N: título en N1 y, desde N2, los CUI que no son PC (vacía si noPC es null)
 * Antes borra el contenido anterior de A3:A y N2:N.
 */
function escribirListasPIM_(insertados, noPC) {
  const hoja = SpreadsheetApp.getActive().getSheetByName(NOMBRE_HOJA);
  if (!hoja) throw new Error('No existe la hoja "' + NOMBRE_HOJA + '". Revisar el nombre de la pestaña.');

  const listaA = [...insertados].sort((a, b) => a - b);
  const listaN = noPC || [];
  // Red de seguridad: sin lista nueva NO se vacía la columna N. Antes se hacía clearContent()
  // de N2:N siempre y solo se reponía la lista si se había podido calcular, así que un error
  // en Proyección26 dejaba la columna N vacía.
  if (!listaN.length) throw new Error('la lista de CUI que no son PC vino vacía: no se tocó nada');

  // que haya filas y columnas suficientes (solo se agregan, nunca se borran)
  const necesarias = Math.max(FILA_INI_INSERTADOS, FILA_INI_INSERTADOS - 1 + listaA.length, 1 + listaN.length);
  if (necesarias > hoja.getMaxRows()) hoja.insertRowsAfter(hoja.getMaxRows(), necesarias - hoja.getMaxRows());
  if (COL_NOPC > hoja.getMaxColumns()) hoja.insertColumnsAfter(hoja.getMaxColumns(), COL_NOPC - hoja.getMaxColumns());
  const maxFilas = hoja.getMaxRows();

  // Columna A: borrar la lista anterior y pegar la nueva
  hoja.getRange(FILA_INI_INSERTADOS, 1, maxFilas - FILA_INI_INSERTADOS + 1, 1).clearContent();
  if (listaA.length) {
    hoja.getRange(FILA_INI_INSERTADOS, 1, listaA.length, 1)
      .setValues(listaA.map(c => [c])).setNumberFormat('0');
  }

  // Columna N: título, borrar la lista anterior y pegar la nueva
  hoja.getRange(1, COL_NOPC).setValue(TITULO_NOPC).setFontWeight('bold');
  hoja.getRange(2, COL_NOPC, maxFilas - 1, 1).clearContent();
  hoja.getRange(2, COL_NOPC, listaN.length, 1)
    .setValues(listaN.map(c => [c])).setNumberFormat('0');

  // Columnas O, P y Q: ya no se usan. El monto ahora se escribe en Proyección26 (columna I),
  // así que lo que quede aquí es de corridas viejas y solo puede engañar a una fórmula.
  hoja.getRange(2, COL_NOPC + 1, maxFilas - 1, 3).clearContent();
  SpreadsheetApp.flush();
}

function esCui_(v) { return /^\d{7}$/.test(String(v).trim().replace(/\.0+$/, '')); }
function norm_(t) { return String(t).replace(/\r\n/g, '\n').trim(); }

/** Para comparar títulos: sin tildes, sin espacios de más y en mayúsculas (igual que el Python). */
function normCmp_(t) {
  return String(t === null || t === undefined ? '' : t)
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ').trim().toUpperCase();
}

/** Correo de avisos, desde las propiedades del proyecto de Apps Script. */
function correoUsuario_() {
  const c = PropertiesService.getScriptProperties().getProperty(PROP_CORREO);
  if (!c) throw new Error('Falta la propiedad "' + PROP_CORREO + '" en Configuración del proyecto > ' +
                          'Propiedades de la secuencia de comandos.');
  return c;
}

function avisar_(asunto, cuerpo) {
  const texto = cuerpo + '\n\nLibro: ' + SpreadsheetApp.getActive().getUrl();
  try {
    MailApp.sendEmail(correoUsuario_(), asunto, texto);
  } catch (e) {
    // que el aviso no se pierda del todo si el correo no está configurado o falla el envío
    console.error('NO se pudo avisar por correo (' + e.message + ')\n' + asunto + '\n\n' + texto);
  }
}
function listar_(lista) { return lista.map(p => p[0] + ' – ' + p[1]).join('\n'); }
function colLetra_(c) { let s = ''; while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - m) / 26); } return s; }
