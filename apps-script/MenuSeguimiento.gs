/**
 * Menú del Seguimiento: traer y borrar órdenes.
 *
 * Lee la pestaña "Maestro" (nombre exacto, no la "Maestro (anterior)") de cada
 * archivo "ORDEN DE PRODUCCION ..." que esté en la MISMA carpeta que este
 * Seguimiento, y llena las columnas A a Q de la hoja "Seguimiento".
 * Las columnas R a AD se calculan solas copiando las fórmulas de la fila 5.
 */

var SEG_HOJA = 'Seguimiento';
var SEG_FILA_INICIO = 5;   // fila 5 = plantilla con las fórmulas R:AD (no se borra)
var SEG_COLS_DATOS = 17;   // A..Q
var SEG_COL_FORMULAS = 18; // R
var SEG_NUM_FORMULAS = 13; // R..AD

var MESES_ = {
  ENERO: 0, FEBRERO: 1, MARZO: 2, ABRIL: 3, MAYO: 4, JUNIO: 5, JULIO: 6, AGOSTO: 7,
  SEPTIEMBRE: 8, SETIEMBRE: 8, OCTUBRE: 9, NOVIEMBRE: 10, DICIEMBRE: 11
};

function onOpen() {
  SpreadsheetApp.getUi().createMenu('📋 Órdenes')
    .addItem('📥 Traer órdenes de la carpeta', 'traerOrdenesCarpeta')
    .addItem('🗑️ Borrar órdenes del Seguimiento', 'borrarOrdenesSeguimiento')
    .addSeparator()
    .addItem('🔄 Actualizar todo', 'actualizarTodo')
    .addToUi();
}

function traerOrdenesCarpeta() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName(SEG_HOJA);
  var tz = ss.getSpreadsheetTimeZone();
  var carpeta = DriveApp.getFileById(ss.getId()).getParents().next();

  // Guardar "Estado" y "¿Enviado?" (P y Q) que ya estaban escritos a mano.
  var previos = {};
  var ultima = hoja.getLastRow();
  if (ultima >= SEG_FILA_INICIO) {
    hoja.getRange(SEG_FILA_INICIO, 1, ultima - SEG_FILA_INICIO + 1, SEG_COLS_DATOS).getValues()
      .forEach(function (f) {
        if (f[0] === '' && f[1] === '') return;
        previos[claveFila_(f, tz)] = [f[15], f[16]];
      });
  }

  var filas = [];
  var vistas = {};
  var avisos = [];
  var archivos = carpeta.getFilesByType(MimeType.GOOGLE_SHEETS);
  while (archivos.hasNext()) {
    var arch = archivos.next();
    var nombre = arch.getName();
    if (!/^ORDEN DE PRODUCCI[OÓ]N/i.test(nombre)) continue;

    var libro = SpreadsheetApp.openById(arch.getId());
    var maestro = libro.getSheetByName('Maestro');
    if (!maestro) { avisos.push(nombre + ': no tiene pestaña "Maestro"'); continue; }

    var fecha = fechaDeNombre_(nombre);
    if (!fecha) {
      var h1 = libro.getSheets()[0].getRange('H1').getValue();
      if (h1 instanceof Date) fecha = h1;
    }
    if (!fecha) { avisos.push(nombre + ': no pude leer la fecha'); continue; }

    var n = maestro.getLastRow() - 1;
    if (n < 1) continue;
    maestro.getRange(2, 1, n, 15).getValues().forEach(function (r) {
      // A OP, B MAQ, D Sku, E SKU Comp, F Producto, G Material, H Componente,
      // I Cant, J Color, K Sacos, L Sacos kg, M Ciclo, N Peso, O Cav
      if (r[0] === '' || String(r[3]).trim() === '') return;
      var fila = [fecha, r[0], r[1], String(r[3]), r[4], r[5], r[6], r[7], r[8],
                  limpiarColor_(r[9]), r[10], r[11], r[12], r[13], r[14], '', ''];
      var k = claveFila_(fila, tz);
      if (vistas[k]) return; // misma orden repetida en Parte 1 y Parte 2
      vistas[k] = true;
      if (previos[k]) { fila[15] = previos[k][0]; fila[16] = previos[k][1]; }
      filas.push(fila);
    });
  }

  // Más reciente primero; dentro del día, por número de OP.
  filas.sort(function (a, b) { return (b[0] - a[0]) || (Number(a[1]) - Number(b[1])); });

  // Limpiar lo anterior: A:Q desde la fila 5 y las fórmulas R:AD desde la fila 6.
  ultima = hoja.getLastRow();
  if (ultima >= SEG_FILA_INICIO) {
    hoja.getRange(SEG_FILA_INICIO, 1, ultima - SEG_FILA_INICIO + 1, SEG_COLS_DATOS).clearContent();
    if (ultima > SEG_FILA_INICIO) {
      hoja.getRange(SEG_FILA_INICIO + 1, SEG_COL_FORMULAS, ultima - SEG_FILA_INICIO, SEG_NUM_FORMULAS).clearContent();
    }
  }

  if (filas.length) {
    hoja.getRange(SEG_FILA_INICIO, 4, filas.length, 1).setNumberFormat('@'); // Sku como texto
    hoja.getRange(SEG_FILA_INICIO, 1, filas.length, SEG_COLS_DATOS).setValues(filas);
    hoja.getRange(SEG_FILA_INICIO, 1, filas.length, 1).setNumberFormat('dd/mm/yyyy');
    if (filas.length > 1) {
      hoja.getRange(SEG_FILA_INICIO, SEG_COL_FORMULAS, 1, SEG_NUM_FORMULAS)
        .copyTo(hoja.getRange(SEG_FILA_INICIO + 1, SEG_COL_FORMULAS, filas.length - 1, SEG_NUM_FORMULAS));
    }
  }

  if (typeof actualizarTodo === 'function') {
    try { actualizarTodo(); } catch (e) { avisos.push('Actualizar todo: ' + e.message); }
  }

  var msg = 'Órdenes traídas: ' + filas.length;
  if (avisos.length) msg += '\n\nAvisos:\n' + avisos.join('\n');
  SpreadsheetApp.getUi().alert(msg);
}

function borrarOrdenesSeguimiento() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.alert('Borrar órdenes', '¿Seguro que quieres borrar todas las órdenes del Seguimiento?\n' +
    '(La fila 5 con las fórmulas se queda.)', ui.ButtonSet.YES_NO);
  if (r !== ui.Button.YES) return;

  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SEG_HOJA);
  var ultima = hoja.getLastRow();
  if (ultima < SEG_FILA_INICIO) return;
  hoja.getRange(SEG_FILA_INICIO, 1, ultima - SEG_FILA_INICIO + 1, SEG_COLS_DATOS).clearContent();
  if (ultima > SEG_FILA_INICIO) {
    hoja.getRange(SEG_FILA_INICIO + 1, SEG_COL_FORMULAS, ultima - SEG_FILA_INICIO, SEG_NUM_FORMULAS).clearContent();
  }
}

// "ORDEN DE PRODUCCION 7 DE OCTUBRE DE 2026 Parte 1" -> 07/10/2026
function fechaDeNombre_(nombre) {
  var m = String(nombre).toUpperCase()
    .match(/(\d{1,2})\s+(?:DE\s+)?([A-ZÁÉÍÓÚ]+)\s+(?:DE\s+)?(\d{4})/);
  if (!m || !(m[2] in MESES_)) return null;
  return new Date(Number(m[3]), MESES_[m[2]], Number(m[1]));
}

// "TURQUESA (MB 0689)" -> "TURQUESA", "VERDE PASTEL (SIN CÓDIGO)" -> "VERDE PASTEL"
// para que el Catálogo de Colores lo reconozca. Los nombres viejos quedan igual.
function limpiarColor_(c) {
  return String(c)
    .replace(/\s*\([^()]*\d{4}\)\s*$/, '')
    .replace(/\s*\(SIN C[OÓ]DIGO\)\s*$/i, '');
}

function claveFila_(f, tz) {
  var fecha = f[0] instanceof Date ? Utilities.formatDate(f[0], tz, 'yyyyMMdd') : String(f[0]);
  return [fecha, f[1], String(f[4]).trim(), String(f[9]).trim(), f[8]].join('|');
}
