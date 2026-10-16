/**
 * Botones para el Maestro de la Orden de Producción.
 *
 *  - Reorganizar Maestro: ordena las filas del Maestro por número de OP
 *    (las filas de una misma OP conservan su orden).
 *  - Mover OP: lleva una OP a otra posición y corre las demás
 *    (ej. la OP 5 pasa a ser la 2; la 2, 3 y 4 pasan a ser 3, 4 y 5).
 *  - Intercambiar OP: cambia dos OP entre sí (ej. la 3 por la 7).
 *
 * Cada hoja 01..14 muestra la OP con su mismo número, así que al cambiar
 * el número de OP en el Maestro la orden pasa a la hoja correspondiente.
 *
 * Solo se mueven las columnas que se escriben a mano; las columnas con
 * fórmula se recalculan solas.
 */

var HOJA_MAESTRO = 'Maestro';
var PRIMERA_FILA = 2;
var ULTIMA_FILA = 250;
// Columnas que se escriben a mano: A OP, B MAQ, D Sku, E SKU Componente,
// G Material, I Cantidad, J Color, R ¿Producto importante?
var COLUMNAS_DATOS = [1, 2, 4, 5, 7, 9, 10, 18];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Orden de Producción')
    .addItem('Reorganizar Maestro por OP', 'reorganizarMaestro')
    .addItem('Mover OP a otra posición', 'moverOP')
    .addItem('Intercambiar dos OP', 'intercambiarOP')
    .addToUi();
}

function reorganizarMaestro() {
  var filas = leerFilas_();
  escribirFilas_(ordenarPorOP_(filas));
  SpreadsheetApp.getActive().toast('Maestro reorganizado por OP.', 'Listo', 4);
}

function moverOP() {
  var ui = SpreadsheetApp.getUi();
  var filas = leerFilas_();
  var ops = opsExistentes_(filas);
  if (ops.length < 2) {
    ui.alert('Se necesitan al menos dos OP en el Maestro.');
    return;
  }
  var origen = pedirOP_('Mover OP', '¿Qué OP quieres mover? (' + ops.join(', ') + ')', ops);
  if (origen === null) return;
  var destino = pedirOP_('Mover OP', '¿A qué posición la llevas? (' + ops.join(', ') + ')', ops);
  if (destino === null || destino === origen) return;

  // Nuevo orden: se saca la OP de su lugar y se inserta en la posición destino.
  var orden = ops.filter(function (op) { return op !== origen; });
  orden.splice(ops.indexOf(destino), 0, origen);
  var nuevoNumero = {};
  orden.forEach(function (op, i) { nuevoNumero[op] = ops[i]; });

  aplicarNumeros_(filas, nuevoNumero);
  escribirFilas_(ordenarPorOP_(filas));
  SpreadsheetApp.getActive().toast('La OP ' + origen + ' ahora es la OP ' + destino + '.', 'Listo', 5);
}

function intercambiarOP() {
  var ui = SpreadsheetApp.getUi();
  var filas = leerFilas_();
  var ops = opsExistentes_(filas);
  if (ops.length < 2) {
    ui.alert('Se necesitan al menos dos OP en el Maestro.');
    return;
  }
  var a = pedirOP_('Intercambiar OP', 'Primera OP (' + ops.join(', ') + ')', ops);
  if (a === null) return;
  var b = pedirOP_('Intercambiar OP', 'Segunda OP (' + ops.join(', ') + ')', ops);
  if (b === null || a === b) return;

  var nuevoNumero = {};
  nuevoNumero[a] = b;
  nuevoNumero[b] = a;
  aplicarNumeros_(filas, nuevoNumero);
  escribirFilas_(ordenarPorOP_(filas));
  SpreadsheetApp.getActive().toast('Se intercambiaron la OP ' + a + ' y la OP ' + b + '.', 'Listo', 5);
}

// ---------- auxiliares ----------

function hoja_() {
  var hoja = SpreadsheetApp.getActive().getSheetByName(HOJA_MAESTRO);
  if (!hoja) throw new Error('No se encontró la hoja "' + HOJA_MAESTRO + '".');
  return hoja;
}

/** Devuelve las filas con datos: cada una es la lista de valores de COLUMNAS_DATOS. */
function leerFilas_() {
  var total = ULTIMA_FILA - PRIMERA_FILA + 1;
  var valores = hoja_().getRange(PRIMERA_FILA, 1, total, 18).getValues();
  var filas = [];
  valores.forEach(function (fila) {
    var datos = COLUMNAS_DATOS.map(function (c) { return fila[c - 1]; });
    var tieneAlgo = datos.some(function (v) { return v !== '' && v !== null; });
    if (tieneAlgo) filas.push(datos);
  });
  return filas;
}

/** Escribe las filas desde la fila 2 y deja vacías las que sobran. */
function escribirFilas_(filas) {
  var hoja = hoja_();
  var total = ULTIMA_FILA - PRIMERA_FILA + 1;
  COLUMNAS_DATOS.forEach(function (col, i) {
    var columna = [];
    for (var f = 0; f < total; f++) {
      columna.push([f < filas.length ? filas[f][i] : '']);
    }
    hoja.getRange(PRIMERA_FILA, col, total, 1).setValues(columna);
  });
  SpreadsheetApp.flush();
}

/** Orden estable por OP; las filas sin OP quedan al final. */
function ordenarPorOP_(filas) {
  return filas
    .map(function (fila, i) { return { fila: fila, i: i, op: numeroOP_(fila[0]) }; })
    .sort(function (x, y) {
      if (x.op === null && y.op === null) return x.i - y.i;
      if (x.op === null) return 1;
      if (y.op === null) return -1;
      return x.op - y.op || x.i - y.i;
    })
    .map(function (x) { return x.fila; });
}

function opsExistentes_(filas) {
  var vistos = {};
  filas.forEach(function (fila) {
    var op = numeroOP_(fila[0]);
    if (op !== null) vistos[op] = true;
  });
  return Object.keys(vistos).map(Number).sort(function (a, b) { return a - b; });
}

function aplicarNumeros_(filas, nuevoNumero) {
  filas.forEach(function (fila) {
    var op = numeroOP_(fila[0]);
    if (op !== null && nuevoNumero.hasOwnProperty(op)) fila[0] = nuevoNumero[op];
  });
}

function numeroOP_(valor) {
  if (valor === '' || valor === null) return null;
  var n = Number(valor);
  return isNaN(n) ? null : n;
}

function pedirOP_(titulo, mensaje, ops) {
  var ui = SpreadsheetApp.getUi();
  var r = ui.prompt(titulo, mensaje, ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return null;
  var op = numeroOP_(r.getResponseText().trim());
  if (op === null || ops.indexOf(op) === -1) {
    ui.alert('"' + r.getResponseText() + '" no es una OP del Maestro.');
    return null;
  }
  return op;
}
