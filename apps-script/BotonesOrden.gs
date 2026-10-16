/**
 * Botones extra para el menú "Órdenes de Producción".
 *
 *  - boReorganizar: ordena el Maestro por número de OP (sin cambiar números).
 *  - boMoverOP: lleva una OP a otra posición y corre las demás
 *    (ej. la OP 5 pasa a ser la 2; la 2, 3 y 4 pasan a ser 3, 4 y 5).
 *  - boIntercambiarOP: cambia dos OP entre sí (ej. la 3 por la 7).
 *
 * Cada hoja 01..14 muestra la OP con su mismo número, así que al cambiar
 * el número de OP en el Maestro la orden pasa a la hoja correspondiente.
 * Solo se mueven las columnas que se escriben a mano; las de fórmula se
 * recalculan solas.
 *
 * Todo lo interno vive dentro de BO_ para no chocar con otros archivos
 * del proyecto.
 */

function boReorganizar() { BO_.reorganizar(); }
function boMoverOP() { BO_.mover(); }
function boIntercambiarOP() { BO_.intercambiar(); }

var BO_ = (function () {
  var HOJA = 'Maestro';
  var PRIMERA = 2;
  var ULTIMA = 250;
  // A OP, B MAQ, D Sku, E SKU Componente, G Material, I Cantidad, J Color,
  // R ¿Producto importante?
  var COLS = [1, 2, 4, 5, 7, 9, 10, 18];

  function hoja() {
    var h = SpreadsheetApp.getActive().getSheetByName(HOJA);
    if (!h) throw new Error('No se encontró la hoja "' + HOJA + '".');
    return h;
  }

  function leer() {
    var total = ULTIMA - PRIMERA + 1;
    var valores = hoja().getRange(PRIMERA, 1, total, 18).getValues();
    var filas = [];
    valores.forEach(function (fila) {
      var datos = COLS.map(function (c) { return fila[c - 1]; });
      if (datos.some(function (v) { return v !== '' && v !== null; })) filas.push(datos);
    });
    return filas;
  }

  function escribir(filas) {
    var h = hoja();
    var total = ULTIMA - PRIMERA + 1;
    COLS.forEach(function (col, i) {
      var columna = [];
      for (var f = 0; f < total; f++) columna.push([f < filas.length ? filas[f][i] : '']);
      h.getRange(PRIMERA, col, total, 1).setValues(columna);
    });
    SpreadsheetApp.flush();
  }

  function numero(v) {
    if (v === '' || v === null) return null;
    var n = Number(v);
    return isNaN(n) ? null : n;
  }

  /** Orden estable por OP; las filas sin OP quedan al final. */
  function ordenar(filas) {
    return filas
      .map(function (fila, i) { return { fila: fila, i: i, op: numero(fila[0]) }; })
      .sort(function (x, y) {
        if (x.op === null && y.op === null) return x.i - y.i;
        if (x.op === null) return 1;
        if (y.op === null) return -1;
        return x.op - y.op || x.i - y.i;
      })
      .map(function (x) { return x.fila; });
  }

  function opsDe(filas) {
    var vistos = {};
    filas.forEach(function (fila) {
      var op = numero(fila[0]);
      if (op !== null) vistos[op] = true;
    });
    return Object.keys(vistos).map(Number).sort(function (a, b) { return a - b; });
  }

  function renumerar(filas, mapa) {
    filas.forEach(function (fila) {
      var op = numero(fila[0]);
      if (op !== null && mapa.hasOwnProperty(op)) fila[0] = mapa[op];
    });
  }

  function pedir(titulo, mensaje, ops) {
    var ui = SpreadsheetApp.getUi();
    var r = ui.prompt(titulo, mensaje + ' (' + ops.join(', ') + ')', ui.ButtonSet.OK_CANCEL);
    if (r.getSelectedButton() !== ui.Button.OK) return null;
    var op = numero(r.getResponseText().trim());
    if (op === null || ops.indexOf(op) === -1) {
      ui.alert('"' + r.getResponseText() + '" no es una OP del Maestro.');
      return null;
    }
    return op;
  }

  function preparar() {
    var filas = leer();
    var ops = opsDe(filas);
    if (ops.length < 2) {
      SpreadsheetApp.getUi().alert('Se necesitan al menos dos OP en el Maestro.');
      return null;
    }
    return { filas: filas, ops: ops };
  }

  return {
    reorganizar: function () {
      escribir(ordenar(leer()));
      SpreadsheetApp.getActive().toast('Maestro reorganizado por OP.', 'Listo', 4);
    },

    mover: function () {
      var d = preparar();
      if (!d) return;
      var origen = pedir('Mover OP', '¿Qué OP quieres mover?', d.ops);
      if (origen === null) return;
      var destino = pedir('Mover OP', '¿A qué posición la llevas?', d.ops);
      if (destino === null || destino === origen) return;

      var orden = d.ops.filter(function (op) { return op !== origen; });
      orden.splice(d.ops.indexOf(destino), 0, origen);
      var mapa = {};
      orden.forEach(function (op, i) { mapa[op] = d.ops[i]; });

      renumerar(d.filas, mapa);
      escribir(ordenar(d.filas));
      SpreadsheetApp.getActive().toast('La OP ' + origen + ' ahora es la OP ' + destino + '.', 'Listo', 5);
    },

    intercambiar: function () {
      var d = preparar();
      if (!d) return;
      var a = pedir('Intercambiar OP', 'Primera OP', d.ops);
      if (a === null) return;
      var b = pedir('Intercambiar OP', 'Segunda OP', d.ops);
      if (b === null || a === b) return;

      var mapa = {};
      mapa[a] = b;
      mapa[b] = a;
      renumerar(d.filas, mapa);
      escribir(ordenar(d.filas));
      SpreadsheetApp.getActive().toast('Se intercambiaron la OP ' + a + ' y la OP ' + b + '.', 'Listo', 5);
    }
  };
})();
