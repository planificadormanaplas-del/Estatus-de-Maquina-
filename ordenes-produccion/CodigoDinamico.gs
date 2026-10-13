/**
 * GENERADOR DE ÓRDENES DE PRODUCCIÓN — v23 (Maestro dinámico)
 * ------------------------------------------------------------
 * Gran cambio de esta versión: el Maestro ya NO tiene 30 bloques de 8
 * filas escondidas. Empieza vacío (solo encabezados) y:
 *   - al generar una orden se AGREGAN sus filas (una por color),
 *   - al agregar un color se inserta una fila más en esa orden,
 *   - al borrar una orden se ELIMINAN sus filas.
 * Cada fila lleva el número de OP en la columna A.
 *
 * Para que esto funcione, la hoja ya viene preparada así:
 *   - Las columnas calculadas del Maestro (TON, Producto, Componentes,
 *     Sacos, Sacos kg, Ciclo, Peso, Cav., Und/Hora, Cantidad Final y
 *     Material Recomendado) tienen UNA fórmula en el encabezado (fila 1)
 *     que se aplica sola a todas las filas. No escribas nada debajo de
 *     esas columnas.
 *   - Las hojas 01-14 ya no apuntan a filas fijas: buscan en el Maestro
 *     las filas cuya OP sea igual a su celda H4 (No:). Por eso se pueden
 *     agregar y borrar filas sin romper nada.
 *   - Las filas vacías quedan en blanco: el formato (verde, bordes…) se
 *     pone solo al crear la orden y se quita al borrarla o vaciar.
 *   - "Gestionar → Borrar" pregunta si borrar SOLO la fila (color) donde
 *     estás o la orden completa.
 *   - "Guardar copia en Drive" la guarda directo en la carpeta Seguimiento.
 *   - Máximo 30 OP y 8 colores por OP.
 *   - La lista de Color sale de la hoja "Colores" (catálogo con código).
 *   - Si pegas datos en una orden, las listas desplegables se vuelven a
 *     poner solas; también está el botón "🧩 Reparar listas desplegables".
 *   - La columna Q es UND/DÍA = Und/Hora × 10 h (turno de 6:00 a 16:00).
 *     También es fórmula del encabezado (antes era "Volver al Índice").
 *
 * Ya no hay límite de 30 órdenes. Cada orden sigue teniendo máximo 8
 * colores, porque las hojas 01-14 tienen 8 renglones de color. Solo las
 * OP 1-14 tienen hoja de impresión.
 *
 * Se mantiene todo lo demás de la v22: asistente paso a paso (máquina con
 * aviso de estatus, SKU, componente con genéricos de familia, producto
 * importante, colores), Historial, Organizar (mueve también Tipo de
 * Material y Observaciones de las hojas 01-14), copia en Drive, onEdit.
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📋 Órdenes de Producción')
    .addItem('➕ Generar nueva orden', 'generarNuevaOrden')
    .addItem('➕➕ Generar varias órdenes de una vez', 'generarVariasOrdenes')
    .addItem('🔧 Gestionar esta orden (agregar / borrar)', 'gestionarOrden')
    .addItem('🩹 Borrar SIN guardar en Historial (para corregir)', 'borrarOrdenSinHistorial')
    .addItem('🔄 Organizar (cerrar huecos en la numeración)', 'organizarOrdenes')
    .addItem('🧩 Reparar listas desplegables', 'repararListas')
    .addSeparator()
    .addItem('📁 Guardar copia de hoy en Drive', 'guardarCopiaEnDrive')
    .addSeparator()
    .addItem('🧹 Vaciar el Maestro (borra TODAS las órdenes)', 'vaciarMaestro')
    .addItem('🗑️ Borrar TODO el Historial', 'borrarHistorial')
    .addToUi();
}

// ---- Ajusta esto si tu Maestro cambia de estructura ----
var HOJA_MAESTRO = 'Maestro';
var HOJA_COMPONENTES = 'Componentes ';
var HOJA_MAQ_TONS = 'Maq y Tons ';
var HOJA_LISTA_PRODUCTOS = 'Lista de Productos';
var FILA_INICIO = 2;          // primera fila de datos
var MAX_COLORES = 8;          // colores por orden (renglones de las hojas 01-14)
var MAX_OP = 30;              // número máximo de órdenes (OP 1 a 30)
var HOJAS_DE_ORDEN = 14;      // OP 1-14 tienen hoja propia (01-14)
var FILAS_DE_RESERVA = 50;    // filas en blanco que se dejan siempre al final
var COL_OP = 1;               // A = OP
var COL_MAQ = 2;              // B = Máquina
var COL_SKU = 4;              // D = Sku
var COL_COMPONENTE = 5;       // E = SKu Componente
var COL_PRODUCTO = 6;         // F = Producto (fórmula)
var COL_MATERIAL = 7;         // G = Material
var COL_CANTIDAD = 9;         // I = Cantidad
var COL_COLOR = 10;           // J = Color
var COL_SACOS = 11;           // K = Sacos (fórmula)
var COL_UND_DIA = 17;         // Q = Und/Día (fórmula: Und/Hora × horas del turno)
var COL_IMPORTANTE = 18;      // R = ¿Producto importante?
var COL_CANTIDAD_FINAL = 19;  // S = Cantidad final (fórmula)
var COL_COMPONENTES_CODIGO = 2;          // B en Componentes (código, ej: 0004-0/0)
var COL_COMPONENTES_SKU = 3;             // C en Componentes (a qué SKU pertenece)
var COL_COMPONENTES_NOMBRE = 5;          // E en Componentes (nombre completo)
var COL_LISTA_PRODUCTOS_SKU = 2;         // B en Lista de Productos (código)
var COL_LISTA_PRODUCTOS_NOMBRE = 3;      // C en Lista de Productos (nombre)
var COL_MAQ_ESTATUS = 6;                 // F en Maq y Tons (Estatus Actual)
var ESTATUS_OPERATIVA = 'Operativa';
// Carpeta "Seguimiento" de Drive donde se guardan las copias con fecha.
var CARPETA_COPIAS_ID = '1O0DMmXaRL6dfLTsXqbA4bdx7aIYm7Wwh';
// ----------------------------------------------------------

// Columnas que se escriben a mano (o desde el asistente). Las demás son
// fórmulas del encabezado y NO se deben tocar.
var COLUMNAS_ENTRADA = [COL_OP, COL_MAQ, COL_SKU, COL_COMPONENTE, COL_MATERIAL, COL_CANTIDAD, COL_COLOR, COL_IMPORTANTE];

// Un SKU escrito como "0" es un marcador de "por asignar todavía".
function esPlaceholderSku(valor) {
  return valor === 0 || valor === '0';
}

// ---------- Lectura de las órdenes del Maestro ----------

// Última fila que tiene un número de OP en la columna A (o 1 si no hay).
function ultimaFilaConDatos(sheet) {
  var total = sheet.getMaxRows() - 1;
  if (total < 1) return 1;
  var ops = sheet.getRange(FILA_INICIO, COL_OP, total, 1).getValues();
  for (var i = ops.length - 1; i >= 0; i--) {
    if (ops[i][0] !== '') return FILA_INICIO + i;
  }
  return 1;
}

// Devuelve las órdenes en el orden en que están en la hoja:
// [{op, fila, n, filas}] — fila = primera fila, n = cuántas filas (colores),
// filas = valores A..S de esas filas (filas[i][COL_SKU - 1], etc.).
function leerOrdenes(sheet) {
  var ultima = ultimaFilaConDatos(sheet);
  if (ultima < FILA_INICIO) return [];
  var datos = sheet.getRange(FILA_INICIO, 1, ultima - FILA_INICIO + 1, COL_CANTIDAD_FINAL).getValues();
  var ordenes = [];
  var actual = null;
  for (var i = 0; i < datos.length; i++) {
    var op = datos[i][COL_OP - 1];
    if (op === '') { actual = null; continue; }
    if (actual && actual.op === op && actual.fila + actual.n === FILA_INICIO + i) {
      actual.n++;
      actual.filas.push(datos[i]);
    } else {
      actual = { op: op, fila: FILA_INICIO + i, n: 1, filas: [datos[i]] };
      ordenes.push(actual);
    }
  }
  return ordenes;
}

function buscarOrden(ordenes, op) {
  for (var i = 0; i < ordenes.length; i++) {
    if (ordenes[i].op === op) return ordenes[i];
  }
  return null;
}

// La orden que contiene la celda activa del Maestro, o null (después de avisar).
function ordenDeCeldaActiva(ui, sheet, accion) {
  if (sheet.getName() !== HOJA_MAESTRO) {
    ui.alert('Primero ponte en una celda dentro del Maestro, en la orden que quieres ' + accion + '.');
    return null;
  }
  var fila = sheet.getActiveCell().getRow();
  var ordenes = leerOrdenes(sheet);
  for (var i = 0; i < ordenes.length; i++) {
    var o = ordenes[i];
    if (fila >= o.fila && fila < o.fila + o.n) return o;
  }
  ui.alert('Ponte en una fila que tenga una orden (con número de OP en la columna A).');
  return null;
}

// Se asegura de que siempre queden filas en blanco al final del Maestro.
function asegurarReserva(sheet, filasNecesarias) {
  var libres = sheet.getMaxRows() - ultimaFilaConDatos(sheet);
  var faltan = (filasNecesarias || 0) + FILAS_DE_RESERVA - libres;
  if (faltan > 0) {
    sheet.insertRowsAfter(sheet.getMaxRows(), faltan);
    limpiarFormatoReserva(sheet);
  }
}

// Borra filas del Maestro sin dejarlo nunca sin filas debajo del encabezado.
function eliminarFilas(sheet, fila, n) {
  if (sheet.getMaxRows() - (fila + n - 1) < 1) {
    sheet.insertRowsAfter(sheet.getMaxRows(), FILAS_DE_RESERVA);
  }
  sheet.deleteRows(fila, n);
  asegurarReserva(sheet, 0);
  limpiarFormatoReserva(sheet);
}

// Escribe una columna con el mismo valor en n filas.
function escribirColumna(sheet, fila, col, n, valor) {
  var m = [];
  for (var i = 0; i < n; i++) m.push([valor]);
  sheet.getRange(fila, col, n, 1).setValues(m);
}

// ---------- Formato de las filas ----------

// Las filas vacías del Maestro quedan en blanco. El COLOR de fondo lo pone
// solo un formato condicional de la hoja: verde/gris/crema si la fila tiene
// número de OP en la columna A, blanco si no (aunque se borre a mano con
// Suprimir). Al crear una orden, el código pone además bordes, letra y
// formato de números. Una entrada por columna, de A a T ("fondo" queda
// solo como referencia de los colores de la regla condicional).
var VERDE = '#D9F2D9', GRIS = '#EDEDED', CREMA = '#EAF6FC';
var FORMATO_COLUMNAS = [
  { fondo: VERDE, alinear: 'center', numero: '0',        tam: 10 }, // A OP
  { fondo: VERDE, alinear: 'center', numero: '0',        tam: 10 }, // B MAQ (número, para que TON la encuentre)
  { fondo: VERDE, alinear: 'center', numero: '0',        tam: 10 }, // C TON
  { fondo: VERDE, alinear: 'center', numero: '@',        tam: 11 }, // D Sku
  { fondo: VERDE, alinear: 'center', numero: '@',        tam: 11 }, // E Componente
  { fondo: VERDE, alinear: 'left',   numero: '@',        tam: 10 }, // F Producto
  { fondo: VERDE, alinear: 'left',   numero: '@',        tam: 10 }, // G Material
  { fondo: GRIS,  alinear: 'left',   numero: '@',        tam: 9  }, // H Componentes
  { fondo: VERDE, alinear: 'right',  numero: '#,##0.00', tam: 10 }, // I Cantidad
  { fondo: VERDE, alinear: 'left',   numero: '@',        tam: 9  }, // J Color
  { fondo: VERDE, alinear: 'right',  numero: '#,##0.00', tam: 10 }, // K Sacos
  { fondo: VERDE, alinear: 'right',  numero: '#,##0.00', tam: 10 }, // L Sacos kg
  { fondo: VERDE, alinear: 'right',  numero: '#,##0.00', tam: 10 }, // M Ciclo
  { fondo: VERDE, alinear: 'right',  numero: '#,##0.00', tam: 10 }, // N Peso
  { fondo: VERDE, alinear: 'right',  numero: '#,##0.00', tam: 10 }, // O Cav.
  { fondo: GRIS,  alinear: 'right',  numero: '#,##0',    tam: 11 }, // P Und/Hora
  { fondo: GRIS,  alinear: 'right',  numero: '#,##0',    tam: 11 }, // Q Und/Día
  { fondo: null,  alinear: 'center', numero: '@',        tam: 11 }, // R ¿Importante?
  { fondo: CREMA, alinear: 'center', numero: '#,##0',    tam: 10 }, // S Cantidad final
  { fondo: null,  alinear: 'center', numero: '@',        tam: 11 }  // T Material recomendado
];

// Da el formato de orden a n filas (desde "fila").
function aplicarFormatoFilas(sheet, fila, n) {
  var rango = sheet.getRange(fila, 1, n, FORMATO_COLUMNAS.length);
  var repetir = function (campo) {
    var fila1 = FORMATO_COLUMNAS.map(function (f) { return f[campo]; });
    var m = [];
    for (var i = 0; i < n; i++) m.push(fila1.slice());
    return m;
  };
  rango.setBackground(null)
    .setHorizontalAlignments(repetir('alinear'))
    .setNumberFormats(repetir('numero'))
    .setFontSizes(repetir('tam'))
    .setFontFamily('Arial')
    .setFontWeight('bold')
    .setVerticalAlignment('middle')
    .setWrap(true)
    .setBorder(true, true, true, true, true, true);
}

// Listas desplegables de cada fila de orden (las mismas del Maestro original).
// El Componente (E) tiene su propia lista según el SKU.
var LISTAS_FILA = [
  { col: COL_MAQ,        rango: ['Maq y Tons ', 'A4:A'] },
  { col: COL_SKU,        rango: ['Lista de Productos', 'B2:B'] },
  { col: COL_MATERIAL,   rango: ['Materia Prima ', 'F3:F'] },
  { col: COL_COLOR,      rango: ['Colores', 'A2:A'] },   // catálogo de colores (nombre + código)
  { col: COL_IMPORTANTE, lista: ['Sí', 'No'] }
];

function aplicarListasFilas(sheet, fila, n) {
  var ss = sheet.getParent();
  LISTAS_FILA.forEach(function (l) {
    var regla = SpreadsheetApp.newDataValidation().setAllowInvalid(false);
    if (l.lista) {
      regla.requireValueInList(l.lista, true);
    } else {
      var hoja = ss.getSheetByName(l.rango[0]);
      if (!hoja) return;
      regla.requireValueInRange(hoja.getRange(l.rango[1]), true);
    }
    sheet.getRange(fila, l.col, n, 1).setDataValidation(regla.build());
  });
  sheet.getRange(fila, COL_CANTIDAD, n, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireNumberGreaterThanOrEqualTo(0).setAllowInvalid(false).build());
}

// Deja en blanco (sin formato ni listas) todas las filas debajo de la última orden.
function limpiarFormatoReserva(sheet) {
  var ultima = ultimaFilaConDatos(sheet);
  var resto = sheet.getMaxRows() - ultima;
  if (resto < 1) return;
  var rango = sheet.getRange(ultima + 1, 1, resto, sheet.getMaxColumns());
  rango.clearFormat();
  rango.clearDataValidations();
}

// ---------- Número de OP ----------

// El primer número de OP que no esté usado (ni en la hoja ni en "excluir").
function siguienteOPLibre(ordenes, excluir) {
  var usados = {};
  ordenes.forEach(function (o) { usados[o.op] = true; });
  for (var n = 1; n <= MAX_OP; n++) {
    if (!usados[n] && !excluir[n]) return n;
  }
  return null; // ya están usadas las 30
}

// Pregunta qué número de OP usar. 0 o en blanco = el siguiente libre.
// Devuelve el número, o null si cancelaron.
function pedirNumeroOP(ui, sheet, excluir) {
  excluir = excluir || {};
  while (true) {
    var ordenes = leerOrdenes(sheet);
    var sugerido = siguienteOPLibre(ordenes, excluir);
    if (sugerido === null) {
      ui.alert('Ya hay ' + MAX_OP + ' órdenes (el máximo). Borra o termina alguna antes de crear otra.');
      return null;
    }
    var resp = ui.prompt(
      'Número de orden (OP)',
      '¿Qué número de OP quieres usar? Escribe 0, o déjalo en blanco, para usar el siguiente libre (' + sugerido + ').\n' +
      '(Solo las OP 1 a ' + HOJAS_DE_ORDEN + ' tienen hoja de impresión.)',
      ui.ButtonSet.OK_CANCEL
    );
    if (resp.getSelectedButton() !== ui.Button.OK) return null;

    var texto = resp.getResponseText().trim();
    if (texto === '' || texto === '0') return sugerido;

    var numero = parseInt(texto, 10);
    if (isNaN(numero) || texto !== numero.toString() || numero < 1 || numero > MAX_OP) {
      ui.alert('Escribe un número de OP del 1 al ' + MAX_OP + ', o 0 / en blanco para automático.');
      continue;
    }
    if (excluir[numero]) {
      ui.alert('Ya usaste la OP ' + numero + ' en esta misma tanda. Elige otro número.');
      continue;
    }
    if (buscarOrden(ordenes, numero)) {
      ui.alert('La OP ' + numero + ' ya existe. Elige otro número, o bórrala primero si quieres reusarla.');
      continue;
    }
    if (numero > HOJAS_DE_ORDEN) {
      var seguir = ui.alert('La OP ' + numero + ' no tiene hoja de impresión (solo hay de la 1 a la ' + HOJAS_DE_ORDEN + '). ¿Usarla de todos modos?', ui.ButtonSet.YES_NO);
      if (seguir !== ui.Button.YES) continue;
    }
    return numero;
  }
}

// ---------- Paso 1: Máquina ----------

// Devuelve [{numero, estatus}] desde A4 hasta el final de "Maq y Tons".
function listaMaquinasValidas() {
  var hoja = encontrarHoja(HOJA_MAQ_TONS, ['maq', 'ton']);
  if (!hoja) return [];
  var ultimaFila = hoja.getLastRow();
  if (ultimaFila < 4) return [];
  var datos = hoja.getRange(4, 1, ultimaFila - 3, COL_MAQ_ESTATUS).getValues();
  var lista = [];
  for (var i = 0; i < datos.length; i++) {
    if (datos[i][0] === '') continue;
    lista.push({
      numero: datos[i][0].toString().trim(),
      estatus: (datos[i][COL_MAQ_ESTATUS - 1] || '').toString().trim()
    });
  }
  return lista;
}

// Busca una hoja por nombre exacto primero; si no la encuentra, busca
// cualquier hoja cuyo nombre contenga todas las palabras clave (sin
// importar mayúsculas/espacios). Así no se rompe si el nombre real
// tiene un espacio de más o de menos.
function encontrarHoja(nombreExacto, palabrasClave) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var directa = ss.getSheetByName(nombreExacto);
  if (directa) return directa;

  var hojas = ss.getSheets();
  for (var i = 0; i < hojas.length; i++) {
    var nombre = hojas[i].getName().toLowerCase();
    var coincideTodo = true;
    for (var k = 0; k < palabrasClave.length; k++) {
      if (nombre.indexOf(palabrasClave[k].toLowerCase()) === -1) { coincideTodo = false; break; }
    }
    if (coincideTodo) return hojas[i];
  }
  return null;
}

// Devuelve el número de máquina como texto, o null si cancelaron
function pedirMaquina(ui) {
  var maquinas = listaMaquinasValidas();
  var validas = maquinas.map(function (m) { return m.numero; });
  if (validas.length === 0) {
    ui.alert('⚠️ Aviso', 'No pude encontrar la lista de máquinas (revisa que la hoja "Maq y Tons" exista y tenga datos desde la fila 4). Puedes escribir el número de máquina de todos modos, no se va a validar contra ninguna lista.', ui.ButtonSet.OK);
  }
  var listaTexto = validas.length > 0 ? ('Máquinas disponibles:\n' + validas.join(', ')) : '';
  var primeraVez = true;
  while (true) {
    var mensaje = primeraVez
      ? (listaTexto + '\n\n¿Cuál es el número de máquina para esta orden?')
      : ('Esa máquina no existe. Inténtalo otra vez.\n\n' + listaTexto + '\n\n¿Cuál es el número de máquina?');
    var resp = ui.prompt('Máquina', mensaje, ui.ButtonSet.OK_CANCEL);
    if (resp.getSelectedButton() !== ui.Button.OK) return null;
    var valor = resp.getResponseText().trim();
    if (valor === '') { primeraVez = false; continue; }
    if (validas.length === 0) return valor;

    var idx = validas.indexOf(valor);
    if (idx === -1) { primeraVez = false; continue; }

    var estatus = maquinas[idx].estatus;
    if (estatus && estatus !== ESTATUS_OPERATIVA) {
      var seguir = ui.alert(
        '⚠️ Máquina ' + valor,
        'Según "Maq y Tons", el estatus de esta máquina es: "' + estatus + '".\n\n¿Quieres usarla de todos modos?',
        ui.ButtonSet.YES_NO
      );
      if (seguir !== ui.Button.YES) { primeraVez = true; continue; }
    }
    return valor;
  }
}

// ---------- Paso 2: Producto / SKU (con búsqueda) ----------

function buscarSkus(texto) {
  var hoja = encontrarHoja(HOJA_LISTA_PRODUCTOS, ['lista', 'producto']);
  if (!hoja) return [];
  var ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return [];
  var datos = hoja.getRange(2, COL_LISTA_PRODUCTOS_SKU, ultimaFila - 1, 2).getValues();
  var textoUp = texto.toUpperCase();
  var resultados = [];
  for (var i = 0; i < datos.length; i++) {
    var codigo = datos[i][0];
    var nombre = datos[i][1];
    if (!codigo) continue;
    var codigoStr = codigo.toString().trim();
    var nombreStr = (nombre || '').toString();
    if (codigoStr.toUpperCase().indexOf(textoUp) !== -1 || nombreStr.toUpperCase().indexOf(textoUp) !== -1) {
      resultados.push({ codigo: codigoStr, nombre: nombreStr.trim() });
    }
  }
  return resultados;
}

// Devuelve el codigo de SKU elegido, o null si cancelaron
function pedirSku(ui) {
  while (true) {
    var resp = ui.prompt(
      'Producto (SKU)',
      'Escribe el código del producto, o parte del nombre para buscarlo (ej: "jarra" o "0048-0"):',
      ui.ButtonSet.OK_CANCEL
    );
    if (resp.getSelectedButton() !== ui.Button.OK) return null;
    var texto = resp.getResponseText().trim();
    if (texto === '') continue;

    var resultados = buscarSkus(texto);
    if (resultados.length === 0) {
      ui.alert('No encontré ningún producto con "' + texto + '". Vamos a intentarlo otra vez.');
      continue;
    }

    var exacto = null;
    for (var i = 0; i < resultados.length; i++) {
      if (resultados[i].codigo.toUpperCase() === texto.toUpperCase()) { exacto = resultados[i]; break; }
    }
    if (exacto) return exacto.codigo;
    if (resultados.length === 1) return resultados[0].codigo;

    var mostrar = resultados.slice(0, 20);
    var lista = mostrar.map(function (r) { return r.codigo + ' — ' + r.nombre; }).join('\n');
    var extra = resultados.length > 20 ? ('\n… y ' + (resultados.length - 20) + ' más. Escribe algo más específico si no está en la lista.') : '';

    var resp2 = ui.prompt(
      'Encontré ' + resultados.length + ' productos',
      lista + extra + '\n\nEscribe el código EXACTO del que quieres:',
      ui.ButtonSet.OK_CANCEL
    );
    if (resp2.getSelectedButton() !== ui.Button.OK) return null;
    var elegido = resp2.getResponseText().trim();

    var valido = null;
    for (var j = 0; j < resultados.length; j++) {
      if (resultados[j].codigo.toUpperCase() === elegido.toUpperCase()) { valido = resultados[j]; break; }
    }
    if (valido) return valido.codigo;
    ui.alert('Ese código no estaba en la lista que te mostré. Empecemos de nuevo la búsqueda.');
  }
}

// ---------- Paso 3: Componente (filtrado por SKU) ----------

// Busca en la hoja "Componentes " todas las filas cuya columna Sku (C)
// coincida EXACTO con el sku dado.
function buscarComponentesPorSkuExacto(sku) {
  var hoja = encontrarHoja(HOJA_COMPONENTES, ['componente']);
  if (!hoja) return [];
  var ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return [];
  var datos = hoja.getRange(2, 1, ultimaFila - 1, 5).getValues(); // A:E
  var skuBuscado = sku.toString().trim();
  var opciones = [];
  for (var i = 0; i < datos.length; i++) {
    var codigo = datos[i][COL_COMPONENTES_CODIGO - 1];
    var skuFila = datos[i][COL_COMPONENTES_SKU - 1];
    var nombre = datos[i][COL_COMPONENTES_NOMBRE - 1];
    if (codigo && skuFila && skuFila.toString().trim() === skuBuscado) {
      opciones.push({ codigo: codigo.toString().trim(), nombre: (nombre || '').toString().trim() });
    }
  }
  return opciones;
}

// La "familia" de un SKU es todo lo que va antes del último guion.
// Ej: "6536-1" -> "6536"   |   "0004-0" -> "0004"
function familiaDeSku(sku) {
  var idx = sku.lastIndexOf('-');
  if (idx === -1) return sku;
  return sku.substring(0, idx);
}

// 1) Componentes cuyo Sku sea EXACTO igual al SKU de la orden.
// 2) Si no hay, los "genéricos" de la familia (Sku "FAMILIA-0"), caso de
//    maceteros o mesas donde SKU-1, SKU-2... comparten el mismo molde.
// Devuelve {opciones, generico}.
function opcionesDeComponenteConOrigen(sku) {
  var directas = buscarComponentesPorSkuExacto(sku);
  if (directas.length > 0) return { opciones: directas, generico: false };

  var familia = familiaDeSku(sku.toString().trim()) + '-0';
  return { opciones: buscarComponentesPorSkuExacto(familia), generico: true };
}

function opcionesDeComponente(sku) {
  return opcionesDeComponenteConOrigen(sku).opciones;
}

// Devuelve el codigo de componente elegido, '' si no hay ninguno, o null si cancelaron
function pedirComponente(ui, sku) {
  var resultado = opcionesDeComponenteConOrigen(sku);
  var opciones = resultado.opciones;
  var aviso = resultado.generico
    ? ('(Este SKU no tiene componente propio — son los componentes GENÉRICOS que comparte toda la familia "' + familiaDeSku(sku) + '")\n\n')
    : '';

  if (opciones.length === 0) {
    ui.alert('No encontré componentes para el SKU "' + sku + '" (ni exactos ni genéricos de la familia). Lo vas a tener que elegir luego directo en la hoja.');
    return '';
  }
  if (opciones.length === 1) {
    if (resultado.generico) {
      ui.alert('Componente asignado (genérico de familia): ' + opciones[0].codigo + ' — ' + opciones[0].nombre);
    }
    return opciones[0].codigo;
  }

  var lista = opciones.map(function (o, idx) { return (idx + 1) + ') ' + o.codigo + ' — ' + o.nombre; }).join('\n');
  var primeraVez = true;
  while (true) {
    var mensaje = primeraVez
      ? (aviso + 'Componentes de este SKU:\n\n' + lista + '\n\nEscribe el NÚMERO de la lista (ej: 1):')
      : ('Ese número no está en la lista. Inténtalo otra vez:\n\n' + lista + '\n\nEscribe el NÚMERO de la lista:');
    var resp = ui.prompt('Componente', mensaje, ui.ButtonSet.OK_CANCEL);
    if (resp.getSelectedButton() !== ui.Button.OK) return null;
    var elegido = resp.getResponseText().trim();

    var numero = parseInt(elegido, 10);
    if (!isNaN(numero) && elegido === numero.toString() && numero >= 1 && numero <= opciones.length) {
      return opciones[numero - 1].codigo;
    }

    // Por si prefiere escribir el código completo, también funciona
    for (var j = 0; j < opciones.length; j++) {
      if (opciones[j].codigo.toUpperCase() === elegido.toUpperCase()) return opciones[j].codigo;
    }
    primeraVez = false;
  }
}

function reglaValidacionComponente(opciones) {
  if (opciones.length === 0) return null;
  var codigos = opciones.map(function (o) { return o.codigo; });
  return SpreadsheetApp.newDataValidation()
    .requireValueInList(codigos, true)
    .setAllowInvalid(false)
    .build();
}

function aplicarValidacionComponente(rango, opciones) {
  var regla = reglaValidacionComponente(opciones);
  if (regla) rango.setDataValidation(regla);
}

// ---------- Paso 4: ¿Producto importante? ----------

// Devuelve 'Sí' o 'No'
function pedirImportante(ui) {
  var resp = ui.alert(
    'Producto importante',
    '¿Este producto/color es de los que NUNCA se pueden redondear hacia abajo?\n' +
    '(o sea, aunque sobren pocos gramos, siempre hay que subir un saco más)\n\n' +
    'Sí = nunca bajar el saco\nNo = redondeo normal (el más cercano)',
    ui.ButtonSet.YES_NO
  );
  return resp === ui.Button.YES ? 'Sí' : 'No';
}

// ---------- Paso 5: cuántos colores ----------

// Devuelve un número entre 1 y MAX_COLORES, o null si cancelaron.
// Si escriben algo inválido, vuelve a preguntar.
function pedirNumeroColores(ui, titulo) {
  var mensaje = '¿Cuántos colores va a tener esta orden? (1 a ' + MAX_COLORES + ')';
  while (true) {
    var resp = ui.prompt(titulo, mensaje, ui.ButtonSet.OK_CANCEL);
    if (resp.getSelectedButton() !== ui.Button.OK) return null;
    var texto = resp.getResponseText().trim();
    var n = parseInt(texto, 10);
    if (!isNaN(n) && texto === n.toString() && n >= 1 && n <= MAX_COLORES) return n;
    mensaje = 'Escribe solo un número entre 1 y ' + MAX_COLORES + '.\n\n¿Cuántos colores va a tener esta orden?';
  }
}

// ---------- Menú: generar orden(es) ----------

// Inserta las filas de una orden nueva en su lugar (ordenadas por OP) y
// llena sus datos. Devuelve la primera fila.
function insertarOrden(sheet, op, n, maquina, sku, componente, importante) {
  var ordenes = leerOrdenes(sheet);
  var fila = null;
  for (var i = 0; i < ordenes.length; i++) {
    if (ordenes[i].op > op) { fila = ordenes[i].fila; break; }
  }
  if (fila) {
    sheet.insertRowsBefore(fila, n);
  } else {
    asegurarReserva(sheet, n);
    fila = ultimaFilaConDatos(sheet) + 1;
  }
  llenarFilasDeOrden(sheet, fila, n, op, maquina, sku, componente, importante);
  return fila;
}

// Escribe OP, Máquina, Sku, Componente (con su lista) e Importante en n
// filas, y deja vacías las demás columnas de entrada.
function llenarFilasDeOrden(sheet, fila, n, op, maquina, sku, componente, importante) {
  aplicarFormatoFilas(sheet, fila, n);
  aplicarListasFilas(sheet, fila, n);
  for (var c = 0; c < COLUMNAS_ENTRADA.length; c++) {
    sheet.getRange(fila, COLUMNAS_ENTRADA[c], n, 1).clearContent();
  }
  escribirColumna(sheet, fila, COL_OP, n, op);
  var maqNumero = Number(maquina);
  escribirColumna(sheet, fila, COL_MAQ, n, (maquina !== '' && !isNaN(maqNumero)) ? maqNumero : maquina);
  escribirColumna(sheet, fila, COL_SKU, n, sku);

  var rangoComp = sheet.getRange(fila, COL_COMPONENTE, n, 1);
  rangoComp.clearDataValidations();
  if (sku) aplicarValidacionComponente(rangoComp, opcionesDeComponente(sku.toString()));
  if (componente) escribirColumna(sheet, fila, COL_COMPONENTE, n, componente);
  if (importante) escribirColumna(sheet, fila, COL_IMPORTANTE, n, importante);
}

// Hace todas las preguntas de una orden y la escribe. Devuelve {op, fila,
// texto}, o null si cancelaron.
function asistenteUnaOrden(ui, ss, sheet, excluir, tituloColores) {
  var op = pedirNumeroOP(ui, sheet, excluir);
  if (op === null) return null;

  var maquina = pedirMaquina(ui);
  if (maquina === null) return null;

  var sku = pedirSku(ui);
  if (sku === null) return null;

  var componente = pedirComponente(ui, sku);
  if (componente === null) return null;

  var importante = pedirImportante(ui);

  var n = pedirNumeroColores(ui, tituloColores + ' (OP ' + op + ')');
  if (n === null) return null;

  var fila = insertarOrden(sheet, op, n, maquina, sku, componente, importante);
  limpiarTipoDeMaterialDeHoja(ss, op);
  excluir[op] = true;

  return { op: op, fila: fila, texto: 'OP ' + op + ' (máquina ' + maquina + ', SKU ' + sku + '): ' + n + ' fila(s)' };
}

function generarNuevaOrden() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(HOJA_MAESTRO);
  if (!sheet) { ui.alert('No encontré la hoja "' + HOJA_MAESTRO + '".'); return; }

  var r = asistenteUnaOrden(ui, ss, sheet, {}, 'Nueva orden');
  if (!r) return;

  sheet.setActiveRange(sheet.getRange(r.fila, COL_CANTIDAD));
  ss.toast(r.texto + '. Solo falta poner Cantidad y Color en cada fila.', '✅ Nueva orden creada', 10);
}

function generarVariasOrdenes() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(HOJA_MAESTRO);
  if (!sheet) { ui.alert('No encontré la hoja "' + HOJA_MAESTRO + '".'); return; }

  var respCantidad = ui.prompt('Generar varias órdenes', '¿Cuántas órdenes nuevas vas a crear?', ui.ButtonSet.OK_CANCEL);
  if (respCantidad.getSelectedButton() !== ui.Button.OK) return;
  var totalOrdenes = parseInt(respCantidad.getResponseText(), 10);
  if (isNaN(totalOrdenes) || totalOrdenes < 1) { ui.alert('Escribe un número válido.'); return; }

  var resumen = [];
  var usadas = {};
  for (var k = 0; k < totalOrdenes; k++) {
    ui.alert('Orden ' + (k + 1) + ' de ' + totalOrdenes);
    var r = asistenteUnaOrden(ui, ss, sheet, usadas, 'Orden ' + (k + 1) + ' de ' + totalOrdenes);
    if (!r) break;
    resumen.push(r.texto);
  }

  if (resumen.length > 0) {
    ui.alert('Listo:\n\n' + resumen.join('\n'));
  }
}

// ---------- Menú: gestionar una orden existente ----------

function gestionarOrden() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();

  var orden = ordenDeCeldaActiva(ui, sheet, 'gestionar');
  if (!orden) return;

  var resp = ui.alert(
    'Orden OP ' + orden.op,
    '¿Qué quieres hacer con esta orden?\n\n' +
    'Sí → Agregar otro color/producto a esta orden\n' +
    'No → Borrar (te pregunto si solo una fila o toda la orden)\n' +
    'Cancelar → No hacer nada',
    ui.ButtonSet.YES_NO_CANCEL
  );

  if (resp === ui.Button.YES) {
    agregarColor(ui, ss, sheet, orden);
  } else if (resp === ui.Button.NO) {
    var filaActiva = sheet.getActiveCell().getRow();
    if (orden.n === 1) {
      borrarOrdenCompleta(ui, ss, sheet, orden);
      return;
    }
    var color = orden.filas[filaActiva - orden.fila][COL_COLOR - 1];
    var cual = ui.alert(
      '🗑️ Borrar en la OP ' + orden.op,
      '¿Qué quieres borrar?\n\n' +
      'Sí → Solo la fila ' + filaActiva + (color ? ' (color ' + color + ')' : '') + '\n' +
      'No → La orden COMPLETA (' + orden.n + ' filas)\n' +
      'Cancelar → No borrar nada',
      ui.ButtonSet.YES_NO_CANCEL
    );
    if (cual === ui.Button.YES) {
      borrarUnaFila(ui, ss, sheet, orden, filaActiva);
    } else if (cual === ui.Button.NO) {
      borrarOrdenCompleta(ui, ss, sheet, orden);
    }
  }
}

function agregarColor(ui, ss, sheet, orden) {
  if (orden.n >= MAX_COLORES) {
    ui.alert('Esta orden ya tiene los ' + MAX_COLORES + ' colores que caben en su hoja de impresión.');
    return;
  }
  var primera = orden.filas[0];
  var filaNueva = orden.fila + orden.n;
  sheet.insertRowsAfter(filaNueva - 1, 1);
  llenarFilasDeOrden(sheet, filaNueva, 1, orden.op, primera[COL_MAQ - 1], primera[COL_SKU - 1],
                     primera[COL_COMPONENTE - 1], primera[COL_IMPORTANTE - 1]);
  sheet.setActiveRange(sheet.getRange(filaNueva, COL_CANTIDAD));
  ss.toast('Se agregó una fila más a la orden OP ' + orden.op + '. Ya tiene el mismo SKU/Componente, solo falta Cantidad y Color.', '✅', 8);
}

// Borra solo una fila (un color) de la orden, guardándola antes en el Historial.
function borrarUnaFila(ui, ss, sheet, orden, fila) {
  var datos = orden.filas[fila - orden.fila];
  var confirmar = ui.alert(
    '🗑️ Borrar fila ' + fila,
    'Se va a guardar esta fila en el Historial y se va a eliminar del Maestro. ' +
    'La OP ' + orden.op + ' queda con ' + (orden.n - 1) + ' fila(s). ¿Seguro?',
    ui.ButtonSet.YES_NO
  );
  if (confirmar !== ui.Button.YES) return;
  guardarEnHistorial(filasParaHistorial([datos], orden.op, 'Borrado de una fila', new Date()));
  eliminarFilas(sheet, fila, 1);
  ss.toast('Se borró la fila ' + fila + ' de la OP ' + orden.op + '. El resto de la orden sigue igual.', '🗑️ Listo', 6);
}

function borrarOrdenCompleta(ui, ss, sheet, orden) {
  var confirmar = ui.alert(
    '🗑️ Borrar orden OP ' + orden.op,
    'Esto va a guardar la orden en el Historial y eliminar sus ' + orden.n + ' fila(s) del Maestro.\n\n¿Seguro?',
    ui.ButtonSet.YES_NO
  );
  if (confirmar !== ui.Button.YES) return;
  guardarEnHistorial(filasParaHistorial(orden.filas, orden.op, 'Borrado individual', new Date()));
  eliminarFilas(sheet, orden.fila, orden.n);
  limpiarTipoDeMaterialDeHoja(ss, orden.op);
  limpiarObservacionesDeHoja(ss, orden.op);
  ss.toast('La orden OP ' + orden.op + ' quedó guardada en el Historial y borrada del Maestro.', '🗑️ Listo', 6);
}

// ---------- Borrar sin dejar rastro en el Historial (para corregir errores) ----------

function borrarOrdenSinHistorial() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();

  var orden = ordenDeCeldaActiva(ui, sheet, 'borrar');
  if (!orden) return;

  var confirmar = ui.alert(
    '🩹 Borrar orden OP ' + orden.op + ' (sin Historial)',
    'Esto elimina la orden SIN guardar copia en el Historial, y renumera las órdenes ' +
    'siguientes para que no quede un hueco en la numeración.\n\n' +
    'Úsalo solo para corregir un error mientras la estás llenando (SKU equivocado, etc). ' +
    'Si la orden ya estaba terminada o en proceso de verdad, mejor usa "🔧 Gestionar esta orden".\n\n' +
    '¿Seguro que quieres continuar?',
    ui.ButtonSet.YES_NO
  );
  if (confirmar !== ui.Button.YES) return;

  eliminarFilas(sheet, orden.fila, orden.n);
  limpiarTipoDeMaterialDeHoja(ss, orden.op);
  limpiarObservacionesDeHoja(ss, orden.op);
  var movimientos = compactarOrdenes(ss, sheet);
  var extra = movimientos.length > 0 ? (' Se renumeraron ' + movimientos.length + ' orden(es) para cerrar el hueco.') : '';
  ss.toast('La orden OP ' + orden.op + ' quedó borrada (sin registrar en Historial).' + extra, '🩹 Listo', 8);
}

// ---------- Organizar: cerrar huecos en la numeración ----------

function organizarOrdenes() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(HOJA_MAESTRO);
  if (!sheet) { ui.alert('No encontré la hoja "' + HOJA_MAESTRO + '".'); return; }

  var confirmar = ui.alert(
    '🔄 Organizar órdenes',
    'Esto ordena las órdenes por número y las renumera 1, 2, 3… para cerrar los huecos ' +
    '(ej: si no hay OP 2, la OP 3 pasa a ser la OP 2).\n\n' +
    'El "Tipo de Material" y las "Observaciones" de las hojas 01-14 se mueven junto con su orden.\n\n' +
    'Los números de las órdenes que se muevan van a cambiar. ¿Seguro?',
    ui.ButtonSet.YES_NO
  );
  if (confirmar !== ui.Button.YES) return;

  var movimientos = compactarOrdenes(ss, sheet);
  if (movimientos.length === 0) {
    ui.alert('No había huecos que cerrar — ya estaba todo organizado.');
  } else {
    ui.alert('Listo, se renumeraron ' + movimientos.length + ' orden(es):\n\n' + movimientos.join('\n'));
  }
}

// Ordena las órdenes por OP (sin mezclar sus filas) y las renumera 1..k.
// Mueve también Tipo de Material / Observaciones de las hojas 01-14.
// Devuelve la lista de cambios de número.
function compactarOrdenes(ss, sheet) {
  var ordenes = leerOrdenes(sheet);
  if (ordenes.length === 0) return [];

  var ordenadas = ordenes.slice().sort(function (a, b) { return a.op - b.op; });
  var movimientos = [];
  var mismoOrden = true;
  for (var i = 0; i < ordenadas.length; i++) {
    if (ordenadas[i] !== ordenes[i]) mismoOrden = false;
    if (ordenadas[i].op !== i + 1) movimientos.push('OP ' + ordenadas[i].op + '  →  OP ' + (i + 1));
  }
  if (movimientos.length === 0 && mismoOrden) return [];

  // Guardar Tipo de Material / Observaciones antes de mover nada.
  var extras = {};
  ordenadas.forEach(function (o) { extras[o.op] = leerExtrasDeHoja(ss, o.op); });

  // Reescribir las columnas de entrada, ya ordenadas y renumeradas.
  var primeraFila = ordenes[0].fila;
  var totalFilas = 0;
  ordenadas.forEach(function (o) { totalFilas += o.n; });
  for (var c = 0; c < COLUMNAS_ENTRADA.length; c++) {
    var col = COLUMNAS_ENTRADA[c];
    var columna = [];
    ordenadas.forEach(function (o, idx) {
      o.filas.forEach(function (f) { columna.push([col === COL_OP ? idx + 1 : f[col - 1]]); });
    });
    sheet.getRange(primeraFila, col, totalFilas, 1).setValues(columna);
  }
  // Si había filas sueltas (sin OP) entre las órdenes, lo que sobra al final se limpia.
  var ultima = ordenes[ordenes.length - 1].fila + ordenes[ordenes.length - 1].n - 1;
  var sobran = ultima - (primeraFila + totalFilas - 1);
  if (sobran > 0) {
    for (var c2 = 0; c2 < COLUMNAS_ENTRADA.length; c2++) {
      sheet.getRange(primeraFila + totalFilas, COLUMNAS_ENTRADA[c2], sobran, 1).clearContent();
    }
    sheet.getRange(primeraFila + totalFilas, COL_COMPONENTE, sobran, 1).clearDataValidations();
    limpiarFormatoReserva(sheet);
  }

  // Listas de Componente según el SKU de cada orden.
  var fila = primeraFila;
  ordenadas.forEach(function (o) {
    var rango = sheet.getRange(fila, COL_COMPONENTE, o.n, 1);
    rango.clearDataValidations();
    var sku = o.filas[0][COL_SKU - 1];
    if (sku) aplicarValidacionComponente(rango, opcionesDeComponente(sku.toString()));
    fila += o.n;
  });

  // Hojas 01-14: cada número nuevo recibe lo de su orden; las que quedan sin orden se limpian.
  for (var nuevo = 1; nuevo <= HOJAS_DE_ORDEN; nuevo++) {
    var origen = ordenadas[nuevo - 1];
    if (origen && origen.op === nuevo) continue; // no cambió
    escribirExtrasEnHoja(ss, nuevo, origen ? extras[origen.op] : null);
  }

  return movimientos;
}

// ---------- Historial: archivar una orden antes de borrarla ----------

function obtenerOCrearHistorial() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName('Historial');
  if (!hoja) {
    hoja = ss.insertSheet('Historial');
    var encabezados = ['Fecha archivado', 'Motivo', 'OP', 'Máquina', 'Sku', 'Producto',
                        'Componente', 'Material', 'Cantidad', 'Cantidad Final', 'Color',
                        'Sacos', 'Estado'];
    var rango = hoja.getRange(1, 1, 1, encabezados.length);
    rango.setValues([encabezados]);
    rango.setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#045C82')
      .setHorizontalAlignment('center').setWrap(true);
    hoja.setFrozenRows(1);
    hoja.setColumnWidth(1, 130);
    hoja.setColumnWidth(6, 180);
  }
  return hoja;
}

// Convierte las filas con datos de una orden (ya leídas) en filas de Historial.
function filasParaHistorial(filas, numeroOP, motivo, fecha) {
  var maquina = filas[0][COL_MAQ - 1];
  var estado = ''; // la columna Estado ya no existe en el Maestro
  var salida = [];
  for (var i = 0; i < filas.length; i++) {
    var f = filas[i];
    var cantidad = f[COL_CANTIDAD - 1];
    var color = f[COL_COLOR - 1];
    var sku = f[COL_SKU - 1];
    if (cantidad === '' && color === '' && (sku === '' || esPlaceholderSku(sku))) continue;
    salida.push([
      fecha,
      motivo,
      numeroOP,
      maquina,
      sku,
      f[COL_PRODUCTO - 1],
      f[COL_COMPONENTE - 1],
      f[COL_MATERIAL - 1],
      cantidad,
      f[COL_CANTIDAD_FINAL - 1],
      color,
      f[COL_SACOS - 1],
      estado
    ]);
  }
  return salida;
}

function guardarEnHistorial(filasHistorial) {
  if (filasHistorial.length === 0) return;
  var historial = obtenerOCrearHistorial();
  historial.getRange(historial.getLastRow() + 1, 1, filasHistorial.length, filasHistorial[0].length)
    .setValues(filasHistorial);
}

// ---------- Guardar una copia con fecha en Drive ----------

function guardarCopiaEnDrive() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var mesesEs = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio',
                 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
  var hoy = new Date();
  var mesActual = hoy.getMonth() + 1; // 1-12
  var anioActual = hoy.getFullYear();

  var respDia = ui.prompt(
    'Guardar copia en Drive',
    '¿De qué día es esta orden? Escribe el número del día (hoy sería ' + hoy.getDate() + '):',
    ui.ButtonSet.OK_CANCEL
  );
  if (respDia.getSelectedButton() !== ui.Button.OK) return;
  var diaTexto = respDia.getResponseText().trim();
  var dia = parseInt(diaTexto, 10);
  if (isNaN(dia) || diaTexto !== dia.toString() || dia < 1 || dia > 31) {
    ui.alert('Escribe solo el número del día (ej: 19).');
    return;
  }

  var respMes = ui.prompt(
    'Guardar copia en Drive',
    '¿De qué mes es esta orden? Escribe el número del mes, del 1 al 12 (este mes sería ' + mesActual + '):',
    ui.ButtonSet.OK_CANCEL
  );
  if (respMes.getSelectedButton() !== ui.Button.OK) return;
  var mesTexto = respMes.getResponseText().trim();
  var mesNumero = parseInt(mesTexto, 10);
  if (isNaN(mesNumero) || mesTexto !== mesNumero.toString() || mesNumero < 1 || mesNumero > 12) {
    ui.alert('Escribe un número de mes válido, del 1 al 12.');
    return;
  }

  // Si el mes elegido "ya pasó" hace más de medio año, es una orden
  // adelantada del año que viene (ej: enero, estando en diciembre).
  var anio = (mesActual - mesNumero > 6) ? anioActual + 1 : anioActual;

  var diasDelMes = new Date(anio, mesNumero, 0).getDate();
  if (dia > diasDelMes) {
    ui.alert(mesesEs[mesNumero - 1] + ' de ' + anio + ' solo tiene ' + diasDelMes + ' días. Revisa la fecha.');
    return;
  }

  var nombre = 'ORDEN DE PRODUCCION ' + dia + ' DE ' + mesesEs[mesNumero - 1].toUpperCase() + ' DE ' + anio;
  var donde;
  var copia;
  try {
    var carpeta = DriveApp.getFolderById(CARPETA_COPIAS_ID);
    copia = DriveApp.getFileById(ss.getId()).makeCopy(nombre, carpeta);
    donde = 'en la carpeta "' + carpeta.getName() + '" de tu Drive';
  } catch (err) {
    // Si la carpeta no se encuentra (o no hay permiso), se guarda junto al original.
    copia = DriveApp.getFileById(ss.copy(nombre).getId());
    donde = 'en la misma carpeta que este archivo (no pude usar la carpeta de Seguimiento: ' + err.message + ')';
  }

  ui.alert(
    '✅ Copia guardada en Drive',
    'Se creó la copia con el nombre:\n\n"' + nombre + '"\n\n' +
    'Quedó guardada ' + donde + '. ' +
    'La puedes abrir, compartir o descargar como Excel cuando quieras.\n\n' +
    'Enlace: ' + copia.getUrl(),
    ui.ButtonSet.OK
  );
}

// ---------- Vaciar todo ----------

function vaciarMaestro() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(HOJA_MAESTRO);
  if (!sheet) { ui.alert('No encontré la hoja "' + HOJA_MAESTRO + '".'); return; }

  var confirmacion = ui.alert(
    '⚠️ Vaciar el Maestro',
    'Esto va a eliminar TODAS las órdenes del Maestro (quedan solo los encabezados), ' +
    'y va a limpiar el campo "Tipo de Material" y la caja de "Observaciones" en las 14 hojas de orden (01-14).\n\n' +
    'Las órdenes se guardan antes en el Historial.\n\n' +
    'Esta acción no se puede deshacer. ¿Seguro que quieres continuar?',
    ui.ButtonSet.YES_NO
  );
  if (confirmacion !== ui.Button.YES) return;

  var fecha = new Date();
  var paraHistorial = [];
  leerOrdenes(sheet).forEach(function (o) {
    paraHistorial = paraHistorial.concat(filasParaHistorial(o.filas, o.op, 'Vaciado masivo', fecha));
  });
  guardarEnHistorial(paraHistorial);

  var ultima = ultimaFilaConDatos(sheet);
  if (ultima >= FILA_INICIO) eliminarFilas(sheet, FILA_INICIO, ultima - FILA_INICIO + 1);

  limpiarTipoDeMaterial(ss);
  limpiarObservaciones(ss);

  ss.toast('Las órdenes quedaron guardadas en el Historial y el Maestro quedó vacío. ' +
           'El "Tipo de Material" y "Observaciones" de las 14 hojas quedaron en blanco.', '🧹 Listo', 8);
}

// ---------- Hojas 01-14: Tipo de Material y Observaciones ----------

// "Tipo de Material" (texto fijo, no fórmula): celda G21, justo debajo del
// encabezado "TIPO DE MATERIAL". "Observaciones": celda E34 (fusionada
// E34:H34). Ajusta aquí si alguna vez cambian de lugar.
var CELDA_TIPO_MATERIAL = 'G21';
var CELDA_OBSERVACIONES = 'E34';

// Devuelve la hoja 01-14 de una OP, o null si esa OP no tiene hoja propia.
function hojaDeOrden(ss, numeroOP) {
  if (numeroOP < 1 || numeroOP > HOJAS_DE_ORDEN) return null;
  var nombre = (numeroOP < 10 ? '0' : '') + numeroOP;
  return ss.getSheetByName(nombre);
}

function limpiarTipoDeMaterialDeHoja(ss, numeroOP) {
  var hoja = hojaDeOrden(ss, numeroOP);
  if (hoja) hoja.getRange(CELDA_TIPO_MATERIAL).clearContent();
}

function limpiarObservacionesDeHoja(ss, numeroOP) {
  var hoja = hojaDeOrden(ss, numeroOP);
  if (hoja) hoja.getRange(CELDA_OBSERVACIONES).clearContent();
}

function limpiarTipoDeMaterial(ss) {
  for (var i = 1; i <= HOJAS_DE_ORDEN; i++) limpiarTipoDeMaterialDeHoja(ss, i);
}

function limpiarObservaciones(ss) {
  for (var i = 1; i <= HOJAS_DE_ORDEN; i++) limpiarObservacionesDeHoja(ss, i);
}

// Lee el Tipo de Material y las Observaciones de la hoja de una OP (o null).
function leerExtrasDeHoja(ss, numeroOP) {
  var hoja = hojaDeOrden(ss, numeroOP);
  if (!hoja) return null;
  return {
    tipoMaterial: hoja.getRange(CELDA_TIPO_MATERIAL).getValue(),
    observaciones: hoja.getRange(CELDA_OBSERVACIONES).getValue()
  };
}

// Escribe (o borra, si extras es null) el Tipo de Material y las
// Observaciones en la hoja de una OP.
function escribirExtrasEnHoja(ss, numeroOP, extras) {
  var hoja = hojaDeOrden(ss, numeroOP);
  if (!hoja) return;
  var tipo = hoja.getRange(CELDA_TIPO_MATERIAL);
  var obs = hoja.getRange(CELDA_OBSERVACIONES);
  if (extras && extras.tipoMaterial !== '') tipo.setValue(extras.tipoMaterial); else tipo.clearContent();
  if (extras && extras.observaciones !== '') obs.setValue(extras.observaciones); else obs.clearContent();
}

// ---------- Borrar todo el Historial ----------

function borrarHistorial() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName('Historial');
  if (!hoja) {
    ui.alert('Todavía no existe la pestaña "Historial" (no se ha borrado ninguna orden con historial todavía, así que no hay nada que borrar).');
    return;
  }

  var confirmacion = ui.alert(
    '⚠️ Borrar TODO el Historial',
    'Esto borra TODO el registro de órdenes que ya se archivaron — el rastro completo de lo que se ha producido antes.\n\n' +
    'Esta acción es independiente de "Vaciar el Maestro" (una no afecta a la otra) y NO se puede deshacer. ' +
    '¿Seguro que quieres continuar?',
    ui.ButtonSet.YES_NO
  );
  if (confirmacion !== ui.Button.YES) return;

  var ultimaFila = hoja.getLastRow();
  if (ultimaFila > 1) {
    hoja.getRange(2, 1, ultimaFila - 1, hoja.getLastColumn()).clearContent();
  }
  ss.toast('El Historial quedó vacío.', '🗑️ Listo', 5);
}

// Vuelve a poner las listas desplegables (y la del Componente según su SKU)
// en las filas indicadas que tengan número de OP.
function restaurarListasEnFilas(sheet, filaIni, filaFin) {
  var n = filaFin - filaIni + 1;
  var datos = sheet.getRange(filaIni, 1, n, COL_COMPONENTE).getValues();
  var cache = {};
  for (var i = 0; i < n; i++) {
    if (datos[i][COL_OP - 1] === '') continue;
    var fila = filaIni + i;
    aplicarListasFilas(sheet, fila, 1);
    var sku = datos[i][COL_SKU - 1];
    if (sku !== '' && !esPlaceholderSku(sku)) {
      var clave = sku.toString();
      if (!cache[clave]) cache[clave] = opcionesDeComponente(clave);
      aplicarValidacionComponente(sheet.getRange(fila, COL_COMPONENTE), cache[clave]);
    }
  }
}

// ¿Alguna celda editada, de una columna con lista, se quedó sin su lista?
function faltaAlgunaLista(rango) {
  var conLista = {};
  LISTAS_FILA.forEach(function (l) { conLista[l.col] = true; });
  conLista[COL_CANTIDAD] = true;
  var reglas = rango.getDataValidations();
  var col0 = rango.getColumn();
  for (var i = 0; i < reglas.length; i++) {
    for (var j = 0; j < reglas[i].length; j++) {
      if (conLista[col0 + j] && !reglas[i][j]) return true;
    }
  }
  return false;
}

function repararListas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(HOJA_MAESTRO);
  if (!sheet) { SpreadsheetApp.getUi().alert('No encontré la hoja "' + HOJA_MAESTRO + '".'); return; }
  var ultima = ultimaFilaConDatos(sheet);
  if (ultima >= FILA_INICIO) restaurarListasEnFilas(sheet, FILA_INICIO, ultima);
  limpiarFormatoReserva(sheet);
  ss.toast('Las listas desplegables quedaron puestas en todas las filas con orden.', '🧩 Listo', 6);
}

// ---------- Automático: filtrar Componentes según el SKU (si editas directo en la hoja) ----------

function onEdit(e) {
  var rango = e.range;
  var sheet = rango.getSheet();
  if (sheet.getName() !== HOJA_MAESTRO) return;

  var colIni = rango.getColumn();
  var colFin = rango.getLastColumn();
  var filaIni = Math.max(rango.getRow(), FILA_INICIO);
  var filaFin = Math.min(rango.getLastRow(), sheet.getMaxRows());
  if (filaIni > filaFin || filaFin - filaIni > 50) return; // pegados enormes: no lo hacemos en vivo

  // Al pegar (Ctrl+V) Google borra la lista desplegable de esas celdas:
  // si a alguna celda editada le falta su lista, se vuelven a poner.
  if (faltaAlgunaLista(rango)) restaurarListasEnFilas(sheet, filaIni, filaFin);

  var tocaSku = colIni <= COL_SKU && COL_SKU <= colFin;
  var tocaComp = colIni <= COL_COMPONENTE && COL_COMPONENTE <= colFin;
  if (!tocaSku && !tocaComp) return;

  for (var fila = filaIni; fila <= filaFin; fila++) {
    if (tocaSku) {
      actualizarComponentesPorSkuEnHoja(sheet, fila);
    } else {
      validarComponenteContraSku(sheet, fila);
    }
  }
}

function actualizarComponentesPorSkuEnHoja(sheet, fila) {
  var sku = sheet.getRange(fila, COL_SKU).getValue();
  var celdaComponente = sheet.getRange(fila, COL_COMPONENTE);
  if (sku === '' || esPlaceholderSku(sku)) {
    // Se borró el SKU: el componente viejo ya no tiene sentido.
    celdaComponente.clearDataValidations();
    celdaComponente.clearContent();
    return;
  }

  var resultado = opcionesDeComponenteConOrigen(sku.toString());
  var opciones = resultado.opciones;
  if (opciones.length === 0) {
    celdaComponente.clearDataValidations();
    SpreadsheetApp.getActiveSpreadsheet().toast('No encontré componentes para el SKU "' + sku + '" (ni exactos ni genéricos).', '⚠️ Revisar SKU', 6);
    return;
  }

  aplicarValidacionComponente(celdaComponente, opciones);
  var etiqueta = resultado.generico ? ('📋 Genéricos de la familia "' + familiaDeSku(sku.toString()) + '"') : '📋 Componentes de este SKU';

  if (opciones.length === 1) {
    celdaComponente.setValue(opciones[0].codigo);
    var titulo = resultado.generico ? '✅ Componente genérico asignado' : '✅ Componente asignado';
    SpreadsheetApp.getActiveSpreadsheet().toast(opciones[0].codigo + ' — ' + opciones[0].nombre, titulo, 6);
  } else {
    var actual = celdaComponente.getValue().toString().trim().toUpperCase();
    var sigueValido = opciones.some(function (o) { return o.codigo.toUpperCase() === actual; });
    if (!sigueValido) celdaComponente.clearContent();
    var detalle = opciones.map(function (o) { return o.codigo + ' — ' + o.nombre; }).join('\n');
    SpreadsheetApp.getActiveSpreadsheet().toast(detalle, etiqueta + ' — elige uno', 15);
  }
}

function validarComponenteContraSku(sheet, fila) {
  var sku = sheet.getRange(fila, COL_SKU).getValue();
  var comp = sheet.getRange(fila, COL_COMPONENTE).getValue();
  if (!sku || !comp) return;

  var compStr = comp.toString().trim().toUpperCase();
  var validos = opcionesDeComponente(sku.toString()); // ya incluye el respaldo genérico
  var coincide = validos.some(function (o) { return o.codigo.toUpperCase() === compStr; });

  if (!coincide) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'El componente "' + comp + '" no parece ser del SKU "' + sku + '" (ni exacto ni genérico de su familia). Revísalo.',
      '⚠️ Componente no coincide', 8
    );
  }
}
