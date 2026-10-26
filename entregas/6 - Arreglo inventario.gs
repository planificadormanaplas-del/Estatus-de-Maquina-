/* Arreglo de "Traer inventario": el reporte del sistema ahora trae 7 columnas
   (Codigo, Nombre, Unidad, Almacen, Peso, Existencia, Total Peso) y el codigo
   viejo leia la columna equivocada. Este archivo reemplaza esa funcion.
   Debe quedar como el ULTIMO archivo de la lista. */

function leerInventarioDesdeCarpeta() {
  var ui = SpreadsheetApp.getUi();
  if (CARPETA_INVENTARIO_ID === 'PON_AQUI_EL_ID_DE_TU_CARPETA_DE_INVENTARIO') {
    ui.alert('Falta configurar la carpeta de inventario. Busca la línea que dice ' +
      'CARPETA_INVENTARIO_ID en el código y pon ahí el ID de tu carpeta (pídeselo a Claude si no sabes cuál es).');
    return;
  }

  var folder = DriveApp.getFolderById(CARPETA_INVENTARIO_ID);
  var files = folder.getFiles();
  var masReciente = null;
  while (files.hasNext()) {
    var f = files.next();
    if (!masReciente || f.getLastUpdated().getTime() > masReciente.getLastUpdated().getTime()) {
      masReciente = f;
    }
  }
  if (!masReciente) {
    ui.alert('No encontré ningún archivo en esa carpeta.');
    return;
  }

  // El archivo puede llegar de dos formas: ya convertido a Google Sheets por
  // Drive (lo normal), o todavía como el binario crudo sin convertir.
  // Detectamos cuál es y leemos cada uno a su manera.
  // Cada fila queda como [Codigo, Nombre, Existencia]. La columna Existencia
  // se busca por su encabezado, asi da igual cuantas columnas traiga el
  // reporte del sistema (antes eran 5; ahora trae tambien Peso y Total Peso).
  var filas;
  if (masReciente.getMimeType() === MimeType.GOOGLE_SHEETS) {
    var libroInv = SpreadsheetApp.openById(masReciente.getId());
    var hojaInv = libroInv.getSheets()[0];
    filas = filasInventarioDesdeTabla(hojaInv.getDataRange().getValues());
    if (!filas) {
      ui.alert('No encontré la columna "Existencia" en "' + masReciente.getName() + '". No cambié nada en Inv.');
      return;
    }
  } else {
    var strings = extraerStringsBinario(masReciente.getBlob().getBytes());
    if (strings.length < 10) {
      ui.alert('No pude leer el archivo "' + masReciente.getName() + '" — no tiene la estructura esperada. No cambié nada en Inv.');
      return;
    }
    filas = filasInventarioDesdeBinario(strings);
  }

  var materiaPrima = [], colorantes = [];
  var yaLeidas = {}; // filas idénticas (código + nombre + existencia) entran una sola vez
  filas.forEach(function (fila) {
    var cod = String(fila[0] || '').trim();
    var nombre = fila[1];
    if (cod === 'Total' || cod === '') return;
    var existencia = (typeof fila[2] === 'number') ? fila[2] : parseNumeroInventario(fila[2]);
    if (existencia === null || existencia === undefined) return;
    var firma = cod + '|' + String(nombre).trim() + '|' + existencia;
    if (yaLeidas[firma]) return;
    yaLeidas[firma] = true;
    // Solo materia prima (01-) y colorantes (02-). Los codigos 04- son bolsas y
    // empaques, y los 13- son productos de limpieza y aceites: no van en Inv.
    if (cod.indexOf('01-') === 0) {
      // Resina pre-coloreada (ej. "PP CARGADO AZUL REY", "POLIPROPILENO VINO 5072")
      // va con Colorantes aunque su código sea de Materia Prima — es resina que
      // ya viene con el color puesto, no un aditivo que se dosifica aparte.
      var esResinaCargada = esResinaPreColoreada(nombre);
      if (esResinaCargada) {
        colorantes.push([cod, nombre, '', existencia]);
      } else {
        materiaPrima.push([cod, nombre, detectarCodigoInterno(nombre), existencia]);
      }
    } else if (cod.indexOf('02-') === 0) {
      colorantes.push([cod, nombre, '', existencia]);
    }
  });
  materiaPrima.sort(function (a, b) { return b[3] - a[3]; });
  colorantes.sort(function (a, b) { return b[3] - a[3]; });

  if (materiaPrima.length === 0 && colorantes.length === 0) {
    ui.alert('Leí el archivo "' + masReciente.getName() + '" pero no encontré ningún código que ' +
      'empiece con "01-" o "02-" — revisa que sea el archivo correcto.');
    return;
  }

  // Seguro: si el archivo trae muy pocos productos, algo salio mal al leerlo.
  // En ese caso no se borra el inventario que ya estaba.
  if (materiaPrima.length + colorantes.length < 20) {
    ui.alert('Solo pude leer ' + (materiaPrima.length + colorantes.length) + ' producto(s) de "' +
      masReciente.getName() + '". Parece que el archivo no se leyó bien, así que NO cambié el inventario. ' +
      'Revisa que sea el reporte completo de existencias.');
    return;
  }

  // --- Escribir en "Inv", en los mismos dos bloques de siempre ---
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja1 = ss.getSheetByName('Inv');
  hoja1.clear();

  // Mismos colores del diseño nuevo (Inicio / Dashboard)
  var HEADER_FILL_MP = '#1DB37A', HEADER_FILL_COL = '#F2436B', HEADER_FILL = '#0B1F33';
  var FILL_MP = '#EAF8F1', FILL_COL = '#FDEBEF';

  var r = 1;
  hoja1.getRange(r, 1, 1, 4).merge().setValue('🧪 MATERIA PRIMA')
    .setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(12).setBackground(HEADER_FILL_MP);
  r++;
  hoja1.getRange(r, 1, 1, 4)
    .setValues([['Código', 'Nombre', 'Código interno (para cruzar con Seguimiento)', 'Existencia (kg)']])
    .setFontColor('#FFFFFF').setFontWeight('bold').setBackground(HEADER_FILL)
    .setHorizontalAlignment('center');
  r++;
  var mpStart = r;
  if (materiaPrima.length > 0) {
    // columnas A a C como texto forzado: evita que Sheets convierta solo
    // codigos como "0240" o "2811" a numero (y pierda el cero adelante).
    // Todo el bloque se escribe de una vez (antes era fila por fila).
    hoja1.getRange(r, 1, materiaPrima.length, 3).setNumberFormat('@');
    hoja1.getRange(r, 1, materiaPrima.length, 4).setValues(materiaPrima).setBackground(FILL_MP);
    r += materiaPrima.length;
  }
  var mpEnd = r - 1;
  hoja1.getRange(r, 2).setValue('Subtotal Materia Prima').setFontWeight('bold');
  hoja1.getRange(r, 4).setFormula('=SUM(D' + mpStart + ':D' + mpEnd + ')').setFontWeight('bold');
  r += 2;

  hoja1.getRange(r, 1, 1, 4).merge().setValue('🎨 COLORANTES / MASTERBATCH')
    .setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(12).setBackground(HEADER_FILL_COL);
  r++;
  hoja1.getRange(r, 1, 1, 4)
    .setValues([['Código', 'Nombre', '(no aplica para colorante)', 'Existencia (kg)']])
    .setFontColor('#FFFFFF').setFontWeight('bold').setBackground(HEADER_FILL)
    .setHorizontalAlignment('center');
  r++;
  var colStart = r;
  if (colorantes.length > 0) {
    hoja1.getRange(r, 1, colorantes.length, 1).setNumberFormat('@');
    hoja1.getRange(r, 1, colorantes.length, 4).setValues(colorantes).setBackground(FILL_COL);
    r += colorantes.length;
  }
  var colEnd = r - 1;
  hoja1.getRange(r, 2).setValue('Subtotal Colorantes').setFontWeight('bold');
  hoja1.getRange(r, 4).setFormula('=SUM(D' + colStart + ':D' + colEnd + ')').setFontWeight('bold');

  hoja1.setColumnWidth(1, 110).setColumnWidth(2, 380).setColumnWidth(3, 220).setColumnWidth(4, 110);
  // Existencias con separador de miles, igual en todas las filas
  hoja1.getRange(3, 4, Math.max(r - 2, 1), 1).setNumberFormat('#,##0');

  // cruzar automaticamente con el consumo, igual que el boton de siempre
  actualizarInventarioCore();
  actualizarPortada();

  ui.alert(
    'Inventario actualizado desde "' + masReciente.getName() + '"\n' +
    '(modificado: ' + masReciente.getLastUpdated().toLocaleString() + ')\n\n' +
    '- Materia Prima: ' + materiaPrima.length + ' productos\n' +
    '- Colorantes: ' + colorantes.length + ' productos\n\n' +
    'Ya se cruzó contra el consumo — revisa "Consumo por Material" y "Resumen Consumo".'
  );
}

function normalizarEncabezadoInv(s) {
  return String(s || '').toUpperCase().replace(/[^A-Z]/g, '');
}

function esNumeroInventario(s) {
  return /^-?[\d.]+(,\d+)?$/.test(String(s).trim());
}

// Reporte ya convertido a Google Sheets: ubica las columnas por su encabezado.
function filasInventarioDesdeTabla(datos) {
  for (var f = 0; f < Math.min(datos.length, 15); f++) {
    var enc = datos[f].map(normalizarEncabezadoInv);
    var cEx = enc.indexOf('EXISTENCIA');
    if (cEx === -1) continue;
    var cCod = 0, cNom = enc.indexOf('NOMBRE');
    for (var c = 0; c < enc.length; c++) {
      if (/^C.?DIGO$/.test(enc[c])) { cCod = c; break; }
    }
    if (cNom === -1) cNom = cCod + 1;
    return datos.slice(f + 1).map(function (r) { return [String(r[cCod] || ''), String(r[cNom] || ''), r[cEx]]; });
  }
  return null;
}

// Reporte en formato viejo (binario): las celdas vacias no dejan rastro en el
// archivo, asi que no se puede contar de 5 en 5. Cada producto empieza con su
// codigo y termina en el siguiente codigo o en una fila "Total"; adentro van
// nombre, unidad, almacen y los numeros. Si hay tres
// numeros son Peso, Existencia y Total Peso; si hay uno solo es la Existencia.
function filasInventarioDesdeBinario(strings) {
  var filas = [], actual = null;
  strings.forEach(function (s) {
    var t = String(s).trim();
    if (/^\d{2,3}-\d{2,4}-[\dA-Z\/\-]+$/i.test(t)) {
      if (actual) filas.push(actual);
      actual = { cod: t, resto: [] };
    } else if (/^TOTAL/i.test(t)) {
      if (actual) filas.push(actual);
      actual = null;
    } else if (actual) {
      actual.resto.push(t);
    }
  });
  if (actual) filas.push(actual);
  return filas.map(function (x) {
    var nums = x.resto.slice(1).filter(esNumeroInventario);
    var ex = nums.length >= 3 ? nums[1] : (nums.length ? nums[nums.length - 1] : null);
    return [x.cod, x.resto[0] || '', ex];
  });
}

// Nombres de la columna C de Inv ("Codigo interno"). Reemplaza la lista del
// archivo anterior: el 30R ahora se llama con su nombre completo 30R40CD y los
// materiales que no se usan en las ordenes tambien llevan un nombre corto.
ALIAS_MATERIAL_MP = [
  ['20H10NA', '20H10NA'], ['11H01A', '11H01A'], ['60C90ND', '60C90ND'],
  ['2213', '2213'], ['20400', '2213'], ['2111', '2213'],
  ['0240', '0240'], ['6200B', '6200B'], ['0506C', '0506C'],
  ['2811', '2811'], ['2908', '2908'], ['40H92C', '40H92C'],
  ['30R40CD', '30R40CD'], ['30R', '30R40CD'], ['20H01NA', '20H01NA'], ['02R01CA-1', '02R01CA-1'],
  ['35R80', '35R80'], ['35R81', '35R81'], ['35R82', '35R82'], ['35R83', '35R83'],
  ['05C06', '05C06'], ['478', '478'],
  ['18H86', '18H86'], ['45R60CD', '45R60CD'],
  ['GHB1MA40.328', 'CHEMITON AZUL AGUA'], ['GHB1MA30.171', 'CHEMITON VERDE'], ['CHEMITON COLOR AMARILLO', 'CHEMITON AMARILLO'],
  ['M-90', 'ACETAL M-90'], ['K-300', 'ACETAL K-300'], ['RUBBAPEX', 'CAUCHO 7008'],
  ['AG12AO', 'ABS AG12AO'], ['1500-1060', 'NYLON 1500'], ['5203HE', 'PS CRISTAL 5203HE'],
  ['649D', 'PS CRISTAL 649D'], ['303-80111', 'PC LEXAN 303'], ['141-7041', 'PC LEXAN 141'],
  ['12-70-0000', 'GOMA 12-70'], ['BLANCO MI :11', 'PP BLANCO MI11']
];
ALIAS_MATERIAL_MP.sort(function (a, b) { return b[0].length - a[0].length; });
DICCIONARIO_MATERIALES['30R40CD'] = { nombre: 'PP Copolímero 30R40CD', familia: 'PPC' };

// Inicio y Dashboard se calculan solos con formulas: "Actualizar todo" no
// tiene que redibujarlos (si alguna parte vieja todavia lo hace, esto lo evita).
function actualizarDashboard() {
  SpreadsheetApp.flush();
}

function actualizarPortada() {
  SpreadsheetApp.flush();
}
