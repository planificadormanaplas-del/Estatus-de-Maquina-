/* Arreglo de "Leer inventario mas reciente (carpeta)".
   - Lee los reportes del sistema (.xls de 5 o 7 columnas) y el reporte semanal
     de materia prima (.xlsx con "Saldo Final").
   - Ya no borra el inventario: actualiza la existencia de lo que ya estaba,
     AGREGA los codigos nuevos y deja igual lo que no venga en el archivo.
   Este archivo reemplaza esas funciones de las partes anteriores.
   Debe quedar como el ULTIMO archivo de la lista. */

function leerInventarioDesdeCarpeta() {
  var ui = SpreadsheetApp.getUi();
  var folder = DriveApp.getFolderById(CARPETA_INVENTARIO_ID);
  var files = folder.getFiles();
  var masReciente = null;
  while (files.hasNext()) {
    var f = files.next();
    if (f.isTrashed()) continue;
    if (!masReciente || f.getLastUpdated().getTime() > masReciente.getLastUpdated().getTime()) masReciente = f;
  }
  if (!masReciente) {
    ui.alert('No encontré ningún archivo en la carpeta de inventario.');
    return;
  }

  var filas = leerFilasArchivoInventario(masReciente, ui);
  if (!filas) return;

  // Lo que trae el archivo: solo materia prima (01-) y colorantes (02-).
  // Los 04- son bolsas y empaques, y los 13- productos de limpieza y aceites.
  var delArchivo = {};
  filas.forEach(function (fila) {
    var cod = String(fila[0] || '').trim();
    if (!/^0[12]-\d{2}-\d{4}$/.test(cod)) return;
    var ex = (typeof fila[2] === 'number') ? fila[2] : parseNumeroInventario(fila[2]);
    if (ex === null || ex === undefined || ex === '') return;
    delArchivo[cod] = { nombre: String(fila[1] || '').replace(/\s+/g, ' ').trim(), ex: ex };
  });
  var nLeidos = Object.keys(delArchivo).length;
  if (nLeidos < 20) {
    ui.alert('Solo pude leer ' + nLeidos + ' producto(s) de "' + masReciente.getName() + '". ' +
      'Parece que el archivo no se leyó bien, así que NO cambié el inventario.');
    return;
  }

  // Lo que ya hay en Inv (se conserva lo que no venga en el archivo)
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja1 = ss.getSheetByName('Inv');
  var todos = {};
  if (hoja1.getLastRow() > 0) {
    hoja1.getRange(1, 1, hoja1.getLastRow(), 4).getValues().forEach(function (r) {
      var cod = String(r[0] || '').trim();
      if (!/^0\d-\d{2}-\d{4}$/.test(cod)) return;
      var ex = (typeof r[3] === 'number') ? r[3] : parseNumeroInventario(r[3]);
      todos[cod] = { nombre: String(r[1] || '').trim(), pista: String(r[2] || '').trim(), ex: ex === null ? 0 : ex };
    });
  }

  var agregados = [], actualizados = 0;
  Object.keys(delArchivo).forEach(function (cod) {
    var n = delArchivo[cod];
    if (todos[cod]) {
      if (todos[cod].ex !== n.ex) actualizados++;
      todos[cod].ex = n.ex;
      if (!todos[cod].nombre) todos[cod].nombre = n.nombre;
    } else {
      todos[cod] = { nombre: n.nombre, pista: '', ex: n.ex };
      agregados.push(cod + '  ' + n.nombre + '  (' + n.ex + ' kg)');
    }
  });
  var noVinieron = Object.keys(todos).filter(function (cod) { return !delArchivo[cod]; }).length;

  var materiaPrima = [], colorantes = [];
  Object.keys(todos).forEach(function (cod) {
    var x = todos[cod];
    if (cod.indexOf('01-') === 0 && !esResinaPreColoreada(x.nombre)) {
      materiaPrima.push([cod, x.nombre, x.pista || detectarCodigoInterno(x.nombre), x.ex]);
    } else {
      colorantes.push([cod, x.nombre, '', x.ex]);
    }
  });
  materiaPrima.sort(function (a, b) { return b[3] - a[3]; });
  colorantes.sort(function (a, b) { return b[3] - a[3]; });

  escribirHojaInv(hoja1, materiaPrima, colorantes);
  actualizarInventarioCore();
  actualizarPortada();

  ui.alert(
    'Inventario actualizado desde "' + masReciente.getName() + '"\n\n' +
    '- Leídos del archivo: ' + nLeidos + '\n' +
    '- Con existencia nueva: ' + actualizados + '\n' +
    '- Códigos NUEVOS agregados: ' + agregados.length + '\n' +
    (agregados.length ? '   ' + agregados.slice(0, 15).join('\n   ') + (agregados.length > 15 ? '\n   ...' : '') + '\n' : '') +
    '- Ya estaban y no venían en el archivo (se dejaron igual): ' + noVinieron + '\n\n' +
    'Materia Prima: ' + materiaPrima.length + ' productos · Colorantes: ' + colorantes.length + ' productos.\n' +
    'Ya se cruzó contra el consumo.'
  );
}

// Devuelve [Codigo, Nombre, Existencia] por fila, o null si no se pudo leer.
function leerFilasArchivoInventario(archivo, ui) {
  var tipo = archivo.getMimeType();
  var nombre = archivo.getName();
  if (tipo === MimeType.GOOGLE_SHEETS) {
    return filasDeTablaOAviso(SpreadsheetApp.openById(archivo.getId()), nombre, ui);
  }
  var esXlsx = tipo === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' || /\.xlsx$/i.test(nombre);
  if (esXlsx) {
    // Un .xlsx hay que convertirlo a Hoja de Google para poder leerlo.
    if (typeof Drive === 'undefined') {
      ui.alert('"' + nombre + '" es un archivo .xlsx. Para leerlo, activa una sola vez el servicio Drive API: ' +
        'en Apps Script, a la izquierda, "Servicios" (+) ▸ Drive API ▸ Agregar. ' +
        'O en Drive, abre el archivo con "Abrir con ▸ Hojas de cálculo de Google" y vuelve a usar este botón. ' +
        'No cambié nada en Inv.');
      return null;
    }
    var copia;
    try {
      copia = Drive.Files.copy({ name: 'tmp inventario', mimeType: MimeType.GOOGLE_SHEETS }, archivo.getId());
    } catch (e) {
      copia = Drive.Files.copy({ title: 'tmp inventario', mimeType: MimeType.GOOGLE_SHEETS }, archivo.getId(), { convert: true });
    }
    try {
      return filasDeTablaOAviso(SpreadsheetApp.openById(copia.id), nombre, ui);
    } finally {
      DriveApp.getFileById(copia.id).setTrashed(true);
    }
  }
  var strings = extraerStringsBinario(archivo.getBlob().getBytes());
  if (strings.length < 10) {
    ui.alert('No pude leer "' + nombre + '": no tiene la estructura esperada. No cambié nada en Inv.');
    return null;
  }
  return filasInventarioDesdeBinario(strings);
}

function filasDeTablaOAviso(libro, nombre, ui) {
  var filas = filasInventarioDesdeTabla(libro.getSheets()[0].getDataRange().getValues());
  if (!filas) ui.alert('No encontré la columna "Existencia" ni "Saldo Final" en "' + nombre + '". No cambié nada en Inv.');
  return filas;
}

// Escribe "Inv" en los mismos dos bloques de siempre.
function escribirHojaInv(hoja1, materiaPrima, colorantes) {
  var HEADER_FILL_MP = '#1DB37A', HEADER_FILL_COL = '#F2436B', HEADER_FILL = '#0B1F33';
  var FILL_MP = '#EAF8F1', FILL_COL = '#FDEBEF';
  hoja1.clear();
  var r = 1;
  hoja1.getRange(r, 1, 1, 4).merge().setValue('🧪 MATERIA PRIMA')
    .setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(12).setBackground(HEADER_FILL_MP);
  r++;
  hoja1.getRange(r, 1, 1, 4)
    .setValues([['Código', 'Nombre', 'Código interno (para cruzar con Seguimiento)', 'Existencia (kg)']])
    .setFontColor('#FFFFFF').setFontWeight('bold').setBackground(HEADER_FILL).setHorizontalAlignment('center');
  r++;
  var mpStart = r;
  if (materiaPrima.length) {
    hoja1.getRange(r, 1, materiaPrima.length, 3).setNumberFormat('@');
    hoja1.getRange(r, 1, materiaPrima.length, 4).setValues(materiaPrima).setBackground(FILL_MP);
    r += materiaPrima.length;
  }
  hoja1.getRange(r, 2).setValue('Subtotal Materia Prima').setFontWeight('bold');
  hoja1.getRange(r, 4).setFormula('=SUM(D' + mpStart + ':D' + (r - 1) + ')').setFontWeight('bold');
  r += 2;
  hoja1.getRange(r, 1, 1, 4).merge().setValue('🎨 COLORANTES / MASTERBATCH')
    .setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(12).setBackground(HEADER_FILL_COL);
  r++;
  hoja1.getRange(r, 1, 1, 4)
    .setValues([['Código', 'Nombre', '(no aplica para colorante)', 'Existencia (kg)']])
    .setFontColor('#FFFFFF').setFontWeight('bold').setBackground(HEADER_FILL).setHorizontalAlignment('center');
  r++;
  var colStart = r;
  if (colorantes.length) {
    hoja1.getRange(r, 1, colorantes.length, 1).setNumberFormat('@');
    hoja1.getRange(r, 1, colorantes.length, 4).setValues(colorantes).setBackground(FILL_COL);
    r += colorantes.length;
  }
  hoja1.getRange(r, 2).setValue('Subtotal Colorantes').setFontWeight('bold');
  hoja1.getRange(r, 4).setFormula('=SUM(D' + colStart + ':D' + (r - 1) + ')').setFontWeight('bold');
  hoja1.setColumnWidth(1, 110).setColumnWidth(2, 380).setColumnWidth(3, 220).setColumnWidth(4, 110);
  hoja1.getRange(3, 4, Math.max(r - 2, 1), 1).setNumberFormat('#,##0');
  hoja1.setFrozenRows(2);
}

function normalizarEncabezadoInv(s) {
  return String(s || '').toUpperCase().replace(/[^A-Z]/g, '');
}

function esNumeroInventario(s) {
  return /^-?[\d.]+(,\d+)?$/.test(String(s).trim());
}

// Reporte ya convertido a Google Sheets: ubica las columnas por su encabezado
// ("Saldo Final" en el reporte semanal, "Existencia" en el del sistema).
function filasInventarioDesdeTabla(datos) {
  for (var f = 0; f < Math.min(datos.length, 15); f++) {
    var enc = datos[f].map(normalizarEncabezadoInv);
    var cEx = enc.indexOf('SALDOFINAL');
    if (cEx === -1) cEx = enc.indexOf('EXISTENCIA');
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
  ['12-70-0000', 'GOMA 12-70'], ['BLANCO MI :11', 'PP BLANCO MI11'], ['BLANCO MI:11', 'PP BLANCO MI11'],
  ['6200 B', '6200B'], ['HI-2053', '2811'], ['5618', '2811'], ['20H92N', '20H92N']
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
