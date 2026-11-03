/* Arreglo de "Leer inventario mas reciente (carpeta)".
   - Lee los DOS reportes de la carpeta: el semanal de los lunes (.xlsx con
     "Saldo Final") y el de la plataforma (.xls de 5 o 7 columnas). Si un
     codigo viene en los dos, gana el semanal de los lunes.
   - Ya no borra el inventario: actualiza la existencia de lo que ya estaba,
     AGREGA los codigos nuevos y deja igual lo que no venga en el archivo.
   Este archivo reemplaza esas funciones de las partes anteriores.
   Debe quedar como el ULTIMO archivo de la lista. */

function leerInventarioDesdeCarpeta() {
  var ui = SpreadsheetApp.getUi();
  var folder = DriveApp.getFolderById(CARPETA_INVENTARIO_ID);

  // Archivos de la carpeta, del mas nuevo al mas viejo
  var lista = [];
  var it = folder.getFiles();
  while (it.hasNext()) {
    var f = it.next();
    if (!f.isTrashed()) lista.push(f);
  }
  lista.sort(function (x, y) { return y.getLastUpdated().getTime() - x.getLastUpdated().getTime(); });
  if (!lista.length) {
    ui.alert('No encontré ningún archivo en la carpeta de inventario.');
    return;
  }

  // Se usan dos reportes: el SEMANAL de los lunes (columna "Saldo Final") y
  // el de la PLATAFORMA (el .xls de nombre raro, columna "Existencia").
  // De cada tipo se toma el mas reciente.
  var semanal = null, plataforma = null, avisos = [];
  for (var i = 0; i < lista.length && (!semanal || !plataforma); i++) {
    var leido = leerReporteInventario(lista[i]);
    if (leido.error) { avisos.push(leido.error); continue; }
    if (leido.tipo === 'semanal' && !semanal) semanal = leido;
    if (leido.tipo === 'plataforma' && !plataforma) plataforma = leido;
  }
  if (!semanal && !plataforma) {
    ui.alert('No pude leer ningún reporte de inventario de la carpeta. No cambié nada en Inv.' +
      (avisos.length ? '\n\n' + avisos.slice(0, 3).join('\n') : ''));
    return;
  }

  // Primero la plataforma y encima el semanal: si un codigo viene en los dos,
  // gana el semanal de los lunes. Lo que solo viene en uno tambien entra.
  var delArchivo = {}, deDonde = {};
  var etiqueta = function (rep) {
    return (rep.tipo === 'semanal' ? 'Semanal ' : 'Plataforma ') +
      Utilities.formatDate(rep.fecha, Session.getScriptTimeZone(), 'dd/MM');
  };
  [plataforma, semanal].forEach(function (rep) {
    if (!rep) return;
    var et = etiqueta(rep);
    Object.keys(rep.datos).forEach(function (cod) {
      delArchivo[cod] = rep.datos[cod];
      delArchivo[cod].fuente = et;
      deDonde[cod] = rep.tipo;
    });
  });

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja1 = ss.getSheetByName('Inv');
  var todos = {};
  if (hoja1.getLastRow() > 0) {
    var nCols = Math.min(Math.max(hoja1.getLastColumn(), 4), 8);
    hoja1.getRange(1, 1, hoja1.getLastRow(), nCols).getValues().forEach(function (r) {
      var cod = String(r[0] || '').trim();
      if (!/^0\d-\d{2}-\d{4}$/.test(cod)) return;
      var ex = (typeof r[3] === 'number') ? r[3] : parseNumeroInventario(r[3]);
      todos[cod] = { nombre: String(r[1] || '').trim(), pista: String(r[2] || '').trim(), ex: ex === null ? 0 : ex,
        si: '', com: '', con: '', fuente: r[7] ? String(r[7]) + ' (anterior)' : 'Anterior' };
    });
  }

  var agregados = [], actualizados = 0;
  Object.keys(delArchivo).forEach(function (cod) {
    var n = delArchivo[cod];
    if (todos[cod]) {
      if (todos[cod].ex !== n.ex) actualizados++;
      todos[cod].ex = n.ex;
      todos[cod].si = n.si; todos[cod].com = n.com; todos[cod].con = n.con; todos[cod].fuente = n.fuente;
      if (!todos[cod].nombre) todos[cod].nombre = n.nombre;
    } else {
      todos[cod] = { nombre: n.nombre, pista: '', ex: n.ex, si: n.si, com: n.com, con: n.con, fuente: n.fuente };
      agregados.push(cod + '  ' + n.nombre + '  (' + n.ex + ' kg)');
    }
  });
  var noVinieron = Object.keys(todos).filter(function (cod) { return !delArchivo[cod]; }).length;
  var soloPlataforma = Object.keys(deDonde).filter(function (cod) { return deDonde[cod] === 'plataforma'; }).length;

  var materiaPrima = [], colorantes = [];
  Object.keys(todos).forEach(function (cod) {
    var x = todos[cod];
    if (cod.indexOf('01-') === 0 && !esResinaPreColoreada(x.nombre)) {
      materiaPrima.push([cod, x.nombre, x.pista || detectarCodigoInterno(x.nombre), x.ex, x.si, x.com, x.con, x.fuente]);
    } else {
      colorantes.push([cod, x.nombre, '', x.ex, x.si, x.com, x.con, x.fuente]);
    }
  });
  materiaPrima.sort(function (a, b) { return b[3] - a[3]; });
  colorantes.sort(function (a, b) { return b[3] - a[3]; });

  var titulo = semanal ? 'Semana al ' + Utilities.formatDate(semanal.fecha, Session.getScriptTimeZone(), 'dd/MM/yyyy') : '';
  escribirHojaInv(hoja1, materiaPrima, colorantes, titulo);
  actualizarInventarioCore();
  actualizarPortada();

  var fecha = function (rep) { return Utilities.formatDate(rep.fecha, Session.getScriptTimeZone(), 'dd/MM/yyyy'); };
  ui.alert(
    'Inventario actualizado.\n\n' +
    '📅 Semanal (lunes): ' + (semanal ? '"' + semanal.archivo.getName() + '" (' + fecha(semanal) + ') — ' + Object.keys(semanal.datos).length + ' productos' : 'no encontré ninguno') + '\n' +
    '💻 Plataforma: ' + (plataforma ? '"' + plataforma.archivo.getName() + '" (' + fecha(plataforma) + ') — ' + Object.keys(plataforma.datos).length + ' productos' : 'no encontré ninguno') + '\n' +
    (semanal && plataforma ? '   Si un código viene en los dos, se usó el semanal. ' + soloPlataforma + ' código(s) solo venían en el de la plataforma.\n' : '') +
    '\n- Con existencia nueva: ' + actualizados + '\n' +
    '- Códigos NUEVOS agregados: ' + agregados.length + '\n' +
    (agregados.length ? '   ' + agregados.slice(0, 15).join('\n   ') + (agregados.length > 15 ? '\n   ...' : '') + '\n' : '') +
    '- Ya estaban y no venían en ningún reporte (se dejaron igual): ' + noVinieron + '\n\n' +
    'Materia Prima: ' + materiaPrima.length + ' productos · Colorantes: ' + colorantes.length + ' productos.\n' +
    'Ya se cruzó contra el consumo.' +
    (!semanal ? '\n\n⚠️ OJO: NO se usó el reporte semanal de los lunes, solo el de la plataforma.' +
      (avisos.length ? '\nMotivo:\n- ' + avisos.slice(0, 3).join('\n- ') : '\nNo hay ningún reporte semanal (.xlsx con "Saldo Final") en la carpeta.') : '')
  );
}

// Lee un archivo de la carpeta. Devuelve { tipo: 'semanal' | 'plataforma',
// archivo, datos: { codigo: { nombre, ex } } } o { error: 'texto' }.
function leerReporteInventario(archivo) {
  var nombre = archivo.getName();
  var tipoMime = archivo.getMimeType();
  var tabla = null, filas = null, tipo = 'plataforma';
  var esXlsx = tipoMime === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' || /\.xlsx$/i.test(nombre);
  try {
    if (tipoMime === MimeType.GOOGLE_SHEETS) {
      tabla = SpreadsheetApp.openById(archivo.getId()).getSheets()[0].getDataRange().getValues();
    } else if (esXlsx) {
      // El .xlsx se abre aqui mismo (es un zip con la tabla en texto). Solo si
      // eso falla se intenta convertirlo con el servicio Drive API.
      try { tabla = tablaDesdeXlsx(archivo.getBlob()); } catch (e) { tabla = null; }
      if (!tabla || !tabla.length) {
        if (typeof Drive === 'undefined') return { error: 'No pude abrir el .xlsx "' + nombre + '".' };
        var copia;
        try {
          copia = Drive.Files.copy({ name: 'tmp inventario', mimeType: MimeType.GOOGLE_SHEETS }, archivo.getId());
        } catch (e) {
          copia = Drive.Files.copy({ title: 'tmp inventario', mimeType: MimeType.GOOGLE_SHEETS }, archivo.getId(), { convert: true });
        }
        try {
          tabla = SpreadsheetApp.openById(copia.id).getSheets()[0].getDataRange().getValues();
        } finally {
          DriveApp.getFileById(copia.id).setTrashed(true);
        }
      }
    } else {
      var strings = extraerStringsBinario(archivo.getBlob().getBytes());
      if (strings.length < 10) return { error: '"' + nombre + '" no tiene la estructura de un reporte de inventario.' };
      filas = filasInventarioDesdeBinario(strings);
    }
  } catch (e) {
    return { error: 'No pude abrir "' + nombre + '": ' + e.message };
  }
  if (tabla) {
    var leidas = filasInventarioDesdeTabla(tabla);
    if (!leidas) return { error: '"' + nombre + '" no tiene columna "Saldo Final" ni "Existencia".' };
    filas = leidas.filas;
    tipo = leidas.columna === 'SALDOFINAL' ? 'semanal' : 'plataforma';
    var fechaReporte = leidas.fecha;
  }

  var datos = {};
  filas.forEach(function (fila) {
    var cod = String(fila[0] || '').trim();
    if (!/^0[12]-\d{2}-\d{4}$/.test(cod)) return;
    var ex = (typeof fila[2] === 'number') ? fila[2] : parseNumeroInventario(fila[2]);
    if (ex === null || ex === undefined || ex === '') return;
    var num = function (v) { return (typeof v === 'number') ? v : (v === '' || v === undefined ? '' : parseNumeroInventario(v)); };
    datos[cod] = { nombre: String(fila[1] || '').replace(/\s+/g, ' ').trim(), ex: ex,
      si: num(fila[3]), com: num(fila[4]), con: num(fila[5]) };
  });
  if (Object.keys(datos).length < 20) {
    return { error: '"' + nombre + '": solo pude leer ' + Object.keys(datos).length + ' producto(s), no lo usé.' };
  }
  return { tipo: tipo, archivo: archivo, datos: datos, fecha: fechaReporte || archivo.getLastUpdated() };
}

// Escribe "Inv" en dos bloques, con las columnas del reporte de los lunes.
// A:D son las que usa el resto del archivo (Codigo, Nombre, Codigo interno,
// Existencia); E:H son informativas (Saldo inicial, Compras, Consumo, Fuente).
function escribirHojaInv(hoja1, materiaPrima, colorantes, titulo) {
  var HEADER_FILL_MP = '#1DB37A', HEADER_FILL_COL = '#F2436B', HEADER_FILL = '#0B1F33';
  var FILL_MP = '#EAF8F1', FILL_COL = '#FDEBEF', NC = 8;
  if (hoja1.getMaxColumns() < NC) hoja1.insertColumnsAfter(hoja1.getMaxColumns(), NC - hoja1.getMaxColumns());
  var filasNecesarias = materiaPrima.length + colorantes.length + 12;
  if (hoja1.getMaxRows() < filasNecesarias) hoja1.insertRowsAfter(hoja1.getMaxRows(), filasNecesarias - hoja1.getMaxRows());
  if (hoja1.getFilter()) hoja1.getFilter().remove();
  hoja1.clear();
  var encabezado = function (fila, colC) {
    hoja1.getRange(fila, 1, 1, NC)
      .setValues([['Código', 'Nombre', colC, 'Existencia (kg) = Saldo final', 'Saldo inicial', 'Compras', 'Consumo', 'Fuente']])
      .setFontColor('#FFFFFF').setFontWeight('bold').setBackground(HEADER_FILL)
      .setHorizontalAlignment('center').setWrap(true);
  };
  var bloque = function (fila, datos, fill) {
    if (!datos.length) return fila;
    hoja1.getRange(fila, 1, datos.length, 3).setNumberFormat('@');
    hoja1.getRange(fila, 1, datos.length, NC).setValues(datos).setBackground(fill);
    return fila + datos.length;
  };
  var subtotal = function (fila, desde, texto) {
    hoja1.getRange(fila, 2).setValue(texto).setFontWeight('bold');
    ['D', 'E', 'F', 'G'].forEach(function (c, k) {
      hoja1.getRange(fila, 4 + k).setFormula('=SUM(' + c + desde + ':' + c + (fila - 1) + ')').setFontWeight('bold');
    });
  };

  var r = 1;
  hoja1.getRange(r, 1, 1, NC).merge().setValue('🧪 MATERIA PRIMA' + (titulo ? '   ·   ' + titulo : ''))
    .setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(12).setBackground(HEADER_FILL_MP);
  r++;
  encabezado(r, 'Código interno (para cruzar con Seguimiento)');
  r++;
  var mpStart = r;
  r = bloque(r, materiaPrima, FILL_MP);
  subtotal(r, mpStart, 'Subtotal Materia Prima');
  r += 2;
  hoja1.getRange(r, 1, 1, NC).merge().setValue('🎨 COLORANTES / MASTERBATCH' + (titulo ? '   ·   ' + titulo : ''))
    .setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(12).setBackground(HEADER_FILL_COL);
  r++;
  encabezado(r, '(no aplica para colorante)');
  r++;
  var colStart = r;
  r = bloque(r, colorantes, FILL_COL);
  subtotal(r, colStart, 'Subtotal Colorantes');

  hoja1.setColumnWidth(1, 110).setColumnWidth(2, 380).setColumnWidth(3, 200).setColumnWidth(4, 120)
    .setColumnWidth(5, 95).setColumnWidth(6, 85).setColumnWidth(7, 85).setColumnWidth(8, 130);
  hoja1.getRange(3, 4, Math.max(r - 2, 1), 4).setNumberFormat('#,##0');
  hoja1.getRange(3, 8, Math.max(r - 2, 1), 1).setFontSize(9).setFontColor('#5C6B7A').setHorizontalAlignment('center');
  // Lo que vino solo del reporte de la plataforma se marca en azul claro
  var reglas = hoja1.getConditionalFormatRules().filter(function (x) {
    var c = x.getBooleanCondition();
    return !(c && String(c.getCriteriaValues()[0]).indexOf('Plataforma') !== -1);
  });
  reglas.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=REGEXMATCH($H3&"","^Plataforma")')
    .setBackground('#E3EFFD').setRanges([hoja1.getRange(3, 1, Math.max(r - 2, 1), NC)]).build());
  hoja1.setConditionalFormatRules(reglas);
  hoja1.setFrozenRows(2);
}

// ---- Lector de .xlsx sin servicios extra ----
// Un .xlsx es un zip; adentro, la hoja esta en xl/worksheets/sheetN.xml y los
// textos repetidos en xl/sharedStrings.xml. Devuelve la primera hoja como
// tabla (arreglo de filas), igual que getValues().
function tablaDesdeXlsx(blob) {
  var archivos = Utilities.unzip(blob.setContentType('application/zip'));
  var porNombre = {};
  archivos.forEach(function (f) { porNombre[f.getName()] = f.getDataAsString('UTF-8'); });
  var textos = porNombre['xl/sharedStrings.xml'] ? leerTextosCompartidosXlsx(porNombre['xl/sharedStrings.xml']) : [];
  return leerHojaXlsx(porNombre[rutaPrimeraHojaXlsx(porNombre)], textos);
}

function rutaPrimeraHojaXlsx(porNombre) {
  var libro = porNombre['xl/workbook.xml'] || '', rels = porNombre['xl/_rels/workbook.xml.rels'] || '';
  var m = /<sheet\b[^>]*\br:id="([^"]+)"/.exec(libro);
  if (m) {
    var re = /<Relationship\b([^>]*)\/?>/g, x;
    while ((x = re.exec(rels))) {
      var id = /\bId="([^"]+)"/.exec(x[1]), destino = /\bTarget="([^"]+)"/.exec(x[1]);
      if (id && destino && id[1] === m[1]) {
        var ruta = destino[1].replace(/^\//, '');
        return ruta.indexOf('xl/') === 0 ? ruta : 'xl/' + ruta;
      }
    }
  }
  return Object.keys(porNombre).filter(function (n) { return /^xl\/worksheets\/[^\/]+\.xml$/.test(n); }).sort()[0];
}

function textoXml(s) {
  return String(s).replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, function (m, n) { return String.fromCharCode(Number(n)); }).replace(/&amp;/g, '&');
}

function juntarTextosXml(xml) {
  var partes = [], re = /<t\b[^>]*>([\s\S]*?)<\/t>/g, m;
  xml = xml.replace(/<rPh\b[\s\S]*?<\/rPh>/g, '');
  while ((m = re.exec(xml))) partes.push(textoXml(m[1]));
  return partes.join('');
}

function leerTextosCompartidosXlsx(xml) {
  var lista = [], re = /<si\b[^>]*>([\s\S]*?)<\/si>/g, m;
  while ((m = re.exec(xml))) lista.push(juntarTextosXml(m[1]));
  return lista;
}

function leerHojaXlsx(xml, textos) {
  var tabla = [], reFila = /<row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g, f;
  var columna = function (letras) {
    var n = 0;
    for (var i = 0; i < letras.length; i++) n = n * 26 + (letras.charCodeAt(i) - 64);
    return n - 1;
  };
  while ((f = reFila.exec(xml || ''))) {
    var numFila = /\br="(\d+)"/.exec(f[1]);
    var fila = numFila ? Number(numFila[1]) - 1 : tabla.length;
    var valores = [], reCelda = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g, c, siguiente = 0;
    while ((c = reCelda.exec(f[2] || ''))) {
      var ref = /\br="([A-Z]+)\d+"/.exec(c[1]);
      var col = ref ? columna(ref[1]) : siguiente;
      siguiente = col + 1;
      var tipo = (/\bt="([^"]+)"/.exec(c[1]) || [])[1] || 'n';
      var v = /<v>([\s\S]*?)<\/v>/.exec(c[2] || '');
      var valor = '';
      if (tipo === 's') valor = v ? (textos[Number(v[1])] || '') : '';
      else if (tipo === 'inlineStr') valor = juntarTextosXml(c[2] || '');
      else if (tipo === 'str' || tipo === 'e') valor = v ? textoXml(v[1]) : '';
      else if (tipo === 'b') valor = v ? v[1] === '1' : '';
      else valor = v ? Number(v[1]) : '';
      valores[col] = valor;
    }
    tabla[fila] = valores;
  }
  var ancho = 0;
  tabla.forEach(function (r) { if (r && r.length > ancho) ancho = r.length; });
  for (var i = 0; i < tabla.length; i++) {
    var r = tabla[i] || [];
    for (var k = 0; k < ancho; k++) if (r[k] === undefined) r[k] = '';
    tabla[i] = r;
  }
  return tabla;
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
    var columna = enc.indexOf('SALDOFINAL') !== -1 ? 'SALDOFINAL' : 'EXISTENCIA';
    var cEx = enc.indexOf(columna);
    if (cEx === -1) continue;
    var cCod = 0, cNom = enc.indexOf('NOMBRE');
    var cSI = enc.indexOf('SALDOINICIAL'), cCom = enc.indexOf('COMPRAS'), cCon = enc.indexOf('CONSUMO');
    for (var c = 0; c < enc.length; c++) {
      if (/^C.?DIGO$/.test(enc[c])) { cCod = c; break; }
    }
    if (cNom === -1) cNom = cCod + 1;
    // Fecha del reporte: la primera fecha que aparezca arriba de los titulos
    var fecha = null;
    for (var ff = 0; ff < f && !fecha; ff++) {
      var llenas = datos[ff].filter(function (v) { return v !== '' && v !== null && v !== undefined; });
      llenas.forEach(function (v) { if (!fecha && Object.prototype.toString.call(v) === '[object Date]') fecha = v; });
      // En el .xlsx leido directo la fecha llega como numero de serie de Excel
      if (!fecha && llenas.length === 1 && typeof llenas[0] === 'number' && llenas[0] > 40000 && llenas[0] < 60000) {
        var d = new Date(Math.round((llenas[0] - 25569) * 86400000));
        fecha = new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
      }
    }
    return {
      columna: columna,
      fecha: fecha,
      filas: datos.slice(f + 1).map(function (r) {
        return [String(r[cCod] || ''), String(r[cNom] || ''), r[cEx],
          cSI === -1 ? '' : r[cSI], cCom === -1 ? '' : r[cCom], cCon === -1 ? '' : r[cCon]];
      })
    };
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

// Recetas de Material: el nombre de cada componente sin el porcentaje ni lo
// que va entre parentesis ("60% 05C06" -> "05C06", "2811(HD2000-GP)" -> "2811").
// Reemplaza la version anterior, que dejaba el "60%" pegado y la receta
// salia como "(definir receta)" en Consumo por Material.
function formulaComponente(fila, slot) {
  return '=IF($A' + fila + '="","",IFERROR(UPPER(TRIM(REGEXREPLACE(REGEXREPLACE(INDEX(SPLIT($A' + fila +
    ',"+"),1,' + slot + '),"\\([^)]*\\)",""),"^\\s*\\d+(\\.\\d+)?\\s*%",""))),""))';
}

// ============================================================
// HISTORICO TOTAL: junta el Seguimiento de todos los cortes guardados en el
// archivo Historico (mas el periodo actual, si tiene ordenes) en una sola
// hoja, y arma un resumen de todo junto. Se crea en el archivo Historico:
//   "📚 Todos los cortes — Seguimiento"  (todas las ordenes, una tras otra)
//   "📈 Resumen de todos los cortes"     (por corte, material, color y producto)
// Se puede volver a generar cuantas veces se quiera: siempre se rehace entero.
// ============================================================
var HOJA_TODOS_SEGUIMIENTO = '📚 Todos los cortes — Seguimiento';
var HOJA_TODOS_RESUMEN = '📈 Resumen de todos los cortes';

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Seguimiento')
    .addItem('📥 Traer órdenes de la carpeta', 'importarDesdeCarpeta')
    .addItem('🔍 Verificar cambios (sincronizar con la carpeta)', 'verificarCambios')
    .addItem('🔄 Actualizar todo', 'actualizarTodo')
    .addItem('🏠 Actualizar portada (Inicio)', 'actualizarPortada')
    .addItem('🍳 Buscar recetas de material faltantes', 'detectarRecetasFaltantesBoton')
    .addItem('📥 Leer inventario más reciente (carpeta)', 'leerInventarioDesdeCarpeta')
    .addItem('📦 Cruzar inventario (Inv)', 'actualizarInventario')
    .addItem('🧹 Eliminar filas duplicadas', 'eliminarFilasDuplicadas')
    .addSeparator()
    .addItem('🚫 Anular orden(es) seleccionada(s)', 'anularOrdenesSeleccionadas')
    .addItem('↩️ Quitar anulación', 'quitarAnulacion')
    .addSeparator()
    .addItem('📅 Corte del mes (cerrar y archivar)', 'corteDelMesYJuntar')
    .addItem('📚 Juntar todos los cortes (histórico total)', 'generarHistoricoTotal')
    .addSeparator()
    .addItem('🔒 Congelar y proteger fórmulas', 'configurarProteccionYFormato')
    .addItem('🔧 Reparar errores (#ERROR!)', 'repararErroresConocidos')
    .addItem('🎨 Colorear pestañas por tipo', 'colorearPestanas')
    .addItem('🗑️ Borrar', 'limpiarDatosSeptiembre')
    .addToUi();
}

// El corte de siempre y, si se hizo, el historico total se rehace solo.
function corteDelMesYJuntar() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Seguimiento');
  var habia = contarOrdenesSeguimiento(sh);
  var maxAntes = null;
  if (habia > 0) {
    ordenesDeHojaSeguimiento(sh).filas.forEach(function (r) {
      if (Object.prototype.toString.call(r[0]) === '[object Date]' && (!maxAntes || r[0] > maxAntes)) maxAntes = r[0];
    });
  }
  corteDelMes();
  if (habia > 0 && contarOrdenesSeguimiento(sh) === 0) {
    if (maxAntes) PropertiesService.getScriptProperties().setProperty('ULTIMA_FECHA_CORTADA',
      String(new Date(maxAntes.getFullYear(), maxAntes.getMonth(), maxAntes.getDate()).getTime()));
    try { generarHistoricoTotal(true); } catch (e) {}
  }
}

function contarOrdenesSeguimiento(sh) {
  if (!sh || sh.getLastRow() < 5) return 0;
  return sh.getRange(5, 9, sh.getLastRow() - 4, 1).getValues()
    .filter(function (r) { return typeof r[0] === 'number'; }).length;
}

// Filas de ordenes de una hoja con el formato de Seguimiento (30 columnas, A:AD).
function ordenesDeHojaSeguimiento(hoja) {
  if (!hoja || hoja.getLastRow() < 2) return { encabezado: null, filas: [] };
  var datos = hoja.getRange(1, 1, hoja.getLastRow(), Math.min(hoja.getLastColumn(), 30)).getValues();
  var fEnc = -1;
  for (var f = 0; f < Math.min(datos.length, 10) && fEnc === -1; f++) {
    if (/^FECHA EMISI/.test(normalizar(datos[f][0]))) fEnc = f;
  }
  if (fEnc === -1) return { encabezado: null, filas: [] };
  var filas = datos.slice(fEnc + 1).filter(function (r) {
    return typeof r[8] === 'number' && r[0] !== '' && String(r[15] || '').trim().toUpperCase() !== 'ANULADA';
  })
    .map(function (r) { while (r.length < 30) r.push(''); return r; });
  return { encabezado: datos[fEnc], filas: filas };
}

// Algunas ordenes viejas tienen dia y mes al reves (1 de septiembre guardado
// como 9 de enero). Si la fecha cae fuera de los 70 dias antes del corte y al
// voltearla si cae adentro, se voltea.
function corregirFechaOrden(f, fechaCorte) {
  if (Object.prototype.toString.call(f) !== '[object Date]') return f;
  var desde = new Date(fechaCorte.getTime() - 70 * 86400000), hasta = new Date(fechaCorte.getTime() + 86400000);
  if (f >= desde && f <= hasta) return f;
  if (f.getDate() > 12) return f;
  var v = new Date(f.getFullYear(), f.getDate() - 1, f.getMonth() + 1);
  return (v >= desde && v <= hasta) ? v : f;
}

function fechaDeCorte(etiqueta) {
  var m = /(\d{2})-(\d{2})-(\d{4})/.exec(etiqueta);
  return m ? new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1])) : new Date(2999, 0, 1);
}

function generarHistoricoTotal(silencioso) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var historico = obtenerArchivoHistorico();

  // 1) Juntar: cada corte guardado + el periodo actual
  var cortes = [];
  historico.getSheets().forEach(function (h) {
    var m = /^(Cierre \d{2}-\d{2}-\d{4}(?: \(\d+\))?) — Seguimiento$/.exec(h.getName());
    if (!m) return;
    var leido = ordenesDeHojaSeguimiento(h);
    var fc = fechaDeCorte(m[1]);
    leido.filas.forEach(function (r) { r[0] = corregirFechaOrden(r[0], fc); });
    if (leido.filas.length) cortes.push({ etiqueta: m[1], fecha: fc, filas: leido.filas, enc: leido.encabezado });
  });
  cortes.sort(function (a, b) { return a.fecha - b.fecha; });
  var actual = ordenesDeHojaSeguimiento(ss.getSheetByName('Seguimiento'));
  if (actual.filas.length) cortes.push({ etiqueta: 'Actual (sin cerrar)', fecha: new Date(), filas: actual.filas, enc: actual.encabezado });
  if (!cortes.length) {
    if (!silencioso) SpreadsheetApp.getUi().alert('No encontré órdenes en ningún corte del archivo "' + historico.getName() + '".');
    return;
  }

  var encabezado = ['Corte'].concat((actual.encabezado || cortes[cortes.length - 1].enc).map(function (x) { return String(x); }));
  var todas = [];
  cortes.forEach(function (c) { c.filas.forEach(function (r) { todas.push([c.etiqueta].concat(r)); }); });

  var hTodo = historico.getSheetByName(HOJA_TODOS_SEGUIMIENTO) || historico.insertSheet(HOJA_TODOS_SEGUIMIENTO, 0);
  hTodo.clear();
  if (hTodo.getMaxRows() < todas.length + 3) hTodo.insertRowsAfter(hTodo.getMaxRows(), todas.length + 3 - hTodo.getMaxRows());
  if (hTodo.getMaxColumns() < 31) hTodo.insertColumnsAfter(hTodo.getMaxColumns(), 31 - hTodo.getMaxColumns());
  hTodo.getRange(1, 1).setValue('📚 TODAS LAS ÓRDENES DE TODOS LOS CORTES  ·  ' + todas.length + ' renglones  ·  generado el ' +
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'))
    .setFontWeight('bold').setFontSize(12);
  hTodo.getRange(2, 1, 1, 31).setValues([encabezado]).setFontWeight('bold').setFontColor('#FFFFFF')
    .setBackground('#0B1F33').setWrap(true).setVerticalAlignment('middle');
  hTodo.getRange(3, 1, todas.length, 31).setValues(todas);
  hTodo.getRange(3, 2, todas.length, 1).setNumberFormat('dd/MM/yyyy');
  hTodo.getRange(3, 20, todas.length, 11).setNumberFormat('#,##0.00');
  hTodo.getRange(3, 20, todas.length, 1).setNumberFormat('0.0%');
  hTodo.setFrozenRows(2);
  hTodo.setFrozenColumns(2);

  // 2) Resumen
  var K = { fecha: 1, cant: 9, prodComp: 5, prod: 6, color: 10, codColor: 18, kgCol: 20, kgTot: 21, kgMP: 22 };
  var n = cortes.length, etiquetas = cortes.map(function (c) { return c.etiqueta; });
  var porCorte = cortes.map(function (c) {
    var fechas = {}, mp = 0, col = 0, tot = 0, min = null, max = null;
    c.filas.forEach(function (r) {
      var f = r[K.fecha - 1];
      if (Object.prototype.toString.call(f) === '[object Date]') {
        fechas[f.toDateString()] = 1;
        if (!min || f < min) min = f;
        if (!max || f > max) max = f;
      }
      mp += Number(r[K.kgMP - 1]) || 0; col += Number(r[K.kgCol - 1]) || 0; tot += Number(r[K.kgTot - 1]) || 0;
    });
    return [c.etiqueta, min || '', max || '', c.filas.length, Object.keys(fechas).length, mp, col, tot];
  });
  var sumaCol = function (i) { return porCorte.reduce(function (s, x) { return s + (Number(x[i]) || 0); }, 0); };
  porCorte.push(['TOTAL', '', '', sumaCol(3), sumaCol(4), sumaCol(5), sumaCol(6), sumaCol(7)]);

  var acumular = function (mapa, clave, nombre, i, kg) {
    if (!clave || !kg) return;
    if (!mapa[clave]) { mapa[clave] = { nombre: nombre, kg: [] }; for (var j = 0; j < n; j++) mapa[clave].kg.push(0); }
    mapa[clave].kg[i] += kg;
  };
  var mats = {}, cols = {}, prods = {}, nombresProd = {};
  cortes.forEach(function (c, i) {
    var comp = {};
    c.filas.forEach(function (r) {
      for (var k = 0; k < 4; k++) {
        var nom = String(r[22 + k * 2] || '').trim(), kg = Number(r[23 + k * 2]) || 0;
        if (!nom) continue;
        var clave = detectarCodigoInterno(nom) || normalizar(nom);
        acumular(mats, clave, infoMaterial(clave).nombre || clave, i, kg);
      }
      // Colores por nombre: los codigos de los cortes viejos no siempre coinciden
      var nomColor = String(r[K.color - 1] || '').replace(/\s*\((MB|CARGADO|CHEMITON|SIN C[ÓO]DIGO)[^)]*\)\s*$/i, '')
        .replace(/\s+/g, ' ').trim().toUpperCase();
      acumular(cols, normalizar(nomColor), nomColor, i, Number(r[K.kgCol - 1]) || 0);
      // Productos: pieza principal (componente /0) por codigo de producto
      var pc = String(r[K.prodComp - 1] || '').trim(), base = pc.split('-')[0] || String(r[K.prod - 1]);
      if (!comp[base]) comp[base] = { principal: 0, piezas: {} };
      var q = Number(r[K.cant - 1]) || 0;
      if (/\/0$/.test(pc)) comp[base].principal += q;
      var nomProd = String(r[K.prod - 1] || '').replace(/\s+/g, ' ').trim().toUpperCase();
      if (!nombresProd[base]) nombresProd[base] = { principal: [], otros: [] };
      (/\/0$/.test(pc) ? nombresProd[base].principal : nombresProd[base].otros).push(nomProd);
      comp[base].piezas[pc] = (comp[base].piezas[pc] || 0) + q;
    });
    Object.keys(comp).forEach(function (b) {
      var g = comp[b], mayor = 0;
      Object.keys(g.piezas).forEach(function (p) { if (g.piezas[p] > mayor) mayor = g.piezas[p]; });
      acumular(prods, b, b, i, g.principal > 0 ? g.principal : mayor);
    });
  });
  // Nombre de cada producto: lo que tienen en comun todos sus nombres (sin el color)
  Object.keys(prods).forEach(function (b) {
    var ns = nombresProd[b].principal.length ? nombresProd[b].principal : nombresProd[b].otros, pre = ns[0] || b;
    ns.forEach(function (x) { var k = 0; while (k < pre.length && pre.charAt(k) === x.charAt(k)) k++; pre = pre.slice(0, k); });
    // Si quedo cortado a mitad de palabra, se completa la palabra
    var primero = ns[0] || '';
    while (pre.length < primero.length && /\S/.test(primero.charAt(pre.length)) && /\S$/.test(pre)) pre += primero.charAt(pre.length);
    var limpio = pre.replace(/(?:[\s\/]+(?:C\/?|COLOR|DE|CON))?[\s\/]*$/, '').trim();
    prods[b].nombre = (limpio.length < 4 ? (ns[0] || b) : limpio) + '  (' + b + ')';
  });
  var tabla = function (mapa, tope) {
    return Object.keys(mapa).map(function (k) {
      var t = mapa[k].kg.reduce(function (s, x) { return s + x; }, 0);
      return [mapa[k].nombre].concat(mapa[k].kg).concat([t]);
    }).sort(function (a, b) { return b[b.length - 1] - a[a.length - 1]; }).slice(0, tope || 1000);
  };

  var hRes = historico.getSheetByName(HOJA_TODOS_RESUMEN) || historico.insertSheet(HOJA_TODOS_RESUMEN, 0);
  hRes.clear();
  var bloques = [
    { titulo: '📅 POR CORTE', enc: ['Corte', 'Desde', 'Hasta', 'Renglones', 'Días con producción', 'Kg materia prima', 'Kg colorante', 'Kg total'], filas: porCorte, color: '#0B1F33' },
    { titulo: '🧪 MATERIA PRIMA CONSUMIDA (kg)', enc: ['Material'].concat(etiquetas).concat(['TOTAL']), filas: tabla(mats), color: '#1DB37A' },
    { titulo: '🎨 COLORANTE CONSUMIDO (kg)', enc: ['Color'].concat(etiquetas).concat(['TOTAL']), filas: tabla(cols), color: '#F2436B' },
    { titulo: '🏭 PRODUCTOS MÁS FABRICADOS (piezas, producto completo)', enc: ['Producto'].concat(etiquetas).concat(['TOTAL']), filas: tabla(prods, 40), color: '#00A6E0' }
  ];
  var ancho = Math.max(8, n + 2);
  var filasNec = bloques.reduce(function (s, b) { return s + b.filas.length + 4; }, 4);
  if (hRes.getMaxRows() < filasNec) hRes.insertRowsAfter(hRes.getMaxRows(), filasNec - hRes.getMaxRows());
  if (hRes.getMaxColumns() < ancho) hRes.insertColumnsAfter(hRes.getMaxColumns(), ancho - hRes.getMaxColumns());
  hRes.getRange(1, 1).setValue('📈 RESUMEN DE TODOS LOS CORTES  ·  ' + etiquetas.join('  ·  '))
    .setFontWeight('bold').setFontSize(13);
  hRes.getRange(2, 1).setValue('Generado el ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') +
    '. Se rehace con el menú Seguimiento ▸ 📚 Juntar todos los cortes.').setFontColor('#5C6B7A');
  var r = 4;
  bloques.forEach(function (b) {
    hRes.getRange(r, 1).setValue(b.titulo).setFontWeight('bold').setFontSize(11).setFontColor(b.color);
    r++;
    hRes.getRange(r, 1, 1, b.enc.length).setValues([b.enc]).setFontWeight('bold').setFontColor('#FFFFFF')
      .setBackground(b.color).setWrap(true).setHorizontalAlignment('center');
    r++;
    if (b.filas.length) {
      hRes.getRange(r, 1, b.filas.length, b.enc.length).setValues(b.filas);
      hRes.getRange(r, 2, b.filas.length, b.enc.length - 1).setNumberFormat('#,##0');
      if (b.titulo.indexOf('POR CORTE') !== -1) {
        hRes.getRange(r, 2, b.filas.length, 2).setNumberFormat('dd/MM/yyyy');
        hRes.getRange(r + b.filas.length - 1, 1, 1, b.enc.length).setFontWeight('bold').setBackground('#EEF2F7');
      }
      hRes.getRange(r, b.enc.length, b.filas.length, 1).setFontWeight('bold');
      r += b.filas.length;
    }
    r += 2;
  });
  hRes.setColumnWidth(1, 300);
  for (var c = 2; c <= ancho; c++) hRes.setColumnWidth(c, 115);
  hRes.setFrozenColumns(1);

  // Link desde "Historico de Cortes" de este archivo
  var hc = ss.getSheetByName('Histórico de Cortes');
  if (hc) {
    var url = 'https://docs.google.com/spreadsheets/d/' + historico.getId() + '/edit#gid=';
    hc.getRange(1, 10).setFormula('=HYPERLINK("' + url + hRes.getSheetId() + '","📈 Ver resumen de todos los cortes")').setFontWeight('bold');
    hc.getRange(2, 10).setFormula('=HYPERLINK("' + url + hTodo.getSheetId() + '","📚 Ver todas las órdenes juntas")').setFontWeight('bold');
    hc.setColumnWidth(10, 260);
  }

  if (!silencioso) {
    SpreadsheetApp.getUi().alert('📚 Histórico total listo',
      'Junté ' + cortes.length + ' corte(s): ' + etiquetas.join(', ') + '.\n\n' +
      '- ' + todas.length + ' renglones de órdenes en "' + HOJA_TODOS_SEGUIMIENTO + '"\n' +
      '- Resumen por corte, material, color y producto en "' + HOJA_TODOS_RESUMEN + '"\n\n' +
      'Están en el archivo "' + historico.getName() + '". En "Histórico de Cortes" (columna J) tienes los links directos.',
      SpreadsheetApp.getUi().ButtonSet.OK);
  }
}


// Traer ordenes de la carpeta: version corregida (no vuelve a traer los
// dias que ya se cerraron en un corte y no falla si una orden viene repetida).
function importarDesdeCarpetaCore() {
  var CARPETA_ID = "1O0DMmXaRL6dfLTsXqbA4bdx7aIYm7Wwh";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName("Seguimiento");
  var lastRow = sh.getLastRow();
  var existentes = {};
  if (lastRow >= 5) {
    var datosActuales = sh.getRange(5, 1, lastRow - 4, 16).getValues();
    var notasSacos = sh.getRange(5, 11, lastRow - 4, 1).getNotes();
    for (var idxExist = 0; idxExist < datosActuales.length; idxExist++) {
      var rExist = datosActuales[idxExist];
      var anulada = String(rExist[15] || "").trim().toUpperCase() === "ANULADA";
      // En una anulada, los sacos reales estan en la nota (para comparar bien)
      var guardados = anulada ? leerNotaSacos(notasSacos[idxExist][0]) : null;
      if (guardados) { rExist[10] = guardados[0]; rExist[11] = guardados[1]; }
      rExist = rExist.slice(0, 15);
      var fecha = rExist[0], op = rExist[1], sku = rExist[3], color = rExist[9];
      if (fecha && op) {
        var keyExist = claveOrden(formatearFecha(fecha), op, sku, color);
        existentes[keyExist] = {
          row: 5 + idxExist,
          firma: firmaOrden(rExist.slice(2)),
          anulada: anulada
        };
      }
    }
  }
  var MESES = {
    ENERO: 1,
    FEBRERO: 2,
    MARZO: 3,
    ABRIL: 4,
    MAYO: 5,
    JUNIO: 6,
    JULIO: 7,
    AGOSTO: 8,
    SEPTIEMBRE: 9,
    OCTUBRE: 10,
    NOVIEMBRE: 11,
    DICIEMBRE: 12
  };
  var folder = DriveApp.getFolderById(CARPETA_ID);
  var files = folder.getFiles();
  var nuevasFilas = [];
  var filasActualizar = [];
  var archivosProcesados = 0;
  var archivosOmitidos = [];
  var detallePorArchivo = [];
  var tiempoInicio = (new Date).getTime();
  var LIMITE_MS = 4.5 * 60 * 1e3;
  var seCortoPorTiempo = false;
  var clavesOrigen = {}, fechasLeidas = {}, fechasConError = {};
  // Los dias que ya se cerraron en un corte no se vuelven a traer
  var ultimaCortada = ultimaFechaArchivada();
  var archivosYaCortados = 0;
  while (files.hasNext()) {
    if ((new Date).getTime() - tiempoInicio > LIMITE_MS) {
      seCortoPorTiempo = true;
      break;
    }
    var file = files.next();
    if (file.getMimeType() !== MimeType.GOOGLE_SHEETS) continue;
    var titulo = file.getName();
    var fecha = parsearFecha(titulo, MESES);
    if (!fecha) {
      archivosOmitidos.push(titulo + " (no se pudo leer la fecha del nombre del archivo)");
      continue;
    }
    if (ultimaCortada && new Date(fecha.anio, fecha.mes - 1, fecha.dia) <= ultimaCortada) {
      archivosYaCortados++;
      continue;
    }
    var fechaTexto = (fecha.dia < 10 ? "0" + fecha.dia : fecha.dia) + "/" + (fecha.mes < 10 ? "0" + fecha.mes : fecha.mes) + "/" + fecha.anio;
    var fechaISO = fecha.anio + "-" + (fecha.mes < 10 ? "0" + fecha.mes : fecha.mes) + "-" + (fecha.dia < 10 ? "0" + fecha.dia : fecha.dia);
    var libroExterno;
    try {
      libroExterno = SpreadsheetApp.openById(file.getId());
    } catch (e) {
      archivosOmitidos.push(titulo + " (no se pudo abrir: " + e.message + ")");
      fechasConError[fechaTexto] = true;
      continue;
    }
    var maestro = libroExterno.getSheetByName("Maestro");
    if (!maestro) {
      archivosOmitidos.push(titulo + ' (no tiene pesta\xf1a "Maestro")');
      fechasConError[fechaTexto] = true;
      continue;
    }
    var datos;
    try {
      var filasALeer = Math.min(250, maestro.getMaxRows());
      datos = maestro.getRange(1, 1, filasALeer, 15).getValues();
    } catch (e) {
      archivosOmitidos.push(titulo + " (error leyendo Maestro: " + e.message + ")");
      fechasConError[fechaTexto] = true;
      continue;
    }
    var opActual = null;
    var filasEsteArchivo = 0;
    var filasActualizadasEsteArchivo = 0;
    var filasYaExistianEsteArchivo = 0;
    for (var i = 1; i < datos.length; i++) {
      var fila = datos[i];
      var opCelda = fila[0];
      if (opCelda !== "" && opCelda !== null) opActual = opCelda;
      var cantidad = fila[8];
      if (cantidad === "" || cantidad === null) continue;
      var sku = fila[3], color = fila[9];
      var key = claveOrden(fechaTexto, opActual, sku, color);
      clavesOrigen[key] = true;
      var camposComparables = [ fila[1], fila[3], fila[4], fila[5], fila[6], fila[7], fila[8], fila[9], fila[10], fila[11], fila[12], fila[13], fila[14] ];
      var firmaNueva = firmaOrden(camposComparables);
      var existente = existentes[key];
      if (existente && existente.row === null) {
        // Ya se agrego en esta misma lectura (la orden viene repetida en otro
        // archivo): se queda con los datos mas recientes, sin fila nueva.
        if (existente.firma !== firmaNueva) {
          nuevasFilas[existente.idx] = [ fechaISO, opActual, fila[1], fila[3], fila[4], fila[5], fila[6], fila[7], fila[8], fila[9], fila[10], fila[11], fila[12], fila[13], fila[14], "", "" ];
          existente.firma = firmaNueva;
        }
        filasYaExistianEsteArchivo++;
        continue;
      }
      if (existente) {
        if (existente.firma === firmaNueva) {
          filasYaExistianEsteArchivo++;
        } else {
          filasActualizar.push({
            row: existente.row,
            valores: camposComparables,
            anulada: existente.anulada
          });
          existentes[key] = {
            row: existente.row,
            firma: firmaNueva,
            anulada: existente.anulada
          };
          filasActualizadasEsteArchivo++;
        }
        continue;
      }
      existentes[key] = {
        row: null,
        firma: firmaNueva,
        idx: nuevasFilas.length
      };
      filasEsteArchivo++;
      nuevasFilas.push([ fechaISO, opActual, fila[1], fila[3], fila[4], fila[5], fila[6], fila[7], fila[8], fila[9], fila[10], fila[11], fila[12], fila[13], fila[14], "", "" ]);
    }
    archivosProcesados++;
    fechasLeidas[fechaTexto] = true;
    var detalleTxt = titulo + " [" + fechaTexto + "]: " + filasEsteArchivo + " nuevas";
    if (filasActualizadasEsteArchivo > 0) detalleTxt += ", " + filasActualizadasEsteArchivo + " actualizadas";
    if (filasYaExistianEsteArchivo > 0) detalleTxt += ", " + filasYaExistianEsteArchivo + " sin cambios";
    detallePorArchivo.push(detalleTxt);
  }
  if (filasActualizar.length > 0) {
    filasActualizar.filter(function (u) { return typeof u.row === "number"; }).forEach((function(u) {
      sh.getRange(u.row, 4, 1, 1).setNumberFormat("@");
      var vals = u.valores.map(valorParaHoja);
      if (u.anulada) {
        // Orden anulada: los sacos nuevos se guardan en la nota, no en la celda
        sh.getRange(u.row, 11).setNote(notaSacos(vals[8], vals[9]));
        vals[8] = ""; vals[9] = "";
      }
      sh.getRange(u.row, 3, 1, 13).setValues([ vals ]);
    }));
    SpreadsheetApp.flush();
  }
  if (nuevasFilas.length > 0) {
    var filaDestino = sh.getLastRow() + 1;
    var faltan = filaDestino + nuevasFilas.length - 1 - sh.getMaxRows();
    if (faltan > 0) sh.insertRowsAfter(sh.getMaxRows(), faltan + 200);
    sh.getRange(filaDestino, 4, nuevasFilas.length, 1).setNumberFormat("@");
    sh.getRange(filaDestino, 1, nuevasFilas.length, 17).setValues(nuevasFilas.map((function(f) {
      return f.map((function(v, i) {
        return i >= 2 ? valorParaHoja(v) : v;
      }));
    })));
    sh.getRange(filaDestino, 1, nuevasFilas.length, 1).setNumberFormat("dd/mm/yyyy");
    var plantilla = sh.getRange(5, 18, 1, 13);
    plantilla.copyTo(sh.getRange(filaDestino, 18, nuevasFilas.length, 13));
    SpreadsheetApp.flush();
    var rangoNuevas = sh.getRange(filaDestino, 1, nuevasFilas.length, sh.getLastColumn());
    sh.moveRows(rangoNuevas, 5);
    SpreadsheetApp.flush();
  }
  actualizarResumenColores();
  actualizarResumenMateriales();
  actualizarInventarioCore();
  actualizarDashboard();
  var filasConError = detectarFilasConError();
  SpreadsheetApp.flush();
  actualizarPortada();
  var mensaje = "Importaci\xf3n desde la carpeta de Drive \u2014 " + (new Date).toLocaleString() + "\n" + "- Archivos revisados: " + archivosProcesados + "\n" + "- \xd3rdenes nuevas agregadas: " + nuevasFilas.length + "\n" + "- \xd3rdenes actualizadas (ya exist\xedan, cambi\xf3 alg\xfan dato): " + filasActualizar.length + "\n\n" + "Detalle por archivo:\n" + detallePorArchivo.join("\n");
  if (seCortoPorTiempo) {
    mensaje += "\n\n\u23f1\ufe0f Se detuvo antes de terminar todos los archivos para no pasarse del " + "l\xedmite de tiempo de Google. No hay problema \u2014 dale clic otra vez y sigue justo donde " + "se qued\xf3, sin duplicar lo que ya trajo.";
  }
  if (archivosYaCortados > 0) {
    mensaje += "\n\n📅 " + archivosYaCortados + " archivo(s) de d\xedas que ya se cerraron en un corte (hasta el " +
      Utilities.formatDate(ultimaCortada, Session.getScriptTimeZone(), "dd/MM/yyyy") + ") no se volvieron a traer.";
  }
  if (archivosOmitidos.length > 0) {
    mensaje += "\n\n\u26a0\ufe0f Archivos que no pude leer (rev\xedsalos a mano):\n" + archivosOmitidos.join("\n");
  }
  if (filasConError.length > 0) {
    mensaje += "\n\n\u26a0\ufe0f Filas con dato faltante o mal escrito (revisa Cantidad y Peso): " + filasConError.join(", ");
  }
  return {
    mensaje: mensaje,
    nuevas: nuevasFilas.length,
    actualizadas: filasActualizar.length,
    clavesOrigen: clavesOrigen,
    fechasLeidas: fechasLeidas,
    fechasConError: fechasConError,
    seCortoPorTiempo: seCortoPorTiempo
  };
}

// Ultimo dia que ya quedo guardado en un corte. Se calcula del archivo
// Historico (la fecha de orden mas reciente del ultimo corte) y se recuerda.
function ultimaFechaArchivada() {
  var props = PropertiesService.getScriptProperties();
  var guardada = props.getProperty('ULTIMA_FECHA_CORTADA');
  if (guardada) return new Date(Number(guardada));
  try {
    var historico = obtenerArchivoHistorico(), mejor = null, hoja = null;
    historico.getSheets().forEach(function (h) {
      var m = /^Cierre (\d{2})-(\d{2})-(\d{4})(?: \(\d+\))? — Seguimiento$/.exec(h.getName());
      if (!m) return;
      var f = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
      if (!mejor || f > mejor) { mejor = f; hoja = h; }
    });
    if (!hoja) return null;
    var max = null;
    ordenesDeHojaSeguimiento(hoja).filas.forEach(function (r) {
      var f = corregirFechaOrden(r[0], mejor);
      if (Object.prototype.toString.call(f) === '[object Date]' && (!max || f > max)) max = f;
    });
    if (max) {
      max = new Date(max.getFullYear(), max.getMonth(), max.getDate());
      props.setProperty('ULTIMA_FECHA_CORTADA', String(max.getTime()));
    }
    return max;
  } catch (e) {
    return null;
  }
}


// ============================================================
// ORDENES ANULADAS: en vez de borrar una orden se marca "ANULADA" en la
// columna P (Estado). La fila se queda en la hoja con una franja roja y no se
// cuenta en nada (sus kg quedan vacios y los conteos la saltan).
// ============================================================
function filasSeleccionadasSeguimiento() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getActiveSheet();
  if (sh.getName() !== 'Seguimiento') return null;
  var filas = {};
  var lista = sh.getActiveRangeList();
  (lista ? lista.getRanges() : [sh.getActiveRange()]).forEach(function (r) {
    for (var f = r.getRow(); f < r.getRow() + r.getNumRows(); f++) if (f >= 5 && f <= sh.getLastRow()) filas[f] = true;
  });
  return { hoja: sh, filas: Object.keys(filas).map(Number).sort(function (a, b) { return a - b; }) };
}

function marcarAnulacion(anular) {
  var ui = SpreadsheetApp.getUi();
  var sel = filasSeleccionadasSeguimiento();
  if (!sel) {
    ui.alert('Primero ve a la pestaña "Seguimiento" y selecciona la fila (o las filas) de la orden.');
    return;
  }
  var sh = sel.hoja, cambiadas = [];
  sel.filas.forEach(function (f) {
    var fila = sh.getRange(f, 1, 1, 16).getValues()[0];
    if (fila[8] === '' || fila[8] === null) return;
    var estado = String(fila[15] || '').trim().toUpperCase();
    if (anular && estado !== 'ANULADA') { sh.getRange(f, 16).setValue('ANULADA'); guardarYVaciarSacos(sh, f); cambiadas.push(f); }
    if (!anular && estado === 'ANULADA') { sh.getRange(f, 16).setValue(''); devolverSacos(sh, f); cambiadas.push(f); }
  });
  if (!cambiadas.length) {
    ui.alert(anular ? 'Las filas seleccionadas no tienen órdenes o ya estaban anuladas.' : 'Ninguna de las filas seleccionadas estaba anulada.');
    return;
  }
  SpreadsheetApp.flush();
  actualizarResumenColores();
  actualizarResumenMateriales();
  actualizarInventarioCore();
  SpreadsheetApp.flush();
  ui.alert((anular ? '🚫 Anulada(s) ' : '↩️ Se quitó la anulación de ') + cambiadas.length + ' orden(es): fila(s) ' + cambiadas.join(', ') + '.\n\n' +
    (anular ? 'Quedan en la hoja con franja roja y ya no se cuentan en consumo, inventario ni Dashboard. ' +
              'Los SACOS y SACOS (kg) quedaron vacíos (lo que tenían está guardado en la nota de la celda de SACOS).'
            : 'Vuelven a contarse normal en consumo, inventario y Dashboard, y recuperaron sus SACOS y SACOS (kg).'));
}

function anularOrdenesSeleccionadas() { marcarAnulacion(true); }
function quitarAnulacion() { marcarAnulacion(false); }

// Plantilla de la fila 5 (columnas R a AD), igual a la que esta en la hoja, con
// los kg en blanco cuando la orden esta ANULADA. Reemplaza la de la parte 3.
function escribirPlantillaFila5(sh) {
  var formulas = ["=IF(TRIM(J5)=\"\",\"\",LET(txt,UPPER(TRIM(J5)),grupo,IFERROR(VLOOKUP(TRIM(J5),Colores!$A$2:$H,8,0),\"\"),cod,IFERROR(REGEXEXTRACT(txt,\"\\((?:MB|CARGADO|CHEMITON)\\s*(\\d{4})\\)\"),\"\"),grupoCod,IF(OR(grupo<>\"\",cod=\"\"),grupo,IFERROR(INDEX(FILTER(Colores!$H$2:$H,RIGHT(Colores!$C$2:$C,4)=cod,Colores!$H$2:$H<>\"\"),1),\"\")),primero,IF(grupoCod=\"\",\"\",TRIM(INDEX(SPLIT(grupoCod,\",\"),1,1))),base,TRIM(REGEXREPLACE(txt,\"\\s*\\([^()]*(?:MB|CARGADO|CHEMITON|SIN C[ÓO]DIGO)[^()]*\\)\\s*$\",\"\")),alias,SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(base,\"STANSHOME\",\"STANHOME\"),\"NATURAL 70%\",\"NATURAL (70%)\"),\"VERDE PASTEL\",\"VERDE MENTA (PASTEL)\"),nom,IF(primero<>\"\",primero,alias),esMueble,OR(ISNUMBER(SEARCH(\"11H01A+05C06+REF+MOLIDO\",SUBSTITUTE(G5,\" \",\"\"))),REGEXMATCH(UPPER(F5),\"SILLA|ESCALEBANCO|BUTACA|\\bMESA\")),nomFinal,IF(nom=\"ARENA\",IF(esMueble,\"ARENA (CON UV)\",\"ARENA (SIN UV)\"),IF(AND(esMueble,ISNUMBER(MATCH(nom&\" (SILLAS)\",'Catálogo Colores'!$B$2:$B$500,0))),nom&\" (SILLAS)\",nom)),IFERROR(INDEX('Catálogo Colores'!$E$2:$E$500,MATCH(nomFinal,'Catálogo Colores'!$B$2:$B$500,0)),\"(sin mapear)\")))", "=IF(OR(ISNUMBER(SEARCH(\"molido\",G5)),ISNUMBER(SEARCH(\"molino\",G5))),0.015,0.02)", "=IFERROR(IF(OR(I5=\"\",UPPER(TRIM(P5))=\"ANULADA\"),\"\",I5*N5/1000*S5),\"(revisar dato)\")", "=IFERROR(IF(OR(I5=\"\",UPPER(TRIM(P5))=\"ANULADA\"),\"\",I5*N5/1000),\"(revisar dato)\")", "=IFERROR(IF(U5=\"\",\"\",U5-T5),\"(revisar dato)\")", "=IF(G5=\"\",\"\",IF(ISNUMBER(SEARCH(\"+\",G5)),IFERROR(INDEX('Recetas de Material'!$B$3:$B$198,MATCH(REGEXREPLACE(UPPER(G5),\"\\([^)]*\\)|\\s\",\"\"),ARRAYFORMULA(REGEXREPLACE(UPPER('Recetas de Material'!$A$3:$A$198),\"\\([^)]*\\)|\\s\",\"\")),0)),\"(definir receta)\"),G5))", "=IF(G5=\"\",\"\",IF(V5=\"\",\"\",IFERROR(LET(f,MATCH(REGEXREPLACE(UPPER(G5),\"\\([^)]*\\)|\\s\",\"\"),ARRAYFORMULA(REGEXREPLACE(UPPER('Recetas de Material'!$A$3:$A$198),\"\\([^)]*\\)|\\s\",\"\")),0),t,SUM(INDEX('Recetas de Material'!$C$3:$C$198,f),INDEX('Recetas de Material'!$E$3:$E$198,f),INDEX('Recetas de Material'!$G$3:$G$198,f),INDEX('Recetas de Material'!$I$3:$I$198,f)),IF(NOT(ISNUMBER(SEARCH(\"+\",G5))),V5,IF(t=0,0,V5*N(INDEX('Recetas de Material'!$C$3:$C$198,f))/t))),V5)))", "=IF(G5=\"\",\"\",IF(AND(ISNUMBER(SEARCH(\"+\",G5)),ISNUMBER(MATCH(REGEXREPLACE(UPPER(G5),\"\\([^)]*\\)|\\s\",\"\"),ARRAYFORMULA(REGEXREPLACE(UPPER('Recetas de Material'!$A$3:$A$198),\"\\([^)]*\\)|\\s\",\"\")),0))),IFERROR(INDEX('Recetas de Material'!$D$3:$D$198,MATCH(REGEXREPLACE(UPPER(G5),\"\\([^)]*\\)|\\s\",\"\"),ARRAYFORMULA(REGEXREPLACE(UPPER('Recetas de Material'!$A$3:$A$198),\"\\([^)]*\\)|\\s\",\"\")),0)),\"\"),\"\"))", "=IF(G5=\"\",\"\",IF(V5=\"\",\"\",IF(ISNUMBER(SEARCH(\"+\",G5)),IFERROR(LET(f,MATCH(REGEXREPLACE(UPPER(G5),\"\\([^)]*\\)|\\s\",\"\"),ARRAYFORMULA(REGEXREPLACE(UPPER('Recetas de Material'!$A$3:$A$198),\"\\([^)]*\\)|\\s\",\"\")),0),t,SUM(INDEX('Recetas de Material'!$C$3:$C$198,f),INDEX('Recetas de Material'!$E$3:$E$198,f),INDEX('Recetas de Material'!$G$3:$G$198,f),INDEX('Recetas de Material'!$I$3:$I$198,f)),IF(t=0,0,V5*N(INDEX('Recetas de Material'!$E$3:$E$198,f))/t)),0),0)))", "=IF(G5=\"\",\"\",IF(AND(ISNUMBER(SEARCH(\"+\",G5)),ISNUMBER(MATCH(REGEXREPLACE(UPPER(G5),\"\\([^)]*\\)|\\s\",\"\"),ARRAYFORMULA(REGEXREPLACE(UPPER('Recetas de Material'!$A$3:$A$198),\"\\([^)]*\\)|\\s\",\"\")),0))),IFERROR(INDEX('Recetas de Material'!$F$3:$F$198,MATCH(REGEXREPLACE(UPPER(G5),\"\\([^)]*\\)|\\s\",\"\"),ARRAYFORMULA(REGEXREPLACE(UPPER('Recetas de Material'!$A$3:$A$198),\"\\([^)]*\\)|\\s\",\"\")),0)),\"\"),\"\"))", "=IF(G5=\"\",\"\",IF(V5=\"\",\"\",IF(ISNUMBER(SEARCH(\"+\",G5)),IFERROR(LET(f,MATCH(REGEXREPLACE(UPPER(G5),\"\\([^)]*\\)|\\s\",\"\"),ARRAYFORMULA(REGEXREPLACE(UPPER('Recetas de Material'!$A$3:$A$198),\"\\([^)]*\\)|\\s\",\"\")),0),t,SUM(INDEX('Recetas de Material'!$C$3:$C$198,f),INDEX('Recetas de Material'!$E$3:$E$198,f),INDEX('Recetas de Material'!$G$3:$G$198,f),INDEX('Recetas de Material'!$I$3:$I$198,f)),IF(t=0,0,V5*N(INDEX('Recetas de Material'!$G$3:$G$198,f))/t)),0),0)))", "=IF(G5=\"\",\"\",IF(AND(ISNUMBER(SEARCH(\"+\",G5)),ISNUMBER(MATCH(REGEXREPLACE(UPPER(G5),\"\\([^)]*\\)|\\s\",\"\"),ARRAYFORMULA(REGEXREPLACE(UPPER('Recetas de Material'!$A$3:$A$198),\"\\([^)]*\\)|\\s\",\"\")),0))),IFERROR(INDEX('Recetas de Material'!$H$3:$H$198,MATCH(REGEXREPLACE(UPPER(G5),\"\\([^)]*\\)|\\s\",\"\"),ARRAYFORMULA(REGEXREPLACE(UPPER('Recetas de Material'!$A$3:$A$198),\"\\([^)]*\\)|\\s\",\"\")),0)),\"\"),\"\"))", "=IF(G5=\"\",\"\",IF(V5=\"\",\"\",IF(ISNUMBER(SEARCH(\"+\",G5)),IFERROR(LET(f,MATCH(REGEXREPLACE(UPPER(G5),\"\\([^)]*\\)|\\s\",\"\"),ARRAYFORMULA(REGEXREPLACE(UPPER('Recetas de Material'!$A$3:$A$198),\"\\([^)]*\\)|\\s\",\"\")),0),t,SUM(INDEX('Recetas de Material'!$C$3:$C$198,f),INDEX('Recetas de Material'!$E$3:$E$198,f),INDEX('Recetas de Material'!$G$3:$G$198,f),INDEX('Recetas de Material'!$I$3:$I$198,f)),IF(t=0,0,V5*N(INDEX('Recetas de Material'!$I$3:$I$198,f))/t)),0),0)))"];
  sh.getRange(5, 18, 1, formulas.length).setFormulas([formulas]);
  SpreadsheetApp.flush();
}

// Igual que la de la parte 1, pero sin contar las ordenes anuladas
// (la usa el corte del mes para los totales de "Historico de Cortes").
function calcularDatosDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName("Seguimiento");
  var lastRow = sh.getLastRow();
  var porFecha = {};
  var porProducto = {};
  var porColor = {};
  var totalOrdenes = 0, totalKgColor = 0, totalKgMp = 0;
  var productosDistintos = {};
  if (lastRow >= 5) {
    var n = lastRow - 4;
    var bloque = sh.getRange(5, 1, n, 22).getValues();
    for (var i = 0; i < n; i++) {
      var fila = bloque[i];
      var cantidad = fila[8];
      if (cantidad === "" || cantidad === null) continue;
      if (String(fila[15] || "").trim().toUpperCase() === "ANULADA") continue;
      var fecha = fila[0];
      var producto = fila[5];
      var color = fila[9];
      var kgColor = fila[19];
      var kgMp = fila[21];
      kgColor = typeof kgColor === "number" ? kgColor : 0;
      kgMp = typeof kgMp === "number" ? kgMp : 0;
      totalOrdenes++;
      totalKgColor += kgColor;
      totalKgMp += kgMp;
      if (producto) productosDistintos[producto] = true;
      var fKey = Object.prototype.toString.call(fecha) === "[object Date]" ? Utilities.formatDate(fecha, Session.getScriptTimeZone(), "yyyy-MM-dd") : String(fecha || "").trim();
      if (fKey) porFecha[fKey] = (porFecha[fKey] || 0) + kgColor + kgMp;
      if (producto) {
        if (!porProducto[producto]) porProducto[producto] = {
          cantidad: 0,
          kgMp: 0
        };
        porProducto[producto].cantidad += cantidad;
        porProducto[producto].kgMp += kgMp;
      }
      if (color) porColor[color] = (porColor[color] || 0) + kgColor;
    }
  }
  var diasDistintos = Object.keys(porFecha).length;
  var productosCount = Object.keys(productosDistintos).length;
  var fechasOrdenadas = Object.keys(porFecha).sort();
  var topProductos = Object.keys(porProducto).map((function(p) {
    return [ p, porProducto[p].cantidad ];
  })).sort((function(a, b) {
    return b[1] - a[1];
  })).slice(0, 8);
  var topColores = Object.keys(porColor).map((function(c) {
    return [ c, Math.round(porColor[c] * 100) / 100 ];
  })).sort((function(a, b) {
    return b[1] - a[1];
  })).slice(0, 8);
  return {
    totalOrdenes: totalOrdenes,
    diasDistintos: diasDistintos,
    productosCount: productosCount,
    totalKgMp: totalKgMp,
    totalKgColor: totalKgColor,
    fechas: fechasOrdenadas,
    kgPorFecha: fechasOrdenadas.map((function(f) {
      return Math.round(porFecha[f] * 100) / 100;
    })),
    topProductos: topProductos,
    topColores: topColores
  };
}


// SACOS (K) y SACOS kg (L) de una orden anulada se vacian para que no se
// sumen; lo que tenian queda guardado en la nota de la celda de SACOS.
var MARCA_NOTA_SACOS = 'ANULADA · antes:';

function guardarYVaciarSacos(sh, fila) {
  var celda = sh.getRange(fila, 11, 1, 2), v = celda.getValues()[0];
  if (v[0] === '' && v[1] === '' && String(sh.getRange(fila, 11).getNote()).indexOf(MARCA_NOTA_SACOS) === 0) return;
  sh.getRange(fila, 11).setNote(notaSacos(v[0], v[1]));
  celda.setValues([['', '']]);
}

function notaSacos(sacos, kg) {
  return MARCA_NOTA_SACOS + ' SACOS=' + sacos + ' · SACOS(kg)=' + kg + '\n(Se devuelven solos al quitar la anulación)';
}

function leerNotaSacos(nota) {
  nota = String(nota || '');
  if (nota.indexOf(MARCA_NOTA_SACOS) !== 0) return null;
  var a = /SACOS=([^·\n]*)/.exec(nota), b = /SACOS\(kg\)=([^·\n]*)/.exec(nota);
  var num = function (x) { x = String(x || '').trim(); return x === '' || isNaN(Number(x)) ? x : Number(x); };
  return [a ? num(a[1]) : '', b ? num(b[1]) : ''];
}

function devolverSacos(sh, fila) {
  var guardado = leerNotaSacos(sh.getRange(fila, 11).getNote());
  if (!guardado) return;
  sh.getRange(fila, 11, 1, 2).setValues([guardado]);
  sh.getRange(fila, 11).clearNote();
}

// Si se escribe o se borra "ANULADA" a mano en la columna Estado (P)
function onEdit(e) {
  try {
    var r = e && e.range;
    if (!r || r.getSheet().getName() !== 'Seguimiento') return;
    var c1 = r.getColumn(), c2 = c1 + r.getNumColumns() - 1;
    if (c1 > 16 || c2 < 16) return;
    var sh = r.getSheet();
    for (var f = Math.max(5, r.getRow()); f < r.getRow() + r.getNumRows(); f++) {
      var estado = String(sh.getRange(f, 16).getValue() || '').trim().toUpperCase();
      if (estado === 'ANULADA') guardarYVaciarSacos(sh, f);
      else devolverSacos(sh, f);
    }
  } catch (err) {}
}
