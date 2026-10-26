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
