/**
 * PÁGINA WEB "RECETAS Y MATERIALES" — Seguimiento Manaplas
 * ---------------------------------------------------------------
 * Buscador de productos, recetas y materias primas, con imagen.
 * Lee (no modifica) estas pestañas:
 *   - Recetas de Material   → composición de cada material mezclado
 *   - Seguimiento           → qué productos usan cada material y cuántos kg
 *   - Consumo por Material  → existencia, consumo promedio y días de cobertura
 *   - Imágenes (opcional)   → Clave | URL de la imagen
 *
 * Se abre con el mismo link de la aplicación web agregando ?page=recetas
 * (ver el cambio de una línea en doGet que explica el README).
 */

// Llamada desde doGet cuando el link trae ?page=recetas
function paginaRecetas() {
  var html = HtmlService.createTemplateFromFile('RecetasHTML');
  html.datos = datosRecetas();
  html.ahora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  return html.evaluate()
    .setTitle('Recetas y materiales — Manaplas')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function datosRecetas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tz = Session.getScriptTimeZone();

  function num(v) {
    if (typeof v === 'number') return v;
    var n = parseFloat(String(v == null ? '' : v).replace(/,/g, ''));
    return isNaN(n) ? null : n;
  }
  function txt(v) { return String(v == null ? '' : v).replace(/\s+/g, ' ').trim(); }

  // --- Recetas de Material (fila 3 en adelante) ---
  var recetas = [];
  var shR = ss.getSheetByName('Recetas de Material');
  if (shR && shR.getLastRow() >= 3) {
    shR.getRange(3, 1, shR.getLastRow() - 2, 10).getValues().forEach(function (r) {
      var material = txt(r[0]);
      if (!material) return;
      var comps = [];
      for (var k = 0; k < 4; k++) {
        var c = txt(r[1 + k * 2]);
        if (c) comps.push({ c: c, p: num(r[2 + k * 2]) || 0 });
      }
      recetas.push({ material: material, componentes: comps, total: num(r[9]) });
    });
  }

  // --- Seguimiento (fila 5 en adelante, columnas A..AD) ---
  var ordenes = [];
  var shS = ss.getSheetByName('Seguimiento');
  if (shS && shS.getLastRow() >= 5) {
    shS.getRange(5, 1, shS.getLastRow() - 4, 30).getValues().forEach(function (r) {
      var cantidad = num(r[8]);
      if (cantidad === null || cantidad === 0) return;
      var fecha = r[0];
      var comps = [];
      for (var k = 0; k < 4; k++) {
        var c = txt(r[22 + k * 2]);
        var kg = num(r[23 + k * 2]) || 0;
        if (c && kg) comps.push([c, Math.round(kg * 100) / 100]);
      }
      ordenes.push([
        Object.prototype.toString.call(fecha) === '[object Date]' ? Utilities.formatDate(fecha, tz, 'dd/MM/yyyy') : txt(fecha),
        txt(r[1]),            // 1 OP
        txt(r[2]),            // 2 Máquina
        txt(r[3]),            // 3 SKU
        txt(r[5]),            // 4 Producto
        txt(r[6]),            // 5 Material
        txt(r[7]),            // 6 Componente (pieza)
        cantidad,             // 7 Cantidad
        txt(r[9]),            // 8 Color
        num(r[13]),           // 9 Peso (g)
        Math.round((num(r[20]) || 0) * 100) / 100, // 10 Kg Material Total
        comps                 // 11 [[componente, kg], ...]
      ]);
    });
  }

  // --- Consumo por Material (fila 3 en adelante, A..K) ---
  var materiales = [];
  var shM = ss.getSheetByName('Consumo por Material');
  if (shM && shM.getLastRow() >= 3) {
    shM.getRange(3, 1, shM.getLastRow() - 2, 11).getValues().forEach(function (r) {
      var nombre = txt(r[0]);
      if (!nombre || nombre === 'TOTAL') return;
      materiales.push({
        codigo: nombre, kg: num(r[1]) || 0, abc: txt(r[5]),
        existencia: num(r[6]), existenciaTxt: txt(r[6]),
        promedio: num(r[7]), dias: num(r[8]),
        nombre: txt(r[9]) || nombre, familia: txt(r[10])
      });
    });
  }

  // --- Imágenes (opcional): Clave | URL ---
  var imagenes = {};
  var shI = ss.getSheetByName('Imágenes');
  if (shI && shI.getLastRow() >= 2) {
    shI.getRange(2, 1, shI.getLastRow() - 1, 2).getValues().forEach(function (r) {
      var clave = txt(r[0]).toUpperCase();
      var url = urlImagen_(txt(r[1]));
      if (clave && url) imagenes[clave] = url;
    });
  }

  return { recetas: recetas, ordenes: ordenes, materiales: materiales, imagenes: imagenes,
           url: ScriptApp.getService().getUrl() };
}

// Convierte un link de Google Drive en un link de imagen que el navegador puede mostrar.
// Cualquier otro link (https://...) se deja igual.
function urlImagen_(url) {
  if (!/^https?:\/\//i.test(url)) return '';
  var m = url.match(/drive\.google\.com\/.*?(?:\/d\/|[?&]id=)([\w-]{20,})/);
  return m ? 'https://drive.google.com/thumbnail?id=' + m[1] + '&sz=w1200' : url;
}

/**
 * Ejecútala UNA vez desde el editor (▶ Ejecutar) para crear la pestaña "Imágenes"
 * con todas las claves ya escritas: cada producto, cada receta y cada materia prima.
 * Tú solo pegas el link de la imagen en la columna B. Si la vuelves a correr,
 * solo agrega las claves nuevas; nunca borra lo que ya pegaste.
 */
function crearPestanaImagenes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Imágenes');
  if (!sh) {
    sh = ss.insertSheet('Imágenes');
    sh.getRange(1, 1, 1, 3).setValues([['Clave', 'URL de la imagen (link de Drive o de internet)', 'Tipo']])
      .setFontWeight('bold').setBackground('#1F4E5F').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 360); sh.setColumnWidth(2, 460); sh.setColumnWidth(3, 110);
  }
  var existentes = {};
  if (sh.getLastRow() >= 2) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().forEach(function (r) {
      existentes[String(r[0]).trim().toUpperCase()] = true;
    });
  }
  var d = datosRecetas();
  var nuevas = [];
  function agregar(clave, tipo) {
    clave = String(clave || '').trim();
    if (!clave || existentes[clave.toUpperCase()]) return;
    existentes[clave.toUpperCase()] = true;
    nuevas.push([clave, '', tipo]);
  }
  d.ordenes.forEach(function (o) { agregar(o[4], 'Producto'); });
  d.recetas.forEach(function (r) { agregar(r.material, 'Receta'); });
  d.materiales.forEach(function (m) { agregar(m.codigo, 'Materia prima'); });
  if (nuevas.length) sh.getRange(sh.getLastRow() + 1, 1, nuevas.length, 3).setValues(nuevas);
  var msg = nuevas.length
    ? 'Listo: se agregaron ' + nuevas.length + ' claves a la pestaña "Imágenes". Pega los links en la columna B.'
    : 'La pestaña "Imágenes" ya tenía todas las claves.';
  // Desde el editor no siempre hay ventana de la hoja para mostrar alertas.
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); }
}
