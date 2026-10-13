/**
 * BOTÓN DE ACTUALIZACIÓN — Seguimiento de Producción Manaplas
 * ---------------------------------------------------------------
 * Qué hace, en orden, cada vez que lo ejecutas:
 *   1. Revisa la pestaña "Seguimiento" fila por fila. Si una fila tiene
 *      Cantidad (columna I) pero le faltan las fórmulas de las columnas
 *      R a AD (Código Color, % Colorante, Kg, Componentes...), se las
 *      copia de la fila 5 (que siempre debe tener fórmulas correctas)
 *      y las ajusta solas a su propia fila.
 *   2. Reconstruye por completo "Resumen Consumo" (colorante por color)
 *      con TODOS los colores que aparezcan hoy en Seguimiento — si hay
 *      uno nuevo que nunca habías usado, se agrega solo.
 *   3. Reconstruye por completo "Consumo por Material" con TODOS los
 *      materiales que aparezcan hoy — igual, si hay uno nuevo se agrega
 *      solo, leyendo directo de las columnas Componente 1-4 / Kg
 *      Componente 1-4 que ya calcula la hoja Seguimiento.
 *
 * NO inventa números: solo copia fórmulas que ya existen y suma lo que
 * la hoja Seguimiento ya calculó. Si un color no está en "Catálogo
 * Colores" seguirá saliendo "(sin mapear)", y si un material compuesto
 * no está en "Recetas de Material" seguirá saliendo "(definir receta)"
 * — exactamente igual que ahora, solo que ya no toca hacerlo a mano.
 */

// ============================================================
// FUNCIÓN PRINCIPAL — esta es la que conecta al botón
// ============================================================
function actualizarTodo() {
  try {
    var reparadas = repararFormulasFaltantes();
    SpreadsheetApp.flush(); // fuerza a que Sheets recalcule antes de leer los totales

    var filasConError = detectarFilasConError();

    actualizarResumenColores();
    actualizarResumenMateriales();
    // Los resúmenes se reordenan en cada corrida, así que el cruce de
    // inventario se vuelve a calcular aquí mismo; si no, la existencia de
    // un material quedaba pegada al renglón de otro.
    actualizarInventarioCore();
    actualizarDashboard();
    SpreadsheetApp.flush();
    actualizarPortada();

    var mensaje = 'Seguimiento actualizado.\n' +
      '- Filas reparadas (les faltaban fórmulas): ' + reparadas + '\n' +
      '- Resumen Consumo (colorante): reconstruido\n' +
      '- Consumo por Material: reconstruido\n' +
      '- Cruce de inventario (Días de Cobertura): recalculado\n' +
      '- Dashboard: reconstruido\n' +
      '- Portada (Inicio): actualizada';

    if (filasConError.length > 0) {
      mensaje += '\n\n⚠️ OJO: estas filas de "Seguimiento" tienen un dato faltante o mal ' +
        'escrito (por eso no calculan bien, seguramente ahí ves el #VALUE!): fila(s) ' +
        filasConError.join(', ') + '. Revisa que Cantidad y Peso sean números en esas filas.';
    }

    SpreadsheetApp.getUi().alert(mensaje);
  } catch (e) {
    SpreadsheetApp.getUi().alert(
      'Se interrumpió con un error. Copia este texto completo y compártelo:\n\n' + e.message
    );
  }
}

// Revisa si alguna fila de Seguimiento con datos tiene #VALUE!/#N/A/etc. en sus columnas calculadas
function detectarFilasConError() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Seguimiento');
  var lastRow = sh.getLastRow();
  var filas = [];
  if (lastRow < 5) return filas;

  var n = lastRow - 4;
  var cantidades = sh.getRange(5, 9, n, 1).getValues();   // I
  var kgColorante = sh.getRange(5, 20, n, 1).getValues(); // T

  for (var i = 0; i < n; i++) {
    var cantidad = cantidades[i][0];
    if (cantidad === '' || cantidad === null) continue;
    var t = kgColorante[i][0];
    if (typeof t === 'string' && t.indexOf('#') === 0) filas.push(5 + i);
  }
  return filas;
}

// ============================================================
// 1) Repara filas con datos pero sin fórmulas calculadas
// ============================================================
function repararFormulasFaltantes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Seguimiento');
  var lastRow = sh.getLastRow();
  var startRow = 5;
  var templateRange = sh.getRange(5, 18, 1, 13); // columnas R a AD (18 a 30)
  var reparadas = 0;
  if (lastRow < startRow) return 0;

  var n = lastRow - startRow + 1;
  var cantidades = sh.getRange(startRow, 9, n, 1).getValues();    // I
  var formulasR = sh.getRange(startRow, 18, n, 1).getFormulas();  // R

  for (var i = 0; i < n; i++) {
    var r = startRow + i;
    var cantidad = cantidades[i][0];
    var formulaCodigoColor = formulasR[i][0];
    if (cantidad !== '' && cantidad !== null && formulaCodigoColor === '') {
      templateRange.copyTo(sh.getRange(r, 18, 1, 13));
      reparadas++;
    }
  }
  return reparadas;
}

// ============================================================
// 2) Reconstruye "Resumen Consumo" (colorante por color) desde cero
// ============================================================
function actualizarResumenColores() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Seguimiento');
  var resumen = ss.getSheetByName('Resumen Consumo');
  var lastRow = sh.getLastRow();

  var datos = {}; // codigo -> {kg, n, nombre}
  if (lastRow >= 5) {
    var n = lastRow - 4;
    var codigosCol = sh.getRange(5, 18, n, 1).getValues(); // R
    var nombresCol = sh.getRange(5, 10, n, 1).getValues(); // J
    var kgCol = sh.getRange(5, 20, n, 1).getValues();      // T
    for (var i = 0; i < n; i++) {
      var cod = codigosCol[i][0], nombre = nombresCol[i][0], kg = kgCol[i][0];
      if (cod && cod !== '' && typeof kg === 'number') {
        if (!datos[cod]) datos[cod] = { kg: 0, n: 0, nombre: nombre };
        datos[cod].kg += kg;
        datos[cod].n += 1;
      }
    }
  }

  var codigos = Object.keys(datos).sort(function (a, b) { return datos[b].kg - datos[a].kg; });
  var totalKg = codigos.reduce(function (s, c) { return s + datos[c].kg; }, 0);
  var totalN = codigos.reduce(function (s, c) { return s + datos[c].n; }, 0);

  // limpiar filas anteriores (deja margen amplio). Se limpian también H a K
  // (cruce de inventario) porque el orden de los colores cambia en cada
  // corrida; actualizarInventarioCore() las vuelve a llenar en el orden nuevo.
  resumen.getRange(3, 1, 300, 11).clearContent();
  resumen.getRange(3, 1, 300, 7).setBackground(null);
  // columna A como texto: si no, "0855" se guarda como el número 855
  resumen.getRange(3, 1, 300, 1).setNumberFormat('@');

  var r = 3, acumulado = 0;
  var ABC_COLOR = { A: '#D9EAD3', B: '#FFF2CC', C: '#F4F4F4' };
  codigos.forEach(function (cod) {
    var kg = datos[cod].kg, n = datos[cod].n;
    var pct = totalKg > 0 ? kg / totalKg : 0;
    acumulado += pct;
    var clase = acumulado <= 0.8 ? 'A' : (acumulado <= 0.95 ? 'B' : 'C');
    resumen.getRange(r, 1, 1, 7).setValues([[cod, datos[cod].nombre, kg, n, pct, acumulado, clase]]);
    if ((r - 3) % 2 === 1) resumen.getRange(r, 1, 1, 7).setBackground('#F8F9FA');
    resumen.getRange(r, 7).setBackground(ABC_COLOR[clase]).setFontWeight('bold')
      .setHorizontalAlignment('center');
    if (cod === '(sin mapear)') resumen.getRange(r, 1, 1, 7).setBackground('#FFC7CE');
    r++;
  });
  resumen.getRange(r, 2, 1, 3).setValues([['TOTAL', totalKg, totalN]]);
  resumen.getRange(r, 1, 1, 7).setFontWeight('bold');

  if (codigos.length > 0) {
    resumen.getRange(3, 3, codigos.length, 1).setNumberFormat('#,##0.00');
    resumen.getRange(3, 5, codigos.length, 2).setNumberFormat('0.0%');
    // Formato condicional en "Días de Cobertura" (col K, la llena el botón de
    // inventario): rojo si quedan menos de 7 días, verde si hay margen sano.
    var reglaRoja = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(7)
      .setBackground('#F4C7C3').setFontColor('#990000')
      .setRanges([resumen.getRange(3, 11, codigos.length, 1)])
      .build();
    var reglaVerde = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(7)
      .setBackground('#D9EAD3').setFontColor('#274E13')
      .setRanges([resumen.getRange(3, 11, codigos.length, 1)])
      .build();
    resumen.setConditionalFormatRules([reglaRoja, reglaVerde]);
  }

  resumen.getRange(1, 9).setValue(
    'Última actualización: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')
  );

  // --- Grafica de barras: top 10 colores por kg consumido ---
  var topN = Math.min(10, codigos.length);
  if (topN > 0) {
    resumen.getCharts().forEach(function (ch) { resumen.removeChart(ch); });
    var rangoNombres = resumen.getRange(3, 2, topN, 1); // Color (nombre)
    var rangoKg = resumen.getRange(3, 3, topN, 1);      // Kg Consumidos
    var chart = resumen.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(rangoNombres)
      .addRange(rangoKg)
      .setPosition(3, 13, 0, 0)
      .setOption('title', 'Top ' + topN + ' colores por kg consumido')
      .setOption('width', 620)
      .setOption('height', 320)
      .setOption('legend', { position: 'none' })
      .setOption('colors', ['#B45F06'])
      .build();
    resumen.insertChart(chart);
  }
}

// ============================================================
// 3) Reconstruye "Consumo por Material" desde cero
// ============================================================
// Quita un porcentaje pegado al inicio del nombre de un componente, ej.
// "10%0506C" -> "0506C", "50% 40H92C" -> "40H92C". El % en sí no se pierde
// (sigue guardado en su propia columna en "Recetas de Material"), esto solo
// limpia el NOMBRE para que el mismo material se sume junto sin importar de
// qué receta vino.
function limpiarNombreComponente(nombre) {
  return String(nombre).replace(/^\s*\d+(\.\d+)?%\s*/, '').trim();
}

function actualizarResumenMateriales() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Seguimiento');
  var resumen = ss.getSheetByName('Consumo por Material');
  var lastRow = sh.getLastRow();

  var datos = {}; // nombre material -> {kg, n}

  if (lastRow >= 5) {
    var n = lastRow - 4;
    var bloque = sh.getRange(5, 23, n, 8).getValues(); // W..AD de un jalón (8 columnas)
    var pares = [[0, 1], [2, 3], [4, 5], [6, 7]]; // offsets: W/X, Y/Z, AA/AB, AC/AD
    for (var i = 0; i < n; i++) {
      for (var p = 0; p < pares.length; p++) {
        var nombre = bloque[i][pares[p][0]];
        var kg = bloque[i][pares[p][1]];
        if (nombre && nombre !== '' && typeof kg === 'number' && kg > 0) {
          // quita un "%" pegado al inicio del nombre (ej. "10%0506C" -> "0506C")
          // para que se sume junto con el mismo material aunque venga de una
          // receta distinta con otro %.
          var nombreLimpio = limpiarNombreComponente(nombre);
          // Si el texto trae un código interno conocido adentro (ej. "2213 ( LLB5000 )",
          // "20400 excell", "2811  (HD2000-GP)"), se agrupa bajo ESE código en vez de
          // quedarse como su propio renglón suelto por la descripción/marca pegada.
          // Mayúsculas siempre, para que "Molido" y "MOLIDO" (u otra
          // diferencia de mayúscula/minúscula entre recetas) se sumen
          // juntos sin importar cómo se haya escrito cada vez.
          var claveAgrupacion = (detectarCodigoInterno(nombreLimpio) || nombreLimpio).toUpperCase();
          if (!datos[claveAgrupacion]) datos[claveAgrupacion] = { kg: 0, n: 0 };
          datos[claveAgrupacion].kg += kg;
          datos[claveAgrupacion].n += 1;
        }
      }
    }
  }

  var materiales = Object.keys(datos).sort(function (a, b) { return datos[b].kg - datos[a].kg; });
  var totalKg = materiales.reduce(function (s, m) { return s + datos[m].kg; }, 0);

  // A a I: también se limpia el cruce de inventario (G, H, I) porque el orden
  // de los materiales cambia en cada corrida; actualizarInventarioCore() lo
  // vuelve a llenar en el orden nuevo.
  resumen.getRange(3, 1, 300, 9).clearContent();
  resumen.getRange(3, 1, 300, 6).setBackground(null);
  resumen.getRange(3, 10, 300, 2).clearContent();
  resumen.getRange(2, 10, 1, 2).setValues([['Nombre Descriptivo', 'Familia']])
    .setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#1F4E5F')
    .setHorizontalAlignment('center');

  var FAMILIA_COLOR = {
    PPH: '#D9EAD3', PPC: '#FFF2CC', PEAD: '#CFE2F3', PEBD: '#D0E0E3',
    Aditivo: '#EAD1DC', Reciclado: '#E6E6E6', PS: '#F4CCCC'
  };

  var r = 3, acumulado = 0;
  var ABC_COLOR = { A: '#D9EAD3', B: '#FFF2CC', C: '#F4F4F4' };
  materiales.forEach(function (m) {
    var kg = datos[m].kg, n = datos[m].n;
    var pct = totalKg > 0 ? kg / totalKg : 0;
    acumulado += pct;
    var clase = acumulado <= 0.8 ? 'A' : (acumulado <= 0.95 ? 'B' : 'C');
    var info = infoMaterial(m);
    // columna A como texto forzado: si no, un codigo puramente numerico como
    // "0240" o "2908" se convierte solo en numero y pierde el cero adelante
    // (y luego ya no hace match con Inv).
    resumen.getRange(r, 1).setNumberFormat('@');
    resumen.getRange(r, 1, 1, 6).setValues([[m, kg, n, pct, acumulado, clase]]);
    resumen.getRange(r, 10, 1, 2).setValues([[info.nombre, info.familia]]);
    // banda alterna para que se lea mas facil
    if ((r - 3) % 2 === 1) resumen.getRange(r, 1, 1, 6).setBackground('#F8F9FA');
    resumen.getRange(r, 6).setBackground(ABC_COLOR[clase]).setFontWeight('bold')
      .setHorizontalAlignment('center');
    if (info.familia && FAMILIA_COLOR[info.familia]) {
      resumen.getRange(r, 11).setBackground(FAMILIA_COLOR[info.familia]).setFontWeight('bold')
        .setHorizontalAlignment('center');
    }
    if (m === '(definir receta)') resumen.getRange(r, 1, 1, 6).setBackground('#FFC7CE');
    r++;
  });
  resumen.getRange(r, 1, 1, 2).setValues([['TOTAL', totalKg]]);
  resumen.getRange(r, 1, 1, 6).setFontWeight('bold').setBorder(true, false, false, false, false, false);

  if (materiales.length > 0) {
    resumen.getRange(3, 2, materiales.length, 1).setNumberFormat('#,##0.00');
    resumen.getRange(3, 4, materiales.length, 2).setNumberFormat('0.0%');
    // Formato condicional en "Días de Cobertura" (col I, la llena el botón de
    // inventario): rojo si quedan menos de 7 días, verde si hay margen sano.
    var reglaRoja = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(7)
      .setBackground('#F4C7C3').setFontColor('#990000')
      .setRanges([resumen.getRange(3, 9, materiales.length, 1)])
      .build();
    var reglaVerde = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(7)
      .setBackground('#D9EAD3').setFontColor('#274E13')
      .setRanges([resumen.getRange(3, 9, materiales.length, 1)])
      .build();
    resumen.setConditionalFormatRules([reglaRoja, reglaVerde]);
  }

  resumen.getRange(1, 10).setValue(
    'Última actualización: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')
  );

  // --- Grafica de barras: top 10 materiales por kg consumido ---
  var topN = Math.min(10, materiales.length);
  if (topN > 0) {
    resumen.getCharts().forEach(function (ch) { resumen.removeChart(ch); });
    var rangoGrafica = resumen.getRange(3, 1, topN, 2); // Material + Kg Consumidos
    var chart = resumen.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(rangoGrafica)
      .setPosition(3, 12, 0, 0)
      .setOption('title', 'Top ' + topN + ' materiales por kg consumido')
      .setOption('width', 620)
      .setOption('height', 320)
      .setOption('legend', { position: 'none' })
      .setOption('colors', ['#1F4E5F'])
      .build();
    resumen.insertChart(chart);
  }
}

// ============================================================
// 4) Reconstruye el "Dashboard" (KPIs + gráficas) desde cero, leyendo
//    directo de Seguimiento. Si la pestaña no existe, no hace nada.
// ============================================================
// Calcula todos los agregados del Dashboard leyendo Seguimiento en vivo.
// La usan tanto actualizarDashboard() (escribe en la pestaña) como doGet()
// (la pagina web bonita) — asi los dos siempre muestran exactamente lo mismo.
function calcularDatosDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Seguimiento');

  var lastRow = sh.getLastRow();
  var porFecha = {};       // 'yyyy-MM-dd' -> kg total
  var porProducto = {};    // producto -> {cantidad, kgMp}
  var porColor = {};       // color -> kg colorante
  var totalOrdenes = 0, totalKgColor = 0, totalKgMp = 0;
  var productosDistintos = {};

  if (lastRow >= 5) {
    var n = lastRow - 4;
    var bloque = sh.getRange(5, 1, n, 22).getValues(); // A..V (hasta Kg Materia Prima)
    for (var i = 0; i < n; i++) {
      var fila = bloque[i];
      var cantidad = fila[8]; // I
      if (cantidad === '' || cantidad === null) continue;
      var fecha = fila[0];    // A
      var producto = fila[5]; // F
      var color = fila[9];    // J
      var kgColor = fila[19]; // T
      var kgMp = fila[21];    // V
      kgColor = typeof kgColor === 'number' ? kgColor : 0;
      kgMp = typeof kgMp === 'number' ? kgMp : 0;

      totalOrdenes++;
      totalKgColor += kgColor;
      totalKgMp += kgMp;
      if (producto) productosDistintos[producto] = true;

      var fKey = (Object.prototype.toString.call(fecha) === '[object Date]')
        ? Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'yyyy-MM-dd')
        : String(fecha || '').trim();
      if (fKey) porFecha[fKey] = (porFecha[fKey] || 0) + kgColor + kgMp;

      if (producto) {
        if (!porProducto[producto]) porProducto[producto] = { cantidad: 0, kgMp: 0 };
        porProducto[producto].cantidad += cantidad;
        porProducto[producto].kgMp += kgMp;
      }
      if (color) porColor[color] = (porColor[color] || 0) + kgColor;
    }
  }

  var diasDistintos = Object.keys(porFecha).length;
  var productosCount = Object.keys(productosDistintos).length;
  var fechasOrdenadas = Object.keys(porFecha).sort();
  var topProductos = Object.keys(porProducto)
    .map(function (p) { return [p, porProducto[p].cantidad]; })
    .sort(function (a, b) { return b[1] - a[1]; })
    .slice(0, 8);
  var topColores = Object.keys(porColor)
    .map(function (c) { return [c, Math.round(porColor[c] * 100) / 100]; })
    .sort(function (a, b) { return b[1] - a[1]; })
    .slice(0, 8);

  return {
    totalOrdenes: totalOrdenes, diasDistintos: diasDistintos, productosCount: productosCount,
    totalKgMp: totalKgMp, totalKgColor: totalKgColor,
    fechas: fechasOrdenadas, kgPorFecha: fechasOrdenadas.map(function (f) { return Math.round(porFecha[f] * 100) / 100; }),
    topProductos: topProductos, topColores: topColores
  };
}

// ============================================================
// Aplicación web: una página HTML bonita con el Dashboard en vivo.
// Para publicarla: Implementar > Nueva implementación > Aplicación web >
// Ejecutar como "Yo", Quién puede acceder "Cualquier persona con la cuenta
// de Google" (o "Cualquier persona" si la quieres compartir fuera de tu
// organización). Te da un link que puedes abrir o compartir — cada vez que
// alguien lo abre, se calcula fresco desde Seguimiento, no es una foto vieja.
// ============================================================
// Lee "Consumo por Material" y "Resumen Consumo" tal cual quedaron la última
// vez que corriste "📦 Cruzar inventario (Inv)" — no recalcula el cruce,
// solo lo muestra. Si quieres ver números más frescos de inventario, corre
// ese botón en la hoja antes de abrir la página web.
function leerInventarioParaDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var consMat = ss.getSheetByName('Consumo por Material');
  var resColor = ss.getSheetByName('Resumen Consumo');

  var materiales = [];
  if (consMat) {
    var lastRowMat = consMat.getLastRow();
    if (lastRowMat >= 3) {
      var datosMat = consMat.getRange(3, 1, lastRowMat - 2, 11).getValues(); // A..K
      datosMat.forEach(function (r) {
        var nombre = String(r[0] || '').trim();
        if (!nombre || nombre === 'TOTAL') return;
        materiales.push({
          nombre: nombre, kgConsumido: typeof r[1] === 'number' ? r[1] : 0,
          existenciaActual: r[6], promedioDiario: r[7], diasCobertura: r[8],
          nombreDescriptivo: r[9] || nombre, familia: r[10] || ''
        });
      });
    }
  }

  var colores = [];
  if (resColor) {
    var lastRowCol = resColor.getLastRow();
    if (lastRowCol >= 3) {
      var datosCol = resColor.getRange(3, 1, lastRowCol - 2, 11).getValues(); // A..K
      datosCol.forEach(function (r) {
        var codigo = String(r[0] || '').trim();
        if (!codigo || r[1] === 'TOTAL') return;
        colores.push({
          codigo: codigo, nombre: r[1], kgConsumido: typeof r[2] === 'number' ? r[2] : 0,
          existencia: r[7], promedioDiario: r[8], diasCobertura: r[10]
        });
      });
    }
  }

  // Lo más urgente primero: menos días de cobertura arriba. Lo que no tiene
  // match/número (texto tipo "sin match") se manda al final, no estorba.
  function porUrgencia(a, b) {
    var da = typeof a.diasCobertura === 'number' ? a.diasCobertura : Infinity;
    var db = typeof b.diasCobertura === 'number' ? b.diasCobertura : Infinity;
    return da - db;
  }
  materiales.sort(porUrgencia);
  colores.sort(porUrgencia);

  var alertaMateriales = materiales.filter(function (m) { return typeof m.diasCobertura === 'number' && m.diasCobertura < 7; }).length;
  var alertaColores = colores.filter(function (c) { return typeof c.diasCobertura === 'number' && c.diasCobertura < 7; }).length;

  return {
    materiales: materiales, colores: colores,
    alertas: alertaMateriales + alertaColores
  };
}

function doGet(e) {
  var d = calcularDatosDashboard();
  var inv = leerInventarioParaDashboard();
  var ahora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');

  var html = HtmlService.createTemplateFromFile('DashboardHTML');
  html.datos = d;
  html.inventario = inv;
  html.ahora = ahora;
  return html.evaluate()
    .setTitle('Seguimiento Manaplas — Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function actualizarDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ws = ss.getSheetByName('Dashboard');
  var sh = ss.getSheetByName('Seguimiento');
  if (!ws || !sh) return;
  var d = calcularDatosDashboard();

  // --- KPIs (fila 6, columnas A,C,E,G,I,K en pares) ---
  // Se escriben como números, no como texto con toLocaleString(): el servidor
  // de Google formatea "49503" como "49.503" y la hoja (en_US) lo leía como
  // 49,503 kg -> el Dashboard mostraba "50" en vez de 49,503.
  var kpiValues = [
    d.totalOrdenes,
    d.diasDistintos,
    d.productosCount,
    Math.round(d.totalKgMp),
    Math.round(d.totalKgColor),
    Math.round(d.totalKgMp + d.totalKgColor)
  ];
  for (var k = 0; k < 6; k++) {
    ws.getRange(6, 1 + k * 2).setNumberFormat('#,##0').setValue(kpiValues[k]);
  }

  // --- Tablas de apoyo para las 3 graficas (misma posicion que el diseno original) ---
  var DATA_START = 32;
  var fechasOrdenadas = d.fechas;
  var topProductos = d.topProductos; // [nombre, cantidad]
  var topColores = d.topColores;     // [nombre, kg]

  // limpiar el area de datos vieja (suficiente margen para hasta 31 dias)
  ws.getRange(DATA_START, 1, 32, 2).clearContent();
  ws.getRange(DATA_START, 5, 32, 2).clearContent();
  ws.getRange(DATA_START, 9, 32, 2).clearContent();

  fechasOrdenadas.forEach(function (f, idx) {
    ws.getRange(DATA_START + 1 + idx, 1).setValue(f);
    ws.getRange(DATA_START + 1 + idx, 2).setValue(d.kgPorFecha[idx]);
  });
  var lastFechaRow = DATA_START + fechasOrdenadas.length;

  topProductos.forEach(function (row, idx) {
    var nombreCorto = row[0].length > 28 ? row[0].substring(0, 26) + '…' : row[0];
    ws.getRange(DATA_START + 1 + idx, 5).setValue(nombreCorto);
    ws.getRange(DATA_START + 1 + idx, 6).setValue(row[1]);
  });
  var lastProdRow = DATA_START + topProductos.length;

  topColores.forEach(function (row, idx) {
    ws.getRange(DATA_START + 1 + idx, 9).setValue(row[0]);
    ws.getRange(DATA_START + 1 + idx, 10).setValue(row[1]);
  });
  var lastColorRow = DATA_START + topColores.length;

  // --- Recrear las 3 graficas con el rango correcto ---
  ws.getCharts().forEach(function (ch) { ws.removeChart(ch); });

  if (fechasOrdenadas.length > 0) {
    var line = ws.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(ws.getRange(DATA_START, 1, lastFechaRow - DATA_START + 1, 2))
      .setPosition(9, 1, 0, 0)
      .setOption('title', 'Consumo total de materia prima por día (kg)')
      .setOption('width', 620).setOption('height', 300)
      .setOption('legend', { position: 'none' })
      .setOption('colors', ['#0092CF'])
      .build();
    ws.insertChart(line);
  }
  if (topProductos.length > 0) {
    var bar1 = ws.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(ws.getRange(DATA_START, 5, lastProdRow - DATA_START + 1, 2))
      .setPosition(9, 8, 0, 0)
      .setOption('title', 'Top ' + topProductos.length + ' productos por cantidad producida (piezas)')
      .setOption('width', 620).setOption('height', 300)
      .setOption('legend', { position: 'none' })
      .setOption('colors', ['#1F4E5F'])
      .build();
    ws.insertChart(bar1);
  }
  if (topColores.length > 0) {
    var bar2 = ws.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(ws.getRange(DATA_START, 9, lastColorRow - DATA_START + 1, 2))
      .setPosition(26, 1, 0, 0)
      .setOption('title', 'Top ' + topColores.length + ' colores por kg de colorante consumido')
      .setOption('width', 620).setOption('height', 300)
      .setOption('legend', { position: 'none' })
      .setOption('colors', ['#C0533E'])
      .build();
    ws.insertChart(bar2);
  }
}

// ============================================================
// 5) Trae las órdenes nuevas desde la pestaña "Maestro" de cada archivo
//    en la carpeta de Drive configurada abajo (CARPETA_ID), sin duplicar lo que ya existe.
//    La lógica real vive en importarDesdeCarpetaCore() (sin UI), para que
//    la pueda usar tanto el botón del menú como el activador automático.
// ============================================================
function importarDesdeCarpetaCore() {
  var CARPETA_ID = '1O0DMmXaRL6dfLTsXqbA4bdx7aIYm7Wwh'; // ID de la carpeta de Drive donde caen los archivos de órdenes, de cualquier mes

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Seguimiento');

  var lastRow = sh.getLastRow();
  // Para cada orden ya existente guardamos en qué fila está y una "firma"
  // (Máquina..Cavidad) para poder detectar si el dato cambió en el archivo
  // de origen, no solo si la orden ya existe.
  var existentes = {};
  if (lastRow >= 5) {
    var datosActuales = sh.getRange(5, 1, lastRow - 4, 15).getValues(); // A a O
    for (var idxExist = 0; idxExist < datosActuales.length; idxExist++) {
      var rExist = datosActuales[idxExist];
      var fecha = rExist[0], op = rExist[1], sku = rExist[3], color = rExist[9];
      if (fecha && op) {
        var keyExist = formatearFecha(fecha) + '|' + op + '|' + sku + '|' + color;
        existentes[keyExist] = { row: 5 + idxExist, firma: JSON.stringify(rExist.slice(2)) };
      }
    }
  }

  var MESES = {
    'ENERO': 1, 'FEBRERO': 2, 'MARZO': 3, 'ABRIL': 4, 'MAYO': 5, 'JUNIO': 6,
    'JULIO': 7, 'AGOSTO': 8, 'SEPTIEMBRE': 9, 'OCTUBRE': 10, 'NOVIEMBRE': 11, 'DICIEMBRE': 12
  };

  var folder = DriveApp.getFolderById(CARPETA_ID);
  var files = folder.getFiles();

  var nuevasFilas = [];
  var filasActualizar = []; // {row, valores} de órdenes que ya existían pero cambiaron
  var archivosProcesados = 0;
  var archivosOmitidos = [];
  var detallePorArchivo = [];
  var tiempoInicio = new Date().getTime();
  var LIMITE_MS = 4.5 * 60 * 1000; // paramos antes de los 6 minutos que da Google, con margen
  var seCortoPorTiempo = false;

  while (files.hasNext()) {
    if (new Date().getTime() - tiempoInicio > LIMITE_MS) {
      seCortoPorTiempo = true;
      break; // mejor parar a tiempo y guardar lo que ya se hizo, que cortarse a la mitad
    }
    var file = files.next();
    if (file.getMimeType() !== MimeType.GOOGLE_SHEETS) continue;

    var titulo = file.getName();
    var fecha = parsearFecha(titulo, MESES);
    if (!fecha) {
      archivosOmitidos.push(titulo + ' (no se pudo leer la fecha del nombre del archivo)');
      continue;
    }
    var fechaTexto = (fecha.dia < 10 ? '0' + fecha.dia : fecha.dia) + '/' +
      (fecha.mes < 10 ? '0' + fecha.mes : fecha.mes) + '/' + fecha.anio;
    // Lo que se ESCRIBE en la hoja va en formato año-mes-día. Si se escribe
    // "08/10/2026", la hoja (en_US, mes/día) lo guarda como 10 de agosto, y
    // "29/09/2026" queda como texto porque no existe el mes 29.
    // fechaTexto (día/mes) se sigue usando solo para comparar órdenes.
    var fechaISO = fecha.anio + '-' + (fecha.mes < 10 ? '0' + fecha.mes : fecha.mes) + '-' +
      (fecha.dia < 10 ? '0' + fecha.dia : fecha.dia);

    var libroExterno;
    try {
      libroExterno = SpreadsheetApp.openById(file.getId());
    } catch (e) {
      archivosOmitidos.push(titulo + ' (no se pudo abrir: ' + e.message + ')');
      continue;
    }
    var maestro = libroExterno.getSheetByName('Maestro');
    if (!maestro) {
      archivosOmitidos.push(titulo + ' (no tiene pestaña "Maestro")');
      continue;
    }

    var datos;
    try {
      // Solo leemos hasta la fila 250 (30 órdenes de 8 filas cada una) en vez de
      // las 984 filas completas de la plantilla — es de sobra para un día normal
      // y hace la lectura varias veces más rápida.
      var filasALeer = Math.min(250, maestro.getMaxRows());
      datos = maestro.getRange(1, 1, filasALeer, 15).getValues();
    } catch (e) {
      archivosOmitidos.push(titulo + ' (error leyendo Maestro: ' + e.message + ')');
      continue;
    }

    var opActual = null;
    var filasEsteArchivo = 0;
    var filasActualizadasEsteArchivo = 0;
    var filasYaExistianEsteArchivo = 0;
    for (var i = 1; i < datos.length; i++) {
      var fila = datos[i];
      var opCelda = fila[0];
      if (opCelda !== '' && opCelda !== null) opActual = opCelda;

      var cantidad = fila[8];
      if (cantidad === '' || cantidad === null) continue;

      var sku = fila[3], color = fila[9];
      var key = fechaTexto + '|' + opActual + '|' + sku + '|' + color;
      var camposComparables = [fila[1], fila[3], fila[4], fila[5], fila[6], fila[7],
        fila[8], fila[9], fila[10], fila[11], fila[12], fila[13], fila[14]];
      var firmaNueva = JSON.stringify(camposComparables);

      var existente = existentes[key];
      if (existente) {
        if (existente.firma === firmaNueva) {
          filasYaExistianEsteArchivo++; // igual que antes, no hay nada que hacer
        } else {
          // La orden ya existía pero algo cambió en el archivo de origen
          // (peso, cantidad, material...) — actualizamos esa fila en vez de
          // dejarla desactualizada o crear una fila duplicada.
          filasActualizar.push({ row: existente.row, valores: camposComparables });
          existentes[key] = { row: existente.row, firma: firmaNueva };
          filasActualizadasEsteArchivo++;
        }
        continue;
      }

      existentes[key] = { row: null, firma: firmaNueva }; // la fila real se sabe al escribirla
      filasEsteArchivo++;
      nuevasFilas.push([
        fechaISO, opActual, fila[1], fila[3], fila[4], fila[5], fila[6], fila[7],
        fila[8], fila[9], fila[10], fila[11], fila[12], fila[13], fila[14], '', ''
      ]);
    }
    archivosProcesados++;
    var detalleTxt = titulo + ' [' + fechaTexto + ']: ' + filasEsteArchivo + ' nuevas';
    if (filasActualizadasEsteArchivo > 0) detalleTxt += ', ' + filasActualizadasEsteArchivo + ' actualizadas';
    if (filasYaExistianEsteArchivo > 0) detalleTxt += ', ' + filasYaExistianEsteArchivo + ' sin cambios';
    detallePorArchivo.push(detalleTxt);
  }

  // Aplicamos primero las actualizaciones a filas que ya existían (mismo
  // Fecha+OP+Sku+Color, pero cambió algún dato como peso o cantidad).
  if (filasActualizar.length > 0) {
    filasActualizar.forEach(function (u) {
      sh.getRange(u.row, 3, 1, 13).setValues([u.valores]); // C a O
    });
    SpreadsheetApp.flush();
  }

  if (nuevasFilas.length > 0) {
    // 1) Escribimos las filas nuevas al final (posición segura y simple)
    var filaDestino = sh.getLastRow() + 1;
    sh.getRange(filaDestino, 1, nuevasFilas.length, 17).setValues(nuevasFilas);
    sh.getRange(filaDestino, 1, nuevasFilas.length, 1).setNumberFormat('dd/mm/yyyy');
    // Copiamos la fórmula de la fila 5 hacia TODAS las filas nuevas en una sola llamada
    // (en vez de una llamada por fila, que era mucho más lento con muchas filas nuevas)
    var plantilla = sh.getRange(5, 18, 1, 13);
    plantilla.copyTo(sh.getRange(filaDestino, 18, nuevasFilas.length, 13));
    SpreadsheetApp.flush();

    // 2) Las movemos al principio (justo después del encabezado, fila 5),
    //    empujando hacia abajo lo que ya tenías. Las fórmulas se ajustan solas.
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
  // Antes la importación no actualizaba "Inicio", y la portada se quedaba
  // con los números en 0 del último "Borrar".
  actualizarPortada();

  var mensaje = 'Importación desde la carpeta de Drive — ' + new Date().toLocaleString() + '\n' +
    '- Archivos revisados: ' + archivosProcesados + '\n' +
    '- Órdenes nuevas agregadas: ' + nuevasFilas.length + '\n' +
    '- Órdenes actualizadas (ya existían, cambió algún dato): ' + filasActualizar.length + '\n\n' +
    'Detalle por archivo:\n' + detallePorArchivo.join('\n');
  if (seCortoPorTiempo) {
    mensaje += '\n\n⏱️ Se detuvo antes de terminar todos los archivos para no pasarse del ' +
      'límite de tiempo de Google. No hay problema — dale clic otra vez y sigue justo donde ' +
      'se quedó, sin duplicar lo que ya trajo.';
  }
  if (archivosOmitidos.length > 0) {
    mensaje += '\n\n⚠️ Archivos que no pude leer (revísalos a mano):\n' + archivosOmitidos.join('\n');
  }
  if (filasConError.length > 0) {
    mensaje += '\n\n⚠️ Filas con dato faltante o mal escrito (revisa Cantidad y Peso): ' +
      filasConError.join(', ');
  }
  return { mensaje: mensaje, nuevas: nuevasFilas.length };
}

// Botón del menú: corre la importación y muestra el resultado en pantalla
function importarDesdeCarpeta() {
  var resultado = importarDesdeCarpetaCore();
  SpreadsheetApp.getUi().alert(resultado.mensaje);
}

// Activador automático: corre lo mismo, pero como no hay pantalla para mostrar
// un cuadro de diálogo, te manda el mismo resumen por correo en su lugar.
function importarDesdeCarpetaAutomatico() {
  var resultado = importarDesdeCarpetaCore();
  var destinatario = Session.getEffectiveUser().getEmail();
  MailApp.sendEmail(
    destinatario,
    'Seguimiento Manaplas — importación automática (' + resultado.nuevas + ' nuevas)',
    resultado.mensaje
  );
}

function formatearFecha(valor) {
  if (Object.prototype.toString.call(valor) === '[object Date]') {
    // Con la zona horaria de la hoja (no la del script): si las dos zonas no
    // coinciden, una fecha a medianoche podía leerse como el día anterior y la
    // orden se volvía a importar como si fuera nueva.
    var zona = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    return Utilities.formatDate(valor, zona, 'dd/MM/yyyy');
  }
  return String(valor).trim();
}

function parsearFecha(titulo, meses) {
  var m1 = titulo.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m1) return { dia: parseInt(m1[1], 10), mes: parseInt(m1[2], 10), anio: parseInt(m1[3], 10) };
  // El "DE" antes del año es opcional, para aceptar tanto "1 DE OCTUBRE DE 2026"
  // como "1 DE OCTUBRE 2026" (con o sin texto extra después, como "Parte 2").
  var m2 = titulo.match(/(\d{1,2})\s+DE\s+([A-ZÁÉÍÓÚÑ]+)\s+(?:DE\s+)?(\d{4})/i);
  if (m2) {
    var mesNombre = m2[2].toUpperCase()
      .replace('Á', 'A').replace('É', 'E').replace('Í', 'I').replace('Ó', 'O').replace('Ú', 'U');
    var mesNum = meses[mesNombre];
    if (mesNum) return { dia: parseInt(m2[1], 10), mes: mesNum, anio: parseInt(m2[3], 10) };
  }
  return null;
}

// ============================================================
// 4) Borra los datos de prueba (columnas A a Q) SIN tocar las fórmulas (R a AD)
// ============================================================
function limpiarDatosSeptiembre() {
  var ui = SpreadsheetApp.getUi();
  var respuesta = ui.alert(
    'Vas a borrar todas las órdenes capturadas',
    'Esto borra las columnas A a Q (Fecha, OP, Máquina, Producto, Material, Componente, Cantidad, ' +
    'Color, Sacos, Ciclo, Peso, Cavidad, Estado, Enviado) desde la fila 5 en adelante.\n\n' +
    'Las fórmulas de las columnas R a AD (Código Color, % Colorante, Kg, Componentes...) NO se tocan, ' +
    'se quedan listas para cuando vuelvas a escribir datos.\n\n' +
    '¿Confirmas que quieres borrar todo?',
    ui.ButtonSet.YES_NO
  );

  if (respuesta !== ui.Button.YES) {
    ui.alert('Cancelado. No se borró nada.');
    return;
  }

  limpiarDatosSeptiembreInterno();

  ui.alert(
    'Listo. Se borraron los datos de las órdenes, los resúmenes, el Dashboard y el cruce de inventario — todo quedó en cero.\n\n' +
    'Las fórmulas de R a AD en Seguimiento siguen ahí (ahora se ven en blanco porque no hay Cantidad que calcular). ' +
    'Cuando vuelvas a cargar órdenes, corre "📦 Cruzar inventario (Inv)" otra vez para refrescar la existencia.'
  );
}

// Lógica real de "Borrar", sin diálogos — así la puede reusar tanto el botón
// 🗑️ Borrar como el 📅 Corte del mes (que borra después de archivar).
function limpiarDatosSeptiembreInterno() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Seguimiento');
  var lastRow = sh.getLastRow();

  if (lastRow >= 5) {
    sh.getRange(5, 1, lastRow - 4, 17).clearContent(); // columnas A(1) a Q(17)
  }

  SpreadsheetApp.flush();
  actualizarResumenColores();
  actualizarResumenMateriales();
  actualizarDashboard();

  // El inventario (Existencia, Consumo Promedio Diario, Días de Cobertura) lo
  // escribe aparte el botón "Cruzar inventario" — actualizarResumenColores/
  // Materiales no lo toca, así que sin esto se quedaba con números viejos
  // "huérfanos" después de borrar. Lo limpiamos aquí también para que un
  // borrado de verdad deje TODO en cero.
  var consMat = ss.getSheetByName('Consumo por Material');
  var resColor = ss.getSheetByName('Resumen Consumo');
  if (consMat) consMat.getRange(3, 7, 300, 3).clearContent();   // G,H,I
  if (resColor) resColor.getRange(3, 8, 300, 4).clearContent(); // H,I,J,K
  SpreadsheetApp.flush();
  actualizarPortada();
  SpreadsheetApp.flush();
}

// ============================================================
// Devuelve el Google Sheet APARTE donde se guardan los cortes ("Seguimiento
// Manaplas — Histórico"), creándolo la primera vez que hace falta. El ID
// queda guardado en las Propiedades del Script, así que los cortes
// siguientes reutilizan el mismo archivo en vez de crear uno nuevo cada vez.
// ============================================================
function obtenerArchivoHistorico() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('HISTORICO_ID');

  if (id) {
    try {
      return SpreadsheetApp.openById(id);
    } catch (e) {
      // El archivo guardado ya no se puede abrir (se borró, se movió a la papelera, etc.)
      // — seguimos abajo y se crea uno nuevo.
    }
  }

  var nuevo = SpreadsheetApp.create('Seguimiento Manaplas — Histórico');
  props.setProperty('HISTORICO_ID', nuevo.getId());

  // Lo deja en la misma carpeta de Drive que el archivo principal, para que sea fácil de encontrar
  try {
    var archivoActual = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
    var carpetas = archivoActual.getParents();
    if (carpetas.hasNext()) {
      var carpeta = carpetas.next();
      var archivoNuevo = DriveApp.getFileById(nuevo.getId());
      carpeta.addFile(archivoNuevo);
      DriveApp.getRootFolder().removeFile(archivoNuevo);
    }
  } catch (e) {
    // Si no se pudo mover de carpeta no importa — el archivo igual quedó creado y funciona,
    // solo que aparece en la raíz de Drive en vez de al lado del archivo principal.
  }

  return nuevo;
}

// Borra la pestaña vacía ("Hoja 1", "Sheet1", etc.) que Google Sheets le pone
// sola a todo archivo nuevo, una vez que el histórico ya tiene pestañas reales.
function limpiarPestanaDefaultDeHistorico(historico) {
  if (historico.getNumSheets() < 2) return;
  var hojas = historico.getSheets();
  for (var i = 0; i < hojas.length; i++) {
    var nombre = hojas[i].getName();
    var vacia = hojas[i].getLastRow() === 0 && hojas[i].getLastColumn() === 0;
    if (vacia && /^(Hoja \d+|Sheet\d+)$/.test(nombre)) {
      historico.deleteSheet(hojas[i]);
    }
  }
}

// ============================================================
// 6b) CORTE DEL MES — guarda una copia congelada de "Seguimiento",
//     "Consumo por Material", "Resumen Consumo" y "Dashboard" TAL COMO
//     ESTÁN en este momento (con los números ya calculados) en el archivo
//     APARTE "Seguimiento Manaplas — Histórico", y después deja la hoja de
//     trabajo en cero para el mes que sigue. El historial archivado nunca
//     se borra ni se vuelve a tocar.
// ============================================================
function corteDelMes() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shSept = ss.getSheetByName('Seguimiento');
  var lastRow = shSept.getLastRow();

  if (lastRow < 5) {
    ui.alert('No hay nada que cerrar. La hoja "Seguimiento" está vacía, todavía no hay órdenes capturadas.');
    return;
  }

  var fechaCorte = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MM-yyyy');
  var etiqueta = 'Cierre ' + fechaCorte;
  var historico = obtenerArchivoHistorico();

  var confirmar = ui.alert(
    '📅 Corte del mes — ' + etiqueta,
    'Esto va a hacer DOS cosas:\n\n' +
    '1) GUARDAR una copia de "Seguimiento", "Consumo por Material", "Resumen Consumo" y "Dashboard" ' +
    'tal como están AHORA MISMO (con todos los números ya calculados), con el nombre "' + etiqueta + '", ' +
    'en el archivo APARTE "' + historico.getName() + '" (no en este archivo). ' +
    'Esa copia queda guardada para siempre y no se vuelve a tocar ni a recalcular.\n\n' +
    '2) BORRAR la hoja de trabajo "Seguimiento" (igual que el botón 🗑️ Borrar), para que quede lista ' +
    'para capturar las órdenes que entren de ahora en adelante — aunque ya no sea septiembre, seguimos ' +
    'usando la misma hoja y los mismos botones.\n\n' +
    '¿Confirmas que quieres hacer el corte?',
    ui.ButtonSet.YES_NO
  );

  if (confirmar !== ui.Button.YES) {
    ui.alert('Cancelado. No se hizo ningún corte.');
    return;
  }

  // Se calcula ANTES de tocar nada, mientras "Seguimiento" todavía tiene los
  // datos de este periodo, para poder registrar los totales en el Histórico.
  var datosResumen = calcularDatosDashboard();

  try {
    var nombresOrigen = ['Seguimiento', 'Consumo por Material', 'Resumen Consumo', 'Dashboard'];
    var nombresArchivo = [];
    var nombreSeptArchivado = '';
    var temporales = []; // copias intermedias dentro de este archivo, se borran al final

    nombresOrigen.forEach(function (nombre) {
      var origen = ss.getSheetByName(nombre);
      if (!origen) return;

      // 1) Copia temporal DENTRO de este mismo archivo. Mientras está aquí,
      //    sus fórmulas todavía pueden mirar "Recetas de Material", "Catálogo
      //    Colores", etc. y calcular bien.
      var copiaTemp = origen.copyTo(ss);
      var rangoTemp = copiaTemp.getDataRange();
      if (rangoTemp.getNumRows() > 0 && rangoTemp.getNumColumns() > 0) {
        rangoTemp.setValues(rangoTemp.getValues());
      }
      temporales.push(copiaTemp);

      // 2) Esa copia YA CONGELADA (puros valores, sin fórmulas que se puedan
      //    romper al cambiar de archivo) se manda al archivo histórico aparte.
      var nombreNuevo = etiqueta + ' — ' + nombre;
      var base = nombreNuevo, n = 2;
      while (historico.getSheetByName(nombreNuevo)) {
        nombreNuevo = base + ' (' + n + ')';
        n++;
      }
      var copiaFinal = copiaTemp.copyTo(historico);
      copiaFinal.setName(nombreNuevo);
      nombresArchivo.push(nombreNuevo);
      if (nombre === 'Seguimiento') nombreSeptArchivado = nombreNuevo;
    });

    // Borra las copias temporales que quedaron en ESTE archivo — ya cumplieron su función
    temporales.forEach(function (hoja) { ss.deleteSheet(hoja); });

    // Quita la pestaña vacía por defecto del histórico, si todavía está
    limpiarPestanaDefaultDeHistorico(historico);

    // Deja registro en la pestaña "Histórico de Cortes" de ESTE archivo (la crea si no existe)
    registrarEnHistoricoCortes(ss, historico, etiqueta, nombreSeptArchivado, datosResumen);

    SpreadsheetApp.flush();

    // Deja la hoja de trabajo en cero, lista para lo que sigue
    limpiarDatosSeptiembreInterno();

    ss.setActiveSheet(shSept);

    ui.alert(
      '✅ Corte completo',
      'Quedó guardado en el archivo "' + historico.getName() + '" con estas pestañas nuevas:\n\n' +
      nombresArchivo.join('\n') + '\n\n' +
      'En "Histórico de Cortes" (aquí en este archivo) se agregó una fila con el resumen y un link ' +
      'directo a esa copia archivada.\n\n' +
      'La hoja "Seguimiento" (y sus resúmenes y el Dashboard) ya quedaron en cero, listos para las órdenes ' +
      'que vayan entrando de ahora en adelante — sigue usando el mismo botón "📥 Traer órdenes de la carpeta" ' +
      'como siempre, sea octubre o el mes que sea.',
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert('Se interrumpió con un error durante el corte. Copia este texto y compártelo:\n\n' + e.message);
  }
}

// Mantiene una pestaña "Histórico de Cortes" con una fila por cada corte que
// se haya hecho, con link directo a la pestaña archivada de ese cierre
// (en el archivo histórico aparte).
function registrarEnHistoricoCortes(ss, historico, etiqueta, nombreSeptArchivado, datosResumen) {
  var NOMBRE = 'Histórico de Cortes';
  var hist = ss.getSheetByName(NOMBRE);

  if (!hist) {
    hist = ss.insertSheet(NOMBRE);
    var encabezados = ['Fecha del corte', 'Nombre del cierre', 'Órdenes', 'Días con producción',
      'Kg Materia Prima', 'Kg Colorante', 'Kg Total', 'Ir al detalle'];
    hist.getRange(1, 1, 1, encabezados.length).setValues([encabezados]);
    hist.getRange(1, 1, 1, encabezados.length)
      .setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#1F4E5F');
    hist.setFrozenRows(1);
    hist.setColumnWidths(1, 1, 110);
    hist.setColumnWidth(2, 150);
    hist.setColumnWidth(8, 160);
  }

  var fila = hist.getLastRow() + 1;
  var linkFormula = '';
  var hojaArchivada = nombreSeptArchivado ? historico.getSheetByName(nombreSeptArchivado) : null;
  if (hojaArchivada) {
    var urlDetalle = 'https://docs.google.com/spreadsheets/d/' + historico.getId() + '/edit#gid=' + hojaArchivada.getSheetId();
    linkFormula = '=HYPERLINK("' + urlDetalle + '", "Ver copia archivada")';
  }

  hist.getRange(fila, 1, 1, 8).setValues([[
    new Date(),
    etiqueta,
    datosResumen.totalOrdenes,
    datosResumen.diasDistintos,
    Math.round(datosResumen.totalKgMp),
    Math.round(datosResumen.totalKgColor),
    Math.round(datosResumen.totalKgMp + datosResumen.totalKgColor),
    linkFormula
  ]]);
  hist.getRange(fila, 1).setNumberFormat('dd/MM/yyyy');

  // Mueve el Histórico justo después de "Seguimiento" para que sea fácil de encontrar
  var idxSept = ss.getSheetByName('Seguimiento').getIndex();
  ss.setActiveSheet(hist);
  ss.moveActiveSheet(idxSept + 1);
}

// ============================================================
// 6) Encuentra materiales compuestos (con "+") que se usan en Seguimiento
//    pero NO están definidos en "Recetas de Material", y los agrega ahí
//    solos, resaltados, listos para que solo falte poner el %.
// ============================================================
function detectarRecetasFaltantes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Seguimiento');
  var recetas = ss.getSheetByName('Recetas de Material');

  var lastRow = sh.getLastRow();
  var materialesUsados = {};
  if (lastRow >= 5) {
    var n = lastRow - 4;
    var materialesCol = sh.getRange(5, 7, n, 1).getValues(); // G
    var cantidadCol = sh.getRange(5, 9, n, 1).getValues();   // I
    for (var i = 0; i < n; i++) {
      var material = materialesCol[i][0], cantidad = cantidadCol[i][0];
      if (cantidad === '' || cantidad === null) continue;
      if (typeof material === 'string' && material.indexOf('+') !== -1) {
        materialesUsados[material] = true;
      }
    }
  }

  var recetasLastRow = recetas.getLastRow();
  var definidos = {};
  if (recetasLastRow >= 3) {
    var col = recetas.getRange(3, 1, recetasLastRow - 2, 1).getValues();
    col.forEach(function (r) { if (r[0]) definidos[r[0]] = true; });
  }

  var faltantes = Object.keys(materialesUsados).filter(function (m) { return !definidos[m]; });

  if (faltantes.length === 0) {
    return {
      mensaje: 'No hay materiales compuestos pendientes de receta.\n' +
        'Todo lo que se usa en Seguimiento ya está definido en "Recetas de Material".',
      nuevos: 0
    };
  }

  var filaLibre = 3;
  while (recetas.getRange(filaLibre, 1).getValue() !== '') filaLibre++;

  faltantes.forEach(function (m) {
    recetas.getRange(filaLibre, 1).setValue(m);
    for (var slot = 1; slot <= 4; slot++) {
      var colNombre = 2 + (slot - 1) * 2; // B, D, F, H
      var colPct = colNombre + 1;         // C, E, G, I
      recetas.getRange(filaLibre, colNombre).setFormula(formulaComponente(filaLibre, slot));
    }
    recetas.getRange(filaLibre, 10).setFormula(
      '=IF($A' + filaLibre + '="","",SUM(C' + filaLibre + ',E' + filaLibre + ',G' + filaLibre + ',I' + filaLibre + '))'
    );
    recetas.getRange(filaLibre, 1, 1, 10).setBackground('#FFC7CE');
    [3, 5, 7, 9].forEach(function (c) { recetas.getRange(filaLibre, c).setBackground('#FFFFCC'); });
    filaLibre++;
  });

  var mensaje = 'Se encontraron ' + faltantes.length + ' material(es) compuesto(s) sin receta:\n\n' +
    faltantes.join('\n') +
    '\n\nYa los agregué en "Recetas de Material" (fila en rojo, celdas de % en amarillo). ' +
    'Solo falta que escribas el % o partes de cada componente — el resto se calcula solo.';

  return { mensaje: mensaje, nuevos: faltantes.length };
}

function detectarRecetasFaltantesBoton() {
  var resultado = detectarRecetasFaltantes();
  SpreadsheetApp.getUi().alert(resultado.mensaje);
}

// Fórmula que extrae el componente N (1 a 4) de un texto tipo "A+B+C", separado
// por "+". Usa SPLIT + INDEX (más simple y confiable) en vez del truco viejo
// de MID/SUBSTITUTE/REPT, que a veces daba #ERROR! sin razón clara. Si el
// material tiene menos partes que "slot", IFERROR devuelve "" sin explotar.
function formulaComponente(fila, slot) {
  return '=IF($A' + fila + '="","",IFERROR(TRIM(INDEX(SPLIT($A' + fila + ',"+"),1,' + slot + ')),""))';
}

// ============================================================
// 6c) REPARAR ERRORES CONOCIDOS — corrige de una vez:
//     1) Las fórmulas viejas de "Recetas de Material" que a veces daban
//        #ERROR! (las reemplaza por la versión SPLIT, más confiable).
//     2) Los links rotos de "Histórico de Cortes" (usaban un formato de
//        URL que a veces Google Sheets no aceptaba desde un script).
//     3) Reconstruye "Inicio" con los links ya corregidos.
//     Es seguro correrlo las veces que quieras.
// ============================================================
function repararErroresConocidos() {
  var ui = SpreadsheetApp.getUi();
  try {
    var localeInfo = repararConfiguracionRegional();
    var plantillaRestaurada = repararPlantillaSeptiembre();
    var reparadasRecetas = repararFormulasRecetas();
    var reparadosLinks = repararLinksHistoricoCortes();
    actualizarPortada();
    ui.alert(
      '✅ Reparación completa.\n\n' +
      (localeInfo.cambiado ? '- Configuración regional: estaba en "' + localeInfo.anterior + '" (por eso salían los #ERROR!), la puse en "en_US".\n' : '') +
      (plantillaRestaurada ? '- Plantilla de fórmulas de "Seguimiento" (fila 5): restaurada, faltaba.\n' : '') +
      '- Fórmulas corregidas en "Recetas de Material": ' + reparadasRecetas + '\n' +
      '- Links corregidos en "Histórico de Cortes": ' + reparadosLinks + '\n' +
      '- "Inicio" reconstruida con los links buenos.\n\n' +
      'Si en algún lado todavía ves #ERROR! o Kg en 0, avísame con una captura.'
    );
  } catch (e) {
    ui.alert('Se interrumpió con un error durante la reparación. Copia este texto y compártelo:\n\n' + e.message);
  }
}

// Pone la hoja en configuración regional "Estados Unidos" (en_US) si no lo
// estaba ya. Con otra configuración regional (por ejemplo España o alguna
// que use ";" en vez de "," para separar los argumentos de las fórmulas),
// TODAS las fórmulas que este código escribe con setFormula() salen
// #ERROR! aunque el texto de la fórmula esté perfecto — este es el
// arreglo real y definitivo para eso, no hace falta tocar nada a mano.
function repararConfiguracionRegional() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var actual = ss.getSpreadsheetLocale();
  if (actual === 'en_US') return { cambiado: false, anterior: actual };
  ss.setSpreadsheetLocale('en_US');
  SpreadsheetApp.flush();
  return { cambiado: true, anterior: actual };
}

// Restaura las fórmulas "plantilla" de la fila 5 de Seguimiento (columnas R a AD)
// si se perdieron — sin esa plantilla, "Reparar fórmulas faltantes" no tiene de
// dónde copiar y todo el Kg de Materia Prima/Colorante sale en 0. Solo actúa si
// de verdad faltan (no pisa nada si ya están bien puestas).
function repararPlantillaSeptiembre() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Seguimiento');
  if (!sh) return false;

  var formulaActual = sh.getRange(5, 18).getFormula(); // R5
  var faltaba = (formulaActual === '');

  if (faltaba) {
    escribirPlantillaFila5(sh);
  }

  // Vuelve a escribir TAL CUAL las fórmulas que ya haya en R:AD (fila 5 en
  // adelante). No cambia nada de contenido, pero fuerza a Sheets a volver a
  // interpretarlas — esto es lo que de verdad arregla los #ERROR! que deja
  // un cambio de configuración regional, no solo en la fila 5 sino en TODAS
  // las filas que ya tenían la fórmula puesta.
  var lastRowTocar = sh.getLastRow();
  if (lastRowTocar >= 5) {
    var rangoTocar = sh.getRange(5, 18, lastRowTocar - 4, 13); // R5:AD<última fila>
    rangoTocar.setFormulas(rangoTocar.getFormulas());
    SpreadsheetApp.flush();
  }

  // Ahora sí, copia esa plantilla ya restaurada/re-interpretada a todas las
  // filas que tengan Cantidad pero les falte la fórmula (esto ya existía,
  // solo lo reutilizamos), y refresca los resúmenes y el Dashboard.
  repararFormulasFaltantes();
  SpreadsheetApp.flush();
  actualizarResumenColores();
  actualizarResumenMateriales();
  actualizarInventarioCore();
  actualizarDashboard();

  return faltaba;
}

// Escribe desde cero las 13 fórmulas "plantilla" de la fila 5 (columnas R a AD)
function escribirPlantillaFila5(sh) {
  var formulas = [
    '=IF(TRIM(J5)="","",IFERROR(INDEX(\'Catálogo Colores\'!$E$2:$E$67,MATCH(TRIM(J5),\'Catálogo Colores\'!$B$2:$B$67,0)),"(sin mapear)"))',
    '=IF(OR(ISNUMBER(SEARCH("molido",G5)),ISNUMBER(SEARCH("molino",G5))),0.015,0.02)',
    '=IFERROR(IF(I5="","",I5*N5/1000*S5),"(revisar dato)")',
    '=IFERROR(IF(I5="","",I5*N5/1000),"(revisar dato)")',
    '=IFERROR(IF(U5="","",U5-T5),"(revisar dato)")',
    '=IF(ISNUMBER(SEARCH("+",G5)),IFERROR(INDEX(\'Recetas de Material\'!$B$3:$B$30,MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0)),"(definir receta)"),G5)',
    '=IF(V5="","",IF((IF(ISNUMBER(SEARCH("+",G5)),IFERROR(INDEX(\'Recetas de Material\'!$C$3:$C$30,MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0)),100),100)+IF(AND(ISNUMBER(SEARCH("+",G5)),ISNUMBER(MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0))),IFERROR(INDEX(\'Recetas de Material\'!$E$3:$E$30,MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0)),0),0)+IF(AND(ISNUMBER(SEARCH("+",G5)),ISNUMBER(MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0))),IFERROR(INDEX(\'Recetas de Material\'!$G$3:$G$30,MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0)),0),0)+IF(AND(ISNUMBER(SEARCH("+",G5)),ISNUMBER(MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0))),IFERROR(INDEX(\'Recetas de Material\'!$I$3:$I$30,MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0)),0),0))=0,0,V5*(IF(ISNUMBER(SEARCH("+",G5)),IFERROR(INDEX(\'Recetas de Material\'!$C$3:$C$30,MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0)),100),100))/(IF(ISNUMBER(SEARCH("+",G5)),IFERROR(INDEX(\'Recetas de Material\'!$C$3:$C$30,MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0)),100),100)+IF(AND(ISNUMBER(SEARCH("+",G5)),ISNUMBER(MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0))),IFERROR(INDEX(\'Recetas de Material\'!$E$3:$E$30,MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0)),0),0)+IF(AND(ISNUMBER(SEARCH("+",G5)),ISNUMBER(MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0))),IFERROR(INDEX(\'Recetas de Material\'!$G$3:$G$30,MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0)),0),0)+IF(AND(ISNUMBER(SEARCH("+",G5)),ISNUMBER(MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0))),IFERROR(INDEX(\'Recetas de Material\'!$I$3:$I$30,MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0)),0),0))))',
    '=IF(AND(ISNUMBER(SEARCH("+",G5)),ISNUMBER(MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0))),IFERROR(INDEX(\'Recetas de Material\'!$D$3:$D$30,MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0)),""),"")',
    '=IF(V5="","",IF((IF(ISNUMBER(SEARCH("+",G5)),IFERROR(INDEX(\'Recetas de Material\'!$C$3:$C$30,MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0)),100),100)+IF(AND(ISNUMBER(SEARCH("+",G5)),ISNUMBER(MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0))),IFERROR(INDEX(\'Recetas de Material\'!$E$3:$E$30,MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0)),0),0)+IF(AND(ISNUMBER(SEARCH("+",G5)),ISNUMBER(MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0))),IFERROR(INDEX(\'Recetas de Material\'!$G$3:$G$30,MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0)),0),0)+IF(AND(ISNUMBER(SEARCH("+",G5)),ISNUMBER(MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0))),IFERROR(INDEX(\'Recetas de Material\'!$I$3:$I$30,MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0)),0),0))=0,0,V5*(IF(AND(ISNUMBER(SEARCH("+",G5)),ISNUMBER(MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0))),IFERROR(INDEX(\'Recetas de Material\'!$E$3:$E$30,MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0)),0),0))/(IF(ISNUMBER(SEARCH("+",G5)),IFERROR(INDEX(\'Recetas de Material\'!$C$3:$C$30,MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0)),100),100)+IF(AND(ISNUMBER(SEARCH("+",G5)),ISNUMBER(MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0))),IFERROR(INDEX(\'Recetas de Material\'!$E$3:$E$30,MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0)),0),0)+IF(AND(ISNUMBER(SEARCH("+",G5)),ISNUMBER(MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0))),IFERROR(INDEX(\'Recetas de Material\'!$G$3:$G$30,MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0)),0),0)+IF(AND(ISNUMBER(SEARCH("+",G5)),ISNUMBER(MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0))),IFERROR(INDEX(\'Recetas de Material\'!$I$3:$I$30,MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0)),0),0))))',
    '=IF(AND(ISNUMBER(SEARCH("+",G5)),ISNUMBER(MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0))),IFERROR(INDEX(\'Recetas de Material\'!$F$3:$F$30,MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0)),""),"")',
    '=IF(V5="","",IF((IF(ISNUMBER(SEARCH("+",G5)),IFERROR(INDEX(\'Recetas de Material\'!$C$3:$C$30,MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0)),100),100)+IF(AND(ISNUMBER(SEARCH("+",G5)),ISNUMBER(MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0))),IFERROR(INDEX(\'Recetas de Material\'!$E$3:$E$30,MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0)),0),0)+IF(AND(ISNUMBER(SEARCH("+",G5)),ISNUMBER(MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0))),IFERROR(INDEX(\'Recetas de Material\'!$G$3:$G$30,MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0)),0),0)+IF(AND(ISNUMBER(SEARCH("+",G5)),ISNUMBER(MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0))),IFERROR(INDEX(\'Recetas de Material\'!$I$3:$I$30,MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0)),0),0))=0,0,V5*(IF(AND(ISNUMBER(SEARCH("+",G5)),ISNUMBER(MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0))),IFERROR(INDEX(\'Recetas de Material\'!$G$3:$G$30,MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0)),0),0))/(IF(ISNUMBER(SEARCH("+",G5)),IFERROR(INDEX(\'Recetas de Material\'!$C$3:$C$30,MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0)),100),100)+IF(AND(ISNUMBER(SEARCH("+",G5)),ISNUMBER(MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0))),IFERROR(INDEX(\'Recetas de Material\'!$E$3:$E$30,MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0)),0),0)+IF(AND(ISNUMBER(SEARCH("+",G5)),ISNUMBER(MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0))),IFERROR(INDEX(\'Recetas de Material\'!$G$3:$G$30,MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0)),0),0)+IF(AND(ISNUMBER(SEARCH("+",G5)),ISNUMBER(MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0))),IFERROR(INDEX(\'Recetas de Material\'!$I$3:$I$30,MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0)),0),0))))',
    '=IF(AND(ISNUMBER(SEARCH("+",G5)),ISNUMBER(MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0))),IFERROR(INDEX(\'Recetas de Material\'!$H$3:$H$30,MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0)),""),"")',
    '=IF(V5="","",IF((IF(ISNUMBER(SEARCH("+",G5)),IFERROR(INDEX(\'Recetas de Material\'!$C$3:$C$30,MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0)),100),100)+IF(AND(ISNUMBER(SEARCH("+",G5)),ISNUMBER(MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0))),IFERROR(INDEX(\'Recetas de Material\'!$E$3:$E$30,MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0)),0),0)+IF(AND(ISNUMBER(SEARCH("+",G5)),ISNUMBER(MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0))),IFERROR(INDEX(\'Recetas de Material\'!$G$3:$G$30,MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0)),0),0)+IF(AND(ISNUMBER(SEARCH("+",G5)),ISNUMBER(MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0))),IFERROR(INDEX(\'Recetas de Material\'!$I$3:$I$30,MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0)),0),0))=0,0,V5*(IF(AND(ISNUMBER(SEARCH("+",G5)),ISNUMBER(MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0))),IFERROR(INDEX(\'Recetas de Material\'!$I$3:$I$30,MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0)),0),0))/(IF(ISNUMBER(SEARCH("+",G5)),IFERROR(INDEX(\'Recetas de Material\'!$C$3:$C$30,MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0)),100),100)+IF(AND(ISNUMBER(SEARCH("+",G5)),ISNUMBER(MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0))),IFERROR(INDEX(\'Recetas de Material\'!$E$3:$E$30,MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0)),0),0)+IF(AND(ISNUMBER(SEARCH("+",G5)),ISNUMBER(MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0))),IFERROR(INDEX(\'Recetas de Material\'!$G$3:$G$30,MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0)),0),0)+IF(AND(ISNUMBER(SEARCH("+",G5)),ISNUMBER(MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0))),IFERROR(INDEX(\'Recetas de Material\'!$I$3:$I$30,MATCH(G5,\'Recetas de Material\'!$A$3:$A$30,0)),0),0))))'
  ];

  for (var i = 0; i < formulas.length; i++) {
    sh.getRange(5, 18 + i).setFormula(formulas[i]); // 18 = R
  }
  SpreadsheetApp.flush();
}

function repararFormulasRecetas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var recetas = ss.getSheetByName('Recetas de Material');
  if (!recetas) return 0;
  var lastRow = recetas.getLastRow();
  if (lastRow < 3) return 0;

  var corregidas = 0;
  for (var fila = 3; fila <= lastRow; fila++) {
    for (var slot = 1; slot <= 4; slot++) {
      var colNombre = 2 + (slot - 1) * 2; // B, D, F, H
      recetas.getRange(fila, colNombre).setFormula(formulaComponente(fila, slot));
    }
    recetas.getRange(fila, 10).setFormula(
      '=IF($A' + fila + '="","",SUM(C' + fila + ',E' + fila + ',G' + fila + ',I' + fila + '))'
    );
    corregidas++;
  }
  SpreadsheetApp.flush();
  return corregidas;
}

function repararLinksHistoricoCortes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hist = ss.getSheetByName('Histórico de Cortes');
  if (!hist) return 0;
  var lastRow = hist.getLastRow();
  if (lastRow < 2) return 0;

  // El corte ahora archiva en un Google Sheet aparte; si ya existe lo usamos
  // para buscar ahí primero, y si no, o si el corte es de antes de este
  // cambio, buscamos dentro de este mismo archivo (como se hacía antes).
  var props = PropertiesService.getScriptProperties();
  var historicoId = props.getProperty('HISTORICO_ID');
  var historico = null;
  if (historicoId) {
    try { historico = SpreadsheetApp.openById(historicoId); } catch (e) { historico = null; }
  }

  var corregidos = 0;
  var etiquetas = hist.getRange(2, 2, lastRow - 1, 1).getValues(); // columna B
  for (var i = 0; i < etiquetas.length; i++) {
    var fila = i + 2;
    var etiqueta = String(etiquetas[i][0] || '').trim();
    if (!etiqueta) continue;

    // Busca la pestaña archivada de esa etiqueta (puede tener sufijo "(2)" si hubo choque de nombres),
    // primero en el archivo histórico aparte y, si no aparece ahí, dentro de este mismo archivo
    // (cortes viejos, de antes de que el corte empezara a archivar aparte) — probando también el
    // nombre viejo ("Septiembre") por si ese corte se hizo antes de renombrar la hoja.
    var candidatos = [historico, ss];
    var nombresPosibles = [etiqueta + ' — Seguimiento', etiqueta + ' — Septiembre'];
    var hojaArchivada = null;
    var archivoDeLaHoja = null;
    for (var c = 0; c < candidatos.length && !hojaArchivada; c++) {
      var libro = candidatos[c];
      if (!libro) continue;
      for (var n = 0; n < nombresPosibles.length && !hojaArchivada; n++) {
        var encontrada = libro.getSheetByName(nombresPosibles[n]);
        if (encontrada) { hojaArchivada = encontrada; archivoDeLaHoja = libro; }
      }
      if (!hojaArchivada) {
        var todas = libro.getSheets();
        for (var s = 0; s < todas.length; s++) {
          var nombreHoja = todas[s].getName();
          if (nombreHoja.indexOf(etiqueta + ' — Seguimiento') === 0 || nombreHoja.indexOf(etiqueta + ' — Septiembre') === 0) {
            hojaArchivada = todas[s];
            archivoDeLaHoja = libro;
            break;
          }
        }
      }
    }
    if (!hojaArchivada) continue;

    var url = 'https://docs.google.com/spreadsheets/d/' + archivoDeLaHoja.getId() + '/edit#gid=' + hojaArchivada.getSheetId();
    hist.getRange(fila, 8).setFormula('=HYPERLINK("' + url + '", "Ver copia archivada")');
    corregidos++;
  }
  SpreadsheetApp.flush();
  return corregidos;
}

// ============================================================
// 7) Cruza el consumo (Seguimiento) contra las existencias pegadas en "Inv"
//    - Materia Prima: se agrupa por TU columna C de "código interno" (columna
//      escrita a mano) -> es exacta, no hay ambigüedad.
//    - Colorantes: NO se puede usar el código (no coincide con Catálogo
//      Colores), así que se busca el NOMBRE del color como texto dentro del
//      nombre del producto en Inv. Si hay más de una coincidencia o
//      ninguna, se marca en vez de adivinar.
// ============================================================
function normalizar(s) {
  return String(s || '').toUpperCase()
    .replace(/Á/g, 'A').replace(/É/g, 'E').replace(/Í/g, 'I').replace(/Ó/g, 'O').replace(/Ú/g, 'U')
    .trim();
}

function parseExistencia(v) {
  if (typeof v === 'number') return v;
  var n = parseFloat(String(v).replace(/\./g, '').replace(',', '.'));
  return isNaN(n) ? null : n;
}

// Botón del menú: cruza y muestra el resultado en pantalla
function actualizarInventario() {
  SpreadsheetApp.getUi().alert(actualizarInventarioCore());
}

// La lógica real del cruce, sin cuadros de diálogo, para que también la puedan
// usar "Actualizar todo" y la importación automática (que corre sin pantalla).
// Devuelve el mensaje de resultado.
function actualizarInventarioCore() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja1 = ss.getSheetByName('Inv');
  var sh = ss.getSheetByName('Seguimiento');
  var consMat = ss.getSheetByName('Consumo por Material');
  var resColor = ss.getSheetByName('Resumen Consumo');
  if (!hoja1) return 'No encuentro una pestaña llamada "Inv".';

  var lastRow = hoja1.getLastRow();
  var datos = hoja1.getRange(1, 1, lastRow, 4).getValues(); // A Código, B Nombre, C tu pista, D Existencia

  var existenciaPorCodigoInterno = {};
  var colorantes = [];

  datos.forEach(function (fila) {
    var cod = String(fila[0] || '').trim();
    var nombre = String(fila[1] || '').trim();
    var hint = String(fila[2] || '').trim();
    var existencia = parseExistencia(fila[3]);
    if (existencia === null) return;

    if (cod.indexOf('01-') === 0 && hint) {
      existenciaPorCodigoInterno[hint] = (existenciaPorCodigoInterno[hint] || 0) + existencia;
    } else if (cod.indexOf('02-') === 0 && nombre) {
      colorantes.push({ nombre: normalizar(nombre), existencia: existencia });
    }
  });

  // --- Días distintos con producción registrada en Seguimiento, para sacar un
  //     promedio de consumo diario y de ahí los "Días de cobertura" ---
  var lastRowSep = sh.getLastRow();
  var diasDistintos = 0;
  if (lastRowSep >= 5) {
    var nSep = lastRowSep - 4;
    var fechasCol = sh.getRange(5, 1, nSep, 1).getValues();
    var cantidadCol = sh.getRange(5, 9, nSep, 1).getValues();
    var setFechas = {};
    for (var k = 0; k < nSep; k++) {
      var cant = cantidadCol[k][0];
      if (cant === '' || cant === null) continue;
      var f = fechasCol[k][0];
      var key = (Object.prototype.toString.call(f) === '[object Date]')
        ? Utilities.formatDate(f, Session.getScriptTimeZone(), 'yyyy-MM-dd')
        : String(f || '').trim();
      if (key) setFechas[key] = true;
    }
    diasDistintos = Object.keys(setFechas).length;
  }

  // --- Consumo por Material: columnas G (Total Inventario), H (Restante), I (Días de cobertura) ---
  var lastRowMat = consMat.getLastRow();
  consMat.getRange(2, 7, 1, 3).setValues([['Existencia Actual', 'Consumo Promedio Diario', 'Días de Cobertura']]);
  if (lastRowMat >= 3) {
    var nMat = lastRowMat - 2;
    var materiales = consMat.getRange(3, 1, nMat, 1).getValues();
    var consumidos = consMat.getRange(3, 2, nMat, 1).getValues();
    // ordenamos las pistas de más larga a más corta, para que "20H10NA" gane
    // sobre una pista corta tipo "20H" si ambas existieran
    var pistas = Object.keys(existenciaPorCodigoInterno).sort(function (a, b) { return b.length - a.length; });
    for (var i = 0; i < nMat; i++) {
      var mat = String(materiales[i][0] || '').trim();
      var matNorm = normalizar(mat);
      var r = 3 + i;
      if (!mat || mat === 'TOTAL') continue;
      // 1) intento directo: el texto del material contiene alguna pista tal cual
      var pistaEncontrada = pistas.filter(function (p) { return matNorm.indexOf(normalizar(p)) !== -1; })[0];
      // 2) si no, uso el mismo diccionario de sinónimos (alias) que usamos para
      //    armar Inv — así "20400 Excelent" también encuentra el "2213",
      //    aunque el texto de la orden no diga "2213" literalmente
      if (!pistaEncontrada) {
        var codigoAlias = detectarCodigoInterno(mat);
        if (codigoAlias && existenciaPorCodigoInterno[codigoAlias] !== undefined) {
          pistaEncontrada = codigoAlias;
        }
      }
      if (pistaEncontrada) {
        // La Existencia de Inv YA es la foto real de hoy (ya incluye lo
        // consumido y lo que entró nuevo) — no se le resta el consumo del mes,
        // eso sería contarlo dos veces. Los "Días de Cobertura" salen de
        // dividir esa existencia entre el ritmo de consumo diario.
        var exist = existenciaPorCodigoInterno[pistaEncontrada];
        var promedioDiario = diasDistintos > 0 ? consumidos[i][0] / diasDistintos : 0;
        consMat.getRange(r, 7).setValue(exist);
        consMat.getRange(r, 8).setValue(Math.round(promedioDiario * 100) / 100);
        if (promedioDiario > 0) {
          consMat.getRange(r, 9).setValue(Math.round((exist / promedioDiario) * 10) / 10);
        } else {
          consMat.getRange(r, 9).setValue('sin consumo reciente');
        }
      } else {
        consMat.getRange(r, 7).setValue('sin match');
        consMat.getRange(r, 8).setValue('');
        consMat.getRange(r, 9).setValue('');
      }
    }
  }

  // --- Resumen Consumo: columnas H (Existencia), I (Restante), J (match usado), K (Días de cobertura) ---
  var lastRowCol = resColor.getLastRow();
  resColor.getRange(2, 8, 1, 4).setValues([['Existencia', 'Consumo Promedio Diario', 'Coincidencia usada', 'Días de Cobertura']]);
  if (lastRowCol >= 3) {
    var nCol = lastRowCol - 2;
    var coloresNombre = resColor.getRange(3, 2, nCol, 1).getValues();
    var consumidosColor = resColor.getRange(3, 3, nCol, 1).getValues();
    for (var j = 0; j < nCol; j++) {
      var colorNombre = normalizar(coloresNombre[j][0]);
      var r2 = 3 + j;
      if (!colorNombre || colorNombre === 'TOTAL') continue;
      var matches = colorantes.filter(function (c) { return c.nombre.indexOf(colorNombre) !== -1; });
      if (matches.length === 1) {
        // Igual que en materiales: la Existencia ya es la foto de hoy, no se
        // le resta el consumo del mes otra vez.
        var promedioDiarioColor = diasDistintos > 0 ? consumidosColor[j][0] / diasDistintos : 0;
        resColor.getRange(r2, 8).setValue(matches[0].existencia);
        resColor.getRange(r2, 9).setValue(Math.round(promedioDiarioColor * 100) / 100);
        resColor.getRange(r2, 10).setValue(matches[0].nombre);
        if (promedioDiarioColor > 0) {
          resColor.getRange(r2, 11).setValue(Math.round((matches[0].existencia / promedioDiarioColor) * 10) / 10);
        } else {
          resColor.getRange(r2, 11).setValue('sin consumo reciente');
        }
      } else if (matches.length > 1) {
        resColor.getRange(r2, 8).setValue('AMBIGUO (' + matches.length + ')');
        resColor.getRange(r2, 9).setValue('');
        resColor.getRange(r2, 10).setValue(matches.map(function (m) { return m.nombre; }).join(' | '));
        resColor.getRange(r2, 11).setValue('');
      } else {
        resColor.getRange(r2, 8).setValue('sin match');
        resColor.getRange(r2, 9).setValue('');
        resColor.getRange(r2, 10).setValue('');
        resColor.getRange(r2, 11).setValue('');
      }
    }
  }

  var pistasEncontradas = Object.keys(existenciaPorCodigoInterno);
  return 'Inventario cruzado contra "Inv" (con ' + diasDistintos + ' días de producción registrados ' +
    'en Seguimiento para el promedio diario).\n\n' +
    '- Materia Prima: "Consumo por Material", columnas G, H e I (Días de Cobertura).\n' +
    '- Colorantes: "Resumen Consumo", columnas H a K — revisa la columna J ' +
    '("Coincidencia usada") para confirmar que cada cruce tiene sentido, sobre todo ' +
    'donde diga AMBIGUO o sin match.\n\n' +
    '🔍 DIAGNÓSTICO: encontré ' + pistasEncontradas.length + ' código(s) interno(s) en Inv: [' +
    pistasEncontradas.join(', ') + ']';
}

// ============================================================
// 10) Congela encabezados y protege (con advertencia) las columnas que
//     llena el script solo, para que nadie las borre sin querer por
//     accidente. "Advertencia" y no bloqueo total: sigue pudiéndose
//     editar si hace falta, pero primero pregunta "¿estás segura?".
// ============================================================
function configurarProteccionYFormato() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var congelar = {
    'Seguimiento': 4,
    'Consumo por Material': 2,
    'Resumen Consumo': 2,
    'Recetas de Material': 2,
    'Dashboard': 6,
    'Inv': 2,
    'Catálogo Colores': 1
  };
  Object.keys(congelar).forEach(function (nombre) {
    var sh = ss.getSheetByName(nombre);
    if (sh) sh.setFrozenRows(congelar[nombre]);
  });

  function protegerConAdvertencia(nombreHoja, rangoA1, descripcion) {
    var sh = ss.getSheetByName(nombreHoja);
    if (!sh) return;
    // Si ya existe una protección con esta misma descripción, no duplicar
    var yaExiste = sh.getProtections(SpreadsheetApp.ProtectionType.RANGE)
      .some(function (p) { return p.getDescription() === descripcion; });
    if (yaExiste) return;
    var p = sh.getRange(rangoA1).protect();
    p.setDescription(descripcion);
    p.setWarningOnly(true); // avisa antes de editar, no bloquea (para no encerrarte a ti misma)
  }

  protegerConAdvertencia('Seguimiento', 'R5:AD2000', 'Fórmulas calculadas (Código Color, Kg, Componentes...) — no se editan a mano');
  protegerConAdvertencia('Consumo por Material', 'G3:K1000', 'Cruce de inventario — lo llena el botón "Cruzar inventario"');
  protegerConAdvertencia('Resumen Consumo', 'H3:K1000', 'Cruce de inventario — lo llena el botón "Cruzar inventario"');

  ui.alert(
    '✅ Listo.\n\n' +
    '- Encabezados congelados en Seguimiento, Consumo por Material, Resumen Consumo, Recetas de Material y Dashboard.\n' +
    '- Las columnas que llenan los botones automáticamente (fórmulas de Seguimiento y el cruce de inventario) ' +
    'ahora muestran una advertencia antes de dejarte editarlas a mano, para evitar borrones accidentales.\n\n' +
    'Esto es seguro de correr de nuevo cuando quieras, no duplica protecciones.'
  );
}

// ============================================================
// 11) PORTADA — pestaña "Inicio" con KPIs, alerta de materiales en riesgo
//     y links directos a cada pestaña. Se puede correr sola desde el menú,
//     o se actualiza automáticamente cada vez que corres "Actualizar todo".
// ============================================================
function actualizarPortada() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var NOMBRE = 'Inicio';
  var inicio = ss.getSheetByName(NOMBRE);
  if (!inicio) {
    inicio = ss.insertSheet(NOMBRE, 0);
  } else {
    inicio.clear();
  }
  inicio.setHiddenGridlines(true);
  inicio.setColumnWidth(1, 260);
  inicio.setColumnWidths(2, 5, 130);

  var d = calcularDatosDashboard();
  var inv = leerInventarioParaDashboard();
  var enRiesgo = inv.materiales.filter(function (m) { return typeof m.diasCobertura === 'number' && m.diasCobertura < 7; }).length
    + inv.colores.filter(function (c) { return typeof c.diasCobertura === 'number' && c.diasCobertura < 7; }).length;
  // Cuántos renglones tienen de verdad un número de días de cobertura. Si es 0,
  // el cruce no se ha hecho y NO se puede decir que todo está bien.
  var conCobertura = inv.materiales.filter(function (m) { return typeof m.diasCobertura === 'number'; }).length
    + inv.colores.filter(function (c) { return typeof c.diasCobertura === 'number'; }).length;

  // --- Encabezado ---
  inicio.getRange('A1:F1').merge().setValue('📊 SEGUIMIENTO MANAPLAS')
    .setFontSize(20).setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#1F4E5F')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  inicio.setRowHeight(1, 46);
  inicio.getRange('A2:F2').merge().setValue('Consumo de materia prima y colorantes — Producción e Inventarios')
    .setFontStyle('italic').setFontColor('#5C6B73').setHorizontalAlignment('center');

  // --- Alerta ---
  inicio.getRange('A4:F4').merge();
  if (enRiesgo > 0) {
    inicio.getRange('A4').setValue('⚠️  ' + enRiesgo + ' material(es)/color(es) con menos de 7 días de cobertura — revisa "Consumo por Material" / "Resumen Consumo"')
      .setBackground('#FFC7CE').setFontColor('#9C0006');
  } else if (conCobertura === 0) {
    inicio.getRange('A4').setValue('⚠️  Cobertura sin calcular — corre "📦 Cruzar inventario (Inv)" o "📥 Leer inventario más reciente"')
      .setBackground('#FFF2CC').setFontColor('#7F6000');
  } else {
    inicio.getRange('A4').setValue('✅  Todos los materiales y colores con inventario cruzado tienen 7 días de cobertura o más')
      .setBackground('#D9EAD3').setFontColor('#274E13');
  }
  inicio.getRange('A4:F4').setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
  inicio.setRowHeight(4, 30);

  // --- KPIs ---
  var kpis = [
    ['Órdenes registradas', d.totalOrdenes],
    ['Días con producción', d.diasDistintos],
    ['Productos distintos', d.productosCount],
    ['Kg Materia Prima', Math.round(d.totalKgMp)],
    ['Kg Colorante', Math.round(d.totalKgColor)],
    ['Kg Total del periodo', Math.round(d.totalKgMp + d.totalKgColor)]
  ];
  kpis.forEach(function (kv, idx) {
    var col = idx + 1;
    inicio.getRange(6, col).setValue(kv[0]).setFontSize(9).setFontColor('#5C6B73').setWrap(true);
    inicio.getRange(7, col).setValue(kv[1]).setFontSize(16).setFontWeight('bold').setFontColor('#1F4E5F');
  });
  inicio.getRange(6, 1, 2, 6).setBorder(true, true, true, true, true, true, '#E8EEF1', SpreadsheetApp.BorderStyle.SOLID);

  // --- Navegación ---
  inicio.getRange('A9').setValue('Ir directo a:').setFontWeight('bold').setFontColor('#1F4E5F');
  var destinos = ['Seguimiento', 'Dashboard', 'Consumo por Material', 'Resumen Consumo', 'Inv', 'Histórico de Cortes', 'Recetas de Material'];
  var fila = 10;
  var urlBase = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/edit#gid=';
  destinos.forEach(function (nombre) {
    var hoja = ss.getSheetByName(nombre);
    if (!hoja) return;
    inicio.getRange(fila, 1).setFormula('=HYPERLINK("' + urlBase + hoja.getSheetId() + '", "→ ' + nombre + '")')
      .setFontColor('#0092CF').setFontWeight('bold');
    fila++;
  });

  // Si ya existe el archivo histórico aparte (donde caen los cortes), agrega un link directo a él
  var historicoId = PropertiesService.getScriptProperties().getProperty('HISTORICO_ID');
  if (historicoId) {
    inicio.getRange(fila, 1).setFormula('=HYPERLINK("https://docs.google.com/spreadsheets/d/' + historicoId + '/edit", "→ Histórico (archivo aparte)")')
      .setFontColor('#0092CF').setFontWeight('bold');
    fila++;
  }

  var ahora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  inicio.getRange(fila + 1, 1).setValue('Última actualización: ' + ahora).setFontSize(8).setFontColor('#5C6B73');

  ss.setActiveSheet(inicio);
  ss.moveActiveSheet(1);
}

// ============================================================
// 12) Colorea TODAS las pestañas por tipo, para que sea fácil ubicarse
//     de un vistazo en la barra de pestañas de abajo:
//     azul = Inicio, celeste = Seguimiento (dato vivo), verde = resúmenes,
//     coral = Dashboard, gris = hojas de referencia/soporte,
//     gris oscuro = Histórico de Cortes, gris claro = pestañas archivadas.
// ============================================================
function colorearPestanas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var COLORES = {
    'Inicio': '#1F4E5F',
    'Seguimiento': '#0092CF',
    'Consumo por Material': '#548235',
    'Resumen Consumo': '#548235',
    'Dashboard': '#C0533E',
    'Inv': '#8E7CC3',
    'Catálogo Colores': '#8E7CC3',
    'Recetas de Material': '#8E7CC3',
    'Histórico de Cortes': '#163A47'
  };

  ss.getSheets().forEach(function (sh) {
    var nombre = sh.getName();
    if (COLORES[nombre]) {
      sh.setTabColor(COLORES[nombre]);
    } else if (nombre.indexOf('Cierre ') === 0) {
      sh.setTabColor('#B7C1C6'); // pestañas archivadas, ya congeladas
    }
  });

  SpreadsheetApp.getUi().alert(
    '✅ Listo. Todas las pestañas quedaron coloreadas por tipo:\n\n' +
    '🔵 Inicio    🩵 Seguimiento    🟢 Resúmenes    🟠 Dashboard\n' +
    '🟣 Inv / Catálogo / Recetas (referencia)    ⚫ Histórico de Cortes    ⚪ Cierres archivados'
  );
}

// ============================================================
// Crea el menú "Seguimiento" al abrir el archivo (alternativa al botón)
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Seguimiento')
    .addItem('📥 Traer órdenes de la carpeta', 'importarDesdeCarpeta')
    .addItem('🔄 Actualizar todo', 'actualizarTodo')
    .addItem('🏠 Actualizar portada (Inicio)', 'actualizarPortada')
    .addItem('🍳 Buscar recetas de material faltantes', 'detectarRecetasFaltantesBoton')
    .addItem('📥 Leer inventario más reciente (carpeta)', 'leerInventarioDesdeCarpeta')
    .addItem('📦 Cruzar inventario (Inv)', 'actualizarInventario')
    .addItem('🧹 Eliminar filas duplicadas', 'eliminarFilasDuplicadas')
    .addItem('📅 Corte del mes (cerrar y archivar)', 'corteDelMes')
    .addItem('🔒 Congelar y proteger fórmulas', 'configurarProteccionYFormato')
    .addItem('🔧 Reparar errores (#ERROR!)', 'repararErroresConocidos')
    .addItem('🎨 Colorear pestañas por tipo', 'colorearPestanas')
    .addItem('🗑️ Borrar', 'limpiarDatosSeptiembre')
    .addToUi();
}

// ============================================================
// 9) Busca filas de Seguimiento con exactamente el mismo
//    Fecha + OP + Sku + Color, y borra las copias repetidas
//    (deja una sola de cada orden real).
// ============================================================
function eliminarFilasDuplicadas() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Seguimiento');
  var lastRow = sh.getLastRow();
  if (lastRow < 5) { ui.alert('No hay datos que revisar.'); return; }

  var n = lastRow - 4;
  var datos = sh.getRange(5, 1, n, 10).getValues(); // A..J
  var vistos = {};
  var filasABorrar = [];

  for (var i = 0; i < n; i++) {
    var fila = datos[i];
    var fecha = fila[0], op = fila[1], sku = fila[3], color = fila[9];
    if ((fecha === '' || fecha === null) && (op === '' || op === null)) continue; // fila vacía
    var fKey = (Object.prototype.toString.call(fecha) === '[object Date]')
      ? Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : String(fecha || '').trim();
    var key = fKey + '|' + op + '|' + sku + '|' + color;
    if (vistos[key]) {
      filasABorrar.push(5 + i); // número de fila real en la hoja
    } else {
      vistos[key] = true;
    }
  }

  if (filasABorrar.length === 0) {
    ui.alert('No encontré filas duplicadas (mismo Fecha + OP + Sku + Color). Todo está en orden.');
    return;
  }

  var respuesta = ui.alert(
    'Encontré ' + filasABorrar.length + ' fila(s) duplicada(s)',
    'Tienen exactamente el mismo Fecha + OP + Sku + Color que otra fila ya existente. ' +
    'Voy a borrar las copias repetidas y dejar una sola de cada orden. ¿Confirmas?',
    ui.ButtonSet.YES_NO
  );
  if (respuesta !== ui.Button.YES) { ui.alert('Cancelado. No se borró nada.'); return; }

  // borrar de abajo hacia arriba para que no se recorran los números de fila
  filasABorrar.sort(function (a, b) { return b - a; });
  filasABorrar.forEach(function (r) { sh.deleteRow(r); });

  SpreadsheetApp.flush();
  actualizarResumenColores();
  actualizarResumenMateriales();
  actualizarInventarioCore();
  actualizarDashboard();
  SpreadsheetApp.flush();

  ui.alert(
    'Listo. Se borraron ' + filasABorrar.length + ' filas duplicadas.\n\n' +
    'Los resúmenes (Resumen Consumo, Consumo por Material) y el Dashboard ya se recalcularon solos.'
  );
}

// ============================================================
// 8) Lee el archivo de existencias MÁS RECIENTE de una carpeta de Drive
//    (el formato binario viejo, sin importar el nombre raro que tenga),
//    lo descifra, y llena "Inv" solo. Luego cruza el inventario igual
//    que el botón de siempre.
//
//    OJO: reemplaza CARPETA_INVENTARIO_ID de abajo con el ID de TU carpeta
//    de inventario antes de usar este botón (el ID es la parte del link
//    de Drive después de /folders/).
// ============================================================
var CARPETA_INVENTARIO_ID = '1b0oi6-EGYdC8yVOjX1l56HAwkjaLWDM7';

// Cada entrada es [palabra que buscamos en el nombre del producto, código interno
// al que pertenece]. Varias palabras pueden apuntar al MISMO código cuando son
// sustitutos/homotipos del mismo material (ej. "20400" y "2111" son el mismo
// material que "2213", solo que de otra marca).
var ALIAS_MATERIAL_MP = [
  ['20H10NA', '20H10NA'], ['11H01A', '11H01A'], ['60C90ND', '60C90ND'],
  ['2213', '2213'], ['20400', '2213'], ['2111', '2213'],
  ['0240', '0240'], ['6200B', '6200B'], ['0506C', '0506C'],
  ['2811', '2811'], ['2908', '2908'], ['40H92C', '40H92C'],
  ['30R40CD', '30R'], ['30R', '30R'], ['20H01NA', '20H01NA'], ['02R01CA-1', '02R01CA-1'],
  ['35R80', '35R80'], ['35R81', '35R81'], ['35R82', '35R82'], ['35R83', '35R83'],
  ['05C06', '05C06'], ['478', '478']
];
ALIAS_MATERIAL_MP.sort(function (a, b) { return b[0].length - a[0].length; });

function detectarCodigoInterno(nombre) {
  var up = String(nombre).toUpperCase();
  for (var k = 0; k < ALIAS_MATERIAL_MP.length; k++) {
    if (up.indexOf(ALIAS_MATERIAL_MP[k][0]) !== -1) return ALIAS_MATERIAL_MP[k][1];
  }
  return '';
}

// ============================================================
// Diccionario de materiales: nombre legible + familia de resina
// (PPH=Homopolímero, PPC=Copolímero, PEAD=Polietileno alta densidad,
// PEBD=Polietileno baja densidad, PS=Poliestireno) para cada código interno.
// Sacado de las descripciones reales de tu inventario. Si aparece un código
// nuevo que no está aquí, agrégalo a esta lista.
// ============================================================
var DICCIONARIO_MATERIALES = {
  '11H01A':  { nombre: 'PP Homopolímero 11H01A',              familia: 'PPH' },
  '20H10NA': { nombre: 'PP Homopolímero 20H10NA',              familia: 'PPH' },
  '20H01NA': { nombre: 'PP Homopolímero 20H01NA',              familia: 'PPH' },
  '40H92C':  { nombre: 'PP Homopolímero 40H92C',               familia: 'PPH' },
  '05C06':   { nombre: 'PP Copolímero 05C06',                  familia: 'PPC' },
  '60C90ND': { nombre: 'PP Copolímero 60C90ND',                familia: 'PPC' },
  '02R01CA-1': { nombre: 'PP Copolímero Random 02R01CA-1',     familia: 'PPC' },
  '30R':     { nombre: 'PP Copolímero 30R40CD',                familia: 'PPC' },
  '35R80':   { nombre: 'PP Copolímero 35R80',                  familia: 'PPC' },
  '35R81':   { nombre: 'PP Copolímero 35R81',                  familia: 'PPC' },
  '35R82':   { nombre: 'PP Copolímero 35R82',                  familia: 'PPC' },
  '35R83':   { nombre: 'PP Copolímero 35R83',                  familia: 'PPC' },
  '6200B':   { nombre: 'PEAD 6200B (HD5502H)',                 familia: 'PEAD' },
  '2811':    { nombre: 'PEAD 2811 (HD2000-GP)',                familia: 'PEAD' },
  '2908':    { nombre: 'PEAD 2908 (Alathon/Esenttia)',         familia: 'PEAD' },
  '2213':    { nombre: 'PEBD Lineal 2213 (Excelene/Star Resin/LLB5000)', familia: 'PEBD' },
  '0240':    { nombre: 'PEBD 0240 (Westlake EF412JM)',         familia: 'PEBD' },
  '0506C':   { nombre: 'Aditivo 0506C',                        familia: 'Aditivo' },
  'REF':     { nombre: 'Refuerzo',                             familia: 'Aditivo' },
  'MOLIDO':  { nombre: 'Molido (reciclado interno)',           familia: 'Reciclado' },
  '478':     { nombre: 'Poliestireno de Alto Impacto STYRON 478', familia: 'PS' },
  'PPH':     { nombre: 'PP Homopolímero (grado no especificado)', familia: 'PPH' },
  'PPC':     { nombre: 'PP Copolímero (grado no especificado)',   familia: 'PPC' }
};

function infoMaterial(codigoOnombre) {
  var d = DICCIONARIO_MATERIALES[codigoOnombre];
  if (d) return d;
  return { nombre: codigoOnombre, familia: '' };
}

// Detecta resina que ya viene coloreada de fábrica (no colorante que se dosifica
// aparte). Si más adelante aparece otro patrón de nombre que también sea de este
// tipo, agrégalo a esta lista de palabras.
var PALABRAS_RESINA_CARGADA = ['CARGADO', 'CARGADA', 'POLIPROPILENO VINO'];
function esResinaPreColoreada(nombre) {
  var up = String(nombre).toUpperCase();
  for (var k = 0; k < PALABRAS_RESINA_CARGADA.length; k++) {
    if (up.indexOf(PALABRAS_RESINA_CARGADA[k]) !== -1) return true;
  }
  return false;
}

function parseNumeroInventario(s) {
  var n = parseFloat(String(s).replace(/\./g, '').replace(',', '.'));
  return isNaN(n) ? null : n;
}

// Extrae todos los textos con prefijo de longitud (2 bytes little-endian)
// del archivo binario — así está armado el formato viejo del sistema.
function extraerStringsBinario(bytes) {
  var strings = [];
  var i = 0;
  var n = bytes.length;
  while (i < n - 2) {
    var b0 = bytes[i] & 0xFF;
    var b1 = bytes[i + 1] & 0xFF;
    var length = b0 | (b1 << 8);
    if (length > 0 && length < 120 && i + 2 + length <= n) {
      var chars = [];
      var esImprimible = true;
      for (var j = 0; j < length; j++) {
        var byteVal = bytes[i + 2 + j] & 0xFF;
        if (byteVal < 32) { esImprimible = false; break; }
        chars.push(byteVal);
      }
      if (esImprimible) {
        var texto = '';
        for (var m = 0; m < chars.length; m++) texto += String.fromCharCode(chars[m]);
        strings.push(texto);
        i += 2 + length;
        continue;
      }
    }
    i++;
  }
  return strings;
}

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
  var filas; // arreglo de [Codigo, Nombre, Unidad, Almacen, Existencia]
  if (masReciente.getMimeType() === MimeType.GOOGLE_SHEETS) {
    var libroInv = SpreadsheetApp.openById(masReciente.getId());
    var hojaInv = libroInv.getSheets()[0];
    var datosInv = hojaInv.getDataRange().getValues(); // fila 0 = encabezado
    filas = datosInv.slice(1).map(function (r) {
      return [String(r[0] || ''), String(r[1] || ''), String(r[2] || ''), String(r[3] || ''), r[4]];
    });
  } else {
    var bytes = masReciente.getBlob().getBytes();
    var strings = extraerStringsBinario(bytes);
    if (strings.length < 10) {
      ui.alert('No pude leer el archivo "' + masReciente.getName() + '" — no tiene la estructura esperada.');
      return;
    }
    var valores = strings.slice(5); // los primeros 5 textos son encabezados
    filas = [];
    for (var i = 0; i + 4 < valores.length; i += 5) {
      filas.push([valores[i], valores[i + 1], valores[i + 2], valores[i + 3], valores[i + 4]]);
    }
  }

  var materiaPrima = [], colorantes = [];
  filas.forEach(function (fila) {
    var cod = String(fila[0] || '').trim();
    var nombre = fila[1];
    if (cod === 'Total' || cod === '') return;
    var existencia = (typeof fila[4] === 'number') ? fila[4] : parseNumeroInventario(fila[4]);
    if (existencia === null || existencia === undefined) return;
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

  // --- Escribir en "Inv", en los mismos dos bloques de siempre ---
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja1 = ss.getSheetByName('Inv');
  hoja1.clear();

  var HEADER_FILL_MP = '#548235', HEADER_FILL_COL = '#B45F06', HEADER_FILL = '#1F4E5F';
  var FILL_MP = '#E2EFDA', FILL_COL = '#FCE4D6';

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
  materiaPrima.forEach(function (row) {
    // columnas A y C como texto forzado: evita que Sheets convierta solo
    // codigos como "0240" o "2811" a numero (y pierda el cero adelante)
    hoja1.getRange(r, 1, 1, 2).setNumberFormat('@');
    hoja1.getRange(r, 3, 1, 1).setNumberFormat('@');
    hoja1.getRange(r, 1, 1, 4).setValues([row]).setBackground(FILL_MP);
    r++;
  });
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
  colorantes.forEach(function (row) {
    hoja1.getRange(r, 1, 1, 4).setValues([row]).setBackground(FILL_COL);
    r++;
  });
  var colEnd = r - 1;
  hoja1.getRange(r, 2).setValue('Subtotal Colorantes').setFontWeight('bold');
  hoja1.getRange(r, 4).setFormula('=SUM(D' + colStart + ':D' + colEnd + ')').setFontWeight('bold');

  hoja1.setColumnWidth(1, 110).setColumnWidth(2, 380).setColumnWidth(3, 220).setColumnWidth(4, 110);

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
