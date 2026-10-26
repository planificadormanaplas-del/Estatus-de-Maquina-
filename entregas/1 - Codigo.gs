function actualizarTodo() {
  try {
    var reparadas = repararFormulasFaltantes();
    SpreadsheetApp.flush();
    var filasConError = detectarFilasConError();
    actualizarResumenColores();
    actualizarResumenMateriales();
    actualizarInventarioCore();
    actualizarDashboard();
    SpreadsheetApp.flush();
    actualizarPortada();
    var mensaje = "Seguimiento actualizado.\n" + "- Filas reparadas (les faltaban f\xf3rmulas): " + reparadas + "\n" + "- Resumen Consumo (colorante): reconstruido\n" + "- Consumo por Material: reconstruido\n" + "- Cruce de inventario (D\xedas de Cobertura): recalculado\n" + "- Dashboard: reconstruido\n" + "- Portada (Inicio): actualizada";
    if (filasConError.length > 0) {
      mensaje += '\n\n\u26a0\ufe0f OJO: estas filas de "Seguimiento" tienen un dato faltante o mal ' + "escrito (por eso no calculan bien, seguramente ah\xed ves el #VALUE!): fila(s) " + filasConError.join(", ") + ". Revisa que Cantidad y Peso sean n\xfameros en esas filas.";
    }
    SpreadsheetApp.getUi().alert(mensaje);
  } catch (e) {
    SpreadsheetApp.getUi().alert("Se interrumpi\xf3 con un error. Copia este texto completo y comp\xe1rtelo:\n\n" + e.message);
  }
}

function detectarFilasConError() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName("Seguimiento");
  var lastRow = sh.getLastRow();
  var filas = [];
  if (lastRow < 5) return filas;
  var n = lastRow - 4;
  var cantidades = sh.getRange(5, 9, n, 1).getValues();
  var kgColorante = sh.getRange(5, 20, n, 1).getValues();
  for (var i = 0; i < n; i++) {
    var cantidad = cantidades[i][0];
    if (cantidad === "" || cantidad === null) continue;
    var t = kgColorante[i][0];
    if (typeof t === "string" && t.indexOf("#") === 0) filas.push(5 + i);
  }
  return filas;
}

function repararFormulasFaltantes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName("Seguimiento");
  var lastRow = sh.getLastRow();
  var startRow = 5;
  var templateRange = sh.getRange(5, 18, 1, 13);
  var reparadas = 0;
  if (lastRow < startRow) return 0;
  var n = lastRow - startRow + 1;
  var cantidades = sh.getRange(startRow, 9, n, 1).getValues();
  var formulasR = sh.getRange(startRow, 18, n, 1).getFormulas();
  for (var i = 0; i < n; i++) {
    var r = startRow + i;
    var cantidad = cantidades[i][0];
    var formulaCodigoColor = formulasR[i][0];
    if (cantidad !== "" && cantidad !== null && formulaCodigoColor === "") {
      templateRange.copyTo(sh.getRange(r, 18, 1, 13));
      reparadas++;
    }
  }
  return reparadas;
}

function actualizarResumenColores() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName("Seguimiento");
  var resumen = ss.getSheetByName("Resumen Consumo");
  var lastRow = sh.getLastRow();
  var datos = {};
  if (lastRow >= 5) {
    var n = lastRow - 4;
    var codigosCol = sh.getRange(5, 18, n, 1).getValues();
    var nombresCol = sh.getRange(5, 10, n, 1).getValues();
    var kgCol = sh.getRange(5, 20, n, 1).getValues();
    for (var i = 0; i < n; i++) {
      var cod = codigosCol[i][0], nombre = nombresCol[i][0], kg = kgCol[i][0];
      if (cod && cod !== "" && typeof kg === "number") {
        if (!datos[cod]) datos[cod] = {
          kg: 0,
          n: 0,
          nombre: nombre
        };
        datos[cod].kg += kg;
        datos[cod].n += 1;
      }
    }
  }
  var codigos = Object.keys(datos).sort((function(a, b) {
    return datos[b].kg - datos[a].kg;
  }));
  var totalKg = codigos.reduce((function(s, c) {
    return s + datos[c].kg;
  }), 0);
  var totalN = codigos.reduce((function(s, c) {
    return s + datos[c].n;
  }), 0);
  resumen.getRange(3, 1, 300, 11).clearContent();
  resumen.getRange(3, 1, 300, 7).setBackground(null).setFontWeight("normal");
  resumen.getRange(3, 1, 300, 1).setNumberFormat("@");
  var acumulado = 0;
  var ABC_COLOR = {
    A: "#D9EAD3",
    B: "#FFF2CC",
    C: "#F4F4F4"
  };
  var valores = [], fondos = [];
  codigos.forEach((function(cod, idx) {
    var kg = datos[cod].kg, n = datos[cod].n;
    var pct = totalKg > 0 ? kg / totalKg : 0;
    acumulado += pct;
    var clase = acumulado <= .8 ? "A" : acumulado <= .95 ? "B" : "C";
    valores.push([ cod, datos[cod].nombre, kg, n, pct, acumulado, clase ]);
    var fondo = cod === "(sin mapear)" ? "#FFC7CE" : idx % 2 === 1 ? "#F8F9FA" : null;
    var filaFondos = [ fondo, fondo, fondo, fondo, fondo, fondo, cod === "(sin mapear)" ? fondo : ABC_COLOR[clase] ];
    fondos.push(filaFondos);
  }));
  if (valores.length > 0) {
    resumen.getRange(3, 1, valores.length, 7).setValues(valores).setBackgrounds(fondos);
    resumen.getRange(3, 7, valores.length, 1).setFontWeight("bold").setHorizontalAlignment("center");
  }
  var r = 3 + valores.length;
  resumen.getRange(r, 2, 1, 3).setValues([ [ "TOTAL", totalKg, totalN ] ]);
  resumen.getRange(r, 1, 1, 7).setFontWeight("bold");
  if (codigos.length > 0) {
    resumen.getRange(3, 3, codigos.length, 1).setNumberFormat("#,##0.00");
    resumen.getRange(3, 5, codigos.length, 2).setNumberFormat("0.0%");
    var reglaRoja = SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(7).setBackground("#F4C7C3").setFontColor("#990000").setRanges([ resumen.getRange(3, 11, codigos.length, 1) ]).build();
    var reglaVerde = SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(7).setBackground("#D9EAD3").setFontColor("#274E13").setRanges([ resumen.getRange(3, 11, codigos.length, 1) ]).build();
    resumen.setConditionalFormatRules([ reglaRoja, reglaVerde ]);
  }
  resumen.getRange(1, 9).setValue("\xdaltima actualizaci\xf3n: " + Utilities.formatDate(new Date, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"));
  var topN = Math.min(10, codigos.length);
  if (topN > 0) {
    resumen.getCharts().forEach((function(ch) {
      resumen.removeChart(ch);
    }));
    var rangoNombres = resumen.getRange(3, 2, topN, 1);
    var rangoKg = resumen.getRange(3, 3, topN, 1);
    var chart = resumen.newChart().setChartType(Charts.ChartType.BAR).addRange(rangoNombres).addRange(rangoKg).setPosition(3, 13, 0, 0).setOption("title", "Top " + topN + " colores por kg consumido").setOption("width", 620).setOption("height", 320).setOption("legend", {
      position: "none"
    }).setOption("colors", [ "#B45F06" ]).build();
    resumen.insertChart(chart);
  }
}

function limpiarNombreComponente(nombre) {
  return String(nombre).replace(/^\s*\d+(\.\d+)?%\s*/, "").trim();
}

function actualizarResumenMateriales() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName("Seguimiento");
  var resumen = ss.getSheetByName("Consumo por Material");
  var lastRow = sh.getLastRow();
  var datos = {};
  if (lastRow >= 5) {
    var n = lastRow - 4;
    var bloque = sh.getRange(5, 23, n, 8).getValues();
    var pares = [ [ 0, 1 ], [ 2, 3 ], [ 4, 5 ], [ 6, 7 ] ];
    for (var i = 0; i < n; i++) {
      for (var p = 0; p < pares.length; p++) {
        var nombre = bloque[i][pares[p][0]];
        var kg = bloque[i][pares[p][1]];
        if (nombre && nombre !== "" && typeof kg === "number" && kg > 0) {
          var nombreLimpio = limpiarNombreComponente(nombre);
          var claveAgrupacion = (detectarCodigoInterno(nombreLimpio) || nombreLimpio).toUpperCase();
          if (claveAgrupacion === "(DEFINIR RECETA)") claveAgrupacion = "(definir receta)";
          if (!datos[claveAgrupacion]) datos[claveAgrupacion] = {
            kg: 0,
            n: 0
          };
          datos[claveAgrupacion].kg += kg;
          datos[claveAgrupacion].n += 1;
        }
      }
    }
  }
  var materiales = Object.keys(datos).sort((function(a, b) {
    return datos[b].kg - datos[a].kg;
  }));
  var totalKg = materiales.reduce((function(s, m) {
    return s + datos[m].kg;
  }), 0);
  resumen.getRange(3, 1, 300, 9).clearContent();
  resumen.getRange(3, 1, 300, 6).setBackground(null).setFontWeight("normal").setBorder(false, false, false, false, false, false);
  resumen.getRange(3, 10, 300, 2).clearContent();
  resumen.getRange(3, 11, 300, 1).setBackground(null).setFontWeight("normal");
  resumen.getRange(2, 10, 1, 2).setValues([ [ "Nombre Descriptivo", "Familia" ] ]).setFontWeight("bold").setFontColor("#FFFFFF").setBackground("#0B1F33").setHorizontalAlignment("center");
  var FAMILIA_COLOR = {
    PPH: "#D9EAD3",
    PPC: "#FFF2CC",
    PEAD: "#CFE2F3",
    PEBD: "#D0E0E3",
    Aditivo: "#EAD1DC",
    Reciclado: "#E6E6E6",
    PS: "#F4CCCC"
  };
  var acumulado = 0;
  var ABC_COLOR = {
    A: "#D9EAD3",
    B: "#FFF2CC",
    C: "#F4F4F4"
  };
  var valores = [], fondos = [], infos = [], fondosFamilia = [];
  materiales.forEach((function(m, idx) {
    var kg = datos[m].kg, n = datos[m].n;
    var pct = totalKg > 0 ? kg / totalKg : 0;
    acumulado += pct;
    var clase = acumulado <= .8 ? "A" : acumulado <= .95 ? "B" : "C";
    var info = infoMaterial(m);
    valores.push([ m, kg, n, pct, acumulado, clase ]);
    infos.push([ info.nombre, info.familia ]);
    var fondo = m === "(definir receta)" ? "#FFC7CE" : idx % 2 === 1 ? "#F8F9FA" : null;
    fondos.push([ fondo, fondo, fondo, fondo, fondo, m === "(definir receta)" ? fondo : ABC_COLOR[clase] ]);
    fondosFamilia.push([ FAMILIA_COLOR[info.familia] || null ]);
  }));
  if (valores.length > 0) {
    resumen.getRange(3, 1, valores.length, 1).setNumberFormat("@");
    resumen.getRange(3, 1, valores.length, 6).setValues(valores).setBackgrounds(fondos);
    resumen.getRange(3, 6, valores.length, 1).setFontWeight("bold").setHorizontalAlignment("center");
    resumen.getRange(3, 10, valores.length, 2).setValues(infos);
    resumen.getRange(3, 11, valores.length, 1).setBackgrounds(fondosFamilia).setFontWeight("bold").setHorizontalAlignment("center");
  }
  var r = 3 + valores.length;
  resumen.getRange(r, 1, 1, 2).setValues([ [ "TOTAL", totalKg ] ]);
  resumen.getRange(r, 1, 1, 6).setFontWeight("bold").setBorder(true, false, false, false, false, false);
  if (materiales.length > 0) {
    resumen.getRange(3, 2, materiales.length, 1).setNumberFormat("#,##0.00");
    resumen.getRange(3, 4, materiales.length, 2).setNumberFormat("0.0%");
    var reglaRoja = SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(7).setBackground("#F4C7C3").setFontColor("#990000").setRanges([ resumen.getRange(3, 9, materiales.length, 1) ]).build();
    var reglaVerde = SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(7).setBackground("#D9EAD3").setFontColor("#274E13").setRanges([ resumen.getRange(3, 9, materiales.length, 1) ]).build();
    resumen.setConditionalFormatRules([ reglaRoja, reglaVerde ]);
  }
  resumen.getRange(1, 10).setValue("\xdaltima actualizaci\xf3n: " + Utilities.formatDate(new Date, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"));
  var topN = Math.min(10, materiales.length);
  if (topN > 0) {
    resumen.getCharts().forEach((function(ch) {
      resumen.removeChart(ch);
    }));
    var rangoGrafica = resumen.getRange(3, 1, topN, 2);
    var chart = resumen.newChart().setChartType(Charts.ChartType.BAR).addRange(rangoGrafica).setPosition(3, 12, 0, 0).setOption("title", "Top " + topN + " materiales por kg consumido").setOption("width", 620).setOption("height", 320).setOption("legend", {
      position: "none"
    }).setOption("colors", [ "#1F4E5F" ]).build();
    resumen.insertChart(chart);
  }
}

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

function leerInventarioParaDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var consMat = ss.getSheetByName("Consumo por Material");
  var resColor = ss.getSheetByName("Resumen Consumo");
  var materiales = [];
  if (consMat) {
    var lastRowMat = consMat.getLastRow();
    if (lastRowMat >= 3) {
      var datosMat = consMat.getRange(3, 1, lastRowMat - 2, 11).getValues();
      datosMat.forEach((function(r) {
        var nombre = String(r[0] || "").trim();
        if (!nombre || nombre === "TOTAL") return;
        materiales.push({
          nombre: nombre,
          kgConsumido: typeof r[1] === "number" ? r[1] : 0,
          existenciaActual: r[6],
          promedioDiario: r[7],
          diasCobertura: r[8],
          nombreDescriptivo: r[9] || nombre,
          familia: r[10] || ""
        });
      }));
    }
  }
  var colores = [];
  if (resColor) {
    var lastRowCol = resColor.getLastRow();
    if (lastRowCol >= 3) {
      var datosCol = resColor.getRange(3, 1, lastRowCol - 2, 11).getValues();
      datosCol.forEach((function(r) {
        var codigo = String(r[0] || "").trim();
        if (!codigo || r[1] === "TOTAL") return;
        colores.push({
          codigo: codigo,
          nombre: r[1],
          kgConsumido: typeof r[2] === "number" ? r[2] : 0,
          existencia: r[7],
          promedioDiario: r[8],
          diasCobertura: r[10]
        });
      }));
    }
  }
  function porUrgencia(a, b) {
    var da = typeof a.diasCobertura === "number" ? a.diasCobertura : Infinity;
    var db = typeof b.diasCobertura === "number" ? b.diasCobertura : Infinity;
    return da - db;
  }
  materiales.sort(porUrgencia);
  colores.sort(porUrgencia);
  var alertaMateriales = materiales.filter((function(m) {
    return typeof m.diasCobertura === "number" && m.diasCobertura < 7;
  })).length;
  var alertaColores = colores.filter((function(c) {
    return typeof c.diasCobertura === "number" && c.diasCobertura < 7;
  })).length;
  return {
    materiales: materiales,
    colores: colores,
    alertas: alertaMateriales + alertaColores
  };
}

function doGet(e) {
  var d = calcularDatosDashboard();
  var inv = leerInventarioParaDashboard();
  var ahora = Utilities.formatDate(new Date, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  var html = HtmlService.createTemplateFromFile("DashboardHTML");
  html.datos = d;
  html.inventario = inv;
  html.ahora = ahora;
  return html.evaluate().setTitle("Seguimiento Manaplas \u2014 Dashboard").addMetaTag("viewport", "width=device-width, initial-scale=1");
}
