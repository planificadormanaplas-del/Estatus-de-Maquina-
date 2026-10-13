/**
 * GENERADOR DE ÓRDENES DE PRODUCCIÓN — v22
 * -----------------------------------------
 * Novedades de esta versión (revisión de errores):
 *
 *  - NUEVO botón "🛠️ Reparar fórmulas del Maestro". En la hoja Maestro
 *    faltaban las fórmulas de TON, Producto, Componentes, Sacos, Sacos (kg),
 *    Ciclo, Peso, Cavidades y Und/Hora en:
 *      · la fila 7 (OP 1, color 6),
 *      · las filas 69-73 (OP 9, colores 4-8),
 *      · y en TODAS las filas de las OP 10 a la 30, excepto la primera.
 *    Por eso, al agregar un segundo color a esas órdenes, Sacos y Cantidad
 *    Final salían en blanco (y el Historial los guardaba vacíos). El botón
 *    vuelve a poner las fórmulas en las 240 filas, sin tocar lo que hayas
 *    escrito a mano. Además, ahora cada vez que se crea una orden, se agrega
 *    un color o se organizan las órdenes, se aseguran las fórmulas de ese
 *    bloque automáticamente.
 *  - Las fórmulas reparadas buscan en columnas completas ('Componentes '!B:N)
 *    en vez de rangos cortados como $B$2:$M1029, que se iban "corriendo" de
 *    fila en fila y dejaban fuera los componentes nuevos al final de la lista.
 *  - "Organizar" y "Borrar SIN Historial" ahora mueven también el "Tipo de
 *    Material" (G21) y las "Observaciones" (E34) de las hojas 01-14 junto con
 *    la orden. Antes, al recorrer la OP 3 a la OP 2, sus observaciones se
 *    quedaban pegadas en la hoja 03.
 *  - Al escoger la máquina, avisa si su estatus en "Maq y Tons" no es
 *    "Operativa" (ej: "Espera Logística").
 *  - Si escribes mal la cantidad de colores, te la vuelve a preguntar en vez
 *    de cancelar toda la orden (o, en "varias órdenes", poner 1 sin avisar).
 *  - "Gestionar → Agregar" ya no agrega filas a una OP vacía.
 *  - "Guardar copia en Drive" valida el día (que exista en ese mes) y, si
 *    eliges un mes que ya pasó hace más de 6 meses (ej: enero estando en
 *    diciembre), usa el año siguiente. Al final muestra el enlace a la copia.
 *  - Al borrar el SKU directo en la hoja, se borra también el Componente y
 *    su lista desplegable. Si pegas varios SKU de una vez, se procesan todas
 *    las filas pegadas (antes solo la primera).
 *  - Mucho más rápido: el Maestro se lee de una sola vez en lugar de celda
 *    por celda (antes "Organizar" o "Vaciar" hacían miles de lecturas).
 *  - Al limpiar un bloque también se quita la lista desplegable vieja del
 *    Componente.
 *
 * Novedades de versiones anteriores (se mantienen):
 *  v21: "Guardar copia de hoy en Drive" pregunta el MES.
 *  v20: botón "🎨 Colorear OP (por bloques)".
 *  v19: la Máquina se escribe en TODAS las filas de la orden.
 *  v18: los comandos de borrar limpian también "Observaciones".
 *  v17: un SKU "0" en el Maestro es "por asignar" (no cuenta como orden).
 *  v16: número de OP "0" o en blanco = automático.
 *  v15: el "Tipo de Material" se borra al generar una orden nueva.
 *  v14: botón "🗑️ Borrar TODO el Historial".
 *  v13: puedes elegir el número de OP.
 *  v10: el Componente se puede elegir por NÚMERO de la lista.
 *  v9:  los Componentes se leen de la hoja "Componentes " (B=código, C=Sku,
 *       E=nombre), comparando el Sku EXACTO.
 *  v8:  si un SKU no tiene componente propio, usa los genéricos "FAMILIA-0".
 *
 * Nada de esto mueve ni inserta filas reales, así que no rompe ningún
 * vínculo con las hojas 01-14.
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📋 Órdenes de Producción')
    .addItem('➕ Generar nueva orden', 'generarNuevaOrden')
    .addItem('➕➕ Generar varias órdenes de una vez', 'generarVariasOrdenes')
    .addItem('🔧 Gestionar esta orden (agregar / borrar)', 'gestionarOrden')
    .addItem('🩹 Borrar SIN guardar en Historial (para corregir)', 'borrarOrdenSinHistorial')
    .addItem('🔄 Organizar (cerrar huecos en la numeración)', 'organizarOrdenes')
    .addItem('🎨 Colorear OP (por bloques)', 'colorearOP')
    .addItem('🛠️ Reparar fórmulas del Maestro', 'repararFormulasMaestro')
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
var FILA_INICIO = 2;          // primera fila de datos (bloque 1)
var FILAS_POR_BLOQUE = 8;     // filas que ocupa cada orden
var TOTAL_BLOQUES = 30;       // cuántos bloques hay preparados
var HOJAS_DE_ORDEN = 14;      // OP 1-14 tienen hoja propia (01-14)
var COL_MAQ = 2;              // B = Máquina
var COL_SKU = 4;              // D = Sku
var COL_COMPONENTE = 5;       // E = SKu Componente
var COL_PRODUCTO = 6;         // F = Producto (fórmula)
var COL_MATERIAL = 7;         // G = Material
var COL_CANTIDAD = 9;         // I = Cantidad
var COL_COLOR = 10;           // J = Color
var COL_SACOS = 11;           // K = Sacos (fórmula)
var COL_ESTADO = 17;          // Q = Estado
var COL_IMPORTANTE = 18;      // R = ¿Producto importante? (columna oculta)
var COL_CANTIDAD_FINAL = 19;  // S = Cantidad final (fórmula)
var COL_COMPONENTES_CODIGO = 2;          // B en Componentes (código, ej: 0004-0/0)
var COL_COMPONENTES_SKU = 3;             // C en Componentes (a qué SKU pertenece)
var COL_COMPONENTES_NOMBRE = 5;          // E en Componentes (nombre completo)
var COL_LISTA_PRODUCTOS_SKU = 2;         // B en Lista de Productos (código)
var COL_LISTA_PRODUCTOS_NOMBRE = 3;      // C en Lista de Productos (nombre)
var COL_MAQ_ESTATUS = 6;                 // F en Maq y Tons (Estatus Actual)
var ESTATUS_OPERATIVA = 'Operativa';
// ----------------------------------------------------------

var TOTAL_FILAS = TOTAL_BLOQUES * FILAS_POR_BLOQUE;

// ---------- Utilidades de lectura en bloque ----------

// Lee de una sola vez las columnas A..S de un rango de filas del Maestro.
// Devuelve una matriz [fila][columna-1] (o sea, datos[i][COL_SKU - 1]).
function leerFilas(sheet, filaInicio, numFilas) {
  return sheet.getRange(filaInicio, 1, numFilas, COL_CANTIDAD_FINAL).getValues();
}

// Un SKU escrito como "0" es un marcador de "por asignar todavía" — no
// cuenta como dato real. Así, ni "Generar nueva orden" ni "Organizar" lo
// tratan como una orden ya existente; para ellos, esa fila sigue vacía.
function esPlaceholderSku(valor) {
  return valor === 0 || valor === '0';
}

// ¿Está vacío un bloque, según sus valores ya leídos (8 filas)?
function filasEstanVacias(filas) {
  for (var i = 0; i < filas.length; i++) {
    var sku = filas[i][COL_SKU - 1];
    var skuCuenta = (sku !== '' && !esPlaceholderSku(sku));
    if (filas[i][COL_CANTIDAD - 1] !== '' || filas[i][COL_COLOR - 1] !== '' || skuCuenta) {
      return false;
    }
  }
  return true;
}

function bloqueEstaVacio(sheet, filaInicio) {
  return filasEstanVacias(leerFilas(sheet, filaInicio, FILAS_POR_BLOQUE));
}

function filaDeBloque(numeroOP) {
  return FILA_INICIO + (numeroOP - 1) * FILAS_POR_BLOQUE;
}

function buscarBloqueLibre(sheet, excluir) {
  excluir = excluir || {};
  var datos = leerFilas(sheet, FILA_INICIO, TOTAL_FILAS);
  for (var b = 0; b < TOTAL_BLOQUES; b++) {
    var numero = b + 1;
    if (excluir[numero]) continue;
    var filas = datos.slice(b * FILAS_POR_BLOQUE, (b + 1) * FILAS_POR_BLOQUE);
    if (filasEstanVacias(filas)) {
      return { numero: numero, fila: filaDeBloque(numero) };
    }
  }
  return null;
}

// Pregunta qué número de OP usar para la orden nueva. Si lo dejan en
// blanco, elige automático el siguiente libre.
// Devuelve: {numero, fila} si hay uno válido, null si no quedan libres
// (modo automático), o undefined si cancelaron.
function pedirNumeroOP(ui, sheet, excluir) {
  excluir = excluir || {};
  while (true) {
    var resp = ui.prompt(
      'Número de orden (OP)',
      '¿Qué número de OP quieres usar? Escribe 0, o déjalo en blanco, para que el sistema elija ' +
      'automáticamente el siguiente libre (1 a ' + TOTAL_BLOQUES + ').',
      ui.ButtonSet.OK_CANCEL
    );
    if (resp.getSelectedButton() !== ui.Button.OK) return undefined;

    var texto = resp.getResponseText().trim();
    if (texto === '' || texto === '0') {
      return buscarBloqueLibre(sheet, excluir);
    }

    var numero = parseInt(texto, 10);
    if (isNaN(numero) || texto !== numero.toString() || numero < 1 || numero > TOTAL_BLOQUES) {
      ui.alert('Escribe un número entre 1 y ' + TOTAL_BLOQUES + ', o 0 (o déjalo en blanco) para automático.');
      continue;
    }

    if (excluir[numero]) {
      ui.alert('Ya usaste la OP ' + numero + ' en esta misma tanda. Elige otro número.');
      continue;
    }

    var fila = filaDeBloque(numero);
    if (!bloqueEstaVacio(sheet, fila)) {
      ui.alert('La OP ' + numero + ' ya tiene datos. Elige otro número, o bórrala primero si quieres reusarla.');
      continue;
    }

    return { numero: numero, fila: fila };
  }
}

function desplegarBloque(sheet, filaBloque, n) {
  sheet.showRows(filaBloque, n);
  if (n < FILAS_POR_BLOQUE) {
    sheet.hideRows(filaBloque + n, FILAS_POR_BLOQUE - n);
  }
}

// Columnas de "entrada" que definen una orden — todo lo demás (Producto,
// Componentes calculado, Sacos, Ciclo, etc.) son fórmulas que se recalculan
// solas en cuanto estas queden en su nueva posición.
var COLUMNAS_BLOQUE = [COL_MAQ, COL_SKU, COL_COMPONENTE, COL_MATERIAL, COL_CANTIDAD, COL_COLOR, COL_ESTADO, COL_IMPORTANTE];

function limpiarBloque(sheet, filaBloque) {
  for (var c = 0; c < COLUMNAS_BLOQUE.length; c++) {
    sheet.getRange(filaBloque, COLUMNAS_BLOQUE[c], FILAS_POR_BLOQUE, 1).clearContent();
  }
  sheet.getRange(filaBloque, COL_COMPONENTE, FILAS_POR_BLOQUE, 1).clearDataValidations();
  sheet.hideRows(filaBloque, FILAS_POR_BLOQUE);
}

// ---------- Fórmulas del Maestro ----------

// Fórmulas "oficiales" de cada columna calculada. {r} se reemplaza por el
// número de fila. Van en sintaxis de Apps Script (comas, nombres en inglés);
// Google Sheets las muestra después con punto y coma, como siempre.
// Son las mismas que ya tenía la hoja, pero buscando en columnas completas.
var FORMULAS_MAESTRO = {
  3:  '=IFERROR(VLOOKUP(B{r},\'Maq y Tons \'!$A$3:$F,3,FALSE),"")',                                   // C = TON
  6:  '=IFERROR(VLOOKUP(D{r},\'Lista de Productos\'!$B:$D,2,FALSE)," ")',                              // F = PRODUCTO
  8:  '=IFERROR(UPPER(VLOOKUP(E{r},\'Componentes \'!$B:$F,5,FALSE))," ")',                             // H = COMPONENTES
  11: '=IF(I{r}="","",IFERROR(IF(R{r}="Sí",ROUNDUP(N{r}*I{r}/25000,0),ROUND(N{r}*I{r}/25000,0))," "))', // K = SACOS
  12: '=IF(K{r}="","",IFERROR(K{r}*25," "))',                                                          // L = SACOS (kg)
  13: '=IFERROR(VLOOKUP(E{r},\'Componentes \'!$B:$N,12,FALSE)," ")',                                  // M = CICLO
  14: '=IFERROR(VLOOKUP(E{r},\'Componentes \'!$B:$N,10,FALSE),"")',                                    // N = PESO
  15: '=IFERROR(VLOOKUP(E{r},\'Componentes \'!$B:$N,13,FALSE),"")',                                    // O = CAV.
  16: '=IFERROR((3600*O{r})/M{r},"")',                                                                // P = UND/HORA
  19: '=IF(I{r}="","",IFERROR(MROUND(K{r}*25000/N{r},10),""))',                                        // S = CANTIDAD FINAL
  20: '=IFERROR(VLOOKUP(E{r},Caracteristicas!$C:$AD,12,0),"")'                                        // T = MATERIAL RECOMENDADO
};

// Pone las fórmulas oficiales en las filas indicadas. Si una celda tiene un
// valor escrito a mano (no fórmula), se respeta y no se toca.
// Devuelve cuántas celdas se arreglaron (estaban vacías) y cuántas se
// respetaron por tener un valor a mano.
function asegurarFormulas(sheet, filaInicio, numFilas) {
  var reparadas = 0;
  var manuales = [];
  for (var colTexto in FORMULAS_MAESTRO) {
    var col = parseInt(colTexto, 10);
    var plantilla = FORMULAS_MAESTRO[colTexto];
    var rango = sheet.getRange(filaInicio, col, numFilas, 1);
    var formulas = rango.getFormulas();
    var valores = rango.getValues();
    var nuevos = [];
    var cambia = false;
    for (var i = 0; i < numFilas; i++) {
      var fila = filaInicio + i;
      if (formulas[i][0] === '' && valores[i][0] !== '' && valores[i][0] !== ' ') {
        // Valor escrito a mano: no lo pisamos.
        nuevos.push([valores[i][0]]);
        manuales.push(rango.getCell(i + 1, 1).getA1Notation());
        continue;
      }
      if (formulas[i][0] === '') {
        reparadas++;
        cambia = true;
      }
      nuevos.push([plantilla.replace(/\{r\}/g, fila)]);
    }
    if (cambia) rango.setValues(nuevos);
  }
  return { reparadas: reparadas, manuales: manuales };
}

function asegurarFormulasBloque(sheet, filaBloque) {
  asegurarFormulas(sheet, filaBloque, FILAS_POR_BLOQUE);
}

function repararFormulasMaestro() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(HOJA_MAESTRO);
  if (!sheet) { ui.alert('No encontré la hoja "' + HOJA_MAESTRO + '".'); return; }

  var confirmar = ui.alert(
    '🛠️ Reparar fórmulas del Maestro',
    'Esto revisa las ' + TOTAL_FILAS + ' filas de las ' + TOTAL_BLOQUES + ' órdenes y vuelve a poner las fórmulas ' +
    'de TON, Producto, Componentes, Sacos, Sacos (kg), Ciclo, Peso, Cavidades, Und/Hora, Cantidad Final y ' +
    'Material Recomendado donde falten.\n\n' +
    'No toca Máquina, Sku, Componente, Material, Cantidad, Color, Estado ni "¿Producto importante?". ' +
    'Si alguna celda calculada tiene un número escrito a mano, se respeta. ¿Continuar?',
    ui.ButtonSet.YES_NO
  );
  if (confirmar !== ui.Button.YES) return;

  var resultado = asegurarFormulas(sheet, FILA_INICIO, TOTAL_FILAS);
  var mensaje = resultado.reparadas === 0
    ? 'No faltaba ninguna fórmula. 👍'
    : ('Se pusieron ' + resultado.reparadas + ' fórmulas que faltaban.');
  if (resultado.manuales.length > 0) {
    var muestra = resultado.manuales.slice(0, 20).join(', ');
    var extra = resultado.manuales.length > 20 ? (' … y ' + (resultado.manuales.length - 20) + ' más') : '';
    mensaje += '\n\nEstas celdas tienen un valor escrito a mano en lugar de fórmula y NO se tocaron ' +
               '(revísalas): ' + muestra + extra;
  }
  ui.alert('🛠️ Listo', mensaje, ui.ButtonSet.OK);
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

// Devuelve un número entre 1 y FILAS_POR_BLOQUE, o null si cancelaron.
// Si escriben algo inválido, vuelve a preguntar.
function pedirNumeroColores(ui, titulo) {
  var mensaje = '¿Cuántos colores va a tener esta orden? (1 a ' + FILAS_POR_BLOQUE + ')';
  while (true) {
    var resp = ui.prompt(titulo, mensaje, ui.ButtonSet.OK_CANCEL);
    if (resp.getSelectedButton() !== ui.Button.OK) return null;
    var texto = resp.getResponseText().trim();
    var n = parseInt(texto, 10);
    if (!isNaN(n) && texto === n.toString() && n >= 1 && n <= FILAS_POR_BLOQUE) return n;
    mensaje = 'Escribe solo un número entre 1 y ' + FILAS_POR_BLOQUE + '.\n\n¿Cuántos colores va a tener esta orden?';
  }
}

// ---------- Menú: generar orden(es) ----------

// Hace todas las preguntas de una orden y la escribe. Devuelve un texto de
// resumen, null si cancelaron, o false si ya no quedan bloques libres.
function asistenteUnaOrden(ui, ss, sheet, excluir, tituloColores) {
  var bloque = pedirNumeroOP(ui, sheet, excluir);
  if (bloque === undefined) return null;
  if (!bloque) return false;

  var maquina = pedirMaquina(ui);
  if (maquina === null) return null;

  var sku = pedirSku(ui);
  if (sku === null) return null;

  var componente = pedirComponente(ui, sku);
  if (componente === null) return null;

  var importante = pedirImportante(ui);

  var n = pedirNumeroColores(ui, tituloColores + ' (OP ' + bloque.numero + ')');
  if (n === null) return null;

  aplicarOrdenAlBloque(sheet, bloque.fila, n, maquina, sku, componente, importante);
  limpiarTipoDeMaterialDeHoja(ss, bloque.numero);
  desplegarBloque(sheet, bloque.fila, n);
  excluir[bloque.numero] = true;

  return {
    numero: bloque.numero,
    fila: bloque.fila,
    texto: 'OP ' + bloque.numero + ' (máquina ' + maquina + ', SKU ' + sku + '): ' + n + ' fila(s)'
  };
}

function generarNuevaOrden() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(HOJA_MAESTRO);
  if (!sheet) { ui.alert('No encontré la hoja "' + HOJA_MAESTRO + '".'); return; }

  var r = asistenteUnaOrden(ui, ss, sheet, {}, 'Nueva orden');
  if (r === false) { ui.alert('No quedan bloques libres en el Maestro.'); return; }
  if (!r) return;

  sheet.setActiveRange(sheet.getRange(r.fila, COL_COLOR));
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
  var bloquesUsados = {};
  for (var k = 0; k < totalOrdenes; k++) {
    ui.alert('Orden ' + (k + 1) + ' de ' + totalOrdenes);
    var r = asistenteUnaOrden(ui, ss, sheet, bloquesUsados, 'Orden ' + (k + 1) + ' de ' + totalOrdenes);
    if (r === false) {
      ui.alert('Ya no quedan bloques libres. Se generaron ' + k + ' de ' + totalOrdenes + ' órdenes.');
      break;
    }
    if (!r) break;
    resumen.push(r.texto);
  }

  if (resumen.length > 0) {
    ui.alert('Listo:\n\n' + resumen.join('\n'));
  }
}

function aplicarOrdenAlBloque(sheet, filaBloque, n, maquina, sku, componente, importante) {
  asegurarFormulasBloque(sheet, filaBloque);

  var columna = function (valor) {
    var m = [];
    for (var i = 0; i < n; i++) m.push([valor]);
    return m;
  };
  sheet.getRange(filaBloque, COL_MAQ, n, 1).setValues(columna(maquina));
  sheet.getRange(filaBloque, COL_SKU, n, 1).setValues(columna(sku));
  if (componente) {
    var rangoComp = sheet.getRange(filaBloque, COL_COMPONENTE, n, 1);
    aplicarValidacionComponente(rangoComp, opcionesDeComponente(sku));
    rangoComp.setValues(columna(componente));
  }
  if (importante) {
    sheet.getRange(filaBloque, COL_IMPORTANTE, n, 1).setValues(columna(importante));
  }
}

// ---------- Menú: gestionar una orden existente ----------

// Devuelve {numeroOP, filaBloque} según la celda activa del Maestro, o null
// (después de avisar) si no estás en el lugar correcto.
function bloqueDeCeldaActiva(ui, sheet, accion) {
  if (sheet.getName() !== HOJA_MAESTRO) {
    ui.alert('Primero ponte en una celda dentro del Maestro, en la orden que quieres ' + accion + '.');
    return null;
  }
  var filaActiva = sheet.getActiveCell().getRow();
  var indiceBloque = Math.floor((filaActiva - FILA_INICIO) / FILAS_POR_BLOQUE);
  if (filaActiva < FILA_INICIO || indiceBloque >= TOTAL_BLOQUES) {
    ui.alert('Ponte en una fila dentro de alguna orden.');
    return null;
  }
  return { numeroOP: indiceBloque + 1, filaBloque: FILA_INICIO + indiceBloque * FILAS_POR_BLOQUE };
}

function gestionarOrden() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();

  var pos = bloqueDeCeldaActiva(ui, sheet, 'gestionar');
  if (!pos) return;
  var filaBloque = pos.filaBloque;
  var numeroOP = pos.numeroOP;

  var resp = ui.alert(
    'Orden OP ' + numeroOP,
    '¿Qué quieres hacer con esta orden?\n\n' +
    'Sí → Agregar otro color/producto a esta orden\n' +
    'No → Borrar esta orden completa\n' +
    'Cancelar → No hacer nada',
    ui.ButtonSet.YES_NO_CANCEL
  );

  if (resp === ui.Button.YES) {
    var filas = leerFilas(sheet, filaBloque, FILAS_POR_BLOQUE);
    if (filas[0][COL_SKU - 1] === '' && filas[0][COL_MAQ - 1] === '') {
      ui.alert('La OP ' + numeroOP + ' está vacía. Usa "➕ Generar nueva orden" para crearla.');
      return;
    }
    var visibles = 0;
    for (var i = 0; i < FILAS_POR_BLOQUE; i++) {
      if (!sheet.isRowHiddenByUser(filaBloque + i)) visibles++;
    }
    if (visibles >= FILAS_POR_BLOQUE) {
      ui.alert('Esta orden ya tiene visibles las ' + FILAS_POR_BLOQUE + ' filas disponibles.');
      return;
    }
    var filaNueva = filaBloque + visibles;
    var maquina = filas[0][COL_MAQ - 1];
    var sku = filas[0][COL_SKU - 1];
    var componente = filas[0][COL_COMPONENTE - 1];
    var importante = filas[0][COL_IMPORTANTE - 1];

    asegurarFormulasBloque(sheet, filaBloque);
    if (maquina) sheet.getRange(filaNueva, COL_MAQ).setValue(maquina);
    if (sku) sheet.getRange(filaNueva, COL_SKU).setValue(sku);
    if (componente) {
      var celdaComp = sheet.getRange(filaNueva, COL_COMPONENTE);
      aplicarValidacionComponente(celdaComp, opcionesDeComponente(sku.toString()));
      celdaComp.setValue(componente);
    }
    if (importante) sheet.getRange(filaNueva, COL_IMPORTANTE).setValue(importante);
    sheet.showRows(filaNueva, 1);
    sheet.setActiveRange(sheet.getRange(filaNueva, COL_COLOR));
    ss.toast('Se agregó una fila más a la orden OP ' + numeroOP + '. Ya tiene el mismo SKU/Componente, solo falta Cantidad y Color.', '✅', 8);

  } else if (resp === ui.Button.NO) {
    var confirmar = ui.alert(
      '🗑️ Borrar orden OP ' + numeroOP,
      'Esto va a borrar todos los datos de esta orden y va a esconder sus filas otra vez.\n\nEsta acción no se puede deshacer. ¿Seguro?',
      ui.ButtonSet.YES_NO
    );
    if (confirmar === ui.Button.YES) {
      archivarBloque(sheet, filaBloque, numeroOP, 'Borrado individual');
      limpiarBloque(sheet, filaBloque);
      limpiarTipoDeMaterialDeHoja(ss, numeroOP);
      limpiarObservacionesDeHoja(ss, numeroOP);
      ss.toast('La orden OP ' + numeroOP + ' quedó guardada en el Historial y borrada del Maestro.', '🗑️ Listo', 6);
    }
  }
}

// ---------- Borrar sin dejar rastro en el Historial (para corregir errores) ----------

function borrarOrdenSinHistorial() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();

  var pos = bloqueDeCeldaActiva(ui, sheet, 'borrar');
  if (!pos) return;
  var numeroOP = pos.numeroOP;

  var confirmar = ui.alert(
    '🩹 Borrar orden OP ' + numeroOP + ' (sin Historial)',
    'Esto borra los datos de esta orden SIN guardar copia en el Historial, y recorre las órdenes ' +
    'siguientes hacia arriba para que no quede un hueco en la numeración.\n\n' +
    'Úsalo solo para corregir un error mientras la estás llenando (SKU equivocado, etc). ' +
    'Si la orden ya estaba terminada o en proceso de verdad, mejor usa "🔧 Gestionar esta orden" en vez de este.\n\n' +
    '¿Seguro que quieres continuar?',
    ui.ButtonSet.YES_NO
  );
  if (confirmar !== ui.Button.YES) return;

  limpiarBloque(sheet, pos.filaBloque);
  limpiarTipoDeMaterialDeHoja(ss, numeroOP);
  limpiarObservacionesDeHoja(ss, numeroOP);
  var movimientos = compactarBloques(ss, sheet);
  var extra = movimientos.length > 0 ? (' Se recorrieron ' + movimientos.length + ' orden(es) para cerrar el hueco.') : '';
  ss.toast('La orden OP ' + numeroOP + ' quedó borrada (sin registrar en Historial).' + extra, '🩹 Listo', 8);
}

// ---------- Organizar: cerrar huecos en la numeración de las órdenes ----------

function organizarOrdenes() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(HOJA_MAESTRO);
  if (!sheet) { ui.alert('No encontré la hoja "' + HOJA_MAESTRO + '".'); return; }

  var confirmar = ui.alert(
    '🔄 Organizar órdenes',
    'Esto recorre hacia arriba las órdenes que tengan datos, para cerrar los huecos que hayan ' +
    'quedado en la numeración (ej: si la OP 2 está vacía y la OP 3 tiene datos, la OP 3 pasa a ser la OP 2).\n\n' +
    'El "Tipo de Material" y las "Observaciones" de las hojas 01-14 se mueven junto con su orden.\n\n' +
    'Los números de las órdenes que se muevan van a cambiar. Esta acción no se puede deshacer. ¿Seguro?',
    ui.ButtonSet.YES_NO
  );
  if (confirmar !== ui.Button.YES) return;

  var movimientos = compactarBloques(ss, sheet);

  if (movimientos.length === 0) {
    ui.alert('No había huecos que cerrar — ya estaba todo organizado.');
  } else {
    ui.alert('Listo, se movieron ' + movimientos.length + ' orden(es):\n\n' + movimientos.join('\n'));
  }
}

// Lee todos los bloques, junta los que tienen datos al principio (en su
// mismo orden relativo) y deja los vacíos al final. También mueve el "Tipo
// de Material" y las "Observaciones" de las hojas 01-14 con su orden.
// Devuelve la lista de movimientos hechos (para avisar al usuario).
function compactarBloques(ss, sheet) {
  var datos = leerFilas(sheet, FILA_INICIO, TOTAL_FILAS);
  var bloques = [];
  for (var b = 0; b < TOTAL_BLOQUES; b++) {
    var fila = FILA_INICIO + b * FILAS_POR_BLOQUE;
    var filas = datos.slice(b * FILAS_POR_BLOQUE, (b + 1) * FILAS_POR_BLOQUE);
    var visibles = 0;
    for (var i = 0; i < FILAS_POR_BLOQUE; i++) {
      if (!sheet.isRowHiddenByUser(fila + i)) visibles++;
    }
    bloques.push({
      indice: b,
      tieneDatos: !filasEstanVacias(filas),
      visibles: Math.max(visibles, 1),
      filas: filas,
      extras: leerExtrasDeHoja(ss, b + 1)
    });
  }

  var conDatos = bloques.filter(function (x) { return x.tieneDatos; });
  var vacios = bloques.filter(function (x) { return !x.tieneDatos; });
  var nuevoOrden = conDatos.concat(vacios);

  var movimientos = [];
  for (var pos = 0; pos < nuevoOrden.length; pos++) {
    if (nuevoOrden[pos].indice !== pos) {
      movimientos.push('OP ' + (nuevoOrden[pos].indice + 1) + '  →  OP ' + (pos + 1));
    }
  }
  if (movimientos.length === 0) return [];

  // Armar las columnas de entrada ya en el orden nuevo y escribirlas de una vez.
  for (var c = 0; c < COLUMNAS_BLOQUE.length; c++) {
    var col = COLUMNAS_BLOQUE[c];
    var nuevaColumna = [];
    for (var p = 0; p < nuevoOrden.length; p++) {
      for (var f = 0; f < FILAS_POR_BLOQUE; f++) {
        nuevaColumna.push([nuevoOrden[p].tieneDatos ? nuevoOrden[p].filas[f][col - 1] : '']);
      }
    }
    sheet.getRange(FILA_INICIO, col, TOTAL_FILAS, 1).setValues(nuevaColumna);
  }

  // Listas de Componente, fórmulas, filas visibles y hojas 01-14.
  sheet.getRange(FILA_INICIO, COL_COMPONENTE, TOTAL_FILAS, 1).clearDataValidations();
  sheet.hideRows(FILA_INICIO, TOTAL_FILAS);
  for (var p2 = 0; p2 < nuevoOrden.length; p2++) {
    var origen = nuevoOrden[p2];
    var filaDestino = FILA_INICIO + p2 * FILAS_POR_BLOQUE;
    if (origen.tieneDatos) {
      var skuBloque = origen.filas[0][COL_SKU - 1];
      if (skuBloque) {
        aplicarValidacionComponente(
          sheet.getRange(filaDestino, COL_COMPONENTE, FILAS_POR_BLOQUE, 1),
          opcionesDeComponente(skuBloque.toString())
        );
      }
      asegurarFormulasBloque(sheet, filaDestino);
      desplegarBloque(sheet, filaDestino, origen.visibles);
    }
    if (origen.indice !== p2) {
      escribirExtrasEnHoja(ss, p2 + 1, origen.tieneDatos ? origen.extras : null);
    }
  }

  return movimientos;
}

// ---------- Colorear la columna OP por bloques (azul/gris alternado) ----------

var COLOR_OP_A = '#D9E8F5'; // azul claro
var COLOR_OP_B = '#EDEDED'; // gris claro

function colorearOP() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(HOJA_MAESTRO);
  if (!sheet) { ui.alert('No encontré la hoja "' + HOJA_MAESTRO + '".'); return; }

  var colores = [];
  for (var b = 0; b < TOTAL_BLOQUES; b++) {
    var color = (b % 2 === 0) ? COLOR_OP_A : COLOR_OP_B;
    for (var i = 0; i < FILAS_POR_BLOQUE; i++) colores.push([color]);
  }
  sheet.getRange(FILA_INICIO, 1, TOTAL_FILAS, 1).setBackgrounds(colores);

  ss.toast('La columna OP quedó coloreada por bloques (azul/gris alternado).', '🎨 Listo', 5);
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

// Convierte las filas con datos de un bloque (ya leídas) en filas de Historial.
function filasParaHistorial(filas, numeroOP, motivo, fecha) {
  var maquina = filas[0][COL_MAQ - 1];
  var estado = filas[0][COL_ESTADO - 1];
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

// Guarda una "foto" de cada fila con datos de un bloque antes de borrarlo.
// No guarda nada si el bloque en realidad ya estaba vacío.
function archivarBloque(sheet, filaBloque, numeroOP, motivo) {
  var filas = leerFilas(sheet, filaBloque, FILAS_POR_BLOQUE);
  guardarEnHistorial(filasParaHistorial(filas, numeroOP, motivo, new Date()));
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
  var copia = ss.copy(nombre);

  ui.alert(
    '✅ Copia guardada en Drive',
    'Se creó la copia con el nombre:\n\n"' + nombre + '"\n\n' +
    'Ya está guardada en tu Google Drive (en la misma carpeta que este archivo). ' +
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
    'Esto va a borrar TODOS los datos de TODAS las órdenes, va a esconder todas las filas otra vez, ' +
    'y va a limpiar el campo "Tipo de Material" y la caja de "Observaciones" en las 14 hojas de orden (01-14).\n\n' +
    'Las órdenes con datos se guardan antes en el Historial.\n\n' +
    'Esta acción no se puede deshacer. ¿Seguro que quieres continuar?',
    ui.ButtonSet.YES_NO
  );
  if (confirmacion !== ui.Button.YES) return;

  var datos = leerFilas(sheet, FILA_INICIO, TOTAL_FILAS);
  var fecha = new Date();
  var paraHistorial = [];
  for (var b = 0; b < TOTAL_BLOQUES; b++) {
    var filas = datos.slice(b * FILAS_POR_BLOQUE, (b + 1) * FILAS_POR_BLOQUE);
    paraHistorial = paraHistorial.concat(filasParaHistorial(filas, b + 1, 'Vaciado masivo', fecha));
  }
  guardarEnHistorial(paraHistorial);

  for (var c = 0; c < COLUMNAS_BLOQUE.length; c++) {
    sheet.getRange(FILA_INICIO, COLUMNAS_BLOQUE[c], TOTAL_FILAS, 1).clearContent();
  }
  sheet.getRange(FILA_INICIO, COL_COMPONENTE, TOTAL_FILAS, 1).clearDataValidations();
  sheet.hideRows(FILA_INICIO, TOTAL_FILAS);

  limpiarTipoDeMaterial(ss);
  limpiarObservaciones(ss);

  ss.toast('Las órdenes que tenían datos quedaron guardadas en el Historial. El Maestro quedó vacío, ' +
           'todas las filas ocultas, y el "Tipo de Material" y "Observaciones" de las 14 hojas quedaron en blanco.', '🧹 Listo', 8);
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

// ---------- Automático: filtrar Componentes según el SKU (si editas directo en la hoja) ----------

function onEdit(e) {
  var rango = e.range;
  var sheet = rango.getSheet();
  if (sheet.getName() !== HOJA_MAESTRO) return;

  var colIni = rango.getColumn();
  var colFin = rango.getLastColumn();
  var tocaSku = colIni <= COL_SKU && COL_SKU <= colFin;
  var tocaComp = colIni <= COL_COMPONENTE && COL_COMPONENTE <= colFin;
  if (!tocaSku && !tocaComp) return;

  var filaIni = Math.max(rango.getRow(), FILA_INICIO);
  var filaFin = Math.min(rango.getLastRow(), FILA_INICIO + TOTAL_FILAS - 1);
  if (filaFin - filaIni > 50) return; // pegados enormes: no lo hacemos en vivo

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
