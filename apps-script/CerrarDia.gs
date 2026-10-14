/**
 * Reporte de Producción - botón "Cerrar día".
 *
 * Copia el Reporte Diario (pestaña REP PROD) a la pestaña REP HISTORICO como
 * valores fijos y luego limpia la columna de Piezas Malas para el día siguiente.
 *
 * Instalación:
 *   1. En el Google Sheet: Extensiones > Apps Script.
 *   2. Pegar este archivo completo y guardar.
 *   3. Recargar el Sheet. Aparece el menú "Producción" > "Cerrar día".
 *   4. (Opcional) Para un botón: Insertar > Dibujo, dibujar el botón,
 *      clic en los tres puntos del dibujo > Asignar secuencia de comandos > cerrarDia
 */

const HOJA_DIARIO = 'REP PROD';
const HOJA_HISTORICO = 'REP HISTORICO';
const CELDA_FECHA = 'B2';
const PRIMERA_FILA = 5;      // primera fila de datos del Reporte Diario
const NUM_COLUMNAS = 22;     // columnas A:V (Fecha ... Enlace)
const COL_ESTADO = 6;        // F
const COL_PIEZAS_MALAS = 14; // N

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Producción')
    .addItem('Cerrar día', 'cerrarDia')
    .addToUi();
}

function cerrarDia() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const diario = ss.getSheetByName(HOJA_DIARIO);
  const historico = ss.getSheetByName(HOJA_HISTORICO);
  if (!diario || !historico) {
    ui.alert('No se encontraron las pestañas "' + HOJA_DIARIO + '" y "' + HOJA_HISTORICO + '".');
    return;
  }

  const fecha = diario.getRange(CELDA_FECHA).getValue();
  if (!(fecha instanceof Date)) {
    ui.alert('La fecha en ' + HOJA_DIARIO + '!' + CELDA_FECHA + ' no es válida.');
    return;
  }
  const tz = ss.getSpreadsheetTimeZone();
  const fechaTxt = Utilities.formatDate(fecha, tz, 'dd/MM/yyyy');

  // Solo se guardan las filas de máquinas que tuvieron registro ese día y turno.
  const ultimaFila = diario.getLastRow();
  if (ultimaFila < PRIMERA_FILA) {
    ui.alert('El Reporte Diario está vacío.');
    return;
  }
  const numFilas = ultimaFila - PRIMERA_FILA + 1;
  const filas = diario.getRange(PRIMERA_FILA, 1, numFilas, NUM_COLUMNAS).getValues()
    .filter(function (f) { return f[1] !== '' && f[COL_ESTADO - 1] !== 'Sin registro'; });
  if (filas.length === 0) {
    ui.alert('No hay registros para el ' + fechaTxt + '.');
    return;
  }

  // Si el día ya se cerró antes, se pregunta si se reemplaza.
  const filasPrevias = filasDeFecha_(historico, fechaTxt, tz);
  if (filasPrevias.length > 0) {
    const resp = ui.alert('El ' + fechaTxt + ' ya está cerrado',
      '¿Quieres reemplazarlo con los datos actuales del Reporte Diario?', ui.ButtonSet.YES_NO);
    if (resp !== ui.Button.YES) return;
    // Se borran de abajo hacia arriba para no mover las filas pendientes.
    filasPrevias.reverse().forEach(function (n) { historico.deleteRow(n); });
  } else {
    const resp = ui.alert('Cerrar día ' + fechaTxt,
      'Se guardarán ' + filas.length + ' filas en ' + HOJA_HISTORICO +
      ' y se limpiarán las Piezas Malas. ¿Continuar?', ui.ButtonSet.YES_NO);
    if (resp !== ui.Button.YES) return;
  }

  const cerradoEl = new Date();
  const salida = filas.map(function (f) { return f.concat([cerradoEl]); });
  historico.getRange(historico.getLastRow() + 1, 1, salida.length, NUM_COLUMNAS + 1).setValues(salida);

  diario.getRange(PRIMERA_FILA, COL_PIEZAS_MALAS, numFilas, 1).clearContent();

  ui.alert('Día ' + fechaTxt + ' cerrado: ' + salida.length + ' filas guardadas en ' + HOJA_HISTORICO + '.');
}

/** Números de fila (1-based) del histórico cuya fecha coincide con fechaTxt. */
function filasDeFecha_(historico, fechaTxt, tz) {
  const ultima = historico.getLastRow();
  if (ultima < 2) return [];
  const fechas = historico.getRange(2, 1, ultima - 1, 1).getValues();
  const filas = [];
  fechas.forEach(function (f, i) {
    if (f[0] instanceof Date && Utilities.formatDate(f[0], tz, 'dd/MM/yyyy') === fechaTxt) {
      filas.push(i + 2);
    }
  });
  return filas;
}
