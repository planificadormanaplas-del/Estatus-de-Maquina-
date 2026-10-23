/**
 * MANAPLAS · PLANIFICACIÓN DE TURNOS DE PRODUCCIÓN
 * -------------------------------------------------
 * Google Apps Script para un Google Sheet.
 *
 * Instalación:
 *   1. Abre un Google Sheet nuevo > Extensiones > Apps Script.
 *   2. Borra el contenido, pega este archivo completo y guarda.
 *   3. Ejecuta la función configurarArchivo() y autoriza los permisos.
 *   4. Vuelve a la hoja y RECÁRGALA (F5).
 *
 * Hojas que crea:
 *   Turno 1 · Turno 2 · Turno 3 · Turno 4  → plano de máquinas + tabla de producción
 *   Planificación                           → base de datos (una fila por fecha/turno/máquina)
 *   Listas (oculta)                         → listado de máquinas para validaciones
 *
 * Uso:
 *   · Casilla de cada máquina    → la marca como activa (trabaja ese turno). Se pinta al instante.
 *   · «Marcar / desmarcar todas» → marca o desmarca todas las máquinas del turno.
 *   · Tabla de producción        → una fila por máquina activa, en orden de número. Lo que se
 *                                  escribe en la tabla se guarda solo en Planificación.
 */

// ======================= CONFIGURACIÓN =======================
const META_MAQUINAS = 20;          // máquinas que deberían trabajar por turno
const HOJA_PLAN = 'Planificación';
const HOJA_LISTAS = 'Listas';

const TURNOS = {
  1: { hoja: 'Turno 1', horario: '6:00 AM – 2:30 PM',  descanso: '1 h entre 9:30 AM y 12:00 PM', horas: 7.5, inicioMin: 360,  cambioMolde: false, color: '#f4a261' },
  2: { hoja: 'Turno 2', horario: '2:30 PM – 10:30 PM', descanso: '1 h entre 5:00 PM y 7:00 PM',  horas: 7,   inicioMin: 870,  cambioMolde: false, color: '#2a9d8f' },
  3: { hoja: 'Turno 3', horario: '10:30 PM – 6:00 AM', descanso: '1 h entre 2:00 AM y 3:00 AM',  horas: 6.5, inicioMin: 1350, cambioMolde: false, color: '#264653' },
  4: { hoja: 'Turno 4', horario: '7:00 AM – 4:00 PM',  descanso: '1 h de 12:00 PM a 1:00 PM',    horas: 8,   inicioMin: 420,  cambioMolde: true,  color: '#1f6fd1' },
};

// Columnas de la hoja Planificación (A:M)
const COLS_PLAN = ['Fecha', 'Turno', 'Máquina', 'Producto', 'Referencia', 'Ciclo ideal (s)', 'Ciclo real (s)',
  'Contador inicial', 'Contador final', 'Cavidades', 'Color', 'Material', 'N° operarios'];
const IDX = { fecha: 0, turno: 1, maquina: 2, producto: 3, referencia: 4, cicloIdeal: 5, cicloReal: 6,
  contIni: 7, contFin: 8, cavidades: 9, color: 10, material: 11, operarios: 12 };

const COLOR = {
  inactiva: '#eef1f5', activa: '#b9e4c1', alerta: '#ffc98a',
  borde: '#1f2a37', texto: '#1f2a37', etiqueta: '#5b6676',
  cabecera: '#1f2a37', entrada: '#fff8dc', fijo: '#f1f3f6', lineaTabla: '#d0d5dd',
};

// ======================= CUADRÍCULA =======================
// Columna A = margen. El plano usa "unidades" de 12 px desde la columna B.
const ANCHO_UNIDAD = 12;
const COL0 = 2;
const U = u => COL0 + u - 1;        // zona principal: unidades 1..71
const D = u => U(73) + u;           // zona derecha (Soplado y 125–143): unidades 1..11
const ULTIMA_COL = D(11) + 1;       // columna de margen derecho
const FILA_PLANO_INI = 6, FILA_PLANO_FIN = 32;
const FILA_PLANO_PRINCIPAL_FIN = 17; // última fila del plano en la zona principal (101–124)

const CELDA = {
  fechaSel: [2, U(5)], activas: [2, U(63)], horas: [3, U(44)], todas: [3, U(55)],
  alertas: [33, 1],   // celda auxiliar (texto invisible) con las máquinas en alerta: "|203|215|"
};

/**
 * Distribución de las máquinas según la hoja resumen de asignación de personal.
 * Cada máquina se divide en dos zonas:
 *   num → número de la máquina
 *   chk → casilla para marcarla como activa
 * Bloques de 2 o más filas: la casilla ocupa la última fila. Bloques de 1 fila: las 2 primeras columnas.
 */
function maquinas_() {
  const m = [];
  const add = (id, fila, col, alto, ancho) => {
    const x = { id: id, fila: fila, col: col, alto: alto, ancho: ancho };
    if (alto > 1) {
      x.num = { fila: fila, col: col, alto: alto - 1, ancho: ancho };
      x.chk = { fila: fila + alto - 1, col: col, alto: 1, ancho: ancho };
    } else {
      x.chk = { fila: fila, col: col, alto: 1, ancho: 2 };
      x.num = { fila: fila, col: col + 2, alto: 1, ancho: ancho - 2 };
    }
    m.push(x);
  };
  // Bloque 201–218
  add(201, 6, U(1), 1, 11);
  for (let j = 0; j < 13; j++) add(202 + j, 8, U(4 * j + 1), 4, 3);   // 202–214
  add(218, 6, U(62), 1, 10);
  add(217, 8, U(69), 2, 3);
  add(216, 10, U(62), 1, 10);
  add(215, 11, U(62), 1, 8);
  // Soplado
  add(306, 6, D(7), 1, 4);
  add(302, 8, D(2), 1, 4); add(305, 8, D(7), 1, 4);
  add(301, 10, D(2), 1, 4); add(304, 10, D(7), 1, 4);
  // Bloque 101–124
  for (let i = 0; i < 24; i++) add(101 + i, 14, U(3 * i + 1), 4, 2);
  // Bloque 125–143
  for (let k = 0; k < 19; k++) add(125 + k, 14 + k, D(1), 1, 11);
  return m;
}

function dentro_(r, f, c) {
  return f >= r.fila && f < r.fila + r.alto && c >= r.col && c < r.col + r.ancho;
}
/** Celda (superior izquierda) de la casilla de una máquina. */
function casilla_(sh, m) {
  return sh.getRange(m.chk.fila, m.chk.col);
}
function casillasA1_(sh) {
  return maquinas_().map(m => casilla_(sh, m).getA1Notation());
}

// ======================= TABLA DE PRODUCCIÓN =======================
// Una fila fija por máquina, en orden de número. Solo se muestran las filas de las máquinas activas.
const FILA_TABLA_TITULO = 35, FILA_TABLA_CAB = 36, FILA_TABLA_INI = 37;
const ORDEN_TABLA = maquinas_().map(m => m.id).sort((a, b) => a - b);
const FILA_TABLA_FIN = FILA_TABLA_INI + ORDEN_TABLA.length - 1;
const FILA_AYUDA = FILA_TABLA_FIN + 2;
const FILAS_HOJA = FILA_AYUDA + 1;

// k = campo de Planificación (entrada) o indicador calculado; ancho en unidades de 12 px.
const CAMPOS_TABLA = [
  { k: 'maquina',      titulo: 'MÁQ.',               ancho: 3 },
  { k: 'producto',     titulo: 'PRODUCTO',           ancho: 12, entrada: true, formato: '@' },
  { k: 'referencia',   titulo: 'REFERENCIA',         ancho: 7,  entrada: true, formato: '@' },
  { k: 'color',        titulo: 'COLOR',              ancho: 5,  entrada: true, formato: '@' },
  { k: 'material',     titulo: 'MATERIAL',           ancho: 6,  entrada: true, formato: '@' },
  { k: 'cicloIdeal',   titulo: 'CICLO IDEAL (s)',    ancho: 4,  entrada: true, formato: '0.0',   numero: true },
  { k: 'cicloReal',    titulo: 'CICLO REAL (s)',     ancho: 4,  entrada: true, formato: '0.0',   numero: true },
  { k: 'cavidades',    titulo: 'CAVI- DADES',        ancho: 4,  entrada: true, formato: '0',     numero: true },
  { k: 'contIni',      titulo: 'CONTADOR INICIAL',   ancho: 6,  entrada: true, formato: '#,##0', numero: true },
  { k: 'contFin',      titulo: 'CONTADOR FINAL',     ancho: 6,  entrada: true, formato: '#,##0', numero: true },
  { k: 'operarios',    titulo: 'N° OPERA- RIOS',     ancho: 4,  entrada: true, formato: '0',     numero: true },
  { k: 'golpes',       titulo: 'GOLPES',             ancho: 5,  formato: '#,##0' },
  { k: 'piezas',       titulo: 'PIEZAS PRODUCIDAS',  ancho: 5,  formato: '#,##0' },
  { k: 'teorica',      titulo: 'PROD. TEÓRICA',      ancho: 5,  formato: '#,##0' },
  { k: 'eficiencia',   titulo: 'EFICIEN- CIA CICLO', ancho: 4,  formato: '0.0%' },
  { k: 'cumplimiento', titulo: 'CUMPLI- MIENTO',     ancho: 4,  formato: '0.0%' },
];
(function () {
  let c = COL0;
  CAMPOS_TABLA.forEach(x => { x.col = c; c += x.ancho; });
})();
const CAMPO = {};
CAMPOS_TABLA.forEach(x => CAMPO[x.k] = x);
const ENTRADAS = CAMPOS_TABLA.filter(x => x.entrada);
const COL_ENTRADA_INI = ENTRADAS[0].col;
const COL_ENTRADA_FIN = ENTRADAS[ENTRADAS.length - 1].col + ENTRADAS[ENTRADAS.length - 1].ancho - 1;
const ANCHO_TABLA = CAMPOS_TABLA.reduce((s, x) => s + x.ancho, 0);

function filaTabla_(id) {
  return FILA_TABLA_INI + ORDEN_TABLA.indexOf(Number(id));
}

// ======================= MENÚ Y DISPARADORES =======================
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Turnos')
    .addItem('Actualizar planos y tablas', 'actualizarTodo')
    .addSeparator()
    .addItem('Configurar / reconstruir hojas de turnos', 'configurarArchivo')
    .addToUi();
}

function onEdit(e) {
  if (!e || !e.range) return;
  const rg = e.range, sh = rg.getSheet(), nombre = sh.getName();
  const t = turnoDeHoja_(nombre);
  if (nombre !== HOJA_PLAN && !t) return;

  // Varias casillas marcadas seguidas generan varias ejecuciones: se atienden una a la vez.
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(25000)) {
    SpreadsheetApp.getActive().toast('Hay otra actualización en curso. Vuelve a intentarlo.', 'Turnos', 4);
    return;
  }
  try {
    // Cambios hechos directamente en Planificación → sincroniza planos y tablas
    if (nombre === HOJA_PLAN) {
      if (rg.getRow() === 1 && rg.getNumRows() === 1) return;
      actualizarTodo_();
      return;
    }

    const f = rg.getRow(), c = rg.getColumn();
    if (f === CELDA.fechaSel[0] && c === CELDA.fechaSel[1]) {
      const datos = leerPlan_();
      pintarPlano_(sh, t, datos);
      cargarCasillas_(sh, t, datos);
      cargarTabla_(sh, t, datos);
    } else if (f === CELDA.todas[0] && c === CELDA.todas[1]) {
      const lista = sh.getRangeList(casillasA1_(sh));
      if (rg.getValue() === true) lista.check();
      else lista.uncheck();
      sincronizarCasillas_(sh, t);
    } else if (maquinas_().some(m => dentro_(m.chk, f, c))) {
      sincronizarCasillas_(sh, t);
    } else if (f <= FILA_TABLA_FIN && f + rg.getNumRows() - 1 >= FILA_TABLA_INI &&
               c <= COL_ENTRADA_FIN && c + rg.getNumColumns() - 1 >= COL_ENTRADA_INI) {
      const f0 = Math.max(f, FILA_TABLA_INI), f1 = Math.min(f + rg.getNumRows() - 1, FILA_TABLA_FIN);
      guardarFilas_(sh, t, f0, f1 - f0 + 1);
    }
  } finally {
    lock.releaseLock();
  }
}

function actualizarTodo() { actualizarTodo_(); }

// ======================= FÓRMULAS Y CONFIGURACIÓN REGIONAL =======================
// En configuraciones como es_ES las fórmulas separan los argumentos con ";" en lugar de ",".
// Las fórmulas de este archivo se escriben con "," y fx_() las adapta a la configuración del archivo.
let sep_ = null;
function separador_() {
  if (sep_) return sep_;
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(HOJA_LISTAS) || ss.getSheets()[0];
  const c = sh.getRange(1, sh.getMaxColumns());
  c.setFormula('=IF(TRUE,1,0)');
  SpreadsheetApp.flush();
  sep_ = c.getValue() === 1 ? ',' : ';';
  c.clearContent();
  return sep_;
}
function fx_(formula) {
  return separador_() === ',' ? formula : formula.replace(/,/g, ';');
}

// ======================= CONFIGURACIÓN DEL ARCHIVO =======================
function configurarArchivo() {
  const ss = SpreadsheetApp.getActive();
  prepararListas_(ss);
  prepararPlan_(ss);

  [1, 2, 3, 4].forEach(t => {
    const s = ss.getSheetByName(TURNOS[t].hoja);
    if (s) ss.deleteSheet(s);
  });

  // Se construye el Turno 4 y se copia para los demás (más rápido)
  const plantilla = ss.insertSheet(TURNOS[4].hoja);
  construirHojaTurno_(plantilla);
  ajustarTurno_(plantilla, 4);
  SpreadsheetApp.flush();
  [1, 2, 3].forEach(t => ajustarTurno_(plantilla.copyTo(ss).setName(TURNOS[t].hoja), t));

  [1, 2, 3, 4].forEach((t, i) => {
    const s = ss.getSheetByName(TURNOS[t].hoja);
    proteger_(s);
    ss.setActiveSheet(s);
    ss.moveActiveSheet(i + 1);
  });
  ss.setActiveSheet(ss.getSheetByName(HOJA_PLAN));
  ss.moveActiveSheet(5);

  // Elimina hojas vacías por defecto ("Hoja 1", "Sheet1"...)
  const nuestras = [HOJA_PLAN, HOJA_LISTAS].concat([1, 2, 3, 4].map(t => TURNOS[t].hoja));
  ss.getSheets().forEach(s => {
    if (nuestras.indexOf(s.getName()) === -1 && s.getLastRow() === 0 && s.getLastColumn() === 0) ss.deleteSheet(s);
  });

  actualizarTodo_();
  ss.setActiveSheet(ss.getSheetByName(TURNOS[4].hoja));
  ss.toast('Listo. Recarga la página (F5).', 'Turnos configurados', 8);
}

function prepararListas_(ss) {
  let sh = ss.getSheetByName(HOJA_LISTAS);
  if (!sh) sh = ss.insertSheet(HOJA_LISTAS);
  sh.clear();
  sh.getRange(1, 1).setValue('Máquinas');
  sh.getRange(2, 1, ORDEN_TABLA.length, 1).setValues(ORDEN_TABLA.map(x => [x]));
  sh.hideSheet();
}

function prepararPlan_(ss) {
  let sh = ss.getSheetByName(HOJA_PLAN);
  if (!sh) sh = ss.insertSheet(HOJA_PLAN);
  const SB = SpreadsheetApp.BorderStyle;
  if (sh.getMaxColumns() < 15) sh.insertColumnsAfter(sh.getMaxColumns(), 15 - sh.getMaxColumns());

  sh.getRange(1, 1, 1, 13).setValues([COLS_PLAN]);
  sh.getRange(1, 14).setFormula(fx_('={"Golpes";ARRAYFORMULA(IF((H2:H<>"")*(I2:I<>""),I2:I-H2:H,""))}'));
  sh.getRange(1, 15).setFormula(fx_('={"Piezas";ARRAYFORMULA(IF((N2:N<>"")*(J2:J<>""),N2:N*J2:J,""))}'));
  sh.getRange(1, 1, 1, 15).setFontWeight('bold').setFontColor('#ffffff').setBackground(COLOR.cabecera)
    .setWrap(true).setVerticalAlignment('middle').setHorizontalAlignment('center')
    .setBorder(true, true, true, true, true, true, '#ffffff', SB.SOLID);
  sh.getRange(1, 14, 1, 2).setBackground('#5b6676');
  sh.setFrozenRows(1);
  sh.setRowHeight(1, 36);

  const n = sh.getMaxRows() - 1;
  const dv = SpreadsheetApp.newDataValidation;
  sh.getRange(2, 1, n, 1).setNumberFormat('dd/mm/yyyy')
    .setDataValidation(dv().requireDate().setAllowInvalid(false).build());
  sh.getRange(2, 2, n, 1).setHorizontalAlignment('center')
    .setDataValidation(dv().requireValueInList(['1', '2', '3', '4'], true).setAllowInvalid(false).build());
  sh.getRange(2, 3, n, 1).setHorizontalAlignment('center')
    .setDataValidation(dv().requireValueInRange(ss.getSheetByName(HOJA_LISTAS).getRange('A2:A200'), true)
      .setAllowInvalid(false).build());
  sh.getRange(2, 5, n, 1).setNumberFormat('@');
  sh.getRange(2, 6, n, 2).setNumberFormat('0.0');
  sh.getRange(2, 8, n, 2).setNumberFormat('#,##0');
  sh.getRange(2, 10, n, 1).setNumberFormat('0');
  sh.getRange(2, 13, n, 1).setNumberFormat('0');
  sh.getRange(2, 14, n, 2).setNumberFormat('#,##0').setBackground('#f1f3f6');
  const num = dv().requireNumberGreaterThanOrEqualTo(0).setAllowInvalid(false).setHelpText('Solo números (≥ 0).').build();
  [6, 7, 8, 9, 10, 13].forEach(c => sh.getRange(2, c, n, 1).setDataValidation(num));

  const anchos = [90, 55, 70, 220, 120, 80, 80, 100, 100, 75, 100, 120, 80, 85, 90];
  anchos.forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.setTabColor('#5b6676');
}

// ======================= CONSTRUCCIÓN DE LA HOJA DE TURNO =======================
/** Letra(s) de una columna: 1 → A, 27 → AA. */
function letraCol_(c) {
  let s = '';
  for (; c > 0; c = Math.floor((c - 1) / 26)) s = String.fromCharCode(65 + (c - 1) % 26) + s;
  return s;
}

function caja_(sh, fila, col, alto, ancho) {
  const r = sh.getRange(fila, col, alto, ancho);
  if (alto > 1 || ancho > 1) r.merge();
  return r;
}
function etiqueta_(r, txt) {
  return r.setValue(txt).setFontSize(8).setFontWeight('bold').setFontColor(COLOR.etiqueta);
}
function seccion_(r, txt) {
  return r.setValue(txt).setFontSize(9).setFontWeight('bold').setFontColor(COLOR.texto);
}

function construirHojaTurno_(sh) {
  const SB = SpreadsheetApp.BorderStyle;

  // Tamaño, cuadrícula y estilo base
  if (sh.getMaxColumns() < ULTIMA_COL) sh.insertColumnsAfter(sh.getMaxColumns(), ULTIMA_COL - sh.getMaxColumns());
  if (sh.getMaxColumns() > ULTIMA_COL) sh.deleteColumns(ULTIMA_COL + 1, sh.getMaxColumns() - ULTIMA_COL);
  if (sh.getMaxRows() < FILAS_HOJA) sh.insertRowsAfter(sh.getMaxRows(), FILAS_HOJA - sh.getMaxRows());
  if (sh.getMaxRows() > FILAS_HOJA) sh.deleteRows(FILAS_HOJA + 1, sh.getMaxRows() - FILAS_HOJA);
  sh.setHiddenGridlines(true);
  sh.setColumnWidth(1, 16);
  sh.setColumnWidths(COL0, ULTIMA_COL - COL0, ANCHO_UNIDAD);
  sh.setColumnWidth(ULTIMA_COL, 16);
  const alto = { 1: 28, 2: 22, 3: 20, 4: 6, 5: 18, 6: 22, 7: 6, 8: 22, 9: 22, 10: 22, 11: 22, 12: 8, 13: 18,
    33: 8, 34: 8, [FILA_TABLA_TITULO]: 24, [FILA_TABLA_CAB]: 30, [FILA_TABLA_FIN + 1]: 8, [FILA_AYUDA]: 30 };
  for (let f = 14; f <= 32; f++) alto[f] = 19;
  Object.keys(alto).forEach(f => sh.setRowHeight(Number(f), alto[f]));
  sh.setRowHeights(FILA_TABLA_INI, ORDEN_TABLA.length, 21);
  sh.getRange(1, 1, FILAS_HOJA, ULTIMA_COL).setFontFamily('Roboto').setFontSize(9)
    .setFontColor(COLOR.texto).setVerticalAlignment('middle').setBackground('#ffffff');

  // ---- Encabezado ----
  caja_(sh, 1, U(1), 1, 71).setFontSize(13).setFontWeight('bold');
  caja_(sh, 1, D(1), 3, 11).setFontSize(8).setFontWeight('bold').setHorizontalAlignment('center').setWrap(true)
    .setBorder(true, true, true, true, null, null, COLOR.borde, SB.SOLID);

  etiqueta_(caja_(sh, 2, U(1), 1, 4), 'FECHA');
  const sel = caja_(sh, 2, U(5), 1, 8);
  sel.setNumberFormat('dd/mm/yyyy').setHorizontalAlignment('center').setFontWeight('bold').setFontSize(10)
    .setBackground(COLOR.entrada).setBorder(true, true, true, true, null, null, '#c9a227', SB.SOLID_MEDIUM)
    .setDataValidation(SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false)
      .setHelpText('Fecha a ver o planificar (doble clic abre el calendario).').build());
  const hoy = new Date();
  sel.setValue(new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate()));
  etiqueta_(caja_(sh, 2, U(14), 1, 5), 'HORARIO');
  caja_(sh, 2, U(19), 1, 12).setFontWeight('bold');
  etiqueta_(caja_(sh, 2, U(32), 1, 5), 'DESCANSO');
  caja_(sh, 2, U(37), 1, 17);
  etiqueta_(caja_(sh, 2, U(55), 1, 8), 'MÁQ. ACTIVAS');
  const activas = caja_(sh, 2, U(63), 1, 9).setFontWeight('bold').setFontSize(10).setFontColor('#b26a00');

  // Leyenda
  [[U(1), COLOR.activa, 'Activa (trabaja)'], [U(12), COLOR.inactiva, 'Sin planificar'],
   [U(23), COLOR.alerta, 'Alerta de molde']].forEach(([c, col, txt]) => {
    sh.getRange(3, c).setBackground(col).setBorder(true, true, true, true, null, null, COLOR.borde, SB.SOLID);
    caja_(sh, 3, c + 1, 1, 9).setValue(txt).setFontSize(8).setFontColor(COLOR.etiqueta);
  });

  // Horas efectivas del turno (las usa la producción teórica)
  etiqueta_(caja_(sh, 3, U(34), 1, 10), 'HORAS EFECTIVAS');
  caja_(sh, CELDA.horas[0], CELDA.horas[1], 1, 5).setNumberFormat('0.0 "h"').setFontWeight('bold')
    .setHorizontalAlignment('left');

  // Marcar / desmarcar todas
  caja_(sh, CELDA.todas[0], CELDA.todas[1], 1, 2).insertCheckboxes().setHorizontalAlignment('center');
  caja_(sh, CELDA.todas[0], CELDA.todas[1] + 2, 1, 15).setValue('Marcar / desmarcar todas')
    .setFontSize(8).setFontWeight('bold').setFontColor(COLOR.texto);

  // ---- Plano ----
  seccion_(caja_(sh, 5, U(1), 1, 20), 'MÁQUINAS 201–218');
  seccion_(caja_(sh, 13, U(1), 1, 20), 'MÁQUINAS 101–124');
  seccion_(caja_(sh, 13, D(1), 1, 11), 'MÁQUINAS 125–143');
  seccion_(caja_(sh, 5, D(1), 1, 11), ' SOPLADO');
  sh.getRange(5, D(1), 7, 11).setBorder(true, true, true, true, null, null, '#7a8594', SB.DASHED);

  const ms = maquinas_();
  const bloques = ms.map(m => sh.getRange(m.fila, m.col, m.alto, m.ancho));
  sh.getRangeList(bloques.map(r => r.getA1Notation())).setBackground(COLOR.inactiva)
    .setBorder(true, true, true, true, null, null, COLOR.borde, SB.SOLID_MEDIUM);

  const nums = ms.map(m => caja_(sh, m.num.fila, m.num.col, m.num.alto, m.num.ancho));
  sh.getRangeList(nums.map(r => r.getA1Notation())).setNumberFormat('@')
    .setFontFamily('Roboto Mono').setFontWeight('bold').setFontSize(9);
  const anchas = [], angostas = [], bajas = [];
  ms.forEach((m, i) => {
    const ancha = m.num.ancho >= 4;
    nums[i].setValue(ancha ? ' ' + m.id : String(m.id));
    const a1 = nums[i].getA1Notation();
    if (ancha) anchas.push(a1);
    else if (m.num.alto > 1) angostas.push(a1);
    else bajas.push(a1);
  });
  sh.getRangeList(anchas).setHorizontalAlignment('left').setVerticalAlignment('middle');
  sh.getRangeList(angostas).setHorizontalAlignment('center').setVerticalAlignment('top');
  sh.getRangeList(bajas).setHorizontalAlignment('center').setVerticalAlignment('middle').setFontSize(8);

  const chks = ms.map(m => caja_(sh, m.chk.fila, m.chk.col, 1, m.chk.ancho).getA1Notation());
  sh.getRangeList(chks).insertCheckboxes().setHorizontalAlignment('center').setVerticalAlignment('middle');

  // Colores del plano con formato condicional (se ven al instante al marcar una casilla).
  // Prioridad: alerta de molde > activa > sin planificar (fondo base).
  const abs = r => r.getA1Notation().replace(/^([A-Z]+)(\d+)$/, '$$$1$$$2');
  const celAlertas = abs(sh.getRange(CELDA.alertas[0], CELDA.alertas[1]));
  const regla = SpreadsheetApp.newConditionalFormatRule;
  const reglas = [];
  ms.forEach((m, i) => {
    const chk = abs(casilla_(sh, m));
    reglas.push(regla().whenFormulaSatisfied(fx_(`=AND(${chk},ISNUMBER(SEARCH("|${m.id}|",${celAlertas})))`))
      .setBackground(COLOR.alerta).setRanges([bloques[i]]).build());
    reglas.push(regla().whenFormulaSatisfied(`=${chk}`)
      .setBackground(COLOR.activa).setRanges([bloques[i]]).build());
  });

  // Contador de máquinas activas (cuenta las casillas marcadas del plano)
  const zonaPrincipal = sh.getRange(FILA_PLANO_INI, U(1), FILA_PLANO_PRINCIPAL_FIN - FILA_PLANO_INI + 1, 71).getA1Notation();
  const zonaDerecha = sh.getRange(FILA_PLANO_INI, D(1), FILA_PLANO_FIN - FILA_PLANO_INI + 1, 11).getA1Notation();
  const cuenta = `(COUNTIF(${zonaPrincipal},TRUE)+COUNTIF(${zonaDerecha},TRUE))`;
  sh.getRange(CELDA.activas[0], CELDA.activas[1]).setFormula(fx_(`=${cuenta}&" / ${META_MAQUINAS}"`));
  activas.setHorizontalAlignment('left');
  reglas.push(regla().whenFormulaSatisfied(fx_(`=${cuenta}>${META_MAQUINAS}`))
    .setFontColor('#b3261e').setRanges([activas]).build());
  reglas.push(regla().whenFormulaSatisfied(fx_(`=${cuenta}=${META_MAQUINAS}`))
    .setFontColor('#1b7a3a').setRanges([activas]).build());

  sh.getRange(CELDA.alertas[0], CELDA.alertas[1]).setFontColor('#ffffff').setFontSize(6)
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

  // ---- Tabla de producción ----
  const n = ORDEN_TABLA.length;
  const fechaA1 = abs(sh.getRange(CELDA.fechaSel[0], CELDA.fechaSel[1]));
  caja_(sh, FILA_TABLA_TITULO, COL0, 1, ANCHO_TABLA).setBackground(COLOR.cabecera).setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(10);
  sh.getRange(FILA_TABLA_TITULO, COL0).setFormula(fx_(
    `="  INFORMACIÓN DE PRODUCCIÓN · MÁQUINAS ACTIVAS · "&IF(ISNUMBER(${fechaA1}),TEXT(${fechaA1},"dd/mm/yyyy"),"sin fecha")`));

  CAMPOS_TABLA.forEach(x => {
    caja_(sh, FILA_TABLA_CAB, x.col, 1, x.ancho).setValue(x.titulo);
    sh.getRange(FILA_TABLA_INI, x.col, n, x.ancho).mergeAcross().setNumberFormat(x.formato || '0')
      .setBackground(x.entrada ? COLOR.entrada : COLOR.fijo)
      .setHorizontalAlignment(x.formato === '@' ? 'left' : 'center');
  });
  sh.getRange(FILA_TABLA_CAB, COL0, 1, ANCHO_TABLA).setBackground('#dfe3ea').setFontSize(7).setFontWeight('bold')
    .setFontColor(COLOR.texto).setWrap(true).setHorizontalAlignment('center');
  sh.getRange(FILA_TABLA_CAB, COL0, n + 1, ANCHO_TABLA)
    .setBorder(true, true, true, true, true, true, COLOR.lineaTabla, SB.SOLID);

  const maq = CAMPO.maquina;
  sh.getRange(FILA_TABLA_INI, maq.col, n, 1).setValues(ORDEN_TABLA.map(id => [id]));
  sh.getRange(FILA_TABLA_INI, maq.col, n, maq.ancho).setFontFamily('Roboto Mono').setFontWeight('bold');

  const num = SpreadsheetApp.newDataValidation().requireNumberGreaterThanOrEqualTo(0).setAllowInvalid(false)
    .setHelpText('Solo números (≥ 0).').build();
  ENTRADAS.filter(x => x.numero).forEach(x => sh.getRange(FILA_TABLA_INI, x.col, n, x.ancho).setDataValidation(num));

  // Indicadores calculados por fila
  const horas = abs(sh.getRange(CELDA.horas[0], CELDA.horas[1]));
  const formulas = {
    golpes:       r => `=IF(AND(ISNUMBER(${r.contFin}),ISNUMBER(${r.contIni})),${r.contFin}-${r.contIni},"")`,
    piezas:       r => `=IF(AND(ISNUMBER(${r.golpes}),ISNUMBER(${r.cavidades})),${r.golpes}*${r.cavidades},"")`,
    teorica:      r => `=IF(AND(ISNUMBER(${r.cicloIdeal}),ISNUMBER(${r.cavidades}),N(${r.cicloIdeal})>0),ROUNDDOWN(${horas}*3600/${r.cicloIdeal}*${r.cavidades},0),"")`,
    eficiencia:   r => `=IF(AND(ISNUMBER(${r.cicloIdeal}),ISNUMBER(${r.cicloReal}),N(${r.cicloReal})>0),${r.cicloIdeal}/${r.cicloReal},"")`,
    cumplimiento: r => `=IF(AND(ISNUMBER(${r.piezas}),ISNUMBER(${r.teorica}),N(${r.teorica})>0),${r.piezas}/${r.teorica},"")`,
  };
  Object.keys(formulas).forEach(k => {
    const filas = ORDEN_TABLA.map((id, i) => {
      const r = {};
      CAMPOS_TABLA.forEach(x => r[x.k] = letraCol_(x.col) + (FILA_TABLA_INI + i));
      return [fx_(formulas[k](r))];
    });
    sh.getRange(FILA_TABLA_INI, CAMPO[k].col, n, 1).setFormulas(filas);
  });

  // Fila naranja si la máquina tiene alerta de molde; contador final en rojo si es menor que el inicial
  const filaTabla = sh.getRange(FILA_TABLA_INI, COL0, n, ANCHO_TABLA);
  const maqA1 = '$' + letraCol_(maq.col) + FILA_TABLA_INI;
  reglas.push(regla().whenFormulaSatisfied(fx_(`=ISNUMBER(SEARCH("|"&${maqA1}&"|",${celAlertas}))`))
    .setBackground(COLOR.alerta).setRanges([filaTabla]).build());
  const ini0 = '$' + letraCol_(CAMPO.contIni.col) + FILA_TABLA_INI;
  const fin0 = '$' + letraCol_(CAMPO.contFin.col) + FILA_TABLA_INI;
  reglas.push(regla().whenFormulaSatisfied(fx_(`=AND(ISNUMBER(${ini0}),ISNUMBER(${fin0}),${fin0}<${ini0})`))
    .setFontColor('#b3261e').setBold(true)
    .setRanges([sh.getRange(FILA_TABLA_INI, CAMPO.contFin.col, n, CAMPO.contFin.ancho)]).build());
  sh.setConditionalFormatRules(reglas);

  caja_(sh, FILA_AYUDA, COL0, 1, ANCHO_TABLA).setWrap(true).setVerticalAlignment('top').setFontSize(8)
    .setFontColor(COLOR.etiqueta)
    .setValue('Cómo usar: 1) Elige la fecha arriba. 2) Marca la casilla de cada máquina que trabajará en el turno ' +
      '(o usa «Marcar / desmarcar todas»); su fila aparece en la tabla, en orden de número. 3) Completa los campos ' +
      'amarillos de la tabla: se guardan solos en la hoja Planificación. Para quitar una máquina con datos, borra sus ' +
      'datos de la tabla y desmarca la casilla. Turno 3: la fecha es la del día en que inicia el turno (10:30 PM).');

  sh.hideRows(FILA_TABLA_INI, n);
}

function ajustarTurno_(sh, t) {
  const T = TURNOS[t];
  sh.getRange(1, U(1)).setValue(`MANAPLAS · PLANIFICACIÓN DE PRODUCCIÓN · ${T.hoja.toUpperCase()}${T.cambioMolde ? ' (STANDARD)' : ''}`);
  sh.getRange(2, U(19)).setValue(T.horario);
  sh.getRange(2, U(37)).setValue(T.descanso);
  sh.getRange(CELDA.horas[0], CELDA.horas[1]).setValue(T.horas);
  sh.getRange(1, D(1))
    .setValue(T.cambioMolde ? 'TURNO STANDARD\nCambio de molde permitido' : 'SOLO PRODUCCIÓN\nNo se cambian moldes')
    .setBackground(T.cambioMolde ? '#dcebfb' : '#fde9d6');
  sh.setTabColor(T.color);
}

/** Protección de advertencia: avisa si se intenta escribir sobre el plano, las etiquetas o los cálculos. */
function proteger_(sh) {
  sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
  const p = sh.protect().setDescription('Plano y tabla (solo advertencia)').setWarningOnly(true);
  p.setUnprotectedRanges([
    sh.getRange(CELDA.fechaSel[0], CELDA.fechaSel[1], 1, 8),
    sh.getRange(CELDA.todas[0], CELDA.todas[1], 1, 2),
    sh.getRange(FILA_TABLA_INI, COL_ENTRADA_INI, ORDEN_TABLA.length, COL_ENTRADA_FIN - COL_ENTRADA_INI + 1),
  ].concat(maquinas_().map(m => sh.getRange(m.chk.fila, m.chk.col, 1, m.chk.ancho))));
}

// ======================= LÓGICA =======================
let tz_ = null;
function tz() { return tz_ || (tz_ = SpreadsheetApp.getActive().getSpreadsheetTimeZone()); }

function claveFecha_(v) {
  return (v instanceof Date && !isNaN(v)) ? Utilities.formatDate(v, tz(), 'yyyy-MM-dd') : '';
}
function turnoDeHoja_(nombre) {
  for (const t in TURNOS) if (TURNOS[t].hoja === nombre) return Number(t);
  return 0;
}
function norm_(s) { return String(s == null ? '' : s).trim().toUpperCase(); }

function ultimaFilaPlan_(sh) {
  const v = sh.getRange(1, 1, sh.getMaxRows(), 3).getValues();
  for (let i = v.length - 1; i >= 0; i--) if (v[i].some(x => x !== '' && x !== null)) return i + 1;
  return 1;
}

/** Lee la hoja Planificación como una lista de registros. */
function leerPlan_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(HOJA_PLAN);
  if (!sh) return [];
  const ult = ultimaFilaPlan_(sh);
  if (ult < 2) return [];
  const out = [];
  sh.getRange(2, 1, ult - 1, 13).getValues().forEach((r, i) => {
    const fecha = claveFecha_(r[IDX.fecha]), turno = Number(r[IDX.turno]), maq = String(r[IDX.maquina]).trim();
    if (!fecha || !TURNOS[turno] || !maq) return;
    out.push({ fila: i + 2, fecha: fecha, turno: turno, maquina: maq, ref: String(r[IDX.referencia]).trim(), valores: r });
  });
  return out;
}

/** Registros de una fecha y turno, por número de máquina. */
function registros_(datos, clave, t) {
  const rec = {};
  if (clave) datos.forEach(r => { if (r.fecha === clave && r.turno === t) rec[r.maquina] = r; });
  return rec;
}

/** true si el registro tiene algún dato de producción además de fecha, turno y máquina. */
function tieneDatos_(r) {
  return r.valores.slice(IDX.producto).some(x => x !== '' && x !== null);
}

/** Minutos absolutos de inicio del turno (para ordenar cronológicamente). */
function tiempo_(r) {
  const p = r.fecha.split('-').map(Number);
  return Date.UTC(p[0], p[1] - 1, p[2]) / 60000 + TURNOS[r.turno].inicioMin;
}

/**
 * En los turnos 1, 2 y 3 no se cambian moldes: si la referencia difiere de la del
 * último turno anterior de esa misma máquina, devuelve ese registro anterior.
 */
function alertaMolde_(datos, r) {
  if (TURNOS[r.turno].cambioMolde) return null;
  const ref = norm_(r.ref);
  if (!ref) return null;
  const t0 = tiempo_(r);
  let prev = null, tPrev = -Infinity;
  datos.forEach(o => {
    if (o.maquina !== r.maquina || !norm_(o.ref)) return;
    const to = tiempo_(o);
    if (to < t0 && to > tPrev) { prev = o; tPrev = to; }
  });
  return (prev && norm_(prev.ref) !== ref) ? prev : null;
}

/**
 * Actualiza la lista de máquinas con alerta de molde. Los colores del plano y de la tabla
 * los pone el formato condicional (casillas y esta lista).
 */
function pintarPlano_(sh, t, datos) {
  const clave = claveFecha_(sh.getRange(CELDA.fechaSel[0], CELDA.fechaSel[1]).getValue());
  const alertas = [];
  if (clave) datos.forEach(r => {
    if (r.fecha === clave && r.turno === t && alertaMolde_(datos, r)) alertas.push(r.maquina);
  });
  sh.getRange(CELDA.alertas[0], CELDA.alertas[1]).setValue(alertas.length ? '|' + alertas.join('|') + '|' : '');
}

/** Marca las casillas de las máquinas que tienen registro para la fecha y turno de la hoja. */
function cargarCasillas_(sh, t, datos) {
  const rec = registros_(datos, claveFecha_(sh.getRange(CELDA.fechaSel[0], CELDA.fechaSel[1]).getValue()), t);
  const si = [], no = [];
  maquinas_().forEach(m => (rec[String(m.id)] ? si : no).push(casilla_(sh, m).getA1Notation()));
  if (si.length) sh.getRangeList(si).check();
  if (no.length) sh.getRangeList(no).uncheck();
}

/**
 * Llena la tabla con los registros de la fecha y turno de la hoja y muestra solo las filas
 * de las máquinas activas. Con `ids` solo reescribe los valores de esas máquinas (así no pisa
 * lo que se esté escribiendo en otras filas).
 */
function cargarTabla_(sh, t, datos, ids) {
  const rec = registros_(datos, claveFecha_(sh.getRange(CELDA.fechaSel[0], CELDA.fechaSel[1]).getValue()), t);
  const ancho = COL_ENTRADA_FIN - COL_ENTRADA_INI + 1;
  const filaValores = id => {
    const fila = new Array(ancho).fill('');
    const r = rec[String(id)];
    if (r) ENTRADAS.forEach(x => fila[x.col - COL_ENTRADA_INI] = r.valores[IDX[x.k]]);
    return fila;
  };

  if (ids === undefined) {
    sh.getRange(FILA_TABLA_INI, COL_ENTRADA_INI, ORDEN_TABLA.length, ancho).setValues(ORDEN_TABLA.map(filaValores));
  } else {
    ids.forEach(id => sh.getRange(filaTabla_(id), COL_ENTRADA_INI, 1, ancho).setValues([filaValores(id)]));
  }

  // Muestra u oculta las filas por tramos seguidos
  let i = 0;
  while (i < ORDEN_TABLA.length) {
    const visible = !!rec[String(ORDEN_TABLA[i])];
    let j = i;
    while (j + 1 < ORDEN_TABLA.length && !!rec[String(ORDEN_TABLA[j + 1])] === visible) j++;
    if (visible) sh.showRows(FILA_TABLA_INI + i, j - i + 1);
    else sh.hideRows(FILA_TABLA_INI + i, j - i + 1);
    i = j + 1;
  }
}

/**
 * Lleva el estado de las casillas a la hoja Planificación:
 *   marcada sin registro      → crea el registro (fecha, turno, máquina) y muestra su fila en la tabla
 *   desmarcada con registro   → lo borra si está vacío; si tiene producción, vuelve a marcar la casilla
 * Lee siempre el estado actual de todas las casillas, así que varios clics seguidos quedan bien.
 */
function sincronizarCasillas_(sh, t) {
  const ss = SpreadsheetApp.getActive();
  const fechaV = sh.getRange(CELDA.fechaSel[0], CELDA.fechaSel[1]).getValue();
  const clave = claveFecha_(fechaV);
  if (!clave) {
    ss.toast('Elige una fecha válida arriba para registrar las máquinas activas.', 'Sin fecha', 5);
    return;
  }

  const plano = sh.getRange(FILA_PLANO_INI, COL0, FILA_PLANO_FIN - FILA_PLANO_INI + 1, ULTIMA_COL - COL0).getValues();
  const marcada = m => plano[m.chk.fila - FILA_PLANO_INI][m.chk.col - COL0] === true;
  const rec = registros_(leerPlan_(), clave, t);

  const nuevas = [], borrar = [], conservar = [];
  let n = 0;
  maquinas_().forEach(m => {
    const r = rec[String(m.id)];
    if (marcada(m)) {
      n++;
      if (!r) nuevas.push(m.id);
    } else if (r) {
      if (tieneDatos_(r)) conservar.push(m);
      else borrar.push(r.fila);
    }
  });

  const plan = ss.getSheetByName(HOJA_PLAN);
  borrar.sort((a, b) => b - a).forEach(f => plan.deleteRow(f));
  if (nuevas.length) {
    const ini = ultimaFilaPlan_(plan) + 1;
    const fin = ini + nuevas.length - 1;
    if (fin > plan.getMaxRows()) plan.insertRowsAfter(plan.getMaxRows(), fin - plan.getMaxRows() + 200);
    plan.getRange(ini, 1, nuevas.length, 13).setValues(nuevas.map(id => {
      const fila = new Array(13).fill('');
      fila[IDX.fecha] = fechaV; fila[IDX.turno] = t; fila[IDX.maquina] = id;
      return fila;
    }));
  }

  if (conservar.length) {
    sh.getRangeList(conservar.map(m => casilla_(sh, m).getA1Notation())).check();
    n += conservar.length;
  }
  cargarTabla_(sh, t, leerPlan_(), nuevas);

  if (conservar.length) {
    ss.toast(`Máquina(s) ${conservar.map(m => m.id).join(', ')}: tienen producción guardada y siguen activas. ` +
      'Para quitarlas, borra sus datos de la tabla y desmarca la casilla.', 'No se desmarcaron', 8);
  } else if (n > META_MAQUINAS) {
    ss.toast(`Hay ${n} máquinas activas y la meta por turno es ${META_MAQUINAS}.`, TURNOS[t].hoja, 5);
  }
}

/** Guarda en Planificación las filas de la tabla editadas (f0 .. f0+n-1). */
function guardarFilas_(sh, t, f0, n) {
  const ss = SpreadsheetApp.getActive();
  const fechaV = sh.getRange(CELDA.fechaSel[0], CELDA.fechaSel[1]).getValue();
  const clave = claveFecha_(fechaV);
  if (!clave) {
    ss.toast('Elige una fecha válida arriba antes de cargar datos.', 'No se guardó', 5);
    return;
  }

  const filas = sh.getRange(f0, COL0, n, ANCHO_TABLA).getValues();
  const plan = ss.getSheetByName(HOJA_PLAN);
  const rec = registros_(leerPlan_(), clave, t);
  const nuevas = [];
  filas.forEach((fila, i) => {
    const id = ORDEN_TABLA[f0 + i - FILA_TABLA_INI];
    if (sh.isRowHiddenByUser(f0 + i)) return;
    const valores = new Array(13).fill('');
    valores[IDX.fecha] = fechaV; valores[IDX.turno] = t; valores[IDX.maquina] = id;
    ENTRADAS.forEach(x => valores[IDX[x.k]] = fila[x.col - COL0]);
    const r = rec[String(id)];
    if (r) plan.getRange(r.fila, IDX.producto + 1, 1, 13 - IDX.producto).setValues([valores.slice(IDX.producto)]);
    else nuevas.push(valores);
  });
  if (nuevas.length) {
    const ini = ultimaFilaPlan_(plan) + 1;
    const fin = ini + nuevas.length - 1;
    if (fin > plan.getMaxRows()) plan.insertRowsAfter(plan.getMaxRows(), fin - plan.getMaxRows() + 200);
    plan.getRange(ini, 1, nuevas.length, 13).setValues(nuevas);
  }

  pintarTodos_(leerPlan_());
}

function pintarTodos_(datos) {
  [1, 2, 3, 4].forEach(t => {
    const sh = SpreadsheetApp.getActive().getSheetByName(TURNOS[t].hoja);
    if (sh) pintarPlano_(sh, t, datos);
  });
}

function actualizarTodo_() {
  const datos = leerPlan_();
  [1, 2, 3, 4].forEach(t => {
    const sh = SpreadsheetApp.getActive().getSheetByName(TURNOS[t].hoja);
    if (!sh) return;
    pintarPlano_(sh, t, datos);
    cargarCasillas_(sh, t, datos);
    cargarTabla_(sh, t, datos);
  });
}
