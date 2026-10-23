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
 *   Listas (oculta)                         → máquinas y su tonelaje
 *
 * Uso:
 *   · Casilla de cada máquina        → la marca como activa y muestra su fila en la tabla.
 *   · «Marcar / desmarcar todas»     → todas las máquinas. «Línea 1/2/3» → solo esa línea.
 *   · Tabla de producción            → una fila por máquina activa, en orden de número.
 *   · «Enviar a Planificación»       → guarda la selección y la tabla en la hoja Planificación.
 *   · «Descartar cambios»            → vuelve a lo último enviado para esa fecha y turno.
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

// Columnas de la hoja Planificación (A:O). P y Q son fórmulas (Golpes y Piezas).
const COLS_PLAN = ['Fecha', 'Turno', 'Máquina', 'Ton', 'OP', 'Producto', 'Referencia', 'Ciclo ideal (s)',
  'Ciclo real (s)', 'Contador inicial', 'Contador final', 'Cavidades', 'Color', 'Material', 'N° operarios'];
const IDX = { fecha: 0, turno: 1, maquina: 2, ton: 3, op: 4, producto: 5, referencia: 6, cicloIdeal: 7,
  cicloReal: 8, contIni: 9, contFin: 10, cavidades: 11, color: 12, material: 13, operarios: 14 };
const N_PLAN = COLS_PLAN.length;

// Grupos de selección rápida
const GRUPOS = [
  { nombre: 'Línea 1', desde: 101, hasta: 143 },
  { nombre: 'Línea 2', desde: 201, hasta: 218 },
  { nombre: 'Línea 3', desde: 301, hasta: 306 },
];

const COLOR = {
  inactiva: '#eef1f5', activa: '#b9e4c1', alerta: '#ffc98a',
  borde: '#1f2a37', texto: '#1f2a37', etiqueta: '#5b6676',
  cabecera: '#1f2a37', entrada: '#fff8dc', fijo: '#f1f3f6', lineaTabla: '#d0d5dd', enviar: '#1f6fd1',
};

// ======================= CUADRÍCULA =======================
// Columna A = margen. El plano usa "unidades" de 29 px desde la columna B (58 unidades de ancho):
//   1–43  Líneas 2 (201–218) y 1 (101–124) · 45–50  Máquinas 125–143 · 52–57  Soplado
// Todas las máquinas tienen el mismo tamaño (3 × 2 celdas) y el plano (filas 1–25) cabe
// en una pantalla Full HD sin desplazarse.
const ANCHO_UNIDAD = 29;
const COL0 = 2;
const N_UNIDADES = 58;
const U = u => COL0 + u - 1;
const COL_125 = U(45);              // columnas de las máquinas 125–134 y 135–143
const S = u => U(51) + u;           // zona Soplado: unidades 1..6
const ULTIMA_COL = U(N_UNIDADES) + 1; // columna de margen derecho
const ANCHO_TOTAL = N_UNIDADES;
// Las máquinas van en "franjas" de 3 filas: número, casilla y un separador fino, así el espacio
// entre filas de máquinas es siempre el mismo. Hay 10 franjas (filas 6–35).
const FILA_PLANO_INI = 6;
const franja_ = k => FILA_PLANO_INI + 3 * k;
const FILA_PLANO_FIN = franja_(9) + 1;
const ALTO_FILA_PLANO = 24, ALTO_SEPARADOR = 8;

const CELDA = {
  fechaSel: [2, U(4)], activas: [2, U(36)], horas: [3, U(28)], todas: [3, U(33)],
  grupos: [[4, U(33)], [4, U(38)], [4, U(43)]],
  // Celdas auxiliares en la columna A (texto invisible)
  alertas: [FILA_PLANO_FIN + 2, 1],     // máquinas con alerta de molde: "|203|215|"
  pendiente: [FILA_PLANO_FIN + 3, 1],   // TRUE si hay cambios sin enviar
  fechaCargada: [FILA_PLANO_FIN + 4, 1], // fecha cuyos datos muestra la hoja
};
const ANCHO_FECHA = 5;

/**
 * Distribución de las máquinas según la hoja resumen de asignación de personal.
 * Todas las máquinas miden 3 × 2 celdas y se dividen en dos zonas:
 *   num → número de la máquina (fila de arriba)
 *   chk → casilla para marcarla como activa (fila de abajo)
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
  const q = (id, fila, col) => add(id, fila, col, 2, 3);
  // Bloque 201–218: 201 encima de la 202, 202–214 en fila y 218–215 en columna a la derecha
  q(201, franja_(0), U(1));
  for (let j = 0; j < 13; j++) q(202 + j, franja_(1), U(3 * j + 1));
  q(218, franja_(0), U(41)); q(217, franja_(1), U(41)); q(216, franja_(2), U(41)); q(215, franja_(3), U(41));
  // Bloque 101–124 en dos filas: 101–112 y 113–124
  // (la franja 2 queda para el título «MÁQUINAS 101–124»)
  for (let i = 0; i < 24; i++) q(101 + i, franja_(3 + Math.floor(i / 12)), U(3 * (i % 12) + 1));
  // Bloque 125–143 en dos columnas: 125–134 y 135–143
  for (let k = 0; k < 19; k++) q(125 + k, franja_(k % 10), COL_125 + 3 * Math.floor(k / 10));
  // Soplado
  q(306, franja_(0), S(4));
  q(302, franja_(1), S(1)); q(305, franja_(1), S(4));
  q(301, franja_(2), S(1)); q(304, franja_(2), S(4));
  return m;
}

function dentro_(r, f, c) {
  return f >= r.fila && f < r.fila + r.alto && c >= r.col && c < r.col + r.ancho;
}
/** Celda (superior izquierda) de la casilla de una máquina. */
function casilla_(sh, m) {
  return sh.getRange(m.chk.fila, m.chk.col);
}

// ======================= TABLA DE PRODUCCIÓN =======================
// Una fila fija por máquina, en orden de número. Solo se muestran las filas de las máquinas activas.
const FILA_TABLA_TITULO = FILA_PLANO_FIN + 5, FILA_ACCIONES = FILA_TABLA_TITULO + 1,
  FILA_TABLA_CAB = FILA_TABLA_TITULO + 2, FILA_TABLA_INI = FILA_TABLA_TITULO + 3;
const ORDEN_TABLA = maquinas_().map(m => m.id).sort((a, b) => a - b);
const FILA_TABLA_FIN = FILA_TABLA_INI + ORDEN_TABLA.length - 1;
const FILA_AYUDA = FILA_TABLA_FIN + 2;
const FILAS_HOJA = FILA_AYUDA + 1;

const ACCION = { enviar: [FILA_ACCIONES, U(1)], descartar: [FILA_ACCIONES, U(12)], estado: [FILA_ACCIONES, U(20)] };

// k = campo de Planificación; ancho en unidades (deben sumar N_UNIDADES = 58).
const CAMPOS_TABLA = [
  { k: 'maquina',    titulo: 'MÁQ.',              ancho: 2 },
  { k: 'ton',        titulo: 'Ton',               ancho: 2,  entrada: true, formato: '0',     numero: true },
  { k: 'op',         titulo: 'OP',                ancho: 4,  entrada: true, formato: '@' },
  { k: 'producto',   titulo: 'PRODUCTO',          ancho: 11, entrada: true, formato: '@' },
  { k: 'referencia', titulo: 'REFERENCIA',        ancho: 6,  entrada: true, formato: '@' },
  { k: 'color',      titulo: 'COLOR',             ancho: 4,  entrada: true, formato: '@' },
  { k: 'material',   titulo: 'MATERIAL',          ancho: 5,  entrada: true, formato: '@' },
  { k: 'cicloIdeal', titulo: 'CICLO IDEAL (s)',   ancho: 4,  entrada: true, formato: '0.0',   numero: true },
  { k: 'cicloReal',  titulo: 'CICLO REAL (s)',    ancho: 4,  entrada: true, formato: '0.0',   numero: true },
  { k: 'cavidades',  titulo: 'CAVIDADES',         ancho: 4,  entrada: true, formato: '0',     numero: true },
  { k: 'contIni',    titulo: 'CONTADOR INICIAL',  ancho: 4,  entrada: true, formato: '#,##0', numero: true },
  { k: 'contFin',    titulo: 'CONTADOR FINAL',    ancho: 4,  entrada: true, formato: '#,##0', numero: true },
  { k: 'operarios',  titulo: 'N° OPERARIOS',      ancho: 4,  entrada: true, formato: '0',     numero: true },
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
    // Cambios hechos directamente en Planificación → actualiza las hojas sin cambios pendientes
    if (nombre === HOJA_PLAN) {
      if (rg.getRow() === 1 && rg.getNumRows() === 1) return;
      actualizarTodo_();
      return;
    }

    const f = rg.getRow(), c = rg.getColumn();
    const es = celda => f === celda[0] && c === celda[1];
    if (es(CELDA.fechaSel)) {
      cambiarFecha_(sh, t);
    } else if (es(ACCION.enviar)) {
      if (rg.getValue() === true) { rg.setValue(false); enviar_(sh, t); }
    } else if (es(ACCION.descartar)) {
      if (rg.getValue() === true) {
        rg.setValue(false);
        cargar_(sh, t, leerPlan_());
        SpreadsheetApp.getActive().toast('Se volvió a lo último enviado a Planificación.', 'Cambios descartados', 4);
      }
    } else if (es(CELDA.todas) || CELDA.grupos.some(es)) {
      const g = es(CELDA.todas) ? null : GRUPOS[CELDA.grupos.findIndex(es)];
      const celdas = maquinas_().filter(m => !g || (m.id >= g.desde && m.id <= g.hasta))
        .map(m => casilla_(sh, m).getA1Notation());
      if (rg.getValue() === true) sh.getRangeList(celdas).check();
      else sh.getRangeList(celdas).uncheck();
      aplicarSeleccion_(sh, t);
      marcarPendiente_(sh);
    } else if (maquinas_().some(m => dentro_(m.chk, f, c))) {
      aplicarSeleccion_(sh, t);
      marcarPendiente_(sh);
    } else if (f <= FILA_TABLA_FIN && f + rg.getNumRows() - 1 >= FILA_TABLA_INI &&
               c <= COL_ENTRADA_FIN && c + rg.getNumColumns() - 1 >= COL_ENTRADA_INI) {
      marcarPendiente_(sh);
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

/** Listas: máquina y tonelaje. Conserva los tonelajes ya cargados. */
function prepararListas_(ss) {
  let sh = ss.getSheetByName(HOJA_LISTAS);
  if (!sh) sh = ss.insertSheet(HOJA_LISTAS);
  const ton = tonelajes_();
  sh.clear();
  sh.getRange(1, 1, 1, 2).setValues([['Máquinas', 'Ton']]).setFontWeight('bold');
  sh.getRange(2, 1, ORDEN_TABLA.length, 2).setValues(ORDEN_TABLA.map(id => [id, ton[id] === undefined ? '' : ton[id]]));
  sh.hideSheet();
}

/** Tonelaje por máquina según Listas (columna B). */
function tonelajes_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(HOJA_LISTAS);
  const out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(([id, ton]) => {
    if (id !== '' && ton !== '' && ton !== null) out[Number(id)] = ton;
  });
  return out;
}

function prepararPlan_(ss) {
  let sh = ss.getSheetByName(HOJA_PLAN);
  if (!sh) sh = ss.insertSheet(HOJA_PLAN);
  const SB = SpreadsheetApp.BorderStyle;

  // Versión anterior (sin Ton ni OP): se insertan las dos columnas después de Máquina sin perder datos
  if (sh.getRange(1, 3).getValue() === 'Máquina' && sh.getRange(1, 4).getValue() === 'Producto') {
    sh.insertColumnsAfter(3, 2);
  }
  if (sh.getMaxColumns() < N_PLAN + 2) sh.insertColumnsAfter(sh.getMaxColumns(), N_PLAN + 2 - sh.getMaxColumns());

  sh.getRange(1, 1, 1, N_PLAN).setValues([COLS_PLAN]);
  sh.getRange(1, N_PLAN + 1).setFormula(fx_('={"Golpes";ARRAYFORMULA(IF((J2:J<>"")*(K2:K<>""),K2:K-J2:J,""))}'));
  sh.getRange(1, N_PLAN + 2).setFormula(fx_('={"Piezas";ARRAYFORMULA(IF((P2:P<>"")*(L2:L<>""),P2:P*L2:L,""))}'));
  sh.getRange(1, 1, 1, N_PLAN + 2).setFontWeight('bold').setFontColor('#ffffff').setBackground(COLOR.cabecera)
    .setWrap(true).setVerticalAlignment('middle').setHorizontalAlignment('center')
    .setBorder(true, true, true, true, true, true, '#ffffff', SB.SOLID);
  sh.getRange(1, N_PLAN + 1, 1, 2).setBackground('#5b6676');
  sh.setFrozenRows(1);
  sh.setRowHeight(1, 36);

  const n = sh.getMaxRows() - 1;
  const col = k => IDX[k] + 1;
  const dv = SpreadsheetApp.newDataValidation;
  sh.getRange(2, 1, n, N_PLAN).setDataValidation(null);
  sh.getRange(2, col('fecha'), n, 1).setNumberFormat('dd/mm/yyyy')
    .setDataValidation(dv().requireDate().setAllowInvalid(false).build());
  sh.getRange(2, col('turno'), n, 1).setHorizontalAlignment('center')
    .setDataValidation(dv().requireValueInList(['1', '2', '3', '4'], true).setAllowInvalid(false).build());
  sh.getRange(2, col('maquina'), n, 1).setHorizontalAlignment('center').setNumberFormat('0')
    .setDataValidation(dv().requireValueInRange(ss.getSheetByName(HOJA_LISTAS).getRange('A2:A200'), true)
      .setAllowInvalid(false).build());
  ['op', 'producto', 'referencia', 'color', 'material'].forEach(k => sh.getRange(2, col(k), n, 1).setNumberFormat('@'));
  sh.getRange(2, col('ton'), n, 1).setNumberFormat('0').setHorizontalAlignment('center');
  sh.getRange(2, col('cicloIdeal'), n, 2).setNumberFormat('0.0');
  sh.getRange(2, col('contIni'), n, 2).setNumberFormat('#,##0');
  sh.getRange(2, col('cavidades'), n, 1).setNumberFormat('0');
  sh.getRange(2, col('operarios'), n, 1).setNumberFormat('0');
  sh.getRange(2, N_PLAN + 1, n, 2).setNumberFormat('#,##0').setBackground('#f1f3f6');
  const num = dv().requireNumberGreaterThanOrEqualTo(0).setAllowInvalid(false).setHelpText('Solo números (≥ 0).').build();
  ['ton', 'cicloIdeal', 'cicloReal', 'contIni', 'contFin', 'cavidades', 'operarios']
    .forEach(k => sh.getRange(2, col(k), n, 1).setDataValidation(num));

  const anchos = [90, 55, 70, 60, 100, 220, 120, 80, 80, 100, 100, 75, 100, 120, 80, 85, 90];
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
  return r.setValue(txt).setFontSize(9).setFontWeight('bold').setFontColor(COLOR.etiqueta);
}
function seccion_(r, txt) {
  return r.setValue(txt).setFontSize(10).setFontWeight('bold').setFontColor(COLOR.texto);
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
  const alto = { 1: 34, 2: 28, 3: 24, 4: 24, 5: 24, [FILA_PLANO_FIN + 1]: 8,
    [CELDA.alertas[0]]: 6, [CELDA.pendiente[0]]: 6, [CELDA.fechaCargada[0]]: 6,
    [FILA_TABLA_TITULO]: 28, [FILA_ACCIONES]: 32, [FILA_TABLA_CAB]: 34, [FILA_TABLA_FIN + 1]: 10, [FILA_AYUDA]: 36 };
  for (let f = FILA_PLANO_INI; f <= FILA_PLANO_FIN; f++) {
    alto[f] = (f - FILA_PLANO_INI) % 3 === 2 ? ALTO_SEPARADOR : ALTO_FILA_PLANO;
  }
  Object.keys(alto).forEach(f => sh.setRowHeight(Number(f), alto[f]));
  sh.setRowHeights(FILA_TABLA_INI, ORDEN_TABLA.length, 26);
  sh.getRange(1, 1, FILAS_HOJA, ULTIMA_COL).setFontFamily('Roboto').setFontSize(10)
    .setFontColor(COLOR.texto).setVerticalAlignment('middle').setBackground('#ffffff');

  // ---- Encabezado ----
  caja_(sh, 1, U(1), 1, 47).setFontSize(16).setFontWeight('bold');
  caja_(sh, 1, U(49), 4, 10).setFontSize(10).setFontWeight('bold').setHorizontalAlignment('center').setWrap(true)
    .setBorder(true, true, true, true, null, null, COLOR.borde, SB.SOLID);

  etiqueta_(caja_(sh, 2, U(1), 1, 3), 'FECHA');
  const sel = caja_(sh, CELDA.fechaSel[0], CELDA.fechaSel[1], 1, ANCHO_FECHA);
  sel.setNumberFormat('dd/mm/yyyy').setHorizontalAlignment('center').setFontWeight('bold').setFontSize(12)
    .setBackground(COLOR.entrada).setBorder(true, true, true, true, null, null, '#c9a227', SB.SOLID_MEDIUM)
    .setDataValidation(SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false)
      .setHelpText('Fecha a ver o planificar (doble clic abre el calendario).').build());
  const hoy = new Date();
  sel.setValue(new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate()));
  etiqueta_(caja_(sh, 2, U(10), 1, 3), 'HORARIO');
  caja_(sh, 2, U(13), 1, 7).setFontWeight('bold').setFontSize(11);
  etiqueta_(caja_(sh, 2, U(21), 1, 3), 'DESCANSO');
  caja_(sh, 2, U(24), 1, 8).setFontSize(11);
  etiqueta_(caja_(sh, 2, U(32), 1, 4), 'MÁQ. ACTIVAS');
  const activas = caja_(sh, CELDA.activas[0], CELDA.activas[1], 1, 4).setFontWeight('bold').setFontSize(12).setFontColor('#b26a00');

  // Leyenda
  [[U(1), COLOR.activa, 'Activa (trabaja)'], [U(8), COLOR.inactiva, 'Sin planificar'],
   [U(15), COLOR.alerta, 'Alerta de molde']].forEach(([c, col, txt]) => {
    sh.getRange(3, c).setBackground(col).setBorder(true, true, true, true, null, null, COLOR.borde, SB.SOLID);
    caja_(sh, 3, c + 1, 1, 5).setValue(txt).setFontSize(9).setFontColor(COLOR.etiqueta);
  });

  // Horas efectivas del turno
  etiqueta_(caja_(sh, 3, U(22), 1, 6), 'HORAS EFECTIVAS');
  caja_(sh, CELDA.horas[0], CELDA.horas[1], 1, 3).setNumberFormat('0.0 "h"').setFontWeight('bold')
    .setHorizontalAlignment('left');

  // Selección rápida: todas y por línea
  sh.getRange(CELDA.todas[0], CELDA.todas[1]).insertCheckboxes().setHorizontalAlignment('center');
  caja_(sh, CELDA.todas[0], CELDA.todas[1] + 1, 1, 14).setValue('Marcar / desmarcar todas')
    .setFontSize(10).setFontWeight('bold');
  GRUPOS.forEach((g, i) => {
    const [f, c] = CELDA.grupos[i];
    sh.getRange(f, c).insertCheckboxes().setHorizontalAlignment('center');
    caja_(sh, f, c + 1, 1, 4).setValue(`${g.nombre} (${g.desde}–${g.hasta})`).setFontSize(9).setFontWeight('bold');
  });

  // ---- Plano ----
  seccion_(caja_(sh, 5, U(1), 1, 20), 'MÁQUINAS 201–218');
  seccion_(caja_(sh, franja_(2) + 1, U(1), 1, 20), 'MÁQUINAS 101–124');
  seccion_(caja_(sh, 5, COL_125, 1, 6), 'MÁQUINAS 125–143');
  seccion_(caja_(sh, 5, S(1), 1, 6), ' SOPLADO');
  sh.getRange(5, S(1), franja_(2) + 2 - 5, 6).setBorder(true, true, true, true, null, null, '#7a8594', SB.DASHED);

  const ms = maquinas_();
  const bloques = ms.map(m => sh.getRange(m.fila, m.col, m.alto, m.ancho));
  sh.getRangeList(bloques.map(r => r.getA1Notation())).setBackground(COLOR.inactiva)
    .setBorder(true, true, true, true, null, null, COLOR.borde, SB.SOLID_MEDIUM);

  const nums = ms.map(m => caja_(sh, m.num.fila, m.num.col, m.num.alto, m.num.ancho));
  sh.getRangeList(nums.map(r => r.getA1Notation())).setNumberFormat('@')
    .setFontFamily('Roboto Mono').setFontWeight('bold').setFontSize(14);
  const anchas = [], angostas = [], bajas = [];
  ms.forEach((m, i) => {
    const ancha = m.num.ancho >= 4;
    nums[i].setValue(ancha ? '  ' + m.id : String(m.id));
    const a1 = nums[i].getA1Notation();
    if (ancha) anchas.push(a1);
    else if (m.num.alto > 1) angostas.push(a1);
    else bajas.push(a1);
  });
  sh.getRangeList(anchas).setHorizontalAlignment('left').setVerticalAlignment('middle');
  sh.getRangeList(angostas).setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRangeList(bajas).setHorizontalAlignment('center').setVerticalAlignment('middle');

  const chks = ms.map(m => caja_(sh, m.chk.fila, m.chk.col, 1, m.chk.ancho).getA1Notation());
  sh.getRangeList(chks).insertCheckboxes().setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setFontSize(14);

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
  const plano = sh.getRange(FILA_PLANO_INI, U(1), FILA_PLANO_FIN - FILA_PLANO_INI + 1, N_UNIDADES).getA1Notation();
  const cuenta = `COUNTIF(${plano},TRUE)`;
  sh.getRange(CELDA.activas[0], CELDA.activas[1]).setFormula(fx_(`=${cuenta}&" / ${META_MAQUINAS}"`));
  activas.setHorizontalAlignment('left');
  reglas.push(regla().whenFormulaSatisfied(fx_(`=${cuenta}>${META_MAQUINAS}`))
    .setFontColor('#b3261e').setRanges([activas]).build());
  reglas.push(regla().whenFormulaSatisfied(fx_(`=${cuenta}=${META_MAQUINAS}`))
    .setFontColor('#1b7a3a').setRanges([activas]).build());

  [CELDA.alertas, CELDA.pendiente, CELDA.fechaCargada].forEach(([f, c]) => sh.getRange(f, c)
    .setFontColor('#ffffff').setFontSize(6).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP));
  sh.getRange(CELDA.pendiente[0], CELDA.pendiente[1]).setValue(false);

  // ---- Tabla de producción ----
  const n = ORDEN_TABLA.length;
  const fechaA1 = abs(sh.getRange(CELDA.fechaSel[0], CELDA.fechaSel[1]));
  caja_(sh, FILA_TABLA_TITULO, COL0, 1, ANCHO_TOTAL).setBackground(COLOR.cabecera).setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(11);
  sh.getRange(FILA_TABLA_TITULO, COL0).setFormula(fx_(
    `="  INFORMACIÓN DE PRODUCCIÓN · MÁQUINAS ACTIVAS · "&IF(ISNUMBER(${fechaA1}),TEXT(${fechaA1},"dd/mm/yyyy"),"sin fecha")`));

  // Botones: enviar a Planificación, descartar cambios y estado
  sh.getRange(ACCION.enviar[0], ACCION.enviar[1]).insertCheckboxes().setHorizontalAlignment('center')
    .setBackground(COLOR.enviar).setFontColor('#ffffff').setFontSize(12);
  caja_(sh, FILA_ACCIONES, ACCION.enviar[1] + 1, 1, 9).setValue('ENVIAR A PLANIFICACIÓN')
    .setBackground(COLOR.enviar).setFontColor('#ffffff').setFontWeight('bold').setFontSize(11);
  sh.getRange(ACCION.descartar[0], ACCION.descartar[1]).insertCheckboxes().setHorizontalAlignment('center')
    .setFontSize(12);
  caja_(sh, FILA_ACCIONES, ACCION.descartar[1] + 1, 1, 6).setValue('Descartar cambios')
    .setFontWeight('bold').setFontColor('#b3261e');
  caja_(sh, ACCION.estado[0], ACCION.estado[1], 1, ANCHO_TOTAL - ACCION.estado[1] + COL0)
    .setFontStyle('italic').setFontColor(COLOR.etiqueta);
  const pendA1 = abs(sh.getRange(CELDA.pendiente[0], CELDA.pendiente[1]));
  const estadoRg = sh.getRange(ACCION.estado[0], ACCION.estado[1]);

  CAMPOS_TABLA.forEach(x => {
    caja_(sh, FILA_TABLA_CAB, x.col, 1, x.ancho).setValue(x.titulo);
    sh.getRange(FILA_TABLA_INI, x.col, n, x.ancho).mergeAcross().setNumberFormat(x.formato || '0')
      .setBackground(x.entrada ? COLOR.entrada : COLOR.fijo)
      .setHorizontalAlignment(x.formato === '@' ? 'left' : 'center');
  });
  sh.getRange(FILA_TABLA_CAB, COL0, 1, ANCHO_TOTAL).setBackground('#dfe3ea').setFontSize(9).setFontWeight('bold')
    .setFontColor(COLOR.texto).setWrap(true).setHorizontalAlignment('center');
  sh.getRange(FILA_TABLA_CAB, COL0, n + 1, ANCHO_TOTAL)
    .setBorder(true, true, true, true, true, true, COLOR.lineaTabla, SB.SOLID);

  const maq = CAMPO.maquina;
  sh.getRange(FILA_TABLA_INI, maq.col, n, 1).setValues(ORDEN_TABLA.map(id => [id]));
  sh.getRange(FILA_TABLA_INI, maq.col, n, maq.ancho).setFontFamily('Roboto Mono').setFontWeight('bold');

  const num = SpreadsheetApp.newDataValidation().requireNumberGreaterThanOrEqualTo(0).setAllowInvalid(false)
    .setHelpText('Solo números (≥ 0).').build();
  ENTRADAS.filter(x => x.numero).forEach(x => sh.getRange(FILA_TABLA_INI, x.col, n, x.ancho).setDataValidation(num));

  // Estado en rojo si hay cambios sin enviar
  reglas.push(regla().whenFormulaSatisfied(`=${pendA1}`).setFontColor('#b3261e').setBold(true)
    .setRanges([estadoRg]).build());
  // Fila naranja si la máquina tiene alerta de molde; contador final en rojo si es menor que el inicial
  const filaTabla = sh.getRange(FILA_TABLA_INI, COL0, n, ANCHO_TOTAL);
  const maqA1 = '$' + letraCol_(maq.col) + FILA_TABLA_INI;
  reglas.push(regla().whenFormulaSatisfied(fx_(`=ISNUMBER(SEARCH("|"&${maqA1}&"|",${celAlertas}))`))
    .setBackground(COLOR.alerta).setRanges([filaTabla]).build());
  const ini0 = '$' + letraCol_(CAMPO.contIni.col) + FILA_TABLA_INI;
  const fin0 = '$' + letraCol_(CAMPO.contFin.col) + FILA_TABLA_INI;
  reglas.push(regla().whenFormulaSatisfied(fx_(`=AND(ISNUMBER(${ini0}),ISNUMBER(${fin0}),${fin0}<${ini0})`))
    .setFontColor('#b3261e').setBold(true)
    .setRanges([sh.getRange(FILA_TABLA_INI, CAMPO.contFin.col, n, CAMPO.contFin.ancho)]).build());
  sh.setConditionalFormatRules(reglas);

  caja_(sh, FILA_AYUDA, COL0, 1, ANCHO_TOTAL).setWrap(true).setVerticalAlignment('top').setFontSize(9)
    .setFontColor(COLOR.etiqueta)
    .setValue('Cómo usar: 1) Elige la fecha arriba. 2) Marca las máquinas que trabajarán en el turno (una por una, ' +
      'todas o por línea); su fila aparece en la tabla, en orden de número. 3) Completa los campos amarillos. ' +
      '4) Marca «Enviar a Planificación» para guardar todo en la base de datos. Antes de cambiar de fecha envía o ' +
      'descarta los cambios. El tonelaje se recuerda por máquina. Turno 3: la fecha es la del día en que inicia el turno (10:30 PM).');

  sh.hideRows(FILA_TABLA_INI, n);
}

function ajustarTurno_(sh, t) {
  const T = TURNOS[t];
  sh.getRange(1, U(1)).setValue(`MANAPLAS · PLANIFICACIÓN DE PRODUCCIÓN · ${T.hoja.toUpperCase()}${T.cambioMolde ? ' (STANDARD)' : ''}`);
  sh.getRange(2, U(13)).setValue(T.horario);
  sh.getRange(2, U(24)).setValue(T.descanso);
  sh.getRange(CELDA.horas[0], CELDA.horas[1]).setValue(T.horas);
  sh.getRange(1, U(49))
    .setValue(T.cambioMolde ? 'TURNO STANDARD\nCambio de molde permitido' : 'SOLO PRODUCCIÓN\nNo se cambian moldes')
    .setBackground(T.cambioMolde ? '#dcebfb' : '#fde9d6');
  sh.setTabColor(T.color);
}

/** Protección de advertencia: avisa si se intenta escribir sobre el plano o las etiquetas. */
function proteger_(sh) {
  sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
  const p = sh.protect().setDescription('Plano y tabla (solo advertencia)').setWarningOnly(true);
  p.setUnprotectedRanges([
    sh.getRange(CELDA.fechaSel[0], CELDA.fechaSel[1], 1, ANCHO_FECHA),
    sh.getRange(CELDA.todas[0], CELDA.todas[1]),
    sh.getRange(ACCION.enviar[0], ACCION.enviar[1]),
    sh.getRange(ACCION.descartar[0], ACCION.descartar[1]),
    sh.getRange(FILA_TABLA_INI, COL_ENTRADA_INI, ORDEN_TABLA.length, COL_ENTRADA_FIN - COL_ENTRADA_INI + 1),
  ].concat(CELDA.grupos.map(([f, c]) => sh.getRange(f, c)))
   .concat(maquinas_().map(m => sh.getRange(m.chk.fila, m.chk.col, 1, m.chk.ancho))));
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
function vacio_(v) { return v === '' || v === null; }

function ultimaFilaPlan_(sh) {
  const v = sh.getRange(1, 1, sh.getMaxRows(), 3).getValues();
  for (let i = v.length - 1; i >= 0; i--) if (v[i].some(x => !vacio_(x))) return i + 1;
  return 1;
}

/** Lee la hoja Planificación como una lista de registros. */
function leerPlan_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(HOJA_PLAN);
  if (!sh) return [];
  const ult = ultimaFilaPlan_(sh);
  if (ult < 2) return [];
  const out = [];
  sh.getRange(2, 1, ult - 1, N_PLAN).getValues().forEach((r, i) => {
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

/** Estado actual de las casillas del plano: { id: true/false }. */
function leerCasillas_(sh) {
  const plano = sh.getRange(FILA_PLANO_INI, COL0, FILA_PLANO_FIN - FILA_PLANO_INI + 1, ULTIMA_COL - COL0).getValues();
  const out = {};
  maquinas_().forEach(m => out[m.id] = plano[m.chk.fila - FILA_PLANO_INI][m.chk.col - COL0] === true);
  return out;
}

/** Muestra solo las filas de la tabla de las máquinas activas (por tramos seguidos). */
function mostrarFilas_(sh, activa) {
  let i = 0;
  while (i < ORDEN_TABLA.length) {
    const visible = !!activa[ORDEN_TABLA[i]];
    let j = i;
    while (j + 1 < ORDEN_TABLA.length && !!activa[ORDEN_TABLA[j + 1]] === visible) j++;
    if (visible) sh.showRows(FILA_TABLA_INI + i, j - i + 1);
    else sh.hideRows(FILA_TABLA_INI + i, j - i + 1);
    i = j + 1;
  }
}

/** Casillas «todas» y «Línea N» marcadas solo si todas sus máquinas están activas. */
function actualizarGrupos_(sh, activa) {
  const ids = Object.keys(activa).map(Number);
  sh.getRange(CELDA.todas[0], CELDA.todas[1]).setValue(ids.every(id => activa[id]));
  GRUPOS.forEach((g, i) => sh.getRange(CELDA.grupos[i][0], CELDA.grupos[i][1])
    .setValue(ids.filter(id => id >= g.desde && id <= g.hasta).every(id => activa[id])));
}

/**
 * Tras marcar o desmarcar casillas: muestra/oculta las filas de la tabla y, en las filas
 * activas sin tonelaje, pone el tonelaje guardado de la máquina.
 */
function aplicarSeleccion_(sh, t) {
  const activa = leerCasillas_(sh);
  const ton = tonelajes_();
  const c = CAMPO.ton.col;
  const rg = sh.getRange(FILA_TABLA_INI, c, ORDEN_TABLA.length, 1);
  const valores = rg.getValues();
  let cambio = false;
  ORDEN_TABLA.forEach((id, i) => {
    if (activa[id] && vacio_(valores[i][0]) && ton[id] !== undefined) { valores[i][0] = ton[id]; cambio = true; }
  });
  if (cambio) rg.setValues(valores);
  mostrarFilas_(sh, activa);
  actualizarGrupos_(sh, activa);
}

function marcarPendiente_(sh) {
  sh.getRange(CELDA.pendiente[0], CELDA.pendiente[1]).setValue(true);
  sh.getRange(ACCION.estado[0], ACCION.estado[1])
    .setValue('⚠ Hay cambios sin enviar: marca «Enviar a Planificación» para guardarlos.');
}

/**
 * Carga en la hoja lo guardado en Planificación para la fecha y turno: casillas, tabla,
 * filas visibles y alertas. Descarta los cambios sin enviar.
 */
function cargar_(sh, t, datos) {
  const fechaV = sh.getRange(CELDA.fechaSel[0], CELDA.fechaSel[1]).getValue();
  const rec = registros_(datos, claveFecha_(fechaV), t);
  const ton = tonelajes_();

  const si = [], no = [], activa = {};
  maquinas_().forEach(m => {
    activa[m.id] = !!rec[String(m.id)];
    (activa[m.id] ? si : no).push(casilla_(sh, m).getA1Notation());
  });
  if (si.length) sh.getRangeList(si).check();
  if (no.length) sh.getRangeList(no).uncheck();

  const ancho = COL_ENTRADA_FIN - COL_ENTRADA_INI + 1;
  sh.getRange(FILA_TABLA_INI, COL_ENTRADA_INI, ORDEN_TABLA.length, ancho).setValues(ORDEN_TABLA.map(id => {
    const fila = new Array(ancho).fill('');
    const r = rec[String(id)];
    if (r) ENTRADAS.forEach(x => fila[x.col - COL_ENTRADA_INI] = r.valores[IDX[x.k]]);
    const iTon = CAMPO.ton.col - COL_ENTRADA_INI;
    if (r && vacio_(fila[iTon]) && ton[id] !== undefined) fila[iTon] = ton[id];
    return fila;
  }));

  mostrarFilas_(sh, activa);
  actualizarGrupos_(sh, activa);
  pintarPlano_(sh, t, datos);
  sh.getRange(CELDA.pendiente[0], CELDA.pendiente[1]).setValue(false);
  sh.getRange(CELDA.fechaCargada[0], CELDA.fechaCargada[1]).setValue(fechaV);
  sh.getRange(ACCION.estado[0], ACCION.estado[1]).setValue(si.length
    ? `Datos cargados desde Planificación: ${si.length} máquina(s).`
    : 'Sin datos enviados para esta fecha y turno.');
}

/** Cambio de fecha: si hay cambios sin enviar, se vuelve a la fecha anterior para no perderlos. */
function cambiarFecha_(sh, t) {
  const pendiente = sh.getRange(CELDA.pendiente[0], CELDA.pendiente[1]).getValue() === true;
  const anterior = sh.getRange(CELDA.fechaCargada[0], CELDA.fechaCargada[1]).getValue();
  if (pendiente && claveFecha_(anterior)) {
    sh.getRange(CELDA.fechaSel[0], CELDA.fechaSel[1]).setValue(anterior);
    SpreadsheetApp.getActive().toast('Tienes cambios sin enviar en esta fecha. Marca «Enviar a Planificación» ' +
      'o «Descartar cambios» y luego cambia la fecha.', 'No se cambió la fecha', 8);
    return;
  }
  cargar_(sh, t, leerPlan_());
}

/**
 * Envía la selección y la tabla a Planificación para la fecha y turno de la hoja:
 *   máquina marcada    → crea o actualiza su registro
 *   máquina desmarcada → borra su registro si existía
 */
function enviar_(sh, t) {
  const ss = SpreadsheetApp.getActive();
  const fechaV = sh.getRange(CELDA.fechaSel[0], CELDA.fechaSel[1]).getValue();
  const clave = claveFecha_(fechaV);
  if (!clave) {
    ss.toast('Elige una fecha válida arriba antes de enviar.', 'No se envió', 5);
    return;
  }

  const activa = leerCasillas_(sh);
  const tabla = sh.getRange(FILA_TABLA_INI, COL0, ORDEN_TABLA.length, ANCHO_TOTAL).getValues();
  const filaDe = id => tabla[ORDEN_TABLA.indexOf(id)];
  const val = (id, k) => filaDe(id)[CAMPO[k].col - COL0];

  const errores = ORDEN_TABLA.filter(id => activa[id] && typeof val(id, 'contIni') === 'number' &&
    typeof val(id, 'contFin') === 'number' && val(id, 'contFin') < val(id, 'contIni'));
  if (errores.length) {
    ss.toast(`Máquina(s) ${errores.join(', ')}: el contador final es menor que el inicial. Corrígelo y vuelve a enviar.`,
      'No se envió', 8);
    return;
  }

  const plan = ss.getSheetByName(HOJA_PLAN);
  const rec = registros_(leerPlan_(), clave, t);
  const nuevas = [], borrar = [];
  let actualizadas = 0;
  ORDEN_TABLA.forEach(id => {
    const r = rec[String(id)];
    if (!activa[id]) {
      if (r) borrar.push(r.fila);
      return;
    }
    const fila = new Array(N_PLAN).fill('');
    fila[IDX.fecha] = fechaV; fila[IDX.turno] = t; fila[IDX.maquina] = id;
    ENTRADAS.forEach(x => fila[IDX[x.k]] = val(id, x.k));
    if (r) {
      plan.getRange(r.fila, 1, 1, N_PLAN).setValues([fila]);
      actualizadas++;
    } else {
      nuevas.push(fila);
    }
  });
  if (nuevas.length) {
    const ini = ultimaFilaPlan_(plan) + 1;
    const fin = ini + nuevas.length - 1;
    if (fin > plan.getMaxRows()) plan.insertRowsAfter(plan.getMaxRows(), fin - plan.getMaxRows() + 200);
    plan.getRange(ini, 1, nuevas.length, N_PLAN).setValues(nuevas);
  }
  borrar.sort((a, b) => b - a).forEach(f => plan.deleteRow(f));

  // El tonelaje queda guardado por máquina para las próximas veces
  const listas = ss.getSheetByName(HOJA_LISTAS);
  const ids = listas.getRange(2, 1, Math.max(listas.getLastRow() - 1, 1), 1).getValues().map(r => Number(r[0]));
  ORDEN_TABLA.forEach(id => {
    const ton = val(id, 'ton'), i = ids.indexOf(id);
    if (activa[id] && typeof ton === 'number' && i >= 0) listas.getRange(i + 2, 2).setValue(ton);
  });

  sh.getRange(CELDA.pendiente[0], CELDA.pendiente[1]).setValue(false);
  sh.getRange(CELDA.fechaCargada[0], CELDA.fechaCargada[1]).setValue(fechaV);
  pintarTodos_(leerPlan_());
  const hora = Utilities.formatDate(new Date(), tz(), 'h:mm a');
  const total = nuevas.length + actualizadas;
  sh.getRange(ACCION.estado[0], ACCION.estado[1]).setValue(`✓ Enviado a Planificación a las ${hora}: ${total} máquina(s)` +
    (borrar.length ? `, ${borrar.length} quitada(s).` : '.'));
  ss.toast(`${nuevas.length} nueva(s), ${actualizadas} actualizada(s), ${borrar.length} quitada(s).`, 'Enviado a Planificación', 5);
}

function pintarTodos_(datos) {
  [1, 2, 3, 4].forEach(t => {
    const sh = SpreadsheetApp.getActive().getSheetByName(TURNOS[t].hoja);
    if (sh) pintarPlano_(sh, t, datos);
  });
}

/** Recarga las hojas de turno desde Planificación (las que tienen cambios sin enviar solo actualizan alertas). */
function actualizarTodo_() {
  const datos = leerPlan_();
  [1, 2, 3, 4].forEach(t => {
    const sh = SpreadsheetApp.getActive().getSheetByName(TURNOS[t].hoja);
    if (!sh) return;
    if (sh.getRange(CELDA.pendiente[0], CELDA.pendiente[1]).getValue() === true) pintarPlano_(sh, t, datos);
    else cargar_(sh, t, datos);
  });
}
