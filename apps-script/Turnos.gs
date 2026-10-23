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
 *   Turno 1 · Turno 2 · Turno 3 · Turno 4  → plano de máquinas + asignación de personal
 *   Planificación                           → base de datos (una fila por fecha/turno/máquina)
 *   Personal por turno                      → base de datos (una fila por fecha/turno)
 *   Listas (oculta)                         → máquinas y su tonelaje
 *
 * Uso:
 *   · Casilla de cada máquina        → la marca como activa y agrega su fila a la tabla.
 *   · «Marcar / desmarcar todas»     → todas las máquinas. «Línea 1/2/3» → solo esa línea.
 *   · Tabla de máquinas activas      → una fila por máquina activa, en orden de número, con
 *                                      sus operarios.
 *   · Personal del turno             → puestos de la propuesta de turnos (distintos en el Turno 4).
 *   · Resumen (fila 4)               → lo asignado frente a lo que pide la propuesta.
 *   · «Enviar a Planificación»       → guarda máquinas y personal de la fecha y el turno.
 *   · «Descartar cambios»            → vuelve a lo último enviado para esa fecha y turno.
 */

// ======================= CONFIGURACIÓN =======================
const HOJA_PLAN = 'Planificación';
const HOJA_PERSONAL = 'Personal por turno';
const HOJA_LISTAS = 'Listas';

const TURNOS = {
  1: { hoja: 'Turno 1', horario: '6:00 AM – 2:30 PM',  descanso: '1 h entre 9:30 AM y 12:00 PM', horas: 7.5, color: '#f4a261',
       personal: 'produccion', metas: { maquinas: 12, operarios: 16, produccion: 21, tecnicos: 6 } },
  2: { hoja: 'Turno 2', horario: '2:30 PM – 10:30 PM', descanso: '1 h entre 5:00 PM y 7:00 PM',  horas: 8,   color: '#2a9d8f',
       personal: 'produccion', metas: { maquinas: 12, operarios: 16, produccion: 21, tecnicos: 6 } },
  3: { hoja: 'Turno 3', horario: '10:30 PM – 6:00 AM', descanso: '1 h entre 2:00 AM y 3:00 AM',  horas: 6.5, color: '#264653',
       personal: 'produccion', metas: { maquinas: 12, operarios: 16, produccion: 21, tecnicos: 6 } },
  4: { hoja: 'Turno 4', horario: '7:00 AM – 4:00 PM',  descanso: '1 h de 12:00 PM a 1:00 PM',    horas: 8,   color: '#1f6fd1',
       personal: 'adicional', metas: { tecnicos: 5 }, standard: true },
};

// Columnas de la hoja Planificación (A:M)
const COLS_PLAN = ['Fecha', 'Turno', 'Máquina', 'Ton', 'OP', 'Producto', 'Referencia', 'Color', 'Material',
  'Operario 1', 'Operario 2', 'Operario 3'];
const IDX = { fecha: 0, turno: 1, maquina: 2, ton: 3, op: 4, producto: 5, referencia: 6, color: 7, material: 8,
  operario1: 9, operario2: 10, operario3: 11 };
const N_PLAN = COLS_PLAN.length;

// Personal del turno según la propuesta de turnos (cuadro al lado de la tabla).
// grupo: 'produccion' (cuenta como personal de producción) o 'tecnico' (personal técnico).
const PERSONAL = {
  produccion: [   // Turnos 1, 2 y 3
    { titulo: 'Supervisor de turno', n: 1, grupo: 'produccion' },
    { titulo: 'Asistente de producción', n: 2, grupo: 'produccion' },
    { titulo: 'Puente y distribución (montacargas)', n: 2, grupo: 'produccion' },
    { titulo: 'Técnico de inyección', n: 2, grupo: 'tecnico' },
    { titulo: 'Técnico eléctrico', n: 1, grupo: 'tecnico' },
    { titulo: 'Técnico hidráulico', n: 1, grupo: 'tecnico' },
    { titulo: 'Asistente de inyección', n: 2, grupo: 'tecnico' },
  ],
  adicional: [    // Turno 4
    { titulo: 'Técnico de molde', n: 2, grupo: 'tecnico' },
    { titulo: 'Molino', n: 1, grupo: 'tecnico' },
    { titulo: 'Pintura / Mezcla', n: 2, grupo: 'tecnico' },
  ],
};
const puestos_ = t => PERSONAL[TURNOS[t].personal];
const nNombres_ = t => puestos_(t).reduce((s, p) => s + p.n, 0);
const N_NOMBRES_MAX = Math.max(...Object.keys(PERSONAL).map(k => PERSONAL[k].reduce((s, p) => s + p.n, 0)));
// Hoja «Personal por turno»: una fila por persona
const COLS_PERSONAL = ['Fecha', 'Turno', 'Puesto', 'N°', 'Nombre'];

// Grupos de selección rápida
const GRUPOS = [
  { nombre: 'Línea 1', desde: 101, hasta: 143 },
  { nombre: 'Línea 2', desde: 201, hasta: 218 },
  { nombre: 'Línea 3', desde: 301, hasta: 306 },
];

const COLOR = {
  inactiva: '#eef1f5', activa: '#b9e4c1',
  borde: '#1f2a37', texto: '#1f2a37', etiqueta: '#5b6676',
  cabecera: '#1f2a37', entrada: '#fff8dc', fijo: '#f1f3f6', lineaTabla: '#d0d5dd', enviar: '#1f6fd1',
};

// ======================= CUADRÍCULA =======================
// Columna A = margen. El plano usa "unidades" de 37 px desde la columna B (45 unidades de ancho):
//   1–33  Línea 2 (201–214) y Línea 1 (101–123 y, al lado, 125–143) · 36–38  Máquinas 218–215
//   40–45  Línea 3 (Soplado)
// Todas las máquinas tienen el mismo tamaño (3 × 2 celdas) y el plano cabe en una pantalla
// Full HD sin desplazarse.
const ANCHO_UNIDAD = 37;
const COL0 = 2;
const N_UNIDADES = 45;
const U = u => COL0 + u - 1;
const COL_125 = U(23);              // cuadrícula de 4 × 3 máquinas (125–143), entre 101–123 y 218–215
const COL_DER = U(36);              // columna 218–215
const S = u => U(39) + u;           // zona Soplado: unidades 1..6
const ULTIMA_COL = U(N_UNIDADES) + 1; // columna de margen derecho
const ANCHO_TOTAL = N_UNIDADES;
// Las máquinas van en "franjas" de 3 filas: número, casilla y un separador fino, así el espacio
// entre filas de máquinas es siempre el mismo. Hay 6 franjas (filas 6–22).
// Encabezado: título (1) · tarjetas de datos (3 etiqueta, 4 valor) · resumen (6) · títulos del plano (8)
const FILA_ETIQ = 3, FILA_VALOR = 4, FILA_RESUMEN = 6, FILA_SECCION = 8;
const FILA_PLANO_INI = 9;
const franja_ = k => FILA_PLANO_INI + 3 * k;
const FILA_PLANO_FIN = franja_(5) + 1;
const ALTO_FILA_PLANO = 36, ALTO_SEPARADOR = 8;

const CELDA = {
  fechaSel: [FILA_VALOR, U(1)], activas: [FILA_VALOR, U(27)], horas: [FILA_VALOR, U(22)],
  todas: [FILA_ETIQ, U(33)], grupos: [[FILA_VALOR, U(33)], [FILA_VALOR, U(37)], [FILA_VALOR, U(41)]],
  // Celdas auxiliares en la columna A (texto invisible)
  pendiente: [FILA_PLANO_FIN + 3, 1],   // TRUE si hay cambios sin enviar
  fechaCargada: [FILA_PLANO_FIN + 4, 1], // fecha cuyos datos muestra la hoja
};
const ANCHO_FECHA = 5;
// Tarjetas del encabezado: [título, columna, ancho]
const TARJETAS = { fecha: ['FECHA', U(1), 5], horario: ['HORARIO', U(7), 6], descanso: ['DESCANSO', U(14), 7],
  horas: ['HORAS EFECTIVAS', U(22), 4], activas: ['MÁQ. ACTIVAS', U(27), 4] };

// Máquinas en uso, en el orden en que aparecen en el plano
const LINEA2 = [202, 203, 204, 206, 207, 209, 210, 211, 212, 213, 214];   // fila bajo la 201
const LINEA1 = [101, 102, 103, 107, 108, 109, 110, 111, 113, 116, 117, 119, 121, 123]; // dos filas de 7
const BLOQUE_125 = [125, 127, 129, 131, 132, 133, 134, 135, 136, 139, 140, 143];        // tres filas de 4

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
  // Bloque 201–218: 201 encima de la 202, la fila 202–214 y 218–215 en columna a la derecha
  q(201, franja_(0), U(1));
  LINEA2.forEach((id, j) => q(id, franja_(1), U(3 * j + 1)));
  q(218, franja_(0), COL_DER); q(217, franja_(1), COL_DER); q(216, franja_(2), COL_DER); q(215, franja_(3), COL_DER);
  // Bloque 101–123 en dos filas de 7 (la franja 2 queda para el título)
  LINEA1.forEach((id, i) => q(id, franja_(3 + Math.floor(i / 7)), U(3 * (i % 7) + 1)));
  // Bloque 125–143: tres filas de 4 a la derecha de 101–123, empezando a su misma altura
  BLOQUE_125.forEach((id, i) => q(id, franja_(3 + Math.floor(i / 4)), COL_125 + 3 * (i % 4)));
  // Línea 3 (Soplado), junto a la columna 218–215
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

// ======================= TABLA DE MÁQUINAS Y PERSONAL =======================
// Las máquinas activas se escriben una debajo de otra, en orden de número (sin filas ocultas en
// medio). A la derecha va el cuadro con el personal del turno.
const FILA_ACCIONES = FILA_PLANO_FIN + 5, FILA_TABLA_TITULO = FILA_ACCIONES + 1,
  FILA_TABLA_CAB = FILA_ACCIONES + 2, FILA_TABLA_INI = FILA_ACCIONES + 3;
const ORDEN_TABLA = maquinas_().map(m => m.id).sort((a, b) => a - b);
const N_FILAS_TABLA = Math.max(ORDEN_TABLA.length, N_NOMBRES_MAX);
const FILA_TABLA_FIN = FILA_TABLA_INI + N_FILAS_TABLA - 1;
const FILA_AYUDA = FILA_TABLA_FIN + 2;
const FILAS_HOJA = FILA_AYUDA + 1;

const ACCION = { enviar: [FILA_ACCIONES, U(1)], descartar: [FILA_ACCIONES, U(10)], estado: [FILA_ACCIONES, U(17)] };

// k = campo de Planificación; ancho en unidades (la tabla ocupa las unidades 1–36).
const CAMPOS_TABLA = [
  { k: 'maquina',    titulo: 'MÁQ.',                 ancho: 2 },
  { k: 'ton',        titulo: 'TON',                  ancho: 2, entrada: true, formato: '0', numero: true },
  { k: 'op',         titulo: 'OP',                   ancho: 3, entrada: true, formato: '@' },
  { k: 'producto',   titulo: 'PRODUCTO',             ancho: 7, entrada: true, formato: '@' },
  { k: 'referencia', titulo: 'REFERENCIA',           ancho: 4, entrada: true, formato: '@' },
  { k: 'color',      titulo: 'COLOR',                ancho: 3, entrada: true, formato: '@' },
  { k: 'material',   titulo: 'MATERIAL',             ancho: 3, entrada: true, formato: '@' },
  { k: 'operario1',  titulo: 'OPERARIO 1',           ancho: 4, entrada: true, formato: '@' },
  { k: 'operario2',  titulo: 'OPERARIO 2',           ancho: 4, entrada: true, formato: '@' },
  { k: 'operario3',  titulo: 'OPERARIO 3',           ancho: 4, entrada: true, formato: '@' },
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

// Cuadro de personal: puesto (4 unidades) y nombre (4 unidades), a la derecha de la tabla
const COL_PUESTO = U(ANCHO_TABLA + 2), COL_NOMBRE = COL_PUESTO + 4, ANCHO_PERSONAL = 8;

// Resumen frente a la propuesta (fila 4): etiqueta de 3 unidades y valor de 2
const RESUMEN = [
  { k: 'maquinas',   titulo: 'MÁQUINAS',   col: U(6) },
  { k: 'operarios',  titulo: 'OPERARIOS',  col: U(12) },
  { k: 'produccion', titulo: 'PRODUCCIÓN', col: U(18) },
  { k: 'tecnicos',   titulo: 'TÉCNICOS',   col: U(24) },
];

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
               ((c <= COL_ENTRADA_FIN && c + rg.getNumColumns() - 1 >= COL_ENTRADA_INI) ||
                (c <= COL_NOMBRE + 3 && c + rg.getNumColumns() - 1 >= COL_NOMBRE))) {
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
  prepararPersonal_(ss);

  [1, 2, 3, 4].forEach(t => {
    const s = ss.getSheetByName(TURNOS[t].hoja);
    if (s) ss.deleteSheet(s);
  });

  // Se construye el Turno 4 y se copia para los demás (más rápido)
  const plantilla = ss.insertSheet(TURNOS[4].hoja);
  construirHojaTurno_(plantilla);
  SpreadsheetApp.flush();
  [1, 2, 3].forEach(t => plantilla.copyTo(ss).setName(TURNOS[t].hoja));

  [1, 2, 3, 4].forEach((t, i) => {
    const s = ss.getSheetByName(TURNOS[t].hoja);
    ajustarTurno_(s, t);
    proteger_(s, t);
    ss.setActiveSheet(s);
    ss.moveActiveSheet(i + 1);
  });
  ss.setActiveSheet(ss.getSheetByName(HOJA_PLAN));
  ss.moveActiveSheet(5);
  ss.setActiveSheet(ss.getSheetByName(HOJA_PERSONAL));
  ss.moveActiveSheet(6);

  // Elimina hojas vacías por defecto ("Hoja 1", "Sheet1"...)
  const nuestras = [HOJA_PLAN, HOJA_PERSONAL, HOJA_LISTAS].concat([1, 2, 3, 4].map(t => TURNOS[t].hoja));
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

/**
 * Reacomoda una hoja de datos a las columnas `cols` buscando cada columna por su nombre en la
 * fila 1: conserva los datos de las columnas que siguen existiendo y quita las demás.
 */
function migrarColumnas_(sh, cols) {
  const ultCol = sh.getLastColumn(), ultFila = sh.getLastRow();
  const cab = ultCol ? sh.getRange(1, 1, 1, ultCol).getValues()[0].map(String) : [];
  if (cab.length === cols.length && cols.every((c, i) => cab[i] === c)) return;
  const datos = (ultFila >= 2 && ultCol) ? sh.getRange(2, 1, ultFila - 1, ultCol).getValues() : [];
  const desde = cols.map(c => cab.indexOf(c));
  sh.clear();
  sh.getRange(1, 1, 1, sh.getMaxColumns()).setDataValidation(null);
  if (sh.getMaxColumns() < cols.length) sh.insertColumnsAfter(sh.getMaxColumns(), cols.length - sh.getMaxColumns());
  sh.getRange(1, 1, 1, cols.length).setValues([cols]);
  const filas = datos.filter(r => r.some(x => !vacio_(x))).map(r => desde.map(i => i >= 0 ? r[i] : ''));
  if (filas.length) {
    if (filas.length + 1 > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), filas.length + 1 - sh.getMaxRows());
    sh.getRange(2, 1, filas.length, cols.length).setValues(filas);
  }
}

/** Formato de la fila de títulos de una hoja de datos. */
function cabecera_(sh, n) {
  sh.getRange(1, 1, 1, n).setFontWeight('bold').setFontColor('#ffffff').setBackground(COLOR.cabecera)
    .setWrap(true).setVerticalAlignment('middle').setHorizontalAlignment('center')
    .setBorder(true, true, true, true, true, true, '#ffffff', SpreadsheetApp.BorderStyle.SOLID);
  sh.setFrozenRows(1);
  sh.setRowHeight(1, 36);
}

function prepararPlan_(ss) {
  let sh = ss.getSheetByName(HOJA_PLAN);
  if (!sh) sh = ss.insertSheet(HOJA_PLAN);
  migrarColumnas_(sh, COLS_PLAN);
  cabecera_(sh, N_PLAN);

  const n = sh.getMaxRows() - 1;
  const col = k => IDX[k] + 1;
  const dv = SpreadsheetApp.newDataValidation;
  sh.getRange(2, 1, n, sh.getMaxColumns()).setDataValidation(null);
  sh.getRange(2, col('fecha'), n, 1).setNumberFormat('dd/mm/yyyy')
    .setDataValidation(dv().requireDate().setAllowInvalid(false).build());
  sh.getRange(2, col('turno'), n, 1).setHorizontalAlignment('center')
    .setDataValidation(dv().requireValueInList(['1', '2', '3', '4'], true).setAllowInvalid(false).build());
  sh.getRange(2, col('maquina'), n, 1).setHorizontalAlignment('center').setNumberFormat('0')
    .setDataValidation(dv().requireValueInRange(ss.getSheetByName(HOJA_LISTAS).getRange('A2:A200'), true)
      .setAllowInvalid(false).build());
  sh.getRange(2, col('op'), n, N_PLAN - IDX.op).setNumberFormat('@');
  sh.getRange(2, col('ton'), n, 1).setNumberFormat('0').setHorizontalAlignment('center')
    .setDataValidation(dv().requireNumberGreaterThanOrEqualTo(0).setAllowInvalid(false)
      .setHelpText('Solo números (≥ 0).').build());

  const anchos = [90, 55, 70, 60, 100, 220, 120, 100, 120, 150, 150, 150, 170];
  anchos.forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.setTabColor('#5b6676');
}

function prepararPersonal_(ss) {
  let sh = ss.getSheetByName(HOJA_PERSONAL);
  if (!sh) sh = ss.insertSheet(HOJA_PERSONAL);
  // Formato anterior (una columna por puesto): los puestos cambiaron, se empieza de nuevo
  const cab = sh.getLastColumn() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String) : [];
  if (cab.join('|') !== COLS_PERSONAL.join('|')) {
    sh.clear();
    sh.getRange(1, 1, 1, COLS_PERSONAL.length).setValues([COLS_PERSONAL]);
  }
  cabecera_(sh, COLS_PERSONAL.length);
  const n = sh.getMaxRows() - 1;
  sh.getRange(2, 1, n, 1).setNumberFormat('dd/mm/yyyy');
  sh.getRange(2, 2, n, 1).setHorizontalAlignment('center');
  sh.getRange(2, 3, n, 1).setNumberFormat('@');
  sh.getRange(2, 4, n, 1).setHorizontalAlignment('center');
  sh.getRange(2, 5, n, 1).setNumberFormat('@');
  [90, 55, 260, 45, 220].forEach((w, i) => sh.setColumnWidth(i + 1, w));
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
  const alto = { 1: 40, 2: 10, [FILA_ETIQ]: 22, [FILA_VALOR]: 30, 5: 10, [FILA_RESUMEN]: 30, 7: 10, [FILA_SECCION]: 24,
    [FILA_PLANO_FIN + 1]: 8,
    [FILA_PLANO_FIN + 2]: 6, [CELDA.pendiente[0]]: 6, [CELDA.fechaCargada[0]]: 6,
    [FILA_TABLA_TITULO]: 28, [FILA_ACCIONES]: 32, [FILA_TABLA_CAB]: 34, [FILA_TABLA_FIN + 1]: 10, [FILA_AYUDA]: 36 };
  for (let f = FILA_PLANO_INI; f <= FILA_PLANO_FIN; f++) {
    alto[f] = (f - FILA_PLANO_INI) % 3 === 2 ? ALTO_SEPARADOR : ALTO_FILA_PLANO;
  }
  Object.keys(alto).forEach(f => sh.setRowHeight(Number(f), alto[f]));
  sh.setRowHeights(FILA_TABLA_INI, ORDEN_TABLA.length, 26);
  sh.getRange(1, 1, FILAS_HOJA, ULTIMA_COL).setFontFamily('Roboto').setFontSize(10)
    .setFontColor(COLOR.texto).setVerticalAlignment('middle').setBackground('#ffffff');

  // ---- Encabezado ----
  // Barra de título
  caja_(sh, 1, U(1), 1, N_UNIDADES).setFontSize(16).setFontWeight('bold').setFontColor('#ffffff')
    .setBackground(COLOR.cabecera);

  // Tarjetas de datos: etiqueta arriba (gris) y valor abajo, con recuadro
  Object.keys(TARJETAS).forEach(k => {
    const [titulo, col, ancho] = TARJETAS[k];
    caja_(sh, FILA_ETIQ, col, 1, ancho).setValue('  ' + titulo).setFontSize(8).setFontWeight('bold')
      .setFontColor(COLOR.etiqueta).setBackground(COLOR.fijo);
    caja_(sh, FILA_VALOR, col, 1, ancho).setFontSize(12).setFontWeight('bold').setHorizontalAlignment('center');
    sh.getRange(FILA_ETIQ, col, 2, ancho).setBorder(true, true, true, true, null, null, '#c5ccd6', SB.SOLID);
  });
  const sel = sh.getRange(CELDA.fechaSel[0], CELDA.fechaSel[1], 1, ANCHO_FECHA);
  sel.setNumberFormat('dd/mm/yyyy').setBackground(COLOR.entrada)
    .setBorder(true, true, true, true, null, null, '#c9a227', SB.SOLID_MEDIUM)
    .setDataValidation(SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false)
      .setHelpText('Fecha a ver o planificar (doble clic abre el calendario).').build());
  const hoy = new Date();
  sel.setValue(new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate()));
  sh.getRange(CELDA.horas[0], CELDA.horas[1]).setNumberFormat('0.0 "h"');
  const activas = sh.getRange(CELDA.activas[0], CELDA.activas[1], 1, TARJETAS.activas[2]);

  // Selección rápida: todas (arriba) y por línea (abajo), en un recuadro
  const marco = sh.getRange(FILA_ETIQ, U(32), 2, N_UNIDADES - 31);
  marco.setBackground(COLOR.fijo).setBorder(true, true, true, true, null, null, '#c5ccd6', SB.SOLID);
  sh.getRange(CELDA.todas[0], CELDA.todas[1]).insertCheckboxes().setHorizontalAlignment('center');
  caja_(sh, CELDA.todas[0], CELDA.todas[1] + 1, 1, 12).setValue('Marcar / desmarcar todas')
    .setFontSize(10).setFontWeight('bold');
  GRUPOS.forEach((g, i) => {
    const [f, c] = CELDA.grupos[i];
    sh.getRange(f, c).insertCheckboxes().setHorizontalAlignment('center');
    caja_(sh, f, c + 1, 1, 3).setValue(`${g.nombre} (${g.desde}–${g.hasta})`).setFontSize(9).setFontWeight('bold');
  });

  // Franja del resumen frente a la propuesta (los valores los pone ajustarTurno_) y leyenda
  sh.getRange(FILA_RESUMEN, U(1), 1, N_UNIDADES).setBackground(COLOR.fijo)
    .setBorder(true, true, true, true, null, null, '#c5ccd6', SB.SOLID);
  caja_(sh, FILA_RESUMEN, U(1), 1, 4).setValue('  PROPUESTA').setFontSize(9).setFontWeight('bold')
    .setFontColor('#ffffff').setBackground(COLOR.cabecera);
  RESUMEN.forEach(x => {
    etiqueta_(caja_(sh, FILA_RESUMEN, x.col, 1, 3), x.titulo).setHorizontalAlignment('right');
    caja_(sh, FILA_RESUMEN, x.col + 3, 1, 2).setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center');
  });
  [[U(36), COLOR.activa, 'Activa (trabaja)'], [U(41), COLOR.inactiva, 'Sin planificar']].forEach(([c, col, txt]) => {
    sh.getRange(FILA_RESUMEN, c).setBackground(col).setBorder(true, true, true, true, null, null, COLOR.borde, SB.SOLID);
    caja_(sh, FILA_RESUMEN, c + 1, 1, 4).setValue(txt).setFontSize(9).setFontColor(COLOR.texto);
  });

  // ---- Plano ----
  seccion_(caja_(sh, FILA_SECCION, U(1), 1, 20), 'LÍNEA 2 · MÁQUINAS 201–218');
  seccion_(caja_(sh, franja_(2) + 1, U(1), 1, 20), 'LÍNEA 1 · MÁQUINAS 101–143');
  seccion_(caja_(sh, FILA_SECCION, S(1), 1, 6), ' LÍNEA 3 · SOPLADO');
  sh.getRange(FILA_SECCION, S(1), franja_(2) + 2 - FILA_SECCION, 6).setBorder(true, true, true, true, null, null, '#7a8594', SB.DASHED);

  const ms = maquinas_();
  const bloques = ms.map(m => sh.getRange(m.fila, m.col, m.alto, m.ancho));
  sh.getRangeList(bloques.map(r => r.getA1Notation())).setBackground(COLOR.inactiva)
    .setBorder(true, true, true, true, null, null, COLOR.borde, SB.SOLID_MEDIUM);

  const nums = ms.map(m => caja_(sh, m.num.fila, m.num.col, m.num.alto, m.num.ancho));
  sh.getRangeList(nums.map(r => r.getA1Notation())).setNumberFormat('@')
    .setFontFamily('Roboto Mono').setFontWeight('bold').setFontSize(14)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  ms.forEach((m, i) => nums[i].setValue(String(m.id)));

  const chks = ms.map(m => caja_(sh, m.chk.fila, m.chk.col, 1, m.chk.ancho).getA1Notation());
  sh.getRangeList(chks).insertCheckboxes().setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setFontSize(14);

  // Colores del plano con formato condicional (se ven al instante al marcar una casilla).
  const abs = r => r.getA1Notation().replace(/^([A-Z]+)(\d+)$/, '$$$1$$$2');
  const regla = SpreadsheetApp.newConditionalFormatRule;
  const reglas = [];
  ms.forEach((m, i) => {
    const chk = abs(casilla_(sh, m));
    reglas.push(regla().whenFormulaSatisfied(`=${chk}`)
      .setBackground(COLOR.activa).setRanges([bloques[i]]).build());
  });

  // Contador de máquinas activas: casillas marcadas del plano / total de máquinas
  const plano = sh.getRange(FILA_PLANO_INI, U(1), FILA_PLANO_FIN - FILA_PLANO_INI + 1, N_UNIDADES).getA1Notation();
  sh.getRange(CELDA.activas[0], CELDA.activas[1])
    .setFormula(fx_(`=COUNTIF(${plano},TRUE)&" / ${ORDEN_TABLA.length}"`));
  activas.setHorizontalAlignment('center');

  [CELDA.pendiente, CELDA.fechaCargada].forEach(([f, c]) => sh.getRange(f, c)
    .setFontColor('#ffffff').setFontSize(6).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP));
  sh.getRange(CELDA.pendiente[0], CELDA.pendiente[1]).setValue(false);

  // ---- Tabla de máquinas activas y cuadro de personal ----
  const n = N_FILAS_TABLA;
  const fechaA1 = abs(sh.getRange(CELDA.fechaSel[0], CELDA.fechaSel[1]));
  caja_(sh, FILA_TABLA_TITULO, COL0, 1, ANCHO_TABLA).setBackground(COLOR.cabecera).setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(11);
  sh.getRange(FILA_TABLA_TITULO, COL0).setFormula(fx_(
    `="  INFORMACIÓN DE PRODUCCIÓN · MÁQUINAS ACTIVAS · "&IF(ISNUMBER(${fechaA1}),TEXT(${fechaA1},"dd/mm/yyyy"),"sin fecha")`));
  caja_(sh, FILA_TABLA_TITULO, COL_PUESTO, 1, ANCHO_PERSONAL).setValue('  PERSONAL DEL TURNO')
    .setBackground(COLOR.cabecera).setFontColor('#ffffff').setFontWeight('bold').setFontSize(11);

  // Botones: enviar a Planificación, descartar cambios y estado
  sh.getRange(ACCION.enviar[0], ACCION.enviar[1]).insertCheckboxes().setHorizontalAlignment('center')
    .setBackground(COLOR.enviar).setFontColor('#ffffff').setFontSize(12);
  caja_(sh, FILA_ACCIONES, ACCION.enviar[1] + 1, 1, 7).setValue('ENVIAR A PLANIFICACIÓN')
    .setBackground(COLOR.enviar).setFontColor('#ffffff').setFontWeight('bold').setFontSize(11);
  sh.getRange(ACCION.descartar[0], ACCION.descartar[1]).insertCheckboxes().setHorizontalAlignment('center')
    .setFontSize(12);
  caja_(sh, FILA_ACCIONES, ACCION.descartar[1] + 1, 1, 5).setValue('Descartar cambios')
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
  sh.getRange(FILA_TABLA_CAB, COL0, 1, ANCHO_TABLA).setBackground('#dfe3ea').setFontSize(9).setFontWeight('bold')
    .setFontColor(COLOR.texto).setWrap(true).setHorizontalAlignment('center');
  sh.getRange(FILA_TABLA_CAB, COL0, n + 1, ANCHO_TABLA)
    .setBorder(true, true, true, true, true, true, COLOR.lineaTabla, SB.SOLID);

  const maq = CAMPO.maquina;
  sh.getRange(FILA_TABLA_INI, maq.col, n, maq.ancho).setFontFamily('Roboto Mono').setFontWeight('bold');
  sh.getRange(FILA_TABLA_INI, CAMPO.ton.col, n, CAMPO.ton.ancho).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireNumberGreaterThanOrEqualTo(0).setAllowInvalid(false).setHelpText('Solo números (≥ 0).').build());

  // Cuadro de personal: encabezado (los puestos de cada turno los pone ajustarTurno_)
  caja_(sh, FILA_TABLA_CAB, COL_PUESTO, 1, 4).setValue('PUESTO');
  caja_(sh, FILA_TABLA_CAB, COL_NOMBRE, 1, 4).setValue('NOMBRE');
  sh.getRange(FILA_TABLA_CAB, COL_PUESTO, 1, ANCHO_PERSONAL).setBackground('#dfe3ea').setFontSize(9)
    .setFontWeight('bold').setHorizontalAlignment('center')
    .setBorder(true, true, true, true, true, true, COLOR.lineaTabla, SB.SOLID);

  // Estado en rojo si hay cambios sin enviar
  reglas.push(regla().whenFormulaSatisfied(`=${pendA1}`).setFontColor('#b3261e').setBold(true)
    .setRanges([estadoRg]).build());
  // Filas de la tabla en blanco si no tienen máquina
  const filaTabla = sh.getRange(FILA_TABLA_INI, COL0, n, ANCHO_TABLA);
  const maqA1 = '$' + letraCol_(maq.col) + FILA_TABLA_INI;
  reglas.push(regla().whenFormulaSatisfied(`=${maqA1}=""`).setBackground('#ffffff').setRanges([filaTabla]).build());
  sh.setConditionalFormatRules(reglas);

  caja_(sh, FILA_AYUDA, COL0, 1, ANCHO_TOTAL).setWrap(true).setVerticalAlignment('top').setFontSize(9)
    .setFontColor(COLOR.etiqueta)
    .setValue('Cómo usar: 1) Elige la fecha arriba. 2) Marca las máquinas que trabajarán en el turno (una por una, ' +
      'todas o por línea); se agregan a la tabla en orden de número. 3) Completa los campos amarillos de la tabla ' +
      'y el personal del turno. 4) Marca «Enviar a Planificación» para guardar todo. Antes de cambiar de fecha envía o ' +
      'descarta los cambios. El tonelaje se recuerda por máquina. En los turnos 2 y 3 un técnico de inyección puede ' +
      'cubrir la supervisión. Turno 3: la fecha es la del día en que inicia el turno (10:30 PM).');
}

function ajustarTurno_(sh, t) {
  const T = TURNOS[t];
  sh.getRange(1, U(1)).setValue(`MANAPLAS · PLANIFICACIÓN DE PRODUCCIÓN · ${T.hoja.toUpperCase()}${T.standard ? ' (STANDARD)' : ''}`);
  sh.getRange(FILA_VALOR, TARJETAS.horario[1]).setValue(T.horario);
  sh.getRange(FILA_VALOR, TARJETAS.descanso[1]).setValue(T.descanso);
  sh.getRange(CELDA.horas[0], CELDA.horas[1]).setValue(T.horas);
  sh.setTabColor(T.color);

  // Cuadro de personal con los puestos del turno
  const SB = SpreadsheetApp.BorderStyle;
  const nombres = nNombres_(t);
  let fila = FILA_TABLA_INI;
  const rangos = { produccion: [], tecnico: [] };
  puestos_(t).forEach(p => {
    caja_(sh, fila, COL_PUESTO, p.n, 4).setValue(p.titulo).setFontWeight('bold').setFontSize(9).setWrap(true)
      .setBackground(COLOR.fijo).setVerticalAlignment('middle');
    rangos[p.grupo].push(sh.getRange(fila, COL_NOMBRE, p.n, 1).getA1Notation());
    fila += p.n;
  });
  sh.getRange(FILA_TABLA_INI, COL_NOMBRE, nombres, 4).mergeAcross().setNumberFormat('@').setBackground(COLOR.entrada);
  sh.getRange(FILA_TABLA_CAB, COL_PUESTO, nombres + 1, ANCHO_PERSONAL)
    .setBorder(true, true, true, true, true, true, COLOR.lineaTabla, SB.SOLID);

  // Resumen: lo asignado frente a la meta de la propuesta (rojo si falta, verde si se cumple)
  const plano = sh.getRange(FILA_PLANO_INI, U(1), FILA_PLANO_FIN - FILA_PLANO_INI + 1, N_UNIDADES).getA1Notation();
  const ops = sh.getRange(FILA_TABLA_INI, CAMPO.operario1.col, N_FILAS_TABLA,
    CAMPO.operario3.col + CAMPO.operario3.ancho - CAMPO.operario1.col).getA1Notation();
  const contar = lista => lista.length ? lista.map(a => `COUNTA(${a})`).join('+') : '0';
  const cuenta = {
    maquinas: `COUNTIF(${plano},TRUE)`,
    operarios: `COUNTA(${ops})`,
    produccion: `COUNTA(${ops})+${contar(rangos.produccion)}`,
    tecnicos: contar(rangos.tecnico),
  };
  const regla = SpreadsheetApp.newConditionalFormatRule;
  const reglas = sh.getConditionalFormatRules();
  RESUMEN.forEach(x => {
    const celda = sh.getRange(FILA_RESUMEN, x.col + 3);
    const meta = T.metas[x.k];
    celda.setFormula(fx_(meta ? `=(${cuenta[x.k]})&" / ${meta}"` : `=(${cuenta[x.k]})&""`));
    if (!meta) return;
    const rg = sh.getRange(FILA_RESUMEN, x.col + 3, 1, 2);
    reglas.push(regla().whenFormulaSatisfied(fx_(`=(${cuenta[x.k]})<${meta}`)).setFontColor('#b3261e').setRanges([rg]).build());
    reglas.push(regla().whenFormulaSatisfied(fx_(`=(${cuenta[x.k]})>=${meta}`)).setFontColor('#1b7a3a').setRanges([rg]).build());
  });
  sh.setConditionalFormatRules(reglas);
}

/** Protección de advertencia: avisa si se intenta escribir sobre el plano o las etiquetas. */
function proteger_(sh, t) {
  sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
  const p = sh.protect().setDescription('Plano y tabla (solo advertencia)').setWarningOnly(true);
  p.setUnprotectedRanges([
    sh.getRange(CELDA.fechaSel[0], CELDA.fechaSel[1], 1, ANCHO_FECHA),
    sh.getRange(CELDA.todas[0], CELDA.todas[1]),
    sh.getRange(ACCION.enviar[0], ACCION.enviar[1]),
    sh.getRange(ACCION.descartar[0], ACCION.descartar[1]),
    sh.getRange(FILA_TABLA_INI, COL_ENTRADA_INI, N_FILAS_TABLA, COL_ENTRADA_FIN - COL_ENTRADA_INI + 1),
    sh.getRange(FILA_TABLA_INI, COL_NOMBRE, nNombres_(t), 4),
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
    out.push({ fila: i + 2, fecha: fecha, turno: turno, maquina: maq, valores: r });
  });
  return out;
}

/** Registros de una fecha y turno, por número de máquina. */
function registros_(datos, clave, t) {
  const rec = {};
  if (clave) datos.forEach(r => { if (r.fecha === clave && r.turno === t) rec[r.maquina] = r; });
  return rec;
}

/** Estado actual de las casillas del plano: { id: true/false }. */
function leerCasillas_(sh) {
  const plano = sh.getRange(FILA_PLANO_INI, COL0, FILA_PLANO_FIN - FILA_PLANO_INI + 1, ULTIMA_COL - COL0).getValues();
  const out = {};
  maquinas_().forEach(m => out[m.id] = plano[m.chk.fila - FILA_PLANO_INI][m.chk.col - COL0] === true);
  return out;
}

/** Lee la tabla de la hoja: { id: { campo: valor } } de las filas que tienen máquina. */
function leerTabla_(sh) {
  const out = {};
  sh.getRange(FILA_TABLA_INI, COL0, N_FILAS_TABLA, ANCHO_TABLA).getValues().forEach(fila => {
    const id = Number(fila[CAMPO.maquina.col - COL0]);
    if (!id) return;
    const v = {};
    ENTRADAS.forEach(x => v[x.k] = fila[x.col - COL0]);
    out[id] = v;
  });
  return out;
}

/**
 * Escribe la tabla con las máquinas activas una debajo de otra, en orden de número.
 * `valores(id)` da los campos de cada máquina. Muestra al menos las filas del cuadro de personal.
 */
function escribirTabla_(sh, t, activa, valores) {
  const ids = ORDEN_TABLA.filter(id => activa[id]);
  const filas = [];
  for (let i = 0; i < N_FILAS_TABLA; i++) {
    const fila = new Array(ANCHO_TABLA).fill('');
    if (i < ids.length) {
      const v = valores(ids[i]) || {};
      fila[CAMPO.maquina.col - COL0] = ids[i];
      ENTRADAS.forEach(x => fila[x.col - COL0] = v[x.k] === undefined ? '' : v[x.k]);
    }
    filas.push(fila);
  }
  sh.getRange(FILA_TABLA_INI, COL0, N_FILAS_TABLA, ANCHO_TABLA).setValues(filas);
  const visibles = Math.max(ids.length, nNombres_(t));
  sh.showRows(FILA_TABLA_INI, visibles);
  if (visibles < N_FILAS_TABLA) sh.hideRows(FILA_TABLA_INI + visibles, N_FILAS_TABLA - visibles);
}

/** Campos de una máquina desde su registro de Planificación, con el tonelaje guardado si falta. */
function valoresRegistro_(r, ton, id) {
  const v = {};
  ENTRADAS.forEach(x => v[x.k] = r ? r.valores[IDX[x.k]] : '');
  if (vacio_(v.ton) && ton[id] !== undefined) v.ton = ton[id];
  return v;
}

/** Casillas «todas» y «Línea N» marcadas solo si todas sus máquinas están activas. */
function actualizarGrupos_(sh, activa) {
  const ids = Object.keys(activa).map(Number);
  sh.getRange(CELDA.todas[0], CELDA.todas[1]).setValue(ids.every(id => activa[id]));
  GRUPOS.forEach((g, i) => sh.getRange(CELDA.grupos[i][0], CELDA.grupos[i][1])
    .setValue(ids.filter(id => id >= g.desde && id <= g.hasta).every(id => activa[id])));
}

/**
 * Tras marcar o desmarcar casillas: vuelve a escribir la tabla con las máquinas activas en orden.
 * Lo ya escrito en la tabla se conserva; una máquina que se agrega trae lo último enviado para
 * esta fecha y turno (o al menos su tonelaje).
 */
function aplicarSeleccion_(sh, t) {
  const activa = leerCasillas_(sh);
  const actual = leerTabla_(sh);
  const rec = registros_(leerPlan_(), claveFecha_(sh.getRange(CELDA.fechaSel[0], CELDA.fechaSel[1]).getValue()), t);
  const ton = tonelajes_();
  escribirTabla_(sh, t, activa, id => actual[id] || valoresRegistro_(rec[String(id)], ton, id));
  actualizarGrupos_(sh, activa);
}

function marcarPendiente_(sh) {
  sh.getRange(CELDA.pendiente[0], CELDA.pendiente[1]).setValue(true);
  sh.getRange(ACCION.estado[0], ACCION.estado[1])
    .setValue('⚠ Hay cambios sin enviar: marca «Enviar a Planificación» para guardarlos.');
}

/**
 * Carga en la hoja lo guardado en Planificación para la fecha y turno: casillas, tabla,
 * filas visibles y personal. Descarta los cambios sin enviar.
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

  escribirTabla_(sh, t, activa, id => valoresRegistro_(rec[String(id)], ton, id));
  sh.getRange(FILA_TABLA_INI, COL_NOMBRE, nNombres_(t), 1)
    .setValues(leerPersonal_(claveFecha_(fechaV), t).map(x => [x]));
  actualizarGrupos_(sh, activa);
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
  const tabla = leerTabla_(sh);
  const val = (id, k) => (tabla[id] && tabla[id][k] !== undefined) ? tabla[id][k] : '';

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

  // Personal del turno
  guardarPersonal_(fechaV, clave, t,
    sh.getRange(FILA_TABLA_INI, COL_NOMBRE, nNombres_(t), 1).getValues().map(r => r[0]));

  sh.getRange(CELDA.pendiente[0], CELDA.pendiente[1]).setValue(false);
  sh.getRange(CELDA.fechaCargada[0], CELDA.fechaCargada[1]).setValue(fechaV);
  const hora = Utilities.formatDate(new Date(), tz(), 'h:mm a');
  const total = nuevas.length + actualizadas;
  sh.getRange(ACCION.estado[0], ACCION.estado[1]).setValue(`✓ Enviado a Planificación a las ${hora}: ${total} máquina(s)` +
    (borrar.length ? `, ${borrar.length} quitada(s).` : '.'));
  ss.toast(`${nuevas.length} nueva(s), ${actualizadas} actualizada(s), ${borrar.length} quitada(s).`, 'Enviado a Planificación', 5);
}

/** Casillas de nombre del cuadro de personal de un turno, en orden: [{ puesto, n }]. */
function casillasPersonal_(t) {
  const out = [];
  puestos_(t).forEach(p => { for (let i = 1; i <= p.n; i++) out.push({ puesto: p.titulo, n: i }); });
  return out;
}

/** Filas de «Personal por turno» de una fecha y turno: [{ fila, puesto, n, nombre }]. */
function filasPersonal_(clave, t) {
  const sh = SpreadsheetApp.getActive().getSheetByName(HOJA_PERSONAL);
  if (!sh || !clave || sh.getLastRow() < 2) return [];
  const out = [];
  sh.getRange(2, 1, sh.getLastRow() - 1, COLS_PERSONAL.length).getValues().forEach((r, i) => {
    if (claveFecha_(r[0]) === clave && Number(r[1]) === t) out.push({ fila: i + 2, puesto: r[2], n: Number(r[3]), nombre: r[4] });
  });
  return out;
}

/** Nombres guardados para el cuadro de personal de una fecha y turno (vacíos si no hay). */
function leerPersonal_(clave, t) {
  const guardado = filasPersonal_(clave, t);
  return casillasPersonal_(t).map(c => {
    const g = guardado.find(x => x.puesto === c.puesto && x.n === c.n);
    return g ? g.nombre : '';
  });
}

/** Reemplaza el personal guardado de una fecha y turno por los nombres del cuadro (una fila por persona). */
function guardarPersonal_(fechaV, clave, t, nombres) {
  const sh = SpreadsheetApp.getActive().getSheetByName(HOJA_PERSONAL);
  if (!sh) return;
  filasPersonal_(clave, t).map(x => x.fila).sort((a, b) => b - a).forEach(f => sh.deleteRow(f));
  const filas = casillasPersonal_(t).map((c, i) => [fechaV, t, c.puesto, c.n, nombres[i]])
    .filter(f => !vacio_(f[4]));
  if (!filas.length) return;
  const ini = Math.max(sh.getLastRow(), 1) + 1;
  const fin = ini + filas.length - 1;
  if (fin > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), fin - sh.getMaxRows() + 200);
  sh.getRange(ini, 1, filas.length, COLS_PERSONAL.length).setValues(filas);
}

/** Recarga las hojas de turno desde Planificación (salvo las que tienen cambios sin enviar). */
function actualizarTodo_() {
  const datos = leerPlan_();
  [1, 2, 3, 4].forEach(t => {
    const sh = SpreadsheetApp.getActive().getSheetByName(TURNOS[t].hoja);
    if (!sh) return;
    if (sh.getRange(CELDA.pendiente[0], CELDA.pendiente[1]).getValue() !== true) cargar_(sh, t, datos);
  });
}
