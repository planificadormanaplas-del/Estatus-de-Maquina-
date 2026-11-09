Necesito que construyas desde cero un archivo de Excel para planificar los turnos de producción de una planta de inyección de plástico (Manaplas). Hoy lo tengo en Google Sheets con Apps Script y lo voy a migrar a Excel. Quiero que las hojas de turno se vean y funcionen igual que ahora: lo describo abajo con todo detalle.

Te voy a pasar **una imagen con la distribución de las máquinas en la planta**. Usa esa imagen como referencia del plano: misma posición relativa de cada máquina, mismos grupos y mismo orden. Si la imagen y lo que describo abajo no coinciden en algo, manda la imagen y pregúntame antes de decidir.

## Cómo quiero que trabajes
1. Antes de construir nada, hazme estas preguntas y espera mis respuestas:
   - Qué versión de Excel usamos (Excel 365 de escritorio, Excel 2016/2019, Excel en la web) y si se puede usar un archivo con macros (.xlsm). De esto depende cómo hacer las casillas y el botón de enviar (ver «Decisiones técnicas»).
   - Cualquier cosa de la imagen que no se lea bien (números de máquina, grupos).
2. Antes de entregar el archivo final, muéstrame una **vista previa en imagen** de la hoja de un turno (plano + tabla) con datos de ejemplo, para aprobar el diseño. Hago ajustes visuales seguido: prefiero ver una vista previa por cada cambio grande.
3. Prueba el archivo antes de entregarlo (que abra sin errores, que las fórmulas calculen, que las macros compilen si las hay) y dime qué probaste y qué no pudiste probar.
4. Háblame en español, con pasos cortos. No soy programador: si tengo que hacer algo en Excel (habilitar macros, pegar código, etc.), explícame paso a paso.

## Estructura del archivo
- **Turno 1, Turno 2, Turno 3, Turno 4**: una hoja por turno, con el plano de máquinas, la tabla de máquinas activas y el cuadro de personal.
- **Planificación**: base de datos de máquinas, una fila por fecha + turno + máquina.
- **Personal por turno**: base de datos del personal, una fila por persona (Fecha, Turno, Puesto, N°, Nombre).
- **Listas** (oculta): lista de máquinas y su tonelaje (Ton), que se recuerda por máquina.

### Turnos
| Turno | Horario | Descanso (cada persona elige 1 h dentro de ese intervalo) | Horas efectivas |
|---|---|---|---|
| 1 | 6:00 AM – 2:30 PM | 1 h entre 9:30 AM y 12:00 PM | 7,5 h |
| 2 | 2:30 PM – 10:30 PM | 1 h entre 5:00 PM y 7:00 PM | 8 h |
| 3 | 10:30 PM – 6:00 AM | 1 h entre 2:00 AM y 3:00 AM | 6,5 h |
| 4 (Standard) | 7:00 AM – 4:00 PM | 1 h de 12:00 PM a 1:00 PM | 8 h |

En el Turno 3, la fecha es la del día en que empieza el turno (10:30 PM).

### Máquinas (47 en total)
- **Línea 1** (101–143): 101, 102, 103, 107, 108, 109, 110, 111, 113, 116, 117, 119, 121, 123, 125, 127, 129, 131, 132, 133, 134, 135, 136, 139, 140, 143
- **Línea 2** (201–218): 201, 202, 203, 204, 206, 207, 209, 210, 211, 212, 213, 214, 215, 216, 217, 218
- **Línea 3 · Soplado** (301–306): 301, 302, 304, 305, 306

## Hoja de turno (el diseño actual, que quiero mantener)
Todo el encabezado y el plano deben verse **completos en una pantalla Full HD sin hacer scroll** (aprox. 1.700 × 650 px). La tabla va debajo.

### Encabezado, de arriba abajo
1. **Barra de título** oscura de lado a lado: «MANAPLAS · PLANIFICACIÓN DE PRODUCCIÓN · TURNO N» (en el Turno 4: «TURNO 4 (STANDARD)»).
2. **Fila de tarjetas**, cada una en un recuadro con la etiqueta pequeña en gris arriba y el valor en negrita abajo:
   - **FECHA**: celda editable en amarillo, con validación de fecha (dd/mm/aaaa).
   - **HORARIO**, **DESCANSO** y **HORAS EFECTIVAS** del turno.
   - **MÁQ. ACTIVAS**: «n / 47», donde n son las máquinas marcadas en el plano.
   - A la derecha, un recuadro de **selección rápida**: «Marcar / desmarcar todas» y, debajo, «Línea 1 (101–143)», «Línea 2 (201–218)» y «Línea 3 (301–306)». Cada una marca o desmarca todas las máquinas de su grupo.
3. **Franja «PROPUESTA»**: compara lo asignado con las metas del turno y muestra el número en rojo si falta y en verde si se cumple. Al final de la franja va la leyenda de colores: ■ Activa (trabaja) en verde y ■ Sin planificar en gris.
   - Turnos 1, 2 y 3: MÁQUINAS x / 12 · OPERARIOS x / 16 · PRODUCCIÓN x / 21 · TÉCNICOS x / 6.
   - Turno 4: solo TÉCNICOS x / 5. Máquinas, operarios y producción se muestran sin meta.
   - OPERARIOS = nombres escritos en las columnas Operario 1–3 de la tabla. PRODUCCIÓN = operarios + Supervisor de turno + Asistentes de producción + Puente y distribución. TÉCNICOS = suma de los puestos técnicos del cuadro de personal.

### Plano de máquinas
- Distribución **según la imagen que te paso**, con los títulos «LÍNEA 2 · MÁQUINAS 201–218», «LÍNEA 1 · MÁQUINAS 101–143» y «LÍNEA 3 · SOPLADO» (Soplado dentro de un recuadro punteado).
- **Todas las máquinas del mismo tamaño** (hoy cada una mide unos 111 × 72 px): número arriba en letra grande (monoespaciada, negrita) y casilla de selección abajo, dentro del mismo bloque con borde grueso.
- Máquinas de un mismo grupo **pegadas entre sí**. Entre filas de máquinas, **el mismo espacio pequeño en todo el plano** (simetría). Sin huecos grandes ni grupos apartados.
- **Seleccionable**: al marcar la casilla, el bloque entero se pone **verde al instante** (formato condicional, sin esperar a ninguna macro). Sin marcar se ve gris claro. Se pueden marcar varias a la vez.

### Botones (encima de la tabla)
- **«ENVIAR A PLANIFICACIÓN»** (botón azul): guarda la fecha y el turno en Planificación. Las máquinas marcadas se crean o se actualizan, y las desmarcadas se borran de esa fecha y turno. El personal se guarda en Personal por turno. El tonelaje escrito se recuerda por máquina en Listas.
- **«Descartar cambios»**: vuelve a cargar lo último enviado para esa fecha y turno.
- **Texto de estado** al lado: «⚠ Hay cambios sin enviar…» en rojo, o «✓ Enviado a las h:mm: n máquinas».
- **Cambio de fecha**: carga lo guardado de esa fecha y turno (máquinas marcadas, tabla y personal). Si hay cambios sin enviar, no deja cambiar de fecha y avisa.

### Tabla «INFORMACIÓN DE PRODUCCIÓN · MÁQUINAS ACTIVAS · dd/mm/aaaa»
- **Una fila por máquina marcada**, una debajo de otra, **siempre en orden de número de máquina** y sin filas vacías en medio. Al marcar una máquina nueva, lo ya escrito en las otras filas **no se pierde**.
- Columnas: **MÁQ.** (fija) · **TON** (número; se completa solo con el tonelaje guardado de la máquina) · **OP** · **PRODUCTO** · **REFERENCIA** · **COLOR** · **MATERIAL** · **OPERARIO 1** · **OPERARIO 2** · **OPERARIO 3**. Los campos editables van en amarillo claro.

### Cuadro «PERSONAL DEL TURNO» (a la derecha de la tabla)
Columnas PUESTO | NOMBRE, con una celda de nombre por persona. Si el puesto tiene varias personas, la celda del puesto va combinada a lo alto.
- **Turnos 1, 2 y 3**: Supervisor de turno (1) · Asistente de producción (2) · Puente y distribución, con montacargas (2) · Técnico de inyección (2) · Técnico eléctrico (1) · Técnico hidráulico (1) · Asistente de inyección (2).
  - Para PRODUCCIÓN cuentan Supervisor, Asistentes de producción y Puente.
  - Para TÉCNICOS cuentan Técnico de inyección, eléctrico, hidráulico y Asistente de inyección.
- **Turno 4**: Técnico de molde (2) · Molino (1) · Pintura / Mezcla (2). Todos cuentan como técnicos.
- Nota en la ayuda de la hoja: en los turnos 2 y 3, un técnico de inyección puede cubrir la supervisión.

### Bases de datos
- **Planificación**: Fecha · Turno · Máquina · Ton · OP · Producto · Referencia · Color · Material · Operario 1 · Operario 2 · Operario 3.
- **Personal por turno**: Fecha · Turno · Puesto · N° · Nombre.
- Encabezados con fondo oscuro, fila 1 inmovilizada y validaciones: fecha válida, turno 1–4, máquina de la lista, Ton numérico.

### Estilo general
- Fuente Roboto o similar. Colores: azul oscuro para títulos (#1f2a37), verde para activa (#b9e4c1), gris claro para sin planificar (#eef1f5), amarillo claro para campos editables (#fff8dc) y azul para el botón de enviar (#1f6fd1).
- Sin líneas de cuadrícula en las hojas de turno.
- Cada hoja de turno con un color de pestaña distinto.

## Decisiones técnicas (decídelas conmigo según mi versión de Excel)
- **Casillas de las máquinas.**
  - En Excel 365 actual se pueden usar las **casillas nativas en celda** (Insertar → Casilla), que guardan VERDADERO/FALSO y funcionan con formato condicional.
  - En versiones anteriores, usa una celda con lista desplegable («✔» o vacío), o una macro que alterne la marca al hacer clic.
  - Proponme la opción que funcione en mi versión, y que el color verde aparezca sin retraso.
- **Lo que necesita lógica** (tabla ordenada que conserva lo escrito, enviar, descartar, bloquear el cambio de fecha con pendientes, marcar por línea):
  - Con .xlsm, hazlo con VBA en módulos bien comentados en español.
  - Si no se pueden usar macros, dime qué partes se pueden hacer solo con fórmulas (por ejemplo FILTRAR/ORDENAR para listar las máquinas activas) y qué se perdería, antes de decidir.
- **Configuración regional:** está en español (coma decimal). Las fórmulas que escribas por código deben funcionar en esa configuración.
- **Protección:** no protejas las hojas. Lo quiero editable libremente.
- Si generas el archivo con Python (openpyxl o similar), dame también el script, para poder regenerarlo.

## Entrega
- El archivo de Excel (.xlsx o .xlsm, según lo que decidamos).
- Instrucciones cortas para abrirlo y usarlo: habilitar macros si aplica, cómo marcar máquinas, cómo enviar.
- Si hay VBA, el código también por separado, por si hay que volver a pegarlo.
