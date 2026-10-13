# Revisión de "Orden De Produccion" (Google Sheets) — v22

## Ya aplicado directamente en la hoja (13/10/2026)

- **Maestro, filas 2–241**: fórmulas puestas en C (TON), F, H, K, L, M, N, O, P, S y T, usando columnas completas. Las filas que no tenían fórmulas (fila 7, filas 69–73 y las OP 10–30) ya las tienen.
- **Hojas 01–14, H5**: código de OP cambiado a `=TEXT(H1;"ddmmyyyy")&"-"&TEXT(H4;"00")` (ej: `08102026-01`).
- **Maestro, U3**: se borró el espacio suelto.
- Nota: en las OP 9–30 la columna C (TON) está combinada por orden; la fórmula vive en la primera fila de cada orden y está bien así.

## Errores encontrados en la hoja

| # | Dónde | Problema | Arreglo |
|---|-------|----------|---------|
| 1 ✅ | Maestro, fila 7 (OP 1, color 6) | No tiene ninguna fórmula en C–P (TON, Producto, Sacos, Peso…). | Botón **🛠️ Reparar fórmulas del Maestro** |
| 2 ✅ | Maestro, filas 69–73 (OP 9, colores 4–8) | Igual: faltan las fórmulas. | Igual |
| 3 ✅ | Maestro, OP 10 a OP 30 | Solo la **primera fila** de cada orden tiene fórmulas. Desde el 2º color, Sacos, Sacos (kg) y Cantidad Final salen en blanco y el Historial los guarda vacíos. | Igual, y ahora el código pone las fórmulas solo al crear/agregar/organizar |
| 4 ✅ | Maestro, columnas M, N, O | Los rangos de búsqueda se "corren" (`$B$2:$M1029`, `$N1037`, `$N1038`, `$N1044`…). Hoy alcanza, pero los componentes que se agreguen al final de la lista no se encontrarían. | Las fórmulas reparadas usan columnas completas (`'Componentes '!$B:$N`) |
| 5 ✅ | Hojas 01–14, celda H5 | `=DAY(H1)&MONTH(H1)&YEAR(H1)&"-"&…` produce el mismo código para el 1/11/2026 y el 11/1/2026 (`1112026`). | **Aplicado**: `=TEXT(H1;"ddmmyyyy")&"-"&TEXT(H4;"00")` |
| 6 | Hojas 01–14, H11:H18 (Sacos) | Calcula `Cantidad × Peso / 25000` sin redondear, así que no coincide con los Sacos del Maestro (que redondean y respetan "producto importante"). | **A decidir**: si deben ser iguales, usar `=IF(Maestro!K2="";"";Maestro!K2)` |
| 7 | Maestro, columna K | Con `ROUND`, una orden pequeña de un producto "no importante" puede dar **0 sacos** (y Cantidad Final 0). | **A decidir**: `MAX(1; ROUND(...))` |
| 8 | Maestro, Q1 | El encabezado dice "⟵ Volver al Índice", pero el código guarda ahí el **Estado** y lo copia al Historial. | Revisar que Q sea realmente la columna de Estado |
| 9 ✅ | Maestro, U3 | Hay un espacio `" "` suelto. | Sin importancia, se puede borrar |

## Errores corregidos en el código

- **Organizar / Borrar sin Historial**: al recorrer órdenes, el "Tipo de Material" (G21) y las "Observaciones" (E34) se quedaban en la hoja vieja. Ahora se mueven con su orden.
- **Cantidad de colores inválida**: antes cancelaba toda la orden (o ponía 1 sin avisar en "varias órdenes"). Ahora vuelve a preguntar.
- **Gestionar → Agregar** sobre una OP vacía agregaba una fila vacía. Ahora avisa.
- **Guardar copia en Drive**: aceptaba días como "45" o "19abc"; ahora valida que el día exista en ese mes. Una orden adelantada de enero hecha en diciembre usa el año siguiente. Muestra el enlace de la copia.
- **onEdit**: si pegabas varios SKU, solo procesaba la primera fila. Al borrar el SKU, el Componente y su lista quedaban pegados.
- **Lista desplegable vieja** del Componente quedaba en las filas limpiadas.
- **Máquina**: avisa si su estatus en "Maq y Tons" no es "Operativa".
- **Velocidad**: el Maestro se lee de una sola vez en vez de celda por celda (Organizar, Vaciar y Buscar bloque libre hacían miles de lecturas).

## Cómo instalarlo

1. En la hoja: **Extensiones → Apps Script**.
2. Reemplaza todo el código por el de `Codigo.gs` y guarda.
3. Recarga la hoja y ejecuta una vez **📋 Órdenes de Producción → 🛠️ Reparar fórmulas del Maestro**.

## v23 — Maestro dinámico (en la copia de prueba)

Archivo de prueba: **Orden De Produccion - PRUEBA Maestro dinámico** (copia del original en Drive).
Código: `CodigoDinamico.gs` (el original sigue usando `Codigo.gs` v22).

Qué cambió en la copia:
- **Maestro**: se eliminaron las 967 filas preparadas/ocultas. Quedan los encabezados y filas en blanco visibles. Fila 1 congelada.
- **Columnas calculadas** (C, F, H, K–P, S, T): una sola fórmula `ARRAYFORMULA` en el encabezado que se aplica a todas las filas. No escribir debajo de esas columnas.
- **Hojas 01–14**: B6, D6, H6, B7, A11:C18, A20:A27 y H22:H26 ahora buscan en el Maestro las filas con `OP = H4` (ej: `INDEX(FILTER(Maestro!$D:$D;Maestro!$A:$A=$H$4);1)`), en vez de apuntar a filas fijas.
- **Columna OP**: se colorea sola (impar azul / par gris) con formato condicional.
- Probado con 2 órdenes de ejemplo: el Maestro calculó TON, Producto, Sacos, Peso, etc., y las hojas 01 y 02 mostraron cada una su orden. Los datos de prueba se borraron.

Qué hace el código v23:
- Generar orden → inserta sus filas (ordenadas por número de OP).
- Gestionar → Agregar color → inserta una fila dentro de la orden (máx. 8).
- Gestionar → Borrar / Borrar sin Historial / Vaciar → elimina las filas.
- Organizar → ordena por OP y renumera 1, 2, 3… moviendo Tipo de Material y Observaciones.
- Ya no hay límite de 30 órdenes; solo las OP 1–14 tienen hoja de impresión (avisa si eliges una mayor).
- Se quitaron "Colorear OP" y "Reparar fórmulas" (ya no hacen falta).

### Ajustes a la copia (después de la prueba)
- Se quitó el fondo azul de la columna OP y se devolvió el formato original a las filas del Maestro: verde claro en datos, gris en Componentes / Und-Hora / Und-Día, crema en Cantidad Final, bordes y Arial negrita.
- Columna Q: se quitó "⟵ Volver al Índice" (no se usaba) y ahora es **UND/DÍA** = Und/Hora × 10 h (turno de 6:00 a 16:00), calculada sola desde el encabezado.
- `CodigoDinamico.gs` ya no escribe en la columna Q.
- Filas vacías del Maestro en blanco (sin formato). `CodigoDinamico.gs` pone el formato verde/gris con bordes solo en las filas de cada orden al crearla o al agregar un color, y deja en blanco las filas que quedan libres al borrar, organizar o vaciar.
- El color de las filas ahora es un formato condicional (`=$A2<>""`): verde/gris/crema solo si la fila tiene OP. Si se borran los datos a mano (Suprimir), la fila vuelve a verse en blanco.
