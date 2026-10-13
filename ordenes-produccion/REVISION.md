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
