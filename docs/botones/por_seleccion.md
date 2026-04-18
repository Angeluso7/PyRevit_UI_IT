# Botón Por Selección

## Estado
Iteración 4 con formato XLSX alineado al comportamiento original de `script_2.py` y `script_3.py`.

## Formato aplicado
- Se reservan las filas 1 a 5 antes de los encabezados de datos.
- Los encabezados reales parten en la fila 6.
- Los encabezados de la fila 6 usan dos colores: gris para columnas de solo lectura y azul para columnas modificables.
- Se escriben en A1, A2 y A3: nombre de planilla, fecha de generación y total de activos.
- Se agrega leyenda visual en I2:J3 para distinguir columnas de solo lectura y modificables.
- Los encabezados admiten reemplazos desde `script.json` y división por `/` para crear encabezados multinivel.
- Cuando hay textos repetidos consecutivos en niveles superiores del encabezado, se combinan horizontalmente.
- El CSV temporal se elimina después de generar el XLSX, igual que en el flujo original.

## Configuración usada
- `reemplazos_encabezados` y `reemplazos_de_nombres` se leen desde `script.json`.
- La ruta puede configurarse mediante `SCRIPT_JSON_PATH` en `lib/config/settings.py`.

## Observación
La lógica quedó integrada en un solo script CPython para mantener el flujo más simple, pero reproduce la intención funcional de los dos scripts originales compartidos.
