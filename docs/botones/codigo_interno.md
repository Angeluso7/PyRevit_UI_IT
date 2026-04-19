# Botón Código interno

## Estado
Iteración 6 con refactorización del botón `Código interno.pushbutton` según la misma estructura aplicada antes al botón `Por Selección.pushbutton`.

## Cambios principales
- `script.py` quedó reducido a punto de entrada y validación de vista 3D.
- La lectura de `CodIntBIM` desde host y vínculos se movió a `lib/services/codint_service.py`.
- La lógica de filtros y validación de categorías incompatibles se movió a `lib/revit/filter_utils.py`.
- La interfaz CPython Tkinter quedó encapsulada en `scripts_cpython/codint_selector.py`.
- Se preserva el flujo de selección por `CodIntBIM` y por estado `Asignados / No asignados`.

## Robustecimientos aplicados
- Evita fallos directos cuando un filtro contiene categorías donde `CodIntBIM` no aplica.
- Informa categorías incompatibles excluidas antes de reasignar reglas.
- Se centraliza el uso de `CPYTHON_EXE` y `DATA_DIR` desde configuración.
- Se mantiene el uso de JSON temporal en `data`, pero ya desacoplado del botón principal.

## Próxima mejora recomendada
- Extraer la espera bloqueante del selector a una integración reusable si otros botones usarán el mismo patrón CPython + JSON temporal.
