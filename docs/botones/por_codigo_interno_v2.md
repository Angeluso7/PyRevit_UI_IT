# Por Código Interno - iteración v2

Esta versión reincorpora para el primer renglón la lógica funcional antigua del botón para actualizar filtros por `CodIntBIM`.

## Criterio aplicado

- Se mantiene intacta la lógica del segundo renglón (`Asignados` / `No asignados`).
- El primer renglón vuelve a usar la estrategia funcional previa:
  - localizar `f_element_x` y `f_element_y`,
  - obtener el Id del parámetro `CodIntBIM` desde `ParameterBindings`,
  - actualizar la regla de igualdad para `f_element_x`,
  - actualizar la regla de desigualdad para `f_element_y`,
  - aplicar ambos filtros a la vista,
  - activar visibilidad solo del filtro objetivo.

## Cambio respecto a la iteración anterior

Se eliminó la dependencia de encontrar un elemento de muestra en la vista activa para construir la regla del parámetro en el primer renglón.
