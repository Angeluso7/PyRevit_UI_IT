# Por Código Interno

Esta iteración del botón elimina la discriminación automática de categorías en los filtros `f_element_x` y `f_element_y`.

## Cambio principal

El script ya no modifica las categorías del filtro ni muestra alertas de categorías incompatibles con `CodIntBIM`.

## Nuevo comportamiento

- Renglón 1: toma el `CodIntBIM` seleccionado y solo actualiza las reglas de `f_element_x` y `f_element_y`.
- `f_element_x`: se actualiza con igualdad al código seleccionado.
- `f_element_y`: se actualiza con desigualdad al código seleccionado.
- Se respetan las categorías configuradas manualmente en el modelo.
- Renglón 2: mantiene la activación por filtros `c_cod_int` y `s_cod_int`.

## Condición importante

Para que funcione correctamente, la vista activa debe contener al menos un elemento con el parámetro `CodIntBIM`, ya que el Id del parámetro se toma desde una muestra de la vista activa.
