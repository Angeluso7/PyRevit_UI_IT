# Guía de Implementación - Sistema de Exportación BIM v2.0

## Resumen General

Se ha creado una **nueva solución** que reemplaza el flujo anterior (`ui_comparacion.py` + `leer_xlsm_codigos.py`). 

Esta nueva solución:
- ✅ **Extrae directamente desde Revit** usando la Revit API (No requiere Excel intermediario)
- ✅ **Accede a modelos linkeados** y extrae ElementIds, CodIntBIM, parámetros
- ✅ **Permite seleccionar carpeta de salida** (no carpeta fija)
- ✅ **Genera Excel con estructura completa** según especificaciones
- ✅ **Crea hoja de Índice** con listado de tablas
- ✅ **Crea hoja de Excepciones** para elementos sin CodIntBIM
- ✅ **Usa diccionario `codigos_planillas`** del `script.json` para mapear prefijos a nombres de tabla

---

## Estructura de Archivos

### Ubicación recomendada:

```
%APPDATA%\Roaming\MyPyRevitExtention\PyRevitIT.extension\
├── data/
│   ├── script.json                           (EXISTENTE)
│   ├── formatear_tablas_excel_v2.py         (NUEVO)
└── commands/
    └── extraer_modelos_bim.py               (NUEVO - PyRevit script)
```

---

## Archivos Creados

### 1. `extraer_modelos_bim.py` (PyRevit IronPython)

**Ubicación:** `%APPDATA%\Roaming\MyPyRevitExtention\PyRevitIT.extension\commands\`

**Función:** Script principal que se ejecuta desde pyRevit

**Flujo:**
1. Verifica rutas necesarias (DATA_DIR_EXT, script.json, formatear_script)
2. Carga configuración desde `script.json` (especialmente `codigos_planillas`)
3. **Abre diálogo para seleccionar carpeta de salida**
4. Extrae ElementIds, CodIntBIM y parámetros de todos los modelos linkeados
5. Procesa datos y agrupa por tabla (según primeros 4 caracteres de CodIntBIM)
6. Genera archivo JSON intermedio con datos procesados
7. Ejecuta `formatear_tablas_excel_v2.py` con CPython
8. Muestra mensaje de éxito con ruta del archivo generado

**Dependencias:**
- Revit API (Autodesk.Revit.DB)
- PyRevit forms
- subprocess
- json

---

### 2. `formatear_tablas_excel_v2.py` (CPython)

**Ubicación:** `%APPDATA%\Roaming\MyPyRevitExtention\PyRevitIT.extension\data\`

**Función:** Genera archivo Excel con estructura completa

**Entrada:**
- Archivo JSON intermedio con datos procesados (generado por script PyRevit)
- Ruta donde guardar el Excel final

**Salida:**
- Archivo `.xlsx` con estructura:
  - **Hoja 0: "Índice"** - Listado de tablas
  - **Hojas 1..N: Tablas de datos** - Una por cada tabla de planificación
  - **Hoja Final: "Excepciones"** - Elementos sin CodIntBIM (si existen)

**Características:**
- Encabezados con fondo azul oscuro (FF366092) y texto blanco
- Auto-ajuste de ancho de columnas
- Bordes en todas las celdas
- Primera fila congelada (freeze_panes)
- Alineación de texto con wrap

---

## Estructura del Excel Generado

### Hoja "Índice" (Primera)

```
Fila 3, Columna B: "Listado de Tablas"
Fila 4, Columnas B-C: Headers "Nombre Tabla" | "Código"
Fila 5+: Valores de codigos_planillas | Claves de codigos_planillas
```

### Hojas de Tablas (Una por tabla de planificación)

Orden: Ordenadas por prefijo CM01, CM02, CM03, etc.

**Columnas:**
1. **Parámetros del elemento** (todos aquellos que no sean campos adicionales)
2. **ElementId** - ID del elemento en Revit
3. **CodIntBIM** - Código identificador
4. **Categoría** - Categoría de Revit del elemento
5. **Familia** - Familia (si es aplicable)
6. **Tipo** - Tipo de familia (si es aplicable)
7. **Nombre_RVT** - Nombre del archivo RVT donde está el elemento
8. **Situación CodIntBIM** - "Elemento no Anidado" (configurable)
9. **Elementos** - "único" (un elemento) o "varios" (múltiples elementos)

### Hoja "Excepciones" (Si existen excepciones)

**Elementos sin CodIntBIM**

Mismas columnas que tablas, pero:
- Columna **CodIntBIM** contiene:
  - "No Asignado" - Si el parámetro existe pero está vacío
  - "No existe" - Si el parámetro no existe en el elemento

---

## Cómo Usar

### Desde PyRevit:

1. **Copiar archivos a ubicación correcta:**
   - `extraer_modelos_bim.py` → carpeta `commands`
   - `formatear_tablas_excel_v2.py` → carpeta `data`

2. **En Revit:**
   - Abrir documento maestro con modelos linkeados
   - En pyRevit, buscar y ejecutar: **Extraer Modelos BIM**
   - Se abrirá diálogo para seleccionar **carpeta de salida**
   - El script extrae datos, procesa y genera Excel
   - Mensaje de éxito mostrará ruta del archivo generado

3. **Verificar archivo:**
   - Abrir Excel en carpeta seleccionada
   - Revisar "Índice" con listado de tablas
   - Revisar hojas de datos con estructura de elementos
   - Si hay excepciones, revisar hoja "Excepciones"

---

## Configuración Necesaria en script.json

El archivo `script.json` debe contener:

```json
{
  "codigos_planillas": {
    "CM01": "Líneas de Distribución",
    "CM02": "Transformadores",
    "CM03": "Subestaciones",
    "CM04": "Equipos de Protección",
    ...
  }
}
```

**Donde:**
- **Clave:** Prefijo de 4 caracteres (primeros 4 caracteres del CodIntBIM)
- **Valor:** Nombre de la tabla/hoja en Excel

El script automáticamente:
1. Busca elementos cuyos CodIntBIM comienzan con estos prefijos
2. Agrupa elementos por tabla
3. Crea hojas Excel con nombre de la tabla
4. Ordena hojas en orden de prefijo (CM01, CM02, CM03...)

---

## Flujo Técnico Detallado

### En PyRevit (extraer_modelos_bim.py):

```
1. verificar_rutas()
   ├─ DATA_DIR_EXT existe?
   ├─ script.json existe?
   └─ formatear_tablas_excel_v2.py existe?

2. cargar_config()
   └─ Lee script.json → codigos_planillas

3. seleccionar_carpeta_salida()
   └─ Diálogo de selección de carpeta

4. extraer_modelos_linkeados()
   ├─ Para cada LinkInstance:
   │  ├─ Obtener linked document
   │  └─ Para cada elemento:
   │     ├─ Extraer ElementId
   │     ├─ Extraer CodIntBIM (parámetro)
   │     ├─ Extraer Categoría, Familia, Tipo
   │     ├─ Extraer Nombre_RVT
   │     └─ Extraer todos los parámetros
   └─ Retorna lista de dicts con datos

5. procesar_datos(datos_raw, codigos_planillas)
   ├─ Para cada elemento:
   │  ├─ Extraer prefijo (4 primeros chars de CodIntBIM)
   │  ├─ Mapear prefijo a nombre_tabla
   │  └─ Agrupar en tabla correspondiente
   └─ Retorna elementos_por_tabla + excepciones

6. generar_json_para_formatear()
   └─ Escribe JSON intermedio

7. ejecutar_formateo()
   └─ subprocess.check_call([python3, formatear_script, ...])

8. Mensaje de éxito
```

### En CPython (formatear_tablas_excel_v2.py):

```
1. cargar_datos_json()
   └─ Lee JSON intermedio

2. Crear Workbook

3. generar_hoja_indice()
   ├─ Crear hoja "Índice"
   ├─ B3: "Listado de Tablas"
   ├─ B4-C4: Headers
   └─ B5+: Valores y claves de codigos_planillas

4. Para cada tabla en elementos_por_tabla:
   ├─ generar_hoja_tabla()
   │  ├─ Crear hoja con nombre de tabla
   │  ├─ Escribir headers (parámetros + campos adicionales)
   │  ├─ Escribir datos de elementos
   │  └─ Aplicar formato
   └─ Siguiente tabla

5. Si hay excepciones:
   ├─ generar_hoja_excepciones()
   │  ├─ Crear hoja "Excepciones"
   │  ├─ Escribir datos excepcionales
   │  └─ Aplicar formato
   └─ Fin excepciones

6. wb.save(ruta_xlsx_salida)
```

---

## Parámetros Personalizables

En `extraer_modelos_bim.py`, línea ~29-31:

```python
DATA_DIR_EXT = r"..."  # Ruta donde está script.json
FORMATEAR_SCRIPT = r"..."  # Ruta del script de formateo
PYTHON3_EXE = r"..."  # Ruta a python.exe de CPython
```

En `formatear_tablas_excel_v2.py`, línea ~12-14:

```python
HEADER_COLOR = 'FF366092'  # Cambiar color de encabezados
HEADER_FONT_COLOR = 'FFFFFFFF'  # Cambiar color de texto
```

---

## Gestión de Errores

Ambos scripts incluyen validaciones:
- Verificación de archivos requeridos
- Manejo de excepciones en extracciones
- Mensajes de error descriptivos vía `forms.alert()` (PyRevit)

---

## Ejemplo de Datos de Salida

### JSON intermedio (_temp_datos.json):

```json
{
  "elementos_por_tabla": {
    "Líneas de Distribución": [
      {
        "ElementId": 12345,
        "CodIntBIM": "CM01-0001",
        "Voltaje": "13.8 kV",
        "Longitud": "250 m",
        "Categoría": "Generic Models",
        "Familia": "Línea Distribución",
        "Tipo": "Tipo A",
        "Nombre_RVT": "Modelo_Eléctrico.rvt"
      },
      ...
    ],
    "Transformadores": [
      {...},
      ...
    ]
  },
  "excepciones": [
    {
      "elemento": {...},
      "situacion": "No existe"
    }
  ],
  "listado_tablas": {
    "valores": ["Líneas de Distribución", "Transformadores", ...],
    "claves": ["CM01", "CM02", ...]
  }
}
```

---

## Notas Importantes

1. **Parámetro "CodIntBIM" es obligatorio** en los elementos para ser procesados
2. **Los primeros 4 caracteres** del CodIntBIM determinan la tabla
3. **El orden de las hojas** es: Índice → Tablas (CM01, CM02, ...) → Excepciones
4. **Elementos sin CodIntBIM** van automáticamente a "Excepciones"
5. **El archivo JSON intermedio** se elimina automáticamente después de generar Excel
6. **Ancho de columnas** se ajusta automáticamente al contenido (máx 60 caracteres)

---

## Siguientes Pasos

1. ✅ Copiar archivos a carpetas correctas
2. ✅ Verificar `script.json` tiene `codigos_planillas`
3. ✅ Verificar rutas en `extraer_modelos_bim.py`
4. ✅ Verificar ruta a Python 3.13 en ambos scripts
5. ✅ Ejecutar desde PyRevit
6. ✅ Validar estructura del Excel generado

---

**Versión:** 2.0  
**Fecha:** Enero 2026  
**Autor:** BIM Automation / Desarrollo Custom
