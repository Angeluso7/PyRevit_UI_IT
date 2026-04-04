# Diagrama de Arquitectura - Sistema de Exportación BIM v2.0

## Flujo de Ejecución

```
┌─────────────────────────────────────────────────────────────────┐
│                     USUARIO EN REVIT                            │
│            Ejecuta botón "Extraer Modelos BIM"                  │
└────────────────────────┬────────────────────────────────────────┘
                         │
                         ▼
┌─────────────────────────────────────────────────────────────────┐
│          PyRevit Script (IronPython)                            │
│       extraer_modelos_bim.py                                   │
│                                                                 │
│  1. ✓ Verifica rutas necesarias                               │
│  2. ✓ Carga script.json → codigos_planillas                  │
│  3. ✓ Diálogo: Seleccionar carpeta de salida                 │
│  4. ✓ Accede a modelos linkeados                             │
│  5. ✓ Extrae ElementIds, CodIntBIM, parámetros               │
│  6. ✓ Procesa y agrupa por tabla (prefijo CM01, CM02, ...)   │
│  7. ✓ Genera JSON intermedio (_temp_datos.json)              │
└────────────────────────┬────────────────────────────────────────┘
                         │
                         │ subprocess.check_call(
                         │   [python3.exe, formatear_script, ...]
                         │ )
                         ▼
┌─────────────────────────────────────────────────────────────────┐
│        CPython Script                                           │
│    formatear_tablas_excel_v2.py                               │
│                                                                 │
│  1. ✓ Lee JSON intermedio                                     │
│  2. ✓ Crea Workbook Excel                                    │
│  3. ✓ Genera hoja "Índice" (B3-B5-C5)                        │
│  4. ✓ Genera hojas de tablas (CM01, CM02, ...)               │
│  5. ✓ Genera hoja "Excepciones" si hay                       │
│  6. ✓ Aplica formatos: encabezados, bordes, ajustes           │
│  7. ✓ Guarda Excel final                                      │
└────────────────────────┬────────────────────────────────────────┘
                         │
                         ▼
┌─────────────────────────────────────────────────────────────────┐
│              Archivo Excel Generado                             │
│         (carpeta_seleccionada\Exportacion_BIM_*.xlsx)          │
│                                                                 │
│  Hojas:                                                         │
│  ├─ [0] Índice (Listado de Tablas)                            │
│  ├─ [1] Líneas de Distribución (CM01)                         │
│  ├─ [2] Transformadores (CM02)                                │
│  ├─ [3] Subestaciones (CM03)                                  │
│  ├─ ...                                                         │
│  └─ [n] Excepciones (si existen)                              │
│                                                                 │
│  Limpieza: Se elimina _temp_datos.json                        │
└─────────────────────────────────────────────────────────────────┘
                         │
                         ▼
┌─────────────────────────────────────────────────────────────────┐
│              Mensaje de Éxito a Usuario                         │
│     "Archivo generado exitosamente en: [ruta]"                │
└─────────────────────────────────────────────────────────────────┘
```

---

## Estructura de Datos - JSON Intermedio

```json
{
  "elementos_por_tabla": {
    "Líneas de Distribución": [
      {
        "ElementId": 12345,
        "CodIntBIM": "CM01-0001",
        "Voltaje": "13.8 kV",
        "Longitud": "250 m",
        "Capacidad": "1000 A",
        "Categoría": "Generic Models",
        "Familia": "Línea Distribución",
        "Tipo": "Tipo A",
        "Nombre_RVT": "Modelo_Eléctrico.rvt"
      },
      {
        "ElementId": 12346,
        "CodIntBIM": "CM01-0002",
        ...
      }
    ],
    "Transformadores": [
      {
        "ElementId": 23456,
        "CodIntBIM": "CM02-0001",
        "Potencia": "500 kVA",
        "Voltaje_Primario": "13.8 kV",
        "Voltaje_Secundario": "480 V",
        "Categoría": "Electrical Equipment",
        "Familia": "Transformador",
        "Tipo": "Tipo T1",
        "Nombre_RVT": "Modelo_Eléctrico.rvt"
      }
    ]
  },
  "excepciones": [
    {
      "elemento": {
        "ElementId": 34567,
        "CodIntBIM": "",
        "Categoría": "Lighting Devices",
        ...
      },
      "situacion": "No Asignado"
    },
    {
      "elemento": {
        "ElementId": 34568,
        "CodIntBIM": "",
        "Categoría": "Electrical Equipment",
        ...
      },
      "situacion": "No existe"
    }
  ],
  "listado_tablas": {
    "valores": [
      "Líneas de Distribución",
      "Transformadores",
      "Subestaciones"
    ],
    "claves": [
      "CM01",
      "CM02",
      "CM03"
    ]
  }
}
```

---

## Estructura de Excel Generado

```
┌──────────────────────────────────────────────────────────────┐
│ Excel: Exportacion_BIM_[nombre_proyecto].xlsx               │
├──────────────────────────────────────────────────────────────┤
│ [Índice] | [Líneas de...] | [Transf...] | [...] | [Excepciones]
├──────────────────────────────────────────────────────────────┤
│
│ ═════════════════════════════════════════════════════════════
│ HOJA: Índice (Primera)
│ ═════════════════════════════════════════════════════════════
│
│   B3: Listado de Tablas
│
│   B4 │ C4
│  ─────────────────────────────────
│  Nombre Tabla │ Código
│  Líneas de Distribución │ CM01
│  Transformadores │ CM02
│  Subestaciones │ CM03
│
│
│ ═════════════════════════════════════════════════════════════
│ HOJA: Líneas de Distribución (índice 1)
│ ═════════════════════════════════════════════════════════════
│
│   A   │ B   │ C │ ... │ ElementId │ CodIntBIM │ Categoría │ ...
│  ─────┼─────┼───┼─────┼───────────┼───────────┼───────────┼────
│ [Encabezados con fondo azul FF366092]
│ Voltaje │ Long │ Cap │ ... │ 12345 │ CM01-0001 │ Gen.Models │ ...
│ 13.8kV │ 250m │ 1000A │ ... │ 12346 │ CM01-0002 │ Gen.Models │ ...
│
│
│ ═════════════════════════════════════════════════════════════
│ HOJA: Transformadores (índice 2)
│ ═════════════════════════════════════════════════════════════
│
│   A   │ B   │ C │ ... │ ElementId │ CodIntBIM │ Categoría │ ...
│  ─────┼─────┼───┼─────┼───────────┼───────────┼───────────┼────
│ [Encabezados con fondo azul FF366092]
│ Potencia │ VP │ VS │ ... │ 23456 │ CM02-0001 │ Elec.Equip │ ...
│ 500kVA │ 13.8kV │ 480V │ ... │ ...
│
│
│ ═════════════════════════════════════════════════════════════
│ HOJA: Excepciones (Última, si existen excepciones)
│ ═════════════════════════════════════════════════════════════
│
│   ElementId │ CodIntBIM │ Categoría │ Familia │ Tipo │ ...
│  ───────────┼───────────┼───────────┼─────────┼──────┼────
│ [Encabezados con fondo azul FF366092]
│ 34567 │ No Asignado │ Lighting │ Fixture │ T1 │ ...
│ 34568 │ No existe │ Elec.Equip │ Switch │ T2 │ ...
│
└──────────────────────────────────────────────────────────────┘
```

---

## Mapeo de Prefijo a Tabla

```
CodIntBIM              → Prefijo (4 primeros chars) → Tabla Excel
                                                    → Nombre de Hoja
─────────────────────────────────────────────────────────────────
CM01-0001              → CM01 ──────────────────→ Líneas de Distribución
CM01-0002              → CM01 ──────────────────→ Líneas de Distribución
CM01-0003              → CM01 ──────────────────→ Líneas de Distribución

CM02-0001              → CM02 ──────────────────→ Transformadores
CM02-0002              → CM02 ──────────────────→ Transformadores

CM03-0001              → CM03 ──────────────────→ Subestaciones
CM03-0002              → CM03 ──────────────────→ Subestaciones

(sin CodIntBIM o vacío)                         → Excepciones
```

---

## Columnas en Cada Tabla

```
Orden de columnas en Excel:
┌─────────────────────────────────────────────────────────────┐
│ [Parámetros específicos del elemento] │ [Campos adicionales] │
├─────────────────────────────────────────────────────────────┤
│ • Voltaje (parámetro)                 │ • ElementId         │
│ • Longitud (parámetro)                │ • CodIntBIM         │
│ • Capacidad (parámetro)               │ • Categoría         │
│ • Marca (parámetro)                   │ • Familia           │
│ • ... (otros parámetros)              │ • Tipo              │
│                                        │ • Nombre_RVT        │
│                                        │ • Situación CodInt  │
│                                        │ • Elementos         │
└─────────────────────────────────────────────────────────────┘

Ejemplo de Datos:
─────────────────────────────────────────────────────────────────
Voltaje │ Long │ Cap │ ElementId │ CodIntBIM │ Categoría │ ...
─────────────────────────────────────────────────────────────────
13.8kV  │ 250m │ 1kA │ 12345     │ CM01-0001 │ GenModels │ ...
13.8kV  │ 150m │ 2kA │ 12346     │ CM01-0002 │ GenModels │ ...
```

---

## Campos Adicionales Explicados

| Campo | Valor | Fuente | Descripción |
|-------|-------|--------|-------------|
| **ElementId** | Número entero | Revit API | ID único del elemento en Revit |
| **CodIntBIM** | String (ej: CM01-0001) | Parámetro del elemento | Código identificador |
| **Categoría** | Texto (ej: Generic Models) | element.Category.Name | Categoría de Revit |
| **Familia** | Texto (ej: Línea Distribución) | element.Symbol.FamilyName | Nombre de familia |
| **Tipo** | Texto (ej: Tipo A) | element.Symbol.Name | Nombre de tipo/símbolo |
| **Nombre_RVT** | Nombre archivo (ej: Modelo.rvt) | linked_doc.PathName | Archivo RVT origen |
| **Situación CodIntBIM** | "Elemento no Anidado" | Lógica | Anidado o no anidado |
| **Elementos** | "único" o "varios" | Conteo | Cuántos elementos tiene el código |

---

## Requisitos de Configuración

### script.json

```json
{
  "codigos_planillas": {
    "CM01": "Líneas de Distribución",
    "CM02": "Transformadores",
    "CM03": "Subestaciones",
    "CM04": "Equipos de Protección",
    "CM05": "Paneles de Control",
    ...
  }
}
```

### extraer_modelos_bim.py (líneas ~28-31)

```python
DATA_DIR_EXT = os.path.join(
    os.path.expanduser("~"),
    r"AppData\Roaming\MyPyRevitExtention\PyRevitIT.extension\data"
)

SCRIPT_JSON_PATH = os.path.join(DATA_DIR_EXT, 'script.json')
FORMATEAR_SCRIPT = os.path.join(DATA_DIR_EXT, 'formatear_tablas_excel_v2.py')
PYTHON3_EXE = r"C:\Users\Zbook HP\AppData\Local\Programs\Python\Python313\python.exe"
```

Todos estos valores deben ajustarse a tu instalación local.

---

## Casos de Uso

### Caso 1: Elemento con CodIntBIM válido
```
Elemento → CodIntBIM: "CM01-0001"
         → Prefijo: "CM01"
         → Tabla: "Líneas de Distribución" (según script.json)
         → Va a hoja: "Líneas de Distribución"
```

### Caso 2: Elemento con CodIntBIM pero sin datos
```
Elemento → CodIntBIM: "" (parámetro existe pero vacío)
         → Va a hoja: "Excepciones"
         → CodIntBIM mostrará: "No Asignado"
```

### Caso 3: Elemento sin parámetro CodIntBIM
```
Elemento → No tiene parámetro CodIntBIM
         → Va a hoja: "Excepciones"
         → CodIntBIM mostrará: "No existe"
```

### Caso 4: Elemento con prefijo no mapeado
```
Elemento → CodIntBIM: "XX01-0001"
         → Prefijo: "XX01"
         → No existe en script.json
         → Tabla: "Sin_mapeo_XX01"
         → Va a hoja: "Sin_mapeo_XX01"
```

---

## Ventajas de la Nueva Arquitectura

✅ **Acceso directo a Revit** - Sin necesidad de Excel intermediario  
✅ **Todos los parámetros** - Se extraen automáticamente del modelo  
✅ **Flexible** - Usa diccionario configurable de mappings  
✅ **Selección de carpeta** - Usuario elige dónde guardar  
✅ **Excepciones claramente identificadas** - En hoja separada  
✅ **Estructura profesional** - Índice, tablas ordenadas, formato limpio  
✅ **Escalable** - Funciona con cualquier número de tablas  
✅ **Robusto** - Manejo de errores en cada paso  

---

**Versión:** 2.0  
**Fecha:** Enero 2026  
**Listo para implementar:** ✅
