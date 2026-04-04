# Checklist de Implementación & Troubleshooting

## ✅ Pre-Implementación

### Verificación de Carpetas

- [ ] Verificar existencia de: `%APPDATA%\Roaming\MyPyRevitExtention\PyRevitIT.extension\`
- [ ] Verificar existencia de carpeta `data\` dentro de extension
- [ ] Verificar existencia de carpeta `commands\` dentro de extension
- [ ] Verificar que `script.json` existe en carpeta `data\`

### Verificación de Python

- [ ] Python 3.13 está instalado en: `C:\Users\Zbook HP\AppData\Local\Programs\Python\Python313\`
- [ ] Ejecutable python.exe existe en esa ruta
- [ ] openpyxl está instalado en Python 3.13: `pip install openpyxl`

```bash
# Comando para verificar
C:\Users\Zbook HP\AppData\Local\Programs\Python\Python313\python.exe -m pip list | grep openpyxl
```

### Verificación de Revit & PyRevit

- [ ] PyRevit está funcionando en Revit
- [ ] Carpeta `commands\` es reconocida por PyRevit
- [ ] Otros scripts de PyRevit funcionan correctamente

---

## 📦 Pasos de Instalación

### Paso 1: Copiar Scripts

```
Copiar:  extraer_modelos_bim.py
A:       C:\Users\Zbook HP\AppData\Roaming\MyPyRevitExtention\PyRevitIT.extension\commands\

Copiar:  formatear_tablas_excel_v2.py
A:       C:\Users\Zbook HP\AppData\Roaming\MyPyRevitExtention\PyRevitIT.extension\data\
```

### Paso 2: Actualizar script.json

Editar `script.json` y asegurarse que contiene estructura similar a:

```json
{
  "codigos_planillas": {
    "CM01": "Líneas de Distribución",
    "CM02": "Transformadores",
    "CM03": "Subestaciones",
    ...
  }
}
```

**Importante:** Las claves (CM01, CM02, etc.) deben coincidir con los primeros 4 caracteres de tus CodIntBIM

### Paso 3: Actualizar Rutas en Scripts

#### En `extraer_modelos_bim.py` (líneas ~29-31):

```python
# Verifica que estas rutas sean correctas según tu instalación
DATA_DIR_EXT = os.path.join(
    os.path.expanduser("~"),
    r"AppData\Roaming\MyPyRevitExtention\PyRevitIT.extension\data"
)

SCRIPT_JSON_PATH = os.path.join(DATA_DIR_EXT, 'script.json')
FORMATEAR_SCRIPT = os.path.join(DATA_DIR_EXT, 'formatear_tablas_excel_v2.py')
PYTHON3_EXE = r"C:\Users\Zbook HP\AppData\Local\Programs\Python\Python313\python.exe"
```

Si tu username no es "Zbook HP", la ruta de PYTHON3_EXE debe ajustarse.

**Comando para encontrar ruta de Python:**
```bash
where python
# O si tienes múltiples versiones:
C:\Users\[TU_USERNAME]\AppData\Local\Programs\Python\Python313\python.exe --version
```

### Paso 4: Recargar PyRevit

- Cerrar Revit
- Eliminar carpeta `%LOCALAPPDATA%\pyRevit_cache` (si existe)
- Abrir Revit nuevamente
- PyRevit debería cargar los scripts

### Paso 5: Verificar en PyRevit

- Buscar botón "Extraer Modelos BIM" en PyRevit ribbon
- Debería estar disponible si todo se copió correctamente

---

## 🧪 Pruebas Básicas

### Test 1: Verificar Python 3.13

Abrir Command Prompt:
```cmd
C:\Users\Zbook HP\AppData\Local\Programs\Python\Python313\python.exe --version
```

Esperado: `Python 3.13.x`

### Test 2: Verificar openpyxl

```cmd
C:\Users\Zbook HP\AppData\Local\Programs\Python\Python313\python.exe -c "import openpyxl; print(openpyxl.__version__)"
```

Esperado: Número de versión (ej: 3.1.0)

### Test 3: Probar Script de Formateo Manualmente

```cmd
C:\Users\Zbook HP\AppData\Local\Programs\Python\Python313\python.exe ^
  "C:\Users\Zbook HP\AppData\Roaming\MyPyRevitExtention\PyRevitIT.extension\data\formatear_tablas_excel_v2.py" ^
  "C:\path\a\datos_test.json" ^
  "C:\output\test.xlsx"
```

Si funciona, debería crear un archivo Excel en `C:\output\test.xlsx`

### Test 4: Ejecutar desde Revit

1. Abrir documento de Revit con modelos linkeados
2. Ejecutar "Extraer Modelos BIM"
3. Seleccionar carpeta de salida
4. Esperar a que termine
5. Verificar archivo Excel en carpeta seleccionada

---

## ⚠️ Troubleshooting

### Problema 1: "No se encontró la carpeta data de la extensión"

**Causa:** Ruta DATA_DIR_EXT es incorrecta o carpeta no existe

**Solución:**
1. Verifica ruta: `%APPDATA%\Roaming\MyPyRevitExtention\PyRevitIT.extension\data\`
2. Si no existe, crea la carpeta manualmente
3. Verifica permisos de escritura
4. Actualiza ruta en script si username no es "Zbook HP"

```cmd
# Para verificar ruta APPDATA
echo %APPDATA%
```

---

### Problema 2: "No se encontró script.json"

**Causa:** Archivo script.json no está en carpeta data

**Solución:**
1. Verifica que `script.json` existe en `%APPDATA%\Roaming\MyPyRevitExtention\PyRevitIT.extension\data\`
2. Si no existe, crea uno usando `script_ejemplo.json` como referencia
3. Asegúrate que contiene diccionario `codigos_planillas`

---

### Problema 3: "No se encontró formatear_tablas_excel_v2.py"

**Causa:** Script de formateo no está en carpeta data

**Solución:**
1. Copiar `formatear_tablas_excel_v2.py` a `%APPDATA%\Roaming\MyPyRevitExtention\PyRevitIT.extension\data\`
2. Verifica nombre exacto del archivo (sin espacios extra)

---

### Problema 4: "Error al ejecutar script de formateo"

**Causa Posible A:** Python 3.13 no está instalado o ruta es incorrecta

**Solución A:**
```cmd
# Verifica ruta exacta de python.exe
dir "C:\Users\Zbook HP\AppData\Local\Programs\Python\Python313\"
# Debería mostrar: python.exe

# Si no existe, instala Python 3.13:
# https://www.python.org/downloads/
```

**Causa Posible B:** openpyxl no está instalado

**Solución B:**
```cmd
C:\Users\Zbook HP\AppData\Local\Programs\Python\Python313\python.exe -m pip install openpyxl
```

**Causa Posible C:** Permisos de archivo

**Solución C:**
1. Verifica que `formatear_tablas_excel_v2.py` no está en modo lectura
2. Haz click derecho → Propiedades → Seguridad
3. Asegúrate que tu usuario tiene permisos de lectura/escritura

---

### Problema 5: "Error al leer JSON intermedio"

**Causa:** Archivo JSON intermedio está corrupto o mal formado

**Solución:**
1. Verifica que no haya caracteres especiales en datos
2. Valida JSON usando herramienta online: https://jsonlint.com/
3. Verifica encoding UTF-8 en el archivo

---

### Problema 6: "No se encuentran modelos linkeados"

**Causa A:** El documento abierto no tiene modelos linkeados

**Solución A:**
1. Abre documento maestro que sí tenga modelos linkeados
2. Verifica que links están "cargados" (no desactivados)
3. Revisa en Revit: Manage → Manage Links → Revit

**Causa B:** Los elementos no tienen parámetro "CodIntBIM"

**Solución B:**
1. Abre archivo linkedeado
2. Verifica que existe parámetro "CodIntBIM" en elementos
3. Si no existe, crea parámetro compartido (debe ser String)
4. Asigna valores a elementos

---

### Problema 7: "Columnas en Excel están muy anchas/estrechas"

**Solución:**
No es problema. El script auto-ajusta, pero puedes modificar:

En `formatear_tablas_excel_v2.py`, función `auto_ajustar_columnas()`:

```python
ws.column_dimensions[col_letter].width = min(max_len + 2, 60)
# Cambiar 60 por otro valor máximo si lo deseas
```

---

### Problema 8: "Excel se abre con errores de compatibilidad"

**Solución:**
1. Abre en Excel
2. Si muestra alerta, elige "Repair" o "Sí" para reparar
3. El archivo debería funcionar normalmente
4. (Esto es normal en algunos casos)

---

### Problema 9: "No aparecen datos en hojas de tablas"

**Causa Posible A:** CodIntBIM está vacío en elementos

**Solución A:**
1. Verifica que elementos tienen valores en parámetro CodIntBIM
2. Verifica que valores comienzan con prefijos en script.json (CM01, CM02, etc.)
3. Verifica que hay al menos 4 caracteres en CodIntBIM

**Causa Posible B:** Prefijo no está en script.json

**Solución B:**
1. Abre script.json
2. Agrega prefijos faltantes: `"CM04": "Nombre Tabla"`
3. Recarga PyRevit
4. Intenta nuevamente

---

### Problema 10: "Hoja Excepciones está vacía pero debería tener datos"

**Solución:**
1. Verifica que elementos realmente no tienen CodIntBIM
2. Verifica que parámetro "CodIntBIM" existe en elementos
3. Si los elementos sí tienen valores válidos, no van a Excepciones

---

## 🔍 Verificación de Datos

### Verificar que extrajo datos correctamente:

1. Abre archivo Excel generado
2. Ve a hoja "Índice"
3. Deberías ver valores en B5+ y claves en C5+
4. Verifica que coinciden con script.json

### Verificar que agrupó correctamente por tabla:

1. Verifica que cada hoja tiene nombre de tabla (CM01, CM02, etc.)
2. Abre cada hoja
3. Debería haber datos con CodIntBIM del prefijo correspondiente

### Verificar campos adicionales:

1. Ve a última columna de cada hoja
2. Debería haber: ElementId, CodIntBIM, Categoría, Familia, Tipo, Nombre_RVT, Situación, Elementos

---

## 📋 Comandos Útiles para Diagnóstico

```bash
# Listar archivos en carpeta data
dir "%APPDATA%\Roaming\MyPyRevitExtention\PyRevitIT.extension\data\"

# Verificar si Python puede importar openpyxl
C:\Users\Zbook HP\AppData\Local\Programs\Python\Python313\python.exe ^
  -c "from openpyxl import load_workbook; print('OK')"

# Ejecutar script de formateo con output
C:\Users\Zbook HP\AppData\Local\Programs\Python\Python313\python.exe ^
  "C:\path\to\formatear_tablas_excel_v2.py" ^
  "C:\path\to\input.json" ^
  "C:\path\to\output.xlsx"

# Ver logs si hay error
# El error aparecerá en console output
```

---

## 📞 Soporte

Si encuentras problema no listado arriba:

1. **Verifica los logs** - Los mensajes de error son descriptivos
2. **Prueba paso a paso** - Test cada componente por separado
3. **Valida JSON** - Asegúrate que JSON intermedio es válido
4. **Verifica permisos** - Carpetas y archivos necesitan permisos de lectura/escritura
5. **Reinicia Revit** - A veces PyRevit necesita recargar

---

**Versión:** 2.0  
**Última actualización:** Enero 2026  
**Estado:** ✅ Listo para Implementar
