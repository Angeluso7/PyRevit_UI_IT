# -*- coding: utf-8 -*-
__title__ = "Revision\nElementos"
__doc__   = """Version = 2.0
Date    = 22.04.2026
________________________________________________________________
Description:
Extrae todos los elementos de modelos linkeados con CodIntBIM
asignado y genera un informe Excel (.xlsx) con formato vertical
por tabla (CMxx), colores por parametro y hoja de excepciones.
________________________________________________________________
Author: Argenis Angel"""

# ── Imports ──────────────────────────────────────────────────
import os
import sys
import subprocess
import json
import datetime
from pyrevit import forms

# ══════════════════════════════════════════════════════════════
#  1. LOG
# ══════════════════════════════════════════════════════════════
_LOG_FALLBACK = os.path.join(
    os.environ.get("TEMP", os.path.expanduser("~")),
    "revision_elementos_log.txt"
)

def log(msg):
    try:
        ts   = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ruta = LOG_PATH if "LOG_PATH" in globals() else _LOG_FALLBACK
        d    = os.path.dirname(ruta)
        if d and not os.path.exists(d):
            os.makedirs(d)
        with open(ruta, "a", encoding="utf-8") as f:
            f.write(u"[{}] {}\n".format(ts, msg))
    except Exception:
        pass

# ══════════════════════════════════════════════════════════════
#  2. RAIZ DE LA EXTENSION
# ══════════════════════════════════════════════════════════════
def _get_ext_root():
    # Intento 1: navegar 5 niveles desde __file__
    try:
        ruta = os.path.abspath(__file__)
        for _ in range(5):
            ruta = os.path.dirname(ruta)
        if os.path.isdir(os.path.join(ruta, "data", "master")):
            return ruta
        if os.path.isdir(ruta):
            return ruta
    except Exception:
        pass

    # Intento 2: API de pyRevit
    try:
        from pyrevit import extensions as exts
        for ext in exts.get_installed_ui_extensions():
            n = ext.name or ""
            d = ext.directory or ""
            if "PyRevitIT" in n or "PyRevit_UI_IT" in n or "PyRevitIT" in d:
                return d
    except Exception:
        pass

    # Intento 3: ruta conocida del equipo
    ruta_conocida = os.path.join(
        os.environ.get("APPDATA", ""),
        "MyPyRevitExtention", "PyRevitIT.extension"
    )
    if os.path.isdir(ruta_conocida):
        return ruta_conocida

    # Fallback
    try:
        ruta = os.path.abspath(__file__)
        for _ in range(5):
            ruta = os.path.dirname(ruta)
        return ruta
    except Exception:
        return os.environ.get("APPDATA", "")

EXT_ROOT = _get_ext_root()

# ══════════════════════════════════════════════════════════════
#  3. RUTAS BASE
# ══════════════════════════════════════════════════════════════
DATA_MASTER      = os.path.join(EXT_ROOT, "data", "master")
DATA_OUTPUT      = os.path.join(EXT_ROOT, "data", "output")
SCRIPTS_CPYTHON  = os.path.join(EXT_ROOT, "scripts_cpython")

LOG_PATH         = os.path.join(DATA_OUTPUT, "revision_elementos_log.txt")

CONFIG_PROYECTO  = os.path.join(DATA_MASTER, "config_proyecto_activo.json")
SCRIPT_JSON      = os.path.join(DATA_MASTER, "script.json")

TEMP_DATOS_JSON  = os.path.join(DATA_OUTPUT, "_temp_datos.json")
FORMATEAR_SCRIPT = os.path.join(SCRIPTS_CPYTHON, "formatear_revision_xlsx.py")

log("==== Inicio: Revision Elementos ====")
log("EXT_ROOT: {}".format(EXT_ROOT))

# ══════════════════════════════════════════════════════════════
#  4. PYTHON 3 EXE
# ══════════════════════════════════════════════════════════════
CREATE_NO_WINDOW = 0x08000000

def get_python3_exe():
    try:
        with open(CONFIG_PROYECTO, "r", encoding="utf-8") as f:
            cfg = json.load(f)
        ruta = (cfg.get("python3_exe") or "").strip()
        if ruta and os.path.isfile(ruta):
            log("python3_exe desde config: {}".format(ruta))
            return ruta
    except Exception as e:
        log("get_python3_exe config error: {}".format(e))

    candidatos = [
        r"C:\Users\Zbook HP\AppData\Local\Programs\Python\Python313\python.exe",
        r"C:\Users\Zbook HP\AppData\Local\Programs\Python\Python312\python.exe",
        r"C:\Python313\python.exe",
        r"C:\Python312\python.exe",
    ]
    for c in candidatos:
        if os.path.isfile(c):
            log("python3_exe via fallback: {}".format(c))
            return c
    return "python"

PYTHON3_EXE = get_python3_exe()

# ══════════════════════════════════════════════════════════════
#  5. HELPERS
# ══════════════════════════════════════════════════════════════
def nombre_informe_timestamp():
    ahora  = datetime.datetime.now()
    dia    = int(ahora.day)
    mes    = int(ahora.month)
    anio   = int(ahora.year) % 100
    hora   = int(ahora.hour)
    minuto = int(ahora.minute)
    seg    = int(ahora.second)
    return "Inf_Revision_{:02d}{:02d}{:02d}_{:02d}{:02d}{:02d}.xlsx".format(
        dia, mes, anio, hora, minuto, seg
    )

def validar_archivos():
    faltantes = []
    for ruta, nombre in [
        (SCRIPT_JSON,      "script.json (data/master)"),
        (FORMATEAR_SCRIPT, "formatear_revision_xlsx.py (scripts_cpython)"),
    ]:
        if not os.path.exists(ruta):
            faltantes.append(u"  - {} \n    -> {}".format(nombre, ruta))
    if faltantes:
        forms.alert(
            u"Archivos faltantes:\n\n{}".format(u"\n".join(faltantes)),
            title="Error - Revision Elementos"
        )
        return False
    return True

def limpiar_temp_json():
    """Elimina _temp_datos.json si existe."""
    try:
        if os.path.exists(TEMP_DATOS_JSON):
            os.remove(TEMP_DATOS_JSON)
            log("_temp_datos.json eliminado.")
    except Exception as e:
        log("Advertencia al limpiar temp json: {}".format(e))

# ══════════════════════════════════════════════════════════════
#  6. MAIN
# ══════════════════════════════════════════════════════════════
def main():
    if not os.path.exists(DATA_OUTPUT):
        os.makedirs(DATA_OUTPUT)

    if not validar_archivos():
        return

    # Paso 1: seleccionar carpeta destino primero (antes de procesar)
    carpeta_salida = forms.pick_folder(
        title=u"Selecciona carpeta donde guardar el informe .xlsx"
    )
    if not carpeta_salida:
        log("Usuario cancelo seleccion de carpeta.")
        return

    # Paso 2: Extraer datos del modelo Revit → TEMP_DATOS_JSON
    log("Paso 2: importando extraer_modelos_bim...")
    try:
        from extraer_modelos_bim import ejecutar_extraccion_y_json
    except Exception as e:
        forms.alert(
            u"No se pudo importar extraer_modelos_bim.py:\n{}".format(e),
            title="Error"
        )
        log("Error import extraer_modelos_bim: {}".format(e))
        return

    ok = ejecutar_extraccion_y_json(
        script_json_path = SCRIPT_JSON,
        datos_json_path  = TEMP_DATOS_JSON,
        data_master_dir  = DATA_MASTER,
        log_fn           = log
    )
    if not ok:
        log("Extraccion fallida. Abortando.")
        return

    # Paso 3: llamar CPython para generar Excel
    ruta_xlsx = os.path.join(carpeta_salida, nombre_informe_timestamp())
    log("Paso 3: llamando formatear_revision_xlsx.py...")
    log("  xlsx destino: {}".format(ruta_xlsx))

    try:
        salida = subprocess.check_output(
            [
                PYTHON3_EXE,
                FORMATEAR_SCRIPT,
                TEMP_DATOS_JSON,     # sys.argv[1]
                ruta_xlsx,           # sys.argv[2]
                DATA_MASTER          # sys.argv[3] → para hallar colores_parametros.json
            ],
            stderr=subprocess.STDOUT,
            creationflags=CREATE_NO_WINDOW
        )
        log("CPython output: {}".format(salida))
    except subprocess.CalledProcessError as e:
        err = (e.output or b"").decode("utf-8", errors="replace")
        forms.alert(
            u"Error al generar Excel:\n\n{}".format(err[:2000]),
            title="Error - CPython"
        )
        log("CalledProcessError: {}".format(err))
        limpiar_temp_json()
        return
    except Exception as e:
        forms.alert(u"Error inesperado:\n{}".format(e), title="Error")
        log("Error subprocess: {}".format(e))
        limpiar_temp_json()
        return

    # Paso 4: limpiar JSON temporal
    limpiar_temp_json()

    # Confirmacion final
    if os.path.exists(ruta_xlsx):
        forms.alert(
            u"Informe generado exitosamente:\n\n{}".format(ruta_xlsx),
            title=u"Revision Elementos - OK"
        )
        log("Excel OK: {}".format(ruta_xlsx))
    else:
        forms.alert(
            u"El proceso termino sin errores pero no se encontro el archivo:\n{}".format(ruta_xlsx),
            title="Advertencia"
        )
        log("Archivo no encontrado post-proceso: {}".format(ruta_xlsx))

if __name__ == "__main__":
    main()