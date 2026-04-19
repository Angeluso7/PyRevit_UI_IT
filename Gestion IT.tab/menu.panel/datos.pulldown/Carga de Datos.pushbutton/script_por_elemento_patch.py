# -*- coding: utf-8 -*-
# ==============================================================
# PARCHE PARA: Por Elemento.pushbutton/script.py
# Reemplaza el bloque de resolución de CONFIG_PATH / REPO_PATH
# usando config_utils.py centralizado.
#
# INSTRUCCIONES:
#   1. Copia config_utils.py a PyRevitIT.extension/lib/
#   2. Reemplaza el bloque "Rutas comunes de datos" de script.py
#      por este bloque:
# ==============================================================

import clr
import os
import sys
import json
import subprocess

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System.Windows.Forms")

from Autodesk.Revit.DB import (
    FilteredElementCollector, RevitLinkInstance, ViewSchedule,
)
from Autodesk.Revit.UI.Selection import ObjectType
from System.Windows.Forms import MessageBox

# ── Contexto Revit ──────────────────────────────────────────────────────────
app   = __revit__.Application
uidoc = __revit__.ActiveUIDocument
doc   = uidoc.Document

# ── Rutas base con config_utils ─────────────────────────────────────────────
_LIB_DIR = os.path.normpath(
    os.path.join(os.path.dirname(__file__), "..", "..", "lib")
)
if _LIB_DIR not in sys.path:
    sys.path.insert(0, _LIB_DIR)

try:
    from config_utils import (
        get_config_path, get_repo_activo_path,
        DATA_DIR, MASTER_DIR, TEMP_DIR
    )
    _config_utils_ok = True
except ImportError:
    _config_utils_ok = False

# Fallback manual si lib/ todavía no tiene config_utils.py
if not _config_utils_ok:
    DATA_DIR   = os.path.join(os.path.expanduser("~"),
                    r"AppData\Roaming\MyPyRevitExtention\PyRevitIT.extension\data")
    MASTER_DIR = os.path.join(DATA_DIR, "master")
    TEMP_DIR   = os.path.join(DATA_DIR, "temp")

    def get_config_path():
        for c in [os.path.join(MASTER_DIR, "config_proyecto_activo.json"),
                  os.path.join(DATA_DIR,   "config_proyecto_activo.json")]:
            if os.path.isfile(c):
                return c
        raise IOError(u"No se encontró config_proyecto_activo.json")

    def get_repo_activo_path():
        import json as _j
        with open(get_config_path(), "r", encoding="utf-8") as f:
            cfg = _j.load(f)
        ruta = (cfg.get("ruta_repositorio_activo") or "").strip()
        if not ruta:
            raise ValueError(u"ruta_repositorio_activo vacía en config.")
        return ruta

# Carpeta del pushbutton actual
CURRENT_FOLDER      = os.path.dirname(__file__)
SCRIPT_JSON_PATH    = os.path.join(DATA_DIR, "script.json")
PLANILLAS_ORDER_PATH = os.path.join(CURRENT_FOLDER, "planillas_headers_order.json")
REPO_ELEMENTO_TMP_PATH = os.path.join(CURRENT_FOLDER, "repo_elemento_tmp.json")

# Ejecutable CPython (ajustar si cambia la máquina / versión)
PYTHON_EXE = r"C:\Users\Zbook HP\AppData\Local\Programs\Python\Python313\pythonw.exe"
CPYTHON_VIEWER_PATH = os.path.join(CURRENT_FOLDER, "visor_elemento.pyw")

# ── Resolución del repositorio activo ──────────────────────────────────────
try:
    REPO_PATH = get_repo_activo_path()
except (IOError, ValueError) as e:
    MessageBox.Show(u"{}".format(e), u"Config no encontrada")
    REPO_PATH = None

# ── Utilidades JSON ─────────────────────────────────────────────────────────
def load_json(path, show_error=True):
    if not os.path.exists(path):
        if show_error:
            MessageBox.Show(u"No se encontró archivo JSON:\n{}".format(path), u"Error")
        return None
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        if show_error:
            MessageBox.Show(u"Error cargando JSON:\n{}".format(e), u"Error")
        return None

def save_json(data, path, show_error=True):
    try:
        folder = os.path.dirname(path)
        if folder and not os.path.exists(folder):
            os.makedirs(folder)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        if show_error:
            MessageBox.Show(u"Error guardando JSON:\n{}".format(e), u"Error")
        return False

def load_repo():
    if not REPO_PATH:
        return {}
    repo = load_json(REPO_PATH, show_error=False)
    return repo or {}

# ── El resto del script original continúa sin cambios desde aquí ────────────
# (seleccionar_elemento, armar_datos_elemento, etc.)
