# -*- coding: utf-8 -*-
__title__   = "Coordinador"
__doc__     = """Version = 1.0
Date    = 15.06.2024
________________________________________________________________
Description:

This is the placeholder for a .pushbutton
You can use it to start your pyRevit Add-In

________________________________________________________________
How-To:

1. [Hold ALT + CLICK] on the button to open its source folder.
You will be able to override this placeholder.

2. Automate Your Boring Work ;)

________________________________________________________________
TODO:
[FEATURE] - Describe Your ToDo Tasks Here
________________________________________________________________
Last Updates:
- [15.06.2024] v1.0 Change Description
- [10.06.2024] v0.5 Change Description
- [05.06.2024] v0.1 Change Description 
________________________________________________________________
Author: Erik Frits"""

# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝
#==================================================
import os
import sys
import subprocess
import json
import clr
from pyrevit import forms

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import Document, FilteredElementCollector, RevitLinkInstance


# ╦  ╦╔═╗╦═╗╦╔═╗╔╗ ╦  ╔═╗╔═╗
# ╚╗╔╝╠═╣╠╦╝║╠═╣╠╩╗║  ║╣ ╚═╗
#  ╚╝ ╩ ╩╩╚═╩╩ ╩╚═╝╩═╝╚═╝╚═╝
#==================================================
#app    = __revit__.Application
#uidoc  = __revit__.ActiveUIDocument
doc    = __revit__.ActiveUIDocument.Document #type:Document


# ── Rutas centralizadas desde config.paths ────────────────────────────────────
try:
    _this_dir = os.path.dirname(os.path.abspath(__file__))
except Exception:
    _this_dir = os.getcwd()

# pushbutton(1) -> pulldown(2) -> panel(3) -> tab(4) -> EXT_ROOT
_EXT_ROOT = os.path.abspath(os.path.join(_this_dir, '..', '..', '..', '..'))
_LIB_DIR  = os.path.join(_EXT_ROOT, 'lib')
if _LIB_DIR not in sys.path:
    sys.path.insert(0, _LIB_DIR)

try:
    from config.paths import DATA_DIR, MASTER_DIR, TEMP_DIR, CACHE_DIR, \
                             CONFIG_PROYECTO, REGISTRO_PROYECTOS, \
                             SCRIPT_JSON_PATH_LIB, ensure_runtime_dirs
    from core.env_config import get_python_exe
    ensure_runtime_dirs()
    PYTHON_EXE = get_python_exe()
except Exception as _path_err:
    _DATA_DIR       = os.path.join(_EXT_ROOT, 'data')
    DATA_DIR        = _DATA_DIR
    MASTER_DIR      = os.path.join(_DATA_DIR, 'master')
    TEMP_DIR        = os.path.join(_DATA_DIR, 'temp')
    CACHE_DIR       = os.path.join(_DATA_DIR, 'cache')
    CONFIG_PROYECTO = os.path.join(MASTER_DIR, 'config_proyecto_activo.json')
    REGISTRO_PROYECTOS   = os.path.join(MASTER_DIR, 'registro_proyectos.json')
    SCRIPT_JSON_PATH_LIB = os.path.join(MASTER_DIR, 'script.json')
    import glob as _glob
    def _fb_python():
        base = os.path.join(os.path.expanduser('~'), 'AppData', 'Local', 'Programs', 'Python')
        for exe in ('python.exe', 'pythonw.exe'):
            for cand in sorted(_glob.glob(os.path.join(base, 'Python3*', exe)), reverse=True):
                return cand
        for folder in os.environ.get('PATH', '').split(os.pathsep):
            cand = os.path.join(folder.strip(), 'python.exe')
            if os.path.isfile(cand):
                return cand
        return None
    PYTHON_EXE = _fb_python()

# ╔╦╗╔═╗╦╔╗╔
# ║║║╠═╣║║║║
# ╩ ╩╩ ╩╩╝╚╝
#==================================================

_CPYTHON_DIR = os.path.join(_EXT_ROOT, 'scripts_cpython')


def build_docs_info():
    """
    Construye un diccionario con info de GUIDs:
    {
      "activo": {"nombre": "Modelo.rvt", "unique_id": "<guid>"},
      "links": [
        {"nombre": "Link1.rvt", "unique_id": "<guid>"},
        ...
      ]
    }
    El UniqueId del documento se obtiene como GUID basado en su PathName.
    """
    info = {
        "activo": {
            "nombre": os.path.basename(doc.PathName) if doc.PathName else "<sin guardar>",
            "unique_id": doc.ProjectInformation.UniqueId if doc.ProjectInformation else ""
        },
        "links": []
    }
    for li in FilteredElementCollector(doc).OfClass(RevitLinkInstance):
        link_doc = li.GetLinkDocument()
        if link_doc:
            info["links"].append({
                "nombre": os.path.basename(link_doc.PathName),
                "unique_id": link_doc.ProjectInformation.UniqueId if link_doc.ProjectInformation else ""
            })
    return info


def run_cpython_script(script_name, args=None):
    script_path = os.path.join(_CPYTHON_DIR, script_name)
    if not os.path.isfile(script_path):
        forms.alert(
            u"No se encontro el script CPython:\n{}".format(script_path),
            title=u"Error"
        )
        return 1
    if not PYTHON_EXE or not os.path.isfile(PYTHON_EXE):
        forms.alert(
            u"No se encontro Python 3 instalado en este equipo.",
            title=u"Error"
        )
        return 1
    cmd = [PYTHON_EXE, script_path] + (args or [])
    return subprocess.call(cmd)


def main():
    if not os.path.exists(DATA_DIR):
        os.makedirs(DATA_DIR)

    docs_info = build_docs_info()
    docs_info_json = json.dumps(docs_info, ensure_ascii=False)
    rc = run_cpython_script("gestor_proyectos.py", [DATA_DIR, docs_info_json])
    if rc != 0:
        forms.alert(
            u"El gestor de proyectos termino con codigo: {}".format(rc),
            title=u"Advertencia"
        )


if __name__ == '__main__':
    main()
