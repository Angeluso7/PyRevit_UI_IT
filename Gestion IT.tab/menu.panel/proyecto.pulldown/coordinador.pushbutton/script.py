# -*- coding: utf-8 -*-
__title__   = "Coordinador"
__doc__     = """Version = 2.1
Date    = 19.04.2026
________________________________________________________________
Description:

Gestiona proyectos IT vinculando el modelo Revit activo (y sus
links) a un registro central mediante GUIDs.

________________________________________________________________
Last Updates:
- [19.04.2026] v2.1 Fix: gestor_proyectos.py se busca en _this_dir (pushbutton)
- [19.04.2026] v2.0 Fix rutas -> MASTER_DIR; validacion modelo sin guardar
- [15.06.2024] v1.0 Version inicial
________________________________________________________________
Author: Angel Uso"""

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
doc = __revit__.ActiveUIDocument.Document  # type: Document


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
    _DATA_DIR            = os.path.join(_EXT_ROOT, 'data')
    DATA_DIR             = _DATA_DIR
    MASTER_DIR           = os.path.join(_DATA_DIR, 'master')
    TEMP_DIR             = os.path.join(_DATA_DIR, 'temp')
    CACHE_DIR            = os.path.join(_DATA_DIR, 'cache')
    CONFIG_PROYECTO      = os.path.join(MASTER_DIR, 'config_proyecto_activo.json')
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

# gestor_proyectos.py vive en la misma carpeta que este script (pushbutton)
_GESTOR_SCRIPT = os.path.join(_this_dir, 'gestor_proyectos.py')


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
    """
    nombre_activo = os.path.basename(doc.PathName) if doc.PathName else ""
    uid_activo    = doc.ProjectInformation.UniqueId if doc.ProjectInformation else ""

    info = {
        "activo": {
            "nombre":    nombre_activo,
            "unique_id": uid_activo
        },
        "links": []
    }

    for li in FilteredElementCollector(doc).OfClass(RevitLinkInstance):
        link_doc = li.GetLinkDocument()
        if link_doc:
            info["links"].append({
                "nombre":    os.path.basename(link_doc.PathName),
                "unique_id": link_doc.ProjectInformation.UniqueId if link_doc.ProjectInformation else ""
            })
    return info


def run_gestor(args=None):
    if not os.path.isfile(_GESTOR_SCRIPT):
        forms.alert(
            u"No se encontro el script CPython:\n{}".format(_GESTOR_SCRIPT),
            title=u"Error"
        )
        return 1
    if not PYTHON_EXE or not os.path.isfile(PYTHON_EXE):
        forms.alert(
            u"No se encontro Python 3 instalado en este equipo.",
            title=u"Error"
        )
        return 1
    cmd = [PYTHON_EXE, _GESTOR_SCRIPT] + (args or [])
    return subprocess.call(cmd)


def main():
    # Validar que el modelo este guardado
    if not doc.PathName:
        forms.alert(
            u"El modelo activo no ha sido guardado aun.\n"
            u"Guarda el archivo antes de gestionar proyectos.",
            title=u"Modelo sin guardar"
        )
        return

    # Asegurar que existen las carpetas necesarias
    for d in (DATA_DIR, MASTER_DIR, TEMP_DIR, CACHE_DIR):
        if not os.path.exists(d):
            os.makedirs(d)

    docs_info      = build_docs_info()
    docs_info_json = json.dumps(docs_info, ensure_ascii=False)

    # Se pasa MASTER_DIR para que registro y config queden en data/master/
    rc = run_gestor([MASTER_DIR, docs_info_json])
    if rc != 0:
        forms.alert(
            u"El gestor de proyectos termino con codigo: {}".format(rc),
            title=u"Advertencia"
        )


if __name__ == '__main__':
    main()
