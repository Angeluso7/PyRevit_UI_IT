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


# ╔╦╗╔═╗╦╔╗╔
# ║║║╠═╣║║║║
# ╩ ╩╩ ╩╩╝╚╝
#==================================================

# Carpeta común de datos
DATA_DIR = os.path.join(
    os.path.expanduser("~"),
    r"AppData\\Roaming\\MyPyRevitExtention\\PyRevitIT.extension\\data"
)

# Python 3 (ajusta si cambia)
PYTHON_EXE = r"C:\\Users\\Zbook HP\\AppData\\Local\\Programs\\Python\\Python313\\pythonw.exe"


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
        "activo": {},
        "links": []
    }

    if doc is None:
        return info

    # GUID del archivo activo (usando PathName como fuente)
    try:
        import uuid
        path = doc.PathName or ""
        if path:
            uid = str(uuid.uuid5(uuid.NAMESPACE_URL, path))
        else:
            uid = ""
        info["activo"] = {
            "nombre": os.path.basename(path) if path else "(no guardado)",
            "unique_id": uid,
            "path": path
        }
    except Exception:
        info["activo"] = {"nombre": "(desconocido)", "unique_id": "", "path": ""}

    # GUID de archivos linkeados
    try:
        import uuid
        col = FilteredElementCollector(doc).OfClass(RevitLinkInstance)
        seen_paths = set()
        for li in col:
            try:
                link_doc = li.GetLinkDocument()
                if link_doc is None:
                    continue
                lpath = link_doc.PathName or ""
                if not lpath or lpath in seen_paths:
                    continue
                seen_paths.add(lpath)
                luid = str(uuid.uuid5(uuid.NAMESPACE_URL, lpath))
                info["links"].append({
                    "nombre": os.path.basename(lpath),
                    "unique_id": luid,
                    "path": lpath
                })
            except Exception:
                continue
    except Exception:
        pass

    return info


def run_cpython_script(script_name, args_list):
    """Ejecuta un script CPython en la carpeta de este botón."""
    try:
        this_folder = os.path.dirname(os.path.abspath(__file__))
    except Exception:
        this_folder = os.getcwd()

    script_path = os.path.join(this_folder, script_name)

    if not os.path.exists(PYTHON_EXE):
        forms.alert(
            u"No se encontró el ejecutable de Python 3 en:\n{}\n\nAjusta PYTHON_EXE en script.py.".format(
                PYTHON_EXE
            ),
            title="Python no encontrado"
        )
        return 1

    if not os.path.exists(script_path):
        forms.alert(
            u"No se encontró el script CPython '{}'.\n\nRuta:\n{}".format(script_name, script_path),
            title="Script CPython no encontrado"
        )
        return 1

    # Asegurar carpeta de datos
    if not os.path.exists(DATA_DIR):
        os.makedirs(DATA_DIR)

    cmd = [PYTHON_EXE, script_path] + args_list

    try:
        subprocess.Popen(cmd, cwd=this_folder)
        return 0
    except Exception as e:
        forms.alert(
            "Error al ejecutar CPython:\n{}\n\nComando:\n{}".format(e, " ".join(cmd)),
            title="Error CPython"
        )
        return 1


def main():
    if doc is None:
        forms.alert(
            "No hay documento activo.\nAbre un archivo Revit antes de usar el gestor de proyectos.",
            title="Error"
        )
        return

    # Construir info de GUIDs de modelo activo + links
    docs_info = build_docs_info()
    docs_info_json = json.dumps(docs_info)

    # Ejecutar gestor_proyectos.py (CPython) → ventana Tkinter
    rc = run_cpython_script("gestor_proyectos.py", [DATA_DIR, docs_info_json])
    if rc != 0:
        return


if __name__ == "__main__":
    main()