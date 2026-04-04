# -*- coding: utf-8 -*-
__title__   = "Saesa"
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
import clr
from pyrevit import forms

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import Document


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

DATA_DIR = os.path.join(
    os.path.expanduser("~"),
    r"AppData\\Roaming\\MyPyRevitExtention\\PyRevitIT.extension\\data"
)

PYTHON_EXE = r"C:\\Users\\Zbook HP\\AppData\\Local\\Programs\\Python\\Python313\\pythonw.exe"


def run_cpython_script(script_name, args_list):
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
            "No hay documento activo.\nAbre un archivo Revit antes de usar este botón.",
            title="Error"
        )
        return

    # Pasar DATA_DIR a CPython
    run_cpython_script("datos_proyecto.py", [DATA_DIR])


if __name__ == "__main__":
    main()