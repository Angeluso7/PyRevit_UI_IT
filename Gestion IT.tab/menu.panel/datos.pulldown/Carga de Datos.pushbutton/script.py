# -*- coding: utf-8 -*-
__title__   = "Carga de datos"
__doc__     = """Version = 1.0
Date    = 15.06.2024
________________________________________________________________
Description:

This is the placeholder for a .pushbutton in a /pulldown
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
doc = __revit__.ActiveUIDocument.Document


# ╔╦╗╔═╗╦╔╗╔
# ║║║╠═╣║║║║
# ╩ ╩╩ ╩╩╝╚╝
#==================================================

# Carpeta común de datos
DATA_DIR = os.path.join(
    os.path.expanduser("~"),
    r"AppData\Roaming\MyPyRevitExtention\PyRevitIT.extension\data"
)

# Python 3.13 (ajusta si cambia)
PYTHON_EXE = r"C:\Users\Zbook HP\AppData\Local\Programs\Python\Python313\python.exe"

# Repositorio temporal (TXT con filas separadas por ;)
REPO_TMP_PATH = os.path.join(DATA_DIR, "repo_tmp_codigos.txt")


def select_excel_file():
    """Ventana de selección de archivo .xlsm antes de ejecutar CPython."""
    try:
        file_path = forms.pick_file(
            title="Seleccionar archivo Excel CODIGO"
        )
    except Exception as e:
        forms.alert("Error al abrir el cuadro de selección de archivos:\n{}".format(e), title="Error selección")
        return None

    if not file_path:
        return None

    if not file_path.lower().endswith(".xlsm"):
        forms.alert(
            "El archivo seleccionado no es .xlsm.\nSelecciona un archivo Excel con macros (*.xlsm).",
            title="Archivo no válido"
        )
        return None

    return file_path


def run_cpython_script(script_name, args_list):
    """Ejecuta un script CPython en la carpeta de este botón."""
    try:
        this_folder = os.path.dirname(os.path.abspath(__file__))
    except Exception:
        this_folder = os.getcwd()

    script_path = os.path.join(this_folder, script_name)

    if not os.path.exists(PYTHON_EXE):
        forms.alert(
            u"No se encontró el ejecutable de Python 3 en:\n{}\n\nAjusta PYTHON_EXE en script.py.".format(PYTHON_EXE),
            title="Python no encontrado"
        )
        return 1, "", "Python no encontrado"

    if not os.path.exists(script_path):
        forms.alert(
            u"No se encontró el script CPython '{}'.\n\nRuta:\n{}".format(script_name, script_path),
            title="Script CPython no encontrado"
        )
        return 1, "", "Script no encontrado"

    # Asegurar carpeta de datos
    if not os.path.exists(DATA_DIR):
        os.makedirs(DATA_DIR)

    cmd = [PYTHON_EXE, script_path] + args_list
    try:
        proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, cwd=this_folder)
        out, err = proc.communicate()
    except Exception as e:
        forms.alert("Error al ejecutar CPython:\n{}\n\nComando:\n{}".format(e, " ".join(cmd)), title="Error CPython")
        return 1, "", str(e)

    out_str = out.decode("utf-8", errors="ignore").strip()
    err_str = err.decode("utf-8", errors="ignore").strip()
    return proc.returncode, out_str, err_str


def main():
    if doc is None:
        forms.alert("No hay documento activo.\nAbre un archivo Revit antes de ejecutar la herramienta.", title="Error")
        return

    # 1) Seleccionar .xlsm
    xlsm_path = select_excel_file()
    if not xlsm_path:
        forms.alert("Operación cancelada. No se seleccionó archivo Excel válido.", title="Aviso")
        return

    # 2) Ejecutar carga_excel.py (CPython) → TXT temporal
    rc1, out1, err1 = run_cpython_script("carga_excel.py", [xlsm_path, REPO_TMP_PATH])
    if rc1 != 0:
        forms.alert(
            "Error ejecutando carga_excel.py:\n\nSTDERR:\n{}\n\nSTDOUT:\n{}".format(
                err1 or "(vacío)", out1 or "(vacío)"
            ),
            title="Error CPython"
        )
        return

    if not os.path.exists(REPO_TMP_PATH):
        forms.alert(
            "No se generó el repositorio temporal con las filas del Excel.\n\nSalida:\n{}".format(out1 or "(vacío)"),
            title="Sin datos temporales"
        )
        return

    # 3) Ejecutar combinar_datos.py (IronPython) desde el mismo botón
    try:
        this_folder = os.path.dirname(os.path.abspath(__file__))
    except Exception:
        this_folder = os.getcwd()

    combinar_path = os.path.join(this_folder, "combinar_datos.py")
    if not os.path.exists(combinar_path):
        forms.alert(
            u"No se encontró 'combinar_datos.py' en la carpeta del botón.\n\nRuta:\n{}".format(combinar_path),
            title="Error combinación"
        )
        return

    try:
        import imp
        combinar_mod = imp.load_source("combinar_datos", combinar_path)
    except Exception as e:
        forms.alert("Error al cargar combinar_datos.py:\n{}".format(e), title="Error combinación")
        return

    try:
        combinar_mod.main(REPO_TMP_PATH)
    except Exception as e:
        forms.alert("Error dentro de combinar_datos.py:\n{}".format(e), title="Error combinación")
        return

        # 4) Limpiar repositorio temporal
    try:
        if os.path.exists(REPO_TMP_PATH):
            os.remove(REPO_TMP_PATH)
    except Exception as e:
        forms.alert("No se pudo borrar el repositorio temporal:\n{}".format(e), title="Aviso limpieza")


if __name__ == "__main__":
    main()