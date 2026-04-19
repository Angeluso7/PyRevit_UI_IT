# -*- coding: utf-8 -*-
__title__ = "Carga de datos"
__doc__   = """Version = 1.1
Date    = 15.06.2024
________________________________________________________________
Description:
Selecciona un archivo Excel (.xlsm) con hoja CODIGO,
extrae sus filas via CPython (carga_excel.py) y las combina
con el repositorio activo del proyecto (combinar_datos.py).
________________________________________________________________
Last Updates:
- [18.04.2026] v1.1  _find_python() autodetecta pythonw/python.
                     Validaciones tempranas y mensajes claros.
- [15.06.2024] v1.0  Inicio de aplicación.
________________________________________________________________
Author: Argenis Angel"""

# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝
import os
import subprocess
import glob
import clr
from pyrevit import forms

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import Document

# ╦  ╦╔═╗╦═╗╦╔═╗╔╗ ╦  ╔═╗╔═╗
# ╚╗╔╝╠═╣╠╦╝║╠═╣╠╩╗║  ║╣ ╚═╗
#  ╚╝ ╩ ╩╩╚═╩╩ ╩╚═╝╩═╝╚═╝╚═╝
doc = __revit__.ActiveUIDocument.Document

DATA_DIR = os.path.join(
    os.path.expanduser("~"),
    "AppData", "Roaming", "MyPyRevitExtention", "PyRevitIT.extension", "data"
)
REPO_TMP_PATH = os.path.join(DATA_DIR, "repo_tmp_codigos.txt")

try:
    CURRENT_FOLDER = os.path.dirname(os.path.abspath(__file__))
except Exception:
    CURRENT_FOLDER = os.getcwd()


# ╔╦╗╔═╗╦╔╗╔
# ║║║╠═╣║║║║
# ╩ ╩╩ ╩╩╝╚╝

def _find_python():
    """
    Busca python.exe / pythonw.exe automáticamente:
      1. Carpetas Python3xx en AppData\\Local\\Programs\\Python\\
      2. Ejecutable en el PATH del sistema
    Devuelve la ruta completa o None.
    """
    base = os.path.join(os.path.expanduser("~"),
                        "AppData", "Local", "Programs", "Python")
    for exe_name in ("python.exe", "pythonw.exe"):
        if os.path.isdir(base):
            pattern = os.path.join(base, "Python3*", exe_name)
            candidates = sorted(glob.glob(pattern), reverse=True)
            if candidates:
                return candidates[0]
        # Fallback: buscar en PATH
        for folder in os.environ.get("PATH", "").split(os.pathsep):
            candidate = os.path.join(folder, exe_name)
            if os.path.isfile(candidate):
                return candidate
    return None


PYTHON_EXE = _find_python()


def select_excel_file():
    try:
        file_path = forms.pick_file(title="Seleccionar archivo Excel CODIGO")
    except Exception as e:
        forms.alert(u"Error al abrir el cuadro de selección:\n{}".format(e),
                    title="Error selección")
        return None

    if not file_path:
        return None

    if not file_path.lower().endswith(".xlsm"):
        forms.alert(
            u"El archivo seleccionado no es .xlsm.\n"
            u"Selecciona un archivo Excel con macros (*.xlsm).",
            title="Archivo no válido"
        )
        return None
    return file_path


def run_cpython_script(script_name, args_list):
    script_path = os.path.join(CURRENT_FOLDER, script_name)

    if not PYTHON_EXE:
        forms.alert(
            u"No se encontró Python 3 instalado en esta máquina.\n\n"
            u"Instala Python 3.x desde https://www.python.org/downloads/\n"
            u"y asegúrate de marcar 'Add Python to PATH' durante la instalación.",
            title="Python no encontrado"
        )
        return 1, "", "Python no encontrado"

    if not os.path.exists(script_path):
        forms.alert(
            u"No se encontró el script CPython '{}'.\n\nRuta:\n{}".format(
                script_name, script_path),
            title="Script CPython no encontrado"
        )
        return 1, "", "Script no encontrado"

    if not os.path.exists(DATA_DIR):
        os.makedirs(DATA_DIR)

    cmd = [PYTHON_EXE, script_path] + args_list
    try:
        proc = subprocess.Popen(
            cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
            cwd=CURRENT_FOLDER
        )
        out, err = proc.communicate()
    except Exception as e:
        forms.alert(
            u"Error al ejecutar CPython:\n{}\n\nComando:\n{}".format(
                e, " ".join(cmd)),
            title="Error CPython"
        )
        return 1, "", str(e)

    out_str = out.decode("utf-8", errors="ignore").strip()
    err_str = err.decode("utf-8", errors="ignore").strip()
    return proc.returncode, out_str, err_str


def main():
    if doc is None:
        forms.alert("No hay documento activo.\n"
                    "Abre un archivo Revit antes de ejecutar la herramienta.",
                    title="Error")
        return

    # Validación temprana de Python
    if not PYTHON_EXE:
        forms.alert(
            u"No se encontró Python 3 instalado en esta máquina.\n\n"
            u"Instala Python 3.x desde https://www.python.org/downloads/\n"
            u"y asegúrate de marcar 'Add Python to PATH'.",
            title="Python no encontrado"
        )
        return

    # Validación temprana de scripts CPython
    for script_name in ("carga_excel.py",):
        if not os.path.exists(os.path.join(CURRENT_FOLDER, script_name)):
            forms.alert(
                u"No se encontró '{}' en la carpeta del botón.\n\nRuta:\n{}".format(
                    script_name, CURRENT_FOLDER),
                title="Archivo faltante"
            )
            return

    # 1) Seleccionar .xlsm
    xlsm_path = select_excel_file()
    if not xlsm_path:
        return

    # 2) Ejecutar carga_excel.py (CPython) → TXT temporal
    rc1, out1, err1 = run_cpython_script("carga_excel.py", [xlsm_path, REPO_TMP_PATH])
    if rc1 != 0:
        forms.alert(
            u"Error ejecutando carga_excel.py:\n\nSTDERR:\n{}\n\nSTDOUT:\n{}".format(
                err1 or "(vacío)", out1 or "(vacío)"),
            title="Error CPython"
        )
        return

    if not os.path.exists(REPO_TMP_PATH):
        forms.alert(
            u"No se generó el repositorio temporal.\n\nSalida:\n{}".format(
                out1 or "(vacío)"),
            title="Sin datos temporales"
        )
        return

    # 3) Cargar combinar_datos.py via IronPython (mismo proceso Revit)
    combinar_path = os.path.join(CURRENT_FOLDER, "combinar_datos.py")
    if not os.path.exists(combinar_path):
        forms.alert(
            u"No se encontró 'combinar_datos.py'.\n\nRuta:\n{}".format(combinar_path),
            title="Error combinación"
        )
        return

    try:
        import imp
        combinar_mod = imp.load_source("combinar_datos", combinar_path)
    except Exception as e:
        forms.alert(u"Error al cargar combinar_datos.py:\n{}".format(e),
                    title="Error combinación")
        return

    try:
        combinar_mod.main(REPO_TMP_PATH)
    except Exception as e:
        forms.alert(u"Error dentro de combinar_datos.py:\n{}".format(e),
                    title="Error combinación")
        return

    # 4) Limpiar temporal
    try:
        if os.path.exists(REPO_TMP_PATH):
            os.remove(REPO_TMP_PATH)
    except Exception as e:
        forms.alert(u"No se pudo borrar el archivo temporal:\n{}".format(e),
                    title="Aviso limpieza")


if __name__ == "__main__":
    main()
