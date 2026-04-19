# -*- coding: utf-8 -*-
__title__ = "Carga de datos"
__doc__ = """Version = 1.3
Date    = 19.04.2026
________________________________________________________________
Description:
Selecciona un archivo Excel (.xlsm) con hoja CODIGO,
extrae sus filas via CPython (carga_excel.py) y las combina
con el repositorio activo del proyecto (combinar_datos.py).
________________________________________________________________
Last Updates:
- [19.04.2026] v1.3  Script completo y limpio.
                     Eliminada _find_python() duplicada.
                     get_python_exe() unificado en lib/core/env_config.
                     Fallback interno si lib/core no esta disponible.
- [18.04.2026] v1.1  Autodeteccion de python.exe.
- [15.06.2024] v1.0  Inicio de aplicacion.
________________________________________________________________
Author: Argenis Angel"""

# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝
import os
import sys
import glob
import subprocess
import clr
from pyrevit import forms

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import Document

# ╦  ╦╔═╗╦═╗╦╔═╗╔╗ ╦  ╔═╗╔═╗
# ╚╗╔╝╠═╣╠╦╝║╠═╣╠╩╗║  ║╣ ╚═╗
#  ╚╝ ╩ ╩╩╚═╩╩ ╩╚═╝╩═╝╚═╝╚═╝
doc = __revit__.ActiveUIDocument.Document

# ── Rutas centralizadas ────────────────────────────────────────
EXT_ROOT = os.path.join(
    os.path.expanduser("~"),
    "AppData", "Roaming", "MyPyRevitExtention", "PyRevitIT.extension"
)
DATA_DIR    = os.path.join(os.path.expanduser("~"), r"AppData\Roaming\...")
CONFIG_PATH = os.path.join(DATA_DIR, "config_proyecto_activo.json")
MASTER_DIR = os.path.join(DATA_DIR, "master")
TEMP_DIR   = os.path.join(DATA_DIR, "temp")
LIB_DIR    = os.path.join(EXT_ROOT, "lib")

if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

REPO_TMP_PATH = os.path.join(TEMP_DIR, "repo_tmp_codigos.txt")

try:
    CURRENT_FOLDER = os.path.dirname(os.path.abspath(__file__))
except Exception:
    CURRENT_FOLDER = os.getcwd()

# ╔═╗╦ ╦╔╦╗╦ ╦╔═╗╔╗╔  ╔═╗═╗ ╦╔═╗
# ╠═╝╚╗╔╝ ║ ╠═╣║ ║║║║  ║╣ ╔╩╦╝║╣
# ╩   ╚╝  ╩ ╩ ╩╚═╝╝╚╝  ╚═╝╩ ╚═╚═╝

def _find_python_fallback():
    """
    Fallback interno: busca python.exe sin dependencias externas.
    Se usa solo si lib/core/env_config no esta disponible.
    """
    base = os.path.join(
        os.path.expanduser("~"),
        "AppData", "Local", "Programs", "Python"
    )
    for exe_name in ("python.exe", "pythonw.exe"):
        if os.path.isdir(base):
            pattern = os.path.join(base, "Python3*", exe_name)
            candidates = sorted(glob.glob(pattern), reverse=True)
            if candidates:
                return candidates[0]
        for folder in os.environ.get("PATH", "").split(os.pathsep):
            candidate = os.path.join(folder.strip(), exe_name)
            if os.path.isfile(candidate):
                return candidate
    return None


def _resolve_python_exe():
    """
    Intenta usar lib/core/env_config (con cache).
    Si falla, usa el fallback interno.
    """
    try:
        from core.env_config import get_python_exe
        result = get_python_exe()
        if result:
            return result
    except Exception:
        pass
    return _find_python_fallback()


# Resolucion unica al cargar el modulo
PYTHON_EXE = _resolve_python_exe()

# ╔╦╗╔═╗╦╔╗╔
# ║║║╠═╣║║║║
# ╩ ╩╩ ╩╩╝╚╝

def select_excel_file():
    try:
        file_path = forms.pick_file(title=u"Seleccionar archivo Excel CODIGO")
    except Exception as e:
        forms.alert(
            u"Error al abrir el cuadro de seleccion:\n{}".format(e),
            title=u"Error seleccion"
        )
        return None

    if not file_path:
        return None

    if not file_path.lower().endswith(".xlsm"):
        forms.alert(
            u"El archivo seleccionado no es .xlsm.\n"
            u"Selecciona un archivo Excel con macros (*.xlsm).",
            title=u"Archivo no valido"
        )
        return None

    return file_path


def run_cpython_script(script_name, args_list):
    script_path = os.path.join(CURRENT_FOLDER, script_name)

    if not PYTHON_EXE:
        forms.alert(
            u"No se encontro Python 3 instalado en esta maquina.\n\n"
            u"Instala Python 3.x desde https://www.python.org/downloads/\n"
            u"y asegurate de marcar 'Add Python to PATH'.",
            title=u"Python no encontrado"
        )
        return 1, "", "Python no encontrado"

    if not os.path.exists(script_path):
        forms.alert(
            u"No se encontro el script CPython '{}'.\n\nRuta:\n{}".format(
                script_name, script_path),
            title=u"Script CPython no encontrado"
        )
        return 1, "", "Script no encontrado"

    for folder in (DATA_DIR, TEMP_DIR):
        if not os.path.exists(folder):
            os.makedirs(folder)

    cmd = [PYTHON_EXE, script_path] + args_list
    try:
        proc = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            cwd=CURRENT_FOLDER
        )
        out, err = proc.communicate()
    except Exception as e:
        forms.alert(
            u"Error al ejecutar CPython:\n{}\n\nComando:\n{}".format(
                e, " ".join(cmd)),
            title=u"Error CPython"
        )
        return 1, "", str(e)

    out_str = out.decode("utf-8", errors="ignore").strip()
    err_str = err.decode("utf-8", errors="ignore").strip()
    return proc.returncode, out_str, err_str


def main():
    if doc is None:
        forms.alert(
            u"No hay documento activo.\n"
            u"Abre un archivo Revit antes de ejecutar la herramienta.",
            title=u"Error"
        )
        return

    # Validacion temprana de Python
    if not PYTHON_EXE:
        forms.alert(
            u"No se encontro Python 3 instalado en esta maquina.\n\n"
            u"Instala Python 3.x desde https://www.python.org/downloads/\n"
            u"y asegurate de marcar 'Add Python to PATH'.",
            title=u"Python no encontrado"
        )
        return

    # Validacion temprana de scripts CPython
    for script_name in ("carga_excel.py",):
        if not os.path.exists(os.path.join(CURRENT_FOLDER, script_name)):
            forms.alert(
                u"No se encontro '{}' en la carpeta del boton.\n\nRuta:\n{}".format(
                    script_name, CURRENT_FOLDER),
                title=u"Archivo faltante"
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
                err1 or "(vacio)", out1 or "(vacio)"),
            title=u"Error CPython"
        )
        return

    if not os.path.exists(REPO_TMP_PATH):
        forms.alert(
            u"No se genero el repositorio temporal.\n\nSalida:\n{}".format(
                out1 or "(vacio)"),
            title=u"Sin datos temporales"
        )
        return

    # 3) Cargar combinar_datos.py via IronPython (mismo proceso Revit)
    combinar_path = os.path.join(CURRENT_FOLDER, "combinar_datos.py")
    if not os.path.exists(combinar_path):
        forms.alert(
            u"No se encontro 'combinar_datos.py'.\n\nRuta:\n{}".format(combinar_path),
            title=u"Error combinacion"
        )
        return

    try:
        import imp
        combinar_mod = imp.load_source("combinar_datos", combinar_path)
    except Exception as e:
        forms.alert(
            u"Error al cargar combinar_datos.py:\n{}".format(e),
            title=u"Error combinacion"
        )
        return

    try:
        combinar_mod.main(REPO_TMP_PATH)
    except Exception as e:
        forms.alert(
            u"Error dentro de combinar_datos.py:\n{}".format(e),
            title=u"Error combinacion"
        )
        return

    # 4) Limpiar temporal
    try:
        if os.path.exists(REPO_TMP_PATH):
            os.remove(REPO_TMP_PATH)
    except Exception as e:
        forms.alert(
            u"No se pudo borrar el archivo temporal:\n{}".format(e),
            title=u"Aviso limpieza"
        )


if __name__ == "__main__":
    main()
