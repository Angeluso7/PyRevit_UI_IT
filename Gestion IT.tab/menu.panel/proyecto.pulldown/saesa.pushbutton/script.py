# -*- coding: utf-8 -*-
__title__   = "Saesa"
__doc__     = """Version = 2.0
Date    = 2024-06-15
________________________________________________________________
Description:
  Abre el visor de datos del proyecto SAESA activo.
  Requiere que exista config_proyecto_activo.json en data/master/.
________________________________________________________________
Author: Equipo IT"""

# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝
import os
import sys
import subprocess
from pyrevit import forms

# ── Agregar lib/ al sys.path (IronPython no lo hace automáticamente) ──────────
try:
    _this_dir = os.path.dirname(os.path.abspath(__file__))
except Exception:
    _this_dir = os.getcwd()

# Subir 5 niveles desde el pushbutton hasta la raíz de la extensión:
# saesa.pushbutton → proyecto.pulldown → menu.panel → Gestion IT.tab → EXT_ROOT
_ext_root = os.path.normpath(os.path.join(_this_dir, "..", "..", "..", ".."))
_lib_path = os.path.join(_ext_root, "lib")
if _lib_path not in sys.path:
    sys.path.insert(0, _lib_path)

# ── Importar utilidades del proyecto ─────────────────────────────────────────
try:
    from config_utils import MASTER_DIR
except ImportError:
    # Fallback: construir MASTER_DIR desde expanduser
    MASTER_DIR = os.path.join(
        os.path.expanduser("~"),
        "AppData", "Roaming", "MyPyRevitExtention",
        "PyRevitIT.extension", "data", "master"
    )

try:
    from core.env_config import get_python_exe
    PYTHON_EXE = get_python_exe()
except Exception:
    PYTHON_EXE = None

# ── Helpers ───────────────────────────────────────────────────────────────────
def _get_python_exe():
    """Devuelve PYTHON_EXE resuelto; muestra alerta si no se encontró."""
    if PYTHON_EXE and os.path.isfile(PYTHON_EXE):
        return PYTHON_EXE
    forms.alert(
        u"No se encontró un intérprete Python 3 en este equipo.\n\n"
        u"Instala Python 3 o verifica que esté en el PATH del sistema.",
        title=u"Python no encontrado"
    )
    return None

def _run_cpython(script_name, args_list):
    """Lanza un script CPython como subproceso independiente (Popen)."""
    exe = _get_python_exe()
    if not exe:
        return 1

    script_path = os.path.join(_this_dir, script_name)
    if not os.path.isfile(script_path):
        forms.alert(
            u"No se encontró el script CPython '{}'.\nRuta esperada:\n{}".format(
                script_name, script_path
            ),
            title=u"Script no encontrado"
        )
        return 1

    # Crear la carpeta de datos si no existe
    if not os.path.isdir(MASTER_DIR):
        os.makedirs(MASTER_DIR)

    cmd = [exe, script_path] + args_list
    try:
        subprocess.Popen(cmd, cwd=_this_dir)
        return 0
    except Exception as e:
        forms.alert(
            u"Error al lanzar el script CPython:\n{}\n\nComando:\n{}".format(
                e, " ".join(cmd)
            ),
            title=u"Error CPython"
        )
        return 1

# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    # Verificar que haya un documento activo en Revit
    try:
        doc = __revit__.ActiveUIDocument.Document  # noqa: F821
        if doc is None:
            raise AttributeError
    except Exception:
        forms.alert(
            u"No hay documento activo.\nAbre un archivo Revit antes de usar este botón.",
            title=u"Sin documento"
        )
        return

    # Pasar MASTER_DIR como único argumento a datos_proyecto.py
    _run_cpython("datos_proyecto.py", [MASTER_DIR])

main()
