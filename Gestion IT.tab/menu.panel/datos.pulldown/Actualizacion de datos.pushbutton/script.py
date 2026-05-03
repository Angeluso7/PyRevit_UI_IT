# -*- coding: utf-8 -*-
__title__ = "SAESA"
__doc__ = """Version = 2.3
Date = 03.05.2026
________________________________________________________________
Description:

Abre el gestor de datos SAESA del proyecto activo.
________________________________________________________________
Last Updates:
- [03.05.2026] v2.3 Fix rutas: importa config_utils (fuente centralizada)
- [02.05.2026] v2.2 Fix rutas: se pasan TEMP_DIR y MASTER_DIR por separado
- [19.04.2026] v2.1 Fix: datos_proyecto.py se busca en _this_dir (pushbutton)
- [19.04.2026] v2.0 Fix rutas -> MASTER_DIR; crea carpetas necesarias
- [18.04.2026] v1.1 Version anterior
________________________________________________________________
Author: Argenis Angel"""

# ==============================================================================
import os
import sys
import subprocess
from pyrevit import forms

# ── Resolver _this_dir y agregar lib/ al path ─────────────────────────────────
try:
    _this_dir = os.path.dirname(os.path.abspath(__file__))
except Exception:
    _this_dir = os.getcwd()

# pushbutton(1) -> pulldown(2) -> panel(3) -> tab(4) -> EXT_ROOT
_EXT_ROOT = os.path.abspath(os.path.join(_this_dir, '..', '..', '..', '..'))
_LIB_DIR = os.path.join(_EXT_ROOT, 'lib')
if _LIB_DIR not in sys.path:
    sys.path.insert(0, _LIB_DIR)

# ── Importar rutas centralizadas desde config_utils ──────────────────────────
try:
    from config_utils import (
        EXT_ROOT, DATA_DIR, MASTER_DIR, TEMP_DIR, CACHE_DIR
    )
except Exception as _import_err:
    # Fallback robusto: usa expanduser igual que config_utils.py
    EXT_ROOT   = os.path.normpath(os.path.join(
        os.path.expanduser("~"),
        "AppData", "Roaming", "MyPyRevitExtention", "PyRevitIT.extension"
    ))
    DATA_DIR   = os.path.join(EXT_ROOT, "data")
    MASTER_DIR = os.path.join(DATA_DIR, "master")
    TEMP_DIR   = os.path.join(DATA_DIR, "temp")
    CACHE_DIR  = os.path.join(DATA_DIR, "cache")

# ── Resolver ejecutable Python 3 ─────────────────────────────────────────────
try:
    from core.env_config import get_python_exe
    PYTHON_EXE = get_python_exe()
except Exception:
    import glob as _glob

    def _fb_python():
        base = os.path.join(
            os.path.expanduser('~'),
            'AppData', 'Local', 'Programs', 'Python'
        )
        for exe in ('python.exe', 'pythonw.exe'):
            for cand in sorted(
                _glob.glob(os.path.join(base, 'Python3*', exe)),
                reverse=True
            ):
                return cand
        for folder in os.environ.get('PATH', '').split(os.pathsep):
            cand = os.path.join(folder.strip(), 'python.exe')
            if os.path.isfile(cand):
                return cand
        return None

    PYTHON_EXE = _fb_python()

# ==============================================================================
# datos_proyecto.py vive en la misma carpeta que este script (pushbutton)
_DATOS_SCRIPT = os.path.join(_this_dir, 'datos_proyecto.py')


def run_datos_proyecto():
    """Lanza datos_proyecto.py pasando TEMP_DIR y MASTER_DIR por separado."""
    if not os.path.isfile(_DATOS_SCRIPT):
        forms.alert(
            u"No se encontro el script CPython:\n{}".format(_DATOS_SCRIPT),
            title=u"Error"
        )
        return 1

    if not PYTHON_EXE or not os.path.isfile(PYTHON_EXE):
        forms.alert(
            u"No se encontro Python 3 instalado en este equipo.",
            title=u"Error"
        )
        return 1

    # argv[1] = TEMP_DIR   (donde vive datos_tmp.json)
    # argv[2] = MASTER_DIR (donde vive config_proyecto_activo.json)
    cmd = [PYTHON_EXE, _DATOS_SCRIPT, TEMP_DIR, MASTER_DIR]
    return subprocess.call(cmd)


def main():
    # Asegurar que existen todas las carpetas necesarias
    for d in (DATA_DIR, MASTER_DIR, TEMP_DIR, CACHE_DIR):
        if not os.path.exists(d):
            os.makedirs(d)

    rc = run_datos_proyecto()
    if rc != 0:
        forms.alert(
            u"El gestor de datos termino con codigo: {}".format(rc),
            title=u"Advertencia"
        )


if __name__ == '__main__':
    main()
