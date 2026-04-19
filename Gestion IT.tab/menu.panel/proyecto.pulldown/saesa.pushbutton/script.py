# -*- coding: utf-8 -*-
__title__   = "SAESA"
__doc__     = """Version = 2.0
Date    = 19.04.2026
________________________________________________________________
Description:

Abre el gestor de datos SAESA del proyecto activo.
________________________________________________________________
Last Updates:
- [19.04.2026] v2.0 Fix rutas -> MASTER_DIR; crea carpetas necesarias
- [18.04.2026] v1.1 Version anterior
________________________________________________________________
Author: Argenis Angel"""

#==================================================
import os
import sys
import subprocess
import json
from pyrevit import forms

# ── Rutas centralizadas desde config.paths ───────────────────────────────────
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

#==================================================
_CPYTHON_DIR = os.path.join(_EXT_ROOT, 'scripts_cpython')


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
    # Asegurar que existen todas las carpetas necesarias
    for d in (DATA_DIR, MASTER_DIR, TEMP_DIR, CACHE_DIR):
        if not os.path.exists(d):
            os.makedirs(d)

    # Se pasa MASTER_DIR para que registro y config queden en data/master/
    rc = run_cpython_script("datos_proyecto.py", [MASTER_DIR])
    if rc != 0:
        forms.alert(
            u"El gestor de datos termino con codigo: {}".format(rc),
            title=u"Advertencia"
        )


if __name__ == '__main__':
    main()
