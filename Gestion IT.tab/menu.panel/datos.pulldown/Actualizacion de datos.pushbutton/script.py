# -*- coding: utf-8 -*-
__title__ = "SAESA"
__doc__ = """Version = 2.5
Date    = 03.05.2026
________________________________________________________________
Description:

Abre el gestor de datos SAESA del proyecto activo.

Flujo:
  1. generar_datos_tmp.main(MASTER_DIR)
     Lee el repositorio activo y escribe data/master/datos_tmp.json.
  2. subprocess CPython: datos_proyecto.py <ruta_json> <MASTER_DIR>
     Abre el visor/editor Tkinter con el JSON generado.

Rutas: todas resueltas desde __file__ o config_utils. Sin hardcodeo.
________________________________________________________________
Last Updates:
- [03.05.2026] v2.5 Genera datos_tmp.json antes de abrir el visor;
               pasa ruta COMPLETA al json (no directorio) en argv[1];
               elimina comillas embebidas en cmd list.
- [03.05.2026] v2.4 Crea TEMP_DIR/MASTER_DIR antes de lanzar subproceso.
- [03.05.2026] v2.3 Fix rutas: importa config_utils (fuente centralizada).
- [02.05.2026] v2.2 Fix rutas: TEMP_DIR y MASTER_DIR por separado.
- [19.04.2026] v2.1 Fix: datos_proyecto.py se busca en _this_dir.
- [19.04.2026] v2.0 Fix rutas -> MASTER_DIR; crea carpetas necesarias.
- [18.04.2026] v1.1 Version anterior.
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
_LIB_DIR  = os.path.join(_EXT_ROOT, 'lib')
if _LIB_DIR not in sys.path:
    sys.path.insert(0, _LIB_DIR)

# ── Importar rutas centralizadas desde config_utils ──────────────────────────
try:
    from config_utils import (
        EXT_ROOT, DATA_DIR, MASTER_DIR, TEMP_DIR, CACHE_DIR
    )
except Exception:
    # Fallback: mismo metodo que config_utils.py
    EXT_ROOT   = _EXT_ROOT
    DATA_DIR   = os.path.join(EXT_ROOT, 'data')
    MASTER_DIR = os.path.join(DATA_DIR, 'master')
    TEMP_DIR   = os.path.join(DATA_DIR, 'temp')
    CACHE_DIR  = os.path.join(DATA_DIR, 'cache')

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

# ── Rutas a los scripts CPython e IronPython del pushbutton ──────────────────
_DATOS_SCRIPT   = os.path.join(_this_dir, 'datos_proyecto.py')
_GENERAR_SCRIPT = os.path.join(_this_dir, 'generar_datos_tmp.py')

# ==============================================================================

CREATE_NO_WINDOW = 0x08000000


def _ensure_dirs():
    """Crea carpetas de runtime si no existen."""
    for d in (DATA_DIR, MASTER_DIR, TEMP_DIR, CACHE_DIR):
        if not os.path.exists(d):
            try:
                os.makedirs(d)
            except Exception:
                pass


def generar_datos_tmp():
    """
    Llama a generar_datos_tmp.main(MASTER_DIR) (IronPython/Revit API).
    Retorna la ruta completa al datos_tmp.json generado, o None si falla.
    """
    if _this_dir not in sys.path:
        sys.path.insert(0, _this_dir)
    try:
        import generar_datos_tmp as _gen
        return _gen.main(MASTER_DIR)
    except Exception as e:
        forms.alert(
            u"Error generando datos_tmp.json:\n{}".format(e),
            title=u"Error generacion"
        )
        return None


def run_datos_proyecto(json_path):
    """
    Lanza datos_proyecto.py via CPython.
    argv[1] = ruta completa al datos_tmp.json
    argv[2] = MASTER_DIR
    Sin comillas embebidas: subprocess con lista no las necesita.
    """
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

    # argv[1] = ruta COMPLETA al json (datos_proyecto.py la acepta directamente)
    # argv[2] = MASTER_DIR (para resolver config_proyecto_activo.json)
    cmd = [PYTHON_EXE, _DATOS_SCRIPT, json_path, MASTER_DIR]
    return subprocess.call(cmd, creationflags=CREATE_NO_WINDOW)


def main():
    _ensure_dirs()

    # ── Paso 1: generar datos_tmp.json desde el repositorio activo ─────────────
    json_path = generar_datos_tmp()
    if not json_path or not os.path.isfile(json_path):
        # generar_datos_tmp ya mostro el alert de error
        return

    # ── Paso 2: abrir visor CPython con el JSON generado ──────────────────────
    rc = run_datos_proyecto(json_path)
    if rc != 0:
        forms.alert(
            u"El gestor de datos termino con codigo: {}".format(rc),
            title=u"Advertencia"
        )


if __name__ == '__main__':
    main()
