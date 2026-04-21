# -*- coding: utf-8 -*-
__title__ = "Carga de datos"
__doc__   = """Version = 2.1
Date    = 21.04.2026
________________________________________________________________
Description:

Carga datos desde una planilla .xlsm (hoja CODIGO) y los combina
con el repositorio activo del proyecto. Usa scripts CPython para
leer el Excel y IronPython/Revit API para cruzar con el modelo.

Requiere data/master/config_proyecto_activo.json con:
  - ruta_repositorio_activo
  (python_exe ya NO es necesario en el JSON → se resuelve
   automáticamente via core.env_config / data/temp/env_cache.json)
________________________________________________________________
How-To:

1. Asegúrate de tener configurado el proyecto activo.
2. Selecciona el archivo .xlsm con la hoja CODIGO.
3. El script lee los códigos, los cruza con el modelo y actualiza
   el repositorio.
________________________________________________________________
Last Updates:
- [21.04.2026] v2.1 python_exe via core.env_config (env_cache.json).
               Rutas desde lib/config/paths.py (CONFIG_PROYECTO,
               TEMP_DIR, MASTER_DIR). Eliminado config_utils.
- [21.04.2026] v2.0 Refactor: rutas centralizadas, CPython en
               scripts_cpython/, temp en data/temp/.
- [15.06.2024] v1.0 Versión original.
________________________________________________________________
Author: Argenis Angel"""

# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝
#==================================================
import os
import sys
import subprocess
from datetime import datetime

import clr
from pyrevit import forms

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import Document

# ── Añadir lib/ al path (resuelto desde la ubicación de este script) ─────────
# Estructura esperada:
#   EXT_ROOT/
#     lib/              ← se añade al path
#       config/paths.py
#       core/env_config.py
#     data/
#       master/config_proyecto_activo.json
#       temp/env_cache.json
#     scripts_cpython/
#       carga_excel.py
#     PyRevitIT.tab/
#       .../<Panel>.panel/
#         Carga de datos.pushbutton/
#           script.py   ← este archivo (5 niveles arriba = EXT_ROOT)

_THIS_DIR = os.path.dirname(os.path.abspath(__file__))
_EXT_ROOT = os.path.normpath(os.path.join(_THIS_DIR, "..", "..", "..", "..", ".."))
_LIB_DIR  = os.path.join(_EXT_ROOT, "lib")
if _LIB_DIR not in sys.path:
    sys.path.insert(0, _LIB_DIR)

# ── Importar rutas centralizadas desde lib/config/paths.py ──────────────────
from config.paths import (
    DATA_DIR,
    MASTER_DIR,
    TEMP_DIR,
    LOG_DIR,
    CONFIG_PROYECTO,        # data/master/config_proyecto_activo.json
    ensure_runtime_dirs,
)

# ──────────────────────────────────────────────────────────────────────────────
# BLOQUE ENV_CONFIG: resolución portátil de python_exe
# ──────────────────────────────────────────────────────────────────────────────
# get_python_exe() realiza la siguiente cascada:
#   1. Lee  data/temp/env_cache.json  → si la ruta es válida, la devuelve.
#   2. Si no existe o el exe fue movido, busca en el sistema:
#        a) AppData\Local\Programs\Python\Python3*\python.exe
#        b) Comando 'where python'
#        c) Recorre el PATH del sistema
#   3. Guarda el resultado en env_cache.json para futuras ejecuciones.
# → Un solo import, un solo call: el mismo bloque sirve para TODOS los botones.
from core.env_config import get_python_exe
# ──────────────────────────────────────────────────────────────────────────────

# ╦  ╦╔═╗╦═╗╦╔═╗╔╗ ╦  ╔═╗╔═╗
# ╚╗╔╝╠═╣╠╦╝║╠═╣╠╩╗║  ║╣ ╚═╗
#  ╚╝ ╩ ╩╩╚═╩╩ ╩╚═╝╩═╝╚═╝╚═╝
#==================================================
try:
    doc = __revit__.ActiveUIDocument.Document
except Exception:
    doc = None

# ╔╦╗╔═╗╦╔╗╔
# ║║║╠═╣║║║║
# ╩ ╩╩ ╩╩╝╚╝
#==================================================

CREATE_NO_WINDOW = 0x08000000

# Scripts CPython
_SCRIPTS_CPYTHON = os.path.join(_EXT_ROOT, "scripts_cpython")
CARGA_EXCEL_PY   = os.path.join(_SCRIPTS_CPYTHON, "carga_excel.py")

# Archivo temporal de intercambio entre CPython e IronPython
REPO_TMP_PATH = os.path.join(TEMP_DIR, "repo_tmp_codigos.txt")


def _log(msg):
    """Log en data/logs/carga_datos.log."""
    try:
        ensure_runtime_dirs()   # garantiza que LOG_DIR exista
        log_path = os.path.join(LOG_DIR, "carga_datos.log")
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(u"[{}] {}\n".format(ts, msg))
    except Exception:
        pass


def _get_repo_activo_path():
    """
    Lee ruta_repositorio_activo desde data/master/config_proyecto_activo.json.
    Lanza ValueError si no está configurado.
    """
    import json
    if not os.path.isfile(CONFIG_PROYECTO):
        raise ValueError(
            u"No se encontró config_proyecto_activo.json en:\n{}".format(CONFIG_PROYECTO)
        )
    with open(CONFIG_PROYECTO, "r", encoding="utf-8") as f:
        cfg = json.load(f)
    ruta = (cfg.get("ruta_repositorio_activo") or "").strip()
    if not ruta:
        raise ValueError(
            u"La clave 'ruta_repositorio_activo' está vacía en config_proyecto_activo.json."
        )
    if not os.path.isfile(ruta):
        raise ValueError(
            u"El repositorio activo no existe en la ruta configurada:\n{}".format(ruta)
        )
    return ruta


def select_excel_file():
    """Abre diálogo para seleccionar un .xlsm."""
    try:
        file_path = forms.pick_file(
            file_ext="xlsm",
            multi_file=False,
            title="Seleccionar planilla Excel CODIGO (.xlsm)"
        )
    except Exception as e:
        forms.alert(u"Error abriendo el selector de archivo:\n{}".format(e), title="Error")
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


def run_carga_excel(python_exe, xlsm_path, tmp_path):
    """
    Ejecuta scripts_cpython/carga_excel.py via CPython.
    Devuelve (returncode, stdout, stderr).
    """
    if not os.path.isfile(python_exe):
        forms.alert(
            u"No se encontró el ejecutable Python 3:\n{}\n\n"
            u"Elimina data/temp/env_cache.json y vuelve a ejecutar "
            u"para que el sistema lo redetecte.".format(python_exe),
            title="Python no encontrado"
        )
        return 1, "", "Python no encontrado"

    if not os.path.isfile(CARGA_EXCEL_PY):
        forms.alert(
            u"No se encontró carga_excel.py en:\n{}".format(CARGA_EXCEL_PY),
            title="Script CPython no encontrado"
        )
        return 1, "", "Script no encontrado"

    # Asegurar carpeta temp
    ensure_runtime_dirs()

    cmd = [python_exe, CARGA_EXCEL_PY, xlsm_path, tmp_path]
    try:
        proc = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            creationflags=CREATE_NO_WINDOW
        )
        out, err = proc.communicate()
    except Exception as e:
        forms.alert(u"Error ejecutando carga_excel.py:\n{}".format(e), title="Error CPython")
        return 1, "", str(e)

    out_str = out.decode("utf-8", errors="ignore").strip()
    err_str = err.decode("utf-8", errors="ignore").strip()
    return proc.returncode, out_str, err_str


def main():
    _log("==== Inicio Carga de Datos ====")

    if doc is None:
        forms.alert(u"No hay documento activo de Revit.", title="Error")
        _log("doc es None")
        return

    # Garantizar que existan todas las carpetas de runtime
    ensure_runtime_dirs()

    # ── Verificar repositorio activo ─────────────────────────────────────────
    try:
        repo_path = _get_repo_activo_path()
    except ValueError as e:
        forms.alert(u"{}".format(e), title="Sin proyecto configurado")
        _log("_get_repo_activo_path: {}".format(e))
        return

    _log("Repositorio activo: {}".format(repo_path))

    # ── Resolver python_exe (bloque env_config) ──────────────────────────────
    # Primera ejecución: busca en el sistema y guarda en env_cache.json.
    # Ejecuciones siguientes: lee directo del cache (rápido).
    python_exe = get_python_exe()
    if not python_exe:
        forms.alert(
            u"No se encontró Python 3 instalado en este equipo.\n\n"
            u"Instala Python 3.x y vuelve a ejecutar el botón.\n"
            u"(También puedes eliminar data/temp/env_cache.json "
            u"para forzar una nueva búsqueda.)",
            title="Python no encontrado"
        )
        _log("get_python_exe devolvió None")
        return

    _log("python_exe resuelto: {}".format(python_exe))

    # ── Verificar script.json ─────────────────────────────────────────────────
    from config.paths import SCRIPT_JSON_PATH_LIB as script_json
    if not os.path.isfile(script_json):
        forms.alert(
            u"No se encontró script.json en:\n{}".format(script_json),
            title="Error"
        )
        _log("script.json no encontrado: {}".format(script_json))
        return

    # ── Seleccionar .xlsm ─────────────────────────────────────────────────────
    xlsm_path = select_excel_file()
    if not xlsm_path:
        _log("Selección de xlsm cancelada")
        return

    _log("xlsm seleccionado: {}".format(xlsm_path))

    # ── CPython: leer Excel → TXT temporal ───────────────────────────────────
    rc, out, err = run_carga_excel(python_exe, xlsm_path, REPO_TMP_PATH)
    if rc != 0:
        forms.alert(
            u"Error al leer el archivo Excel:\n\nSTDERR:\n{}\n\nSTDOUT:\n{}".format(
                err or "(vacío)", out or "(vacío)"
            ),
            title="Error CPython"
        )
        _log("run_carga_excel rc={} err={}".format(rc, err))
        return

    if not os.path.isfile(REPO_TMP_PATH):
        forms.alert(
            u"No se generó el archivo temporal con los datos del Excel.\n\nSalida:\n{}".format(
                out or "(vacío)"
            ),
            title="Sin datos temporales"
        )
        _log("REPO_TMP_PATH no generado: {}".format(REPO_TMP_PATH))
        return

    _log("carga_excel.py OK → {}".format(REPO_TMP_PATH))

    # ── IronPython: combinar datos con el modelo ──────────────────────────────
    try:
        import combinar_datos
        combinar_datos.main(REPO_TMP_PATH)
    except ImportError:
        if _THIS_DIR not in sys.path:
            sys.path.insert(0, _THIS_DIR)
        try:
            import combinar_datos
            combinar_datos.main(REPO_TMP_PATH)
        except Exception as e:
            forms.alert(u"Error ejecutando combinar_datos:\n{}".format(e), title="Error combinación")
            _log("combinar_datos error: {}".format(e))
            return
    except Exception as e:
        forms.alert(u"Error ejecutando combinar_datos:\n{}".format(e), title="Error combinación")
        _log("combinar_datos error: {}".format(e))
        return

    # ── Limpiar temporal ──────────────────────────────────────────────────────
    try:
        if os.path.isfile(REPO_TMP_PATH):
            os.remove(REPO_TMP_PATH)
            _log("Temporal eliminado: {}".format(REPO_TMP_PATH))
    except Exception as e:
        _log("No se pudo borrar temporal: {}".format(e))

    _log("==== Fin Carga de Datos ====")


if __name__ == "__main__":
    main()