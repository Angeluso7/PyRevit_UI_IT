# -*- coding: utf-8 -*-
__title__ = "Carga de datos"
__doc__   = """Version = 2.3
Date    = 21.04.2026
________________________________________________________________
Description:

Carga datos desde una planilla .xlsm (hoja CODIGO) y los combina
con el repositorio activo del proyecto. Usa scripts CPython para
leer el Excel y IronPython/Revit API para cruzar con el modelo.

Requiere data/master/config_proyecto_activo.json con:
  - ruta_repositorio_activo
  (python_exe resuelto automáticamente via core.env_config)
________________________________________________________________
Last Updates:
- [21.04.2026] v2.3 Rutas críticas ancladas al __file__ del botón.
               Evita el problema de paths.py resolviendo EXT_ROOT
               desde AppData en lugar del repo real.
- [21.04.2026] v2.2 _EXT_ROOT y CARGA_EXCEL_PY desde paths.py.
- [21.04.2026] v2.1 python_exe via core.env_config (env_cache.json).
- [21.04.2026] v2.0 Refactor: rutas centralizadas.
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

# ── Resolver EXT_ROOT desde __file__ del botón (siempre confiable) ────────────
# Estructura:
#   EXT_ROOT/
#     lib/                         ← se añade al sys.path
#     data/master/script.json      ← data/master/
#     data/master/config_proyecto_activo.json
#     data/temp/                   ← archivos temporales
#     data/logs/                   ← logs
#     scripts_cpython/carga_excel.py
#     Gestion IT.tab/
#       menu.panel/
#         datos.pulldown/
#           Carga de Datos.pushbutton/
#             script.py            ← ESTE archivo (4 niveles arriba = EXT_ROOT)

_THIS_DIR = os.path.dirname(os.path.abspath(__file__))
_EXT_ROOT = os.path.normpath(os.path.join(_THIS_DIR, "..", "..", "..", ".."))
_LIB_DIR  = os.path.join(_EXT_ROOT, "lib")

if _LIB_DIR not in sys.path:
    sys.path.insert(0, _LIB_DIR)

# ── Rutas críticas ancladas a _EXT_ROOT (no dependen de paths.py __file__) ───
# paths.py puede resolver EXT_ROOT hacia AppData\Roaming cuando pyRevit
# cachea el módulo. Por eso todas las rutas de datos se definen aquí.
_DATA_DIR       = os.path.join(_EXT_ROOT, "data")
_MASTER_DIR     = os.path.join(_DATA_DIR, "master")
_TEMP_DIR       = os.path.join(_DATA_DIR, "temp")
_LOG_DIR        = os.path.join(_DATA_DIR, "logs")

CONFIG_PROYECTO = os.path.join(_MASTER_DIR, "config_proyecto_activo.json")
SCRIPT_JSON     = os.path.join(_MASTER_DIR, "script.json")
CARGA_EXCEL_PY  = os.path.join(_EXT_ROOT, "scripts_cpython", "carga_excel.py")
REPO_TMP_PATH   = os.path.join(_TEMP_DIR, "repo_tmp_codigos.txt")

# ── Importar utilidades desde lib/ (solo funciones, no rutas) ─────────────────
from core.env_config import get_python_exe


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


def _ensure_dirs():
    """Crea carpetas de runtime si no existen."""
    for path in (_DATA_DIR, _MASTER_DIR, _TEMP_DIR, _LOG_DIR):
        if not os.path.exists(path):
            os.makedirs(path)


def _log(msg):
    """Log en data/logs/carga_datos.log."""
    try:
        _ensure_dirs()
        log_path = os.path.join(_LOG_DIR, "carga_datos.log")
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(u"[{}] {}\n".format(ts, msg))
    except Exception:
        pass


def _get_repo_activo_path():
    """
    Lee ruta_repositorio_activo desde data/master/config_proyecto_activo.json.
    Lanza ValueError si no está configurado o el archivo no existe.
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
            u"No se encontró carga_excel.py en:\n{}\n\n"
            u"Verifica que la carpeta scripts_cpython/ existe en:\n{}".format(
                CARGA_EXCEL_PY, _EXT_ROOT
            ),
            title="Script CPython no encontrado"
        )
        return 1, "", "Script no encontrado"

    _ensure_dirs()

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
    _log("EXT_ROOT resuelto: {}".format(_EXT_ROOT))

    if doc is None:
        forms.alert(u"No hay documento activo de Revit.", title="Error")
        _log("doc es None")
        return

    _ensure_dirs()

    # ── Verificar repositorio activo ─────────────────────────────────────────
    try:
        repo_path = _get_repo_activo_path()
    except ValueError as e:
        forms.alert(u"{}".format(e), title="Sin proyecto configurado")
        _log("_get_repo_activo_path: {}".format(e))
        return

    _log("Repositorio activo: {}".format(repo_path))

    # ── Resolver python_exe ──────────────────────────────────────────────────
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

    # ── Verificar script.json en data/master/ ────────────────────────────────
    if not os.path.isfile(SCRIPT_JSON):
        forms.alert(
            u"No se encontró script.json en:\n{}".format(SCRIPT_JSON),
            title="Error"
        )
        _log("script.json no encontrado: {}".format(SCRIPT_JSON))
        return

    _log("script.json OK: {}".format(SCRIPT_JSON))

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
    if _THIS_DIR not in sys.path:
        sys.path.insert(0, _THIS_DIR)

    try:
        import combinar_datos
        combinar_datos.main(REPO_TMP_PATH)
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