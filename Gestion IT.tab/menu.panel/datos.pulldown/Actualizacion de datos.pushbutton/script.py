# -*- coding: utf-8 -*-
__title__ = "SAESA"
__doc__ = """Version = 2.6
Date    = 03.05.2026
________________________________________________________________
Description:

Abre el gestor de datos SAESA del proyecto activo.

Flujo:
  1. _generar_datos_tmp(MASTER_DIR)
     Lee el repositorio activo y escribe data/master/datos_tmp.json.
  2. subprocess CPython: datos_proyecto.py <ruta_json> <MASTER_DIR>
     Abre el visor/editor Tkinter con el JSON generado.

Rutas: todas resueltas desde __file__ o config_utils. Sin hardcodeo.
________________________________________________________________
Last Updates:
- [03.05.2026] v2.6 Consolida generar_datos_tmp.py dentro de este script.
- [03.05.2026] v2.5 Genera datos_tmp.json antes de abrir el visor;
               pasa ruta COMPLETA al json en argv[1].
- [03.05.2026] v2.4 Crea TEMP_DIR/MASTER_DIR antes de lanzar subproceso.
- [03.05.2026] v2.3 Fix rutas: importa config_utils.
- [02.05.2026] v2.2 Fix rutas: TEMP_DIR y MASTER_DIR por separado.
- [19.04.2026] v2.1 Fix: datos_proyecto.py se busca en _this_dir.
- [19.04.2026] v2.0 Fix rutas -> MASTER_DIR.
- [18.04.2026] v1.1 Version anterior.
________________________________________________________________
Author: Argenis Angel"""

# ==============================================================================
import os
import sys
import json
import subprocess
from pyrevit import forms

import clr
clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import FilteredElementCollector, ViewSchedule

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

# ── Documento activo ────────────────────────────────────────────────────────────
try:
    doc = __revit__.ActiveUIDocument.Document
except Exception:
    doc = None

# ── Ruta al script CPython ────────────────────────────────────────────────────
_DATOS_SCRIPT = os.path.join(_this_dir, 'datos_proyecto.py')

CREATE_NO_WINDOW = 0x08000000


# ==============================================================================
# SECCION 1: Generacion de datos_tmp.json (IronPython / Revit API)
# Logica previamente en generar_datos_tmp.py, ahora consolidada aqui.
# ==============================================================================

def _get_repo_path_from_config(master_dir):
    """
    Lee config_proyecto_activo.json desde master_dir y retorna
    la ruta al repositorio activo. Lanza ValueError si falla.
    """
    config_path = os.path.join(master_dir, "config_proyecto_activo.json")
    if not os.path.isfile(config_path):
        raise ValueError(
            u"No se encontro config_proyecto_activo.json en:\n{}".format(config_path)
        )
    with open(config_path, "r", encoding="utf-8") as f:
        cfg = json.load(f)
    ruta = (cfg.get("ruta_repositorio_activo") or "").strip()
    if not ruta:
        raise ValueError(
            u"La clave 'ruta_repositorio_activo' esta vacia en config_proyecto_activo.json."
        )
    if not os.path.isfile(ruta):
        raise ValueError(
            u"El repositorio activo no existe en la ruta configurada:\n{}".format(ruta)
        )
    return ruta


def _get_project_name():
    """Retorna el nombre del proyecto Revit activo o cadena vacia."""
    if doc is None:
        return u""
    try:
        pi = doc.ProjectInformation
        return (pi.Name or pi.Number or u"").strip()
    except Exception:
        return u""


def _build_headers(repo):
    """
    Recorre todas las entradas del repositorio para obtener el conjunto
    completo de claves, poniendo las mas importantes al inicio.
    """
    priority = ["CodIntBIM", "nombre_archivo", "Archivo", "ElementId"]
    seen = set(priority)
    all_keys = list(priority)
    for v in repo.values():
        if isinstance(v, dict):
            for k in v.keys():
                if k not in seen:
                    seen.add(k)
                    all_keys.append(k)
    return all_keys


def _generar_datos_tmp(master_dir):
    """
    Lee el repositorio activo del proyecto y genera datos_tmp.json
    en master_dir. Retorna la ruta completa al archivo generado, o None.
    """
    if doc is None:
        forms.alert(u"No hay documento activo de Revit.", title="Error")
        return None

    # Asegurar que master_dir existe
    if not os.path.isdir(master_dir):
        try:
            os.makedirs(master_dir)
        except Exception as e:
            forms.alert(
                u"No se pudo crear la carpeta master:\n{}\n{}".format(master_dir, e),
                title="Error carpeta"
            )
            return None

    # Leer repositorio activo
    try:
        repo_path = _get_repo_path_from_config(master_dir)
    except ValueError as e:
        forms.alert(u"{}".format(e), title="Sin proyecto configurado")
        return None

    try:
        with open(repo_path, "r", encoding="utf-8", errors="replace") as f:
            repo = json.load(f)
    except Exception as e:
        forms.alert(u"Error leyendo repositorio:\n{}".format(e), title="Error repositorio")
        return None

    if not repo:
        forms.alert(
            u"El repositorio esta vacio.\n"
            u"Ejecuta 'Carga de Datos' primero para cargar datos del proyecto.",
            title="Sin datos"
        )
        return None

    # Construir Headers y Rows
    headers = _build_headers(repo)
    rows = []
    for entry in repo.values():
        if not isinstance(entry, dict):
            continue
        rows.append({h: entry.get(h, u"") or u"" for h in headers})

    payload = {
        "Titulo":   u"SAESA \u2014 Datos del Proyecto",
        "Proyecto": _get_project_name(),
        "Headers":  headers,
        "Rows":     rows
    }

    # Escribir JSON en master_dir
    out_path = os.path.join(master_dir, "datos_tmp.json")
    try:
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(payload, f, indent=2, ensure_ascii=False)
    except Exception as e:
        forms.alert(
            u"Error escribiendo datos_tmp.json:\n{}".format(e),
            title="Error escritura"
        )
        return None

    return out_path


# ==============================================================================
# SECCION 2: Lanzar visor CPython
# ==============================================================================

def _ensure_dirs():
    """Crea carpetas de runtime si no existen."""
    for d in (DATA_DIR, MASTER_DIR, TEMP_DIR, CACHE_DIR):
        if not os.path.exists(d):
            try:
                os.makedirs(d)
            except Exception:
                pass


def _run_datos_proyecto(json_path):
    """
    Lanza datos_proyecto.py via CPython.
    argv[1] = ruta completa al datos_tmp.json
    argv[2] = MASTER_DIR
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

    cmd = [PYTHON_EXE, _DATOS_SCRIPT, json_path, MASTER_DIR]
    return subprocess.call(cmd, creationflags=CREATE_NO_WINDOW)


# ==============================================================================
# MAIN
# ==============================================================================

def main():
    _ensure_dirs()

    # Paso 1: generar datos_tmp.json desde el repositorio activo
    json_path = _generar_datos_tmp(MASTER_DIR)
    if not json_path or not os.path.isfile(json_path):
        return

    # Paso 2: abrir visor CPython con el JSON generado
    rc = _run_datos_proyecto(json_path)
    if rc != 0:
        forms.alert(
            u"El gestor de datos termino con codigo: {}".format(rc),
            title=u"Advertencia"
        )


if __name__ == '__main__':
    main()
