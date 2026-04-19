# -*- coding: utf-8 -*-
"""
lib/config/paths.py
===================
Fuente unica de verdad para todas las rutas del proyecto.

Importar desde cualquier script.py con:

    from config.paths import (
        DATA_DIR, MASTER_DIR, TEMP_DIR, CACHE_DIR, LOG_DIR, EXPORT_DIR,
        CONFIG_PROYECTO, REGISTRO_PROYECTOS,
        SCRIPT_JSON_PATH_LIB, SCRIPT_XLSX_PATH,
        ensure_runtime_dirs
    )

Nota: BASE_DIR / EXT_ROOT apuntan a la raiz de la extension
(la carpeta que contiene lib/, data/, scripts_cpython/, etc.)
Se resuelve a partir de la ubicacion de este mismo archivo:

    lib/config/paths.py  ->  lib/config/  ->  lib/  ->  EXT_ROOT
"""

import os


# ── Raiz de la extension ──────────────────────────────────────────────────────
def get_extension_root():
    current_dir = os.path.dirname(os.path.abspath(__file__))   # lib/config/
    return os.path.abspath(os.path.join(current_dir, '..', '..'))  # EXT_ROOT


BASE_DIR = get_extension_root()   # alias historico
EXT_ROOT = BASE_DIR               # alias semantico


# ── Directorios de datos ───────────────────────────────────────────────────────
DATA_DIR   = os.path.join(BASE_DIR, 'data')
MASTER_DIR = os.path.join(DATA_DIR, 'master')
TEMP_DIR   = os.path.join(DATA_DIR, 'temp')
CACHE_DIR  = os.path.join(DATA_DIR, 'cache')
LOG_DIR    = os.path.join(DATA_DIR, 'logs')
EXPORT_DIR = os.path.join(DATA_DIR, 'exports')


# ── Archivos maestros ──────────────────────────────────────────────────────────
COLORES_JSON_PATH    = os.path.join(MASTER_DIR, 'colores_parametros.json')
SCRIPT_JSON_PATH_LIB = os.path.join(MASTER_DIR, 'script.json')
CONFIG_PROYECTO      = os.path.join(MASTER_DIR, 'config_proyecto_activo.json')
REGISTRO_PROYECTOS   = os.path.join(MASTER_DIR, 'registro_proyectos.json')


# ── Scripts CPython ────────────────────────────────────────────────────────────
_CPYTHON_DIR     = os.path.join(BASE_DIR, 'scripts_cpython')
SCRIPT_XLSX_PATH = os.path.join(_CPYTHON_DIR, 'exportar_planillas_xlsx.py')


# ── Utilidades ─────────────────────────────────────────────────────────────────
def ensure_runtime_dirs():
    """Crea todas las carpetas de runtime si no existen.
    Llamar al inicio de cada boton antes de usar rutas de datos."""
    for path in (DATA_DIR, MASTER_DIR, TEMP_DIR, CACHE_DIR, LOG_DIR, EXPORT_DIR):
        if not os.path.exists(path):
            os.makedirs(path)
