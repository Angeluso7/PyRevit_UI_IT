# -*- coding: utf-8 -*-
import os


def get_extension_root():
    current_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.abspath(os.path.join(current_dir, '..', '..'))


BASE_DIR    = get_extension_root()
DATA_DIR    = os.path.join(BASE_DIR, 'data')
MASTER_DIR  = os.path.join(DATA_DIR, 'master')
CACHE_DIR   = os.path.join(DATA_DIR, 'cache')
TEMP_DIR    = os.path.join(DATA_DIR, 'temp')
LOG_DIR     = os.path.join(DATA_DIR, 'logs')
EXPORT_DIR  = os.path.join(DATA_DIR, 'exports')

# Archivos maestros
COLORES_JSON_PATH    = os.path.join(MASTER_DIR, 'colores_parametros.json')
SCRIPT_JSON_PATH_LIB = os.path.join(MASTER_DIR, 'script.json')
CONFIG_PROYECTO      = os.path.join(MASTER_DIR, 'config_proyecto_activo.json')
REGISTRO_PROYECTOS   = os.path.join(MASTER_DIR, 'registro_proyectos.json')

# Script CPython para exportacion XLSX
SCRIPT_XLSX_PATH = os.path.join(BASE_DIR, 'scripts_cpython', 'exportar_planillas_xlsx.py')


def ensure_runtime_dirs():
    for path in [DATA_DIR, MASTER_DIR, CACHE_DIR, TEMP_DIR, LOG_DIR, EXPORT_DIR]:
        if not os.path.exists(path):
            os.makedirs(path)
