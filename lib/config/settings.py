# -*- coding: utf-8 -*-
"""
settings.py
-----------
Resolución de CPYTHON_EXE en 3 niveles (orden de prioridad):

  1. Variable de entorno:  PYREVIT_IT_CPYTHON
  2. Caché local:          data/temp/env_cache.json  {"python_exe": "..."}
  3. Fallback genérico:   C:\\Python313\\python.exe

Para configurar la ruta en la máquina local sin tocar este archivo:
  - Opción A) setx PYREVIT_IT_CPYTHON "C:\\...\\python.exe" /M
  - Opción B) Editar data/temp/env_cache.json con la ruta correcta
"""
import os
import json
from config.paths import MASTER_DIR, TEMP_DIR, SCRIPT_XLSX_PATH


def _resolver_cpython_exe():
    # 1° Variable de entorno
    desde_env = os.environ.get('PYREVIT_IT_CPYTHON', '').strip()
    if desde_env and os.path.isfile(desde_env):
        return desde_env

    # 2° Cache local: data/temp/env_cache.json
    cache_path = os.path.join(TEMP_DIR, 'env_cache.json')
    if os.path.isfile(cache_path):
        try:
            with open(cache_path, 'r', encoding='utf-8') as f:
                cache = json.load(f)
            ruta = cache.get('python_exe', '').strip()
            if ruta and os.path.isfile(ruta):
                return ruta
        except Exception:
            pass

    # 3° Fallback genérico
    return r'C:\Python313\python.exe'


CPYTHON_EXE         = _resolver_cpython_exe()
ENABLE_CPYTHON_XLSX = True
CPYTHON_TIMEOUT     = 120

SCRIPT_JSON_PATH = os.environ.get(
    'PYREVIT_IT_SCRIPT_JSON',
    os.path.join(MASTER_DIR, 'script.json')
)
