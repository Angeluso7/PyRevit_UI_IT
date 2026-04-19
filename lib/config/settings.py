# -*- coding: utf-8 -*-
"""
settings.py
-----------
Rutas sensibles resueltas desde variables de entorno primero,
con fallback generico. Nunca se hardcodea el nombre de usuario.

Configura antes de usar:
    set PYREVIT_IT_CPYTHON=C:\Python313\python.exe
    set PYREVIT_IT_SCRIPT_JSON=C:\...\data\master\script.json
"""
import os
from config.paths import MASTER_DIR, SCRIPT_XLSX_PATH

CPYTHON_EXE = os.environ.get(
    'PYREVIT_IT_CPYTHON',
    r'C:\Python313\python.exe'
)
ENABLE_CPYTHON_XLSX = True
CPYTHON_TIMEOUT     = 120

SCRIPT_JSON_PATH = os.environ.get(
    'PYREVIT_IT_SCRIPT_JSON',
    os.path.join(MASTER_DIR, 'script.json')
)
