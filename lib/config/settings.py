# -*- coding: utf-8 -*-
import os

CPYTHON_EXE = os.environ.get('PYREVIT_IT_CPYTHON', r'C:\Users\Zbook HP\AppData\Local\Programs\Python\Python313\python.exe')
ENABLE_CPYTHON_XLSX = True
CPYTHON_TIMEOUT = 120

SCRIPT_JSON_PATH = os.environ.get('PYREVIT_IT_SCRIPT_JSON', os.path.join(os.path.expanduser('~'), r'AppData\Roaming\MyPyRevitExtention\PyRevitIT.extension\data\script.json'))
