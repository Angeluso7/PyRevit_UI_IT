# -*- coding: utf-8 -*-
"""
lib/core/env_config.py  (IronPython + CPython compatible)
==========================================================
Modulo compartido para resolver rutas de entorno de forma
portatil entre PCs.  Todos los botones que lancen sub-procesos
CPython importan:

    from core.env_config import get_python_exe

Logica de get_python_exe()
--------------------------
1. Lee la cache  data/temp/env_cache.json
2. Si la ruta guardada apunta a un exe valido -> la devuelve (rapido)
3. Si no existe o el exe fue movido/desinstalado:
   a) Busca en AppData\Local\Programs\Python\Python3*\
   b) Usa el comando 'where python' (Windows)
   c) Recorre el PATH del sistema
4. Guarda el resultado en la cache y lo devuelve
"""

import os
import glob
import json
import subprocess

# TEMP_DIR viene de config.paths para no hardcodear la ruta de usuario.
# Fallback manual solo si se ejecuta fuera del contexto de la extension
# (p.ej. pruebas directas en CPython sin lib/ en sys.path).
try:
    from config.paths import TEMP_DIR as _TEMP_DIR
except ImportError:
    # Fallback: resolver desde la ubicacion de este archivo
    # lib/core/env_config.py -> lib/core/ -> lib/ -> EXT_ROOT -> data/temp
    _here     = os.path.dirname(os.path.abspath(__file__))
    _ext_root = os.path.abspath(os.path.join(_here, '..', '..'))
    _TEMP_DIR = os.path.join(_ext_root, 'data', 'temp')


def _cache_path():
    return os.path.join(_TEMP_DIR, "env_cache.json")


def _read_cache():
    p = _cache_path()
    if not os.path.isfile(p):
        return {}
    try:
        with open(p, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def _write_cache(data):
    p = _cache_path()
    folder = os.path.dirname(p)
    try:
        if not os.path.isdir(folder):
            os.makedirs(folder)
        with open(p, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
    except Exception:
        pass  # fallo de escritura no es fatal


def _search_python_system():
    """
    Busca python.exe en cascada (sin dependencias externas).
    Devuelve la primera ruta valida encontrada, o None.
    """
    base = os.path.join(
        os.path.expanduser("~"), "AppData", "Local", "Programs", "Python"
    )

    # 1. Carpetas Python3xx en AppData del usuario
    if os.path.isdir(base):
        for exe_name in ("python.exe", "pythonw.exe"):
            pattern = os.path.join(base, "Python3*", exe_name)
            candidates = sorted(glob.glob(pattern), reverse=True)
            for c in candidates:
                if os.path.isfile(c):
                    return c

    # 2. Comando 'where' de Windows
    for exe_name in ("python", "pythonw"):
        try:
            out = subprocess.check_output(
                ["where", exe_name],
                stderr=subprocess.DEVNULL,
                shell=True
            ).decode("utf-8", errors="ignore").strip()
            for line in out.splitlines():
                line = line.strip()
                if line and os.path.isfile(line):
                    return line
        except Exception:
            pass

    # 3. Recorrido manual del PATH
    for exe_name in ("python.exe", "pythonw.exe"):
        for folder in os.environ.get("PATH", "").split(os.pathsep):
            candidate = os.path.join(folder.strip(), exe_name)
            if os.path.isfile(candidate):
                return candidate

    return None


def get_python_exe():
    """
    Devuelve la ruta a python.exe portatil entre PCs.
    - Si la cache es valida, la devuelve sin buscar (rapido).
    - Si no hay cache o la ruta no existe, busca y actualiza cache.
    - Devuelve None si Python no esta instalado.
    """
    cache  = _read_cache()
    cached = cache.get("python_exe", "")
    if cached and os.path.isfile(cached):
        return cached

    exe = _search_python_system()
    cache["python_exe"] = exe or ""
    _write_cache(cache)
    return exe or None
