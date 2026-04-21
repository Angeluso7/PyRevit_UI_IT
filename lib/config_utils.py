# -*- coding: utf-8 -*-
"""
lib/config_utils.py
===================
Módulo utilitario centralizado para resolver rutas críticas del proyecto.
Importado tanto desde IronPython (scripts Revit) como desde CPython (visores).

USO:
    import sys, os
    sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'lib'))
    from config_utils import get_config_path, get_repo_activo_path, EXT_ROOT, DATA_DIR, MASTER_DIR
"""

import os
import json

# ── Raíz de la extensión ────────────────────────────────────────────────────
EXT_ROOT = os.path.normpath(os.path.join(
    os.path.expanduser("~"),
    "AppData", "Roaming", "MyPyRevitExtention", "PyRevitIT.extension"
))

DATA_DIR   = os.path.join(EXT_ROOT, "data")
MASTER_DIR = os.path.join(DATA_DIR, "master")
CACHE_DIR  = os.path.join(DATA_DIR, "cache")
TEMP_DIR   = os.path.join(DATA_DIR, "temp")
LOGS_DIR   = os.path.join(DATA_DIR, "logs")

# Candidatos para config_proyecto_activo.json (master tiene prioridad)
_CONFIG_CANDIDATES = [
    os.path.join(MASTER_DIR, "config_proyecto_activo.json"),
    os.path.join(DATA_DIR,   "config_proyecto_activo.json"),
]

# ── Resolución de config ────────────────────────────────────────────────────

def get_config_path():
    """
    Devuelve la ruta al config_proyecto_activo.json que exista en disco.
    Orden de búsqueda: data/master/ → data/ (fallback legacy).
    Lanza FileNotFoundError si no existe en ninguna ubicación.
    """
    for candidate in _CONFIG_CANDIDATES:
        if os.path.isfile(candidate):
            return candidate
    raise IOError(
        u"No se encontró config_proyecto_activo.json.\n"
        u"Ubicaciones revisadas:\n  " + u"\n  ".join(_CONFIG_CANDIDATES) + u"\n\n"
        u"Ejecuta primero el botón 'Configurar Proyecto' para crear este archivo."
    )

def load_config():
    """Carga y devuelve el dict de config_proyecto_activo.json."""
    path = get_config_path()
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)

def get_repo_activo_path():
    """
    Lee config_proyecto_activo.json y devuelve la ruta del repositorio activo.
    Lanza ValueError si la clave no existe o está vacía.
    """
    cfg = load_config()
    ruta = (cfg.get("ruta_repositorio_activo") or "").strip()
    if not ruta:
        raise ValueError(
            u"La clave 'ruta_repositorio_activo' no existe o está vacía en config_proyecto_activo.json."
        )
    return ruta

def get_script_json_path():
    """
    Devuelve la ruta de script.json ubicado en data/master/.
    Candidatos en orden: data/master/script.json → data/script.json (legacy).
    """
    candidates = [
        os.path.join(MASTER_DIR, "script.json"),
        os.path.join(DATA_DIR,   "script.json"),
    ]
    for candidate in candidates:
        if os.path.isfile(candidate):
            return candidate
    # Si no existe ninguno, devuelve la ruta canónica (master) para que
    # el llamador muestre el error correcto con la ruta esperada.
    return candidates[0]

# ── Compatibilidad CPython / IronPython ─────────────────────────────────────
# Ambos entornos usan os.path.expanduser("~") correctamente; no se necesita
# ajuste especial. En CPython puro se puede importar este módulo directamente.
