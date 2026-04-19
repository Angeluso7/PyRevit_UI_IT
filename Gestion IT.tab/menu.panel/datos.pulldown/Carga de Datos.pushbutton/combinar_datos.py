# -*- coding: utf-8 -*-
__title__   = "Combinar Datos"
__doc__     = """Version = 1.4
Date    = 18.04.2026
________________________________________________________________
Description:

Combina los datos del modelo Revit con el repositorio activo.
________________________________________________________________
Author: Argenis Angel"""

# ╬══╬═╦══╬═╬══╬══╬═╦═╦══
# ║║║╠═╣║╚╗ ║║ ═╝
# ╚╚ ╚╚  ╚═╝╚═╝ ╚ ╚═╝
#==================================================
import os
import sys
import json

# ╬  ╬╔═╗╬═╗╬╔╗ ╬  ╔═╗╔═╗
# ╚╗╔╝╠═╣╠╦╝╠╣║╬  ║ ║╕ ║
#  ╚╝ ╚ ╚╚╚══╚╚ ╚╚═╝═╝╚═╝╚═╝
#==================================================

# ── Asegurar lib/ en sys.path y cargar rutas centralizadas ──────────────────
try:
    _this_dir = os.path.dirname(os.path.abspath(__file__))
except Exception:
    _this_dir = os.getcwd()

_EXT_ROOT = os.path.abspath(os.path.join(_this_dir, '..', '..', '..', '..'))
_LIB_DIR  = os.path.join(_EXT_ROOT, 'lib')
if _LIB_DIR not in sys.path:
    sys.path.insert(0, _LIB_DIR)

try:
    from config.paths import MASTER_DIR, CONFIG_PROYECTO
except Exception:
    _DATA_DIR       = os.path.join(_EXT_ROOT, 'data')
    MASTER_DIR      = os.path.join(_DATA_DIR, 'master')
    CONFIG_PROYECTO = os.path.join(MASTER_DIR, 'config_proyecto_activo.json')

_CONFIG_CANDIDATES = [
    CONFIG_PROYECTO,
    os.path.join(os.path.dirname(MASTER_DIR), 'config_proyecto_activo.json'),
]

# ── Utilidades JSON ───────────────────────────────────────────────────────────────

def load_json(path):
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)


def save_json(path, data):
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


# ── Config proyecto activo ──────────────────────────────────────────────────────────

def get_config_path():
    for p in _CONFIG_CANDIDATES:
        if os.path.isfile(p):
            return p
    return None


def load_config():
    cfg_path = get_config_path()
    if not cfg_path:
        raise IOError(
            u"No se encontro config_proyecto_activo.json.\n"
            u"Rutas buscadas:\n" + u"\n".join(_CONFIG_CANDIDATES)
        )
    return load_json(cfg_path)


# ── Logica principal ─────────────────────────────────────────────────────────────────────

def combinar_datos(modelo_data, config):
    """
    Combina los datos extraidos del modelo con la configuracion
    del proyecto activo.
    """
    ruta_repo = config.get('ruta_repositorio_activo', '')
    if not ruta_repo or not os.path.isfile(ruta_repo):
        return modelo_data  # sin repo activo, devuelve modelo sin combinar

    repo_data = load_json(ruta_repo)
    if not isinstance(repo_data, dict):
        return modelo_data

    combined = dict(modelo_data)
    for clave, valores_repo in repo_data.items():
        if clave in combined:
            entry = combined[clave]
            if isinstance(entry, dict) and isinstance(valores_repo, dict):
                entry.update(valores_repo)
        else:
            combined[clave] = valores_repo

    return combined


def main():
    from pyrevit import forms
    try:
        cfg = load_config()
    except IOError as e:
        forms.alert(str(e), title=u"Error de configuracion")
        return

    ruta_repo = cfg.get('ruta_repositorio_activo', '').strip()
    if not ruta_repo:
        forms.alert(
            u"El proyecto activo no tiene 'ruta_repositorio_activo' definida.",
            title=u"Sin repositorio"
        )
        return

    if not os.path.isfile(ruta_repo):
        forms.alert(
            u"No se encontro el repositorio activo:\n{}".format(ruta_repo),
            title=u"Archivo no encontrado"
        )
        return

    forms.alert(
        u"Repositorio activo cargado correctamente:\n{}".format(ruta_repo),
        title=u"OK"
    )


if __name__ == '__main__':
    main()
