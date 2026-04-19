# -*- coding: utf-8 -*-
"""
combinar_datos.py  (IronPython — carga via imp desde Carga de Datos.pushbutton)
=================================================================================
Combina el repositorio temporal (TXT generado por carga_excel.py) con el
repositorio activo del proyecto definido en config_proyecto_activo.json.

Rutas de búsqueda de config_proyecto_activo.json (en orden de prioridad):
  1. data/master/config_proyecto_activo.json   ← nueva estructura
  2. data/config_proyecto_activo.json          ← fallback legacy
"""

import os
import sys
import json

# ── Rutas base ─────────────────────────────────────────────────────────────
_EXT_ROOT = os.path.join(
    os.path.expanduser("~"),
    "AppData", "Roaming", "MyPyRevitExtention", "PyRevitIT.extension"
)
_DATA_DIR    = os.path.join(_EXT_ROOT, "data")
_MASTER_DIR  = os.path.join(_DATA_DIR, "master")

# Rutas candidatas para config_proyecto_activo.json (prioridad: master > raíz)
_CONFIG_CANDIDATES = [
    os.path.join(_MASTER_DIR, "config_proyecto_activo.json"),
    os.path.join(_DATA_DIR,   "config_proyecto_activo.json"),
]

# ── Utilidades JSON ─────────────────────────────────────────────────────────

def _load_json(path):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)

def _save_json(data, path):
    folder = os.path.dirname(path)
    if folder and not os.path.isdir(folder):
        os.makedirs(folder)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

# ── Resolución de config_proyecto_activo.json ───────────────────────────────

def _find_config_path():
    """
    Busca config_proyecto_activo.json en la lista de candidatos.
    Devuelve la primera ruta existente o None.
    """
    for candidate in _CONFIG_CANDIDATES:
        if os.path.isfile(candidate):
            return candidate
    return None

def _get_repo_activo_path():
    """
    Lee config_proyecto_activo.json (donde sea que esté) y devuelve
    la ruta del repositorio activo o None.
    """
    cfg_path = _find_config_path()
    if not cfg_path:
        raise IOError(
            u"No se encontró config_proyecto_activo.json en ninguna ubicación esperada.\n"
            u"Ubicaciones revisadas:\n  " + u"\n  ".join(_CONFIG_CANDIDATES)
        )
    cfg = _load_json(cfg_path)
    ruta = (cfg.get("ruta_repositorio_activo") or "").strip()
    if not ruta:
        raise ValueError(
            u"En '{}' la clave 'ruta_repositorio_activo' está vacía o no existe.".format(cfg_path)
        )
    return ruta

# ── Lógica de combinación ───────────────────────────────────────────────────

def _leer_tmp_txt(txt_path):
    """
    Lee el TXT temporal generado por carga_excel.py.
    Formato: una línea JSON por registro.
    Devuelve lista de dicts.
    """
    registros = []
    with open(txt_path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line:
                try:
                    registros.append(json.loads(line))
                except ValueError:
                    pass  # línea malformada — ignorar
    return registros

def _combinar(repo_activo, nuevos_registros):
    """
    Fusiona nuevos_registros en repo_activo (dict keyed por CodIntBIM o índice).
    Los registros nuevos sobreescriben los existentes con la misma clave.
    """
    for reg in nuevos_registros:
        clave = reg.get("CodIntBIM") or reg.get("CODIGO") or reg.get("codigo")
        if clave:
            repo_activo[clave] = reg
        else:
            # Sin clave identificable: agregar como nuevo con índice único
            idx = len(repo_activo)
            while str(idx) in repo_activo:
                idx += 1
            repo_activo[str(idx)] = reg
    return repo_activo

# ── Punto de entrada ────────────────────────────────────────────────────────

def main(tmp_txt_path):
    """
    Entrada principal llamada desde script.py via imp.load_source().
    tmp_txt_path: ruta al TXT temporal generado por carga_excel.py.
    """
    from System.Windows.Forms import MessageBox  # IronPython

    # 1. Validar TXT temporal
    if not os.path.isfile(tmp_txt_path):
        MessageBox.Show(
            u"No se encontró el archivo temporal de datos:\n{}".format(tmp_txt_path),
            u"Error combinación"
        )
        return

    # 2. Resolver ruta del repositorio activo
    try:
        repo_path = _get_repo_activo_path()
    except (IOError, ValueError) as e:
        MessageBox.Show(unicode(e), u"Config no encontrada")
        return

    # 3. Cargar repositorio existente (puede no existir aún)
    if os.path.isfile(repo_path):
        try:
            repo_activo = _load_json(repo_path)
            if not isinstance(repo_activo, dict):
                repo_activo = {}
        except Exception as e:
            MessageBox.Show(
                u"Error leyendo repositorio activo:\n{}\n\nRuta:\n{}".format(e, repo_path),
                u"Error lectura"
            )
            return
    else:
        repo_activo = {}

    # 4. Leer nuevos registros del TXT temporal
    try:
        nuevos = _leer_tmp_txt(tmp_txt_path)
    except Exception as e:
        MessageBox.Show(
            u"Error leyendo datos temporales:\n{}\n\nRuta:\n{}".format(e, tmp_txt_path),
            u"Error lectura temporal"
        )
        return

    if not nuevos:
        MessageBox.Show(
            u"El archivo temporal no contiene registros válidos.\n\nRuta:\n{}".format(tmp_txt_path),
            u"Sin datos"
        )
        return

    # 5. Combinar y guardar
    try:
        repo_combinado = _combinar(repo_activo, nuevos)
        _save_json(repo_combinado, repo_path)
    except Exception as e:
        MessageBox.Show(
            u"Error guardando repositorio combinado:\n{}\n\nRuta:\n{}".format(e, repo_path),
            u"Error guardado"
        )
        return

    MessageBox.Show(
        u"Carga completada.\n\n"
        u"Registros nuevos/actualizados: {}\n"
        u"Total en repositorio: {}\n\n"
        u"Repositorio:\n{}".format(len(nuevos), len(repo_combinado), repo_path),
        u"Carga de datos"
    )
