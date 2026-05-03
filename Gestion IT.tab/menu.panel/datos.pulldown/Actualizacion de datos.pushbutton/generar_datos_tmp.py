# -*- coding: utf-8 -*-
"""
generar_datos_tmp.py  —  IronPython / Revit API
Lee el repositorio activo del proyecto y genera datos_tmp.json en MASTER_DIR.

Llamado desde script.py con:
    generar_datos_tmp.main(master_dir) -> str  (ruta al JSON generado)

El JSON resultante tiene la estructura que espera datos_proyecto.py:
    {
        "Titulo":   "SAESA — Datos del Proyecto",
        "Proyecto": "<nombre proyecto>",
        "Headers": [...],
        "Rows":    [{...}, ...]
    }
"""
import os
import sys
import json

import clr
from pyrevit import forms

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInParameter

# ── Documento activo ────────────────────────────────────────────────────────
try:
    doc = __revit__.ActiveUIDocument.Document
except Exception:
    doc = None


# ── Helpers ─────────────────────────────────────────────────────────────────

def _get_repo_path(master_dir):
    """
    Lee config_proyecto_activo.json desde master_dir y retorna
    la ruta al JSON de repositorio activo. Lanza ValueError si falla.
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


def _load_repo(repo_path):
    try:
        with open(repo_path, "r", encoding="utf-8", errors="replace") as f:
            return json.load(f)
    except Exception as e:
        forms.alert(u"Error leyendo repositorio:\n{}".format(e), title="Error repositorio")
        return None


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
    seen = set()
    all_keys = list(priority)
    for v in repo.values():
        if isinstance(v, dict):
            for k in v.keys():
                if k not in seen:
                    seen.add(k)
                    if k not in priority:
                        all_keys.append(k)
    return all_keys


# ── Main ─────────────────────────────────────────────────────────────────────

def main(master_dir):
    """
    Genera datos_tmp.json en master_dir y retorna la ruta completa al archivo.
    Devuelve None si ocurre un error.
    """
    if doc is None:
        forms.alert(u"No hay documento activo de Revit.", title="Error")
        return None

    # ── Asegurar que master_dir existe ────────────────────────────────────────
    if not os.path.isdir(master_dir):
        try:
            os.makedirs(master_dir)
        except Exception as e:
            forms.alert(
                u"No se pudo crear la carpeta master:\n{}\n{}".format(master_dir, e),
                title="Error carpeta"
            )
            return None

    # ── Leer repositorio activo ───────────────────────────────────────────────
    try:
        repo_path = _get_repo_path(master_dir)
    except ValueError as e:
        forms.alert(u"{}".format(e), title="Sin proyecto configurado")
        return None

    repo = _load_repo(repo_path)
    if repo is None:
        return None

    if not repo:
        forms.alert(
            u"El repositorio esta vacio.\n"
            u"Ejecuta 'Carga de Datos' primero para cargar datos del proyecto.",
            title="Sin datos"
        )
        return None

    # ── Construir Headers y Rows ──────────────────────────────────────────────
    headers = _build_headers(repo)
    rows = []
    for key, entry in repo.items():
        if not isinstance(entry, dict):
            continue
        row = {}
        for h in headers:
            row[h] = entry.get(h, u"") or u""
        rows.append(row)

    proyecto = _get_project_name()

    payload = {
        "Titulo":   u"SAESA \u2014 Datos del Proyecto",
        "Proyecto": proyecto,
        "Headers":  headers,
        "Rows":     rows
    }

    # ── Escribir JSON en master_dir ───────────────────────────────────────────
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
