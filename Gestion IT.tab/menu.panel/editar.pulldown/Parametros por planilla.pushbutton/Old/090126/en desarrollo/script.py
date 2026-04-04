# -*- coding: utf-8 -*-

__title__ = "Parametros por planilla"

__doc__ = """Version = 1.0
Date = 15.06.2024
________________________________________________________________
Description:
PyRevit pushbutton that launches a planilla selector, builds a
combined dataset (BD + modelo) for a schedule, and opens a Tk viewer.
BD is only updated from the viewer when user saves changes.
________________________________________________________________
Author: Erik Frits"""

import clr
import os
import json
import subprocess
from pyrevit import forms

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    RevitLinkInstance,
    ViewSchedule,
)

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

#--------------------------------------------------
# Rutas comunes

DATA_DIR = os.path.join(
    os.path.expanduser("~"),
    r"AppData\Roaming\MyPyRevitExtention\PyRevitIT.extension\data"
)

CONFIG_PATH       = os.path.join(DATA_DIR, "config_proyecto_activo.json")
SCRIPT_JSON_PATH  = os.path.join(DATA_DIR, "script.json")

BASE_PATH = os.path.dirname(__file__)

PYTHON_EXE        = r"C:\Users\Zbook HP\AppData\Local\Programs\Python\Python313\pythonw.exe"
CPYTHON_SELECTOR  = os.path.join(BASE_PATH, "selector_planillas.pyw")
CPYTHON_VIEWER    = os.path.join(BASE_PATH, "mostrar_tabla_tk.pyw")

# JSONs de intercambio
PLANILLA_META_PATH  = os.path.join(BASE_PATH, "planilla_meta_tmp.json")
PLANILLA_DATA_PATH  = os.path.join(BASE_PATH, "planilla_data_tmp.json")
HEADERS_CACHE_PATH  = os.path.join(BASE_PATH, "planillas_headers_order.json")

#--------------------------------------------------
# Utilidades JSON

def load_json(path, show_error=True, title="Error"):
    if not os.path.exists(path):
        if show_error:
            forms.alert(
                u"No se encontró archivo JSON:\n{}".format(path),
                title=title
            )
        return None
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        if show_error:
            forms.alert(
                u"Error cargando JSON:\n{}\nRuta:\n{}".format(e, path),
                title=title
            )
        return None


def save_json(data, path, show_error=True, title="Error"):
    try:
        folder = os.path.dirname(path)
        if folder and not os.path.exists(folder):
            os.makedirs(folder)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        return True
    except Exception as e:
        if show_error:
            forms.alert(
                u"Error guardando JSON:\n{}\nRuta:\n{}".format(e, path),
                title=title
            )
        return False

#--------------------------------------------------
# Repositorio principal

def get_repo_path_from_config():
    if not os.path.exists(CONFIG_PATH):
        forms.alert(
            u"No se encontró config_proyecto_activo.json en:\n{}\n\n"
            u"No se puede determinar el repositorio de datos activo."
            .format(CONFIG_PATH),
            title="Config no encontrada"
        )
        return None
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            cfg = json.load(f)
        ruta = (cfg.get("ruta_repositorio_activo") or "").strip()
        if not ruta:
            forms.alert(
                u"En config_proyecto_activo.json no se encontró la clave "
                u"'ruta_repositorio_activo' o está vacía.\n\n"
                u"No se puede determinar el repositorio de datos activo.",
                title="Config incompleta"
            )
            return None
        return ruta
    except Exception as e:
        forms.alert(
            u"Error leyendo config_proyecto_activo.json:\n{}\n\n"
            u"No se puede determinar el repositorio de datos activo."
            .format(e),
            title="Error config"
        )
        return None


REPO_PATH = get_repo_path_from_config()


def load_repo():
    if not REPO_PATH:
        return {}
    repo = load_json(REPO_PATH, show_error=False)
    return repo or {}

#--------------------------------------------------
# script.json

def load_script_config():
    data = load_json(SCRIPT_JSON_PATH, show_error=True, title="Error script.json")
    return data or {}


def find_codigo_planilla_from_name(alias_name, config):
    reemplazos = config.get("reemplazos_de_nombres", {}) or {}
    codigos    = config.get("codigos_planillas", {}) or {}
    key = None
    for k, v in reemplazos.items():
        if (v or "").strip() == alias_name.strip():
            key = k
            break
    if not key:
        return None, None
    codigo = codigos.get(key, None)
    return key, codigo

#--------------------------------------------------
# Encabezados desde ViewSchedule + caché local

def get_headers_from_view_schedule(host_doc, nombre_original_planilla):
    try:
        schedules = (
            FilteredElementCollector(host_doc)
            .OfClass(ViewSchedule)
            .WhereElementIsNotElementType()
            .ToElements()
        )
        planilla = next(
            (s for s in schedules if s.Name == nombre_original_planilla),
            None
        )
        if not planilla:
            forms.alert(
                u"No se encontró la planilla '{}' en el modelo."
                .format(nombre_original_planilla),
                title="Planilla no encontrada"
            )
            return []
        headers = []
        for fid in planilla.Definition.GetFieldOrder():
            f = planilla.Definition.GetField(fid)
            if f:
                headers.append(f.GetName())
        if not headers:
            forms.alert(
                u"La planilla '{}' no tiene encabezados de datos."
                .format(nombre_original_planilla),
                title="Encabezados vacíos"
            )
            return []
        return headers
    except Exception as e:
        forms.alert(
            u"Error obteniendo encabezados de la planilla '{}':\n{}"
            .format(nombre_original_planilla, e),
            title="Error encabezados"
        )
        return []


def get_headers_cached(host_doc, nombre_original, codigo_planilla):
    cache = load_json(HEADERS_CACHE_PATH, show_error=False) or {}
    key_cache = "{}::{}".format(nombre_original, codigo_planilla or "")
    if key_cache in cache:
        headers = cache.get(key_cache) or []
    else:
        headers = get_headers_from_view_schedule(host_doc, nombre_original)
        if headers:
            cache[key_cache] = headers
            save_json(cache, HEADERS_CACHE_PATH, show_error=False)
    return headers

#--------------------------------------------------
# Integración con selector Tkinter (CPython)

def run_selector_tk():
    """Lanza selector_planillas.pyw y espera a que escriba planilla_meta_tmp.json."""
    if os.path.exists(PLANILLA_META_PATH):
        try:
            os.remove(PLANILLA_META_PATH)
        except Exception:
            pass

    try:
        subprocess.Popen(
            [PYTHON_EXE, CPYTHON_SELECTOR, PLANILLA_META_PATH],
            shell=False
        ).wait()
    except Exception as e:
        forms.alert(
            u"No se pudo ejecutar selector externo:\n{}".format(e),
            title="Error selector"
        )
        return None

    meta = load_json(PLANILLA_META_PATH, show_error=False) or {}
    if not meta:
        return None
    return meta

#--------------------------------------------------
# Utilidades para combinar modelo + BD (sin guardar en BD)

def make_repo_key(archivo, elem_id_str):
    return u"{}_{}".format(archivo, elem_id_str)


def _norm(val):
    """Normaliza valores vacíos a '-' para la vista."""
    if val in (None, "", " "):
        return "-"
    return val


def get_model_data_for_planilla(codigo_planilla, headers):
    """Recorre links y arma diccionario solo en memoria."""
    if not codigo_planilla or not headers:
        return {}

    result = {}

    link_instances = (
        FilteredElementCollector(doc)
        .OfClass(RevitLinkInstance)
        .WhereElementIsNotElementType()
        .ToElements()
    )

    for link_inst in link_instances:
        linked_doc = link_inst.GetLinkDocument()
        if not linked_doc:
            continue

        elementos = (
            FilteredElementCollector(linked_doc)
            .WhereElementIsNotElementType()
            .ToElements()
        )

        for elem in elementos:
            try:
                p_cod = elem.LookupParameter("CodIntBIM")
                cod_val = p_cod.AsString() if (p_cod and p_cod.HasValue) else ""
                cod_val = (cod_val or "").strip()
                if not cod_val:
                    continue
                if codigo_planilla not in cod_val:
                    continue

                archivo = linked_doc.PathName
                elem_id_str = str(elem.Id.IntegerValue)
                key = make_repo_key(archivo, elem_id_str)

                data = {
                    "Archivo":       archivo,
                    "ElementId":     elem_id_str,
                    "nombre_archivo": os.path.basename(archivo) if archivo else "",
                    "CodIntBIM":     cod_val,
                }

                for p in elem.Parameters:
                    try:
                        pname = p.Definition.Name
                        if pname not in headers:
                            continue
                        if p.StorageType.ToString() == "String":
                            pval = p.AsString()
                        else:
                            pval = p.AsValueString()
                        data[pname] = _norm(pval)
                    except Exception:
                        continue

                # Normalizar también metadatos
                data["Archivo"]       = _norm(data.get("Archivo"))
                data["ElementId"]     = _norm(data.get("ElementId"))
                data["nombre_archivo"] = _norm(data.get("nombre_archivo"))
                data["CodIntBIM"]     = _norm(data.get("CodIntBIM"))

                result[key] = data
            except Exception:
                continue

    return result


def build_combined_dataset(codigo_planilla, headers):
    """Combina BD + modelo solo en memoria. No guarda en BD."""
    repo = load_repo()
    combined = {}

    # 1) Datos del modelo
    model_data = get_model_data_for_planilla(codigo_planilla, headers)
    for key, md in model_data.items():
        combined[key] = dict(md)

    # 2) Incorporar datos existentes en BD
    for key, registro in repo.items():
        try:
            codint = (registro.get("CodIntBIM", "") or "").strip()
            if codigo_planilla and codigo_planilla not in codint:
                continue

            if key in combined:
                base = combined[key]
                # Fusionar: prioridad a BD donde haya datos
                for h in headers:
                    if h in ("Archivo", "ElementId", "nombre_archivo", "CodIntBIM"):
                        continue
                    val_bd = registro.get(h, "")
                    if val_bd not in (None, "", " "):
                        base[h] = _norm(val_bd)
                # Mantener también CodIntBIM y metadatos (normalizados)
                for extra in ("CodIntBIM", "Archivo", "ElementId", "nombre_archivo"):
                    if extra in registro:
                        base[extra] = _norm(registro.get(extra))
                combined[key] = base
            else:
                # Registro solo en BD pero perteneciente a la planilla
                archivo = registro.get("Archivo", "") or ""
                elem_id = registro.get("ElementId", "") or ""
                if not archivo or not elem_id:
                    continue

                fila = {}
                fila["Archivo"]        = _norm(archivo)
                fila["ElementId"]      = _norm(elem_id)
                fila["nombre_archivo"] = _norm(
                    registro.get(
                        "nombre_archivo",
                        os.path.basename(archivo) if archivo else ""
                    )
                )
                fila["CodIntBIM"]      = _norm(codint)

                for h in headers:
                    if h in ("Archivo", "ElementId", "nombre_archivo", "CodIntBIM"):
                        continue
                    fila[h] = _norm(registro.get(h, ""))

                combined[key] = fila
        except Exception:
            continue

    rows = list(combined.values())
    rows.sort(key=lambda r: str(r.get("CodIntBIM", "") or "").lower())
    return rows

#--------------------------------------------------
# Lanzar visor Tkinter (editor de planilla)

def run_viewer_tk(planilla_meta):
    save_json(planilla_meta, PLANILLA_META_PATH, show_error=False)
    try:
        subprocess.Popen(
            [PYTHON_EXE, CPYTHON_VIEWER, PLANILLA_META_PATH],
            shell=False
        )
    except Exception as e:
        forms.alert(
            u"No se pudo ejecutar visor externo:\n{}".format(e),
            title="Error visor"
        )

#--------------------------------------------------
# main

def main():
    if not REPO_PATH:
        return

    script_cfg = load_script_config()
    if not script_cfg:
        return

    meta_sel = run_selector_tk()
    if not meta_sel:
        return

    nombre_original = meta_sel.get("NombrePlanillaOriginal", "")
    nombre_alias    = meta_sel.get("NombrePlanillaAlias", "")
    codigo_planilla = meta_sel.get("CodigoPlanilla", "")

    if not nombre_original or not codigo_planilla:
        forms.alert(
            u"El selector no devolvió NombrePlanillaOriginal o CodigoPlanilla.",
            title="Meta incompleta"
        )
        return

    headers = get_headers_cached(doc, nombre_original, codigo_planilla)
    if not headers:
        return

    filas = build_combined_dataset(codigo_planilla, headers)
    save_json(filas, PLANILLA_DATA_PATH, show_error=False)

    planilla_meta = {
        "Headers":        headers,
        "CodigoPlanilla": codigo_planilla,
        "NombrePlanilla": nombre_alias or nombre_original,
        "DataPath":       PLANILLA_DATA_PATH,
    }

    run_viewer_tk(planilla_meta)


if __name__ == "__main__":
    main()

#==================================================
from Snippets._customprint import kit_button_clicked
kit_button_clicked(btn_name=__title__)
