# -*- coding: utf-8 -*-
__title__   = "Parametros por planilla"
__doc__     = """Version = 1.2
Date    = 19.04.2026
________________________________________________________________
Description:

Editor de parámetros por planilla sobre modelo linkeado,
usando planillas y un repositorio JSON.

________________________________________________________________
Last Updates:
- [19.04.2026] v1.2  load_script_config tolerante; codigo zombie comentado;
                     archivos runtime movidos a data/temp y data/cache.
- [19.04.2026] v1.1  Rutas centralizadas via config.paths.
________________________________________________________________
Author: Erik Frits + ajustes Angeluso"""

# ╬╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝
#==================================================

import clr
import os
import sys
import json
import re
import subprocess

from pyrevit import forms

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    RevitLinkInstance,
    ViewSchedule,
)

# ╬  ╬╔═╗╦═╗╦╔═╗╔╗ ╬  ╔═╗╔═╗
# ╚╗╔╝╠═╣╠╦╝║╠═╣╠╩╗║  ║╣ ╚═╗
#  ╚╝ ╩ ╩╩╚═╩╩ ╩╚═╝╩═╝╚═╝╚═╝
#==================================================
uidoc = __revit__.ActiveUIDocument
doc   = uidoc.Document

# ╔╦╗╔═╗╦╔╗╔
# ║║║╠═╣║║║║
# ╩ ╩╩ ╩╩╝╚╝
#==================================================

# ── Rutas centralizadas desde config.paths ────────────────────────────────────
try:
    _this_dir = os.path.dirname(os.path.abspath(__file__))
except Exception:
    _this_dir = os.getcwd()

# pushbutton(1) -> pulldown(2) -> panel(3) -> tab(4) -> EXT_ROOT
_EXT_ROOT = os.path.abspath(os.path.join(_this_dir, '..', '..', '..', '..'))
_LIB_DIR  = os.path.join(_EXT_ROOT, 'lib')
if _LIB_DIR not in sys.path:
    sys.path.insert(0, _LIB_DIR)

try:
    from config.paths import DATA_DIR, MASTER_DIR, TEMP_DIR, CACHE_DIR, \
                             CONFIG_PROYECTO, REGISTRO_PROYECTOS, \
                             SCRIPT_JSON_PATH_LIB, ensure_runtime_dirs
    from core.env_config import get_python_exe
    ensure_runtime_dirs()
    PYTHON_EXE = get_python_exe()
except Exception as _path_err:
    _DATA_DIR            = os.path.join(_EXT_ROOT, 'data')
    DATA_DIR             = _DATA_DIR
    MASTER_DIR           = os.path.join(_DATA_DIR, 'master')
    TEMP_DIR             = os.path.join(_DATA_DIR, 'temp')
    CACHE_DIR            = os.path.join(_DATA_DIR, 'cache')
    CONFIG_PROYECTO      = os.path.join(MASTER_DIR, 'config_proyecto_activo.json')
    REGISTRO_PROYECTOS   = os.path.join(MASTER_DIR, 'registro_proyectos.json')
    SCRIPT_JSON_PATH_LIB = os.path.join(MASTER_DIR, 'script.json')
    import glob as _glob
    def _fb_python():
        base = os.path.join(os.path.expanduser('~'), 'AppData', 'Local', 'Programs', 'Python')
        for exe in ('python.exe', 'pythonw.exe'):
            for cand in sorted(_glob.glob(os.path.join(base, 'Python3*', exe)), reverse=True):
                return cand
        for folder in os.environ.get('PATH', '').split(os.pathsep):
            cand = os.path.join(folder.strip(), 'python.exe')
            if os.path.isfile(cand):
                return cand
        return None
    PYTHON_EXE = _fb_python()

# Alias locales
CONFIG_PATH      = CONFIG_PROYECTO
SCRIPT_JSON_PATH = SCRIPT_JSON_PATH_LIB

# Scripts CPython — en la carpeta del botón
BASE_PATH        = os.path.dirname(os.path.abspath(__file__))
CPYTHON_SELECTOR = os.path.join(BASE_PATH, "selector_planillas.pyw")
CPYTHON_VIEWER   = os.path.join(BASE_PATH, "mostrar_tabla_tk.pyw")

# Temporales y cache — en data/temp y data/cache
PLANILLA_META_PATH = os.path.join(TEMP_DIR,  "planilla_meta_tmp.json")
PLANILLA_DATA_PATH = os.path.join(TEMP_DIR,  "planilla_data_tmp.json")
HEADERS_CACHE_PATH = os.path.join(CACHE_DIR, "planillas_headers_order.json")
CACHE_MODELO_PATH  = os.path.join(CACHE_DIR, "cache_modelo_por_clave.json")

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
# script.json  — tolerante: si no existe, continua con config vacia

def load_script_config():
    """
    Carga la configuracion desde data/master/script.json.
    Si el archivo no existe o tiene error, devuelve {} y continua
    sin interrumpir el flujo del boton.
    """
    if not os.path.exists(SCRIPT_JSON_PATH):
        # No es un error critico: el boton puede funcionar sin script.json
        return {}
    data = load_json(SCRIPT_JSON_PATH, show_error=True, title="Error script.json")
    return data or {}

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
# Selector externo (Tk)

def run_selector_tk():
    """Lanza selector_planillas.pyw y espera meta en PLANILLA_META_PATH."""
    if os.path.exists(PLANILLA_META_PATH):
        try:
            os.remove(PLANILLA_META_PATH)
        except Exception:
            pass

    if not PYTHON_EXE or not os.path.isfile(PYTHON_EXE):
        forms.alert(
            u"No se encontró Python 3 en este equipo.\n"
            u"Verifica la instalación o configura PYREVIT_PYTHON_EXE.",
            title="Python no encontrado"
        )
        return None

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
# Dataset modelo + BD

def make_repo_key(archivo, elem_id_str):
    return u"{}_{}".format(archivo, elem_id_str)


def _norm(val):
    """Normaliza valores vacíos a '-' para la vista."""
    if val in (None, "", " "):
        return "-"
    return val


def get_filtered_rows_from_model(headers, codigo_planilla):
    if not codigo_planilla:
        return {}, {}, {}

    repo_datos = load_repo()
    if not isinstance(repo_datos, dict):
        repo_datos = {}

    cache_modelo = load_json(CACHE_MODELO_PATH, show_error=False) or {}
    if not isinstance(cache_modelo, dict):
        cache_modelo = {}

    result            = {}
    cods_por_clave    = {}
    valores_por_clave = {}

    filtrados             = []
    element_ids_filtrados = set()

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
                if not p_cod:
                    continue
                cod_str = p_cod.AsString()
                if not cod_str:
                    continue
                cod_str = cod_str.strip()
                if len(cod_str) < 4:
                    continue
                if cod_str[:4] != codigo_planilla:
                    continue

                codint_modelo = cod_str
                elem_id_str   = str(elem.Id.IntegerValue)

                archivo_procesado = linked_doc.PathName
                base, ext = os.path.splitext(archivo_procesado)
                archivo_procesado = re.sub(r"_\d+$", "", base) + ext

                clave_repo = make_repo_key(archivo_procesado, elem_id_str)
                filtrados.append((clave_repo, elem_id_str, codint_modelo, elem, linked_doc))
                element_ids_filtrados.add(elem_id_str)

            except Exception:
                continue

    bd_por_elementid = {}
    if element_ids_filtrados:
        for _, datos_bd in repo_datos.items():
            try:
                if not isinstance(datos_bd, dict):
                    continue
                elem_id_bd = str(datos_bd.get("ElementId", "") or "")
                if elem_id_bd and elem_id_bd in element_ids_filtrados and len(datos_bd.keys()) > 2:
                    bd_por_elementid[elem_id_bd] = datos_bd
            except Exception:
                continue

    for clave_repo, elem_id_str, codint_modelo, elem, linked_doc in filtrados:
        try:
            archivo_procesado = linked_doc.PathName
            base, ext = os.path.splitext(archivo_procesado)
            archivo_procesado = re.sub(r"_\d+$", "", base) + ext

            datos_modelo = {
                "ElementId":      elem_id_str,
                "Archivo":        archivo_procesado,
                "nombre_archivo": os.path.basename(archivo_procesado),
                "CodIntBIM":      codint_modelo,
            }

            for p in elem.Parameters:
                try:
                    pname = p.Definition.Name
                    if pname not in headers:
                        continue
                    if p.StorageType.ToString() == "String":
                        pval = p.AsString() or ""
                    else:
                        pval = p.AsValueString() or ""
                    datos_modelo[pname] = pval
                except Exception:
                    continue

            cache_modelo[clave_repo] = datos_modelo

            datos_bd = bd_por_elementid.get(elem_id_str)

            if datos_bd and isinstance(datos_bd, dict):
                datos = dict(datos_bd)
                datos["ElementId"]      = elem_id_str
                datos["Archivo"]        = archivo_procesado
                datos["nombre_archivo"] = os.path.basename(archivo_procesado)
                cod_bd = datos.get("CodIntBIM", "")
                if cod_bd in (None, "", " "):
                    datos["CodIntBIM"] = codint_modelo
            else:
                datos_cache = cache_modelo.get(clave_repo)
                datos = dict(datos_cache) if datos_cache and isinstance(datos_cache, dict) else datos_modelo

            cod_oficial               = datos.get("CodIntBIM", "") or codint_modelo
            cods_por_clave[clave_repo] = cod_oficial

            fila = {
                "Archivo":        _norm(datos.get("Archivo")),
                "ElementId":      _norm(datos.get("ElementId")),
                "nombre_archivo": _norm(datos.get("nombre_archivo")),
                "CodIntBIM":      _norm(cod_oficial),
            }

            valores_header = {
                "Archivo":        datos.get("Archivo",        archivo_procesado),
                "ElementId":      datos.get("ElementId",      elem_id_str),
                "nombre_archivo": datos.get("nombre_archivo", os.path.basename(archivo_procesado)),
                "CodIntBIM":      cod_oficial,
            }

            for h in headers:
                if h in ("Archivo", "ElementId", "nombre_archivo", "CodIntBIM"):
                    continue
                valor_h = datos.get(h, "")
                fila[h]           = _norm(valor_h)
                valores_header[h] = valor_h

            result[clave_repo]            = fila
            valores_por_clave[clave_repo] = valores_header

        except Exception:
            continue

    save_json(cache_modelo, CACHE_MODELO_PATH, show_error=False)
    return result, cods_por_clave, valores_por_clave


def build_combined_dataset(headers, codigo_planilla):
    model_rows, cods_por_clave, valores_por_clave = get_filtered_rows_from_model(
        headers, codigo_planilla
    )
    rows = list(model_rows.values())
    rows.sort(key=lambda r: str(r.get("CodIntBIM", "") or "").lower())
    return rows, cods_por_clave, valores_por_clave

#--------------------------------------------------
# Lanzar visor Tkinter (editor de planilla)

def run_viewer_tk(planilla_meta):
    save_json(planilla_meta, PLANILLA_META_PATH, show_error=False)
    if not PYTHON_EXE or not os.path.isfile(PYTHON_EXE):
        forms.alert(
            u"No se encontró Python 3 en este equipo.",
            title="Python no encontrado"
        )
        return
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

    # script.json es opcional: si no existe el boton continua igual
    script_cfg = load_script_config()

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

    filas, cods_por_clave, valores_por_clave = build_combined_dataset(headers, codigo_planilla)
    save_json(filas, PLANILLA_DATA_PATH, show_error=False)

    planilla_meta = {
        "Headers":         headers,
        "CodigoPlanilla":  codigo_planilla,
        "NombrePlanilla":  nombre_alias or nombre_original,
        "DataPath":        PLANILLA_DATA_PATH,
        "CodsPorClave":    cods_por_clave,
        "ValoresPorClave": valores_por_clave,
    }

    run_viewer_tk(planilla_meta)


if __name__ == "__main__":
    main()

#==================================================
#🚫 DELETE BELOW
#from Snippets._customprint import kit_button_clicked    
#kit_button_clicked(btn_name=__title__)                  

