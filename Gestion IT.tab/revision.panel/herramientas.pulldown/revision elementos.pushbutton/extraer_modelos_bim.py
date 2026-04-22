# -*- coding: utf-8 -*-
"""
extraer_modelos_bim.py - IronPython (PyRevit helper)
Recorre modelos linkeados, extrae elementos con CodIntBIM,
agrupa por CM y genera _temp_datos.json en data/output.

Se invoca desde script.py via:
    from extraer_modelos_bim import ejecutar_extraccion_y_json
"""

import os
import json
from collections import defaultdict, OrderedDict

from pyrevit import forms

import clr
clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BuiltInCategory,
    CategoryType,
    ViewSchedule,
    SectionType,
)

# ── Categorias a excluir ─────────────────────────────────────
NOMBRES_CATEGORIAS_EXCLUIR = set([
    "Materiales",
    "Circuitos electricos",
    "Elementos de detalle",
    "Informacion de proyecto",
    "Sistemas de tuberias",
    "Tramos de bandeja de cables",
    "Tramos de tubo",
    "Vinculos RVT",
    "Zonas de climatizacion",
    "Lineas",
    "Segmentos de tuberia",
    "Emplazamientos",
    "Equipo medico",
    "Planos",
    # Con tildes (por si el modelo las incluye)
    u"Materiales",
    u"Circuitos el\u00e9ctricos",
    u"Elementos de detalle",
    u"Informaci\u00f3n de proyecto",
    u"Sistemas de tuber\u00edas",
    u"Tramos de bandeja de cables",
    u"Tramos de tubo",
    u"V\u00ednculos RVT",
    u"Zonas de climatizaci\u00f3n",
    u"L\u00edneas",
    u"Segmentos de tuber\u00eda",
    u"Emplazamientos",
    u"Equipo m\u00e9dico",
    u"Planos",
])

# ── Helpers internos ─────────────────────────────────────────
def _cargar_codigos_planillas(script_json_path, log_fn=None):
    try:
        with open(script_json_path, "r", encoding="utf-8") as f:
            cfg = json.load(f)
        return cfg.get("codigos_planillas", {})
    except Exception as e:
        if log_fn:
            log_fn("_cargar_codigos_planillas ERROR: {}".format(e))
        return {}

def _obtener_categorias_modelo_ids(doc, log_fn=None):
    """Solo categorias de modelo, excluyendo lista por nombre."""
    ids = set()
    try:
        cats = doc.Settings.Categories
        for c in cats:
            try:
                if c.CategoryType == CategoryType.Model and not c.IsTagCategory:
                    if c.Name in NOMBRES_CATEGORIAS_EXCLUIR:
                        continue
                    ids.add(c.Id)
            except Exception:
                continue
    except Exception as e:
        if log_fn:
            log_fn("_obtener_categorias_modelo_ids ERROR: {}".format(e))
    return ids

def _obtener_headers_por_schedule(doc, codigos_planillas, log_fn=None):
    """
    Retorna dict {nombre_schedule: [headers desde 'val-sql']}
    """
    result = {}
    try:
        schedules = (
            FilteredElementCollector(doc)
            .OfClass(ViewSchedule)
            .ToElements()
        )
        by_name = {vs.Name: vs for vs in schedules}

        for schedule_name in codigos_planillas.keys():
            vs = by_name.get(schedule_name)
            if not vs:
                result[schedule_name] = []
                continue
            try:
                body  = vs.GetTableData().GetSectionData(SectionType.Body)
                ncols = body.NumberOfColumns

                headers_all  = []
                idx_val_sql  = None
                for c in range(ncols):
                    txt = body.GetCellText(0, c) or ""
                    headers_all.append(txt)
                    if txt.strip().lower() == "val-sql":
                        idx_val_sql = c

                if idx_val_sql is None:
                    result[schedule_name] = headers_all
                else:
                    result[schedule_name] = headers_all[idx_val_sql:]
            except Exception as e:
                if log_fn:
                    log_fn("_headers schedule '{}' error: {}".format(schedule_name, e))
                result[schedule_name] = []
    except Exception as e:
        if log_fn:
            log_fn("_obtener_headers_por_schedule error: {}".format(e))
    return result

def _extraer_elementos_linkeados(doc, codigos_planillas, model_cat_ids, log_fn=None):
    """
    Recorre links, extrae elementos y agrupa por CM.
    Retorna dict con elementos_por_tabla, excepciones, listado_tablas.
    """
    datos_elementos = []
    excepciones     = []
    codigos_validos = set(codigos_planillas.values())

    try:
        link_instances = (
            FilteredElementCollector(doc)
            .OfCategory(BuiltInCategory.OST_RvtLinks)
            .WhereElementIsNotElementType()
            .ToElements()
        )
    except Exception as e:
        if log_fn:
            log_fn("Error obteniendo links: {}".format(e))
        return {"elementos_por_tabla": OrderedDict(), "excepciones": [], "listado_tablas": {}}

    for link_instance in link_instances:
        try:
            linked_doc = link_instance.GetLinkDocument()
            if linked_doc is None:
                continue

            nombre_rvt = (
                os.path.basename(linked_doc.PathName)
                if linked_doc.PathName
                else "SinNombre"
            )
            if log_fn:
                log_fn("  Procesando link: {}".format(nombre_rvt))

            elems = (
                FilteredElementCollector(linked_doc)
                .WhereElementIsNotElementType()
                .ToElements()
            )

            for el in elems:
                try:
                    cat = el.Category
                    if not cat or cat.Id not in model_cat_ids:
                        continue

                    elem_id   = el.Id.IntegerValue
                    categoria = cat.Name if cat else "Sin Categoria"

                    base = {
                        "ElementId"  : elem_id,
                        "CodIntBIM"  : "",
                        u"Categor\u00eda"   : categoria,
                        "Familia"    : "",
                        "Tipo"       : "",
                        "Nombre_RVT" : nombre_rvt,
                    }

                    # Familia / Tipo
                    if hasattr(el, "Symbol") and el.Symbol:
                        try: base["Familia"] = el.Symbol.FamilyName
                        except Exception: pass
                        try: base["Tipo"]    = el.Symbol.Name
                        except Exception: pass

                    # CodIntBIM
                    codint_param = el.LookupParameter("CodIntBIM")
                    tiene_param  = codint_param is not None
                    codint       = ""
                    if tiene_param and codint_param.AsString():
                        codint = codint_param.AsString().strip()

                    base["CodIntBIM"] = codint

                    # ── Excepciones ──
                    if not tiene_param:
                        exc = dict(base)
                        exc["CodIntBIM"] = "No existe"
                        excepciones.append({"elemento": exc, "situacion": "No existe"})
                        continue
                    if not codint:
                        exc = dict(base)
                        exc["CodIntBIM"] = "No Asignado"
                        excepciones.append({"elemento": exc, "situacion": "No Asignado"})
                        continue
                    if len(codint) < 4:
                        exc = dict(base)
                        exc["CodIntBIM"] = "No Asignado"
                        excepciones.append({"elemento": exc, "situacion": "No Asignado"})
                        continue

                    pref = codint[:4]
                    if pref not in codigos_validos:
                        exc = dict(base)
                        exc["CodIntBIM"] = "errado"
                        excepciones.append({"elemento": exc, "situacion": "errado"})
                        continue

                    codigo_cm = pref  # CM01, CM02, ...

                    # ── Parámetros adicionales ──
                    if hasattr(el, "Parameters"):
                        for p in el.Parameters:
                            try:
                                pname = p.Definition.Name
                                if pname in base:
                                    continue
                                stype = p.StorageType.ToString()
                                val   = ""
                                if stype == "String":
                                    val = p.AsString() or ""
                                elif stype == "Integer":
                                    v   = p.AsInteger()
                                    val = "" if v is None else str(v)
                                elif stype == "Double":
                                    v = p.AsDouble()
                                    if v is not None:
                                        val = "{:.2f}".format(round(v, 2))
                                elif stype == "ElementId":
                                    pid = p.AsElementId()
                                    val = "" if pid is None else str(pid.IntegerValue)
                                if val:
                                    base[pname] = val
                            except Exception:
                                continue

                    datos_elementos.append((codigo_cm, base))

                except Exception:
                    continue

        except Exception:
            continue

    # Agrupar por CM
    elementos_por_codigo = defaultdict(list)
    for codigo, eldict in datos_elementos:
        elementos_por_codigo[codigo].append(eldict)

    codigos_ordenados = sorted(set(codigos_planillas.values()))

    od = OrderedDict()
    for codigo in codigos_ordenados:
        if codigo in elementos_por_codigo:
            od[codigo] = elementos_por_codigo[codigo]

    listado_tablas = {
        "valores": codigos_ordenados,
        "claves" : [
            k for k, v in codigos_planillas.items()
            if v in codigos_ordenados
        ]
    }

    return {
        "elementos_por_tabla" : od,
        "excepciones"         : excepciones,
        "listado_tablas"      : listado_tablas,
    }

# ── Funcion principal exportada ──────────────────────────────
def ejecutar_extraccion_y_json(script_json_path, datos_json_path,
                                data_master_dir, log_fn=None):
    """
    Punto de entrada llamado desde script.py.
    Retorna True si el JSON se generó correctamente, False si hubo error.
    """
    # Documento activo (accedido aqui para no fallar al importar)
    try:
        doc = __revit__.ActiveUIDocument.Document
    except Exception as e:
        forms.alert(u"No se pudo acceder al documento Revit:\n{}".format(e), title="Error")
        if log_fn: log_fn("Error doc: {}".format(e))
        return False

    # Cargar codigos_planillas
    codigos_planillas = _cargar_codigos_planillas(script_json_path, log_fn)
    if not codigos_planillas:
        forms.alert(
            u"No se encontraron 'codigos_planillas' en script.json:\n{}".format(script_json_path),
            title="Error"
        )
        return False

    if log_fn:
        log_fn("codigos_planillas cargados: {}".format(len(codigos_planillas)))

    # Categorias de modelo
    model_cat_ids = _obtener_categorias_modelo_ids(doc, log_fn)
    if log_fn:
        log_fn("Categorias modelo: {}".format(len(model_cat_ids)))

    # Headers por schedule
    headers_por_schedule = _obtener_headers_por_schedule(doc, codigos_planillas, log_fn)

    # Extraccion de elementos
    if log_fn: log_fn("Extrayendo elementos de modelos linkeados...")
    datos = _extraer_elementos_linkeados(doc, codigos_planillas, model_cat_ids, log_fn)

    # Construir JSON serializable
    try:
        serializable = {
            "elementos_por_tabla": {
                codigo: [
                    {kk: ("" if vv is None else vv) for kk, vv in el.items()}
                    for el in elems
                ]
                for codigo, elems in datos["elementos_por_tabla"].items()
            },
            "excepciones"    : datos["excepciones"],
            "listado_tablas" : datos["listado_tablas"],
            "headers_por_tabla": {
                # codigo CMxx -> headers desde val-sql
                v: headers_por_schedule.get(k, [])
                for k, v in codigos_planillas.items()
            }
        }
    except Exception as e:
        forms.alert(u"Error serializando datos:\n{}".format(e), title="Error")
        if log_fn: log_fn("Error serializando: {}".format(e))
        return False

    # Guardar JSON temporal
    try:
        d = os.path.dirname(datos_json_path)
        if d and not os.path.exists(d):
            os.makedirs(d)
        with open(datos_json_path, "w", encoding="utf-8") as f:
            json.dump(serializable, f, ensure_ascii=False, indent=2)
        if log_fn:
            log_fn("_temp_datos.json escrito: {}".format(datos_json_path))
        return True
    except Exception as e:
        forms.alert(u"Error al escribir JSON temporal:\n{}".format(e), title="Error")
        if log_fn: log_fn("Error escribiendo JSON: {}".format(e))
        return False