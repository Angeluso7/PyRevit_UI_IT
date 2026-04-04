# -*- coding: utf-8 -*-
# Extrae datos de modelos linkeados y genera _temp_datos.json

import os
import json
from collections import defaultdict, OrderedDict
from pyrevit import forms
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BuiltInCategory,
    CategoryType,
    ViewSchedule,
    SectionType
)

doc = __revit__.ActiveUIDocument.Document


# ---------------------------------------------------------
# Configuración de exclusiones de categorías
# ---------------------------------------------------------
NOMBRES_CATEGORIAS_EXCLUIR = set([
    "Materiales",
    "Circuitos eléctricos",
    "Elementos de detalle",
    "Información de proyecto",
    "Sistemas de tuberías",
    "Tramos de bandeja de cables",
    "Tramos de tubo",
    "Vínculos RVT",
    "Zonas de climatización",
    "Líneas",
    "Segmentos de tubería",
    "Emplazamientos",
    "Equipo médico",
    "Planos"
])


def _cargar_codigos_planillas(script_json_path):
    try:
        with open(script_json_path, 'r', encoding='utf-8') as f:
            cfg = json.load(f)
        return cfg.get('codigos_planillas', {})
    except Exception as e:
        forms.alert("Error al leer script.json:\n{}".format(e), title="Error")
        return {}


def _obtener_categorias_modelo_ids():
    """Solo categorías de modelo, excluyendo una lista por nombre."""
    cats = doc.Settings.Categories
    ids = set()
    for c in cats:
        try:
            if c.CategoryType == CategoryType.Model and not c.IsTagCategory:
                if c.Name in NOMBRES_CATEGORIAS_EXCLUIR:
                    continue
                ids.add(c.Id)
        except:
            continue
    return ids


def _obtener_headers_por_schedule(codigos_planillas):
    """
    nombre_schedule -> lista headers (desde 'val-sql' a la derecha).
    """
    result = {}
    schedules = FilteredElementCollector(doc) \
        .OfClass(ViewSchedule) \
        .ToElements()
    by_name = {vs.Name: vs for vs in schedules}

    for schedule_name in codigos_planillas.keys():
        vs = by_name.get(schedule_name)
        if not vs:
            result[schedule_name] = []
            continue

        try:
            table_data = vs.GetTableData()
            body = table_data.GetSectionData(SectionType.Body)
            ncols = body.NumberOfColumns

            headers_all = []
            idx_val_sql = None
            for c in range(ncols):
                txt = body.GetCellText(0, c) or ""
                headers_all.append(txt)
                if txt.strip().lower() == "val-sql":
                    idx_val_sql = c

            if idx_val_sql is None:
                result[schedule_name] = headers_all
            else:
                result[schedule_name] = headers_all[idx_val_sql:]
        except:
            result[schedule_name] = []

    return result


def _extraer_elementos_linkeados(codigos_planillas, model_cat_ids):
    datos_elementos = []
    excepciones = []

    codigos_validos = set(codigos_planillas.values())

    link_instances = FilteredElementCollector(doc) \
        .OfCategory(BuiltInCategory.OST_RvtLinks) \
        .WhereElementIsNotElementType() \
        .ToElements()

    for link_instance in link_instances:
        try:
            linked_doc = link_instance.GetLinkDocument()
            if linked_doc is None:
                continue

            nombre_rvt = os.path.basename(linked_doc.PathName) if linked_doc.PathName else "SinNombre"

            elems = FilteredElementCollector(linked_doc) \
                .WhereElementIsNotElementType() \
                .ToElements()

            for el in elems:
                try:
                    cat = el.Category
                    if not cat or cat.Id not in model_cat_ids:
                        continue

                    elem_id = el.Id.IntegerValue
                    categoria = cat.Name if cat else "Sin Categoría"

                    codint_param = el.LookupParameter("CodIntBIM")
                    tiene_param = codint_param is not None
                    codint = ""
                    if tiene_param and codint_param.AsString():
                        codint = codint_param.AsString().strip()

                    base = {
                        "ElementId": elem_id,
                        "CodIntBIM": codint,
                        "Categoría": categoria,
                        "Familia": "",
                        "Tipo": "",
                        "Nombre_RVT": nombre_rvt
                    }

                    if hasattr(el, "Symbol") and el.Symbol:
                        try:
                            base["Familia"] = el.Symbol.FamilyName
                        except:
                            pass
                        try:
                            base["Tipo"] = el.Symbol.Name
                        except:
                            pass

                    # Excepciones por CodIntBIM
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

                    # Parámetros adicionales
                    if hasattr(el, "Parameters"):
                        for p in el.Parameters:
                            try:
                                pname = p.Definition.Name
                                if pname in base:
                                    continue
                                stype = p.StorageType.ToString()
                                val = ""
                                if stype == "String":
                                    val = p.AsString() or ""
                                elif stype == "Integer":
                                    v = p.AsInteger()
                                    val = "" if v is None else str(v)
                                elif stype == "Double":
                                    v = p.AsDouble()
                                    val = "" if v is None else str(v)
                                elif stype == "ElementId":
                                    pid = p.AsElementId()
                                    val = "" if pid is None else str(pid.IntegerValue)
                                if val:
                                    base[pname] = val
                            except:
                                pass

                    datos_elementos.append((codigo_cm, base))

                except:
                    continue

        except:
            continue

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
        "claves": [
            k for k, v in codigos_planillas.items()
            if v in codigos_ordenados
        ]
    }

    return {
        "elementos_por_tabla": od,
        "excepciones": excepciones,
        "listado_tablas": listado_tablas
    }


def ejecutar_extraccion_y_json(data_dir_ext):
    script_json = os.path.join(data_dir_ext, "script.json")
    codigos_planillas = _cargar_codigos_planillas(script_json)
    if not codigos_planillas:
        forms.alert("No se encontraron 'codigos_planillas' en script.json", title="Error")
        return False

    model_cat_ids = _obtener_categorias_modelo_ids()
    forms.alert("Extrayendo datos de modelos linkeados...", title="Procesando")

    headers_por_schedule = _obtener_headers_por_schedule(codigos_planillas)
    datos = _extraer_elementos_linkeados(codigos_planillas, model_cat_ids)

    ruta_json = os.path.join(data_dir_ext, "_temp_datos.json")
    try:
        serializable = {
            "elementos_por_tabla": {
                codigo: [
                    {kk: ("" if vv is None else vv) for kk, vv in el.items()}
                    for el in elems
                ]
                for codigo, elems in datos["elementos_por_tabla"].items()
            },
            "excepciones": datos["excepciones"],
            "listado_tablas": datos["listado_tablas"],
            "headers_por_tabla": {
                # código CMxx -> headers desde val-sql
                v: headers_por_schedule.get(k, [])
                for k, v in codigos_planillas.items()
            }
        }
        with open(ruta_json, "w", encoding="utf-8") as f:
            json.dump(serializable, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        forms.alert("Error al escribir JSON temporal:\n{}".format(e), title="Error")
        return False
