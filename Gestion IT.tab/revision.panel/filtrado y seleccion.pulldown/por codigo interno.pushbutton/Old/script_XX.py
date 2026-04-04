# -*- coding: utf-8 -*-
__title__   = "Por Código Interno"
__doc__     = """Version = 1.0
Date    = 01.09.2024
_______________________________________________________________
Description:

Aplicación para seleccionar elementos en modelo por categorías.

________________________________________________________________
How-To:

1. [Hold ALT + CLICK] on the button to open its source folder.
You will be able to override this placeholder.

2. Automate Your Boring Work ;)

________________________________________________________________
TODO:
[FEATURE] - Describe Your ToDo Tasks Here
________________________________________________________________
Last Updates:
- [01.09.2025] v0.1 Inicio de Aplicación. 
________________________________________________________________
Author: Argenis Angel"""

# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝
#==================================================
# Importaciones base para pyRevit script con interfaz WPF y Revit API

import os
import json
import subprocess
import clr
import time

from pyrevit import forms

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    RevitLinkInstance,
    View3D,
    ParameterFilterElement,
    Transaction,
    ParameterFilterRuleFactory,
    ElementParameterFilter,
    ElementId,
)


# ╦  ╦╔═╗╦═╗╦╔═╗╔╗ ╦  ╔═╗╔═╗
# ╚╗╔╝╠═╣╠╦╝║╠═╣╠╩╗║  ║╣ ╚═╗
#  ╚╝ ╩ ╩╩╚═╩╩ ╩╚═╝╩═╝╚═╝╚═╝
#==================================================

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# ╔╦╗╔═╗╦╔╗╔
# ║║║╠═╣║║║║
# ╩ ╩╩ ╩╩╝╚╝
#==================================================


DATA_DIR = os.path.join(
    os.path.expanduser("~"),
    r"AppData\\Roaming\\MyPyRevitExtention\\PyRevitIT.extension\\data"
)
TMP_IN = os.path.join(DATA_DIR, "codint_selector_in.json")
TMP_OUT = os.path.join(DATA_DIR, "codint_selector_out.json")

PYTHON_EXE = r"C:\\Users\\Zbook HP\\AppData\\Local\\Programs\\Python\\Python313\\pythonw.exe"

# Clave del diccionario ElementId ↔ CodIntBIM en el repositorio
KEY_ELEM_DICT = "elementos_codintbim"


# -------------------------------------------------------
# Utilidades JSON proyecto
# -------------------------------------------------------
def cargar_json(ruta, default):
    try:
        if not os.path.exists(ruta):
            return default
        with open(ruta, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return default


def get_repo_activo_path():
    cfg_path = os.path.join(DATA_DIR, "config_proyecto_activo.json")
    cfg = cargar_json(cfg_path, {})
    return cfg.get("ruta_repositorio_activo", "")


def cargar_repo_activo():
    path = get_repo_activo_path()
    if not path or not os.path.exists(path):
        return {}
    return cargar_json(path, {})


# -------------------------------------------------------
# Sincronización modelo ↔ base (ElementId / CodIntBIM)
# -------------------------------------------------------
def sincronizar_repo_con_modelo():
    """
    Sincroniza el diccionario ElementId ↔ CodIntBIM entre el modelo y la base de datos.
    Devuelve (repo_dict, elem_dict).
    """
    repo = cargar_repo_activo() or {}
    elem_dict = repo.get(KEY_ELEM_DICT, {})

    modelo_dict = {}

    # Elementos del documento activo
    for el in FilteredElementCollector(doc).WhereElementIsNotElementType():
        p = el.LookupParameter("CodIntBIM")
        if p and p.HasValue:
            cod = (p.AsString() or "").strip()
            if cod:
                eid_str = str(el.Id.IntegerValue)
                modelo_dict[eid_str] = {
                    "ElementId": el.Id.IntegerValue,
                    "CodIntBIM": cod,
                }

    # Elementos de documentos linkeados (opcional)
    try:
        links = FilteredElementCollector(doc).OfClass(RevitLinkInstance)
        for li in links:
            ldoc = li.GetLinkDocument()
            if not ldoc:
                continue
            for el in FilteredElementCollector(ldoc).WhereElementIsNotElementType():
                p = el.LookupParameter("CodIntBIM")
                if p and p.HasValue:
                    cod = (p.AsString() or "").strip()
                    if cod:
                        eid_str = str(el.Id.IntegerValue)
                        modelo_dict[eid_str] = {
                            "ElementId": el.Id.IntegerValue,
                            "CodIntBIM": cod,
                        }
    except Exception:
        pass

    # Eliminar entradas que ya no están en el modelo
    ids_modelo = set(modelo_dict.keys())
    ids_base = set(elem_dict.keys())
    ids_a_eliminar = ids_base - ids_modelo
    for k in ids_a_eliminar:
        elem_dict.pop(k, None)

    # Agregar/actualizar entradas del modelo
    for k, v in modelo_dict.items():
        elem_dict[k] = v

    repo[KEY_ELEM_DICT] = elem_dict

    repo_path = get_repo_activo_path()
    if repo_path:
        try:
            with open(repo_path, "w", encoding="utf-8") as f:
                json.dump(repo, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    return repo, elem_dict


# -------------------------------------------------------
# Recolección de CodIntBIM para la ventana (usa diccionario sincronizado)
# -------------------------------------------------------
def recoger_codintbim():
    _, elem_dict = sincronizar_repo_con_modelo()

    resultados = {"elementos": []}
    for k, v in elem_dict.items():
        resultados["elementos"].append(
            {
                "doc_path": doc.PathName or "",
                "element_id": v.get("ElementId", None),
                "codintbim": v.get("CodIntBIM", ""),
            }
        )

    return resultados


# -------------------------------------------------------
# Lanzar ventana Tkinter (codint_selector.py)
# -------------------------------------------------------
def lanzar_tkinter_selector():
    try:
        bdir = os.path.dirname(os.path.abspath(__file__))
    except Exception:
        bdir = os.getcwd()

    cpython_script = os.path.join(bdir, "codint_selector.py")

    if not os.path.exists(PYTHON_EXE):
        forms.alert(
            "No se encontró el ejecutable de Python 3:\n{}\nAjusta PYTHON_EXE en script.py.".format(PYTHON_EXE),
            title="Python no encontrado",
        )
        return False

    if not os.path.exists(cpython_script):
        forms.alert(
            "No se encontró codint_selector.py en la carpeta del botón:\n{}".format(cpython_script),
            title="Script CPython no encontrado",
        )
        return False

    datos = recoger_codintbim()
    if not os.path.exists(DATA_DIR):
        os.makedirs(DATA_DIR)
    with open(TMP_IN, "w", encoding="utf-8") as f:
        json.dump(datos, f, ensure_ascii=False, indent=2)

    if os.path.exists(TMP_OUT):
        try:
            os.remove(TMP_OUT)
        except Exception:
            pass

    cmd = [PYTHON_EXE, cpython_script, TMP_IN, TMP_OUT]
    try:
        subprocess.Popen(cmd, cwd=bdir)
        return True
    except Exception as e:
        forms.alert(
            "Error al ejecutar CPython:\n{}\n\nComando:\n{}".format(e, " ".join(cmd)),
            title="Error CPython",
        )
        return False


def leer_salida_tk():
    if not os.path.exists(TMP_OUT):
        return None
    try:
        with open(TMP_OUT, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


# -------------------------------------------------------
# Helper: obtener Id de parámetro por nombre (como en planillas)
# -------------------------------------------------------
def get_param_id_by_name(doc_local, param_name):
    iterator = doc_local.ParameterBindings.ForwardIterator()
    param_id = None
    while iterator.MoveNext():
        definition = iterator.Key
        if definition.Name == param_name:
            param_id = definition.Id
            break
    return param_id


# -------------------------------------------------------
# Lógica renglón 2 (Asignados / No asignados) c_cod_int / s_cod_int
# -------------------------------------------------------
def activar_filtros_por_nombre(vista, filtro_activo, otros_filtros):
    doc_local = vista.Document

    collector = list(FilteredElementCollector(doc_local).OfClass(ParameterFilterElement))

    filtro_a_activar = None
    otros = []
    for fil in collector:
        if fil.Name == filtro_activo:
            filtro_a_activar = fil
        elif fil.Name in otros_filtros:
            otros.append(fil)

    with Transaction(doc_local, "Actualizar filtros CodIntBIM Asignado/No") as t:
        t.Start()

        if filtro_a_activar and not vista.IsFilterApplied(filtro_a_activar.Id):
            vista.AddFilter(filtro_a_activar.Id)
        for fil in otros:
            if not vista.IsFilterApplied(fil.Id):
                vista.AddFilter(fil.Id)

        for fid in list(vista.GetFilters()):
            fil = next((f for f in collector if f.Id == fid), None)
            if fil is None:
                continue

            if fil.Name == filtro_activo:
                vista.SetFilterVisibility(fid, True)
                try:
                    vista.SetIsFilterEnabled(fid, True)
                except:
                    pass
            elif fil.Name in otros_filtros:
                vista.SetFilterVisibility(fid, False)
                try:
                    vista.SetIsFilterEnabled(fid, True)
                except:
                    pass
            else:
                vista.SetFilterVisibility(fid, False)
                try:
                    vista.SetIsFilterEnabled(fid, False)
                except:
                    pass

        t.Commit()


# -------------------------------------------------------
# Lógica renglón 1 (f_element_x / f_element_y) usando patrón de planillas
# -------------------------------------------------------
def modificar_filtros_codint(doc_local, nombres_filtros, valor_codint, nombre_parametro="CodIntBIM"):
    filtro_collector = FilteredElementCollector(doc_local).OfClass(ParameterFilterElement)
    filtros_encontrados = [f for f in filtro_collector if f.Name in nombres_filtros]

    if not filtros_encontrados:
        forms.alert("No se encontraron filtros con nombres {}.".format(nombres_filtros),
                    title="Filtros no encontrados")
        return None, None

    param_id = get_param_id_by_name(doc_local, nombre_parametro)
    if param_id is None or param_id == ElementId.InvalidElementId:
        forms.alert("No se encontró parámetro '{}' para la regla.".format(nombre_parametro),
                    title="Parámetro no encontrado")
        return None, None

    filtro_x_id = None
    filtro_y_id = None

    with Transaction(doc_local, "Modificar reglas filtros f_element_x / f_element_y") as t:
        t.Start()

        for filtro_obj in filtros_encontrados:
            if filtro_obj.Name == "f_element_x":
                regla_nueva = ParameterFilterRuleFactory.CreateEqualsRule(param_id, valor_codint, False)
                filtro_x_id = filtro_obj.Id
            else:
                regla_nueva = ParameterFilterRuleFactory.CreateNotEqualsRule(param_id, valor_codint)
                filtro_y_id = filtro_obj.Id

            filtro_nuevo = ElementParameterFilter(regla_nueva)
            filtro_obj.SetElementFilter(filtro_nuevo)

        t.Commit()

    return filtro_x_id, filtro_y_id


def activar_filtros_codint_en_vista(doc_local, vista_activa, filtro_x_id, filtro_y_id):
    filtros_aplicados = vista_activa.GetFilters()

    with Transaction(doc_local, "Actualizar filtros visibilidad/activación CodIntBIM") as t:
        t.Start()

        for filtro_id in filtros_aplicados:
            filtro_obj = None
            for f in FilteredElementCollector(doc_local).OfClass(ParameterFilterElement):
                if f.Id == filtro_id:
                    filtro_obj = f
                    break
            if filtro_obj is None:
                continue

            if filtro_id == filtro_x_id:
                vista_activa.SetFilterVisibility(filtro_id, True)
                try:
                    vista_activa.SetIsFilterEnabled(filtro_id, True)
                except:
                    pass

            elif filtro_id == filtro_y_id:
                vista_activa.SetFilterVisibility(filtro_id, False)
                try:
                    vista_activa.SetIsFilterEnabled(filtro_id, True)
                except:
                    pass

            else:
                vista_activa.SetFilterVisibility(filtro_id, False)
                try:
                    vista_activa.SetIsFilterEnabled(filtro_id, False)
                except:
                    pass

        t.Commit()


def aplicar_filtros_por_codint(vista, valor_codint):
    doc_local = vista.Document

    filtro_x_id, filtro_y_id = modificar_filtros_codint(
        doc_local,
        ["f_element_x", "f_element_y"],
        valor_codint,
        nombre_parametro="CodIntBIM"
    )
    if filtro_x_id is None or filtro_y_id is None:
        return

    with Transaction(doc_local, "Aplicar filtros f_element_x / f_element_y a la vista") as t:
        t.Start()
        if not vista.IsFilterApplied(filtro_x_id):
            vista.AddFilter(filtro_x_id)
        if not vista.IsFilterApplied(filtro_y_id):
            vista.AddFilter(filtro_y_id)
        t.Commit()

    activar_filtros_codint_en_vista(doc_local, vista, filtro_x_id, filtro_y_id)


# -------------------------------------------------------
# main
# -------------------------------------------------------
def main():
    if doc is None:
        forms.alert("No hay documento activo.", title="Error")
        return

    if not isinstance(doc.ActiveView, View3D):
        forms.alert("La vista activa debe ser una vista 3D.", title="Aviso")
        return

    ok = lanzar_tkinter_selector()
    if not ok:
        return

    for _ in range(300):
        if os.path.exists(TMP_OUT):
            break
        time.sleep(0.1)

    salida = leer_salida_tk()
    if not salida:
        return

    opcion = salida.get("opcion", "")
    vista = doc.ActiveView

    if opcion == "by_codint":
        cod = salida.get("codintbim", "")
        if not cod:
            return
        aplicar_filtros_por_codint(vista, cod)

    elif opcion == "asignados":
        activar_filtros_por_nombre(vista, "c_cod_int", ["s_cod_int"])

    elif opcion == "no_asignados":
        activar_filtros_por_nombre(vista, "s_cod_int", ["c_cod_int"])

    else:
        return


if __name__ == "__main__":
    main()

#==================================================
#🚫 DELETE BELOW
from Snippets._customprint import kit_button_clicked    # Import Reusable Function from 'lib/Snippets/_customprint.py'
kit_button_clicked(btn_name=__title__)                  # Display Default Print Message
