# -*- coding: utf-8 -*-
__title__   = "Por Planilla IT"
__doc__     = """Version = 1.2
Date    = 20.04.2026
________________________________________________________________
Description:

Filtra la vista activa por planilla usando los filtros
'b_x_planilla' y 'b_x_planilla_x'.

FIX v1.2:
- Se elimina el error "Parameter does not apply to this filter's
  categories" reconstruyendo las categorias del filtro para incluir
  SOLO las que tienen CodIntBIM asignado antes de aplicar la regla.
- Si ninguna categoria del filtro soporta CodIntBIM se avisa al
  usuario con un mensaje claro.
________________________________________________________________
Author: Argenis Angel"""

# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝
#==================================================

import clr
import os
import json
import subprocess

clr.AddReference("RevitAPI")
clr.AddReference("RevitServices")
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

import System
import sys as _sys

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    ViewSchedule,
    ParameterFilterElement,
    ParameterFilterRuleFactory,
    ElementId,
    Transaction,
    ElementParameterFilter,
    OverrideGraphicSettings,
)

from System.Collections.Generic import List
from System.Windows.Forms import MessageBox

# ╦  ╦╔═╗╦═╗╦╔═╗╔╗ ╦  ╔═╗╔═╗
# ╚╗╔╝╠═╣╠╦╝║╠═╣╠╩╗║  ║╣ ╚═╗
#  ╚╝ ╩ ╩╩╚═╩╩ ╩╚═╝╩═╝╚═╝╚═╝
#==================================================
doc   = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# ╔╦╗╔═╗╦╔╗╔
# ║║║╠═╣║║║║
# ╩ ╩╩ ╩╩╝╚╝
#==================================================

# ── Rutas centralizadas ──────────────────────────────────────────────────────
_lib = os.path.normpath(os.path.join(os.path.dirname(__file__),
                                     '..', '..', '..', '..', 'lib'))
if _lib not in _sys.path:
    _sys.path.insert(0, _lib)

try:
    from config.paths import DATA_DIR, MASTER_DIR, TEMP_DIR, ensure_runtime_dirs
    from core.env_config import get_python_exe
    ensure_runtime_dirs()
    PYTHON_EXE = get_python_exe()
except Exception:
    _ext_root  = os.path.normpath(os.path.join(os.path.dirname(__file__),
                                                '..', '..', '..', '..'))
    DATA_DIR   = os.path.join(_ext_root, 'data')
    MASTER_DIR = os.path.join(DATA_DIR, 'master')
    TEMP_DIR   = os.path.join(DATA_DIR, 'temp')

    import glob as _glob
    def _fb_python():
        base = os.path.join(os.path.expanduser('~'), 'AppData',
                            'Local', 'Programs', 'Python')
        for exe in ('python.exe', 'pythonw.exe'):
            for cand in sorted(
                    _glob.glob(os.path.join(base, 'Python3*', exe)), reverse=True):
                return cand
        for folder in os.environ.get('PATH', '').split(os.pathsep):
            cand = os.path.join(folder.strip(), 'python.exe')
            if os.path.isfile(cand):
                return cand
        return None
    PYTHON_EXE = _fb_python()

# ── Archivos de datos ────────────────────────────────────────────────────────
archivo_json        = os.path.join(MASTER_DIR, 'script.json')
SELECCION_OUT_PATH  = os.path.join(TEMP_DIR,   'planilla_seleccion_tmp.json')
META_SELECTOR_PATH  = os.path.join(TEMP_DIR,   'planillas_selector_meta.json')

BASE_PATH           = os.path.dirname(__file__)
CPYTHON_SELECTOR_TK = os.path.join(BASE_PATH,  'selector_planillas_tk.pyw')

#--------------------------------------------------
# Utilidades Revit / JSON

def get_param_id_by_name(doc, param_name):
    iterator = doc.ParameterBindings.ForwardIterator()
    while iterator.MoveNext():
        definition = iterator.Key
        if definition.Name == param_name:
            return definition.Id
    return None


def obtener_planillas_desde_documento(doc):
    collector = FilteredElementCollector(doc).OfClass(ViewSchedule)
    return [s.Name for s in collector if not s.IsTemplate]


#--------------------------------------------------
# FIX PRINCIPAL:
# Obtener solo las categorias del filtro donde el parametro SI aplica.
# Esto evita el error:
#   "One of the given rules refers to a parameter that does not
#    apply to this filter's categories. Parameter name: elementFilter"

def get_categorias_con_parametro(doc, param_id, filtro_obj):
    """
    Devuelve una List[ElementId] con las categorias del filtro que
    tienen el parametro param_id asignado en sus bindings.
    Si una categoria no tiene el parametro, se excluye silenciosamente.
    """
    # Construir set de cat_ids donde el parametro esta vinculado
    cats_con_param = set()
    iterator = doc.ParameterBindings.ForwardIterator()
    while iterator.MoveNext():
        defn = iterator.Key
        if defn.Id == param_id:
            binding = iterator.Current
            try:
                for c in binding.Categories:
                    cats_con_param.add(c.Id.IntegerValue)
            except Exception:
                pass
            break

    cat_ids_filtro = filtro_obj.GetCategories()
    cats_validas = List[ElementId]()
    for cid in cat_ids_filtro:
        if cid.IntegerValue in cats_con_param:
            cats_validas.Add(cid)

    return cats_validas


#--------------------------------------------------
# Lanzar selector Tkinter externo (CPython)

def run_selector_tkinter(planillas_doc, ruta_json):
    meta = {
        "planillas_doc": planillas_doc,
        "ruta_json": ruta_json,
    }
    try:
        with open(META_SELECTOR_PATH, "w", encoding="utf-8") as f:
            json.dump(meta, f, indent=2, ensure_ascii=False)
    except Exception as e:
        MessageBox.Show(
            "Error escribiendo meta para selector Tk:\n{}".format(e), "Error")
        return None

    if os.path.exists(SELECCION_OUT_PATH):
        try:
            os.remove(SELECCION_OUT_PATH)
        except Exception:
            pass

    try:
        subprocess.Popen(
            [PYTHON_EXE, CPYTHON_SELECTOR_TK,
             META_SELECTOR_PATH, SELECCION_OUT_PATH],
            shell=False
        ).wait()
    except Exception as e:
        MessageBox.Show(
            "No se pudo ejecutar selector Tkinter externo:\n{}".format(e),
            "Error selector")
        return None

    if not os.path.exists(SELECCION_OUT_PATH):
        return None

    try:
        with open(SELECCION_OUT_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data.get("selected_planilla")
    except Exception as e:
        MessageBox.Show(
            "Error leyendo seleccion del selector Tk:\n{}".format(e), "Error")
        return None


#--------------------------------------------------
# Modificacion de filtros — VERSION CORREGIDA

def modificar_filtros(doc, nombres_filtros, valor_parametro,
                      nombre_parametro="CodIntBIM"):
    """
    Actualiza las reglas de los filtros:
    - Para 'b_x_planilla'   usa Contains.
    - Para 'b_x_planilla_x' usa NotContains.

    FIX: antes de aplicar la regla, restringe las categorias del filtro
    a solo aquellas donde CodIntBIM esta vinculado, evitando el error
    'Parameter does not apply to this filter's categories'.
    """
    filtro_collector    = FilteredElementCollector(doc).OfClass(ParameterFilterElement)
    filtros_encontrados = [f for f in filtro_collector
                           if f.Name in nombres_filtros]

    if not filtros_encontrados:
        MessageBox.Show(
            "No se encontraron filtros con nombres {}.".format(nombres_filtros))
        return None

    param_id = get_param_id_by_name(doc, nombre_parametro)
    if param_id is None or param_id == ElementId.InvalidElementId:
        MessageBox.Show(
            "No se encontro el parametro '{}' en el documento.\n"
            "Verifique que el parametro compartido este cargado.".format(
                nombre_parametro))
        return None

    ids_modificados = []

    with Transaction(doc, "Modificar reglas de filtros") as t:
        t.Start()

        for filtro_obj in filtros_encontrados:

            # ── FIX: obtener solo las cats donde el param aplica ──────────
            cats_validas = get_categorias_con_parametro(doc, param_id, filtro_obj)

            if cats_validas.Count == 0:
                MessageBox.Show(
                    "El parametro '{}' no esta asignado a ninguna de las\n"
                    "categorias del filtro '{}'.\n\n"
                    "Abre Visibility/Graphics > Filters en Revit, edita\n"
                    "el filtro y asegurate de que sus categorias tengan\n"
                    "el parametro compartido cargado.".format(
                        nombre_parametro, filtro_obj.Name))
                t.RollBack()
                return None

            # Actualizar las categorias del filtro (solo las validas)
            try:
                filtro_obj.SetCategories(cats_validas)
            except Exception:
                pass  # si no se puede cambiar, continuar e intentar la regla

            # Crear la regla segun el nombre del filtro
            if filtro_obj.Name == "b_x_planilla":
                regla_nueva = ParameterFilterRuleFactory.CreateContainsRule(
                    param_id, valor_parametro, False)
            else:
                regla_nueva = ParameterFilterRuleFactory.CreateNotContainsRule(
                    param_id, valor_parametro, False)

            filtro_nuevo = ElementParameterFilter(regla_nueva)

            try:
                filtro_obj.SetElementFilter(filtro_nuevo)
            except Exception as ex:
                MessageBox.Show(
                    "Error al aplicar la regla al filtro '{}':\n{}\n\n"
                    "Sugerencia: verifica que el parametro '{}' este\n"
                    "disponible en todas las categorias de ese filtro.".format(
                        filtro_obj.Name, ex, nombre_parametro))
                t.RollBack()
                return None

            ids_modificados.append(filtro_obj.Id)

        t.Commit()

    return ids_modificados


#--------------------------------------------------
# Activacion de filtros en la vista

def activar_filtro_unico(doc, vista_activa, filtro_activar_id):
    """
    - b_x_planilla   : visible y activado.
    - b_x_planilla_x : NO visible, pero ACTIVADO.
    - Resto          : desactivados y sin visibilidad.
    """
    filtros_aplicados = vista_activa.GetFilters()

    filtros_dict = {}
    for f in FilteredElementCollector(doc).OfClass(ParameterFilterElement):
        filtros_dict[f.Id] = f

    with Transaction(doc, "Actualizar filtros visibilidad y activacion") as t:
        t.Start()

        for filtro_id in filtros_aplicados:
            filtro_obj = filtros_dict.get(filtro_id)
            if filtro_obj is None:
                continue

            nombre_filtro = filtro_obj.Name

            if nombre_filtro not in ("b_x_planilla", "b_x_planilla_x"):
                vista_activa.SetFilterVisibility(filtro_id, False)
                vista_activa.SetIsFilterEnabled(filtro_id, False)
                vista_activa.SetFilterOverrides(filtro_id, OverrideGraphicSettings())
                continue

            if nombre_filtro == "b_x_planilla":
                vista_activa.SetFilterVisibility(filtro_id, True)
                vista_activa.SetIsFilterEnabled(filtro_id, True)
                vista_activa.SetFilterOverrides(filtro_id, OverrideGraphicSettings())

            elif nombre_filtro == "b_x_planilla_x":
                vista_activa.SetFilterVisibility(filtro_id, False)
                vista_activa.SetIsFilterEnabled(filtro_id, True)

        t.Commit()


#--------------------------------------------------
# main

def main():
    planillas_doc = obtener_planillas_desde_documento(doc)

    if not os.path.exists(archivo_json):
        MessageBox.Show(
            "No se encontro el archivo de codigos de planillas:\n{}".format(
                archivo_json), "Error")
        return

    if not PYTHON_EXE or not os.path.isfile(PYTHON_EXE):
        MessageBox.Show(
            "No se encontro Python 3 en este equipo.\n"
            "Verifica la instalacion de CPython.", "Error")
        return

    seleccion = run_selector_tkinter(planillas_doc, archivo_json)
    if not seleccion:
        return

    with open(archivo_json, "r", encoding="utf-8") as f:
        data = json.load(f)
    codigos_planillas = data.get("codigos_planillas", {})

    if seleccion not in codigos_planillas:
        MessageBox.Show(
            "No se encontro codigo para la planilla '{}' en el JSON.".format(
                seleccion), "Aviso")
        return

    valor_regla = codigos_planillas[seleccion]

    filtro_ids = modificar_filtros(
        doc, ["b_x_planilla", "b_x_planilla_x"], valor_regla)
    if filtro_ids is None:
        return

    vista_activa = doc.ActiveView
    activar_filtro_unico(doc, vista_activa, filtro_ids[0])

    MessageBox.Show(
        "Filtros actualizados para planilla '{}'.".format(seleccion),
        "Listo")


#--------------------------------------------------
if __name__ == "__main__":
    main()
