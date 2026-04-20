# -*- coding: utf-8 -*-
__title__   = "Por Planilla IT"
__doc__     = """Version = 1.1
Date    = 19.04.2026
________________________________________________________________
Description:

Filtra la vista activa por planilla usando los filtros
'b_x_planilla' y 'b_x_planilla_x'.

- Usa un selector externo en Tkinter (CPython) con filtro de texto
  y seleccion unica de planilla.
- Ajusta las reglas de ambos filtros segun el codigo de la planilla.
- Deja 'b_x_planilla' visible y activado.
- Deja 'b_x_planilla_x' oculto pero con la casilla "Activar filtro" marcada.
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
    BuiltInParameter,
    ElementParameterFilter,
    LogicalAndFilter,
    ElementFilter,
    OverrideGraphicSettings,
)

from System.Collections.Generic import List
from System.Windows.Forms import MessageBox
from System.Drawing import Size, Point


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


def obtener_claves_json(archivo_json):
    if not os.path.exists(archivo_json):
        return []
    with open(archivo_json, "r", encoding="utf-8") as f:
        data = json.load(f)
    return list(data.get("codigos_planillas", {}).keys())

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
# Verificar que param aplica a TODAS las categorias del filtro

def param_aplica_a_categorias(doc, param_id, filtro_obj):
    """
    Devuelve True si el parametro (param_id) aplica a todas las
    categorias asignadas al ParameterFilterElement.
    Revit lanza excepcion si alguna categoria no soporta el parametro.
    """
    try:
        cat_ids = filtro_obj.GetCategories()
        for cat_id in cat_ids:
            cat = doc.Settings.Categories.get_Item(cat_id)
            if cat is None:
                continue
            # Verificar que la categoria tiene el parametro disponible
            found = False
            iterator = doc.ParameterBindings.ForwardIterator()
            while iterator.MoveNext():
                defn = iterator.Key
                if defn.Id == param_id:
                    binding = iterator.Current
                    # SharedParameterElement / InternalDefinition
                    try:
                        cats = binding.Categories
                        for c in cats:
                            if c.Id == cat_id:
                                found = True
                                break
                    except Exception:
                        found = True  # si no se puede leer, asumimos que aplica
                    break
            if not found:
                return False
        return True
    except Exception:
        return True   # ante la duda, intentar y dejar que Revit valide

#--------------------------------------------------
# Modificacion de filtros de parametros

def modificar_filtros(doc, nombres_filtros, valor_parametro,
                      nombre_parametro="CodIntBIM"):
    """
    Actualiza las reglas de los filtros:
    - Para 'b_x_planilla'   usa Contains.
    - Para 'b_x_planilla_x' usa NotContains.

    Antes de aplicar, verifica que el parametro aplique a las
    categorias del filtro para evitar la excepcion de Revit:
    'One of the given rules refers to a parameter that does not
     apply to this filter's categories.'
    """
    filtro_collector  = FilteredElementCollector(doc).OfClass(ParameterFilterElement)
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

            # ── VALIDACION: el param debe aplicar a las categorias del filtro
            if not param_aplica_a_categorias(doc, param_id, filtro_obj):
                MessageBox.Show(
                    "El parametro '{}' no aplica a todas las categorias "
                    "del filtro '{}'.\n"
                    "Ajusta las categorias del filtro en Revit para incluir "
                    "solo elementos que tengan ese parametro.".format(
                        nombre_parametro, filtro_obj.Name))
                t.RollBack()
                return None

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
                    "Sugerencia: revisa que el parametro '{}' este "
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
    # 1) Planillas desde documento y JSON
    planillas_doc = obtener_planillas_desde_documento(doc)

    if not os.path.exists(archivo_json):
        MessageBox.Show(
            "No se encontro el archivo de codigos de planillas:\n{}".format(
                archivo_json), "Error")
        return

    # 2) Selector externo Tkinter (CPython)
    if not PYTHON_EXE or not os.path.isfile(PYTHON_EXE):
        MessageBox.Show(
            "No se encontro Python 3 en este equipo.\n"
            "Verifica la instalacion de CPython.", "Error")
        return

    seleccion = run_selector_tkinter(planillas_doc, archivo_json)
    if not seleccion:
        MessageBox.Show("Operacion cancelada.", "Aviso")
        return

    # 3) Obtener codigo de planilla desde script.json
    with open(archivo_json, "r", encoding="utf-8") as f:
        data = json.load(f)
    codigos_planillas = data.get("codigos_planillas", {})

    if seleccion not in codigos_planillas:
        MessageBox.Show(
            "No se encontro codigo para la planilla '{}' en el JSON.".format(
                seleccion), "Aviso")
        return

    valor_regla = codigos_planillas[seleccion]

    # 4) Actualizar reglas de ambos filtros
    filtro_ids = modificar_filtros(
        doc, ["b_x_planilla", "b_x_planilla_x"], valor_regla)
    if filtro_ids is None:
        return

    # 5) Activar filtros en la vista
    vista_activa = doc.ActiveView
    activar_filtro_unico(doc, vista_activa, filtro_ids[0])

    MessageBox.Show(
        "Filtros 'b_x_planilla' y 'b_x_planilla_x' actualizados\n"
        "para planilla '{}'.".format(seleccion),
        "Informacion")


#--------------------------------------------------
if __name__ == "__main__":
    main()
