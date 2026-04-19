# -*- coding: utf-8 -*-
__title__   = "Por Planilla IT"
__doc__     = """Version = 1.0
Date    = 15.06.2024
________________________________________________________________
Description:

Filtra la vista activa por planilla usando los filtros
'b_x_planilla' y 'b_x_planilla_x'.

- Usa un selector externo en Tkinter (CPython) con filtro de texto
  y selección única de planilla.
- Ajusta las reglas de ambos filtros según el código de la planilla.
- Deja 'b_x_planilla' visible y activado.
- Deja 'b_x_planilla_x' oculto pero con la casilla "Activar filtro" marcada.

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
- [15.06.2024] v1.0 Change Description
- [10.06.2024] v0.5 Change Description
- [05.06.2024] v0.1 Change Description 
________________________________________________________________
Author: Erik Frits"""

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
from System.Windows.Forms import (
    MessageBox,
)
from System.Drawing import Size, Point


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

# Carpeta de datos
_lib = os.path.normpath(os.path.join(os.path.dirname(__file__),
                                     '..', '..', '..', '..', 'lib'))
if _lib not in _sys.path:
    _sys.path.insert(0, _lib)

try:
    from config.paths import DATA_DIR, MASTER_DIR, TEMP_DIR, ensure_runtime_dirs
    from config.settings import CPYTHON_EXE
    ensure_runtime_dirs()
except Exception as _e:
    DATA_DIR   = os.path.join(os.path.expanduser('~'),
                    r'AppData\Roaming\MyPyRevitExtention\PyRevitIT.extension\data')
    MASTER_DIR = os.path.join(DATA_DIR, 'master')
    TEMP_DIR   = os.path.join(DATA_DIR, 'temp')
    CPYTHON_EXE = r'C:\Python313\python.exe'

# ✅ script.json ahora en master/
archivo_json = os.path.join(MASTER_DIR, 'script.json')

# Archivos temporales → temp/
SELECCION_OUT_PATH = os.path.join(TEMP_DIR, 'planilla_seleccion_tmp.json')
META_SELECTOR_PATH = os.path.join(TEMP_DIR, 'planillas_selector_meta.json')

# CPython desde settings (sin nombre de usuario hardcodeado)
PYTHON_EXE = CPYTHON_EXE

BASE_PATH = os.path.dirname(__file__)
CPYTHON_SELECTOR_TK = os.path.join(BASE_PATH, 'selector_planillas_tk.pyw')

#--------------------------------------------------
# Utilidades Revit / JSON

def get_param_id_by_name(doc, param_name):
    iterator = doc.ParameterBindings.ForwardIterator()
    param_id = None
    while iterator.MoveNext():
        definition = iterator.Key
        if definition.Name == param_name:
            param_id = definition.Id
            break
    return param_id


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
    """
    Lanza el selector en Tkinter (CPython) que permite:
    - Filtrar planillas por texto.
    - Seleccionar una sola planilla.
    El selector escribe la selección en SELECCION_OUT_PATH.
    """
    # Preparar meta de entrada
    meta = {
        "planillas_doc": planillas_doc,
        "ruta_json": ruta_json,
    }
    try:
        with open(META_SELECTOR_PATH, "w", encoding="utf-8") as f:
            json.dump(meta, f, indent=2, ensure_ascii=False)
    except Exception as e:
        MessageBox.Show(
            "Error escribiendo meta para selector Tk:\n{}".format(e),
            "Error"
        )
        return None

    # Borrar selección previa
    if os.path.exists(SELECCION_OUT_PATH):
        try:
            os.remove(SELECCION_OUT_PATH)
        except Exception:
            pass

    # Ejecutar CPython + selector Tkinter
    try:
        subprocess.Popen(
            [PYTHON_EXE, CPYTHON_SELECTOR_TK, META_SELECTOR_PATH, SELECCION_OUT_PATH],
            shell=False
        ).wait()
    except Exception as e:
        MessageBox.Show(
            "No se pudo ejecutar selector Tkinter externo:\n{}".format(e),
            "Error selector"
        )
        return None

    # Leer selección
    if not os.path.exists(SELECCION_OUT_PATH):
        # Usuario pudo cancelar ventana
        return None

    try:
        with open(SELECCION_OUT_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data.get("selected_planilla")
    except Exception as e:
        MessageBox.Show(
            "Error leyendo selección del selector Tk:\n{}".format(e),
            "Error"
        )
        return None

#--------------------------------------------------
# Modificación de filtros de parámetros

def modificar_filtros(doc, nombres_filtros, valor_parametro, nombre_parametro="CodIntBIM"):
    """
    Actualiza las reglas de los filtros:
    - Para 'b_x_planilla' usa Contains.
    - Para 'b_x_planilla_x' usa NotContains.
    """
    filtro_collector = FilteredElementCollector(doc).OfClass(ParameterFilterElement)
    filtros_encontrados = [f for f in filtro_collector if f.Name in nombres_filtros]

    if not filtros_encontrados:
        MessageBox.Show("No se encontraron filtros con nombres {}.".format(nombres_filtros))
        return None

    param_id = get_param_id_by_name(doc, nombre_parametro)
    if param_id is None or param_id == ElementId.InvalidElementId:
        MessageBox.Show("No se encontró parámetro '{}' para la regla.".format(nombre_parametro))
        return None

    ids_modificados = []

    with Transaction(doc, "Modificar reglas de filtros") as t:
        t.Start()

        for filtro_obj in filtros_encontrados:
            if filtro_obj.Name == "b_x_planilla":
                regla_nueva = ParameterFilterRuleFactory.CreateContainsRule(
                    param_id, valor_parametro, False
                )
            else:
                # Para b_x_planilla_x usar operador "no contiene"
                regla_nueva = ParameterFilterRuleFactory.CreateNotContainsRule(
                    param_id, valor_parametro, False
                )

            filtro_nuevo = ElementParameterFilter(regla_nueva)
            filtro_obj.SetElementFilter(filtro_nuevo)
            ids_modificados.append(filtro_obj.Id)

        t.Commit()

    return ids_modificados

#--------------------------------------------------
# Activación de filtros en la vista

def activar_filtro_unico(doc, vista_activa, filtro_activar_id):
    """
    - Desactiva todos los filtros aplicados, excepto b_x_planilla y b_x_planilla_x.
    - b_x_planilla: visible y activado.
    - b_x_planilla_x: NO visible, pero activado (casilla 'Activar filtro' marcada).
    """
    filtros_aplicados = vista_activa.GetFilters()

    # Cachear filtros para no recorrer el collector muchas veces
    filtros_dict = {}
    for f in FilteredElementCollector(doc).OfClass(ParameterFilterElement):
        filtros_dict[f.Id] = f

    with Transaction(doc, "Actualizar filtros visibilidad y activacion") as t:
        t.Start()

        for filtro_id in filtros_aplicados:
            filtro_obj = filtros_dict.get(filtro_id, None)
            if filtro_obj is None:
                continue

            nombre_filtro = filtro_obj.Name

            # Otros filtros: desactivados y sin visibilidad
            if nombre_filtro not in ("b_x_planilla", "b_x_planilla_x"):
                vista_activa.SetFilterVisibility(filtro_id, False)
                vista_activa.SetIsFilterEnabled(filtro_id, False)
                vista_activa.SetFilterOverrides(filtro_id, OverrideGraphicSettings())
                continue

            # Filtro principal b_x_planilla: visible y activado
            if nombre_filtro == "b_x_planilla":
                vista_activa.SetFilterVisibility(filtro_id, True)
                vista_activa.SetIsFilterEnabled(filtro_id, True)
                override_settings = OverrideGraphicSettings()
                vista_activa.SetFilterOverrides(filtro_id, override_settings)

            # Filtro complementario b_x_planilla_x:
            # se mantiene oculto pero ACTIVADO.
            elif nombre_filtro == "b_x_planilla_x":
                vista_activa.SetFilterVisibility(filtro_id, False)
                vista_activa.SetIsFilterEnabled(filtro_id, True)

        t.Commit()

#--------------------------------------------------
# main

def main():
    # 1) Planillas desde documento y JSON
    planillas_doc = obtener_planillas_desde_documento(doc)
    planillas_json = obtener_claves_json(archivo_json)
    nombres_combinados = sorted(set(planillas_doc) | set(planillas_json))

    if not nombres_combinados:
        MessageBox.Show("No se encontraron planillas.", "Aviso")
        return

    # 2) Selector externo Tkinter (CPython) con filtro y selección única
    seleccion = run_selector_tkinter(planillas_doc, archivo_json)
    if not seleccion:
        MessageBox.Show("Operación cancelada.", "Aviso")
        return

    # 3) Obtener código de planilla desde script.json
    if not os.path.exists(archivo_json):
        MessageBox.Show(
            "No se encontró el archivo de códigos de planillas:\n{}".format(archivo_json),
            "Aviso"
        )
        return

    with open(archivo_json, "r", encoding="utf-8") as f:
        data = json.load(f)
    codigos_planillas = data.get("codigos_planillas", {})

    if seleccion not in codigos_planillas:
        MessageBox.Show(
            "No se encontró código para la planilla '{}' en el JSON.".format(seleccion),
            "Aviso"
        )
        return

    valor_regla = codigos_planillas[seleccion]

    # 4) Actualizar reglas de ambos filtros
    filtro_ids = modificar_filtros(doc, ["b_x_planilla", "b_x_planilla_x"], valor_regla)
    if filtro_ids is None:
        return

    # 5) Activar filtros en la vista
    vista_activa = doc.ActiveView
    # Usamos el primero solo como referencia; la función internamente gestiona ambos
    activar_filtro_unico(doc, vista_activa, filtro_ids[0])

    MessageBox.Show(
        "Filtros 'b_x_planilla' y 'b_x_planilla_x' actualizados y aplicados "
        "para planilla '{}'.".format(seleccion),
        "Información"
    )

#--------------------------------------------------

if __name__ == "__main__":
    main()


#==================================================
#🚫 DELETE BELOW
from Snippets._customprint import kit_button_clicked    # Import Reusable Function from 'lib/Snippets/_customprint.py'
kit_button_clicked(btn_name=__title__)                  # Display Default Print Message
