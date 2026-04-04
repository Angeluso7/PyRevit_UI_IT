# -*- coding: utf-8 -*-
__title__ = "Por Elemento"

__doc__ = """Version = 1.0
Date = 01.09.2024
_______________________________________________________________
Description:

Selecciona un elemento (host o link), arma sus datos según la planilla
asociada a CodIntBIM tomando primero el repositorio activo
(ruta_repositorio_activo) y, si no hay datos, desde el modelo.
Pasa la información a un visor CPython (Tkinter) en modo solo lectura.
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
# ╩╩ ╩╩ ╚═╝╩╚═ ╩ ╚═╝
#==================================================

import clr
import os
import json
import subprocess

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System.Windows.Forms")

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    RevitLinkInstance,
    ViewSchedule,
)
from Autodesk.Revit.UI.Selection import ObjectType
from System.Windows.Forms import MessageBox

# ╦ ╦╔═╗╦═╗╦╔═╗╔╗ ╦ ╔═╗╔═╗
# ╚╗╔╝╠═╣╠╦╝║╠═╣╠╩╗║ ║╣ ╚═╗
# ╚╝ ╩ ╩╩╚═╩╩ ╩╚═╝╩═╝╚═╝╚═╝
#==================================================

app = __revit__.Application
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

# ╔╦╗╔═╗╦╔╗╔
# ║║║╠═╣║║║║
# ╩ ╩╩ ╩╩╝╚╝
#==================================================

#==================================================
# Rutas comunes de datos

DATA_DIR = os.path.join(
    os.path.expanduser("~"),
    r"AppData\Roaming\MyPyRevitExtention\PyRevitIT.extension\data"
)

CONFIG_PATH = os.path.join(DATA_DIR, "config_proyecto_activo.json")
SCRIPT_JSON_PATH = os.path.join(DATA_DIR, "script.json")

# Carpeta del pushbutton actual
CURRENT_FOLDER = os.path.dirname(__file__)

# JSON local para orden de encabezados por planilla
PLANILLAS_ORDER_PATH = os.path.join(CURRENT_FOLDER, "planillas_headers_order.json")

# JSON temporal para el visor CPython
REPO_ELEMENTO_TMP_PATH = os.path.join(CURRENT_FOLDER, "repo_elemento_tmp.json")

# Ruta del ejecutable CPython y del visor
PYTHON_EXE = r"C:\Users\Zbook HP\AppData\Local\Programs\Python\Python313\pythonw.exe"
CPYTHON_VIEWER_PATH = os.path.join(CURRENT_FOLDER, "visor_elemento.pyw")

#==================================================
# Utilidades JSON

def load_json(path, show_error=True):
    if not os.path.exists(path):
        if show_error:
            MessageBox.Show(
                u"No se encontró archivo JSON:\n{}".format(path),
                "Error"
            )
        return None
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        if show_error:
            MessageBox.Show(
                u"Error cargando JSON:\n{}".format(e),
                "Error"
            )
        return None


def save_json(data, path, show_error=True):
    try:
        folder = os.path.dirname(path)
        if folder and not os.path.exists(folder):
            os.makedirs(folder)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        if show_error:
            MessageBox.Show(
                u"Error guardando JSON:\n{}".format(e),
                "Error"
            )
        return False

#==================================================
# Repositorio activo desde config_proyecto_activo.json

def get_repo_path_from_config():
    if not os.path.exists(CONFIG_PATH):
        MessageBox.Show(
            u"No se encontró config_proyecto_activo.json en:\n{}\n\n"
            u"No se puede determinar el repositorio de datos activo."
            .format(CONFIG_PATH),
            "Config no encontrada"
        )
        return None
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            cfg = json.load(f)
        ruta = (cfg.get("ruta_repositorio_activo") or "").strip()
        if not ruta:
            MessageBox.Show(
                u"En config_proyecto_activo.json no se encontró la clave "
                u"'ruta_repositorio_activo' o está vacía.\n\n"
                u"No se puede determinar el repositorio de datos activo.",
                "Config incompleta"
            )
            return None
        return ruta
    except Exception as e:
        MessageBox.Show(
            u"Error leyendo config_proyecto_activo.json:\n{}\n\n"
            u"No se puede determinar el repositorio de datos activo."
            .format(e),
            "Error config"
        )
        return None


REPO_PATH = get_repo_path_from_config()


def load_repo():
    if not REPO_PATH:
        return {}
    repo = load_json(REPO_PATH, show_error=False)
    return repo or {}

#==================================================
# Selección de elemento

def seleccionar_elemento():
    """Selecciona un elemento (link o host)."""
    try:
        ref = uidoc.Selection.PickObject(
            ObjectType.LinkedElement,
            "Seleccione un elemento (link o modelo)"
        )
        try:
            link_instance = doc.GetElement(ref.ElementId)
            if isinstance(link_instance, RevitLinkInstance):
                linked_doc = link_instance.GetLinkDocument()
                if linked_doc is None:
                    MessageBox.Show(
                        "No se pudo acceder al documento linkeado.",
                        "Error"
                    )
                    return None, None
                elem = linked_doc.GetElement(ref.LinkedElementId)
                if elem is None:
                    MessageBox.Show("Elemento linkeado no encontrado.", "Error")
                    return None, None
                return elem, linked_doc
        except:
            pass
    except:
        try:
            ref = uidoc.Selection.PickObject(
                ObjectType.Element,
                "Seleccione un elemento del modelo"
            )
            elem = doc.GetElement(ref.ElementId)
            if elem is None:
                MessageBox.Show("Elemento no encontrado.", "Error")
                return None, None
            return elem, doc
        except Exception as e:
            MessageBox.Show(
                u"Error/cancelación en selección:\n{}".format(e),
                "Aviso"
            )
            return None, None

    return None, None

#==================================================
# Planilla y orden de encabezados

def load_script_json():
    data = load_json(SCRIPT_JSON_PATH)
    return data or {}


def obtener_encabezados_planilla(host_doc, codintbim_val, script_data):
    codigos_planillas = script_data.get("codigos_planillas", {})
    pref_cod = (codintbim_val or "")[:4]
    clave_planilla = None

    for clave, vals in codigos_planillas.items():
        if isinstance(vals, list):
            if any((v or "").startswith(pref_cod) for v in vals):
                clave_planilla = clave
                break
        elif isinstance(vals, str):
            if vals.startswith(pref_cod):
                clave_planilla = clave
                break

    if not clave_planilla:
        MessageBox.Show(
            u"No se encontró clave de planilla para el código '{}'."
            .format(pref_cod),
            "Aviso"
        )
        return None, None

    try:
        schedules = (
            FilteredElementCollector(host_doc)
            .OfClass(ViewSchedule)
            .WhereElementIsNotElementType()
            .ToElements()
        )

        planilla_obj = next(
            (s for s in schedules if s.Name == clave_planilla), None
        )

        if planilla_obj is None:
            MessageBox.Show(
                u"No se encontró la planilla '{}'."
                .format(clave_planilla),
                "Error"
            )
            return None, None
    except Exception as e:
        MessageBox.Show(
            u"Error buscando planilla:\n{}".format(e),
            "Error"
        )
        return None, None

    try:
        headers_order = []
        for fid in planilla_obj.Definition.GetFieldOrder():
            field = planilla_obj.Definition.GetField(fid)
            if field:
                headers_order.append(field.GetName())
        if not headers_order:
            MessageBox.Show(
                "No se obtuvieron encabezados de la planilla.",
                "Aviso"
            )
            return None, None
    except Exception as e:
        MessageBox.Show(
            u"Error obteniendo encabezados:\n{}".format(e),
            "Error"
        )
        return None, None

    return clave_planilla, headers_order


def obtener_headers_desde_cache_o_planilla(host_doc, codintbim_val, script_data):
    planillas_order = load_json(PLANILLAS_ORDER_PATH, show_error=False) or {}

    clave_planilla, headers_from_model = obtener_encabezados_planilla(
        host_doc, codintbim_val, script_data
    )
    if not headers_from_model:
        return None, []

    if clave_planilla in planillas_order:
        headers = planillas_order.get(clave_planilla) or []
    else:
        headers = headers_from_model
        planillas_order[clave_planilla] = headers_from_model
        save_json(planillas_order, PLANILLAS_ORDER_PATH, show_error=False)

    return clave_planilla, headers

#==================================================
# Construcción de datos para mostrar

def construir_datos_elemento(elem, linked_doc, headers_order, script_data):
    """Usa primero repo activo y, si no hay, modelo."""
    reemplazos_nombres = script_data.get("reemplazos_de_nombres", {})

    repo_datos = load_repo()

    elem_id_str = str(elem.Id.IntegerValue)
    archivo = linked_doc.PathName if linked_doc is not None else (doc.PathName or "")
    base_key = u"{}_{}".format(archivo, elem_id_str)

    data_elemento = repo_datos.get(base_key, None)
    resultado = {}

    try:
        if data_elemento:
            filtrado = {
                k: v for k, v in data_elemento.items()
                if k not in ["ElementId", "Archivo", "nombre_archivo"]
            }
            tmp = {reemplazos_nombres.get(k, k): v for k, v in filtrado.items()}
            for head in headers_order:
                if head in tmp:
                    resultado[head] = tmp[head]
        else:
            parametros = {}
            for p in elem.Parameters:
                try:
                    if p.HasValue:
                        val = p.AsString() or p.AsValueString()
                        if val and val.strip():
                            parametros[p.Definition.Name] = val.strip()
                except:
                    continue
            tmp = {reemplazos_nombres.get(k, k): v for k, v in parametros.items()}
            for head in headers_order:
                if head in tmp:
                    resultado[head] = tmp[head]
    except Exception as e:
        MessageBox.Show(
            u"Error preparando datos del elemento:\n{}".format(e),
            "Error"
        )
        return None, archivo, elem_id_str

    return resultado, archivo, elem_id_str

#==================================================
# main

def main():
    elem, linked_doc = seleccionar_elemento()
    if elem is None:
        return

    script_data = load_script_json()
    if not script_data:
        return

    try:
        p_cod = elem.LookupParameter("CodIntBIM")
        cod_val = p_cod.AsString() if (p_cod and p_cod.HasValue) else ""
    except:
        cod_val = ""

    host_doc = doc

    clave_planilla, headers_order = obtener_headers_desde_cache_o_planilla(
        host_doc, cod_val, script_data
    )
    if not headers_order:
        return

    resultado, archivo, elem_id_str = construir_datos_elemento(
        elem, linked_doc, headers_order, script_data
    )
    if resultado is None:
        return

    json_temporal = {
        "Archivo": archivo,
        "ElementId": elem_id_str,
        "Planilla": clave_planilla,
        "Headers": headers_order,
        "Row": resultado
    }

    if not save_json(json_temporal, REPO_ELEMENTO_TMP_PATH):
        return

    try:
        subprocess.Popen(
            [PYTHON_EXE, CPYTHON_VIEWER_PATH, REPO_ELEMENTO_TMP_PATH],
            shell=False
        )
    except Exception as e:
        MessageBox.Show(
            u"No se pudo ejecutar visor externo:\n{}".format(e),
            "Error"
        )


if __name__ == "__main__":
    main()

#==================================================
#🚫 DELETE BELOW
from Snippets._customprint import kit_button_clicked  # Import Reusable Function from 'lib/Snippets/_customprint.py'
kit_button_clicked(btn_name=__title__)  # Display Default Print Message
