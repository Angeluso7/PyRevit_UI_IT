# -*- coding: utf-8 -*-
__title__ = "Por Elemento"

__doc__ = """Version = 1.0
Date = 01.09.2024
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
# ╩╩ ╩╩ ╚═╝╩╚═ ╩ ╚═╝
#==================================================

import clr
import os
import json
import subprocess

clr.AddReference("RevitAPI")
clr.AddReference("System.Windows.Forms")

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    RevitLinkInstance,
    ViewSchedule
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
# Carpeta común de datos (repositorios principales)
DATA_DIR = os.path.join(
    os.path.expanduser("~"),
    r"AppData\\Roaming\\MyPyRevitExtention\\PyRevitIT.extension\\data"
)

# Carpeta del pushbutton actual (para JSON temporal y visor)
CURRENT_FOLDER = os.path.dirname(__file__)
CPYTHON_SCRIPT_PATH = os.path.join(CURRENT_FOLDER, "codigo_se2.py")

# JSON temporal SOLO para pasar datos al visor (no repositorio)
REPO_ELEMENTO_TMP_PATH = os.path.join(CURRENT_FOLDER, "repo_elemento_tmp.json")

# Repositorios principales en /data
SCRIPT_JSON_PATH = os.path.join(DATA_DIR, "script.json")
REPOSITORIO_DATOS_PATH = os.path.join(DATA_DIR, "repositorio_datos.json")


#---------------- Utilidades JSON ------------------
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


def save_json(data, path):
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        MessageBox.Show(
            u"Error guardando JSON:\n{}".format(e),
            "Error"
        )
        return False


#------------- Lógica principal --------------------
def seleccionar_elemento():
    """Permite seleccionar con clic un elemento (en link o en host)."""
    try:
        # Primero intentar como elemento linkeado
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
        # Si se cancela o falla, intentar como elemento del documento activo
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


def obtener_encabezados_planilla(host_doc, codintbim_val, script_data):
    """Obtiene la clave de planilla y el orden de encabezados."""
    reemplazos_planillas = script_data.get("codigos_planillas", {})

    pref_cod = (codintbim_val or "")[:4]
    clave_planilla = None

    for clave, vals in reemplazos_planillas.items():
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
            u"No se encontró clave de planilla para el código '{}'.".format(pref_cod),
            "Aviso"
        )
        return None, None

    try:
        schedules = (FilteredElementCollector(host_doc)
                     .OfClass(ViewSchedule)
                     .WhereElementIsNotElementType()
                     .ToElements())
        planilla_obj = next(
            (s for s in schedules if s.Name == clave_planilla), None
        )
        if planilla_obj is None:
            MessageBox.Show(
                u"No se encontró la planilla '{}'.".format(clave_planilla),
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


def construir_datos_elemento(elem, linked_doc, headers_order, script_data):
    """Arma el diccionario de datos a mostrar, usando repositorio_datos si existe."""
    # Reemplazos de nombres de parámetros
    reemplazos_nombres = script_data.get("reemplazos_de_nombres", {})

    # 1) Leer repositorio_datos.json
    repo_datos = load_json(REPOSITORIO_DATOS_PATH, show_error=False)
    if repo_datos is None:
        repo_datos = {}

    elem_id_str = str(elem.Id.IntegerValue)
    archivo = linked_doc.PathName if linked_doc is not None else doc.PathName or ""
    base_key = u"{}_{}".format(archivo, elem_id_str)

    data_elemento = repo_datos.get(base_key, None)

    resultado = {}

    try:
        if data_elemento:
            # 2) Si existe en repositorio_datos: usar sus valores
            filtrado = {
                k: v for k, v in data_elemento.items()
                if k not in ["ElementId", "Archivo"]
            }
            tmp = {reemplazos_nombres.get(k, k): v for k, v in filtrado.items()}
            for head in headers_order:
                if head in tmp:
                    resultado[head] = tmp[head]
        else:
            # 3) Si no existe: leer directamente desde el elemento
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


def main():
    # 1) Seleccionar elemento (link o host)
    elem, linked_doc = seleccionar_elemento()
    if elem is None:
        return

    # 2) Cargar script.json desde carpeta data
    script_data = load_json(SCRIPT_JSON_PATH)
    if not script_data:
        return

    # 3) Obtener CodIntBIM y planilla asociada
    try:
        p_cod = elem.LookupParameter("CodIntBIM")
        cod_val = p_cod.AsString() if (p_cod and p_cod.HasValue) else ""
    except:
        cod_val = ""

    host_doc = doc  # las planillas residen en el documento activo
    clave_planilla, headers_order = obtener_encabezados_planilla(
        host_doc, cod_val, script_data
    )
    if not headers_order:
        return

    # 4) Construir los datos desde repositorio_datos o desde el modelo
    resultado, archivo, elem_id_str = construir_datos_elemento(
        elem, linked_doc, headers_order, script_data
    )
    if resultado is None:
        return

    # 5) Armar JSON temporal para visor (estructura similar a repo_elemento_tmp)
    json_temporal = {
        "Archivo": archivo,
        "Headers": headers_order,
        "ElementId": elem_id_str,
        "Row": resultado
    }

    if not save_json(json_temporal, REPO_ELEMENTO_TMP_PATH):
        return

    # 6) Lanzar visor externo en modo solo lectura
    #    (ajusta la ruta de Python si es necesario)
    PYTHON_EXE = r"C:\\Users\\Zbook HP\\AppData\\Local\\Programs\\Python\\Python313\\pythonw.exe"
    try:
        subprocess.Popen([PYTHON_EXE, CPYTHON_SCRIPT_PATH, REPO_ELEMENTO_TMP_PATH])
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
