# -*- coding: utf-8 -*-
__title__   = "Por Grupo de Elementos"
__doc__     = """Version = 1.0
Date    = 15.06.2024
________________________________________________________________
Description:

This is the placeholder for a .pushbutton in a /pulldown
You can use it to start your pyRevit Add-In

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

import os
import json
import clr
import subprocess

from pyrevit import forms

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    RevitLinkInstance,
)

import filtros_codint  # debe estar en la misma carpeta que este script

# ╦  ╦╔═╗╦═╗╦╔═╗╔╗ ╦  ╔═╗╔═╗
# ╚╗╔╝╠═╣╠╦╝║╠═╣╠╩╗║  ║╣ ╚═╗
#  ╚╝ ╩ ╩╩╚═╩╩ ╩╚═╝╩═╝╚═╝╚═╝
#==================================================

doc = __revit__.ActiveUIDocument.Document

# ╔╦╗╔═╗╦╔╗╔
# ║║║╠═╣║║║║
# ╩ ╩╩ ╩╩╝╚╝
#==================================================
# ---------------- rutas y config ---------------- #

def get_data_dir():
    return os.path.join(
        os.path.expanduser("~"),
        r"AppData\\Roaming\\MyPyRevitExtention\\PyRevitIT.extension\\data"
    )


def get_script_json_path():
    return os.path.join(get_data_dir(), "script.json")


def load_config():
    with open(get_script_json_path(), "r", encoding="utf-8") as f:
        return json.load(f)


def get_cache_path():
    this_folder = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(this_folder, "codintbim_cache.json")


# ---------------- utilidades ---------------- #

def get_param_string(elem, pname):
    p = elem.LookupParameter(pname)
    if p and p.HasValue:
        return p.AsString() or ""
    return ""


# ---------------- recolección CodIntBIM vínculos ---------------- #

def collect_codintbim_from_links():
    all_values = set()
    link_instances = FilteredElementCollector(doc).OfClass(RevitLinkInstance)
    for link in link_instances:
        link_doc = link.GetLinkDocument()
        if link_doc is None:
            continue
        collector = FilteredElementCollector(link_doc).WhereElementIsNotElementType()
        for elem in collector:
            cod = get_param_string(elem, "CodIntBIM")
            if not cod:
                continue
            last_dash = cod.rfind("-")
            cod_trim = cod[:last_dash] if last_dash > 0 else cod
            all_values.add(cod_trim)
    return all_values


# ---------------- construcción estructura cache ---------------- #

def build_structure_from_codintbim(cod_list, cfg):
    dict_esp = cfg.get("especialidad", {})
    dict_cod = cfg.get("codigos_elementos", {})

    result_esp = {}
    cod_list_norm = [(c, c.lower()) for c in cod_list]

    for esp_key, esp_name in dict_esp.items():
        if not esp_key:
            continue
        key_low = esp_key.lower()
        cods_for_esp = [c for c, cl in cod_list_norm if key_low in cl]
        if not cods_for_esp:
            continue

        codigos_presentes = set()
        for elem_key, elem_name in dict_cod.items():
            if not elem_key:
                continue
            elem_key_low = elem_key.lower()
            for c, cl in cod_list_norm:
                if c not in cods_for_esp:
                    continue
                if elem_key_low in cl:
                    codigos_presentes.add(elem_name)
                    break

        if codigos_presentes:
            result_esp[esp_name] = {
                "codigos_elementos": sorted(list(codigos_presentes))
            }

    return {
        "codintbim": sorted(list(cod_list)),
        "especialidad": result_esp
    }


# ---------------- UI Tkinter ---------------- #

def lanzar_ui_tk():
    this_folder = os.path.dirname(os.path.abspath(__file__))
    python_exe = r"C:\Users\Zbook HP\AppData\Local\Programs\Python\Python313\python.exe"
    ui_script = os.path.join(this_folder, "ui_tk_codintbim.py")
    cache_json = get_cache_path()
    output_json = os.path.join(this_folder, "ui_tk_codintbim_output.json")

    cmd = [python_exe, ui_script, cache_json, output_json]
    CREATE_NO_WINDOW = 0x08000000

    proc = subprocess.Popen(
        cmd,
        cwd=this_folder,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        creationflags=CREATE_NO_WINDOW
    )
    proc.communicate()

    if not os.path.exists(output_json):
        return None

    with open(output_json, "r", encoding="utf-8") as f:
        return json.load(f)


# ---------------- main ---------------- #

def main():
    cfg = load_config()

    cod_values = collect_codintbim_from_links()
    structure = build_structure_from_codintbim(cod_values, cfg)

    with open(get_cache_path(), "w", encoding="utf-8") as f:
        json.dump(structure, f, ensure_ascii=False, indent=2)

    if not structure.get("especialidad"):
        forms.alert("No hay coincidencias entre CodIntBIM y los diccionarios.", title="Sin datos")
        return

    seleccion = lanzar_ui_tk()
    if not seleccion:
        return

    esp = seleccion.get("especialidad")
    cod_nombre = seleccion.get("codigo_elemento")

    if not cod_nombre:
        forms.alert("No se seleccionó ningún código de elemento.", title="Sin selección")
        return

    dict_cod = cfg.get("codigos_elementos", {})
    clave_seleccionada = None
    for k, v in dict_cod.items():
        if v == cod_nombre:
            clave_seleccionada = k
            break

    filtros_codint.ajustar_reglas_filtros_codintbim(doc, clave_seleccionada)

    forms.alert(
        u"Proceso correcto.\n\nEspecialidad: {}\nCódigo (nombre): {}\nClave usada: {}"
        .format(esp or "(ninguna)", cod_nombre, clave_seleccionada or "(no encontrada)"),
        title="Proceso Correcto"
    )


if __name__ == "__main__":
    main()

#==================================================
#🚫 DELETE BELOW
#from Snippets._customprint import kit_button_clicked     Import Reusable Function #from 'lib/Snippets/_customprint.py'
#kit_button_clicked(btn_name=__title__)                   Display Default Print Message
