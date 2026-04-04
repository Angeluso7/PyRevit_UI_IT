# -*- coding: utf-8 -*-
__title__   = "Parametros por planilla"
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

import clr
import os
import json
import codecs
import subprocess
import re

from pyrevit import forms
from pyrevit.forms import WPFWindow

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    RevitLinkInstance,
    ViewSchedule,
)
from System.Windows.Controls import RadioButton

# ╦  ╦╔═╗╦═╗╦╔═╗╔╗ ╦  ╔═╗╔═╗
# ╚╗╔╝╠═╣╠╦╝║╠═╣╠╩╗║  ║╣ ╚═╗
#  ╚╝ ╩ ╩╩╚═╩╩ ╩╚═╝╩═╝╚═╝╚═╝
#==================================================
#app    = __revit__.Application
#uidoc  = __revit__.ActiveUIDocument
#doc    = __revit__.ActiveUIDocument.Document #type:Document


# ╔╦╗╔═╗╦╔╗╔
# ║║║╠═╣║║║║
# ╩ ╩╩ ╩╩╝╚╝
#==================================================

DATA_DIR = os.path.join(
    os.path.expanduser("~"),
    r"AppData\Roaming\MyPyRevitExtention\PyRevitIT.extension\data",
)

CONFIG_PATH = os.path.join(DATA_DIR, "config_proyecto_activo.json")

BASE_PATH = os.path.dirname(__file__)          # carpeta del pushbutton
CPYTHON_SCRIPT_PATH = os.path.join(BASE_PATH, "mostrar_tabla.pyw")

# JSON temporal SOLO para el visor CPython
REPO_PLANILLA_PATH = os.path.join(BASE_PATH, "repo_planilla_tmp.json")

# Ejecutable de Python (ajusta si cambias versión/ruta)
PYTHON_EXE = r"C:\Users\Zbook HP\AppData\Local\Programs\Python\Python313\pythonw.exe"


# ------------ Repositorio activo desde config_proyecto_activo.json ------------
def get_repo_datos_path_from_config():
    """Devuelve la ruta del repositorio activo (ruta_repositorio_activo)."""
    if not os.path.exists(CONFIG_PATH):
        forms.alert(
            u"No se encontró config_proyecto_activo.json en:\n{}\n\n"
            u"No se puede determinar el repositorio de datos activo."
            .format(CONFIG_PATH),
            title="Config no encontrada",
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
                title="Config incompleta",
            )
            return None
        return ruta
    except Exception as e:
        forms.alert(
            u"Error leyendo config_proyecto_activo.json:\n{}\n\n"
            u"No se puede determinar el repositorio de datos activo."
            .format(e),
            title="Error config",
        )
        return None


REPOSITORIO_DATOS_PATH = get_repo_datos_path_from_config()


def load_repo_datos():
    """Carga el repositorio activo indicado en config_proyecto_activo.json."""
    if not REPOSITORIO_DATOS_PATH:
        return {}

    if os.path.exists(REPOSITORIO_DATOS_PATH):
        try:
            with open(REPOSITORIO_DATOS_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            forms.alert(
                u"Error leyendo el repositorio de datos:\n{}\nRuta:\n{}"
                .format(e, REPOSITORIO_DATOS_PATH),
                title="Error repositorio",
            )
            return {}
    else:
        # Si no existe, se parte de un diccionario vacío
        return {}


def save_repo_datos(data):
    """Guarda el repositorio activo con la estructura estándar."""
    if not REPOSITORIO_DATOS_PATH:
        forms.alert(
            u"No se definió una ruta válida de repositorio en "
            u"config_proyecto_activo.json.\n"
            u"No se pueden guardar datos.",
            title="Repositorio no definido",
        )
        return

    try:
        folder = os.path.dirname(REPOSITORIO_DATOS_PATH)
        if folder and not os.path.exists(folder):
            os.makedirs(folder)
        with open(REPOSITORIO_DATOS_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
    except Exception as e:
        forms.alert(
            u"Error guardando el repositorio de datos:\n{}"
            .format(e),
            title="Error repositorio",
        )


def save_repo_planilla_metadata(headers, codigo_planilla):
    """Guarda en repo_planilla_tmp.json solo metadatos (Headers + CodigoPlanilla)."""
    folder = os.path.dirname(REPO_PLANILLA_PATH)
    if not os.path.exists(folder):
        os.makedirs(folder)

    tmp_data = {
        "Headers": headers,
        "CodigoPlanilla": codigo_planilla,
    }

    with codecs.open(REPO_PLANILLA_PATH, "w", encoding="utf-8") as f:
        json.dump(tmp_data, f, indent=2, ensure_ascii=False)


def launch_cpython_viewer():
    subprocess.Popen(
        [PYTHON_EXE, CPYTHON_SCRIPT_PATH, REPO_PLANILLA_PATH],
        shell=False,
    )


#==================================================
class SelectorWPF(WPFWindow):
    def __init__(self):
        WPFWindow.__init__(self, os.path.join(BASE_PATH, "SelectorWPF.xaml"))

        self.doc = __revit__.ActiveUIDocument.Document
        self.script_config = self.load_script_json()
        self.reemplazos = self.script_config.get(
            "reemplazos_de_nombres", {}
        )
        self.codigos = self.script_config.get("codigos_planillas", {})

        # Lista completa de planillas [(nombre_mostrado, ViewSchedule)]
        self.planillas = self.get_filtered_planillas(self.doc)
        if not self.planillas:
            forms.alert("No se encontraron planillas configuradas.")
            exit()

        # Copia para restaurar después de filtrar
        self._all_planillas = list(self.planillas)
        self.selected = None

        self.populate(self.planillas)
        self.AcceptButton.Click += self.on_accept
        # El TextChanged se engancha desde XAML con OnFilterTextChanged

    def load_script_json(self):
        script_json_path = os.path.join(
            os.path.expanduser("~"),
            r"AppData\Roaming\MyPyRevitExtention\PyRevitIT.extension\data",
            "script.json",
        )
        with open(script_json_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data

    def get_filtered_planillas(self, doc):
        schedules = (
            FilteredElementCollector(doc)
            .OfClass(ViewSchedule)
            .WhereElementIsNotElementType()
            .ToElements()
        )

        claves = set(k.strip().lower() for k in self.reemplazos.keys())

        return [
            (self.reemplazos.get(sch.Name, sch.Name).strip(), sch)
            for sch in schedules
            if sch.Name and sch.Name.strip().lower() in claves
        ]

    def populate(self, planillas_a_mostrar):
        self.PanelRBs.Children.Clear()
        self.selected = None
        group_name = "planillas_group"

        for name, _ in planillas_a_mostrar:
            rb = RadioButton()
            rb.Content = name
            rb.GroupName = group_name
            rb.Checked += self.on_radio_checked
            self.PanelRBs.Children.Add(rb)

    def on_radio_checked(self, sender, e):
        self.selected = sender

    def OnFilterTextChanged(self, sender, e):
        """Handler llamado desde XAML: TextChanged="OnFilterTextChanged"."""
        texto = (self.FilterTextBox.Text or "").strip().lower()

        if not texto:
            # Sin filtro: mostrar todas
            self.planillas = list(self._all_planillas)
        else:
            self.planillas = [
                (name, sch)
                for (name, sch) in self._all_planillas
                if texto in name.lower()
            ]

        self.populate(self.planillas)

    def on_accept(self, sender, e):
        if not self.selected:
            return

        nombre = self.selected.Content

        # Buscar clave original de la planilla según nombre mostrado
        clave = next(
            (k for k, v in self.reemplazos.items() if v.strip() == nombre),
            None,
        )
        if not clave:
            forms.alert(
                "No se encontró clave para la planilla seleccionada."
            )
            return

        codigo = self.codigos.get(clave)
        if not codigo:
            forms.alert(
                "No se encontró código de planilla en script.json."
            )
            return

        planilla_obj = next(
            (sch for nm, sch in self.planillas if nm == nombre), None
        )
        if not planilla_obj:
            forms.alert(
                "No se encontró el objeto ViewSchedule de la planilla."
            )
            return

        # Encabezados en el orden de la planilla
        schedule_fields = planilla_obj.Definition.GetFieldOrder()
        encabezados = [
            planilla_obj.Definition.GetField(fid).GetName()
            for fid in schedule_fields
        ]

        # Cargar repositorio activo
        repo_datos = load_repo_datos()

        # Recolectar filas desde elementos + aplicar sobreescritura desde repositorio
        link_instances = (
            FilteredElementCollector(self.doc)
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
                p_cod = elem.LookupParameter("CodIntBIM")
                if not p_cod:
                    continue

                p_cod_as_str = p_cod.AsString()
                if not p_cod_as_str:
                    continue

                p_cod_as_str = p_cod_as_str.strip()

                # Condición: código de planilla contenido en CodIntBIM
                if codigo not in p_cod_as_str:
                    continue

                elem_id_str = str(elem.Id.IntegerValue)

                archivo_procesado = linked_doc.PathName
                base, ext = os.path.splitext(archivo_procesado)
                archivo_procesado = re.sub(r"_\d+$", "", base) + ext

                # Clave estándar de repositorio
                clave_repo = u"{}_{}".format(archivo_procesado, elem_id_str)

                # Base: datos desde el modelo
                datos = {
                    "ElementId": elem_id_str,
                    "Archivo": archivo_procesado,
                    "nombre_archivo": os.path.basename(archivo_procesado),
                }

                # Llenar parámetros según encabezados de la planilla
                for p in elem.Parameters:
                    try:
                        pname = p.Definition.Name
                        if pname in encabezados:
                            if p.StorageType.ToString() == "String":
                                pval = p.AsString() or ""
                            else:
                                pval = p.AsValueString() or ""
                            datos[pname] = pval
                    except Exception:
                        pass

                # Sobrescribir con valores existentes en el repositorio activo
                datos_repo = repo_datos.get(clave_repo)
                if datos_repo:
                    for k, v in datos_repo.items():
                        if k not in ("Archivo", "ElementId", "nombre_archivo"):
                            datos[k] = v

                # Actualizar repositorio con la versión final
                repo_datos[clave_repo] = datos

        # Guardar repositorio actualizado
        save_repo_datos(repo_datos)

        # Guardar SOLO headers + código planilla para el visor
        save_repo_planilla_metadata(encabezados, codigo)

        # Lanzar visor CPython, que ahora leerá directamente del repositorio
        launch_cpython_viewer()
        self.Close()


#==================================================
def main():
    selector = SelectorWPF()
    selector.show_dialog()


if __name__ == "__main__":
    main()

#==================================================
#🚫 DELETE BELOW
from Snippets._customprint import kit_button_clicked    # Import Reusable Function from 'lib/Snippets/_customprint.py'
kit_button_clicked(btn_name=__title__)                  # Display Default Print Message
