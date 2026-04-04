# -*- coding: utf-8 -*-
__title__   = "Por Planilla IT"
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
clr.AddReference("RevitAPI")
clr.AddReference("RevitServices")
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
import System

from Autodesk.Revit.DB import (
    FilteredElementCollector, ViewSchedule, ParameterFilterElement,
    ParameterFilterRuleFactory, ElementId, Transaction, BuiltInParameter,
    ElementParameterFilter, LogicalAndFilter, ElementFilter,
    OverrideGraphicSettings
)
from System.Collections.Generic import List
from System.Windows.Forms import (
    Form, ComboBox, Button, Label, MessageBox, FormStartPosition,
    ComboBoxStyle, DialogResult
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

# Nueva ruta en carpeta data
DATA_DIR = os.path.join(
    os.path.expanduser("~"),
    r"AppData\\Roaming\\MyPyRevitExtention\\PyRevitIT.extension\\data"
)

if not os.path.exists(DATA_DIR):
    os.makedirs(DATA_DIR)

archivo_json = os.path.join(DATA_DIR, "script.json")

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

class PlanillaSelectorForm(Form):
    def __init__(self, nombres_planilla):
        self.Text = "Selecciona una planilla"
        self.Size = Size(400, 150)
        self.StartPosition = FormStartPosition.CenterScreen

        self.label = Label(Text="Planilla:", Location=Point(10, 20), Size=Size(380, 20))
        self.Controls.Add(self.label)

        self.combo = ComboBox(Location=Point(10, 50), Size=Size(360, 20))
        self.combo.DropDownStyle = ComboBoxStyle.DropDownList

        items_array = System.Array[object](nombres_planilla)
        self.combo.Items.AddRange(items_array)
        if nombres_planilla:
            self.combo.SelectedIndex = 0

        self.Controls.Add(self.combo)

        self.btn_ok = Button(Text="Aceptar", Location=Point(220, 80), Size=Size(75, 30))
        self.btn_ok.Click += self.on_accept
        self.Controls.Add(self.btn_ok)

        self.btn_cancel = Button(Text="Cancelar", Location=Point(300, 80), Size=Size(75, 30))
        self.btn_cancel.Click += self.on_cancel
        self.Controls.Add(self.btn_cancel)

        self.selected_planilla = None

    def on_accept(self, sender, event):
        if self.combo.SelectedIndex == -1:
            MessageBox.Show("Por favor seleccione una planilla", "Aviso")
            return
        self.selected_planilla = self.combo.SelectedItem
        self.DialogResult = DialogResult.OK
        self.Close()

    def on_cancel(self, sender, event):
        self.DialogResult = DialogResult.Cancel
        self.Close()

def modificar_filtros(doc, nombres_filtros, valor_parametro, nombre_parametro="CodIntBIM"):
    filtro_collector = FilteredElementCollector(doc).OfClass(ParameterFilterElement)
    filtros_encontrados = [f for f in filtro_collector if f.Name in nombres_filtros]

    if not filtros_encontrados:
        MessageBox.Show(f"No se encontraron filtros con nombres {nombres_filtros}.")
        return None

    param_id = get_param_id_by_name(doc, nombre_parametro)
    if param_id is None or param_id == ElementId.InvalidElementId:
        MessageBox.Show(f"No se encontró parámetro '{nombre_parametro}' para la regla.")
        return None

    ids_modificados = []
    with Transaction(doc, "Modificar reglas de filtros") as t:
        t.Start()
        for filtro_obj in filtros_encontrados:
            if filtro_obj.Name == "b_x_planilla":
                regla_nueva = ParameterFilterRuleFactory.CreateContainsRule(param_id, valor_parametro, False)
            else:  # Para b_x_planilla_x usar operador "no contiene"
                regla_nueva = ParameterFilterRuleFactory.CreateNotContainsRule(param_id, valor_parametro, False)
            filtro_nuevo = ElementParameterFilter(regla_nueva)
            filtro_obj.SetElementFilter(filtro_nuevo)
            ids_modificados.append(filtro_obj.Id)
        t.Commit()

    return ids_modificados

def activar_filtro_unico(doc, vista_activa, filtro_activar_id):
    filtros_aplicados = vista_activa.GetFilters()
    with Transaction(doc, "Actualizar filtros visibilidad y activacion") as t:
        t.Start()
        for filtro_id in filtros_aplicados:
            filtro_obj = None
            for f in FilteredElementCollector(doc).OfClass(ParameterFilterElement):
                if f.Id == filtro_id:
                    filtro_obj = f
                    break
            if filtro_obj.Name != "b_x_planilla" and filtro_obj.Name != "b_x_planilla_x":
                vista_activa.SetFilterVisibility(filtro_id, False)
                vista_activa.SetIsFilterEnabled(filtro_id, False)
                vista_activa.SetFilterOverrides(filtro_id, OverrideGraphicSettings())
            elif filtro_obj.Name == "b_x_planilla":
                vista_activa.SetFilterVisibility(filtro_id, True)
                vista_activa.SetIsFilterEnabled(filtro_id, True)
                override_settings = OverrideGraphicSettings()
                vista_activa.SetFilterOverrides(filtro_id, override_settings)
            elif filtro_obj.Name == "b_x_planilla_x":
                vista_activa.SetFilterVisibility(filtro_id, False)
        t.Commit()

def main():
    planillas_doc = obtener_planillas_desde_documento(doc)
    planillas_json = obtener_claves_json(archivo_json)
    nombres_combinados = sorted(set(planillas_doc) | set(planillas_json))

    if not nombres_combinados:
        MessageBox.Show("No se encontraron planillas.", "Aviso")
        return

    form = PlanillaSelectorForm(nombres_combinados)
    if form.ShowDialog() != DialogResult.OK:
        MessageBox.Show("Operación cancelada.", "Aviso")
        return

    seleccion = form.selected_planilla

    with open(archivo_json, "r", encoding="utf-8") as f:
        data = json.load(f)
    codigos_planillas = data.get("codigos_planillas", {})

    if seleccion not in codigos_planillas:
        MessageBox.Show(f"No se encontró código para la planilla '{seleccion}' en el JSON.")
        return

    valor_regla = codigos_planillas[seleccion]

    filtro_ids = modificar_filtros(doc, ["b_x_planilla", "b_x_planilla_x"], valor_regla)
    if filtro_ids is None:
        return

    vista_activa = doc.ActiveView
    activar_filtro_unico(doc, vista_activa, filtro_ids[0])  # Asumiendo b_x_planilla primer filtro

    MessageBox.Show(f"Filtros 'b_x_planilla' y 'b_x_planilla_x' actualizados y aplicados para planilla '{seleccion}'.")

if __name__ == "__main__":
    main()


#==================================================
#🚫 DELETE BELOW
from Snippets._customprint import kit_button_clicked    # Import Reusable Function from 'lib/Snippets/_customprint.py'
kit_button_clicked(btn_name=__title__)                  # Display Default Print Message
