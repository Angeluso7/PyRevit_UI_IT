# -*- coding: utf-8 -*-

__title__ = "CodIntBIM Asignado / No Asignado"
__doc__ = "Activa filtros c_cod_int / s_cod_int según selección Asignados / No Asignados."

import wpf
import json
import os

from System.Windows import Window
from pyrevit import script
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    ParameterFilterElement,
    Transaction,
)

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument


def activar_filtros_por_nombre(vista, filtro_activo, otros_filtros):
    filtro_a_activar = None
    otros = []

    collector = FilteredElementCollector(vista.Document).OfClass(ParameterFilterElement)
    for fil in collector:
        if fil.Name == filtro_activo:
            filtro_a_activar = fil
        elif fil.Name in otros_filtros:
            otros.append(fil)

    with Transaction(vista.Document, "Actualizar filtros visibilidad") as t:
        t.Start()

        if filtro_a_activar and not vista.IsFilterApplied(filtro_a_activar.Id):
            vista.AddFilter(filtro_a_activar.Id)
        if filtro_a_activar:
            vista.SetFilterVisibility(filtro_a_activar.Id, True)

        for fil in otros:
            if vista.IsFilterApplied(fil.Id):
                vista.SetFilterVisibility(fil.Id, False)

        t.Commit()


class VentanaInicial(Window):
    def __init__(self, vista_activa):
        self.vista_activa = vista_activa

        xaml_path = script.get_bundle_file("seleccion.xaml")
        wpf.LoadComponent(self, xaml_path)

        self.btnAsignados.Click += self.btnAsignados_Click
        self.btnNoAsignados.Click += self.btnNoAsignados_Click

    def btnAsignados_Click(self, sender, e):
        # elementos con CodIntBIM asignado -> filtro c_cod_int
        activar_filtros_por_nombre(self.vista_activa, "c_cod_int", ["s_cod_int"])
        self.exportar_configuracion()
        self.Close()

    def btnNoAsignados_Click(self, sender, e):
        # elementos sin CodIntBIM -> filtro s_cod_int
        activar_filtros_por_nombre(self.vista_activa, "s_cod_int", ["c_cod_int"])
        self.exportar_configuracion()
        self.Close()

    def exportar_configuracion(self):
        config = {
            "Title": self.Title,
            "Width": self.Width,
            "Height": self.Height,
            "Left": self.Left,
            "Top": self.Top,
            "ResizeMode": str(self.ResizeMode),
            "Buttons": []
        }

        grid = self.Content
        stackpanel = None
        for child in grid.Children:
            if hasattr(child, "Children"):
                stackpanel = child
                break

        if stackpanel:
            for c in stackpanel.Children:
                if hasattr(c, "Content") and hasattr(c, "Name"):
                    config["Buttons"].append({
                        "Name": c.Name,
                        "Content": c.Content,
                        "Width": c.Width,
                        "Height": c.Height,
                        "Margin": str(c.Margin)
                    })

        carpeta_salida = (
            r"C:\Users\Zbook HP\AppData\Roaming\MyPyRevitExtention\PyRevitIT.extension"
            r"\Gestion IT.tab\Menu.panel\Seleccion.pulldown\Reestablecer Vista.pushbutton"
        )
        if not os.path.exists(carpeta_salida):
            os.makedirs(carpeta_salida)

        archivo_json = os.path.join(carpeta_salida, "config_ventana.json")
        with open(archivo_json, "w") as f:
            json.dump(config, f, indent=4)

        script.get_logger().info("Configuración exportada a {0}".format(archivo_json))


def main():
    vista_activa = doc.ActiveView
    ventana = VentanaInicial(vista_activa)
    ventana.ShowDialog()


if __name__ == "__main__":
    main()
