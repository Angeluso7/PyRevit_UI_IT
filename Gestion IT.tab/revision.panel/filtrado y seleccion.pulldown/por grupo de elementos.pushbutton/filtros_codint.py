# -*- coding: utf-8 -*-

import clr
from pyrevit import forms

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    ParameterFilterElement,
    ElementParameterFilter,
    ParameterFilterRuleFactory,
    Transaction,
    View3D,
    OverrideGraphicSettings,
)

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# valor pasado como argumento (clave de codigos_elementos)
valor_codint = ARGS[0] if "ARGS" in globals() and ARGS else ""


def ajustar_reglas_filtros_codintbim(doc_local, valor_cod):
    if not valor_cod:
        forms.alert("No se recibió un valor de CodIntBIM válido.", title="Error")
        return

    col_pf = FilteredElementCollector(doc_local).OfClass(ParameterFilterElement)

    f_x = None
    f_y = None
    filtros_dict = {}
    for f in col_pf:
        filtros_dict[f.Name] = f
        if f.Name == "f_element_x":
            f_x = f
        elif f.Name == "f_element_y":
            f_y = f

    if f_x is None or f_y is None:
        forms.alert(
            "No se encontraron ambos filtros 'f_element_x' y 'f_element_y' en el modelo.",
            title="Filtros no encontrados",
        )
        return

    # elemento de muestra con CodIntBIM en la vista activa
    vista = doc_local.ActiveView
    el_muestra = None
    for el in FilteredElementCollector(doc_local, vista.Id).WhereElementIsNotElementType():
        p = el.LookupParameter("CodIntBIM")
        if p:
            el_muestra = el
            break
    if el_muestra is None:
        forms.alert(
            "No se encontró ningún elemento en la vista con el parámetro 'CodIntBIM'.",
            title="Parámetro no encontrado",
        )
        return

    p_cod = el_muestra.LookupParameter("CodIntBIM")
    param_id = p_cod.Id

    # reglas contiene / no contiene
    rule_contains = ParameterFilterRuleFactory.CreateContainsRule(param_id, valor_cod, False)
    rule_not_contains = ParameterFilterRuleFactory.CreateNotContainsRule(param_id, valor_cod, False)

    filt_contains = ElementParameterFilter(rule_contains)
    filt_not_contains = ElementParameterFilter(rule_not_contains)

    with Transaction(doc_local, "Actualizar filtros CodIntBIM") as t:
        t.Start()

        # asignar reglas a f_element_x / f_element_y
        f_x.SetElementFilter(filt_contains)
        f_y.SetElementFilter(filt_not_contains)

        # activar/desactivar filtros y visibilidad en la vista activa
        filtros_en_vista = set(vista.GetFilters())

        for f in filtros_dict.values():
            fid = f.Id

            if f.Name == "f_element_x":
                # activar filtro y visibilidad
                if fid not in filtros_en_vista:
                    vista.AddFilter(fid)
                vista.SetIsFilterEnabled(fid, True)
                vista.SetFilterVisibility(fid, True)

            elif f.Name == "f_element_y":
                # activar filtro, visibilidad apagada
                if fid not in filtros_en_vista:
                    vista.AddFilter(fid)
                vista.SetIsFilterEnabled(fid, True)
                vista.SetFilterVisibility(fid, False)

            else:
                # todos los demás filtros: desactivar y ocultar
                if fid in filtros_en_vista:
                    vista.SetIsFilterEnabled(fid, False)
                    vista.SetFilterVisibility(fid, False)

            # opcionalmente, limpiar overrides
            vista.SetFilterOverrides(fid, OverrideGraphicSettings())

        t.Commit()

    forms.alert(
        "Filtros 'f_element_x' y 'f_element_y' actualizados correctamente\n"
        "y estados de activación/visibilidad ajustados.",
        title="Actualización correcta"
    )


def main():
    ajustar_reglas_filtros_codintbim(doc, valor_codint)


if __name__ == "__main__":
    main()
