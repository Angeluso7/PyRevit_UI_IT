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
)

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

valor_codint = ARGS[0] if "ARGS" in globals() and ARGS else ""


def ajustar_reglas_filtros_codintbim(doc_local, valor_cod):
    if not valor_cod:
        forms.alert("No se recibió un valor de CodIntBIM válido.", title="Error")
        return

    col_pf = FilteredElementCollector(doc_local).OfClass(ParameterFilterElement)
    f_x = None
    f_y = None
    for f in col_pf:
        if f.Name == "f_element_x":
            f_x = f
        elif f.Name == "f_element_y":
            f_y = f

    if f_x is None or f_y is None:
        forms.alert(
            "No se encontraron ambos filtros 'f_element_x' y 'f_element_y' en el modelo.\n"
            "Crea esos filtros manualmente antes de usar la herramienta.",
            title="Filtros no encontrados",
        )
        return

    el_muestra = None
    vista = doc_local.ActiveView
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

    rule_eq = ParameterFilterRuleFactory.CreateEqualsRule(param_id, valor_cod, False)
    rule_neq = ParameterFilterRuleFactory.CreateNotEqualsRule(param_id, valor_cod, False)

    filt_eq = ElementParameterFilter(rule_eq)
    filt_neq = ElementParameterFilter(rule_neq)

    with Transaction(doc_local, "Actualizar reglas f_element_x / f_element_y") as t:
        t.Start()
        f_x.SetElementFilter(filt_eq)
        f_y.SetElementFilter(filt_neq)
        t.Commit()


def main():
    ajustar_reglas_filtros_codintbim(doc, valor_codint)


if __name__ == "__main__":
    main()
