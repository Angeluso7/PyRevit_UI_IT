# -*- coding: utf-8 -*-
import clr
from pyrevit import forms
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import FilteredElementCollector, ParameterFilterElement, ElementParameterFilter, ParameterFilterRuleFactory, Transaction, ElementId

def get_param_id_by_name(doc_local, param_name):
    iterator = doc_local.ParameterBindings.ForwardIterator()
    param_id = None
    while iterator.MoveNext():
        definition = iterator.Key
        if definition.Name == param_name:
            param_id = definition.Id
            break
    return param_id

def activar_filtros_por_nombre(vista, filtro_activo, otros_filtros):
    doc_local = vista.Document
    collector = list(FilteredElementCollector(doc_local).OfClass(ParameterFilterElement))
    filtro_a_activar = None
    otros = []
    for fil in collector:
        if fil.Name == filtro_activo:
            filtro_a_activar = fil
        elif fil.Name in otros_filtros:
            otros.append(fil)
    with Transaction(doc_local, 'Actualizar filtros CodIntBIM Asignado/No') as t:
        t.Start()
        if filtro_a_activar and not vista.IsFilterApplied(filtro_a_activar.Id):
            vista.AddFilter(filtro_a_activar.Id)
        for fil in otros:
            if not vista.IsFilterApplied(fil.Id):
                vista.AddFilter(fil.Id)
        for fid in list(vista.GetFilters()):
            fil = next((f for f in collector if f.Id == fid), None)
            if fil is None:
                continue
            if fil.Name == filtro_activo:
                vista.SetFilterVisibility(fid, True)
                try:
                    vista.SetIsFilterEnabled(fid, True)
                except Exception:
                    pass
            elif fil.Name in otros_filtros:
                vista.SetFilterVisibility(fid, False)
                try:
                    vista.SetIsFilterEnabled(fid, True)
                except Exception:
                    pass
            else:
                vista.SetFilterVisibility(fid, False)
                try:
                    vista.SetIsFilterEnabled(fid, False)
                except Exception:
                    pass
        t.Commit()
    return True

def modificar_filtros_codint(doc_local, nombres_filtros, valor_codint, nombre_parametro='CodIntBIM'):
    filtro_collector = FilteredElementCollector(doc_local).OfClass(ParameterFilterElement)
    filtros_encontrados = [f for f in filtro_collector if f.Name in nombres_filtros]
    if not filtros_encontrados:
        forms.alert("No se encontraron filtros con nombres {}.".format(nombres_filtros), title='Filtros no encontrados')
        return None, None
    param_id = get_param_id_by_name(doc_local, nombre_parametro)
    if param_id is None or param_id == ElementId.InvalidElementId:
        forms.alert("No se encontró parámetro '{}' para la regla.".format(nombre_parametro), title='Parámetro no encontrado')
        return None, None
    filtro_x_id = None
    filtro_y_id = None
    with Transaction(doc_local, 'Modificar reglas filtros f_element_x / f_element_y') as t:
        t.Start()
        for filtro_obj in filtros_encontrados:
            if filtro_obj.Name == 'f_element_x':
                try:
                    regla_nueva = ParameterFilterRuleFactory.CreateEqualsRule(param_id, valor_codint, False)
                except Exception:
                    regla_nueva = ParameterFilterRuleFactory.CreateEqualsRule(param_id, valor_codint)
                filtro_x_id = filtro_obj.Id
            else:
                try:
                    regla_nueva = ParameterFilterRuleFactory.CreateNotEqualsRule(param_id, valor_codint, False)
                except Exception:
                    regla_nueva = ParameterFilterRuleFactory.CreateNotEqualsRule(param_id, valor_codint)
                filtro_y_id = filtro_obj.Id
            filtro_obj.SetElementFilter(ElementParameterFilter(regla_nueva))
        t.Commit()
    return filtro_x_id, filtro_y_id

def activar_filtros_codint_en_vista(doc_local, vista_activa, filtro_x_id, filtro_y_id):
    filtros_aplicados = vista_activa.GetFilters()
    with Transaction(doc_local, 'Actualizar filtros visibilidad/activación CodIntBIM') as t:
        t.Start()
        for filtro_id in filtros_aplicados:
            filtro_obj = None
            for f in FilteredElementCollector(doc_local).OfClass(ParameterFilterElement):
                if f.Id == filtro_id:
                    filtro_obj = f
                    break
            if filtro_obj is None:
                continue
            if filtro_id == filtro_x_id:
                vista_activa.SetFilterVisibility(filtro_id, True)
                try:
                    vista_activa.SetIsFilterEnabled(filtro_id, True)
                except Exception:
                    pass
            elif filtro_id == filtro_y_id:
                vista_activa.SetFilterVisibility(filtro_id, False)
                try:
                    vista_activa.SetIsFilterEnabled(filtro_id, True)
                except Exception:
                    pass
            else:
                vista_activa.SetFilterVisibility(filtro_id, False)
                try:
                    vista_activa.SetIsFilterEnabled(filtro_id, False)
                except Exception:
                    pass
        t.Commit()

def aplicar_filtros_por_codint(vista, valor_codint):
    doc_local = vista.Document
    filtro_x_id, filtro_y_id = modificar_filtros_codint(doc_local, ['f_element_x', 'f_element_y'], valor_codint, nombre_parametro='CodIntBIM')
    if filtro_x_id is None or filtro_y_id is None:
        return False
    with Transaction(doc_local, 'Aplicar filtros f_element_x / f_element_y a la vista') as t:
        t.Start()
        if not vista.IsFilterApplied(filtro_x_id):
            vista.AddFilter(filtro_x_id)
        if not vista.IsFilterApplied(filtro_y_id):
            vista.AddFilter(filtro_y_id)
        t.Commit()
    activar_filtros_codint_en_vista(doc_local, vista, filtro_x_id, filtro_y_id)
    return True
