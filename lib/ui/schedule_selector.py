# -*- coding: utf-8 -*-
from pyrevit import forms


def select_schedules(schedules, title='Seleccionar planillas'):
    names = sorted([s.Name for s in schedules])
    selected = forms.SelectFromList.show(
        names,
        title=title,
        multiselect=True,
        button_name='Exportar selección'
    )
    if not selected:
        return []
    selected_set = set(selected)
    return [s for s in schedules if s.Name in selected_set]
