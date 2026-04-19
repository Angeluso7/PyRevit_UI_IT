# -*- coding: utf-8 -*-
__title__ = 'Por Código Interno'
__doc__ = 'Selecciona elementos del modelo por CodIntBIM o por estado Asignado / No Asignado usando filtros.'
import os
import time
import clr
from pyrevit import forms
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import View3D
from services.codint_service import lanzar_selector, leer_salida_selector, aplicar_opcion

doc = __revit__.ActiveUIDocument.Document

def main():
    if doc is None:
        forms.alert('No hay documento activo.', title='Error')
        return
    if not isinstance(doc.ActiveView, View3D):
        forms.alert('La vista activa debe ser una vista 3D.', title='Aviso')
        return
    try:
        bundle_dir = os.path.dirname(os.path.abspath(__file__))
    except Exception:
        bundle_dir = os.getcwd()
    selector_script = os.path.join(bundle_dir, 'codint_selector.py')
    if not lanzar_selector(doc, bundle_dir, selector_script):
        return
    for _ in range(300):
        salida = leer_salida_selector()
        if salida:
            break
        time.sleep(0.1)
    else:
        return
    if salida.get('opcion', '') == 'cancelar':
        return
    aplicar_opcion(doc, doc.ActiveView, salida)

if __name__ == '__main__':
    main()
