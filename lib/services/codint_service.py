# -*- coding: utf-8 -*-
import os
import json
import subprocess
from pyrevit import forms
import clr
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import FilteredElementCollector, RevitLinkInstance
from revit.filter_utils import activar_filtros_por_nombre, aplicar_filtros_por_codint
from config.paths import TEMP_DIR, ensure_runtime_dirs
from config.settings import CPYTHON_EXE

# Archivos temporales en data/temp/ (resueltos desde config.paths)
ensure_runtime_dirs()
TMP_IN  = os.path.join(TEMP_DIR, 'codint_selector_in.json')
TMP_OUT = os.path.join(TEMP_DIR, 'codint_selector_out.json')


def recoger_codintbim(doc):
    elementos = []
    for el in FilteredElementCollector(doc).WhereElementIsNotElementType():
        try:
            p = el.LookupParameter('CodIntBIM')
            if not p or not p.HasValue:
                continue
            cod = (p.AsString() or '').strip()
            if not cod:
                continue
            elementos.append({
                'doc_path':   doc.PathName or '',
                'element_id': el.Id.IntegerValue,
                'codintbim':  cod,
            })
        except Exception:
            continue
    try:
        for li in FilteredElementCollector(doc).OfClass(RevitLinkInstance):
            ldoc = li.GetLinkDocument()
            if not ldoc:
                continue
            ruta = ldoc.PathName or ''
            for el in FilteredElementCollector(ldoc).WhereElementIsNotElementType():
                try:
                    p = el.LookupParameter('CodIntBIM')
                    if not p or not p.HasValue:
                        continue
                    cod = (p.AsString() or '').strip()
                    if not cod:
                        continue
                    elementos.append({
                        'doc_path':   ruta,
                        'element_id': el.Id.IntegerValue,
                        'codintbim':  cod,
                    })
                except Exception:
                    continue
    except Exception:
        pass
    return {'elementos': elementos}


def lanzar_selector(doc, bundle_dir, selector_script):
    if not os.path.exists(CPYTHON_EXE):
        forms.alert(
            'No se encontró el ejecutable de Python 3:\n{}\n'
            'Ajusta CPYTHON_EXE en settings.py o define la variable de entorno '
            'PYREVIT_IT_CPYTHON.'.format(CPYTHON_EXE),
            title='Python no encontrado',
        )
        return False
    if not os.path.exists(selector_script):
        forms.alert(
            'No se encontró el selector CPython:\n{}'.format(selector_script),
            title='Script CPython no encontrado',
        )
        return False

    datos = recoger_codintbim(doc)

    try:
        with open(TMP_IN, 'w', encoding='utf-8') as f:
            json.dump(datos, f, ensure_ascii=False, indent=2)
    except Exception as e:
        forms.alert(
            'Error escribiendo archivo de entrada para selector:\n{}'.format(e),
            title='Error IO',
        )
        return False

    if os.path.exists(TMP_OUT):
        try:
            os.remove(TMP_OUT)
        except Exception:
            pass

    try:
        subprocess.Popen([CPYTHON_EXE, selector_script, TMP_IN, TMP_OUT], cwd=bundle_dir)
        return True
    except Exception as e:
        forms.alert(
            'Error al ejecutar CPython:\n{}'.format(e),
            title='Error CPython',
        )
        return False


def leer_salida_selector():
    if not os.path.exists(TMP_OUT):
        return None
    try:
        with open(TMP_OUT, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return None


def aplicar_opcion(doc, vista, salida):
    opcion = salida.get('opcion', '')
    if opcion == 'by_codint':
        cod = salida.get('codintbim', '')
        if not cod:
            forms.alert('No se recibió un CodIntBIM válido.', title='Dato faltante')
            return
        aplicar_filtros_por_codint(vista, cod)
    elif opcion == 'asignados':
        activar_filtros_por_nombre(vista, 'c_cod_int', ['s_cod_int'])
    elif opcion == 'no_asignados':
        activar_filtros_por_nombre(vista, 's_cod_int', ['c_cod_int'])
