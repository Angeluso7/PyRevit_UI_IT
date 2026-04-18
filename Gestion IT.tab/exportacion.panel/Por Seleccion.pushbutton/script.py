# -*- coding: utf-8 -*-
__title__ = 'Exportar planillas'
__doc__ = 'Exporta planillas seleccionadas a CSV y genera XLSX formateado mediante un flujo híbrido controlado.'

import os
import sys

button_dir = os.path.dirname(__file__)
extension_root = os.path.abspath(os.path.join(button_dir, '..', '..', '..'))
lib_dir = os.path.join(extension_root, 'lib')
if lib_dir not in sys.path:
    sys.path.append(lib_dir)

from pyrevit import forms
from config.paths import EXPORT_DIR, LOG_DIR, ensure_runtime_dirs
from revit.schedules import get_non_template_schedules
from ui.schedule_selector import select_schedules
from services.export_service import export_selected_schedules
from core.logging_utils import write_log


def main():
    ensure_runtime_dirs()
    uidoc = __revit__.ActiveUIDocument
    if not uidoc:
        forms.alert('No hay un documento activo en Revit.', title='Error')
        write_log(LOG_DIR, 'exportar_planillas.log', 'Fallo: no hay documento activo.')
        return

    doc = uidoc.Document
    schedules = get_non_template_schedules(doc)
    if not schedules:
        forms.alert('No se encontraron planillas disponibles.', title='Aviso')
        write_log(LOG_DIR, 'exportar_planillas.log', 'Sin planillas disponibles.')
        return

    selected = select_schedules(schedules, title='Seleccionar planillas para exportar')
    if not selected:
        forms.alert('Operación cancelada o sin selección.', title='Aviso')
        write_log(LOG_DIR, 'exportar_planillas.log', 'Operación cancelada por el usuario.')
        return

    script_xlsx = os.path.join(extension_root, 'scripts_cpython', 'exportar_csv_a_xlsx.py')
    resultados = export_selected_schedules(selected, EXPORT_DIR, script_xlsx)

    total_csv = len(resultados)
    total_xlsx = len([r for r in resultados if r['xlsx_generado']])
    errores = [u'- {}: {}'.format(r['nombre'], r['error_xlsx']) for r in resultados if r['error_xlsx']]

    mensaje = u'Se exportaron {} archivos CSV.\nSe generaron {} archivos XLSX.\n\nCarpeta de salida:\n{}'.format(total_csv, total_xlsx, EXPORT_DIR)
    if errores:
        mensaje += u'\n\nErrores en XLSX:\n' + u'\n'.join(errores)

    write_log(LOG_DIR, 'exportar_planillas.log', mensaje)
    forms.alert(mensaje, title='Proceso completado')


if __name__ == '__main__':
    main()
