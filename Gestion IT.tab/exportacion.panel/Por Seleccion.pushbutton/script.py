# -*- coding: utf-8 -*-
__title__ = 'Por Seleccion'
__doc__   = ('Exporta planillas seleccionadas a CSV y genera XLSX con '
             'formato vertical y colores por parametro.')

import os
import sys

_button_dir = os.path.dirname(__file__)
_ext_root   = os.path.abspath(os.path.join(_button_dir, '..', '..', '..'))
_lib_dir    = os.path.join(_ext_root, 'lib')
if _lib_dir not in sys.path:
    sys.path.insert(0, _lib_dir)

from pyrevit import forms
from config.paths           import EXPORT_DIR, LOG_DIR, ensure_runtime_dirs
from revit.schedules        import get_non_template_schedules
from ui.schedule_selector   import select_schedules
from services.export_service import export_selected_schedules
from core.logging_utils     import write_log


def main():
    ensure_runtime_dirs()

    uidoc = __revit__.ActiveUIDocument
    if not uidoc:
        forms.alert(u'No hay un documento activo en Revit.', title=u'Error')
        write_log(LOG_DIR, 'exportar_planillas.log', u'Fallo: no hay documento activo.')
        return

    doc = uidoc.Document
    schedules = get_non_template_schedules(doc)
    if not schedules:
        forms.alert(u'No se encontraron planillas disponibles.', title=u'Aviso')
        write_log(LOG_DIR, 'exportar_planillas.log', u'Sin planillas disponibles.')
        return

    selected = select_schedules(schedules, title=u'Seleccionar planillas para exportar')
    if not selected:
        forms.alert(u'Operacion cancelada o sin seleccion.', title=u'Aviso')
        write_log(LOG_DIR, 'exportar_planillas.log', u'Cancelado por el usuario.')
        return

    resultados = export_selected_schedules(selected, EXPORT_DIR)

    total_csv  = len(resultados)
    total_xlsx = sum(1 for r in resultados if r['xlsx_generado'])
    errores    = [u'- {}: {}'.format(r['nombre'], r['error_xlsx'])
                  for r in resultados if r['error_xlsx']]

    msg = (u'Exportados : {} archivos CSV\n'
           u'Generados  : {} archivos XLSX\n\n'
           u'Carpeta de salida:\n{}').format(total_csv, total_xlsx, EXPORT_DIR)
    if errores:
        msg += u'\n\nErrores en XLSX:\n' + u'\n'.join(errores)

    write_log(LOG_DIR, 'exportar_planillas.log', msg)
    forms.alert(msg, title=u'Proceso completado')


if __name__ == '__main__':
    main()
