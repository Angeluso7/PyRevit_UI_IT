# -*- coding: utf-8 -*-
__title__ = "Actualizacion de datos"
__doc__ = """Version = 2.1
Date    = 22.04.2026
________________________________________________________________
Description:

Lanzador pyRevit del addin RevitBatchUpdater.

Cambios v2.1:
- Logging en data/logs
- Sin rutas rígidas innecesarias para la BD
- El addin resuelve la configuración y el repo activo
________________________________________________________________
Author: Angeluso
"""

import os
import clr
import traceback
from datetime import datetime

from pyrevit import forms


def get_ext_root_data_dir():
    appdata = os.environ.get('APPDATA', '')
    return os.path.join(
        appdata,
        'MyPyRevitExtention',
        'PyRevitIT.extension',
        'data'
    )


def get_logs_dir():
    logs_dir = os.path.join(get_ext_root_data_dir(), 'logs')
    if not os.path.isdir(logs_dir):
        try:
            os.makedirs(logs_dir)
        except Exception:
            pass
    return logs_dir


def get_log_path():
    return os.path.join(get_logs_dir(), 'batchupdater_addin.log')


def log(msg):
    try:
        with open(get_log_path(), 'a') as f:
            f.write('[script.py] {0} | {1}\n'.format(
                datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                msg
            ))
    except Exception:
        pass


def main():
    log('script.py iniciado.')

    try:
        log('Intentando clr.AddReference("RevitBatchUpdater")...')
        clr.AddReference('RevitBatchUpdater')
        log('Referencia a RevitBatchUpdater obtenida correctamente.')

        from RevitBatchUpdater import AddinEntry
        log('Import AddinEntry OK.')

        log('Llamando AddinEntry.LanzarActualizador()...')
        AddinEntry.LanzarActualizador()
        log('AddinEntry.LanzarActualizador() ejecutado sin excepción.')

    except Exception as e:
        log('EXCEPCION en script.py: {0}'.format(e))
        log('Traceback: {0}'.format(traceback.format_exc()))
        forms.alert(
            "No se pudo lanzar el actualizador.\n\n"
            "Verifica que el addin 'RevitBatchUpdater' esté compilado, "
            "copiado y cargado por Revit.\n\n"
            "Revisa también el log:\n{}".format(get_log_path()),
            title="Actualizacion de datos - Error"
        )
        raise


if __name__ == '__main__':
    main()