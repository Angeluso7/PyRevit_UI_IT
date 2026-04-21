# -*- coding: utf-8 -*-
__title__ = "Actualizacion de datos"
__doc__ = """Lanza el actualizador batch mediante ExternalEvent del add-in RevitBatchUpdater."""

import clr
import traceback
import os
from datetime import datetime


def get_ext_root():
    return os.path.join(
        os.environ.get('APPDATA', ''),
        'MyPyRevitExtention',
        'PyRevitIT.extension'
    )


def ensure_dir(path):
    if not os.path.isdir(path):
        try:
            os.makedirs(path)
        except Exception:
            pass
    return path


def get_data_dir():
    return ensure_dir(os.path.join(get_ext_root(), 'data'))


def get_logs_dir():
    return ensure_dir(os.path.join(get_data_dir(), 'logs'))


def log(msg):
    try:
        log_path = os.path.join(get_logs_dir(), 'batchupdater_addin.log')
        with open(log_path, 'a', encoding='utf-8') as f:
            f.write('[script.py] {0} | {1}\n'.format(
                datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                msg
            ))
    except Exception:
        pass


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
    raise