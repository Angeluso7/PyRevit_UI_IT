# -*- coding: utf-8 -*-
import os
import subprocess
from core.exceptions import ExportError

_NO_WINDOW = 0x08000000


def ejecutar_exportacion_xlsx(python_exe, script_path, csv_path, timeout=120, json_path=None):
    """
    Lanza el script CPython que genera el XLSX con formato vertical.
    Valida rutas, usa timeout real y reporta errores detallados.
    """
    if not python_exe or not os.path.isfile(python_exe):
        raise ExportError(
            u'CPython no encontrado: {}\n'
            u'Define la variable de entorno PYREVIT_IT_CPYTHON.'.format(python_exe)
        )
    if not os.path.isfile(script_path):
        raise ExportError(u'Script XLSX no encontrado: {}'.format(script_path))
    if not os.path.isfile(csv_path):
        raise ExportError(u'CSV de entrada no encontrado: {}'.format(csv_path))

    cmd = [python_exe, script_path, csv_path]
    if json_path and os.path.isfile(json_path):
        cmd.append(json_path)

    try:
        proceso = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            creationflags=_NO_WINDOW
        )
        out, err = proceso.communicate(timeout=timeout)
    except subprocess.TimeoutExpired:
        proceso.kill()
        raise ExportError(u'Proceso CPython excedio el limite de {} s.'.format(timeout))
    except OSError as exc:
        raise ExportError(u'Error al lanzar CPython: {}'.format(exc))

    if proceso.returncode != 0:
        msg = (err or out or b'Error desconocido').decode('utf-8', 'ignore')
        raise ExportError(msg)
    return True
