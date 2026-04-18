# -*- coding: utf-8 -*-
import os
import subprocess
from core.exceptions import ExportError


def ejecutar_exportacion_xlsx(python_exe, script_path, csv_path, timeout=120, json_path=None):
    if not python_exe or not os.path.exists(python_exe):
        raise ExportError('No se encontró CPython en la ruta configurada.')
    if not os.path.exists(script_path):
        raise ExportError('No se encontró el script CPython de exportación XLSX.')
    if not os.path.exists(csv_path):
        raise ExportError('No se encontró el archivo CSV de entrada.')

    cmd = [python_exe, script_path, csv_path]
    if json_path:
        cmd.append(json_path)

    proceso = subprocess.Popen(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        creationflags=0x08000000
    )
    out, err = proceso.communicate()
    if proceso.returncode != 0:
        mensaje = (err or out or 'Error desconocido').decode('utf-8', 'ignore')
        raise ExportError(mensaje)
    return True
