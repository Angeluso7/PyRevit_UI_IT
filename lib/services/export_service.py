# -*- coding: utf-8 -*-
from revit.schedules import export_schedule_to_csv
from integration.cpython_excel import ejecutar_exportacion_xlsx
from config.settings import CPYTHON_EXE, ENABLE_CPYTHON_XLSX, CPYTHON_TIMEOUT, SCRIPT_JSON_PATH
from config.paths import SCRIPT_XLSX_PATH
from core.exceptions import ExportError
import os


def export_selected_schedules(schedules, export_folder):
    """
    Exporta cada planilla a CSV y genera el XLSX con formato vertical.
    La ruta del script CPython se resuelve centralmente desde config.paths.
    """
    resultados = []
    for schedule in schedules:
        csv_path      = export_schedule_to_csv(schedule, export_folder)
        xlsx_generado = False
        error_xlsx    = None

        if ENABLE_CPYTHON_XLSX:
            try:
                ejecutar_exportacion_xlsx(
                    CPYTHON_EXE,
                    SCRIPT_XLSX_PATH,
                    csv_path,
                    CPYTHON_TIMEOUT,
                    SCRIPT_JSON_PATH
                )
                xlsx_generado = True
            except ExportError as exc:
                error_xlsx = unicode(exc)

        resultados.append({
            'nombre':        schedule.Name,
            'csv':           csv_path,
            'xlsx':          os.path.splitext(csv_path)[0] + '.xlsx',
            'xlsx_generado': xlsx_generado,
            'error_xlsx':    error_xlsx,
        })
    return resultados
