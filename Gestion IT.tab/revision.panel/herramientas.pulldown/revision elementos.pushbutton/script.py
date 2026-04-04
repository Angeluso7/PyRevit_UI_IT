# -*- coding: utf-8 -*-
__title__ = "Revisión Elementos"
__doc__ = """Version = 1.0
Date = 01.09.2024
_______________________________________________________________
Description:

Exporta a Excel (.xlsx) las tablas de planificación BIM
derivadas de los modelos linkeados, agrupando por CodIntBIM
según diccionario 'codigos_planillas' en script.json.

________________________________________________________________
Author: Argenis Angel"""

# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩ ╚═╝╩╚═ ╩ ╚═╝
#==================================================

import os
import sys
import subprocess
import datetime
from pyrevit import forms

#doc = __revit__.ActiveUIDocument.Document
#uidoc = __revit__.ActiveUIDocument

#==================================================

DATA_DIR_EXT = os.path.join(
    os.path.expanduser("~"),
    r"AppData\Roaming\MyPyRevitExtention\PyRevitIT.extension\data"
)

if not os.path.exists(DATA_DIR_EXT):
    forms.alert(
        "No se encontró la carpeta data de la extensión:\n{}".format(DATA_DIR_EXT),
        title="Error"
    )
    raise SystemExit

FORMATEAR_XLSX_V2 = os.path.join(DATA_DIR_EXT, 'formatear_tablas_excel_v2.py')
SCRIPT_JSON_PATH = os.path.join(DATA_DIR_EXT, 'script.json')

PYTHON3_EXE = r"C:\Users\Zbook HP\AppData\Local\Programs\Python\Python313\python.exe"

def seleccionar_carpeta_salida():
    carpeta = forms.pick_folder(
        title='Selecciona la carpeta donde se guardará el .xlsx'
    )
    return carpeta

def nombre_informe_timestamp():
    """Genera nombre dinámico: Inf_Revision_ddmmaa_hhmmss.xlsx"""
    ahora = datetime.datetime.now()
    return "Inf_Revision_{:02d}{:02d}{:02d}_{:02d}{:02d}{:02d}.xlsx".format(
        ahora.day, ahora.month, ahora.year % 100,
        ahora.hour, ahora.minute, ahora.second
    )

def llamar_cpython_export(carpeta_salida):
    """Llama al script CPython que genera el Excel y verifica si realmente se creó."""
    nombre_xlsx = nombre_informe_timestamp()
    ruta_xlsx_salida = os.path.join(carpeta_salida, nombre_xlsx)

    ruta_json_temp = os.path.join(DATA_DIR_EXT, "_temp_datos.json")

    if not os.path.exists(FORMATEAR_XLSX_V2):
        forms.alert(
            "No se encontró formatear_tablas_excel_v2.py en data:\n{}".format(FORMATEAR_XLSX_V2),
            title="Error"
        )
        return

    if not os.path.exists(ruta_json_temp):
        forms.alert(
            "No se encontró el JSON temporal con datos:\n{}\n"
            "Verifica que la extracción desde Revit se haya ejecutado correctamente."
            .format(ruta_json_temp),
            title="Error"
        )
        return

    try:
        subprocess.check_call(
            [PYTHON3_EXE, FORMATEAR_XLSX_V2, ruta_json_temp, ruta_xlsx_salida],
            stderr=subprocess.STDOUT
        )
    except subprocess.CalledProcessError as e:
        forms.alert(
            "Error al generar Excel (proceso CPython):\n{}".format(e),
            title="Error"
        )
        return

    # Verificación final: ¿el archivo realmente existe?
    if os.path.exists(ruta_xlsx_salida):
        forms.alert(
            "Archivo generado:\n{}".format(ruta_xlsx_salida),
            title="Éxito"
        )
    else:
        forms.alert(
            "El proceso terminó sin errores, pero el archivo no se encontró en:\n{}\n\n"
            "Revisa el script formatear_tablas_excel_v2.py; es posible que esté "
            "guardando el Excel en otra ruta o con otro nombre."
            .format(ruta_xlsx_salida),
            title="Advertencia"
        )

def main():
    if not os.path.exists(SCRIPT_JSON_PATH):
        forms.alert(
            "No se encontró script.json en data:\n{}".format(SCRIPT_JSON_PATH),
            title="Error"
        )
        return

    carpeta_salida = seleccionar_carpeta_salida()
    if not carpeta_salida:
        return

    from extraer_modelos_bim import ejecutar_extraccion_y_json
    ok = ejecutar_extraccion_y_json(DATA_DIR_EXT)
    if not ok:
        return

    llamar_cpython_export(carpeta_salida)

if __name__ == '__main__':
    main()

#==================================================
#🚫 DELETE BELOW
from Snippets._customprint import kit_button_clicked
kit_button_clicked(btn_name=__title__)