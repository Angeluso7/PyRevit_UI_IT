# -*- coding: utf-8 -*-
__title__   = "Revisión Elementos"
__doc__     = """Version = 1.0
Date    = 01.09.2024
_______________________________________________________________
Description:

Exporta a Excel (.xlsx) las tablas de planificación BIM
derivadas de los modelos linkeados, agrupando por CodIntBIM
según diccionario 'codigos_planillas' en script.json.

________________________________________________________________
How-To:

1. [Hold ALT + CLICK] on the button to open its source folder.
You will be able to override this placeholder.

2. Automate Your Boring Work ;)

________________________________________________________________
TODO:
[FEATURE] - Describe Your ToDo Tasks Here
________________________________________________________________
Last Updates:
- [01.09.2025] v0.1 Inicio de Aplicación. 
________________________________________________________________
Author: Argenis Angel"""

# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝
#==================================================
# Importaciones base para pyRevit script con interfaz WPF y Revit API

import os
import sys
import subprocess
from pyrevit import forms


# ╦  ╦╔═╗╦═╗╦╔═╗╔╗ ╦  ╔═╗╔═╗
# ╚╗╔╝╠═╣╠╦╝║╠═╣╠╩╗║  ║╣ ╚═╗
#  ╚╝ ╩ ╩╩╚═╩╩ ╩╚═╝╩═╝╚═╝╚═╝
#==================================================

#doc = __revit__.ActiveUIDocument.Document
#uidoc = __revit__.ActiveUIDocument

# ╔╦╗╔═╗╦╔╗╔
# ║║║╠═╣║║║║
# ╩ ╩╩ ╩╩╝╚╝
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

# Script CPython que hace el formateo Excel
FORMATEAR_XLSX_V2 = os.path.join(DATA_DIR_EXT, 'formatear_tablas_excel_v2.py')
SCRIPT_JSON_PATH = os.path.join(DATA_DIR_EXT, 'script.json')

# Ruta de Python 3 (ajustar a tu instalación)
PYTHON3_EXE = r"C:\Users\Zbook HP\AppData\Local\Programs\Python\Python313\python.exe"


def seleccionar_carpeta_salida():
    carpeta = forms.pick_folder(
        title='Selecciona la carpeta donde se guardará el .xlsx'
    )
    return carpeta


def llamar_cpython_export(carpeta_salida):
    """Llama a un pequeño script CPython que:
       - Lee el modelo (vía JSON generado por otro script)
       - Genera un .xlsx en carpeta_salida
    """
    # Nombre de archivo Excel de salida
    nombre_xlsx = "Exportacion_Tablas_BIM.xlsx"
    ruta_xlsx_salida = os.path.join(carpeta_salida, nombre_xlsx)

    # JSON temporal lo generará el script llamado desde Revit
    # En este flujo, asumimos que ya se generará dentro de DATA_DIR_EXT
    ruta_json_temp = os.path.join(DATA_DIR_EXT, "_temp_datos.json")

    if not os.path.exists(FORMATEAR_XLSX_V2):
        forms.alert(
            "No se encontró formatear_tablas_excel_v2.py en data:\n{}".format(FORMATEAR_XLSX_V2),
            title="Error"
        )
        return

    try:
        # Llamada directa al formatear_tablas_excel_v2.py con JSON ya preparado
        subprocess.check_call(
            [PYTHON3_EXE, FORMATEAR_XLSX_V2, ruta_json_temp, ruta_xlsx_salida],
            stderr=subprocess.STDOUT
        )
        forms.alert(
            "Archivo generado:\n{}".format(ruta_xlsx_salida),
            title="Éxito"
        )
    except subprocess.CalledProcessError as e:
        forms.alert(
            "Error al generar Excel:\n{}".format(e),
            title="Error"
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

    # Aquí va la lógica de extracción de datos desde Revit y
    # generación del JSON temporal _temp_datos.json usando la API.
    # Lo haremos en el mismo contexto de Revit más abajo (ver segundo script).
    from extraer_modelos_bim import ejecutar_extraccion_y_json
    ok = ejecutar_extraccion_y_json(DATA_DIR_EXT)

    if not ok:
        return

    llamar_cpython_export(carpeta_salida)


if __name__ == '__main__':
    main()


#==================================================
#🚫 DELETE BELOW
from Snippets._customprint import kit_button_clicked    # Import Reusable Function from 'lib/Snippets/_customprint.py'
kit_button_clicked(btn_name=__title__)                  # Display Default Print Message
