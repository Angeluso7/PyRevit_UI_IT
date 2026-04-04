# -*- coding: utf-8 -*-
__title__   = "Planilla vs Modelo"
__doc__     = """Version = 1.0
Date    = 01.09.2024
_______________________________________________________________
Description:

Comparación de Planilla con datos y elementos de modelo en función
de la información existente y la faltante.

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
import datetime
from pyrevit import forms

from datetime import datetime  # o al inicio del archivo

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

# ---------------------------------------------------------
# Rutas base
# ---------------------------------------------------------

DATA_DIR_EXT = os.path.join(
    os.path.expanduser("~"),
    r"AppData\\Roaming\\MyPyRevitExtention\\PyRevitIT.extension\\data"
)

if not os.path.exists(DATA_DIR_EXT):
    forms.alert(
        "No se encontró la carpeta data de la extensión:\n{}".format(DATA_DIR_EXT),
        title="Error"
    )
    raise SystemExit

APPDATA_DIR = os.getenv('APPDATA') or os.path.expanduser('~')
DATA_DIR = os.path.join(APPDATA_DIR, 'PyRevitIT', 'data', 'comparacion')
if not os.path.exists(DATA_DIR):
    os.makedirs(DATA_DIR)

PYTHON3_EXE = r"C:\Users\Zbook HP\AppData\Local\Programs\Python\Python313\python.exe"

LEER_XLSM = os.path.join(DATA_DIR_EXT, 'leer_xlsm_codigos.py')
FORMATEAR_XLSX = os.path.join(DATA_DIR_EXT, 'formatear_tablas_excel.py')
UI_COMPARACION = os.path.join(DATA_DIR_EXT, 'ui_comparacion.py')
SCRIPT_JSON_PATH = os.path.join(DATA_DIR_EXT, 'script.json')

# JSON generados desde Revit
MODELO_JSON = os.path.join(DATA_DIR_EXT, 'modelo_codint_por_cm.json')
HEADERS_JSON = os.path.join(DATA_DIR_EXT, 'headers_por_tabla.json')  # ya no se usa en la UI nueva

CREATE_NO_WINDOW = 0x08000000


def seleccionar_xlsm():
    return forms.pick_file(
        file_ext='xlsm',
        multi_file=False,
        title='Selecciona la planilla .xlsm'
    )


def llamar_leer_xlsm(ruta_xlsm):
    """Ejecuta leer_xlsm_codigos.py y devuelve ruta del CODIGO.csv."""
    try:
        salida = subprocess.check_output(
            [PYTHON3_EXE, LEER_XLSM, ruta_xlsm, DATA_DIR],
            stderr=subprocess.STDOUT,
            universal_newlines=True,
            creationflags=CREATE_NO_WINDOW
        )
        lineas = [l.strip() for l in salida.splitlines() if l.strip()]
        if not lineas:
            forms.alert(
                "leer_xlsm_codigos.py no devolvió ninguna ruta.\n\nSalida:\n{}".format(salida),
                title="Error"
            )
            return None
        csv_path = lineas[-1]
        if not os.path.exists(csv_path):
            forms.alert(
                "No se generó el CSV de CODIGO.\n\nTexto devuelto:\n{}".format(salida),
                title="Error"
            )
            return None
        return csv_path
    except subprocess.CalledProcessError as e:
        forms.alert(
            "Error al leer .xlsm:\n{}".format(e.output),
            title="Error"
        )
        return None


def llamar_ui_y_formato(ruta_xlsm, csv_codigos):
    # Timestamp AAAAMMDD_hhmm
    ahora = datetime.now()
    stamp = ahora.strftime("%Y%m%d_%H%M")

    ruta_xlsx_salida = os.path.join(
        os.path.dirname(ruta_xlsm),
        "planilla-modelo_{}.xlsx".format(stamp)
    )

    # Aquí NO se modifica modelo_codint_por_cm.json ni planillas_headers_order.json.
    # Se asume que ya existen y se han generado desde Revit / otro flujo.
    try:
        subprocess.check_call(
            [
                PYTHON3_EXE,
                UI_COMPARACION,
                SCRIPT_JSON_PATH,
                csv_codigos,
                DATA_DIR,
                FORMATEAR_XLSX,
                ruta_xlsx_salida,
                PYTHON3_EXE,   # 6º argumento (python_exe)
                MODELO_JSON,   # 7º argumento (modelo_json)
                HEADERS_JSON   # 8º argumento (headers_json, ignorado por la UI nueva)
            ],
            stderr=subprocess.STDOUT,
            creationflags=CREATE_NO_WINDOW
        )
        forms.alert(
            "Archivo generado:\n{}".format(ruta_xlsx_salida),
            title="Éxito"
        )
    except subprocess.CalledProcessError as e:
        forms.alert(
            "Error en la UI / formateo:\n{}".format(e),
            title="Error"
        )


def main():
    if not os.path.exists(SCRIPT_JSON_PATH):
        forms.alert(
            "No se encontró script.json en data:\n{}".format(SCRIPT_JSON_PATH),
            title="Error"
        )
        return

    if not os.path.exists(LEER_XLSM) or not os.path.exists(UI_COMPARACION):
        forms.alert(
            "Faltan scripts CPython en la carpeta data.\n"
            "Se esperaban:\n{}\n{}".format(LEER_XLSM, UI_COMPARACION),
            title="Error"
        )
        return

    # Aviso si aún no has generado los JSON desde Revit
    if not os.path.exists(MODELO_JSON):
        forms.alert(
            "Falta el archivo de modelo:\n{}\n"
            "Ejecuta primero la rutina que genera modelo_codint_por_cm.json."
            .format(MODELO_JSON),
            title="Error"
        )
        return

    ruta_xlsm = seleccionar_xlsm()
    if not ruta_xlsm:
        return

    csv_codigos = llamar_leer_xlsm(ruta_xlsm)
    if not csv_codigos:
        return

    llamar_ui_y_formato(ruta_xlsm, csv_codigos)


if __name__ == '__main__':
    main()


#==================================================
#🚫 DELETE BELOW
from Snippets._customprint import kit_button_clicked    # Import Reusable Function from 'lib/Snippets/_customprint.py'
kit_button_clicked(btn_name=__title__)                  # Display Default Print Message
