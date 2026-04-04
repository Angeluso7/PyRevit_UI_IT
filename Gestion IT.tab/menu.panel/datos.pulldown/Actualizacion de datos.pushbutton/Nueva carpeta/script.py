# -*- coding: utf-8 -*-
__title__   = "Actualizacion de datos"
__doc__     = """Version = 1.0
Date    = 15.06.2024
________________________________________________________________
Description:

This is the placeholder for a .pushbutton in a /pulldown
You can use it to start your pyRevit Add-In

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
- [15.06.2024] v1.0 Change Description
- [10.06.2024] v0.5 Change Description
- [05.06.2024] v0.1 Change Description 
________________________________________________________________
Author: Erik Frits"""

# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝
#==================================================
import clr
import os
import json
import System
from Autodesk.Revit.UI import TaskDialog
# ╦  ╦╔═╗╦═╗╦╔═╗╔╗ ╦  ╔═╗╔═╗
# ╚╗╔╝╠═╣╠╦╝║╠═╣╠╩╗║  ║╣ ╚═╗
#  ╚╝ ╩ ╩╩╚═╩╩ ╩╚═╝╩═╝╚═╝╚═╝
#==================================================
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

# ╔╦╗╔═╗╦╔╗╔
# ║║║╠═╣║║║║
# ╩ ╩╩ ╩╩╝╚╝
#==================================================

def log(msg):
    appdata = System.Environment.GetFolderPath(System.Environment.SpecialFolder.ApplicationData)
    ext_dir = os.path.join(appdata, "MyPyRevitExtention", "PyRevitIT.extension")
    data_dir = os.path.join(ext_dir, "data")
    if not os.path.isdir(data_dir):
        os.makedirs(data_dir)
    log_path = os.path.join(data_dir, "pyrevit_actualizacion.log")
    with open(log_path, "a") as fp:
        fp.write(u"[script.py] {0} | {1}\n".format(System.DateTime.Now, msg))

try:
    # Carpeta data
    appdata = System.Environment.GetFolderPath(System.Environment.SpecialFolder.ApplicationData)
    ext_dir = os.path.join(appdata, "MyPyRevitExtention", "PyRevitIT.extension")
    data_dir = os.path.join(ext_dir, "data")
    if not os.path.isdir(data_dir):
        os.makedirs(data_dir)
        log("Creada carpeta data: {0}".format(data_dir))
    else:
        log("Usando carpeta data existente: {0}".format(data_dir))

    # 1) Leer config_proyecto_activo.json
    config_path = os.path.join(data_dir, 'config_proyecto_activo.json')
    if not os.path.exists(config_path):
        msg = 'No se encontró "config_proyecto_activo.json" en la carpeta data.'
        log(msg)
        TaskDialog.Show('Actualización de datos', msg)
        raise SystemExit

    with open(config_path, 'r') as fp:
        config = json.load(fp)

    ruta_repo = config.get('ruta_repositorio_activo', '')
    if not ruta_repo:
        msg = 'El archivo config_proyecto_activo.json no contiene "ruta_repositorio_activo".'
        log(msg)
        TaskDialog.Show('Actualización de datos', msg)
        raise SystemExit

    nombre_bd = os.path.basename(ruta_repo)
    bd_path = os.path.join(data_dir, nombre_bd)
    log("Config encontrada. ruta_repo={0}, nombre_bd={1}, bd_path={2}".format(
        ruta_repo, nombre_bd, bd_path))

    # 2) Verificar existencia de BD
    if not os.path.exists(bd_path):
        log("BD no encontrada. Creando BD vacía: {0}".format(bd_path))
        with open(bd_path, 'w') as fp:
            json.dump({}, fp, indent=2)
    else:
        log("BD existente detectada: {0}".format(bd_path))

    # 3) Aviso
    TaskDialog.Show(
        'Actualización de datos',
        'Actualización activada, se verán los cambios en la próxima sesión.'
    )
    log("Script pyRevit finalizado correctamente para el modelo: {0}".format(doc.Title))

except Exception as ex:
    log("EXCEPCION en script.py: {0}".format(ex))
    TaskDialog.Show('Actualización de datos', 'Error en script.py: {0}'.format(ex))


#==================================================
#🚫 DELETE BELOW
from Snippets._customprint import kit_button_clicked    # Import Reusable Function from 'lib/Snippets/_customprint.py'
kit_button_clicked(btn_name=__title__)                  # Display Default Print Message
