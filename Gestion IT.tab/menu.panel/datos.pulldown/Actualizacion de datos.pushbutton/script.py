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
import traceback
import os
from datetime import datetime

# ╦  ╦╔═╗╦═╗╦╔═╗╔╗ ╦  ╔═╗╔═╗
# ╚╗╔╝╠═╣╠╦╝║╠═╣╠╩╗║  ║╣ ╚═╗
#  ╚╝ ╩ ╩╩╚═╩╩ ╩╚═╝╩═╝╚═╝╚═╝
#==================================================
#uidoc = __revit__.ActiveUIDocument
#doc = uidoc.Document

# ╔╦╗╔═╗╦╔╗╔
# ║║║╠═╣║║║║
# ╩ ╩╩ ╩╩╝╚╝
#==================================================

# ---------------------------------------------------------
# Utilidades de log (mismo archivo que el addin)
# ---------------------------------------------------------
def get_data_dir():
    appdata = os.environ.get('APPDATA', '')
    data_dir = os.path.join(appdata, 'MyPyRevitExtention', 'PyRevitIT.extension', 'data')
    if not os.path.isdir(data_dir):
        try:
            os.makedirs(data_dir)
        except:
            pass
    return data_dir

def log(msg):
    try:
        log_path = os.path.join(get_data_dir(), 'batchupdater_addin.log')
        with open(log_path, 'a', encoding='utf-8') as f:
            f.write('[script.py] {0} | {1}\n'.format(
                datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                msg
            ))
    except:
        pass

log('script.py iniciado (modo assembly ya cargado).')

try:
    # 1) Referenciar el ensamblado que ya cargó Revit por el .addin
    log('Intentando clr.AddReference("RevitBatchUpdater")...')
    clr.AddReference('RevitBatchUpdater')
    log('Referencia a RevitBatchUpdater obtenida correctamente.')

    # 2) Importar la clase de entrada del addin
    from RevitBatchUpdater import AddinEntry
    log('Import AddinEntry OK.')

    # 3) Lanzar el ExternalEvent
    log('Llamando AddinEntry.LanzarActualizador()...')
    AddinEntry.LanzarActualizador()
    log('AddinEntry.LanzarActualizador() ejecutado sin excepción.')

except Exception as e:
    log('EXCEPCION en script.py: {0}'.format(e))
    log('Traceback: {0}'.format(traceback.format_exc()))
    raise

#==================================================
#🚫 DELETE BELOW
from Snippets._customprint import kit_button_clicked    # Import Reusable Function from 'lib/Snippets/_customprint.py'
kit_button_clicked(btn_name=__title__)                  # Display Default Print Message
