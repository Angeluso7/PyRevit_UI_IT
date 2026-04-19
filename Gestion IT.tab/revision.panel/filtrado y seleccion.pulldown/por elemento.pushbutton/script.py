# -*- coding: utf-8 -*-
__title__ = "Por Elemento"

__doc__ = """Version = 1.2
Date    = 18.04.2026
_______________________________________________________________
Description:

Selecciona un elemento (host o link), arma sus datos segun la
planilla asociada a CodIntBIM tomando primero el repositorio
activo (ruta_repositorio_activo) y, si no hay datos, desde el
modelo. Pasa la informacion a un visor CPython (Tkinter) en
modo solo lectura.
________________________________________________________________
How-To:
1. [Hold ALT + CLICK] on the button to open its source folder.
2. Automate Your Boring Work ;)
________________________________________________________________
Last Updates:
- [01.09.2025] v0.1  Inicio de Aplicacion.
- [18.04.2026] v1.2  Rutas centralizadas via config.paths.
________________________________________________________________
Author: Argenis Angel"""

# ╦╔╦╗╔═╗╦═╗╦╔═╗╔╦╗╔═╗
# ║║║║╠═╝╠╦╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╩ ╚═╝ ╚ ╚═╝
#==================================================

import clr
import os
import sys
import json
import subprocess

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System.Windows.Forms")

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    RevitLinkInstance,
    ViewSchedule,
)
from Autodesk.Revit.UI.Selection import ObjectType
from System.Windows.Forms import MessageBox

# ╦  ╦╔═╗╦═╗╦╔═╗╔╗ ╦  ╔═╗╔═╗
# ╚╗╔╝╠═╣╠╦╝║╠═╣╠╩╗║  ║╣ ╚═╗
#  ╚╝ ╩ ╩╩╚═╩╩ ╩╚═╝╩═╝╚═╝╚═╝
#==================================================

app   = __revit__.Application
uidoc = __revit__.ActiveUIDocument
doc   = uidoc.Document

# ╔╦╗╔═╗╦╔╗╔
# ║║║╠═╣║║║║
# ╩ ╩╩ ╩╩╝╚╝
#==================================================

# ── Rutas centralizadas desde config.paths ────────────────────────────────────
try:
    _this_dir = os.path.dirname(os.path.abspath(__file__))
except Exception:
    _this_dir = os.getcwd()

# pushbutton(1) -> pulldown(2) -> panel(3) -> tab(4) -> EXT_ROOT
_EXT_ROOT = os.path.abspath(os.path.join(_this_dir, '..', '..', '..', '..'))
_LIB_DIR  = os.path.join(_EXT_ROOT, 'lib')
if _LIB_DIR not in sys.path:
    sys.path.insert(0, _LIB_DIR)

try:
    from config.paths import DATA_DIR, MASTER_DIR, TEMP_DIR, CACHE_DIR, \
                             CONFIG_PROYECTO, REGISTRO_PROYECTOS, \
                             SCRIPT_JSON_PATH_LIB, ensure_runtime_dirs
    from core.env_config import get_python_exe
    ensure_runtime_dirs()
    PYTHON_EXE = get_python_exe()
except Exception as _path_err:
    _DATA_DIR       = os.path.join(_EXT_ROOT, 'data')
    DATA_DIR        = _DATA_DIR
    MASTER_DIR      = os.path.join(_DATA_DIR, 'master')
    TEMP_DIR        = os.path.join(_DATA_DIR, 'temp')
    CACHE_DIR       = os.path.join(_DATA_DIR, 'cache')
    CONFIG_PROYECTO = os.path.join(MASTER_DIR, 'config_proyecto_activo.json')
    REGISTRO_PROYECTOS   = os.path.join(MASTER_DIR, 'registro_proyectos.json')
    SCRIPT_JSON_PATH_LIB = os.path.join(MASTER_DIR, 'script.json')
    import glob as _glob
    def _fb_python():
        base = os.path.join(os.path.expanduser('~'), 'AppData', 'Local', 'Programs', 'Python')
        for exe in ('python.exe', 'pythonw.exe'):
            for cand in sorted(_glob.glob(os.path.join(base, 'Python3*', exe)), reverse=True):
                return cand
        for folder in os.environ.get('PATH', '').split(os.pathsep):
            cand = os.path.join(folder.strip(), 'python.exe')
            if os.path.isfile(cand):
                return cand
        return None
    PYTHON_EXE = _fb_python()

CURRENT_FOLDER        = os.path.dirname(os.path.abspath(__file__))
PLANILLAS_ORDER_PATH  = os.path.join(CURRENT_FOLDER, "planillas_headers_order.json")
REPO_ELEMENTO_TMP     = os.path.join(TEMP_DIR, "repo_elemento_tmp.json")
CPYTHON_VIEWER_PATH   = os.path.join(CURRENT_FOLDER, "visor_elemento.pyw")
CONFIG_PATH           = CONFIG_PROYECTO
SCRIPT_JSON           = SCRIPT_JSON_PATH_LIB

# ╦╔╦╗╔═╗╦═╗╦╔═╗
# ║║║║╠═╝╠╦╝║ ║╠═╣
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╩╚═╝
#==================================================

def load_json(path, default=None):
    if not os.path.isfile(path):
        return default
    try:
        with open(path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return default


def save_json(path, data):
    try:
        folder = os.path.dirname(path)
        if folder and not os.path.exists(folder):
            os.makedirs(folder)
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        return True
    except Exception:
        return False


def load_config():
    data = load_json(CONFIG_PATH)
    if data is None:
        raise IOError(u"No se encontro config_proyecto_activo.json en:\n{}".format(CONFIG_PATH))
    return data


def load_script_json():
    data = load_json(SCRIPT_JSON)
    if data is None:
        raise IOError(u"No se encontro script.json en:\n{}".format(SCRIPT_JSON))
    return data


def load_planillas_order():
    return load_json(PLANILLAS_ORDER_PATH, default={})


def get_active_repo():
    cfg  = load_config()
    ruta = (cfg.get("ruta_repositorio_activo") or "").strip()
    if not ruta or not os.path.isfile(ruta):
        return {}
    return load_json(ruta, default={})


# ╬╦  ╬╔═╗╬  ╬╔╦╗╔╦╗
# ╠──╠╦╝╠═╣╠╦╝║ ║║║║
# ╩  ╩╚═ ╩╚═╝╩╚═ ╚╝╚╝
#==================================================

def get_schedule_fields(schedule):
    """Devuelve lista de (field_id, field_name, param_name) de una planilla."""
    fields = []
    sched_def = schedule.Definition
    for i in range(sched_def.GetFieldCount()):
        field = sched_def.GetField(i)
        try:
            param_name = field.GetName()
        except Exception:
            param_name = ""
        fields.append((field.FieldId, field.GetName(), param_name))
    return fields


def get_element_params(element, field_names):
    """Extrae parametros de un elemento dado una lista de nombres de campo."""
    result = {}
    for name in field_names:
        param = element.LookupParameter(name)
        if param:
            try:
                result[name] = param.AsString() or param.AsValueString() or ""
            except Exception:
                result[name] = ""
        else:
            result[name] = ""
    return result


def find_schedule_for_codint(codint, script_data, doc_context):
    """
    Busca en el documento la planilla ViewSchedule cuyo nombre coincide
    con el codigo de planilla asociado al CodIntBIM.
    Devuelve (schedule, tabla_name) o (None, None).
    """
    codigos = script_data.get("codigos_planillas", {})
    tabla_name = codigos.get(codint)
    if not tabla_name:
        return None, None
    for sched in FilteredElementCollector(doc_context).OfClass(ViewSchedule):
        if sched.Name == tabla_name:
            return sched, tabla_name
    return None, tabla_name


def build_element_data(element, schedule, repo_data, doc_name):
    """Construye el diccionario de datos del elemento para el visor."""
    fields = get_schedule_fields(schedule)
    field_names = [f[1] for f in fields]
    model_params = get_element_params(element, field_names)

    codint = model_params.get("CodIntBIM", "") or model_params.get("Cod.Int.BIM", "")
    repo_entry = repo_data.get(codint, {}) if codint else {}

    merged = {}
    for name in field_names:
        merged[name] = repo_entry.get(name) if repo_entry.get(name) else model_params.get(name, "")

    return {
        "elemento_id":  str(element.Id.IntegerValue),
        "documento":    doc_name,
        "codint":       codint,
        "planilla":     schedule.Name,
        "parametros":   merged,
        "orden_campos": field_names,
    }


# ╦╔╦╗╔═╗╦╔╗╔
# ║║║║╠═╣║║║║
# ╩ ╩╩ ╩╩╝╚╝
#==================================================

def main():
    from pyrevit import forms

    # 1. Cargar datos de configuracion
    try:
        script_data = load_script_json()
        repo_data   = get_active_repo()
    except IOError as e:
        forms.alert(str(e), title=u"Error de configuracion")
        return

    # 2. Pedir al usuario que seleccione un elemento
    try:
        ref = uidoc.Selection.PickObject(ObjectType.Element, u"Selecciona un elemento")
    except Exception:
        return  # usuario cancelo

    element = doc.GetElement(ref.ElementId)
    if element is None:
        forms.alert(u"No se pudo obtener el elemento.", title=u"Error")
        return

    # 3. Obtener CodIntBIM
    p = element.LookupParameter("CodIntBIM") or element.LookupParameter("Cod.Int.BIM")
    codint = (p.AsString() or "").strip() if p else ""
    if not codint:
        forms.alert(
            u"El elemento seleccionado no tiene parametro CodIntBIM.",
            title=u"Sin codigo"
        )
        return

    # 4. Buscar planilla asociada
    schedule, tabla_name = find_schedule_for_codint(codint, script_data, doc)
    if schedule is None:
        msg = u"No se encontro planilla para CodIntBIM: {}".format(codint)
        if tabla_name:
            msg += u"\nNombre buscado: {}".format(tabla_name)
        forms.alert(msg, title=u"Planilla no encontrada")
        return

    # 5. Construir datos y guardar JSON temporal
    elem_data = build_element_data(element, schedule, repo_data, doc.Title)
    if not save_json(REPO_ELEMENTO_TMP, elem_data):
        forms.alert(u"No se pudo guardar el JSON temporal.", title=u"Error")
        return

    # 6. Lanzar visor CPython
    if not os.path.isfile(CPYTHON_VIEWER_PATH):
        forms.alert(
            u"No se encontro el visor:\n{}".format(CPYTHON_VIEWER_PATH),
            title=u"Error"
        )
        return
    if not PYTHON_EXE or not os.path.isfile(PYTHON_EXE):
        forms.alert(u"No se encontro Python 3 en este equipo.", title=u"Error")
        return

    import subprocess
    subprocess.Popen([PYTHON_EXE, CPYTHON_VIEWER_PATH, REPO_ELEMENTO_TMP])


if __name__ == '__main__':
    main()
