# -*- coding: utf-8 -*-
__title__ = "Por Elemento"
__doc__ = (
    "Version = 1.4\n"
    "Date    = 19.04.2026\n"
    "Desc    = Selecciona elemento (host o vinculado), extrae parametros\n"
    "          de su planilla asociada, mezcla con repo activo y abre\n"
    "          visor CPython con soporte de vista agrupada.\n"
    "Author  = Argenis Angel"
)

import clr, os, sys, json, subprocess

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import (
    FilteredElementCollector, RevitLinkInstance, ViewSchedule, ElementId,
)
from Autodesk.Revit.UI.Selection import ObjectType

app   = __revit__.Application
uidoc = __revit__.ActiveUIDocument
doc   = uidoc.Document

# ── Rutas centralizadas ──────────────────────────────────────────────────────
try:
    _this_dir = os.path.dirname(os.path.abspath(__file__))
except Exception:
    _this_dir = os.getcwd()

_EXT_ROOT = os.path.abspath(os.path.join(_this_dir, '..', '..', '..', '..'))
_LIB_DIR  = os.path.join(_EXT_ROOT, 'lib')
if _LIB_DIR not in sys.path:
    sys.path.insert(0, _LIB_DIR)

try:
    from config.paths import (DATA_DIR, MASTER_DIR, TEMP_DIR, CACHE_DIR,
                               CONFIG_PROYECTO, REGISTRO_PROYECTOS,
                               SCRIPT_JSON_PATH_LIB, ensure_runtime_dirs)
    from core.env_config import get_python_exe
    ensure_runtime_dirs()
    PYTHON_EXE = get_python_exe()
except Exception:
    _DATA_DIR            = os.path.join(_EXT_ROOT, 'data')
    DATA_DIR             = _DATA_DIR
    MASTER_DIR           = os.path.join(_DATA_DIR, 'master')
    TEMP_DIR             = os.path.join(_DATA_DIR, 'temp')
    CACHE_DIR            = os.path.join(_DATA_DIR, 'cache')
    CONFIG_PROYECTO      = os.path.join(MASTER_DIR, 'config_proyecto_activo.json')
    REGISTRO_PROYECTOS   = os.path.join(MASTER_DIR, 'registro_proyectos.json')
    SCRIPT_JSON_PATH_LIB = os.path.join(MASTER_DIR, 'script.json')

    import glob as _glob
    def _fb_python():
        base = os.path.join(os.path.expanduser('~'), 'AppData',
                            'Local', 'Programs', 'Python')
        for exe in ('python.exe', 'pythonw.exe'):
            for cand in sorted(
                    _glob.glob(os.path.join(base, 'Python3*', exe)), reverse=True):
                return cand
        for folder in os.environ.get('PATH', '').split(os.pathsep):
            cand = os.path.join(folder.strip(), 'python.exe')
            if os.path.isfile(cand):
                return cand
        return None
    PYTHON_EXE = _fb_python()

CURRENT_FOLDER      = os.path.dirname(os.path.abspath(__file__))
REPO_ELEMENTO_TMP   = os.path.join(TEMP_DIR, "repo_elemento_tmp.json")
CPYTHON_VIEWER_PATH = os.path.join(CURRENT_FOLDER, "visor_elemento.pyw")
CONFIG_PATH         = CONFIG_PROYECTO
SCRIPT_JSON         = SCRIPT_JSON_PATH_LIB

# ── Utilidades JSON ──────────────────────────────────────────────────────────
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

def load_script_json():
    data = load_json(SCRIPT_JSON)
    if data is None:
        raise IOError(u"No se encontro script.json en:\n{}".format(SCRIPT_JSON))
    return data

def get_active_repo():
    try:
        cfg  = load_json(CONFIG_PATH) or {}
        ruta = (cfg.get("ruta_repositorio_activo") or "").strip()
        if not ruta or not os.path.isfile(ruta):
            return {}
        return load_json(ruta, default={}) or {}
    except Exception:
        return {}

# ── Helpers Revit ────────────────────────────────────────────────────────────
def get_schedule_field_names(schedule):
    sched_def = schedule.Definition
    names = []
    for i in range(sched_def.GetFieldCount()):
        try:
            names.append(sched_def.GetField(i).GetName())
        except Exception:
            pass
    return names

def get_element_params(element, field_names):
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

def find_schedule_for_codint(codint, script_data, doc_host, doc_link=None):
    codigos    = script_data.get("codigos_planillas", {})
    tabla_name = codigos.get(codint)
    if not tabla_name:
        return None, None
    for doc_ctx in filter(None, [doc_host, doc_link]):
        for sched in FilteredElementCollector(doc_ctx).OfClass(ViewSchedule):
            if sched.Name == tabla_name:
                return sched, tabla_name
    return None, tabla_name

# ── MAIN ─────────────────────────────────────────────────────────────────────
def main():
    from pyrevit import forms

    try:
        script_data = load_script_json()
        repo_data   = get_active_repo()
    except IOError as e:
        forms.alert(str(e), title=u"Error de configuracion")
        return

    element      = None
    link_doc     = None
    doc_name_src = doc.Title
    ref          = None

    try:
        ref = uidoc.Selection.PickObject(
            ObjectType.LinkedElement,
            u"Selecciona un elemento (host o vinculado)"
        )
    except Exception:
        pass

    if ref is None:
        try:
            ref = uidoc.Selection.PickObject(
                ObjectType.Element,
                u"Selecciona un elemento host"
            )
        except Exception:
            return

    if ref is None:
        return

    try:
        link_elem_id = ref.LinkedElementId
    except Exception:
        link_elem_id = ElementId.InvalidElementId

    if link_elem_id != ElementId.InvalidElementId:
        link_inst_elem = doc.GetElement(ref.ElementId)
        if isinstance(link_inst_elem, RevitLinkInstance):
            link_doc     = link_inst_elem.GetLinkDocument()
            doc_name_src = (link_doc.Title if link_doc
                            else os.path.basename(str(link_inst_elem.Name)))
            element = link_doc.GetElement(link_elem_id) if link_doc else None
    else:
        element = doc.GetElement(ref.ElementId)

    if element is None:
        forms.alert(u"No se pudo obtener el elemento.", title=u"Error")
        return

    p      = (element.LookupParameter("CodIntBIM") or
              element.LookupParameter("Cod.Int.BIM"))
    codint = ((p.AsString() or "").strip()) if p else ""
    if not codint:
        forms.alert(
            u"El elemento seleccionado no tiene CodIntBIM.\n"
            u"ID: {}  Documento: {}".format(element.Id.IntegerValue, doc_name_src),
            title=u"Sin codigo"
        )
        return

    schedule, tabla_name = find_schedule_for_codint(
        codint, script_data, doc, link_doc)

    if schedule is None:
        msg = u"No se encontro planilla para CodIntBIM: {}".format(codint)
        if tabla_name:
            msg += u"\nNombre buscado: {}".format(tabla_name)
        forms.alert(msg, title=u"Planilla no encontrada")
        return

    field_names  = get_schedule_field_names(schedule)
    model_params = get_element_params(element, field_names)

    _fname_base = os.path.splitext(os.path.basename(doc_name_src))[0]
    elem_id_int = element.Id.IntegerValue
    repo_key    = "{}_{}".format(_fname_base, elem_id_int)
    repo_entry  = repo_data.get(repo_key, {}) or {}

    merged = {}
    for name in field_names:
        val_repo  = repo_entry.get(name, "")
        val_model = model_params.get(name, "")
        merged[name] = val_repo if val_repo else val_model

    payload = {
        "Headers"   : field_names,
        "Row"       : merged,
        "ElementId" : str(elem_id_int),
        "Planilla"  : schedule.Name,
        "Archivo"   : doc_name_src,
        "CodIntBIM" : codint,
        "RepoKey"   : repo_key,
        "AllRows"   : [{
            "RepoKey"   : repo_key,
            "ElementId" : str(elem_id_int),
            "Archivo"   : doc_name_src,
            "Row"       : merged,
        }],
    }

    if not save_json(REPO_ELEMENTO_TMP, payload):
        forms.alert(u"No se pudo guardar el JSON temporal.", title=u"Error")
        return

    if not os.path.isfile(CPYTHON_VIEWER_PATH):
        forms.alert(
            u"No se encontro el visor:\n{}".format(CPYTHON_VIEWER_PATH),
            title=u"Error"
        )
        return
    if not PYTHON_EXE or not os.path.isfile(PYTHON_EXE):
        forms.alert(u"No se encontro Python 3 en este equipo.", title=u"Error")
        return

    subprocess.Popen([PYTHON_EXE, CPYTHON_VIEWER_PATH, REPO_ELEMENTO_TMP])


if __name__ == '__main__':
    main()
