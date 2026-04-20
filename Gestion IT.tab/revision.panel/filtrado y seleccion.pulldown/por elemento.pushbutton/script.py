# -*- coding: utf-8 -*-
__title__ = "Por Elemento"
__doc__ = (
    "Version = 1.6\n"
    "Date    = 19.04.2026\n"
    "Desc    = Selecciona elemento (host o vinculado), extrae parametros\n"
    "          de su planilla asociada, mezcla con repo activo y abre\n"
    "          visor CPython (Tkinter) en modo solo lectura.\n"
    "Author  = Argenis Angel"
)

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
    from config.paths import (DATA_DIR, MASTER_DIR, TEMP_DIR,
                               CONFIG_PROYECTO, SCRIPT_JSON_PATH_LIB,
                               get_ruta_repositorio, ensure_runtime_dirs)
    from core.env_config import get_python_exe
    ensure_runtime_dirs()
    PYTHON_EXE = get_python_exe()
except Exception:
    _DATA_DIR            = os.path.join(_EXT_ROOT, 'data')
    DATA_DIR             = _DATA_DIR
    MASTER_DIR           = os.path.join(_DATA_DIR, 'master')
    TEMP_DIR             = os.path.join(_DATA_DIR, 'temp')
    CONFIG_PROYECTO      = os.path.join(MASTER_DIR, 'config_proyecto_activo.json')
    SCRIPT_JSON_PATH_LIB = os.path.join(MASTER_DIR, 'script.json')

    def get_ruta_repositorio(nup):
        nombre = u'repositorio_datos_{}.json'.format(nup)
        return os.path.join(_DATA_DIR, 'proyectos', nombre)

    import glob as _glob
    def _fb_python():
        base = os.path.join(os.path.expanduser('~'), 'AppData',
                            'Local', 'Programs', 'Python')
        for exe in ('pythonw.exe', 'python.exe'):
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
REPO_ELEMENTO_TMP   = os.path.join(TEMP_DIR, 'repo_elemento_tmp.json')
CPYTHON_VIEWER_PATH = os.path.join(CURRENT_FOLDER, 'visor_elemento.pyw')
PLANILLAS_ORDER_PATH = os.path.join(CURRENT_FOLDER, 'planillas_headers_order.json')
CONFIG_PATH         = CONFIG_PROYECTO
SCRIPT_JSON         = SCRIPT_JSON_PATH_LIB

# ── Utilidades JSON ──────────────────────────────────────────────────────────
def load_json(path, show_error=True):
    if not os.path.exists(path):
        if show_error:
            MessageBox.Show(
                u"No se encontro archivo JSON:\n{}".format(path), "Error")
        return None
    try:
        with open(path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        if show_error:
            MessageBox.Show(
                u"Error cargando JSON:\n{}".format(e), "Error")
        return None

def save_json(data, path, show_error=True):
    try:
        folder = os.path.dirname(path)
        if folder and not os.path.exists(folder):
            os.makedirs(folder)
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        if show_error:
            MessageBox.Show(
                u"Error guardando JSON:\n{}".format(e), "Error")
        return False

# ── Repositorio activo (portable) ───────────────────────────────────────────────
def get_active_repo():
    """
    Lee el repositorio activo usando get_ruta_repositorio(nup).
    No depende de rutas absolutas hardcodeadas.
    """
    try:
        cfg = load_json(CONFIG_PATH, show_error=False) or {}
        nup = (cfg.get('nup') or cfg.get('decreto') or '').strip()
        if not nup:
            return {}
        ruta = get_ruta_repositorio(nup)
        if not os.path.isfile(ruta):
            return {}
        return load_json(ruta, show_error=False) or {}
    except Exception:
        return {}

# ── Seleccion de elemento ──────────────────────────────────────────────────────────
def seleccionar_elemento():
    """Intenta primero LinkedElement, luego Element host."""
    try:
        ref = uidoc.Selection.PickObject(
            ObjectType.LinkedElement,
            u"Seleccione un elemento (link o modelo)"
        )
        try:
            link_instance = doc.GetElement(ref.ElementId)
            if isinstance(link_instance, RevitLinkInstance):
                linked_doc = link_instance.GetLinkDocument()
                if linked_doc is None:
                    MessageBox.Show(
                        u"No se pudo acceder al documento vinculado.", "Error")
                    return None, None
                elem = linked_doc.GetElement(ref.LinkedElementId)
                if elem is None:
                    MessageBox.Show(
                        u"Elemento vinculado no encontrado.", "Error")
                    return None, None
                return elem, linked_doc
        except Exception:
            pass
    except Exception:
        pass

    try:
        ref = uidoc.Selection.PickObject(
            ObjectType.Element,
            u"Seleccione un elemento del modelo"
        )
        elem = doc.GetElement(ref.ElementId)
        if elem is None:
            MessageBox.Show(u"Elemento no encontrado.", "Error")
            return None, None
        return elem, doc
    except Exception as e:
        MessageBox.Show(
            u"Seleccion cancelada o error:\n{}".format(e), "Aviso")
        return None, None

# ── Planilla y encabezados ────────────────────────────────────────────────────────
def obtener_encabezados_planilla(host_doc, codintbim_val, script_data):
    """
    Busca la planilla asociada al prefijo de 4 chars del CodIntBIM.
    Busca el prefijo en los VALORES de codigos_planillas y obtiene
    el nombre de planilla (clave), luego extrae encabezados con GetFieldOrder.
    """
    codigos_planillas = script_data.get('codigos_planillas', {})
    pref_cod = (codintbim_val or '')[:4].upper()
    clave_planilla = None

    # Los valores del dict son codigos cortos (ej. 'CM52'); las claves son nombres
    for clave, codigo in codigos_planillas.items():
        codigo_str = (codigo or '') if isinstance(codigo, str) else ''
        if codigo_str.upper() == pref_cod:
            clave_planilla = clave
            break

    if not clave_planilla:
        MessageBox.Show(
            u"No se encontro clave de planilla para el prefijo '{}'."
            .format(pref_cod), u"Aviso")
        return None, None

    # Buscar la ViewSchedule por nombre exacto en el documento host
    reemplazos = script_data.get('reemplazos_de_nombres', {})
    nombre_display = reemplazos.get(clave_planilla, clave_planilla)

    schedules = (
        FilteredElementCollector(host_doc)
        .OfClass(ViewSchedule)
        .WhereElementIsNotElementType()
        .ToElements()
    )
    planilla_obj = next(
        (s for s in schedules
         if s.Name == nombre_display or s.Name == clave_planilla), None
    )

    if planilla_obj is None:
        MessageBox.Show(
            u"No se encontro la planilla '{}'.".format(nombre_display),
            u"Error")
        return None, None

    try:
        headers_order = []
        for fid in planilla_obj.Definition.GetFieldOrder():
            field = planilla_obj.Definition.GetField(fid)
            if field:
                headers_order.append(field.GetName())
        if not headers_order:
            MessageBox.Show(
                u"No se obtuvieron encabezados de la planilla.", u"Aviso")
            return None, None
    except Exception as e:
        MessageBox.Show(
            u"Error obteniendo encabezados:\n{}".format(e), u"Error")
        return None, None

    return planilla_obj.Name, headers_order


def obtener_headers_con_cache(host_doc, codintbim_val, script_data):
    """Encabezados con cache local en planillas_headers_order.json."""
    planillas_order = load_json(PLANILLAS_ORDER_PATH, show_error=False) or {}

    clave_planilla, headers_from_model = obtener_encabezados_planilla(
        host_doc, codintbim_val, script_data)
    if not headers_from_model:
        return None, []

    if clave_planilla in planillas_order:
        headers = planillas_order[clave_planilla] or []
    else:
        headers = headers_from_model
        planillas_order[clave_planilla] = headers_from_model
        save_json(planillas_order, PLANILLAS_ORDER_PATH, show_error=False)

    return clave_planilla, headers

# ── Construccion de datos del elemento ──────────────────────────────────────────
def construir_datos_elemento(elem, linked_doc, headers_order, script_data):
    """
    Usa primero el repositorio activo y, si no hay datos, desde el modelo.
    Aplica reemplazos_de_nombres para mapear columnas.
    """
    reemplazos = script_data.get('reemplazos_de_nombres', {})
    repo_datos = get_active_repo()

    elem_id_str = str(elem.Id.IntegerValue)
    archivo     = ''
    if linked_doc is not None:
        archivo = linked_doc.PathName or ''
    else:
        archivo = doc.PathName or ''

    base_key     = u'{}_{}'.format(archivo, elem_id_str)
    data_elemento = repo_datos.get(base_key, None)
    resultado    = {}

    try:
        if data_elemento:
            filtrado = {
                k: v for k, v in data_elemento.items()
                if k not in ['ElementId', 'Archivo', 'nombre_archivo']
            }
            tmp = {reemplazos.get(k, k): v for k, v in filtrado.items()}
            for head in headers_order:
                resultado[head] = tmp.get(head, '')
        else:
            parametros = {}
            for p in elem.Parameters:
                try:
                    if p.HasValue:
                        val = p.AsString() or p.AsValueString()
                        if val and val.strip():
                            parametros[p.Definition.Name] = val.strip()
                except Exception:
                    continue
            tmp = {reemplazos.get(k, k): v for k, v in parametros.items()}
            for head in headers_order:
                resultado[head] = tmp.get(head, '')
    except Exception as e:
        MessageBox.Show(
            u"Error preparando datos del elemento:\n{}".format(e), u"Error")
        return None, archivo, elem_id_str

    return resultado, archivo, elem_id_str

# ── MAIN ─────────────────────────────────────────────────────────────────────
def main():
    elem, linked_doc = seleccionar_elemento()
    if elem is None:
        return

    script_data = load_json(SCRIPT_JSON) or {}
    if not script_data:
        return

    try:
        p_cod  = elem.LookupParameter('CodIntBIM')
        cod_val = (p_cod.AsString() or '').strip() if (p_cod and p_cod.HasValue) else ''
    except Exception:
        cod_val = ''

    clave_planilla, headers_order = obtener_headers_con_cache(
        doc, cod_val, script_data)
    if not headers_order:
        return

    resultado, archivo, elem_id_str = construir_datos_elemento(
        elem, linked_doc, headers_order, script_data)
    if resultado is None:
        return

    json_temporal = {
        'Archivo'   : archivo,
        'ElementId' : elem_id_str,
        'Planilla'  : clave_planilla,
        'CodIntBIM' : cod_val,
        'Headers'   : headers_order,
        'Row'       : resultado,
    }

    if not save_json(json_temporal, REPO_ELEMENTO_TMP):
        return

    if not os.path.isfile(CPYTHON_VIEWER_PATH):
        MessageBox.Show(
            u"No se encontro el visor:\n{}".format(CPYTHON_VIEWER_PATH),
            u"Error")
        return

    if not PYTHON_EXE or not os.path.isfile(PYTHON_EXE):
        MessageBox.Show(
            u"No se encontro Python 3 en este equipo.", u"Error")
        return

    try:
        subprocess.Popen(
            [PYTHON_EXE, CPYTHON_VIEWER_PATH, REPO_ELEMENTO_TMP],
            shell=False
        )
    except Exception as e:
        MessageBox.Show(
            u"No se pudo ejecutar visor externo:\n{}".format(e), u"Error")


if __name__ == '__main__':
    main()
