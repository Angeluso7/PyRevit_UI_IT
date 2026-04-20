# -*- coding: utf-8 -*-
__title__   = "planilla vs modelo"
__doc__     = """Version = 1.1
Date    = 20.04.2026
Cambios v1.1:
- [FIX] Python detectado automaticamente (sin ruta hardcodeada).
- [FIX] Rutas resueltas desde lib/config/paths.py (portátil).
- [FIX] config_proyecto_activo busca master/ luego data/.
- [FIX] Repo activo soporta ruta directa y NUP portátil.
- [FIX] Eliminado bloque basura kit_button_clicked.
- [FIX] main() compatible con pyRevit (sin __main__).
Author: Argenis Angel"""

import os
import sys
import subprocess
from datetime import datetime

import clr
from pyrevit import forms

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    RevitLinkInstance,
)

# Documento activo
try:
    doc = __revit__.ActiveUIDocument.Document
except Exception:
    doc = None

# --- Agregar lib/ al sys.path ---
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_EXT_ROOT   = os.path.normpath(os.path.join(_SCRIPT_DIR, '..', '..', '..', '..'))
_LIB_DIR    = os.path.join(_EXT_ROOT, 'lib')
if _LIB_DIR not in sys.path:
    sys.path.insert(0, _LIB_DIR)

# --- Importar rutas centralizadas ---
try:
    from config.paths import (
        DATA_DIR        as DATA_DIR_EXT,
        MASTER_DIR,
        TEMP_DIR,
        LOG_DIR,
        ensure_runtime_dirs,
        get_ruta_repositorio,
    )
    from config_utils import get_config_path, load_config
except Exception as _e:
    forms.alert(
        'No se pudo importar lib/config.\nVerifica que lib/ este dentro de la extension.\n\n{}'.format(_e),
        title='Error de importacion'
    )
    raise SystemExit

ensure_runtime_dirs()

# --- Rutas de archivos ---
LOG_PATH              = os.path.join(LOG_DIR,      'planilla_vs_modelo_log.txt')
CONFIG_PROYECTO_ACTIVO = get_config_path()
LEER_XLSM             = os.path.join(DATA_DIR_EXT, 'leer_xlsm_codigos.py')
FORMATEAR_XLSX        = os.path.join(DATA_DIR_EXT, 'formatear_tablas_excel.py')
UI_COMPARACION        = os.path.join(DATA_DIR_EXT, 'ui_comparacion.py')
SCRIPT_JSON_PATH      = os.path.join(MASTER_DIR,   'script.json')
MODELO_JSON           = os.path.join(DATA_DIR_EXT, 'modelo_codint_por_cm.json')
HEADERS_JSON          = os.path.join(DATA_DIR_EXT, 'headers_por_tabla.json')
PLANILLAS_HEADERS_JSON = os.path.join(DATA_DIR_EXT,'planillas_headers_order.json')
DATA_COMPARACION      = os.path.join(TEMP_DIR,     'comparacion')
if not os.path.exists(DATA_COMPARACION):
    os.makedirs(DATA_COMPARACION)
CREATE_NO_WINDOW = 0x08000000

# --- Deteccion automatica de Python 3 ---
def _detectar_python3():
    try:
        cfg = load_config()
        ruta_cfg = (cfg.get('python_exe') or '').strip()
        if ruta_cfg and os.path.isfile(ruta_cfg):
            return ruta_cfg
    except Exception:
        pass
    local_programs = os.path.join(
        os.getenv('LOCALAPPDATA', os.path.expanduser('~')),
        'Programs', 'Python'
    )
    if os.path.isdir(local_programs):
        versiones = sorted(
            [d for d in os.listdir(local_programs) if d.startswith('Python')],
            reverse=True
        )
        for v in versiones:
            exe = os.path.join(local_programs, v, 'python.exe')
            if os.path.isfile(exe):
                return exe
    for candidato in ('python3', 'python'):
        try:
            out = subprocess.check_output(
                [candidato, '--version'],
                stderr=subprocess.STDOUT,
                creationflags=CREATE_NO_WINDOW
            )
            if b'Python 3' in out:
                return candidato
        except Exception:
            continue
    return None

PYTHON3_EXE = _detectar_python3()

# --- Logging ---
def log(msg):
    try:
        ts = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        with open(LOG_PATH, 'a', encoding='utf-8') as f:
            f.write(u'[{}] {}\n'.format(ts, msg))
    except Exception:
        pass

# --- Utilidades JSON ---
def cargar_json(ruta, default):
    import json
    try:
        if not os.path.exists(ruta):
            log('cargar_json: no existe -> {}'.format(ruta))
            return default
        with open(ruta, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        log('cargar_json: error {} -> {}'.format(ruta, e))
        return default

def get_repo_activo_path():
    try:
        cfg = load_config()
    except Exception as e:
        log('get_repo_activo_path: error config -> {}'.format(e))
        return ''
    ruta = (cfg.get('ruta_repositorio_activo') or '').strip()
    if ruta:
        log('get_repo_activo_path (directo): {}'.format(ruta))
        return ruta
    nup = (cfg.get('nup_activo') or cfg.get('nup') or '').strip()
    if nup:
        ruta = get_ruta_repositorio(nup)
        log('get_repo_activo_path (nup): {}'.format(ruta))
        return ruta
    log('get_repo_activo_path: sin ruta ni nup')
    return ''

def cargar_repo_activo():
    ruta = get_repo_activo_path()
    if not ruta or not os.path.exists(ruta):
        log('cargar_repo_activo: no disponible -> {}'.format(ruta))
        return {}
    repo = cargar_json(ruta, {})
    log('cargar_repo_activo: {} registros'.format(len(repo)))
    return repo

# --- Utilidades Revit ---
def get_all_docs_with_links():
    docs = []
    try:
        if doc is None:
            return []
        docs.append((doc, doc.PathName or ''))
        col = FilteredElementCollector(doc).OfClass(RevitLinkInstance)
        seen = set()
        for li in col:
            try:
                link_doc = li.GetLinkDocument()
                if not link_doc:
                    continue
                path = link_doc.PathName or ''
                if path in seen:
                    continue
                seen.add(path)
                docs.append((link_doc, path))
                log('link doc: {}'.format(path))
            except Exception as e:
                log('error RevitLinkInstance: {}'.format(e))
    except Exception as e:
        forms.alert('Error obteniendo links:\n{}'.format(e), title='Error links')
        log('get_all_docs_with_links error: {}'.format(e))
    return docs

def cargar_planillas_headers():
    data = cargar_json(PLANILLAS_HEADERS_JSON, {})
    log('planillas_headers: {} claves'.format(len(data)))
    return data

# --- Generar modelo JSON desde Revit ---
def generar_modelo_json_desde_revit(script_json_path, modelo_json_path):
    import json
    if doc is None:
        forms.alert('No hay documento activo.', title='Error')
        return False
    script_data = cargar_json(script_json_path, {})
    codigos_planillas = script_data.get('codigos_planillas', {})
    if not codigos_planillas:
        forms.alert('script.json no tiene codigos_planillas.', title='Error')
        return False
    cm_to_planilla = {}
    for k, v in codigos_planillas.items():
        try:
            if isinstance(v, str) and len(v) == 4 and v.startswith('CM'):
                cm_to_planilla[v] = k
            elif isinstance(k, str) and len(k) == 4 and k.startswith('CM'):
                cm_to_planilla[k] = v
        except Exception as e:
            log('cm_to_planilla error: {}'.format(e))
    planillas_headers = cargar_planillas_headers()
    repo_activo = cargar_repo_activo()
    indice_bd_eid = {}
    indice_bd_cod = {}
    for k, v in repo_activo.items():
        cod     = (v.get('CodIntBIM') or '').strip()
        eid     = str(v.get('ElementId', '')).strip()
        archivo = (v.get('Archivo') or '').strip()
        if cod:
            indice_bd_cod.setdefault(cod, []).append(v)
        if eid and archivo:
            indice_bd_eid[(archivo, eid)] = v

    def headers_para_cm(cm):
        nombre = cm_to_planilla.get(cm)
        h = None
        if nombre:
            h = planillas_headers.get('{}::{}'.format(nombre, cm))
        if not h:
            h = planillas_headers.get(cm)
        return list(h) if h else []

    datos_por_cm = {}
    usados_cod   = set()
    usados_eid   = set()
    total_docs   = 0
    total_elems  = 0
    try:
        for d, ruta_archivo in get_all_docs_with_links():
            total_docs += 1
            try:
                col = FilteredElementCollector(d).WhereElementIsNotElementType()
            except Exception as e:
                log('error collector {}: {}'.format(ruta_archivo, e))
                continue
            for el in col:
                try:
                    p = el.LookupParameter('CodIntBIM')
                    if not p:
                        continue
                    codint = (p.AsString() or '').strip()
                    if not codint or len(codint) < 4:
                        continue
                    total_elems += 1
                    cm  = codint[:4]
                    eid = str(el.Id.IntegerValue)
                    arc = ruta_archivo or ''
                    fila_bd = indice_bd_eid.get((arc, eid))
                    if not fila_bd:
                        lst = indice_bd_cod.get(codint) or []
                        fila_bd = lst[0] if lst else None
                    if fila_bd:
                        usados_cod.add(codint)
                        usados_eid.add((arc, eid))
                        fila = dict(fila_bd)
                    else:
                        fila = {}
                    fila['CodIntBIM'] = codint
                    fila['ElementId'] = eid
                    fila['Archivo']   = arc
                    try:
                        fila['Categoria'] = el.Category.Name if el.Category else ''
                    except Exception:
                        fila['Categoria'] = ''
                    for h in headers_para_cm(cm):
                        if h == 'CodIntBIM':
                            continue
                        try:
                            pe = el.LookupParameter(h)
                            vm = (pe.AsString() or pe.AsValueString() or '') if pe else ''
                        except Exception:
                            vm = ''
                        vb = fila.get(h)
                        fila[h] = vb if vb not in (None, '') else vm
                    datos_por_cm.setdefault(cm, []).append(fila)
                except Exception as e:
                    log('error elem {}: {}'.format(ruta_archivo, e))
    except Exception as e:
        forms.alert('Error recorriendo modelo:\n{}'.format(e), title='Error modelo')
        log('generar_modelo error: {}'.format(e))
        return False

    # Registros BD no encontrados en modelo
    for k, v in repo_activo.items():
        cod = (v.get('CodIntBIM') or '').strip()
        if not cod or len(cod) < 4:
            continue
        eid = str(v.get('ElementId', '')).strip()
        arc = (v.get('Archivo') or '').strip()
        if (arc, eid) in usados_eid or cod in usados_cod:
            continue
        fila = dict(v)
        fila.setdefault('CodIntBIM', cod)
        fila.setdefault('ElementId', eid)
        fila.setdefault('Archivo',   arc)
        fila.setdefault('Categoria', '')
        datos_por_cm.setdefault(cod[:4], []).append(fila)

    log('modelo: docs={} elems={} CMs={}'.format(total_docs, total_elems, list(datos_por_cm.keys())))
    try:
        with open(modelo_json_path, 'w', encoding='utf-8') as f:
            json.dump(datos_por_cm, f, ensure_ascii=False, indent=2)
        log('modelo guardado: {}'.format(modelo_json_path))
    except Exception as e:
        forms.alert('Error guardando modelo JSON:\n{}'.format(e), title='Error')
        log('error guardando modelo: {}'.format(e))
        return False
    forms.alert(
        'Modelo generado:\n{}\n\nDocs: {} | CodIntBIM: {}'.format(
            modelo_json_path, total_docs, total_elems),
        title='Modelo generado'
    )
    return True

# --- Flujo Excel ---
def seleccionar_xlsm():
    try:
        ruta = forms.pick_file(file_ext='xlsm', multi_file=False, title='Selecciona la planilla .xlsm')
        log('xlsm seleccionado: {}'.format(ruta or '<cancelado>'))
        return ruta
    except Exception as e:
        forms.alert('Error seleccionando xlsm:\n{}'.format(e), title='Error')
        return None

def llamar_leer_xlsm(ruta_xlsm):
    if not PYTHON3_EXE:
        forms.alert(
            'No se encontro Python 3 en este equipo.\n'
            'Instala Python 3 o agrega la clave python_exe en config_proyecto_activo.json.',
            title='Python no encontrado'
        )
        return None
    try:
        salida = subprocess.check_output(
            [PYTHON3_EXE, LEER_XLSM, ruta_xlsm, DATA_COMPARACION],
            stderr=subprocess.STDOUT,
            universal_newlines=True,
            creationflags=CREATE_NO_WINDOW
        )
        log('leer_xlsm salida: {}'.format(salida))
        lineas = [l.strip() for l in salida.splitlines() if l.strip()]
        if not lineas:
            forms.alert('leer_xlsm no devolvio ruta.\n{}'.format(salida), title='Error')
            return None
        csv_path = lineas[-1]
        if not os.path.exists(csv_path):
            forms.alert('CSV no generado.\n{}'.format(salida), title='Error')
            return None
        return csv_path
    except subprocess.CalledProcessError as e:
        forms.alert('Error leer_xlsm:\n{}'.format(e.output), title='Error')
        return None
    except Exception as e:
        forms.alert('Error inesperado leer_xlsm:\n{}'.format(e), title='Error')
        return None

def llamar_ui_y_formato(ruta_xlsm, csv_codigos):
    stamp = datetime.now().strftime('%Y%m%d_%H%M')
    ruta_xlsx_salida = os.path.join(
        os.path.dirname(ruta_xlsm),
        'planilla-modelo_{}.xlsx'.format(stamp)
    )
    try:
        subprocess.check_call(
            [
                PYTHON3_EXE, UI_COMPARACION,
                SCRIPT_JSON_PATH,
                csv_codigos,
                DATA_COMPARACION,
                FORMATEAR_XLSX,
                ruta_xlsx_salida,
                PYTHON3_EXE,
                MODELO_JSON,
                HEADERS_JSON
            ],
            stderr=subprocess.STDOUT,
            creationflags=CREATE_NO_WINDOW
        )
        forms.alert('Archivo generado:\n{}'.format(ruta_xlsx_salida), title='Exito')
        log('xlsx generado: {}'.format(ruta_xlsx_salida))
    except subprocess.CalledProcessError as e:
        forms.alert('Error UI/formateo:\n{}'.format(e), title='Error')
        log('error ui_formato: {}'.format(e))
    except Exception as e:
        forms.alert('Error inesperado UI/formateo:\n{}'.format(e), title='Error')
        log('error inesperado ui_formato: {}'.format(e))

# --- MAIN (compatible pyRevit) ---
def main():
    log('==== Inicio Planilla vs Modelo v1.1 ====')
    if doc is None:
        forms.alert('No hay documento activo.', title='Error')
        return
    if not os.path.exists(SCRIPT_JSON_PATH):
        forms.alert('No se encontro script.json:\n{}'.format(SCRIPT_JSON_PATH), title='Error')
        return
    if not os.path.exists(LEER_XLSM) or not os.path.exists(UI_COMPARACION):
        forms.alert(
            'Faltan scripts CPython en data/:\n{}\n{}'.format(LEER_XLSM, UI_COMPARACION),
            title='Error'
        )
        return
    ok = generar_modelo_json_desde_revit(SCRIPT_JSON_PATH, MODELO_JSON)
    if not ok:
        return
    ruta_xlsm = seleccionar_xlsm()
    if not ruta_xlsm:
        return
    csv_codigos = llamar_leer_xlsm(ruta_xlsm)
    if not csv_codigos:
        return
    llamar_ui_y_formato(ruta_xlsm, csv_codigos)
    log('==== Fin Planilla vs Modelo ====')

main()