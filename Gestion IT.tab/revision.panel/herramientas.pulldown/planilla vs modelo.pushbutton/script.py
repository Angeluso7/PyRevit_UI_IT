# -*- coding: utf-8 -*-
__title__   = "planilla vs modelo"
__doc__     = """Version = 1.3
Date    = 20.04.2026
Cambios v1.3:
- [FIX] EXT_ROOT resuelto desde lib/config/paths.py (100% portable).
- [FIX] Scripts CPython apuntan a nombres reales del repo:
        ui_comparacion.py / formatear_tablas_excel_v2.py
- [FIX] leer_xlsm_codigos.py eliminado (no existe); lectura CSV
        delegada completamente a ui_comparacion.py.
- [FIX] MODELO_JSON y HEADERS_JSON en data/temp/comparacion/.
- [FIX] PLANILLAS_HEADERS_JSON en data/master/.
- [OPT] Validacion de scripts CPython con nombres reales antes de ejecutar.
- [OPT] Popen (no bloqueante) para ui_comparacion.py.
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

# ── Documento activo ──────────────────────────────────────────────────────────
try:
    doc = __revit__.ActiveUIDocument.Document
except Exception:
    doc = None

# ── Resolver EXT_ROOT usando lib/config/paths.py (fuente unica de verdad) ────
# El script esta en:
#   <EXT_ROOT>/Gestion IT.tab/revision.panel/herramientas.pulldown/planilla vs modelo.pushbutton/
# Son 4 niveles desde EXT_ROOT, pero NO usamos __file__ para contar ..
# sino que dejamos que paths.py lo resuelva solo desde su propia ubicacion.

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# Subir 4 niveles: pushbutton -> pulldown -> panel -> tab -> EXT_ROOT
_EXT_ROOT = os.path.normpath(
    os.path.join(_SCRIPT_DIR, '..', '..', '..', '..')
)
_LIB_DIR = os.path.join(_EXT_ROOT, 'lib')

if _LIB_DIR not in sys.path:
    sys.path.insert(0, _LIB_DIR)

# ── Importar rutas centralizadas ──────────────────────────────────────────────
try:
    from config.paths import (
        EXT_ROOT,           # <-- fuente definitiva, resuelto por paths.py
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
        u'No se pudo importar lib/config.\n'
        u'Verifica que lib/ este dentro de la extension.\n\n{}'.format(_e),
        title=u'Error de importacion'
    )
    raise SystemExit

ensure_runtime_dirs()

# ── Carpeta scripts_cpython (desde EXT_ROOT canonico) ────────────────────────
_CPYTHON_DIR = os.path.join(EXT_ROOT, 'scripts_cpython')

# Carpeta temporal para JSONs generados en runtime
DATA_COMPARACION = os.path.join(TEMP_DIR, 'comparacion')
if not os.path.exists(DATA_COMPARACION):
    os.makedirs(DATA_COMPARACION)

# ── Rutas de archivos ─────────────────────────────────────────────────────────
LOG_PATH               = os.path.join(LOG_DIR,          'planilla_vs_modelo_log.txt')
SCRIPT_JSON_PATH       = os.path.join(MASTER_DIR,       'script.json')
PLANILLAS_HEADERS_JSON = os.path.join(MASTER_DIR,       'planillas_headers_order.json')

# Scripts CPython — nombres exactos que existen en el repo
UI_COMPARACION  = os.path.join(_CPYTHON_DIR, 'ui_comparacion.py')
FORMATEAR_XLSX  = os.path.join(_CPYTHON_DIR, 'formatear_tablas_excel_v2.py')

# JSONs generados en runtime (data/temp/comparacion/)
MODELO_JSON  = os.path.join(DATA_COMPARACION, 'modelo_codint_por_cm.json')
HEADERS_JSON = os.path.join(DATA_COMPARACION, 'headers_por_tabla.json')

CREATE_NO_WINDOW = 0x08000000


# ── Logging ───────────────────────────────────────────────────────────────────
def log(msg):
    try:
        ts = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        with open(LOG_PATH, 'a', encoding='utf-8') as f:
            f.write(u'[{}] {}\n'.format(ts, msg))
    except Exception:
        pass


def log_paths():
    """Vuelca todas las rutas criticas al log para diagnostico."""
    log(u'--- RUTAS ---')
    log(u'EXT_ROOT        : {}'.format(EXT_ROOT))
    log(u'_CPYTHON_DIR    : {}'.format(_CPYTHON_DIR))
    log(u'DATA_DIR_EXT    : {}'.format(DATA_DIR_EXT))
    log(u'MASTER_DIR      : {}'.format(MASTER_DIR))
    log(u'TEMP_DIR        : {}'.format(TEMP_DIR))
    log(u'DATA_COMPARACION: {}'.format(DATA_COMPARACION))
    log(u'UI_COMPARACION  : {} | existe={}'.format(UI_COMPARACION,  os.path.exists(UI_COMPARACION)))
    log(u'FORMATEAR_XLSX  : {} | existe={}'.format(FORMATEAR_XLSX,  os.path.exists(FORMATEAR_XLSX)))
    log(u'SCRIPT_JSON     : {} | existe={}'.format(SCRIPT_JSON_PATH, os.path.exists(SCRIPT_JSON_PATH)))
    log(u'PLANILLAS_HDR   : {} | existe={}'.format(PLANILLAS_HEADERS_JSON, os.path.exists(PLANILLAS_HEADERS_JSON)))
    log(u'-------------')


# ── Utilidades JSON ───────────────────────────────────────────────────────────
def cargar_json(ruta, default):
    import json
    try:
        if not os.path.exists(ruta):
            log(u'cargar_json: no existe -> {}'.format(ruta))
            return default
        with open(ruta, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        log(u'cargar_json: error {} -> {}'.format(ruta, e))
        return default


# ── Repo activo (portable) ────────────────────────────────────────────────────
def get_repo_activo_path():
    """
    Prioridad:
      1. ruta_repositorio_activo en config si el archivo existe en disco.
      2. Ruta canonica data/proyectos/repositorio_datos_<nup>.json (portable).
    """
    try:
        cfg = load_config()
    except Exception as e:
        log(u'get_repo_activo_path: error config -> {}'.format(e))
        return ''

    ruta = (cfg.get('ruta_repositorio_activo') or '').strip()
    if ruta and os.path.exists(ruta):
        log(u'get_repo_activo_path (directo): {}'.format(ruta))
        return ruta

    nup = (cfg.get('nup_activo') or cfg.get('nup') or '').strip()
    if nup:
        ruta = get_ruta_repositorio(nup)
        log(u'get_repo_activo_path (nup portable): {}'.format(ruta))
        return ruta

    log(u'get_repo_activo_path: sin ruta ni nup definidos')
    return ''


def cargar_repo_activo():
    ruta = get_repo_activo_path()
    if not ruta or not os.path.exists(ruta):
        log(u'cargar_repo_activo: no disponible -> {}'.format(ruta))
        return {}
    repo = cargar_json(ruta, {})
    log(u'cargar_repo_activo: {} registros'.format(len(repo)))
    return repo


# ── Deteccion automatica de Python 3 ─────────────────────────────────────────
def _detectar_python3():
    # 1. Clave python_exe en config
    try:
        cfg = load_config()
        ruta_cfg = (cfg.get('python_exe') or '').strip()
        if ruta_cfg and os.path.isfile(ruta_cfg):
            log(u'python3 desde config: {}'.format(ruta_cfg))
            return ruta_cfg
    except Exception:
        pass

    # 2. LocalAppData/Programs/Python (instalacion tipica Windows)
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
                log(u'python3 desde LocalAppData: {}'.format(exe))
                return exe

    # 3. PATH del sistema
    for candidato in ('python3', 'python'):
        try:
            out = subprocess.check_output(
                [candidato, '--version'],
                stderr=subprocess.STDOUT,
                creationflags=CREATE_NO_WINDOW
            )
            if b'Python 3' in out:
                log(u'python3 desde PATH: {}'.format(candidato))
                return candidato
        except Exception:
            continue

    log(u'python3: no encontrado')
    return None


PYTHON3_EXE = _detectar_python3()


# ── Utilidades Revit ──────────────────────────────────────────────────────────
def get_all_docs_with_links():
    docs = []
    try:
        if doc is None:
            return []
        docs.append((doc, doc.PathName or ''))
        col  = FilteredElementCollector(doc).OfClass(RevitLinkInstance)
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
                log(u'link doc: {}'.format(path))
            except Exception as e:
                log(u'error RevitLinkInstance: {}'.format(e))
    except Exception as e:
        forms.alert(u'Error obteniendo links:\n{}'.format(e), title=u'Error links')
        log(u'get_all_docs_with_links error: {}'.format(e))
    return docs


def cargar_planillas_headers():
    data = cargar_json(PLANILLAS_HEADERS_JSON, {})
    log(u'planillas_headers: {} claves'.format(len(data)))
    return data


# ── Generar modelo JSON desde Revit ───────────────────────────────────────────
def generar_modelo_json_desde_revit(script_json_path, modelo_json_path):
    import json

    if doc is None:
        forms.alert(u'No hay documento activo.', title=u'Error')
        return False

    script_data       = cargar_json(script_json_path, {})
    codigos_planillas = script_data.get('codigos_planillas', {})
    if not codigos_planillas:
        forms.alert(u'script.json no tiene codigos_planillas.', title=u'Error')
        return False

    # Mapeo CM -> nombre planilla
    cm_to_planilla = {}
    for k, v in codigos_planillas.items():
        try:
            if isinstance(v, str) and len(v) == 4 and v.startswith('CM'):
                cm_to_planilla[v] = k
            elif isinstance(k, str) and len(k) == 4 and k.startswith('CM'):
                cm_to_planilla[k] = v
        except Exception as e:
            log(u'cm_to_planilla error: {}'.format(e))

    planillas_headers = cargar_planillas_headers()
    repo_activo       = cargar_repo_activo()

    # Indices BD para lookups rapidos
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
                log(u'error collector {}: {}'.format(ruta_archivo, e))
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
                        lst     = indice_bd_cod.get(codint) or []
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
                    log(u'error elem {}: {}'.format(ruta_archivo, e))

    except Exception as e:
        forms.alert(u'Error recorriendo modelo:\n{}'.format(e), title=u'Error modelo')
        log(u'generar_modelo error: {}'.format(e))
        return False

    # Registros BD que no aparecen en el modelo
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

    log(u'modelo: docs={} elems={} CMs={}'.format(
        total_docs, total_elems, list(datos_por_cm.keys())
    ))

    try:
        with open(modelo_json_path, 'w', encoding='utf-8') as f:
            json.dump(datos_por_cm, f, ensure_ascii=False, indent=2)
        log(u'modelo guardado: {}'.format(modelo_json_path))
    except Exception as e:
        forms.alert(u'Error guardando modelo JSON:\n{}'.format(e), title=u'Error')
        log(u'error guardando modelo: {}'.format(e))
        return False

    forms.alert(
        u'Modelo generado:\n{}\n\nDocs: {} | CodIntBIM encontrados: {}'.format(
            modelo_json_path, total_docs, total_elems
        ),
        title=u'Modelo generado'
    )
    return True


# ── Seleccion de planilla xlsm ────────────────────────────────────────────────
def seleccionar_xlsm():
    try:
        ruta = forms.pick_file(
            file_ext='xlsm',
            multi_file=False,
            title=u'Selecciona la planilla .xlsm'
        )
        log(u'xlsm seleccionado: {}'.format(ruta or u'<cancelado>'))
        return ruta
    except Exception as e:
        forms.alert(u'Error seleccionando xlsm:\n{}'.format(e), title=u'Error')
        return None


# ── Lanzar ui_comparacion.py (no bloqueante) ─────────────────────────────────
def llamar_ui_comparacion(ruta_xlsm):
    """
    Lanza ui_comparacion.py con todos los argumentos que necesita.
    Usa Popen para no bloquear pyRevit mientras la UI tkinter esta abierta.
    Argumentos esperados por ui_comparacion.py:
      argv[1] = script_json_path
      argv[2] = ruta_xlsm
      argv[3] = data_comparacion (carpeta temp)
      argv[4] = formatear_xlsx (script de formateo)
      argv[5] = ruta_xlsx_salida
      argv[6] = python3_exe
      argv[7] = modelo_json
      argv[8] = headers_json
    """
    if not PYTHON3_EXE:
        forms.alert(
            u'No se encontro Python 3 en este equipo.\n'
            u'Instala Python 3 o agrega la clave "python_exe"\n'
            u'en data/master/config_proyecto_activo.json.',
            title=u'Python no encontrado'
        )
        return

    stamp            = datetime.now().strftime('%Y%m%d_%H%M')
    ruta_xlsx_salida = os.path.join(
        os.path.dirname(ruta_xlsm),
        u'planilla-modelo_{}.xlsx'.format(stamp)
    )

    cmd = [
        PYTHON3_EXE, UI_COMPARACION,
        SCRIPT_JSON_PATH,
        ruta_xlsm,
        DATA_COMPARACION,
        FORMATEAR_XLSX,
        ruta_xlsx_salida,
        PYTHON3_EXE,
        MODELO_JSON,
        HEADERS_JSON,
    ]

    log(u'llamar_ui_comparacion CMD: {}'.format(' '.join(cmd)))

    try:
        proc = subprocess.Popen(
            cmd,
            stderr=subprocess.STDOUT,
            creationflags=CREATE_NO_WINDOW
        )
        log(u'ui_comparacion lanzado (PID {})'.format(proc.pid))
    except Exception as e:
        forms.alert(u'Error lanzando ui_comparacion:\n{}'.format(e), title=u'Error')
        log(u'llamar_ui_comparacion error: {}'.format(e))


# ── MAIN ──────────────────────────────────────────────────────────────────────
def main():
    log(u'==== Inicio Planilla vs Modelo v1.3 ====')
    log_paths()

    # 1. Verificar documento activo
    if doc is None:
        forms.alert(u'No hay documento activo en Revit.', title=u'Error')
        return

    # 2. Verificar script.json
    if not os.path.exists(SCRIPT_JSON_PATH):
        forms.alert(
            u'No se encontro script.json:\n{}'.format(SCRIPT_JSON_PATH),
            title=u'Error'
        )
        return

    # 3. Verificar scripts CPython con nombres reales
    scripts_faltantes = [
        s for s in (UI_COMPARACION, FORMATEAR_XLSX)
        if not os.path.exists(s)
    ]
    if scripts_faltantes:
        forms.alert(
            u'Faltan scripts en scripts_cpython/:\n\n{}'.format(
                u'\n'.join(scripts_faltantes)
            ),
            title=u'Scripts no encontrados'
        )
        return

    # 4. Verificar Python 3
    if not PYTHON3_EXE:
        forms.alert(
            u'No se encontro Python 3.\n'
            u'Agrega la clave "python_exe" en:\n{}'.format(
                os.path.join(MASTER_DIR, 'config_proyecto_activo.json')
            ),
            title=u'Python no encontrado'
        )
        return

    # 5. Generar modelo JSON desde Revit
    ok = generar_modelo_json_desde_revit(SCRIPT_JSON_PATH, MODELO_JSON)
    if not ok:
        return

    # 6. Seleccionar planilla xlsm
    ruta_xlsm = seleccionar_xlsm()
    if not ruta_xlsm:
        return

    # 7. Lanzar comparacion + UI
    llamar_ui_comparacion(ruta_xlsm)
    log(u'==== Fin Planilla vs Modelo v1.3 ====')


main()