# -*- coding: utf-8 -*-
__title__ = "planilla vs modelo"
__doc__ = """Version = 1.6
Date = 20.04.2026
_______________________________________________________________
Description:

Comparación de Planilla con datos y elementos de modelo en función
de la información existente y la faltante. Si no existe
modelo_codint_por_cm.json, lo genera automáticamente recorriendo
el modelo activo y sus links, tomando los parámetros definidos en
planillas_headers_order.json por cada CMxx.

________________________________________________________________
Cambios v1.6:
- [FIX] Eliminadas rutas hardcodeadas a AppData\\Roaming\\MyPyRevitExtention.
- [FIX] Rutas resueltas desde la extensión usando lib/config.
- [FIX] Repositorio activo portable usando get_ruta_repositorio(nup) como fallback.
- [FIX] Flujo continuo: selección xlsm -> análisis -> UI -> exportación.
- [FIX] Selector de carpeta de salida para el xlsx final.
- [FIX] Logging robusto de rutas y validaciones previas.
- [FIX] Compatibilidad con ui_comparacion.py por argumentos portables.
- [UI] Fuerza modo oscuro de pyRevit cuando está disponible.

________________________________________________________________
Author: Argenis Angel
"""

import os
import sys
import subprocess
from datetime import datetime

import clr
from pyrevit import forms

try:
    from pyrevit.userconfig import user_config
    user_config.core.darkmode = True
    user_config.save_changes()
except Exception:
    pass

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    RevitLinkInstance,
)

try:
    doc = __revit__.ActiveUIDocument.Document
except Exception:
    doc = None


# -----------------------------------------------------------------------------
# Resolver lib/ portable
# -----------------------------------------------------------------------------
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_EXT_ROOT = os.path.normpath(os.path.join(_SCRIPT_DIR, "..", "..", "..", ".."))
_LIB_DIR = os.path.join(_EXT_ROOT, "lib")
if _LIB_DIR not in sys.path:
    sys.path.insert(0, _LIB_DIR)

try:
    from config.paths import (
        EXT_ROOT,
        MASTER_DIR,
        TEMP_DIR,
        LOG_DIR,
        ensure_runtime_dirs,
        get_ruta_repositorio,
    )
    from config_utils import load_config
except Exception as _e:
    forms.alert(
        u"No se pudo importar lib/config.\n"
        u"Verifica que la carpeta 'lib' exista dentro de la extensión.\n\n{}".format(_e),
        title=u"Error de importación"
    )
    raise SystemExit

ensure_runtime_dirs()


# -----------------------------------------------------------------------------
# Rutas de trabajo
# -----------------------------------------------------------------------------
_CPYTHON_DIR = os.path.join(EXT_ROOT, "scripts_cpython")
DATA_COMPARACION = os.path.join(TEMP_DIR, "comparacion")

if not os.path.exists(DATA_COMPARACION):
    os.makedirs(DATA_COMPARACION)

LOG_PATH = os.path.join(LOG_DIR, "planilla_vs_modelo_log.txt")

CONFIG_PROYECTO_ACTIVO = os.path.join(MASTER_DIR, "config_proyecto_activo.json")
SCRIPT_JSON_PATH = os.path.join(MASTER_DIR, "script.json")
PLANILLAS_HEADERS_JSON = os.path.join(MASTER_DIR, "planillas_headers_order.json")

LEER_XLSM = os.path.join(_CPYTHON_DIR, "leer_xlsm_codigos.py")
UI_COMPARACION = os.path.join(_CPYTHON_DIR, "ui_comparacion.py")
FORMATEAR_XLSX = os.path.join(_CPYTHON_DIR, "formatear_tablas_excel_v2.py")

MODELO_JSON = os.path.join(DATA_COMPARACION, "modelo_codint_por_cm.json")
HEADERS_JSON = os.path.join(DATA_COMPARACION, "headers_por_tabla.json")

CREATE_NO_WINDOW = 0x08000000


# -----------------------------------------------------------------------------
# Logging
# -----------------------------------------------------------------------------
def log(msg):
    try:
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(u"[{}] {}\n".format(ts, msg))
    except Exception:
        pass


def log_paths():
    log(u"--- RUTAS PLANILLA VS MODELO ---")
    log(u"EXT_ROOT: {}".format(EXT_ROOT))
    log(u"_CPYTHON_DIR: {} | existe={}".format(_CPYTHON_DIR, os.path.exists(_CPYTHON_DIR)))
    log(u"MASTER_DIR: {} | existe={}".format(MASTER_DIR, os.path.exists(MASTER_DIR)))
    log(u"TEMP_DIR: {} | existe={}".format(TEMP_DIR, os.path.exists(TEMP_DIR)))
    log(u"DATA_COMPARACION: {} | existe={}".format(DATA_COMPARACION, os.path.exists(DATA_COMPARACION)))
    log(u"SCRIPT_JSON_PATH: {} | existe={}".format(SCRIPT_JSON_PATH, os.path.exists(SCRIPT_JSON_PATH)))
    log(u"PLANILLAS_HEADERS_JSON: {} | existe={}".format(PLANILLAS_HEADERS_JSON, os.path.exists(PLANILLAS_HEADERS_JSON)))
    log(u"LEER_XLSM: {} | existe={}".format(LEER_XLSM, os.path.exists(LEER_XLSM)))
    log(u"UI_COMPARACION: {} | existe={}".format(UI_COMPARACION, os.path.exists(UI_COMPARACION)))
    log(u"FORMATEAR_XLSX: {} | existe={}".format(FORMATEAR_XLSX, os.path.exists(FORMATEAR_XLSX)))
    log(u"MODELO_JSON: {}".format(MODELO_JSON))
    log(u"HEADERS_JSON: {}".format(HEADERS_JSON))
    log(u"LOG_PATH: {}".format(LOG_PATH))
    log(u"--- FIN RUTAS ---")


# -----------------------------------------------------------------------------
# JSON helpers
# -----------------------------------------------------------------------------
def cargar_json(ruta, default=None):
    import json
    if default is None:
        default = {}
    try:
        if not os.path.exists(ruta):
            log(u"cargar_json: no existe -> {}".format(ruta))
            return default
        with open(ruta, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data
    except Exception as e:
        log(u"cargar_json: error leyendo {} -> {}".format(ruta, e))
        return default


def guardar_json(ruta, data):
    import json
    try:
        base = os.path.dirname(ruta)
        if base and not os.path.exists(base):
            os.makedirs(base)
        with open(ruta, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        log(u"guardar_json: error guardando {} -> {}".format(ruta, e))
        return False


# -----------------------------------------------------------------------------
# Repositorio activo portable
# -----------------------------------------------------------------------------
def get_repo_activo_path():
    try:
        cfg = load_config()
    except Exception as e:
        log(u"get_repo_activo_path: error load_config -> {}".format(e))
        return ""

    ruta = (cfg.get("ruta_repositorio_activo") or "").strip()
    if ruta:
        if os.path.exists(ruta):
            log(u"get_repo_activo_path: ruta absoluta OK -> {}".format(ruta))
            return ruta
        else:
            log(u"get_repo_activo_path: ruta absoluta no existe -> {}".format(ruta))

    nup = (cfg.get("nup_activo") or cfg.get("nup") or "").strip()
    if nup:
        try:
            ruta_rel = get_ruta_repositorio(nup)
            if ruta_rel and os.path.exists(ruta_rel):
                log(u"get_repo_activo_path: ruta portable OK -> {}".format(ruta_rel))
                return ruta_rel
            else:
                log(u"get_repo_activo_path: ruta portable no existe -> {}".format(ruta_rel))
        except Exception as e:
            log(u"get_repo_activo_path: error get_ruta_repositorio({}) -> {}".format(nup, e))

    log(u"get_repo_activo_path: sin ruta disponible")
    return ""


def cargar_repo_activo():
    ruta = get_repo_activo_path()
    if not ruta or not os.path.exists(ruta):
        log(u"cargar_repo_activo: no disponible -> {}".format(ruta))
        return {}
    repo = cargar_json(ruta, {})
    log(u"cargar_repo_activo: {} registros desde {}".format(len(repo), ruta))
    return repo


# -----------------------------------------------------------------------------
# Detección Python 3
# -----------------------------------------------------------------------------
def _detectar_python3():
    try:
        cfg = load_config()
        ruta_cfg = (cfg.get("python_exe") or "").strip()
        if ruta_cfg and os.path.isfile(ruta_cfg):
            log(u"python3 detectado desde config: {}".format(ruta_cfg))
            return ruta_cfg
    except Exception:
        pass

    local_programs = os.path.join(
        os.getenv("LOCALAPPDATA", os.path.expanduser("~")),
        "Programs",
        "Python"
    )

    if os.path.isdir(local_programs):
        versiones = sorted(
            [d for d in os.listdir(local_programs) if d.startswith("Python")],
            reverse=True
        )
        for v in versiones:
            exe = os.path.join(local_programs, v, "python.exe")
            if os.path.isfile(exe):
                log(u"python3 detectado en Programs: {}".format(exe))
                return exe

    for candidato in ("python3", "python"):
        try:
            out = subprocess.check_output(
                [candidato, "--version"],
                stderr=subprocess.STDOUT,
                creationflags=CREATE_NO_WINDOW
            )
            try:
                txt = out.decode("utf-8", "ignore")
            except Exception:
                txt = str(out)
            if "Python 3" in txt:
                log(u"python3 detectado en PATH: {}".format(candidato))
                return candidato
        except Exception:
            continue

    log(u"python3 NO encontrado")
    return None


PYTHON3_EXE = _detectar_python3()


# -----------------------------------------------------------------------------
# Utilidades Revit
# -----------------------------------------------------------------------------
def get_all_docs_with_links():
    docs = []
    try:
        if doc is None:
            log(u"get_all_docs_with_links: doc es None")
            return []

        docs.append((doc, doc.PathName or ""))
        col = FilteredElementCollector(doc).OfClass(RevitLinkInstance)
        seen = set()

        for li in col:
            try:
                link_doc = li.GetLinkDocument()
                if not link_doc:
                    continue
                path = link_doc.PathName or ""
                if path in seen:
                    continue
                seen.add(path)
                docs.append((link_doc, path))
            except Exception as e:
                log(u"get_all_docs_with_links: error RevitLinkInstance -> {}".format(e))
    except Exception as e:
        forms.alert(u"Error obteniendo documentos vinculados:\n{}".format(e), title=u"Error")
        log(u"get_all_docs_with_links: excepción general -> {}".format(e))

    return docs


def cargar_planillas_headers():
    data = cargar_json(PLANILLAS_HEADERS_JSON, {})
    log(u"cargar_planillas_headers: {} claves".format(len(data)))
    return data


# -----------------------------------------------------------------------------
# Generación modelo JSON
# -----------------------------------------------------------------------------
def generar_modelo_json_desde_revit(script_json_path, modelo_json_path):
    if doc is None:
        forms.alert(u"No hay documento activo.", title=u"Error")
        return False

    script_data = cargar_json(script_json_path, {})
    codigos_planillas = script_data.get("codigos_planillas", {})
    if not codigos_planillas:
        forms.alert(u"script.json no contiene 'codigos_planillas'.", title=u"Error script.json")
        log(u"generar_modelo_json_desde_revit: sin codigos_planillas")
        return False

    cm_to_planilla = {}
    for k, v in codigos_planillas.items():
        try:
            if isinstance(v, str) and len(v) == 4 and v.startswith("CM"):
                cm_to_planilla[v] = k
            elif isinstance(k, str) and len(k) == 4 and k.startswith("CM"):
                cm_to_planilla[k] = v
        except Exception as e:
            log(u"cm_to_planilla error: {}".format(e))

    planillas_headers = cargar_planillas_headers()
    repo_activo = cargar_repo_activo()

    indice_bd_eid = {}
    indice_bd_cod = {}
    for _, v in repo_activo.items():
        cod = (v.get("CodIntBIM") or "").strip()
        eid = str(v.get("ElementId", "")).strip()
        archivo = (v.get("Archivo") or "").strip()
        if cod:
            indice_bd_cod.setdefault(cod, []).append(v)
        if eid and archivo:
            indice_bd_eid[(archivo, eid)] = v

    def headers_para_cm(cm):
        nombre = cm_to_planilla.get(cm)
        headers = None
        if nombre:
            headers = planillas_headers.get("{}::{}".format(nombre, cm))
        if not headers:
            headers = planillas_headers.get(cm)
        return list(headers) if headers else []

    datos_por_cm = {}
    usados_cod = set()
    usados_eid = set()
    total_docs = 0
    total_elems = 0

    try:
        for d, ruta_archivo in get_all_docs_with_links():
            total_docs += 1
            try:
                col = FilteredElementCollector(d).WhereElementIsNotElementType()
            except Exception as e:
                log(u"error collector {} -> {}".format(ruta_archivo, e))
                continue

            for el in col:
                try:
                    p = el.LookupParameter("CodIntBIM")
                    if not p:
                        continue
                    codint = (p.AsString() or "").strip()
                    if not codint or len(codint) < 4:
                        continue

                    total_elems += 1
                    cm = codint[:4]
                    eid = str(el.Id.IntegerValue)
                    arc = ruta_archivo or ""

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

                    fila["CodIntBIM"] = codint
                    fila["ElementId"] = eid
                    fila["Archivo"] = arc

                    try:
                        fila["Categoria"] = el.Category.Name if el.Category else ""
                    except Exception:
                        fila["Categoria"] = ""

                    for h in headers_para_cm(cm):
                        if h == "CodIntBIM":
                            continue
                        try:
                            pe = el.LookupParameter(h)
                            vm = (pe.AsString() or pe.AsValueString() or "") if pe else ""
                        except Exception:
                            vm = ""
                        vb = fila.get(h)
                        fila[h] = vb if vb not in (None, "") else vm

                    datos_por_cm.setdefault(cm, []).append(fila)

                except Exception as e:
                    log(u"error procesando elemento en {} -> {}".format(ruta_archivo, e))

    except Exception as e:
        forms.alert(u"Error recorriendo modelo:\n{}".format(e), title=u"Error modelo")
        log(u"generar_modelo_json_desde_revit: excepción general -> {}".format(e))
        return False

    for _, v in repo_activo.items():
        cod = (v.get("CodIntBIM") or "").strip()
        if not cod or len(cod) < 4:
            continue
        eid = str(v.get("ElementId", "")).strip()
        arc = (v.get("Archivo") or "").strip()
        if (arc, eid) in usados_eid or cod in usados_cod:
            continue

        fila = dict(v)
        fila.setdefault("CodIntBIM", cod)
        fila.setdefault("ElementId", eid)
        fila.setdefault("Archivo", arc)
        fila.setdefault("Categoria", "")
        datos_por_cm.setdefault(cod[:4], []).append(fila)

    log(u"modelo generado: docs={} elems={} cms={}".format(
        total_docs, total_elems, sorted(datos_por_cm.keys())
    ))

    if not guardar_json(modelo_json_path, datos_por_cm):
        forms.alert(u"Error guardando modelo JSON:\n{}".format(modelo_json_path), title=u"Error")
        return False

    log(u"modelo JSON guardado -> {}".format(modelo_json_path))
    return True


# -----------------------------------------------------------------------------
# Selección planilla
# -----------------------------------------------------------------------------
def seleccionar_xlsm():
    try:
        ruta = forms.pick_file(
            file_ext="xlsm",
            multi_file=False,
            title=u"Selecciona la planilla .xlsm"
        )
        log(u"seleccionar_xlsm: {}".format(ruta or ""))
        return ruta
    except Exception as e:
        forms.alert(u"Error seleccionando .xlsm:\n{}".format(e), title=u"Error")
        log(u"seleccionar_xlsm error -> {}".format(e))
        return None


def seleccionar_carpeta_salida():
    try:
        ruta = forms.pick_folder(
            title=u"Selecciona la carpeta donde guardar el archivo Excel exportado"
        )
        log(u"seleccionar_carpeta_salida: {}".format(ruta or ""))
        return ruta
    except Exception as e:
        forms.alert(u"Error seleccionando carpeta de salida:\n{}".format(e), title=u"Error")
        log(u"seleccionar_carpeta_salida error -> {}".format(e))
        return None


# -----------------------------------------------------------------------------
# Lectura xlsm -> csv
# -----------------------------------------------------------------------------
def llamar_leer_xlsm(ruta_xlsm):
    if not PYTHON3_EXE:
        forms.alert(
            u"No se encontró Python 3.\n"
            u"Agrega la clave 'python_exe' en config si es necesario.",
            title=u"Python no encontrado"
        )
        return None

    if not os.path.exists(LEER_XLSM):
        forms.alert(u"No se encontró leer_xlsm_codigos.py:\n{}".format(LEER_XLSM), title=u"Error")
        return None

    try:
        salida = subprocess.check_output(
            [PYTHON3_EXE, LEER_XLSM, ruta_xlsm, DATA_COMPARACION],
            stderr=subprocess.STDOUT,
            universal_newlines=True,
            creationflags=CREATE_NO_WINDOW
        )
        log(u"llamar_leer_xlsm salida:\n{}".format(salida))

        lineas = [l.strip() for l in salida.splitlines() if l.strip()]
        if not lineas:
            forms.alert(u"leer_xlsm_codigos.py no devolvió ninguna ruta.", title=u"Error")
            return None

        csv_path = lineas[-1]
        if not os.path.exists(csv_path):
            forms.alert(u"No se generó el CSV esperado.\n\n{}".format(csv_path), title=u"Error")
            return None

        log(u"csv generado -> {}".format(csv_path))
        return csv_path

    except subprocess.CalledProcessError as e:
        try:
            msg = e.output
        except Exception:
            msg = str(e)
        forms.alert(u"Error al leer .xlsm:\n{}".format(msg), title=u"Error")
        log(u"llamar_leer_xlsm CalledProcessError -> {}".format(msg))
        return None
    except Exception as e:
        forms.alert(u"Error inesperado al leer .xlsm:\n{}".format(e), title=u"Error")
        log(u"llamar_leer_xlsm excepción general -> {}".format(e))
        return None


# -----------------------------------------------------------------------------
# UI comparacion
# -----------------------------------------------------------------------------
def llamar_ui_comparacion(ruta_xlsm, carpeta_destino):
    if not PYTHON3_EXE:
        forms.alert(
            u"No se encontró Python 3.\n"
            u"Agrega la clave 'python_exe' en config si es necesario.",
            title=u"Python no encontrado"
        )
        return

    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    ruta_xlsx_salida = os.path.join(
        carpeta_destino,
        u"planilla-modelo_{}.xlsx".format(stamp)
    )

    cmd = [
        PYTHON3_EXE,
        UI_COMPARACION,
        SCRIPT_JSON_PATH,
        ruta_xlsm,
        DATA_COMPARACION,
        FORMATEAR_XLSX,
        ruta_xlsx_salida,
        PYTHON3_EXE,
        MODELO_JSON,
        HEADERS_JSON
    ]

    log(u"llamar_ui_comparacion CMD:\n {}".format(u"\n ".join(cmd)))

    try:
        subprocess.Popen(
            cmd,
            stderr=subprocess.STDOUT,
            creationflags=CREATE_NO_WINDOW
        )
        log(u"ui_comparacion lanzado correctamente")
    except Exception as e:
        forms.alert(u"Error lanzando ui_comparacion:\n{}".format(e), title=u"Error")
        log(u"llamar_ui_comparacion error -> {}".format(e))


# -----------------------------------------------------------------------------
# Main
# -----------------------------------------------------------------------------
def main():
    log(u"==== Inicio Planilla vs Modelo v1.6 ====")
    log_paths()

    if doc is None:
        forms.alert(u"No hay documento activo en Revit.", title=u"Error")
        return

    if not os.path.exists(SCRIPT_JSON_PATH):
        forms.alert(
            u"No se encontró script.json:\n{}".format(SCRIPT_JSON_PATH),
            title=u"Error"
        )
        return

    faltantes = [p for p in (LEER_XLSM, UI_COMPARACION, FORMATEAR_XLSX) if not os.path.exists(p)]
    if faltantes:
        forms.alert(
            u"Faltan scripts en scripts_cpython:\n\n{}".format(u"\n".join(faltantes)),
            title=u"Scripts no encontrados"
        )
        return

    if not PYTHON3_EXE:
        forms.alert(
            u"No se encontró Python 3.\n"
            u"Agrega la clave 'python_exe' en config_proyecto_activo.json.",
            title=u"Python no encontrado"
        )
        return

    ruta_xlsm = seleccionar_xlsm()
    if not ruta_xlsm:
        log(u"Cancelado: no se seleccionó xlsm")
        return

    carpeta_destino = seleccionar_carpeta_salida()
    if not carpeta_destino:
        log(u"Cancelado: no se seleccionó carpeta destino")
        return

    ruta_repo = get_repo_activo_path()
    if not ruta_repo:
        continuar = forms.alert(
            u"No se encontró la ruta del repositorio activo.\n\n"
            u"La comparación se hará solo con los datos del modelo Revit,\n"
            u"sin cruzar información de la base de datos.\n\n"
            u"¿Deseas continuar de todas formas?",
            title=u"Repositorio no encontrado",
            options=[u"Continuar", u"Cancelar"]
        )
        if continuar != u"Continuar":
            log(u"Cancelado por usuario: repositorio no encontrado")
            return

    ok = generar_modelo_json_desde_revit(SCRIPT_JSON_PATH, MODELO_JSON)
    if not ok:
        log(u"Fallo generando MODELO_JSON")
        return

    csv_codigos = llamar_leer_xlsm(ruta_xlsm)
    if not csv_codigos:
        log(u"Fallo leyendo xlsm")
        return

    # ui_comparacion.py actual espera xlsm directo en argv[2], no csv
    llamar_ui_comparacion(ruta_xlsm, carpeta_destino)

    log(u"==== Fin Planilla vs Modelo v1.6 ====")


if __name__ == "__main__":
    main()