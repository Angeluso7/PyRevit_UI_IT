# -*- coding: utf-8 -*-
__title__   = "Planilla vs\nModelo"
__doc__     = """Version = 2.1
Date    = 22.04.2026
________________________________________________________________
Description:
Comparación de planilla XLSM con el modelo Revit activo y sus links.
Genera modelo_codint_por_cm.json recorriendo el modelo y luego
lanza la UI de comparación (CPython3) para exportar el Excel.
________________________________________________________________
Author: Argenis Angel"""

# ── Imports ──────────────────────────────────────────────────
import os
import sys
import subprocess
import json
from datetime import datetime

import clr
from pyrevit import forms, script as pyscript

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    RevitLinkInstance,
)

# ── Documento activo ──────────────────────────────────────────
try:
    doc = __revit__.ActiveUIDocument.Document
except Exception:
    doc = None

# ══════════════════════════════════════════════════════════════
#  1. LOG (debe definirse PRIMERO antes de cualquier otra cosa)
# ══════════════════════════════════════════════════════════════
# Ruta temporal del log antes de conocer EXT_ROOT
_LOG_TEMP = os.path.join(
    os.environ.get("TEMP", os.path.expanduser("~")),
    "planilla_vs_modelo_log.txt"
)

def log(msg):
    try:
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ruta = globals().get("LOG_PATH", _LOG_TEMP)
        directorio = os.path.dirname(ruta)
        if directorio and not os.path.exists(directorio):
            os.makedirs(directorio)
        with open(ruta, "a", encoding="utf-8") as f:
            f.write(u"[{}] {}\n".format(ts, msg))
    except Exception:
        pass

# ══════════════════════════════════════════════════════════════
#  2. RAÍZ DE LA EXTENSIÓN (usa log ya definido)
# ══════════════════════════════════════════════════════════════
def _get_ext_root():
    """
    Obtiene la raíz de la extensión de forma robusta con 3 estrategias:
      1. Navegando desde __file__ (5 niveles arriba desde script.py)
      2. API nativa de pyRevit
      3. Fallback buscando carpetas conocidas en %APPDATA%
    """
    # Intento 1: navegar desde __file__
    # script.py → pushbutton → pulldown → panel → tab → ext_root
    try:
        ruta = os.path.abspath(__file__)
        for _ in range(5):
            ruta = os.path.dirname(ruta)
        # Verificar que tenga sentido (que exista data/master dentro)
        if os.path.isdir(os.path.join(ruta, "data", "master")):
            log("EXT_ROOT via __file__: {}".format(ruta))
            return ruta
        # Aunque no tenga data/master, si la carpeta existe la usamos
        if os.path.isdir(ruta):
            log("EXT_ROOT via __file__ (sin data/master): {}".format(ruta))
            return ruta
    except Exception as e:
        log("_get_ext_root intento 1 error: {}".format(e))

    # Intento 2: API de pyRevit
    try:
        from pyrevit import extensions as exts
        for ext in exts.get_installed_ui_extensions():
            if "PyRevitIT" in (ext.name or "") or "PyRevit_UI_IT" in (ext.directory or ""):
                log("EXT_ROOT via pyRevit API: {}".format(ext.directory))
                return ext.directory
    except Exception as e:
        log("_get_ext_root intento 2 error: {}".format(e))

    # Intento 3: fallback por nombres conocidos
    appdata = os.environ.get("APPDATA", "")
    carpetas_base = [
        os.path.join(appdata, "MyPyRevitExtention"),
        os.path.join(appdata, "pyRevit", "Extensions"),
        os.path.join(appdata, "pyRevit-Master", "Extensions"),
    ]
    nombres = [
        "PyRevitIT.extension",
        "PyRevit_UI_IT",
        "PyRevit_UI_IT.extension",
    ]
    for base in carpetas_base:
        for nombre in nombres:
            candidato = os.path.join(base, nombre)
            if os.path.isdir(candidato):
                log("EXT_ROOT via fallback: {}".format(candidato))
                return candidato

    # Último recurso: dos niveles sobre la carpeta del tab
    try:
        ruta = os.path.abspath(__file__)
        for _ in range(5):
            ruta = os.path.dirname(ruta)
        log("EXT_ROOT ultimo recurso: {}".format(ruta))
        return ruta
    except Exception:
        return appdata

EXT_ROOT = _get_ext_root()

# ══════════════════════════════════════════════════════════════
#  3. RUTAS BASE (calculadas desde EXT_ROOT)
# ══════════════════════════════════════════════════════════════
DATA_MASTER      = os.path.join(EXT_ROOT, "data", "master")
DATA_OUTPUT      = os.path.join(EXT_ROOT, "data", "output")
SCRIPTS_CPYTHON  = os.path.join(EXT_ROOT, "scripts_cpython")

# Actualizar LOG_PATH ahora que conocemos DATA_OUTPUT
LOG_PATH = os.path.join(DATA_OUTPUT, "planilla_vs_modelo_log.txt")

# Archivos JSON de configuración
CONFIG_PROYECTO       = os.path.join(DATA_MASTER, "config_proyecto_activo.json")
SCRIPT_JSON           = os.path.join(DATA_MASTER, "script.json")
PLANILLAS_HEADERS_JSON = os.path.join(DATA_MASTER, "planillas_headers_order.json")

# Archivo de salida del modelo
MODELO_JSON = os.path.join(DATA_OUTPUT, "modelo_codint_por_cm.json")

# Scripts CPython
LEER_XLSM       = os.path.join(SCRIPTS_CPYTHON, "leer_xlsm_codigos.py")
UI_COMPARACION  = os.path.join(SCRIPTS_CPYTHON, "ui_comparacion.py")
FORMATEAR_XLSX  = os.path.join(SCRIPTS_CPYTHON, "formatear_tablas_excel.py")

log("==== Inicio: Planilla vs Modelo ====")
log("EXT_ROOT     : {}".format(EXT_ROOT))
log("DATA_MASTER  : {}".format(DATA_MASTER))
log("SCRIPT_JSON  : {}".format(SCRIPT_JSON))

# ══════════════════════════════════════════════════════════════
#  4. PYTHON 3 EXE
# ══════════════════════════════════════════════════════════════
CREATE_NO_WINDOW = 0x08000000

def get_python3_exe():
    """Lee la ruta de Python3 desde config_proyecto_activo.json, con fallback."""
    try:
        with open(CONFIG_PROYECTO, "r", encoding="utf-8") as f:
            cfg = json.load(f)
        ruta = (cfg.get("python3_exe") or "").strip()
        if ruta and os.path.isfile(ruta):
            log("python3_exe desde config: {}".format(ruta))
            return ruta
    except Exception as e:
        log("get_python3_exe error leyendo config: {}".format(e))

    # Fallback: buscar python en paths conocidos
    candidatos = [
        r"C:\Users\Zbook HP\AppData\Local\Programs\Python\Python313\python.exe",
        r"C:\Users\Zbook HP\AppData\Local\Programs\Python\Python312\python.exe",
        r"C:\Python313\python.exe",
        r"C:\Python312\python.exe",
    ]
    for c in candidatos:
        if os.path.isfile(c):
            log("python3_exe via fallback: {}".format(c))
            return c

    log("python3_exe NO encontrado")
    return "python"

PYTHON3_EXE = get_python3_exe()

# ══════════════════════════════════════════════════════════════
#  5. HELPERS JSON
# ══════════════════════════════════════════════════════════════
def cargar_json(ruta, default=None):
    if default is None:
        default = {}
    try:
        if not os.path.exists(ruta):
            log("cargar_json: no existe -> {}".format(ruta))
            return default
        with open(ruta, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        log("cargar_json ERROR {}: {}".format(ruta, e))
        return default

# ══════════════════════════════════════════════════════════════
#  6. REPO ACTIVO (BD)
# ══════════════════════════════════════════════════════════════
def get_repo_activo():
    cfg = cargar_json(CONFIG_PROYECTO)
    ruta = (cfg.get("ruta_repositorio_activo") or "").strip()
    if not ruta or not os.path.exists(ruta):
        log("get_repo_activo: sin repo o no existe -> {}".format(ruta))
        return {}
    repo = cargar_json(ruta, {})
    log("get_repo_activo: {} registros desde {}".format(len(repo), ruta))
    return repo

# ══════════════════════════════════════════════════════════════
#  7. DOCS REVIT (host + links)
# ══════════════════════════════════════════════════════════════
def get_all_docs():
    docs = []
    if doc is None:
        return docs
    docs.append((doc, doc.PathName or ""))
    try:
        seen = set()
        for li in FilteredElementCollector(doc).OfClass(RevitLinkInstance):
            try:
                ld = li.GetLinkDocument()
                if ld is None:
                    continue
                p = ld.PathName or ""
                if p in seen:
                    continue
                seen.add(p)
                docs.append((ld, p))
            except Exception as e:
                log("get_all_docs link error: {}".format(e))
    except Exception as e:
        log("get_all_docs error: {}".format(e))
    log("get_all_docs: {} documentos".format(len(docs)))
    return docs

# ══════════════════════════════════════════════════════════════
#  8. GENERAR modelo_codint_por_cm.json
# ══════════════════════════════════════════════════════════════
def generar_modelo_json():
    if doc is None:
        forms.alert("No hay documento Revit activo.", title="Error")
        return False

    script_data = cargar_json(SCRIPT_JSON)
    codigos_planillas = script_data.get("codigos_planillas", {})
    if not codigos_planillas:
        forms.alert("script.json no tiene 'codigos_planillas'.", title="Error")
        return False

    # Construir mapa CM → nombre_planilla
    cm_to_planilla = {}
    for k, v in codigos_planillas.items():
        if isinstance(v, str) and v.startswith("CM") and len(v) == 4:
            cm_to_planilla[v] = k
        elif isinstance(k, str) and k.startswith("CM") and len(k) == 4:
            cm_to_planilla[k] = v

    planillas_headers = cargar_json(PLANILLAS_HEADERS_JSON)
    repo_activo       = get_repo_activo()

    # Índices BD para búsqueda rápida
    idx_eid = {}
    idx_cod = {}
    for v in repo_activo.values():
        cod  = (v.get("CodIntBIM") or "").strip()
        eid  = str(v.get("ElementId", "")).strip()
        arch = (v.get("Archivo") or "").strip()
        if cod:
            idx_cod.setdefault(cod, []).append(v)
        if eid and arch:
            idx_eid[(arch, eid)] = v

    def headers_para_cm(cm):
        nombre = cm_to_planilla.get(cm)
        if nombre:
            h = planillas_headers.get("{}::{}".format(nombre, cm))
            if h:
                return list(h)
        h = planillas_headers.get(cm)
        return list(h) if h else []

    datos_por_cm = {}
    usados_eid   = set()
    usados_cod   = set()
    total_elems  = 0

    for d, ruta_arch in get_all_docs():
        try:
            col = FilteredElementCollector(d).WhereElementIsNotElementType()
        except Exception as e:
            log("Error collector {}: {}".format(ruta_arch, e))
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
                cm      = codint[:4]
                eid_str = str(el.Id.IntegerValue)

                fila_bd = idx_eid.get((ruta_arch, eid_str))
                if not fila_bd:
                    lst     = idx_cod.get(codint, [])
                    fila_bd = lst[0] if lst else None

                fila = dict(fila_bd) if fila_bd else {}
                if fila_bd:
                    usados_eid.add((ruta_arch, eid_str))
                    usados_cod.add(codint)

                fila["CodIntBIM"] = codint
                fila["ElementId"] = eid_str
                fila["Archivo"]   = ruta_arch
                try:
                    fila["Categoria"] = el.Category.Name if el.Category else ""
                except Exception:
                    fila["Categoria"] = ""

                for h in headers_para_cm(cm):
                    if h == "CodIntBIM":
                        continue
                    if fila.get(h) not in (None, ""):
                        continue
                    try:
                        pel      = el.LookupParameter(h)
                        fila[h]  = (pel.AsString() or pel.AsValueString() or "") if pel else ""
                    except Exception:
                        fila[h] = ""

                datos_por_cm.setdefault(cm, []).append(fila)
            except Exception as e:
                log("Error elemento: {}".format(e))

    # Registros solo en BD (no encontrados en modelo)
    for v in repo_activo.values():
        cod = (v.get("CodIntBIM") or "").strip()
        if not cod or len(cod) < 4:
            continue
        cm   = cod[:4]
        eid  = str(v.get("ElementId", "")).strip()
        arch = (v.get("Archivo") or "").strip()
        if (arch, eid) in usados_eid or cod in usados_cod:
            continue
        fila = dict(v)
        fila.setdefault("Categoria", "")
        datos_por_cm.setdefault(cm, []).append(fila)

    # Guardar JSON
    try:
        if not os.path.exists(DATA_OUTPUT):
            os.makedirs(DATA_OUTPUT)
        with open(MODELO_JSON, "w", encoding="utf-8") as f:
            json.dump(datos_por_cm, f, ensure_ascii=False, indent=2)
        log("modelo_codint_por_cm.json OK. Elems={}, CMs={}".format(
            total_elems, list(datos_por_cm.keys())))
        return True
    except Exception as e:
        forms.alert("Error guardando modelo JSON:\n{}".format(e), title="Error")
        log("Error guardando modelo JSON: {}".format(e))
        return False

# ══════════════════════════════════════════════════════════════
#  9. SELECCIONAR XLSM
# ══════════════════════════════════════════════════════════════
def seleccionar_xlsm():
    try:
        return forms.pick_file(file_ext="xlsm", multi_file=False,
                               title="Selecciona la planilla .xlsm")
    except Exception as e:
        forms.alert("Error seleccionando .xlsm:\n{}".format(e), title="Error")
        return None

# ══════════════════════════════════════════════════════════════
#  10. LEER XLSM → CSV
# ══════════════════════════════════════════════════════════════
def llamar_leer_xlsm(ruta_xlsm):
    try:
        salida = subprocess.check_output(
            [PYTHON3_EXE, LEER_XLSM, ruta_xlsm, DATA_OUTPUT],
            stderr=subprocess.STDOUT,
            universal_newlines=True,
            creationflags=CREATE_NO_WINDOW
        )
        log("leer_xlsm salida:\n{}".format(salida))
        lineas = [l.strip() for l in salida.splitlines() if l.strip()]
        if not lineas:
            forms.alert("leer_xlsm_codigos.py no retornó ruta CSV.", title="Error")
            return None
        csv_path = lineas[-1]
        if not os.path.exists(csv_path):
            forms.alert("CSV no generado:\n{}".format(csv_path), title="Error")
            return None
        return csv_path
    except subprocess.CalledProcessError as e:
        forms.alert("Error leyendo .xlsm:\n{}".format(e.output), title="Error")
        log("llamar_leer_xlsm CalledProcessError: {}".format(e.output))
        return None
    except Exception as e:
        forms.alert("Error inesperado leer_xlsm:\n{}".format(e), title="Error")
        log("llamar_leer_xlsm error: {}".format(e))
        return None

# ══════════════════════════════════════════════════════════════
#  11. LANZAR UI DE COMPARACIÓN
# ══════════════════════════════════════════════════════════════
def llamar_ui_comparacion(ruta_xlsm, csv_codigos):
    carpeta_destino = forms.pick_folder(
        title="Selecciona carpeta donde guardar el Excel exportado"
    )
    if not carpeta_destino:
        forms.alert("No se seleccionó carpeta. Proceso cancelado.", title="Cancelado")
        log("llamar_ui_comparacion: carpeta cancelada")
        return

    stamp             = datetime.now().strftime("%Y%m%d_%H%M")
    ruta_xlsx_salida  = os.path.join(
        carpeta_destino, "planilla-modelo_{}.xlsx".format(stamp)
    )
    log("llamar_ui_comparacion: salida={}".format(ruta_xlsx_salida))

    try:
        subprocess.check_call(
            [
                PYTHON3_EXE,
                UI_COMPARACION,
                SCRIPT_JSON,        # argv[1]
                csv_codigos,        # argv[2]
                DATA_OUTPUT,        # argv[3]
                FORMATEAR_XLSX,     # argv[4]
                ruta_xlsx_salida,   # argv[5]
                PYTHON3_EXE,        # argv[6]
                MODELO_JSON,        # argv[7]
            ],
            creationflags=CREATE_NO_WINDOW
        )
        forms.alert(
            "Archivo exportado exitosamente:\n{}".format(ruta_xlsx_salida),
            title="Éxito"
        )
        log("Excel generado OK: {}".format(ruta_xlsx_salida))
    except subprocess.CalledProcessError as e:
        forms.alert("Error en UI comparación:\n{}".format(e), title="Error")
        log("llamar_ui_comparacion CalledProcessError: {}".format(e))
    except Exception as e:
        forms.alert("Error inesperado UI:\n{}".format(e), title="Error")
        log("llamar_ui_comparacion error: {}".format(e))

# ══════════════════════════════════════════════════════════════
#  12. MAIN
# ══════════════════════════════════════════════════════════════
def main():
    # Validaciones previas
    if doc is None:
        forms.alert("No hay documento Revit activo.", title="Error")
        return

    # Verificar archivos críticos
    faltantes = []
    for ruta, nombre in [
        (SCRIPT_JSON,    "data/master/script.json"),
        (LEER_XLSM,      "scripts_cpython/leer_xlsm_codigos.py"),
        (UI_COMPARACION, "scripts_cpython/ui_comparacion.py"),
    ]:
        if not os.path.exists(ruta):
            faltantes.append("• {} \n  ({})".format(nombre, ruta))

    if faltantes:
        forms.alert(
            "No se encontraron los siguientes archivos:\n\n{}\n\n"
            "EXT_ROOT resuelto:\n{}".format("\n".join(faltantes), EXT_ROOT),
            title="Error — Archivos faltantes"
        )
        log("Archivos faltantes:\n{}".format("\n".join(faltantes)))
        return

    if not os.path.isfile(PYTHON3_EXE):
        forms.alert(
            "Python 3 no encontrado en:\n{}\n\n"
            "Agrega la clave 'python3_exe' en:\n{}".format(
                PYTHON3_EXE, CONFIG_PROYECTO
            ),
            title="Error — Python 3 no encontrado"
        )
        return

    # Paso 1: seleccionar planilla
    ruta_xlsm = seleccionar_xlsm()
    if not ruta_xlsm:
        log("Selección xlsm cancelada")
        return

    # Paso 2: generar JSON del modelo
    with forms.ProgressBar(title="Analizando modelo Revit...", cancellable=False) as pb:
        pb.update_progress(10, 100)
        ok = generar_modelo_json()
        pb.update_progress(80, 100)

    if not ok:
        log("generar_modelo_json devolvió False")
        return

    # Paso 3: extraer CSV de la planilla
    csv_codigos = llamar_leer_xlsm(ruta_xlsm)
    if not csv_codigos:
        log("leer_xlsm devolvió None")
        return

    # Paso 4: lanzar UI de comparación
    llamar_ui_comparacion(ruta_xlsm, csv_codigos)
    log("==== Fin: Planilla vs Modelo ====")


if __name__ == "__main__":
    main()