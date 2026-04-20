# -*- coding: utf-8 -*-
__title__   = "planilla vs modelo"
__doc__     = """Version = 1.1
Date    = 20.04.2026
_______________________________________________________________
Description:

Comparación de Planilla con datos y elementos de modelo en función
de la información existente y la faltante. Si no existe
modelo_codint_por_cm.json, lo genera automáticamente recorriendo
el modelo activo y sus links, tomando los parámetros definidos en
planillas_headers_order.json por cada CMxx.

________________________________________________________________
Cambios v1.1:
- [CAMBIO 1] Tema dark activado via pyRevit user_config
- [CAMBIO 2] Flujo corregido: usa MODELO_JSON propio, no el de revisión elementos
- [CAMBIO 3] Orden reordenado: selección xlsm primero → análisis modelo → UI (sin doble ejecución)
- [CAMBIO 4] Selector de carpeta destino antes de exportar
________________________________________________________________
Author: Argenis Angel"""

# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝
#==================================================
import os
import sys
import subprocess
from datetime import datetime

import clr
from pyrevit import forms

# ─── CAMBIO 1: Activar tema dark en pyRevit ────────────────────────────────
try:
    from pyrevit.userconfig import user_config
    user_config.core.darkmode = True
    user_config.save_changes()
except Exception:
    pass
# ───────────────────────────────────────────────────────────────────────────

# Revit API
clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    RevitLinkInstance,
)

# ╦  ╦╔═╗╦═╗╦╔═╗╔╗ ╦  ╔═╗╔═╗
# ╚╗╔╝╠═╣╠╦╝║╠═╣╠╩╗║  ║╣ ╚═╗
#  ╚╝ ╩ ╩╩╚═╩╩ ╩╚═╝╩═╝╚═╝╚═╝
#==================================================

try:
    doc = __revit__.ActiveUIDocument.Document
except Exception:
    doc = None

# ╔╦╗╔═╗╦╔╗╔
# ║║║╠═╣║║║║
# ╩ ╩╩ ╩╩╝╚╝
#==================================================

# ---------------------------------------------------------
# Rutas base
# ---------------------------------------------------------

DATA_DIR_EXT = os.path.join(
    os.path.expanduser("~"),
    r"AppData\Roaming\MyPyRevitExtention\PyRevitIT.extension\data"
)

LOG_PATH = os.path.join(DATA_DIR_EXT, "planilla_vs_modelo_log.txt")

CONFIG_PROYECTO_ACTIVO = os.path.join(
    DATA_DIR_EXT,
    "config_proyecto_activo.json"
)


def log(msg):
    """Escribe mensajes en el log con timestamp."""
    try:
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        linea = u"[{}] {}\n".format(ts, msg)
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(linea)
    except Exception:
        pass


try:
    if not os.path.exists(DATA_DIR_EXT):
        forms.alert(
            "No se encontró la carpeta data de la extensión:\n{}".format(DATA_DIR_EXT),
            title="Error"
        )
        log("DATA_DIR_EXT no existe: {}".format(DATA_DIR_EXT))
        raise SystemExit

    APPDATA_DIR = os.getenv('APPDATA') or os.path.expanduser('~')
    DATA_DIR = os.path.join(APPDATA_DIR, 'PyRevitIT', 'data', 'comparacion')
    if not os.path.exists(DATA_DIR):
        os.makedirs(DATA_DIR)
        log("Se creó DATA_DIR: {}".format(DATA_DIR))

    PYTHON3_EXE = r"C:\Users\Zbook HP\AppData\Local\Programs\Python\Python313\python.exe"

    LEER_XLSM      = os.path.join(DATA_DIR_EXT, 'leer_xlsm_codigos.py')
    # ─── CAMBIO 2: formateador específico para planilla vs modelo ──────────
    FORMATEAR_XLSX = os.path.join(DATA_DIR_EXT, 'formatear_tablas_planilla_vs_modelo.py')
    # Fallback al formateador genérico si el específico no existe todavía
    if not os.path.exists(FORMATEAR_XLSX):
        FORMATEAR_XLSX = os.path.join(DATA_DIR_EXT, 'formatear_tablas_excel.py')
    # ───────────────────────────────────────────────────────────────────────
    UI_COMPARACION  = os.path.join(DATA_DIR_EXT, 'ui_comparacion.py')
    SCRIPT_JSON_PATH = os.path.join(DATA_DIR_EXT, 'script.json')

    MODELO_JSON          = os.path.join(DATA_DIR_EXT, 'modelo_codint_por_cm.json')
    HEADERS_JSON         = os.path.join(DATA_DIR_EXT, 'headers_por_tabla.json')
    PLANILLAS_HEADERS_JSON = os.path.join(DATA_DIR_EXT, 'planillas_headers_order.json')

    CREATE_NO_WINDOW = 0x08000000
except Exception as e:
    log("Error inicializando rutas base: {}".format(e))
    forms.alert("Error inicializando rutas base:\n{}".format(e), title="Error")
    raise SystemExit


# ----------------- Utilidades JSON / BD -----------------


def cargar_json(ruta, default):
    import json
    try:
        if not os.path.exists(ruta):
            log("cargar_json: archivo no existe -> {}".format(ruta))
            return default
        with open(ruta, "r", encoding="utf-8") as f:
            data = json.load(f)
        log("cargar_json: cargado OK -> {}".format(ruta))
        return data
    except Exception as e:
        log("cargar_json: error leyendo {} -> {}".format(ruta, e))
        return default


def get_repo_activo_path():
    cfg = cargar_json(CONFIG_PROYECTO_ACTIVO, {})
    ruta = (cfg.get("ruta_repositorio_activo") or "").strip()
    log("get_repo_activo_path: {}".format(ruta or "<no definido>"))
    return ruta


def cargar_repo_activo():
    ruta = get_repo_activo_path()
    if not ruta or not os.path.exists(ruta):
        log("cargar_repo_activo: no hay repo activo o no existe -> {}".format(ruta))
        return {}
    repo = cargar_json(ruta, {})
    log("cargar_repo_activo: {} registros".format(len(repo)))
    return repo


# ----------------- Utilidades Revit -----------------


def get_all_docs_with_links():
    """Devuelve lista de (documento, ruta_archivo) para host + links."""
    docs = []
    try:
        if doc is None:
            log("get_all_docs_with_links: doc es None")
            return []

        main_doc = doc
        docs.append((main_doc, main_doc.PathName or ""))
        log("get_all_docs_with_links: host -> {}".format(main_doc.PathName or ""))

        col_links = FilteredElementCollector(main_doc).OfClass(RevitLinkInstance)
        seen = set()
        for li in col_links:
            try:
                link_doc = li.GetLinkDocument()
                if link_doc is None:
                    continue
                path = link_doc.PathName or ""
                if path in seen:
                    continue
                seen.add(path)
                docs.append((link_doc, path))
                log("get_all_docs_with_links: link -> {}".format(path))
            except Exception as e:
                log("get_all_docs_with_links: error con RevitLinkInstance -> {}".format(e))
                continue
    except Exception as e:
        forms.alert(
            "Error obteniendo documentos linkeados:\n{}".format(e),
            title="Error links"
        )
        log("get_all_docs_with_links: excepción general -> {}".format(e))
    return docs


def cargar_planillas_headers():
    data = cargar_json(PLANILLAS_HEADERS_JSON, {})
    if not data:
        log("cargar_planillas_headers: archivo vacío o no encontrado -> {}".format(PLANILLAS_HEADERS_JSON))
    else:
        log("cargar_planillas_headers: {} claves".format(len(data)))
    return data


def generar_modelo_json_desde_revit(script_json_path, modelo_json_path):
    """
    Recorre doc + links, toma elementos con CodIntBIM y combina:
      - Datos de BD (repositorio activo)
      - Datos del modelo (LookupParameter)
    generando modelo_codint_por_cm.json agrupado por CMxx.
    """
    import json

    if doc is None:
        forms.alert("No hay documento activo.", title="Error")
        log("generar_modelo_json_desde_revit: doc es None")
        return False

    script_data = cargar_json(script_json_path, {})
    codigos_planillas = script_data.get("codigos_planillas", {})
    if not codigos_planillas:
        msg = "En script.json no se encontró 'codigos_planillas'."
        forms.alert(msg, title="Error script.json")
        log("generar_modelo_json_desde_revit: {}".format(msg))
        return False

    cm_to_planilla = {}
    for k, v in codigos_planillas.items():
        try:
            if isinstance(v, str) and len(v) == 4 and v.startswith("CM"):
                cm_to_planilla[v] = k
            elif isinstance(k, str) and len(k) == 4 and k.startswith("CM"):
                cm_to_planilla[k] = v
        except Exception as e:
            log("generar_modelo_json_desde_revit: error armando cm_to_planilla -> {}".format(e))

    log("generar_modelo_json_desde_revit: cm_to_planilla -> {}".format(cm_to_planilla))

    planillas_headers = cargar_planillas_headers()
    repo_activo = cargar_repo_activo()

    indice_bd_eid = {}
    indice_bd_cod = {}
    for k, v in repo_activo.items():
        cod = (v.get("CodIntBIM") or "").strip()
        eid = str(v.get("ElementId", "")).strip()
        archivo = (v.get("Archivo") or "").strip()
        if cod:
            indice_bd_cod.setdefault(cod, []).append(v)
        if eid and archivo:
            indice_bd_eid[(archivo, eid)] = v
    log("generar_modelo_json_desde_revit: indice_bd_cod={}, indice_bd_eid={}".format(
        len(indice_bd_cod), len(indice_bd_eid)
    ))

    def headers_para_cm(cm_codigo):
        nombre_planilla = cm_to_planilla.get(cm_codigo)
        headers = None
        if nombre_planilla:
            clave = "{}::{}".format(nombre_planilla, cm_codigo)
            headers = planillas_headers.get(clave)
            if headers:
                log("headers_para_cm: {} -> {} (clave {})".format(cm_codigo, len(headers), clave))
        if not headers:
            headers = planillas_headers.get(cm_codigo)
            if headers:
                log("headers_para_cm: {} -> {} (clave {})".format(cm_codigo, len(headers), cm_codigo))
        if headers:
            return list(headers)
        log("headers_para_cm: {} sin headers configurados".format(cm_codigo))
        return []

    datos_por_cm = {}
    total_docs = 0
    total_elems_codint = 0

    usados_bd_codint = set()
    usados_bd_eid = set()

    try:
        docs_links = get_all_docs_with_links()
        total_docs = len(docs_links)
        log("generar_modelo_json_desde_revit: documentos a recorrer -> {}".format(total_docs))

        for d, ruta_archivo in docs_links:
            try:
                ruta_log = ruta_archivo or "<sin ruta>"
                log("Recorriendo doc: {}".format(ruta_log))
                col = FilteredElementCollector(d).WhereElementIsNotElementType()
            except Exception as e:
                log("Error creando FilteredElementCollector para {} -> {}".format(ruta_archivo, e))
                continue

            for el in col:
                try:
                    p = el.LookupParameter("CodIntBIM")
                    if not p:
                        continue
                    codint = p.AsString() or ""
                    codint = codint.strip()
                    if not codint or len(codint) < 4:
                        continue

                    total_elems_codint += 1
                    cm_codigo = codint[:4]

                    eid_str = str(el.Id.IntegerValue)
                    archivo = ruta_archivo or ""

                    fila_bd = indice_bd_eid.get((archivo, eid_str))
                    if not fila_bd:
                        lst = indice_bd_cod.get(codint) or []
                        fila_bd = lst[0] if lst else None

                    if fila_bd:
                        usados_bd_codint.add(codint)
                        usados_bd_eid.add((archivo, eid_str))
                        fila = dict(fila_bd)
                    else:
                        fila = {}

                    fila["CodIntBIM"] = codint
                    fila["ElementId"] = eid_str
                    fila["Archivo"] = archivo

                    try:
                        categoria = el.Category.Name if el.Category else ""
                    except Exception as ecat:
                        categoria = ""
                        log("Error leyendo Category para {} -> {}".format(codint, ecat))
                    fila["Categoria"] = categoria

                    headers_cm = headers_para_cm(cm_codigo)
                    for h in headers_cm:
                        if h == "CodIntBIM":
                            continue
                        try:
                            p_el = el.LookupParameter(h)
                            if p_el:
                                val_modelo = p_el.AsString() or p_el.AsValueString() or ""
                            else:
                                val_modelo = ""
                        except Exception as ep:
                            val_modelo = ""
                            log("Error leyendo parámetro '{}' para {} -> {}".format(h, codint, ep))

                        val_bd = fila.get(h)
                        if val_bd not in (None, ""):
                            fila[h] = val_bd
                        else:
                            fila[h] = val_modelo

                    datos_por_cm.setdefault(cm_codigo, []).append(fila)

                except Exception as e_el:
                    log("Error procesando elemento con CodIntBIM en doc {} -> {}".format(ruta_archivo, e_el))
                    continue

    except Exception as e:
        forms.alert(
            "Error recorriendo elementos del modelo y links:\n{}".format(e),
            title="Error modelo"
        )
        log("generar_modelo_json_desde_revit: excepción general -> {}".format(e))
        return False

    for k, v in repo_activo.items():
        cod = (v.get("CodIntBIM") or "").strip()
        if not cod or len(cod) < 4:
            continue
        cm_codigo = cod[:4]
        eid = str(v.get("ElementId", "")).strip()
        archivo = (v.get("Archivo") or "").strip()

        if (archivo, eid) in usados_bd_eid or cod in usados_bd_codint:
            continue

        fila = dict(v)
        fila["CodIntBIM"] = cod
        fila["ElementId"] = eid
        fila["Archivo"] = archivo

        if "Categoria" not in fila:
            fila["Categoria"] = ""

        datos_por_cm.setdefault(cm_codigo, []).append(fila)

    log("generar_modelo_json_desde_revit: total_docs={}, total_elems_codint={}".format(
        total_docs, total_elems_codint
    ))
    log("generar_modelo_json_desde_revit: CM encontrados -> {}".format(list(datos_por_cm.keys())))

    try:
        with open(modelo_json_path, "w", encoding="utf-8") as f:
            import json
            json.dump(datos_por_cm, f, ensure_ascii=False, indent=2)
        log("modelo_codint_por_cm.json guardado en {}".format(modelo_json_path))
    except Exception as e:
        forms.alert(
            "Error guardando modelo_codint_por_cm.json:\n{}".format(e),
            title="Error JSON modelo"
        )
        log("Error guardando modelo_codint_por_cm.json -> {}".format(e))
        return False

    # ─── CAMBIO 3: sin alert intermedia → flujo continuo ──────────────────
    log("modelo generado OK -> docs={}, elementos={}".format(total_docs, total_elems_codint))
    return True
    # ───────────────────────────────────────────────────────────────────────


# ----------------- Flujo Excel/Comparación -----------------


def seleccionar_xlsm():
    try:
        ruta = forms.pick_file(
            file_ext='xlsm',
            multi_file=False,
            title='Selecciona la planilla .xlsm'
        )
        log("seleccionar_xlsm: ruta seleccionada -> {}".format(ruta or "<cancelado>"))
        return ruta
    except Exception as e:
        log("seleccionar_xlsm: error -> {}".format(e))
        forms.alert("Error al seleccionar xlsm:\n{}".format(e), title="Error")
        return None


def llamar_leer_xlsm(ruta_xlsm):
    """Ejecuta leer_xlsm_codigos.py y devuelve ruta del CODIGO.csv."""
    try:
        log("llamar_leer_xlsm: inicio con ruta_xlsm={}".format(ruta_xlsm))
        salida = subprocess.check_output(
            [PYTHON3_EXE, LEER_XLSM, ruta_xlsm, DATA_DIR],
            stderr=subprocess.STDOUT,
            universal_newlines=True,
            creationflags=CREATE_NO_WINDOW
        )
        log("llamar_leer_xlsm: salida:\n{}".format(salida))
        lineas = [l.strip() for l in salida.splitlines() if l.strip()]
        if not lineas:
            msg = "leer_xlsm_codigos.py no devolvió ninguna ruta."
            forms.alert("{}\n\nSalida:\n{}".format(msg, salida), title="Error")
            log("llamar_leer_xlsm: {}".format(msg))
            return None
        csv_path = lineas[-1]
        if not os.path.exists(csv_path):
            msg = "No se generó el CSV de CODIGO."
            forms.alert("{}\n\nTexto devuelto:\n{}".format(msg, salida), title="Error")
            log("llamar_leer_xlsm: {} -> {}".format(msg, csv_path))
            return None
        log("llamar_leer_xlsm: csv generado -> {}".format(csv_path))
        return csv_path
    except subprocess.CalledProcessError as e:
        forms.alert("Error al leer .xlsm:\n{}".format(e.output), title="Error")
        log("llamar_leer_xlsm: CalledProcessError -> {}".format(e.output))
        return None
    except Exception as e:
        forms.alert("Error inesperado al leer .xlsm:\n{}".format(e), title="Error")
        log("llamar_leer_xlsm: excepción general -> {}".format(e))
        return None


def llamar_ui_y_formato(ruta_xlsm, csv_codigos):
    ahora = datetime.now()
    stamp = ahora.strftime("%Y%m%d_%H%M")

    # ─── CAMBIO 4: Selector de carpeta destino ────────────────────────────
    carpeta_destino = forms.pick_folder(
        title="Selecciona la carpeta donde guardar el archivo Excel exportado"
    )
    if not carpeta_destino:
        log("llamar_ui_y_formato: selección de carpeta cancelada")
        forms.alert("No se seleccionó carpeta de destino. Operación cancelada.", title="Cancelado")
        return
    # ───────────────────────────────────────────────────────────────────────

    ruta_xlsx_salida = os.path.join(
        carpeta_destino,
        "planilla-modelo_{}.xlsx".format(stamp)
    )

    log("llamar_ui_y_formato: ruta_xlsm={}, csv_codigos={}, salida={}".format(
        ruta_xlsm, csv_codigos, ruta_xlsx_salida
    ))

    try:
        subprocess.check_call(
            [
                PYTHON3_EXE,
                UI_COMPARACION,
                SCRIPT_JSON_PATH,
                csv_codigos,
                DATA_DIR,
                FORMATEAR_XLSX,      # ← formateador específico planilla vs modelo
                ruta_xlsx_salida,
                PYTHON3_EXE,         # 6º argumento (python_exe)
                MODELO_JSON,         # 7º argumento (modelo_json propio)
                HEADERS_JSON         # 8º argumento
            ],
            stderr=subprocess.STDOUT,
            creationflags=CREATE_NO_WINDOW
        )
        forms.alert(
            "Archivo generado:\n{}".format(ruta_xlsx_salida),
            title="Éxito"
        )
        log("llamar_ui_y_formato: Excel generado OK -> {}".format(ruta_xlsx_salida))
    except subprocess.CalledProcessError as e:
        forms.alert("Error en la UI / formateo:\n{}".format(e), title="Error")
        log("llamar_ui_y_formato: CalledProcessError -> {}".format(e))
    except Exception as e:
        forms.alert("Error inesperado en UI / formateo:\n{}".format(e), title="Error")
        log("llamar_ui_y_formato: excepción general -> {}".format(e))


def main():
    log("==== Inicio Planilla vs Modelo v1.1 ====")

    if doc is None:
        forms.alert("No hay documento activo.", title="Error")
        log("main: doc es None")
        return

    if not os.path.exists(SCRIPT_JSON_PATH):
        forms.alert("No se encontró script.json:\n{}".format(SCRIPT_JSON_PATH), title="Error")
        log("main: script.json no encontrado")
        return

    if not os.path.exists(LEER_XLSM) or not os.path.exists(UI_COMPARACION):
        forms.alert(
            "Faltan scripts CPython en la carpeta data.\nSe esperaban:\n{}\n{}".format(
                LEER_XLSM, UI_COMPARACION
            ),
            title="Error"
        )
        log("main: scripts CPython no encontrados")
        return

    # ─── CAMBIO 3: Seleccionar planilla PRIMERO, luego analizar modelo ──────
    # Así la ventana de UI aparece en flujo continuo sin tener que
    # presionar el botón por segunda vez.
    ruta_xlsm = seleccionar_xlsm()
    if not ruta_xlsm:
        log("main: selección xlsm cancelada")
        return

    # Analizar modelo (sin alert intermedia de éxito)
    log("main: generando MODELO_JSON -> {}".format(MODELO_JSON))
    ok = generar_modelo_json_desde_revit(SCRIPT_JSON_PATH, MODELO_JSON)
    if not ok:
        log("main: generar_modelo_json_desde_revit devolvió False")
        return

    # Leer CSV desde la planilla seleccionada
    csv_codigos = llamar_leer_xlsm(ruta_xlsm)
    if not csv_codigos:
        log("main: llamar_leer_xlsm devolvió None")
        return

    # Abrir UI, seleccionar carpeta y exportar → todo en una sola ejecución
    llamar_ui_y_formato(ruta_xlsm, csv_codigos)
    log("==== Fin Planilla vs Modelo v1.1 ====")
    # ───────────────────────────────────────────────────────────────────────


if __name__ == '__main__':
    main()
