# -*- coding: utf-8 -*-
"""
ui_comparacion.py — CPython 3
Ventana Tkinter (dark mode) para comparar planilla XLSM vs modelo Revit.

Invocado desde script.py con 7 argumentos posicionales:
  sys.argv[1]  = script_json_path        → data/master/script.json
  sys.argv[2]  = csv_codigos_path        → data/output/CODIGO.csv
  sys.argv[3]  = data_output_dir         → data/output/
  sys.argv[4]  = formatear_xlsx_path     → scripts_cpython/formatear_tablas_excel.py
  sys.argv[5]  = ruta_xlsx_salida        → carpeta destino/planilla-modelo_XXXXXX.xlsx
  sys.argv[6]  = python3_exe             → ruta al python.exe
  sys.argv[7]  = modelo_json_path        → data/output/modelo_codint_por_cm.json
"""

import sys
import os
import csv
import json
import subprocess
import traceback
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime

# ══════════════════════════════════════════════════════════════
#  ARGUMENTOS (primero para poder usar rutas en log)
# ══════════════════════════════════════════════════════════════
if len(sys.argv) < 8:
    print(
        "Uso: ui_comparacion.py <script_json> <csv_codigos> <data_output_dir> "
        "<formatear_script> <ruta_xlsx_salida> <python_exe> <modelo_json>",
        file=sys.stderr
    )
    sys.exit(1)

SCRIPT_JSON_PATH   = sys.argv[1].strip('"').strip("'")
CSV_CODIGOS_PATH   = sys.argv[2].strip('"').strip("'")
DATA_OUTPUT_DIR    = sys.argv[3].strip('"').strip("'")
FORMATEAR_SCRIPT   = sys.argv[4].strip('"').strip("'")
RUTA_XLSX_SALIDA   = sys.argv[5].strip('"').strip("'")
PYTHON_EXE         = sys.argv[6].strip('"').strip("'")
MODELO_JSON_PATH   = sys.argv[7].strip('"').strip("'")

# Derivar DATA_MASTER desde DATA_OUTPUT_DIR (sube un nivel → data/ → master/)
DATA_MASTER_DIR    = os.path.join(os.path.dirname(DATA_OUTPUT_DIR), "master")

# Archivos de configuración derivados (sin hardcodear)
CONFIG_PROYECTO    = os.path.join(DATA_MASTER_DIR, "config_proyecto_activo.json")
PLANILLAS_HEADERS  = os.path.join(DATA_MASTER_DIR, "planillas_headers_order.json")
UI_LOG_PATH        = os.path.join(DATA_OUTPUT_DIR, "ui_comparacion_log.txt")

CREATE_NO_WINDOW   = 0x08000000
HEADER_VINCULO     = u"Vínculo RVT: Nombre de archivo"

# ══════════════════════════════════════════════════════════════
#  LOG
# ══════════════════════════════════════════════════════════════
def log(msg):
    try:
        if not os.path.exists(DATA_OUTPUT_DIR):
            os.makedirs(DATA_OUTPUT_DIR)
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(UI_LOG_PATH, "a", encoding="utf-8") as f:
            f.write(u"[{}] {}\n".format(ts, msg))
    except Exception:
        pass

def log_exc(contexto):
    try:
        log("{} -> {}".format(contexto, traceback.format_exc()))
    except Exception:
        pass

# ══════════════════════════════════════════════════════════════
#  COLORES DARK MODE (Tkinter)
# ══════════════════════════════════════════════════════════════
BG_DARK    = "#171614"
SURFACE    = "#1c1b19"
SURFACE2   = "#2d2c2a"
BORDER_CLR = "#393836"
TEXT_PRI   = "#cdccca"
TEXT_MUTED = "#797876"
ACCENT     = "#4f98a3"
SUCCESS    = "#6daa45"
ERROR_CLR  = "#dd6974"
WARNING    = "#fdab43"

COLOR_ESTADOS = {
    "ok"           : "#3a4435",   # verde oscuro
    "falta_modelo" : "#574848",   # rojo oscuro
    "falta_excel"  : "#564b3e",   # amarillo oscuro
    "difiere"      : "#56493e",   # naranja oscuro
    "no_existe"    : "#313b3b",   # gris azulado
}
COLOR_ESTADOS_TEXT = {
    "ok"           : "#6daa45",
    "falta_modelo" : "#dd6974",
    "falta_excel"  : "#fdab43",
    "difiere"      : "#fdab43",
    "no_existe"    : "#797876",
}

# ══════════════════════════════════════════════════════════════
#  HELPERS JSON / CSV
# ══════════════════════════════════════════════════════════════
def cargar_json(ruta, default=None):
    if default is None:
        default = {}
    try:
        if not os.path.exists(ruta):
            log("cargar_json: no existe -> {}".format(ruta))
            return default
        with open(ruta, "r", encoding="utf-8") as f:
            data = json.load(f)
        log("cargar_json OK: {}".format(ruta))
        return data
    except Exception:
        log_exc("cargar_json {}".format(ruta))
        return default

def leer_script_json(path):
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        log("leer_script_json OK: {}".format(path))
        return data
    except Exception:
        log_exc("leer_script_json {}".format(path))
        raise

def leer_csv_codigos(csv_path):
    """Lee el CSV generado por leer_xlsm_codigos.py (delimitador ',')."""
    filas = []
    try:
        with open(csv_path, "r", encoding="utf-8-sig") as f:
            reader = csv.reader(f, delimiter=",")
            for row in reader:
                if not row:
                    continue
                if not any((c or "").strip() for c in row):
                    continue
                filas.append(row)
        log("leer_csv_codigos: {} filas desde {}".format(len(filas), csv_path))
        return filas
    except Exception:
        log_exc("leer_csv_codigos {}".format(csv_path))
        raise

def leer_modelo_por_cm(path_json):
    try:
        if not os.path.exists(path_json):
            log("leer_modelo_por_cm: no existe -> {}".format(path_json))
            return {}
        with open(path_json, "r", encoding="utf-8") as f:
            data = json.load(f)
        result = {cm: lista for cm, lista in data.items() if isinstance(lista, list)}
        log("leer_modelo_por_cm: CMs={} desde {}".format(list(result.keys()), path_json))
        return result
    except Exception:
        log_exc("leer_modelo_por_cm {}".format(path_json))
        raise

# ══════════════════════════════════════════════════════════════
#  REPOSITORIO / BD
# ══════════════════════════════════════════════════════════════
def cargar_repo_activo():
    cfg  = cargar_json(CONFIG_PROYECTO, {})
    ruta = (cfg.get("ruta_repositorio_activo") or "").strip()
    log("cargar_repo_activo: ruta={}".format(ruta))
    if not ruta or not os.path.exists(ruta):
        return {}
    repo = cargar_json(ruta, {})
    log("cargar_repo_activo: {} registros".format(len(repo)))
    return repo

def enriquecer_modelo_con_bd(datos_modelo_cm, repo):
    if not repo:
        return datos_modelo_cm
    try:
        indice_bd = {}
        for v in repo.values():
            eid = str(v.get("ElementId", "")).strip()
            cod = (v.get("CodIntBIM") or "").strip()
            if eid and cod:
                indice_bd[(eid, cod)] = v
        log("enriquecer_modelo_con_bd: {} claves en índice BD".format(len(indice_bd)))
        nuevo = {}
        for cm, lista in datos_modelo_cm.items():
            nueva_lista = []
            for fila in lista:
                cod = (fila.get("CodIntBIM") or "").strip()
                eid = str(fila.get("ElementId", "")).strip()
                bd_row = indice_bd.get((eid, cod))
                if bd_row:
                    combinado = dict(fila)
                    combinado.update(bd_row)
                    nueva_lista.append(combinado)
                else:
                    nueva_lista.append(fila)
            nuevo[cm] = nueva_lista
        return nuevo
    except Exception:
        log_exc("enriquecer_modelo_con_bd")
        raise

# ══════════════════════════════════════════════════════════════
#  HEADERS DE PLANILLA
# ══════════════════════════════════════════════════════════════
def _normalizar_headers(headers_raw):
    if not headers_raw:
        return []
    headers = [h for h in headers_raw if h != HEADER_VINCULO]
    if "CodIntBIM" in headers:
        headers = ["CodIntBIM"] + [h for h in headers if h != "CodIntBIM"]
    else:
        headers = ["CodIntBIM"] + headers
    return headers

def cargar_headers_planilla(nombre_planilla, codigo_cm):
    data = cargar_json(PLANILLAS_HEADERS, {})
    if not data:
        return None
    clave = "{}::{}".format(nombre_planilla, codigo_cm)
    headers = data.get(clave) or data.get(codigo_cm)
    if headers:
        h = _normalizar_headers(headers)
        log("cargar_headers_planilla: {} headers para {}".format(len(h), clave))
        return h
    log("cargar_headers_planilla: sin headers para {}".format(clave))
    return None

# ══════════════════════════════════════════════════════════════
#  CONSTRUCCIÓN DE DATOS
# ══════════════════════════════════════════════════════════════
def construir_excel_por_planilla(filas_csv, codigos_planillas):
    planilla_a_cm = {}
    for k, v in codigos_planillas.items():
        try:
            if isinstance(v, str) and len(v) == 4 and v.startswith("CM"):
                planilla_a_cm[k] = v
            elif isinstance(k, str) and len(k) == 4 and k.startswith("CM"):
                planilla_a_cm[v] = k
        except Exception:
            pass
    log("construir_excel_por_planilla: planilla_a_cm={}".format(planilla_a_cm))

    datos = {}
    # filas_csv[0] es la fila de encabezados del CSV
    encabezado_csv = filas_csv[0] if filas_csv else []

    for row in filas_csv[1:]:  # saltar encabezados
        if not row:
            continue
        primera_celda = (row[0] or "").strip()
        if not primera_celda or len(primera_celda) < 4:
            continue
        pref = primera_celda[:4]

        nombre_planilla = None
        for np, cm in planilla_a_cm.items():
            if cm == pref:
                nombre_planilla = np
                break
        if not nombre_planilla:
            continue

        headers = cargar_headers_planilla(nombre_planilla, pref)
        if not headers:
            headers = _normalizar_headers(
                [h for h in encabezado_csv if h] or
                ["Col{}".format(i) for i in range(1, len(row) + 1)]
            )

        bloque = datos.get(nombre_planilla)
        if not bloque:
            bloque = {"codigo_cm": pref, "headers": headers, "filas": []}
            datos[nombre_planilla] = bloque

        valores = []
        for idx_h in range(len(headers)):
            val = row[idx_h] if idx_h < len(row) else ""
            valores.append(val)

        codint = (valores[0] or "").strip()
        bloque["filas"].append({"CodIntBIM": codint, "valores": valores})

    log("construir_excel_por_planilla: {} planillas".format(len(datos)))
    return datos

def _norm_val(v):
    v = (v or "").strip()
    return "" if v == "-" else v

def _get_val(row_dict, header_name):
    if header_name == HEADER_VINCULO:
        return ""
    v = (row_dict.get(header_name) or "").strip()
    return "" if v == "-" else v

def _comparar_por_planilla(codigo_cm, headers, filas_excel, filas_modelo):
    mapa_excel  = {}
    for fila in filas_excel:
        c = (fila["CodIntBIM"] or "").strip()
        if c:
            mapa_excel.setdefault(c, []).append(fila)

    mapa_modelo = {}
    for row_m in filas_modelo:
        c = (row_m.get("CodIntBIM") or "").strip()
        if c:
            mapa_modelo.setdefault(c, []).append(row_m)

    filas_excel_base  = []
    filas_modelo_base = []
    filas_fusionadas  = []

    for cod in sorted(set(mapa_excel.keys()) | set(mapa_modelo.keys())):
        lista_excel  = mapa_excel.get(cod, [])
        lista_modelo = mapa_modelo.get(cod, [])
        fila_x0      = lista_excel[0]  if lista_excel  else None
        fila_m0      = lista_modelo[0] if lista_modelo else None

        # ── Fusionada ──
        valores_fusion = []
        estados_fusion = []
        for h in headers:
            vx_raw = ""
            vm_raw = ""
            if fila_x0:
                idx = headers.index(h)
                vals = fila_x0["valores"]
                vx_raw = vals[idx] if idx < len(vals) else ""
            if fila_m0:
                vm_raw = _get_val(fila_m0, h)
            vx = _norm_val(vx_raw)
            vm = _norm_val(vm_raw)
            if not vx and not vm:
                valores_fusion.append(""); estados_fusion.append("no_existe")
            elif vx == vm:
                valores_fusion.append(vx_raw); estados_fusion.append("ok")
            elif vx and not vm:
                valores_fusion.append(vx_raw); estados_fusion.append("falta_modelo")
            elif not vx and vm:
                valores_fusion.append(vm_raw); estados_fusion.append("falta_excel")
            else:
                valores_fusion.append(u"Planilla: {}\nModelo: {}".format(vx_raw, vm_raw))
                estados_fusion.append("difiere")
        filas_fusionadas.append({
            "CodIntBIM": cod, "valores": valores_fusion,
            "estado_por_celda": estados_fusion
        })

        # ── Desde Excel ──
        for fila_x in lista_excel:
            vals_x = fila_x["valores"]
            valores, estados = [], []
            fila_m = lista_modelo[0] if lista_modelo else None
            for idx, h in enumerate(headers):
                vx_raw = vals_x[idx] if idx < len(vals_x) else ""
                vx = _norm_val(vx_raw)
                vm_raw = _get_val(fila_m, h) if fila_m else ""
                vm = _norm_val(vm_raw)
                if not vx and not vm:
                    valores.append(""); estados.append("no_existe")
                elif vx == vm:
                    valores.append(vx_raw); estados.append("ok")
                elif vx and not vm:
                    valores.append(vx_raw); estados.append("falta_modelo")
                elif not vx and vm:
                    valores.append(u"Planilla: {}\nModelo: {}".format(vx_raw, vm_raw))
                    estados.append("falta_excel")
                else:
                    valores.append(u"Planilla: {}\nModelo: {}".format(vx_raw, vm_raw))
                    estados.append("difiere")
            filas_excel_base.append({
                "origen": "excel", "CodIntBIM": cod,
                "valores": valores, "estado_por_celda": estados
            })

        # ── Desde Modelo ──
        for fila_m in lista_modelo:
            valores, estados = [], []
            fila_x = lista_excel[0] if lista_excel else None
            for idx, h in enumerate(headers):
                vm_raw = _get_val(fila_m, h)
                vm = _norm_val(vm_raw)
                vx_raw = (fila_x["valores"][idx] if fila_x and idx < len(fila_x["valores"]) else "")
                vx = _norm_val(vx_raw)
                if not vx and not vm:
                    valores.append(""); estados.append("no_existe")
                elif vx == vm:
                    valores.append(vm_raw); estados.append("ok")
                elif vm and not vx:
                    valores.append(vm_raw); estados.append("falta_excel")
                elif not vm and vx:
                    valores.append(u"Modelo: {}\nPlanilla: {}".format(vm_raw, vx_raw))
                    estados.append("falta_modelo")
                else:
                    valores.append(u"Modelo: {}\nPlanilla: {}".format(vm_raw, vx_raw))
                    estados.append("difiere")
            filas_modelo_base.append({
                "origen": "modelo", "CodIntBIM": cod,
                "valores": valores, "estado_por_celda": estados
            })

    return filas_excel_base, filas_modelo_base, filas_fusionadas

def construir_tabla_comparativa(datos_excel_planilla, datos_modelo_cm, codigos_planillas):
    cm_to_planilla = {}
    for k, v in codigos_planillas.items():
        if isinstance(v, str) and len(v) == 4 and v.startswith("CM"):
            cm_to_planilla[v] = k
        elif isinstance(k, str) and len(k) == 4 and k.startswith("CM"):
            cm_to_planilla[k] = v

    datos_comparacion = {}

    for planilla, bloque_excel in datos_excel_planilla.items():
        codigo_cm   = bloque_excel["codigo_cm"]
        headers     = bloque_excel["headers"]
        filas_excel = bloque_excel["filas"]
        filas_modelo= datos_modelo_cm.get(codigo_cm, [])
        fe, fm, ff  = _comparar_por_planilla(codigo_cm, headers, filas_excel, filas_modelo)
        datos_comparacion[planilla] = {
            "codigo_cm": codigo_cm, "headers": headers,
            "filas_excel_base": fe, "filas_modelo_base": fm, "filas_fusionadas": ff
        }

    # CMs en modelo que no están en planilla
    for codigo_cm, lista_modelo in datos_modelo_cm.items():
        ya = any(info["codigo_cm"] == codigo_cm for info in datos_comparacion.values())
        if ya:
            continue
        nombre_planilla = cm_to_planilla.get(codigo_cm, codigo_cm)
        headers = cargar_headers_planilla(nombre_planilla, codigo_cm)
        if not headers:
            keys = sorted(set(k for fila in lista_modelo for k in fila.keys()))
            keys = [k for k in keys if k != HEADER_VINCULO]
            headers = _normalizar_headers(keys) if keys else ["CodIntBIM"]
        fe, fm, ff = _comparar_por_planilla(codigo_cm, headers, [], lista_modelo)
        datos_comparacion[nombre_planilla] = {
            "codigo_cm": codigo_cm, "headers": headers,
            "filas_excel_base": fe, "filas_modelo_base": fm, "filas_fusionadas": ff
        }

    log("construir_tabla_comparativa: {} planillas".format(len(datos_comparacion)))
    return datos_comparacion

# ══════════════════════════════════════════════════════════════
#  EXPORTAR A EXCEL
# ══════════════════════════════════════════════════════════════
def exportar_json_y_formatear(datos_comparacion):
    try:
        datos_por_tabla   = {}
        codigos_ordenados = []
        nombres_ordenados = []

        pares = sorted(
            [(bloque["codigo_cm"], nombre) for nombre, bloque in datos_comparacion.items()],
            key=lambda x: x[0]
        )
        for codigo_cm, nombre_planilla in pares:
            bloque = datos_comparacion[nombre_planilla]
            datos_por_tabla[codigo_cm] = {
                "headers": bloque["headers"],
                "filas"  : bloque.get("filas_fusionadas", [])
            }
            codigos_ordenados.append(codigo_cm)
            nombres_ordenados.append(nombre_planilla)

        datos_export = {
            "listado_tablas": {
                "valores": codigos_ordenados,
                "claves" : nombres_ordenados
            },
            "datos_por_tabla": datos_por_tabla
        }

        json_path = os.path.join(DATA_OUTPUT_DIR, "comparacion.json")
        if not os.path.exists(DATA_OUTPUT_DIR):
            os.makedirs(DATA_OUTPUT_DIR)
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(datos_export, f, ensure_ascii=False, indent=2)
        log("exportar: comparacion.json -> {}".format(json_path))

        subprocess.check_call(
            [PYTHON_EXE, FORMATEAR_SCRIPT, json_path, RUTA_XLSX_SALIDA],
            creationflags=CREATE_NO_WINDOW
        )
        log("exportar: Excel OK -> {}".format(RUTA_XLSX_SALIDA))
        messagebox.showinfo(
            "Exportar",
            "Archivo generado exitosamente:\n{}".format(RUTA_XLSX_SALIDA)
        )
    except subprocess.CalledProcessError as e:
        log_exc("exportar CalledProcessError")
        messagebox.showerror("Error exportando", str(e))
        raise
    except Exception:
        log_exc("exportar error general")
        messagebox.showerror("Error exportando", "Revisa ui_comparacion_log.txt")
        raise

# ══════════════════════════════════════════════════════════════
#  INTERFAZ DARK MODE (Tkinter)
# ══════════════════════════════════════════════════════════════
def aplicar_dark(root):
    root.configure(bg=BG_DARK)
    style = ttk.Style(root)
    style.theme_use("clam")
    style.configure(".",
        background=BG_DARK, foreground=TEXT_PRI,
        fieldbackground=SURFACE, bordercolor=BORDER_CLR,
        troughcolor=SURFACE, selectbackground=ACCENT,
        selectforeground=BG_DARK, font=("Segoe UI", 9)
    )
    style.configure("TFrame",    background=BG_DARK)
    style.configure("TLabel",    background=BG_DARK, foreground=TEXT_PRI)
    style.configure("TButton",
        background=SURFACE2, foreground=TEXT_PRI,
        borderwidth=0, padding=(10, 5), relief="flat"
    )
    style.map("TButton",
        background=[("active", ACCENT), ("pressed", ACCENT)],
        foreground=[("active", BG_DARK)]
    )
    style.configure("Accent.TButton",
        background=ACCENT, foreground=BG_DARK,
        borderwidth=0, padding=(12, 6), relief="flat"
    )
    style.map("Accent.TButton",
        background=[("active", "#227f8b"), ("pressed", "#1a626b")]
    )
    style.configure("TCombobox",
        background=SURFACE2, foreground=TEXT_PRI,
        fieldbackground=SURFACE2, arrowcolor=TEXT_PRI,
        bordercolor=BORDER_CLR, lightcolor=BORDER_CLR, darkcolor=BORDER_CLR
    )
    style.configure("TNotebook",          background=SURFACE,  bordercolor=BORDER_CLR)
    style.configure("TNotebook.Tab",
        background=SURFACE2, foreground=TEXT_MUTED, padding=(12, 5)
    )
    style.map("TNotebook.Tab",
        background=[("selected", BG_DARK)],
        foreground=[("selected", ACCENT)]
    )
    style.configure("Treeview",
        background=SURFACE, foreground=TEXT_PRI,
        fieldbackground=SURFACE, bordercolor=BORDER_CLR,
        rowheight=28
    )
    style.configure("Treeview.Heading",
        background=SURFACE2, foreground=ACCENT,
        relief="flat", borderwidth=0
    )
    style.map("Treeview",
        background=[("selected", ACCENT)],
        foreground=[("selected", BG_DARK)]
    )
    style.configure("Vertical.TScrollbar",
        background=SURFACE2, troughcolor=SURFACE, arrowcolor=TEXT_MUTED, borderwidth=0
    )
    style.configure("Horizontal.TScrollbar",
        background=SURFACE2, troughcolor=SURFACE, arrowcolor=TEXT_MUTED, borderwidth=0
    )

def crear_treeview_tabla(parent):
    frame = ttk.Frame(parent)
    frame.pack(fill=tk.BOTH, expand=True)
    tree = ttk.Treeview(frame, show="headings")
    tree.grid(row=0, column=0, sticky="nsew")
    vsb = ttk.Scrollbar(frame, orient="vertical",   command=tree.yview)
    vsb.grid(row=0, column=1, sticky="ns")
    hsb = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
    hsb.grid(row=1, column=0, sticky="ew")
    frame.rowconfigure(0, weight=1)
    frame.columnconfigure(0, weight=1)
    tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
    for estado, color in COLOR_ESTADOS.items():
        tree.tag_configure(
            estado, background=color,
            foreground=COLOR_ESTADOS_TEXT.get(estado, TEXT_PRI)
        )
    return tree

def limpiar_treeview(tree):
    for item in tree.get_children():
        tree.delete(item)
    tree["columns"] = ()

def poblar_treeview(tree, headers, filas):
    limpiar_treeview(tree)
    if not headers:
        return
    tree["columns"] = list(headers)
    for h in headers:
        tree.heading(h, text=h)
        tree.column(h, width=180, stretch=False, anchor="w")
    if not filas:
        return

    def peor_estado(estados):
        for e in ("difiere", "falta_modelo", "falta_excel", "no_existe"):
            if e in estados:
                return e
        return "ok"

    for fila in filas:
        valores = list(fila.get("valores", []))
        if len(valores) < len(headers):
            valores += [""] * (len(headers) - len(valores))
        else:
            valores = valores[:len(headers)]
        estados = fila.get("estado_por_celda") or []
        if len(estados) < len(headers):
            estados += ["ok"] * (len(headers) - len(estados))
        tree.insert("", "end", values=valores, tags=(peor_estado(estados),))

def mostrar_ui(datos_comparacion):
    root = tk.Tk()
    root.title("Planilla vs Modelo")
    root.geometry("1280x700")
    root.configure(bg=BG_DARK)
    aplicar_dark(root)

    export_realizado = {"value": False}

    # ── Header ──
    frame_header = ttk.Frame(root)
    frame_header.pack(fill=tk.X, padx=12, pady=(10, 4))
    ttk.Label(
        frame_header, text="Planilla vs Modelo",
        font=("Segoe UI", 14, "bold"), foreground=ACCENT
    ).pack(side=tk.LEFT)

    # ── Selector de tabla ──
    frame_top = ttk.Frame(root)
    frame_top.pack(fill=tk.X, padx=12, pady=4)
    ttk.Label(frame_top, text="Tabla:", foreground=TEXT_MUTED).pack(side=tk.LEFT, padx=(0, 6))
    nombres_disponibles = sorted(datos_comparacion.keys())
    selected_nombre = tk.StringVar()
    combo = ttk.Combobox(
        frame_top, textvariable=selected_nombre,
        values=nombres_disponibles, state="readonly", width=55
    )
    combo.pack(side=tk.LEFT, fill=tk.X, expand=True)

    # ── Notebook (Desde Excel / Desde Modelo) ──
    frame_center = ttk.Frame(root)
    frame_center.pack(fill=tk.BOTH, expand=True, padx=12, pady=4)
    notebook = ttk.Notebook(frame_center)
    notebook.pack(fill=tk.BOTH, expand=True)
    frame_excel  = ttk.Frame(notebook)
    frame_modelo = ttk.Frame(notebook)
    notebook.add(frame_excel,  text="  Desde Excel  ")
    notebook.add(frame_modelo, text="  Desde Modelo  ")
    tree_excel  = crear_treeview_tabla(frame_excel)
    tree_modelo = crear_treeview_tabla(frame_modelo)

    # ── Footer ──
    frame_footer = tk.Frame(root, bg=SURFACE, height=48)
    frame_footer.pack(fill=tk.X, side=tk.BOTTOM)
    frame_footer.pack_propagate(False)

    # Leyenda
    leyenda_info = [
        ("ok",           "OK / Coincide"),
        ("falta_modelo", "Falta en Modelo"),
        ("falta_excel",  "Falta en Excel"),
        ("difiere",      "Diferencia"),
        ("no_existe",    "No existe"),
    ]
    frame_leyenda = tk.Frame(frame_footer, bg=SURFACE)
    frame_leyenda.pack(side=tk.LEFT, padx=12, pady=10)
    for estado, texto in leyenda_info:
        color = COLOR_ESTADOS[estado]
        tcolor= COLOR_ESTADOS_TEXT.get(estado, TEXT_PRI)
        lbl = tk.Label(
            frame_leyenda, text="  {}  ".format(texto),
            bg=color, fg=tcolor, font=("Segoe UI", 8),
            relief="flat", padx=4, pady=2
        )
        lbl.pack(side=tk.LEFT, padx=3)

    # Botones
    frame_botones = tk.Frame(frame_footer, bg=SURFACE)
    frame_botones.pack(side=tk.RIGHT, padx=12, pady=8)

    def actualizar_tablas(nombre):
        bloque = datos_comparacion.get(nombre)
        if not bloque:
            limpiar_treeview(tree_excel)
            limpiar_treeview(tree_modelo)
            return
        headers = bloque.get("headers", [])
        poblar_treeview(tree_excel,  headers, bloque.get("filas_excel_base",  []))
        poblar_treeview(tree_modelo, headers, bloque.get("filas_modelo_base", []))

    def on_sel_planilla(event=None):
        nombre = selected_nombre.get()
        if nombre:
            actualizar_tablas(nombre)
    combo.bind("<<ComboboxSelected>>", on_sel_planilla)

    def cmd_exportar():
        exportar_json_y_formatear(datos_comparacion)
        export_realizado["value"] = True

    def cmd_salir():
        if not export_realizado["value"]:
            if not messagebox.askyesno(
                "Salir sin exportar",
                "¿Seguro que quieres salir sin exportar el archivo Excel?"
            ):
                return
        root.destroy()

    btn_exportar = tk.Button(
        frame_botones, text="  Exportar Excel  ",
        bg=ACCENT, fg=BG_DARK, font=("Segoe UI", 9, "bold"),
        relief="flat", padx=10, pady=4, cursor="hand2",
        command=cmd_exportar
    )
    btn_exportar.pack(side=tk.LEFT, padx=5)

    btn_salir = tk.Button(
        frame_botones, text="  Salir  ",
        bg=SURFACE2, fg=TEXT_PRI, font=("Segoe UI", 9),
        relief="flat", padx=10, pady=4, cursor="hand2",
        command=cmd_salir
    )
    btn_salir.pack(side=tk.LEFT, padx=5)

    if nombres_disponibles:
        selected_nombre.set(nombres_disponibles[0])
        actualizar_tablas(nombres_disponibles[0])

    root.protocol("WM_DELETE_WINDOW", cmd_salir)
    root.mainloop()

# ══════════════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════════════
if __name__ == "__main__":
    try:
        log("==== Inicio ui_comparacion.py ====")
        log("Args: script_json={} | csv={} | data_output={} | "
            "formatear={} | xlsx={} | python={} | modelo={}".format(
            SCRIPT_JSON_PATH, CSV_CODIGOS_PATH, DATA_OUTPUT_DIR,
            FORMATEAR_SCRIPT, RUTA_XLSX_SALIDA, PYTHON_EXE, MODELO_JSON_PATH
        ))

        # Validar archivos críticos
        faltantes = []
        for ruta, nombre in [
            (SCRIPT_JSON_PATH,  "script.json"),
            (CSV_CODIGOS_PATH,  "CODIGO.csv"),
            (MODELO_JSON_PATH,  "modelo_codint_por_cm.json"),
            (FORMATEAR_SCRIPT,  "formatear_tablas_excel.py"),
        ]:
            if not os.path.exists(ruta):
                faltantes.append("• {} -> {}".format(nombre, ruta))
        if faltantes:
            msg = "Archivos no encontrados:\n\n{}".format("\n".join(faltantes))
            log(msg)
            messagebox.showerror("Error — Archivos faltantes", msg)
            sys.exit(1)

        cfg               = leer_script_json(SCRIPT_JSON_PATH)
        codigos_planillas = cfg.get("codigos_planillas", {})

        filas_csv            = leer_csv_codigos(CSV_CODIGOS_PATH)
        datos_excel_planilla = construir_excel_por_planilla(filas_csv, codigos_planillas)

        datos_modelo_cm_raw = leer_modelo_por_cm(MODELO_JSON_PATH)
        repo_activo         = cargar_repo_activo()
        datos_modelo_cm     = enriquecer_modelo_con_bd(datos_modelo_cm_raw, repo_activo)

        datos_comparacion   = construir_tabla_comparativa(
            datos_excel_planilla, datos_modelo_cm, codigos_planillas
        )
        log("main: {} planillas en datos_comparacion".format(len(datos_comparacion)))

        mostrar_ui(datos_comparacion)
        log("==== Fin ui_comparacion.py ====")

    except SystemExit:
        raise
    except Exception:
        log_exc("ui_comparacion main error general")
        try:
            messagebox.showerror(
                "Error inesperado",
                "Error en ui_comparacion.py.\nRevisa:\n{}".format(UI_LOG_PATH)
            )
        except Exception:
            pass
        sys.exit(1)