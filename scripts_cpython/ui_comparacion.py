# -*- coding: utf-8 -*-
"""
scripts_cpython/ui_comparacion.py  v1.6
Cambios v1.5:
- [FIX] leer_xlsm_codigos: keep_vba=False + manejo robusto de xlsm con macros.
- [FIX] exportar_json_y_formatear: clave "elementos_por_tabla" (no "datos_por_tabla")
        para que formatear_tablas_excel_v2.py la lea correctamente.
- [FIX] headers_por_tabla incluido en comparacion.json con clave correcta.

Cambios v1.6:
- [FIX] Combobox "Tabla" muestra nombre de planilla legible en lugar del codigo CM raw.
- [ADD] _etiqueta_cm() convierte claves "NombrePlanilla::CM01" -> "NombrePlanilla  (CM01)".
- [ADD] _mapa_etiqueta_clave resuelve la etiqueta visible a la clave interna del dict.
"""

import sys
import os
import json
import subprocess
import traceback
import tkinter as tk
from tkinter import ttk, messagebox

COLOR_ESTADOS = {
    'ok':           '#C6EFCE',
    'falta_modelo': '#FFC7CE',
    'falta_excel':  '#FFEB9C',
    'difiere':      '#F4B084',
    'no_existe':    '#66FFFF',
}

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_EXT_ROOT   = os.path.abspath(os.path.join(_SCRIPT_DIR, '..'))

DATA_DIR_EXT           = os.path.join(_EXT_ROOT, 'data')
PLANILLAS_HEADERS_JSON = os.path.join(DATA_DIR_EXT, 'master', 'planillas_headers_order.json')
CONFIG_PROYECTO_ACTIVO = os.path.join(DATA_DIR_EXT, 'master', 'config_proyecto_activo.json')
PROYECTOS_DIR          = os.path.join(DATA_DIR_EXT, 'proyectos')
UI_LOG_PATH            = os.path.join(DATA_DIR_EXT, 'logs', 'ui_comparacion_log.txt')
HEADER_VINCULO         = u"Vínculo RVT: Nombre de archivo"


def log(msg):
    try:
        from datetime import datetime
        ts    = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        linea = u"[{}] {}\n".format(ts, msg)
        os.makedirs(os.path.dirname(UI_LOG_PATH), exist_ok=True)
        with open(UI_LOG_PATH, "a", encoding="utf-8") as f:
            f.write(linea)
    except Exception:
        pass


def log_exc(contexto):
    try:
        tb = traceback.format_exc()
        log("{} -> {}".format(contexto, tb))
    except Exception:
        pass


# ── JSON ──────────────────────────────────────────────────────────────────────
def cargar_json(ruta, default):
    try:
        if not os.path.exists(ruta):
            log("cargar_json: no existe -> {}".format(ruta))
            return default
        with open(ruta, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        log_exc("cargar_json {}".format(ruta))
        return default


# ── Repo activo ───────────────────────────────────────────────────────────────
def get_repo_activo_path():
    cfg  = cargar_json(CONFIG_PROYECTO_ACTIVO, {})
    ruta = (cfg.get("ruta_repositorio_activo") or "").strip()
    if ruta and os.path.exists(ruta):
        return ruta
    nup = (cfg.get("nup_activo") or "").strip()
    if nup:
        return os.path.join(PROYECTOS_DIR, u'repositorio_datos_{}.json'.format(nup))
    return ""


def cargar_repo_activo():
    ruta = get_repo_activo_path()
    if not ruta or not os.path.exists(ruta):
        log("cargar_repo_activo: no disponible -> {}".format(ruta))
        return {}
    repo = cargar_json(ruta, {})
    log("cargar_repo_activo: {} registros".format(len(repo)))
    return repo


# ── Leer xlsm con openpyxl ────────────────────────────────────────────────────
def leer_xlsm_codigos(ruta_xlsm):
    """
    Lee todas las hojas del .xlsm con openpyxl.
    - keep_vba=False  -> ignora las macros VBA (evita errores de parsing).
    - data_only=True  -> lee valores calculados, no formulas.
    - read_only=True  -> apertura rapida sin cargar el DOM completo.
    Devuelve lista de listas (una por fila donde col[0] empieza con 'CM').
    """
    try:
        import openpyxl
    except ImportError:
        log("leer_xlsm_codigos: openpyxl no instalado")
        messagebox.showerror(
            "Dependencia faltante",
            "openpyxl no esta instalado.\n"
            "Ejecuta: pip install openpyxl"
        )
        raise

    if not os.path.exists(ruta_xlsm):
        raise FileNotFoundError("No se encontro el archivo: {}".format(ruta_xlsm))

    ext = os.path.splitext(ruta_xlsm)[1].lower()
    log("leer_xlsm_codigos: extension={} ruta={}".format(ext, ruta_xlsm))

    filas = []
    try:
        wb = openpyxl.load_workbook(
            ruta_xlsm,
            read_only=True,
            data_only=True,
            keep_vba=False
        )
        for nombre_hoja in wb.sheetnames:
            try:
                ws = wb[nombre_hoja]
                for row in ws.iter_rows(values_only=True):
                    if not row:
                        continue
                    primera = str(row[0] or '').strip()
                    if len(primera) >= 4 and primera[:2] == 'CM':
                        filas.append([str(c or '').strip() for c in row])
            except Exception as e_hoja:
                log("leer_xlsm_codigos: error en hoja '{}' -> {}".format(nombre_hoja, e_hoja))
                continue
        wb.close()
        log("leer_xlsm_codigos: {} filas CM desde {}".format(len(filas), ruta_xlsm))
        return filas

    except Exception:
        log_exc("leer_xlsm_codigos error en {}".format(ruta_xlsm))
        raise


# ── Headers ───────────────────────────────────────────────────────────────────
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
    if not os.path.exists(PLANILLAS_HEADERS_JSON):
        return None
    try:
        with open(PLANILLAS_HEADERS_JSON, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        log_exc("cargar_headers_planilla {}".format(PLANILLAS_HEADERS_JSON))
        return None

    clave   = "{}::{}".format(nombre_planilla, codigo_cm)
    headers = data.get(clave) or data.get(codigo_cm)
    if headers:
        h_norm = _normalizar_headers(headers)
        log("cargar_headers_planilla: {} -> {} cols".format(nombre_planilla, len(h_norm)))
        return h_norm
    return None


# ── Construir Excel por planilla ──────────────────────────────────────────────
def construir_excel_por_planilla(filas_xlsm, codigos_planillas):
    planilla_a_cm = {}
    for k, v in codigos_planillas.items():
        try:
            if isinstance(k, str) and len(k) == 4 and k.startswith("CM"):
                planilla_a_cm[k] = v
            elif isinstance(v, str) and len(v) == 4 and v.startswith("CM"):
                planilla_a_cm[v] = k
        except Exception:
            pass

    log("construir_excel_por_planilla: planilla_a_cm -> {}".format(planilla_a_cm))
    datos_excel_planilla = {}

    for row in filas_xlsm:
        if not row:
            continue
        primera_celda = (row[0] or '').strip()
        if not primera_celda or len(primera_celda) < 4:
            continue
        pref = primera_celda[:4]

        nombre_planilla = planilla_a_cm.get(pref)
        if not nombre_planilla:
            continue

        headers = cargar_headers_planilla(nombre_planilla, pref)
        if not headers:
            headers = _normalizar_headers(["Pa{}".format(i) for i in range(1, len(row) + 1)])

        if nombre_planilla not in datos_excel_planilla:
            datos_excel_planilla[nombre_planilla] = {
                "codigo_cm": pref,
                "headers":   headers,
                "filas":     []
            }

        bloque  = datos_excel_planilla[nombre_planilla]
        valores = []
        for idx_h in range(len(headers)):
            valores.append(row[idx_h] if idx_h < len(row) else '')
        codint = (valores[0] or '').strip()
        bloque["filas"].append({"CodIntBIM": codint, "valores": valores})

    log("construir_excel_por_planilla: {} planillas".format(len(datos_excel_planilla)))
    return datos_excel_planilla


# ── Modelo ────────────────────────────────────────────────────────────────────
def leer_modelo_por_cm(path_json_modelo):
    try:
        if not os.path.exists(path_json_modelo):
            log("leer_modelo_por_cm: no existe -> {}".format(path_json_modelo))
            return {}
        with open(path_json_modelo, 'r', encoding='utf-8') as f:
            data = json.load(f)
        result = {cm: lista for cm, lista in data.items() if isinstance(lista, list)}
        log("leer_modelo_por_cm: CMs -> {}".format(list(result.keys())))
        return result
    except Exception:
        log_exc("leer_modelo_por_cm")
        raise


def enriquecer_modelo_con_bd(datos_modelo_cm, repo):
    if not repo:
        return datos_modelo_cm
    indice_bd = {}
    for k, v in repo.items():
        eid = str(v.get("ElementId", "")).strip()
        cod = (v.get("CodIntBIM") or "").strip()
        if eid and cod:
            indice_bd[(eid, cod)] = v

    nuevo = {}
    for cm, lista in datos_modelo_cm.items():
        nueva_lista = []
        for fila in lista:
            cod    = (fila.get("CodIntBIM") or '').strip()
            eid    = str(fila.get("ElementId", "")).strip()
            bd_row = indice_bd.get((eid, cod))
            if bd_row:
                combinado = dict(fila)
                combinado.update(bd_row)
                nueva_lista.append(combinado)
            else:
                nueva_lista.append(fila)
        nuevo[cm] = nueva_lista
    return nuevo


# ── Comparacion ───────────────────────────────────────────────────────────────
def _norm_val(v):
    v = (v or '').strip()
    return '' if v == '-' else v


def _get_val(row_dict, header_name):
    if header_name == HEADER_VINCULO:
        return ''
    v = (row_dict.get(header_name) or '').strip()
    return '' if v == '-' else v


def _comparar_por_planilla(codigo_cm, headers, filas_excel, filas_modelo):
    mapa_excel  = {}
    for fila in filas_excel:
        c = (fila["CodIntBIM"] or '').strip()
        if c:
            mapa_excel.setdefault(c, []).append(fila)

    mapa_modelo = {}
    for row_m in filas_modelo:
        c = (row_m.get("CodIntBIM") or '').strip()
        if c:
            mapa_modelo.setdefault(c, []).append(row_m)

    filas_tabla_excel  = []
    filas_tabla_modelo = []
    filas_fusionadas   = []
    codigos_union      = sorted(set(mapa_excel.keys()) | set(mapa_modelo.keys()))

    for cod in codigos_union:
        lista_excel  = mapa_excel.get(cod, [])
        lista_modelo = mapa_modelo.get(cod, [])
        fila_x0 = lista_excel[0]  if lista_excel  else None
        fila_m0 = lista_modelo[0] if lista_modelo else None

        valores_fusion = []
        estados_fusion = []
        for h in headers:
            vx_raw = ''
            vm_raw = ''
            if fila_x0:
                idx_h  = headers.index(h)
                vals   = fila_x0["valores"]
                vx_raw = vals[idx_h] if idx_h < len(vals) else ''
            if fila_m0:
                vm_raw = _get_val(fila_m0, h)
            vx = _norm_val(vx_raw)
            vm = _norm_val(vm_raw)

            if not vx and not vm:
                valores_fusion.append(''); estados_fusion.append('no_existe')
            elif vx == vm:
                valores_fusion.append(vx_raw); estados_fusion.append('ok')
            elif vx and not vm:
                valores_fusion.append(vx_raw); estados_fusion.append('falta_modelo')
            elif not vx and vm:
                valores_fusion.append(vm_raw); estados_fusion.append('falta_excel')
            else:
                valores_fusion.append(u"Planilla: {}\nModelo: {}".format(vx_raw, vm_raw))
                estados_fusion.append('difiere')

        filas_fusionadas.append({
            "CodIntBIM": cod, "valores": valores_fusion,
            "estado_por_celda": estados_fusion
        })

        for fila_x in lista_excel:
            vals_x = fila_x["valores"]
            valores, estados = [], []
            fila_m = lista_modelo[0] if lista_modelo else None
            for idx, h in enumerate(headers):
                vx_raw = vals_x[idx] if idx < len(vals_x) else ''
                vx     = _norm_val(vx_raw)
                vm_raw = _get_val(fila_m, h) if fila_m else ''
                vm     = _norm_val(vm_raw)
                if not vx and not vm:
                    valores.append(''); estados.append('no_existe')
                elif vx == vm:
                    valores.append(vx_raw); estados.append('ok')
                elif vx and not vm:
                    valores.append(vx_raw); estados.append('falta_modelo')
                elif not vx and vm:
                    valores.append(u"Planilla: {}\nModelo: {}".format(vx_raw, vm_raw))
                    estados.append('falta_excel')
                else:
                    valores.append(u"Planilla: {}\nModelo: {}".format(vx_raw, vm_raw))
                    estados.append('difiere')
            filas_tabla_excel.append({
                "origen": "excel", "CodIntBIM": cod,
                "valores": valores, "estado_por_celda": estados
            })

        for fila_m in lista_modelo:
            valores, estados = [], []
            fila_x = lista_excel[0] if lista_excel else None
            for idx, h in enumerate(headers):
                vm_raw = _get_val(fila_m, h)
                vm     = _norm_val(vm_raw)
                vals_x = fila_x["valores"] if fila_x else []
                vx_raw = vals_x[idx] if idx < len(vals_x) else ''
                vx     = _norm_val(vx_raw)
                if not vx and not vm:
                    valores.append(''); estados.append('no_existe')
                elif vx == vm:
                    valores.append(vm_raw); estados.append('ok')
                elif vm and not vx:
                    valores.append(vm_raw); estados.append('falta_excel')
                elif not vm and vx:
                    valores.append(u"Modelo: {}\nPlanilla: {}".format(vm_raw, vx_raw))
                    estados.append('falta_modelo')
                else:
                    valores.append(u"Modelo: {}\nPlanilla: {}".format(vm_raw, vx_raw))
                    estados.append('difiere')
            filas_tabla_modelo.append({
                "origen": "modelo", "CodIntBIM": cod,
                "valores": valores, "estado_por_celda": estados
            })

    return filas_tabla_excel, filas_tabla_modelo, filas_fusionadas


def construir_tabla_comparativa(datos_excel_planilla, datos_modelo_cm, codigos_planillas):
    datos_comparacion = {}
    cm_to_planilla    = {}
    for k, v in codigos_planillas.items():
        try:
            if isinstance(v, str) and len(v) == 4 and v.startswith("CM"):
                cm_to_planilla[v] = k
            elif isinstance(k, str) and len(k) == 4 and k.startswith("CM"):
                cm_to_planilla[k] = v
        except Exception:
            pass

    for planilla, bloque_excel in datos_excel_planilla.items():
        codigo_cm    = bloque_excel['codigo_cm']
        headers      = bloque_excel['headers']
        filas_excel  = bloque_excel['filas']
        filas_modelo = datos_modelo_cm.get(codigo_cm, [])
        fe, fm, ff   = _comparar_por_planilla(codigo_cm, headers, filas_excel, filas_modelo)
        datos_comparacion[planilla] = {
            "codigo_cm": codigo_cm, "headers": headers,
            "filas_excel_base": fe, "filas_modelo_base": fm, "filas_fusionadas": ff
        }

    for codigo_cm, lista_modelo in datos_modelo_cm.items():
        if any(i["codigo_cm"] == codigo_cm for i in datos_comparacion.values()):
            continue
        nombre_planilla = cm_to_planilla.get(codigo_cm, codigo_cm)
        headers = cargar_headers_planilla(nombre_planilla, codigo_cm)
        if not headers:
            keys = sorted(set(k for fila in lista_modelo for k in fila.keys())) if lista_modelo else []
            keys = [k for k in keys if k != HEADER_VINCULO]
            headers = _normalizar_headers(["CodIntBIM"] + [k for k in keys if k != "CodIntBIM"])
        fe, fm, ff = _comparar_por_planilla(codigo_cm, headers, [], lista_modelo)
        datos_comparacion[nombre_planilla] = {
            "codigo_cm": codigo_cm, "headers": headers,
            "filas_excel_base": fe, "filas_modelo_base": fm, "filas_fusionadas": ff
        }

    log("construir_tabla_comparativa: {} planillas".format(len(datos_comparacion)))
    return datos_comparacion


# ── Exportar Excel ────────────────────────────────────────────────────────────
def exportar_json_y_formatear(datos_comparacion, data_dir,
                              formatear_script, ruta_xlsx_salida, python_exe):
    """
    Escribe comparacion.json con las claves que formatear_tablas_excel_v2.py espera:
      - "elementos_por_tabla"
      - "headers_por_tabla"
      - "listado_tablas"
    """
    try:
        elementos_por_tabla = {}
        headers_por_tabla   = {}
        codigos, nombres    = [], []

        for nombre_planilla, bloque in datos_comparacion.items():
            codigo_cm = bloque['codigo_cm']
            headers   = bloque['headers']
            filas     = bloque.get('filas_fusionadas', [])

            elementos = []
            for fila in filas:
                vals = fila.get('valores', [])
                elem = {}
                for idx, h in enumerate(headers):
                    elem[h] = vals[idx] if idx < len(vals) else ''
                elementos.append(elem)

            elementos_por_tabla[codigo_cm] = elementos
            headers_por_tabla[codigo_cm]   = headers
            codigos.append(codigo_cm)
            nombres.append(nombre_planilla)

        pares             = sorted(zip(codigos, nombres), key=lambda x: x[0])
        codigos_ordenados = [p[0] for p in pares]
        nombres_ordenados = [p[1] for p in pares]

        datos_export = {
            "listado_tablas": {
                "valores": codigos_ordenados,
                "claves":  nombres_ordenados
            },
            "elementos_por_tabla": elementos_por_tabla,
            "headers_por_tabla":   headers_por_tabla,
        }

        os.makedirs(data_dir, exist_ok=True)
        ruta_json = os.path.join(data_dir, "comparacion.json")
        with open(ruta_json, "w", encoding="utf-8") as f:
            json.dump(datos_export, f, ensure_ascii=False, indent=2)
        log("exportar_json_y_formatear: JSON escrito -> {}".format(ruta_json))

        cmd = [python_exe, formatear_script,
               ruta_json, ruta_xlsx_salida]
        log("exportar_json_y_formatear: cmd -> {}".format(cmd))
        result = subprocess.run(
            cmd,
            capture_output=True, text=True,
            timeout=120
        )
        if result.returncode != 0:
            log("exportar_json_y_formatear: ERROR subprocess -> {}".format(result.stderr))
            messagebox.showerror(
                "Error al exportar",
                "El formateador termino con error:\n{}".format(result.stderr[:500])
            )
        else:
            log("exportar_json_y_formatear: OK -> {}".format(ruta_xlsx_salida))

    except Exception:
        log_exc("exportar_json_y_formatear")
        messagebox.showerror(
            "Error al exportar",
            "Error inesperado al generar el Excel.\nRevisa ui_comparacion_log.txt."
        )


# ── Treeview helpers ──────────────────────────────────────────────────────────
def limpiar_treeview(tree):
    for item in tree.get_children():
        tree.delete(item)
    tree["columns"] = []


def crear_treeview_tabla(parent):
    frame = ttk.Frame(parent)
    frame.pack(fill=tk.BOTH, expand=True)

    vsb = ttk.Scrollbar(frame, orient=tk.VERTICAL)
    hsb = ttk.Scrollbar(frame, orient=tk.HORIZONTAL)
    tree = ttk.Treeview(frame, show="headings",
                        yscrollcommand=vsb.set, xscrollcommand=hsb.set)
    vsb.config(command=tree.yview)
    hsb.config(command=tree.xview)

    vsb.pack(side=tk.RIGHT,  fill=tk.Y)
    hsb.pack(side=tk.BOTTOM, fill=tk.X)
    tree.pack(fill=tk.BOTH, expand=True)

    for estado, color in COLOR_ESTADOS.items():
        tree.tag_configure(estado, background=color)

    return tree


def poblar_treeview(tree, headers, filas_tabla):
    limpiar_treeview(tree)
    if not headers:
        return
    columnas = list(headers)
    tree["columns"] = columnas
    for h in columnas:
        tree.heading(h, text=h)
        tree.column(h, width=200, stretch=False, anchor="w")
    if not filas_tabla:
        return

    def peor_estado(estados):
        for e in ('difiere', 'falta_modelo', 'falta_excel', 'no_existe'):
            if e in estados:
                return e
        return 'ok'

    for fila in filas_tabla:
        valores = list(fila.get('valores', []))
        estados = list(fila.get('estado_por_celda') or [])
        if len(valores) < len(columnas):
            valores += [""] * (len(columnas) - len(valores))
        if len(estados) < len(columnas):
            estados += ['ok'] * (len(columnas) - len(estados))
        tree.insert('', 'end', values=valores[:len(columnas)],
                    tags=(peor_estado(estados),))


def mostrar_ui(datos_comparacion, on_exportar):
    root = tk.Tk()
    root.title("Planilla vs Modelo")
    

