# -*- coding: utf-8 -*-

import sys
import os
import csv
import json
import subprocess
import traceback
import tkinter as tk
from tkinter import ttk, messagebox

COLOR_ESTADOS = {
    'ok': '#C6EFCE',
    'falta_modelo': '#FFC7CE',
    'falta_excel': '#FFEB9C',
    'difiere': '#F4B084',
    'no_existe': '#66FFFF',   # solo para leyenda (UI + Excel)
}

DATA_DIR_EXT = os.path.join(
    os.path.expanduser("~"),
    r"AppData\\Roaming\\MyPyRevitExtention\\PyRevitIT.extension\\data"
)

PLANILLAS_HEADERS_JSON = os.path.join(
    DATA_DIR_EXT,
    "planillas_headers_order.json"
)

CONFIG_PROYECTO_ACTIVO = os.path.join(
    DATA_DIR_EXT,
    "config_proyecto_activo.json"
)

UI_LOG_PATH = os.path.join(DATA_DIR_EXT, "ui_comparacion_log.txt")

HEADER_VINCULO = u"Vínculo RVT: Nombre de archivo"


def log(msg):
    try:
        from datetime import datetime
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        linea = u"[{}] {}\n".format(ts, msg)
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


# ============================================================================
#                     UTILIDADES BD/REPOSITORIO
# ============================================================================

def cargar_json(ruta, default):
    try:
        if not os.path.exists(ruta):
            log("cargar_json: archivo no existe -> {}".format(ruta))
            return default
        with open(ruta, "r", encoding="utf-8") as f:
            data = json.load(f)
        log("cargar_json: cargado OK -> {}".format(ruta))
        return data
    except Exception:
        log_exc("cargar_json error leyendo {}".format(ruta))
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


# ============================================================================
#                       CARGA DE DATOS BASE
# ============================================================================

def leer_script_json(path):
    try:
        with open(path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        log("leer_script_json: OK -> {}".format(path))
        return data
    except Exception:
        log_exc("leer_script_json error en {}".format(path))
        raise


def leer_csv_codigos(csv_codigos):
    filas = []
    try:
        with open(csv_codigos, 'r', encoding='utf-8') as f:
            reader = csv.reader(f, delimiter=';')
            for row in reader:
                if not row:
                    continue
                if not any((c or '').strip() for c in row):
                    continue
                filas.append(row)
        log("leer_csv_codigos: {} filas desde {}".format(len(filas), csv_codigos))
        return filas
    except Exception:
        log_exc("leer_csv_codigos error en {}".format(csv_codigos))
        raise


def leer_modelo_por_cm(path_json_modelo):
    try:
        if not os.path.exists(path_json_modelo):
            log("leer_modelo_por_cm: archivo no existe -> {}".format(path_json_modelo))
            return {}
        with open(path_json_modelo, 'r', encoding='utf-8') as f:
            data = json.load(f)
        result = {}
        for cm, lista in data.items():
            if isinstance(lista, list):
                result[cm] = lista
        log("leer_modelo_por_cm: CM -> {} desde {}".format(list(result.keys()), path_json_modelo))
        return result
    except Exception:
        log_exc("leer_modelo_por_cm error en {}".format(path_json_modelo))
        raise


def enriquecer_modelo_con_bd(datos_modelo_cm, repo):
    """Mantiene orden y estructura del modelo; solo sobreescribe campos con la BD."""
    try:
        if not repo:
            log("enriquecer_modelo_con_bd: repo vacío, se devuelve modelo sin cambios")
            return datos_modelo_cm

        indice_bd = {}
        for k, v in repo.items():
            eid = str(v.get("ElementId", "")).strip()
            cod = (v.get("CodIntBIM") or "").strip()
            if not eid or not cod:
                continue
            indice_bd[(eid, cod)] = v

        log("enriquecer_modelo_con_bd: índice BD con {} claves".format(len(indice_bd)))

        nuevo = {}
        for cm, lista in datos_modelo_cm.items():
            nueva_lista = []
            for fila in lista:
                cod = (fila.get("CodIntBIM") or '').strip()
                eid = str(fila.get("ElementId", "")).strip()
                if not cod or not eid:
                    nueva_lista.append(fila)
                    continue
                bd_row = indice_bd.get((eid, cod))
                if bd_row:
                    combinado = dict(fila)
                    for k, v in bd_row.items():
                        combinado[k] = v
                    nueva_lista.append(combinado)
                else:
                    nueva_lista.append(fila)
            nuevo[cm] = nueva_lista

        log("enriquecer_modelo_con_bd: modelo enriquecido para 1 CM")
        return nuevo
    except Exception:
        log_exc("enriquecer_modelo_con_bd")
        raise


def _normalizar_headers(headers_raw):
    """
    - Elimina HEADER_VINCULO.
    - Asegura que CodIntBIM esté siempre en la primera posición
      (lo inserta si no existe).
    """
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
        log("cargar_headers_planilla: no existe PLANILLAS_HEADERS_JSON -> {}".format(PLANILLAS_HEADERS_JSON))
        return None

    try:
        with open(PLANILLAS_HEADERS_JSON, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        log_exc("cargar_headers_planilla error leyendo {}".format(PLANILLAS_HEADERS_JSON))
        return None

    clave = "{}::{}".format(nombre_planilla, codigo_cm)
    headers = data.get(clave)
    if headers:
        h_norm = _normalizar_headers(headers)
        log("cargar_headers_planilla: {} -> {} (clave {})".format(nombre_planilla, len(h_norm), clave))
        return h_norm

    headers = data.get(codigo_cm)
    if headers:
        h_norm = _normalizar_headers(headers)
        log("cargar_headers_planilla: {} -> {} (clave {})".format(nombre_planilla, len(h_norm), codigo_cm))
        return h_norm

    log("cargar_headers_planilla: sin headers para {} / {}".format(nombre_planilla, codigo_cm))
    return None


def construir_excel_por_planilla(filas_csv, codigos_planillas):
    """
    Usa BD de encabezados normalizada (CodIntBIM primero, sin HEADER_VINCULO)
    y mapea 1:1 headers[i] ↔ fila_planilla[i].
    """
    planilla_a_cm = {}
    for k, v in codigos_planillas.items():
        try:
            if len(k) == 4 and k.startswith("CM"):
                planilla_a_cm[v] = k
            elif isinstance(v, str) and len(v) == 4 and v.startswith("CM"):
                planilla_a_cm[k] = v
        except Exception:
            log_exc("construir_excel_por_planilla armando planilla_a_cm")

    log("construir_excel_por_planilla: planilla_a_cm -> {}".format(planilla_a_cm))

    datos_excel_planilla = {}

    for row in filas_csv:
        if not row:
            continue

        row_filtrada = [c for c in row]

        primera_celda = (row_filtrada[0] or '').strip()
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
            num_cols = len(row_filtrada)
            headers = _normalizar_headers(["Pa{}".format(i) for i in range(1, num_cols + 1)])

        bloque = datos_excel_planilla.get(nombre_planilla)
        if not bloque:
            bloque = {
                "codigo_cm": pref,
                "headers": headers,
                "filas": []
            }
            datos_excel_planilla[nombre_planilla] = bloque

        valores = []
        for idx_h, _ in enumerate(headers):
            idx_csv = idx_h
            val = row_filtrada[idx_csv] if idx_csv < len(row_filtrada) else ''
            valores.append(val)

        codint = (valores[0] or '').strip()

        bloque["filas"].append({
            "CodIntBIM": codint,
            "valores": valores
        })

    log("construir_excel_por_planilla: {} planillas construidas".format(len(datos_excel_planilla)))
    return datos_excel_planilla


def _norm_val(v):
    """Normaliza valores para comparación: trata '-' como vacío."""
    v = (v or '').strip()
    if v == '-':
        return ''
    return v


def _get_val(row_dict, header_name):
    if header_name == HEADER_VINCULO:
        return ''
    v = (row_dict.get(header_name) or '').strip()
    if v == '-':
        return ''
    return v


# ============================================================================
#               COMPARACIÓN EXCEL ↔ MODELO (POR CELDA)
# ============================================================================

def construir_tabla_comparativa(datos_excel_planilla, datos_modelo_cm, codigos_planillas):
    """
    Devuelve:
      datos_comparacion = {
        nombre_planilla: {
          "codigo_cm": ...,
          "headers": [...],
          "filas_excel_base": [...],   # Para pestaña "Desde Excel"
          "filas_modelo_base": [...],  # Para pestaña "Desde Modelo"
          "filas_fusionadas": [...]    # Para Excel exportado
        }
      }
    """
    datos_comparacion = {}

    cm_to_planilla = {}
    for k, v in codigos_planillas.items():
        try:
            if isinstance(v, str) and len(v) == 4 and v.startswith("CM"):
                cm_to_planilla[v] = k
            elif isinstance(k, str) and len(k) == 4 and k.startswith("CM"):
                cm_to_planilla[k] = v
        except Exception:
            log_exc("construir_tabla_comparativa armando cm_to_planilla")

    log("construir_tabla_comparativa: cm_to_planilla -> {}".format(cm_to_planilla))

    for planilla, bloque_excel in datos_excel_planilla.items():
        codigo_cm = bloque_excel['codigo_cm']
        headers = bloque_excel['headers']
        filas_excel = bloque_excel['filas']
        filas_modelo = datos_modelo_cm.get(codigo_cm, [])

        filas_excel_base, filas_modelo_base, filas_fusionadas = _comparar_por_planilla(
            codigo_cm, headers, filas_excel, filas_modelo
        )

        datos_comparacion[planilla] = {
            "codigo_cm": codigo_cm,
            "headers": headers,
            "filas_excel_base": filas_excel_base,
            "filas_modelo_base": filas_modelo_base,
            "filas_fusionadas": filas_fusionadas,
        }

    for codigo_cm, lista_modelo in datos_modelo_cm.items():
        ya_existe = any(
            info["codigo_cm"] == codigo_cm for info in datos_comparacion.values()
        )
        if ya_existe:
            continue

        nombre_planilla = cm_to_planilla.get(codigo_cm, codigo_cm)

        headers = cargar_headers_planilla(nombre_planilla, codigo_cm)
        if not headers:
            if lista_modelo:
                keys = sorted(set(k for fila in lista_modelo for k in fila.keys()))
                keys = [k for k in keys if k != HEADER_VINCULO]
                if "CodIntBIM" in keys:
                    headers = ["CodIntBIM"] + [k for k in keys if k != "CodIntBIM"]
                else:
                    headers = ["CodIntBIM"] + keys
            else:
                headers = ["CodIntBIM"]
            headers = _normalizar_headers(headers)

        filas_excel = []
        filas_modelo = lista_modelo

        filas_excel_base, filas_modelo_base, filas_fusionadas = _comparar_por_planilla(
            codigo_cm, headers, filas_excel, filas_modelo
        )

        datos_comparacion[nombre_planilla] = {
            "codigo_cm": codigo_cm,
            "headers": headers,
            "filas_excel_base": filas_excel_base,
            "filas_modelo_base": filas_modelo_base,
            "filas_fusionadas": filas_fusionadas,
        }

    log("construir_tabla_comparativa: {} planillas en datos_comparacion".format(len(datos_comparacion)))
    return datos_comparacion


def _comparar_por_planilla(codigo_cm, headers, filas_excel, filas_modelo):
    """
    Construye tres vistas:
      - filas_tabla_excel   -> solo Excel, estructurado para pestaña "Desde Excel"
      - filas_tabla_modelo  -> solo Modelo, para pestaña "Desde Modelo"
      - filas_fusionadas    -> una fila por código, con valores fusionados
                               y estados_por_celda, para exportar una sola tabla.
    """
    mapa_excel = {}
    for fila in filas_excel:
        c = (fila["CodIntBIM"] or '').strip()
        if c:
            mapa_excel.setdefault(c, []).append(fila)

    mapa_modelo = {}
    for row_m in filas_modelo:
        c = (row_m.get("CodIntBIM") or '').strip()
        if c:
            mapa_modelo.setdefault(c, []).append(row_m)

    filas_tabla_excel = []
    filas_tabla_modelo = []
    filas_fusionadas = []

    codigos_union = sorted(set(mapa_excel.keys()) | set(mapa_modelo.keys()))

    for cod in codigos_union:
        lista_excel = mapa_excel.get(cod, [])
        lista_modelo = mapa_modelo.get(cod, [])

        # --- Vista fusionada: solo una fila por CodIntBIM (tomas la primera) ---
        fila_x0 = lista_excel[0] if lista_excel else None
        fila_m0 = lista_modelo[0] if lista_modelo else None
        valores_fusion = []
        estados_fusion = []

        for h in headers:
            vx_raw = ''
            vm_raw = ''
            if fila_x0:
                idx_h = headers.index(h)
                vals = fila_x0["valores"]
                if idx_h < len(vals):
                    vx_raw = vals[idx_h]
            if fila_m0:
                vm_raw = _get_val(fila_m0, h)

            vx = _norm_val(vx_raw)
            vm = _norm_val(vm_raw)

            if not vx and not vm:
                valores_fusion.append('')
                estados_fusion.append('no_existe')
            elif vx == vm:
                valores_fusion.append(vx_raw)
                estados_fusion.append('ok')
            else:
                if vx and not vm:
                    valores_fusion.append(vx_raw)
                    estados_fusion.append('falta_modelo')
                elif not vx and vm:
                    valores_fusion.append(vm_raw)
                    estados_fusion.append('falta_excel')
                else:
                    texto = u"Planilla: {}\nModelo: {}".format(vx_raw, vm_raw)
                    valores_fusion.append(texto)
                    estados_fusion.append('difiere')

        filas_fusionadas.append({
            "CodIntBIM": cod,
            "valores": valores_fusion,
            "estado_por_celda": estados_fusion
        })

        # --- Vista Desde Excel (igual que antes, pero con "-" tratado como vacío) ---
        if not lista_excel:
            continue  # no genera filas en "Desde Excel" si no hay planilla

        for fila_x in lista_excel:
            valores_excel = fila_x["valores"]
            valores = []
            estados = []
            fila_m = lista_modelo[0] if lista_modelo else None

            for idx, h in enumerate(headers):
                vx_raw = valores_excel[idx] if idx < len(valores_excel) else ''
                vx = _norm_val(vx_raw)
                if fila_m:
                    vm_raw = _get_val(fila_m, h)
                    vm = _norm_val(vm_raw)
                else:
                    vm_raw = ''
                    vm = ''

                if not vx and not vm:
                    valores.append('')
                    estados.append('no_existe')
                elif vx == vm:
                    valores.append(vx_raw)
                    estados.append('ok')
                else:
                    if vx and not vm:
                        valores.append(vx_raw)
                        estados.append('falta_modelo')
                    elif not vx and vm:
                        texto = u"Planilla: {}\nModelo: {}".format(vx_raw, vm_raw)
                        valores.append(texto)
                        estados.append('falta_excel')
                    else:
                        texto = u"Planilla: {}\nModelo: {}".format(vx_raw, vm_raw)
                        valores.append(texto)
                        estados.append('difiere')

            filas_tabla_excel.append({
                "origen": "excel",
                "CodIntBIM": cod,
                "valores": valores,
                "estado_por_celda": estados
            })

        # --- Vista Desde Modelo (igual que antes) ---
        if not lista_modelo:
            continue

        for fila_m in lista_modelo:
            valores = []
            estados = []
            fila_x = lista_excel[0] if lista_excel else None

            for idx, h in enumerate(headers):
                vm_raw = _get_val(fila_m, h)
                vm = _norm_val(vm_raw)
                if fila_x:
                    vals = fila_x["valores"]
                    vx_raw = vals[idx] if idx < len(vals) else ''
                    vx = _norm_val(vx_raw)
                else:
                    vx_raw = ''
                    vx = ''

                if not vx and not vm:
                    valores.append('')
                    estados.append('no_existe')
                elif vx == vm:
                    valores.append(vm_raw)
                    estados.append('ok')
                else:
                    if vm and not vx:
                        valores.append(vm_raw)
                        estados.append('falta_excel')
                    elif not vm and vx:
                        texto = u"Modelo: {}\nPlanilla: {}".format(vm_raw, vx_raw)
                        valores.append(texto)
                        estados.append('falta_modelo')
                    else:
                        texto = u"Modelo: {}\nPlanilla: {}".format(vm_raw, vx_raw)
                        valores.append(texto)
                        estados.append('difiere')

            filas_tabla_modelo.append({
                "origen": "modelo",
                "CodIntBIM": cod,
                "valores": valores,
                "estado_por_celda": estados
            })

    return filas_tabla_excel, filas_tabla_modelo, filas_fusionadas


# ============================================================================
#                    EXPORTAR A EXCEL
# ============================================================================

def exportar_json_y_formatear(datos_comparacion, data_dir,
                              formatear_script, ruta_xlsx_salida,
                              python_exe):
    """
    Exporta headers, filas_fusionadas (modelo + planilla) y estado_por_celda
    para que formatear_tablas_excel.py pinte cada celda según la leyenda.
    """
    try:
        datos_por_tabla = {}
        codigos = []
        nombres = []

        for nombre_planilla, bloque in datos_comparacion.items():
            codigo_cm = bloque['codigo_cm']
            headers = bloque['headers']
            filas_fusionadas = bloque.get('filas_fusionadas', [])

            datos_por_tabla[codigo_cm] = {
                'headers': headers,
                'filas': filas_fusionadas
            }
            codigos.append(codigo_cm)
            nombres.append(nombre_planilla)

        pares = list(zip(codigos, nombres))
        pares.sort(key=lambda x: x[0])

        codigos_ordenados = [p[0] for p in pares]
        nombres_ordenados = [p[1] for p in pares]

        datos_export = {
            "listado_tablas": {
                "valores": codigos_ordenados,
                "claves": nombres_ordenados
            },
            "datos_por_tabla": datos_por_tabla
        }

        json_path = os.path.join(data_dir, "comparacion.json")
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(datos_export, f, ensure_ascii=False, indent=2)

        log("exportar_json_y_formatear: comparacion.json -> {}".format(json_path))

        CREATE_NO_WINDOW = 0x08000000
        subprocess.check_call(
            [python_exe, formatear_script, json_path, ruta_xlsx_salida],
            stderr=subprocess.STDOUT,
            creationflags=CREATE_NO_WINDOW
        )
        log("exportar_json_y_formatear: Excel generado OK -> {}".format(ruta_xlsx_salida))
    except subprocess.CalledProcessError as e:
        log_exc("exportar_json_y_formatear CalledProcessError")
        messagebox.showerror(
            "Error",
            "Error en formatear_tablas_excel.py:\n{}".format(e)
        )
        raise
    except Exception:
        log_exc("exportar_json_y_formatear error general")
        messagebox.showerror(
            "Error",
            "Error exportando a Excel.\nRevisa ui_comparacion_log.txt."
        )
        raise


# ============================================================================
#                          INTERFAZ UI
# ============================================================================

def mostrar_ui(datos_comparacion, on_exportar):
    root = tk.Tk()
    root.title("Planilla vs Modelo")
    root.geometry("1200x650")

    export_realizado = {"value": False}

    style = ttk.Style(root)
    style.configure("Treeview", rowheight=30)
    style.configure("Treeview.Heading", font=("TkDefaultFont", 9, "bold"))

    frame_main = ttk.Frame(root)
    frame_main.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

    frame_top = ttk.Frame(frame_main)
    frame_top.pack(fill=tk.X, padx=5, pady=5)

    ttk.Label(frame_top, text="Tabla:").pack(side=tk.LEFT, padx=(0, 5))

    nombres_disponibles = sorted(datos_comparacion.keys())
    selected_nombre = tk.StringVar()

    combo_tablas = ttk.Combobox(
        frame_top,
        textvariable=selected_nombre,
        values=nombres_disponibles,
        state="readonly",
        width=50
    )
    combo_tablas.pack(side=tk.LEFT, fill=tk.X, expand=True)

    lbl_titulo = ttk.Label(frame_main, text="Tabla comparativa", font=("TkDefaultFont", 10, "bold"))
    lbl_titulo.pack(anchor=tk.W, padx=5)

    frame_center = ttk.Frame(frame_main)
    frame_center.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

    notebook = ttk.Notebook(frame_center)
    notebook.pack(fill=tk.BOTH, expand=True)

    frame_excel = ttk.Frame(notebook)
    frame_modelo = ttk.Frame(notebook)
    notebook.add(frame_excel, text="Desde Excel")
    notebook.add(frame_modelo, text="Desde Modelo")

    tree_excel = crear_treeview_tabla(frame_excel)
    tree_modelo = crear_treeview_tabla(frame_modelo)

    frame_bottom = ttk.Frame(root)
    frame_bottom.pack(fill=tk.X, padx=5, pady=5, side=tk.BOTTOM)

    frame_leyenda = ttk.Frame(frame_bottom)
    frame_leyenda.pack(side=tk.LEFT, padx=5)

    ttk.Label(frame_leyenda, text="Leyenda:").pack(side=tk.LEFT, padx=(0, 5))

    for estado, color in COLOR_ESTADOS.items():
        if estado == 'ok':
            texto = 'OK / Coincide'
        elif estado == 'falta_modelo':
            texto = 'Falta en Modelo'
        elif estado == 'falta_excel':
            texto = 'Falta en Excel'
        elif estado == 'no_existe':
            texto = 'No existe'
        else:
            texto = 'Diferencia de valor'

        lbl_color = tk.Label(frame_leyenda, width=2, background=color, relief=tk.SUNKEN, bd=1)
        lbl_color.pack(side=tk.LEFT, padx=2)

        lbl_texto = ttk.Label(frame_leyenda, text=texto, font=("TkDefaultFont", 8))
        lbl_texto.pack(side=tk.LEFT, padx=2)

    frame_botones = ttk.Frame(frame_bottom)
    frame_botones.pack(side=tk.RIGHT)

    def actualizar_tablas(nombre):
        bloque = datos_comparacion.get(nombre)
        if not bloque:
            limpiar_treeview(tree_excel)
            limpiar_treeview(tree_modelo)
            return

        headers = bloque.get('headers', [])
        filas_excel = bloque.get('filas_excel_base', [])
        filas_modelo = bloque.get('filas_modelo_base', [])

        poblar_treeview(tree_excel, headers, filas_excel)
        poblar_treeview(tree_modelo, headers, filas_modelo)

    def on_sel_planilla(event=None):
        nombre = selected_nombre.get()
        if nombre:
            actualizar_tablas(nombre)

    combo_tablas.bind("<<ComboboxSelected>>", on_sel_planilla)

    def cmd_exportar():
        on_exportar()
        export_realizado["value"] = True
        messagebox.showinfo(
            "Exportar",
            "Se generará el archivo .xlsx con colores por celda."
        )

    def cmd_salir():
        if not export_realizado["value"]:
            if not messagebox.askyesno(
                "Salir",
                "No has exportado el archivo.\n¿Seguro que quieres salir sin exportar?"
            ):
                return
        root.destroy()

    btn_exportar = ttk.Button(frame_botones, text="Exportar", command=cmd_exportar)
    btn_exportar.pack(side=tk.LEFT, padx=5)

    btn_salir = ttk.Button(frame_botones, text="Salir", command=cmd_salir)
    btn_salir.pack(side=tk.LEFT, padx=5)

    sizegrip = ttk.Sizegrip(root)
    sizegrip.pack(side=tk.BOTTOM, anchor=tk.SE)

    if nombres_disponibles:
        selected_nombre.set(nombres_disponibles[0])
        actualizar_tablas(nombres_disponibles[0])

    root.protocol("WM_DELETE_WINDOW", cmd_salir)
    root.mainloop()


def crear_treeview_tabla(parent):
    frame = ttk.Frame(parent)
    frame.pack(fill=tk.BOTH, expand=True)

    tree = ttk.Treeview(frame, show='headings')
    tree.grid(row=0, column=0, sticky="nsew")

    vsb = ttk.Scrollbar(frame, orient='vertical', command=tree.yview)
    vsb.grid(row=0, column=1, sticky="ns")

    hsb = ttk.Scrollbar(frame, orient='horizontal', command=tree.xview)
    hsb.grid(row=1, column=0, sticky="ew")

    frame.rowconfigure(0, weight=1)
    frame.columnconfigure(0, weight=1)

    tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

    for estado, color in COLOR_ESTADOS.items():
        tree.tag_configure(estado, background=color, foreground='black')

    return tree


def limpiar_treeview(tree):
    for item in tree.get_children():
        tree.delete(item)
    tree["columns"] = ()


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
        if "difiere" in estados:
            return "difiere"
        if "falta_modelo" in estados:
            return "falta_modelo"
        if "falta_excel" in estados:
            return "falta_excel"
        if "no_existe" in estados:
            return "no_existe"
        return "ok"

    for fila in filas_tabla:
        valores = list(fila.get('valores', []))

        if len(valores) < len(columnas):
            valores = valores + [""] * (len(columnas) - len(valores))
        else:
            valores = valores[:len(columnas)]

        estados = fila.get('estado_por_celda') or []
        if len(estados) < len(columnas):
            estados = estados + ['ok'] * (len(columnas) - len(estados))
        else:
            estados = estados[:len(columnas)]

        tag_general = peor_estado(estados)
        tree.insert('', 'end', values=valores, tags=(tag_general,))


# ============================================================================
#                          PROGRAMA PRINCIPAL
# ============================================================================

if __name__ == '__main__':
    try:
        log("==== Inicio ui_comparacion.py ====")

        if len(sys.argv) < 9:
            msg = (
                "Uso: ui_comparacion.py "
                "<script_json> <csv_codigos> <data_dir> "
                "<formatear_script> <ruta_xlsx_salida> "
                "<python_exe> <modelo_json> <headers_json_old>"
            )
            log("main: argumentos insuficientes -> {}".format(sys.argv))
            print(msg, file=sys.stderr)
            sys.exit(120)

        script_json = sys.argv[1]
        csv_codigos = sys.argv[2]
        data_dir = sys.argv[3]
        formatear_script = sys.argv[4]
        ruta_xlsx_salida = sys.argv[5]
        python_exe = sys.argv[6]
        modelo_json = sys.argv[7]
        headers_json_old = sys.argv[8]

        log("Argumentos:\nscript_json={}\ncsv_codigos={}\ndata_dir={}\nformatear_script={}\nruta_xlsx_salida={}\npython_exe={}\nmodelo_json={}\nheaders_json_old={}".format(
            script_json, csv_codigos, data_dir, formatear_script,
            ruta_xlsx_salida, python_exe, modelo_json, headers_json_old
        ))

        cfg = leer_script_json(script_json)
        codigos_planillas = cfg.get('codigos_planillas', {})

        filas_csv = leer_csv_codigos(csv_codigos)
        datos_excel_planilla = construir_excel_por_planilla(
            filas_csv,
            codigos_planillas
        )

        datos_modelo_cm_raw = leer_modelo_por_cm(modelo_json)
        repo_activo = cargar_repo_activo()
        datos_modelo_cm = enriquecer_modelo_con_bd(datos_modelo_cm_raw, repo_activo)

        datos_comparacion = construir_tabla_comparativa(
            datos_excel_planilla,
            datos_modelo_cm,
            codigos_planillas
        )
        log("main: datos_comparacion con {} planillas".format(len(datos_comparacion)))

        def on_export():
            exportar_json_y_formatear(
                datos_comparacion,
                data_dir,
                formatear_script,
                ruta_xlsx_salida,
                python_exe
            )

        mostrar_ui(datos_comparacion, on_export)
        log("==== Fin ui_comparacion.py ====")
    except SystemExit as se:
        log("SystemExit -> {}".format(se))
        raise
    except Exception:
        log_exc("ui_comparacion main error general")
        messagebox.showerror(
            "Error",
            "Error inesperado en ui_comparacion.py.\nRevisa ui_comparacion_log.txt."
        )
        sys.exit(120)
