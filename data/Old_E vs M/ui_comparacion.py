# -*- coding: utf-8 -*-

import sys
import os
import csv
import json
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox

COLOR_ESTADOS = {
    'ok': '#C6EFCE',
    'falta_modelo': '#FFC7CE',
    'falta_excel': '#FFEB9C',
    'difiere': '#F4B084',
}

NOMBRE_PLANILLA_PILOTO = "carga_masiva_barras"
CODIGO_CM_PILOTO = "CM05"

DATA_DIR_EXT = os.path.join(
    os.path.expanduser("~"),
    r"AppData\\Roaming\\MyPyRevitExtention\\PyRevitIT.extension\\data"
)

PLANILLAS_HEADERS_JSON = os.path.join(
    DATA_DIR_EXT,
    "planillas_headers_order.json"
)


def leer_script_json(path):
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)


def leer_csv_codigos(csv_codigos):
    filas = []
    with open(csv_codigos, 'r', encoding='utf-8') as f:
        reader = csv.reader(f, delimiter=';')
        for row in reader:
            if not row:
                continue
            if not any((c or '').strip() for c in row):
                continue
            filas.append(row)
    return filas


def leer_modelo_por_cm(path_json_modelo):
    if not os.path.exists(path_json_modelo):
        return {}
    with open(path_json_modelo, 'r', encoding='utf-8') as f:
        data = json.load(f)
    result = {}
    for cm, lista in data.items():
        if isinstance(lista, list):
            result[cm] = lista
    return result


def cargar_headers_planilla(nombre_planilla, codigo_cm):
    """Devuelve lista de encabezados reales desde planillas_headers_order.json."""
    if not os.path.exists(PLANILLAS_HEADERS_JSON):
        return None

    with open(PLANILLAS_HEADERS_JSON, "r", encoding="utf-8") as f:
        data = json.load(f)

    clave = "{}::{}".format(nombre_planilla, codigo_cm)
    headers = data.get(clave)
    if headers:
        return list(headers)

    headers = data.get(codigo_cm)
    if headers:
        return list(headers)

    return None


def construir_excel_por_planilla(filas_csv, codigos_planillas):
    """
    Agrupa filas del CSV por planilla (según codigos_planillas) y
    construye:
      { nombre_planilla: {codigo_cm, headers, filas:[{CodIntBIM,valores}]} }
    """
    planilla_a_cm = {}
    for k, v in codigos_planillas.items():
        if len(k) == 4 and k.startswith("CM"):
            planilla_a_cm[v] = k
        elif isinstance(v, str) and len(v) == 4 and v.startswith("CM"):
            planilla_a_cm[k] = v

    datos_excel_planilla = {}

    for row in filas_csv:
        codint = (row[0] or '').strip()
        if not codint or len(codint) < 4:
            continue

        pref = codint[:4]  # CMxx

        nombre_planilla = None
        for np, cm in planilla_a_cm.items():
            if cm == pref:
                nombre_planilla = np
                break

        if not nombre_planilla:
            continue

        headers = cargar_headers_planilla(nombre_planilla, pref)
        if not headers:
            num_cols = len(row)
            headers = ["CodIntBIM"] + [
                "Pa{}".format(i) for i in range(2, num_cols + 1)
            ]

        if "CodIntBIM" in headers:
            headers = ["CodIntBIM"] + [h for h in headers if h != "CodIntBIM"]

        bloque = datos_excel_planilla.get(nombre_planilla)
        if not bloque:
            bloque = {
                "codigo_cm": pref,
                "headers": headers,
                "filas": []
            }
            datos_excel_planilla[nombre_planilla] = bloque

        valores = []
        for idx, h in enumerate(headers):
            if h == "CodIntBIM":
                valores.append(codint)
            else:
                val = row[idx] if idx < len(row) else ''
                valores.append(val)

        bloque["filas"].append({
            "CodIntBIM": codint,
            "valores": valores
        })

    return datos_excel_planilla


def _get_val(row_dict, header_name):
    return (row_dict.get(header_name) or '').strip()


def construir_tabla_comparativa(datos_excel_planilla, datos_modelo_cm):
    """
    Compara planilla vs modelo por CodIntBIM y por header.
    - Si coinciden: valor = dato (planilla), estado 'ok'.
    - Si difieren: valor = 'Planilla: X\\nModelo: Y', estado 'difiere'.
    - Si falta en modelo: valor = dato de planilla, estado 'falta_modelo'.
    - Si falta en Excel: valor = dato de modelo, estado 'falta_excel'.
    """
    datos_comparacion = {}

    for planilla, bloque_excel in datos_excel_planilla.items():
        codigo_cm = bloque_excel['codigo_cm']
        headers = bloque_excel['headers']
        filas_excel = bloque_excel['filas']
        filas_modelo = datos_modelo_cm.get(codigo_cm, [])

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

        filas_tabla = []

        for cod, lista_excel in mapa_excel.items():
            lista_modelo = mapa_modelo.get(cod, [])

            if not lista_modelo:
                for fila_x in lista_excel:
                    valores_excel = fila_x["valores"]
                    valores = []
                    estados = []
                    for idx, _ in enumerate(headers):
                        vx = (valores_excel[idx] if idx < len(valores_excel) else '').strip()
                        valores.append(vx)
                        estados.append('falta_modelo')
                    filas_tabla.append({
                        "origen": "solo_excel",
                        "CodIntBIM": cod,
                        "valores": valores,
                        "estado_por_celda": estados
                    })
            else:
                max_len = max(len(lista_excel), len(lista_modelo))
                for i in range(max_len):
                    fila_x = lista_excel[i] if i < len(lista_excel) else None
                    fila_m = lista_modelo[i] if i < len(lista_modelo) else None

                    if fila_x and fila_m:
                        valores_excel = fila_x["valores"]
                        valores = []
                        estados = []
                        for idx, h in enumerate(headers):
                            vx = (valores_excel[idx] if idx < len(valores_excel) else '').strip()
                            vm = _get_val(fila_m, h)

                            if not vx and not vm:
                                valores.append('')
                                estados.append('ok')
                            elif vx == vm:
                                valores.append(vx)
                                estados.append('ok')
                            else:
                                if vx and not vm:
                                    # solo planilla
                                    valores.append(vx)
                                    estados.append('falta_modelo')
                                elif not vx and vm:
                                    # solo modelo
                                    valores.append(vm)
                                    estados.append('falta_excel')
                                else:
                                    # diferencia de valor: mostrar ambos
                                    texto = u"Planilla: {}\nModelo: {}".format(vx, vm)
                                    valores.append(texto)
                                    estados.append('difiere')

                        filas_tabla.append({
                            "origen": "ambos",
                            "CodIntBIM": cod,
                            "valores": valores,
                            "estado_por_celda": estados
                        })

                    elif fila_x and not fila_m:
                        valores_excel = fila_x["valores"]
                        valores = []
                        estados = []
                        for idx, _ in enumerate(headers):
                            vx = (valores_excel[idx] if idx < len(valores_excel) else '').strip()
                            valores.append(vx)
                            estados.append('falta_modelo')
                        filas_tabla.append({
                            "origen": "solo_excel",
                            "CodIntBIM": cod,
                            "valores": valores,
                            "estado_por_celda": estados
                        })

                    elif fila_m and not fila_x:
                        valores = []
                        estados = []
                        for h in headers:
                            vm = _get_val(fila_m, h)
                            valores.append(vm)
                            estados.append('falta_excel')
                        filas_tabla.append({
                            "origen": "solo_modelo",
                            "CodIntBIM": cod,
                            "valores": valores,
                            "estado_por_celda": estados
                        })

        for cod, lista_modelo in mapa_modelo.items():
            if cod in mapa_excel:
                continue
            for fila_m in lista_modelo:
                valores = []
                estados = []
                for h in headers:
                    vm = _get_val(fila_m, h)
                    valores.append(vm)
                    estados.append('falta_excel')
                filas_tabla.append({
                    "origen": "solo_modelo",
                    "CodIntBIM": cod,
                    "valores": valores,
                    "estado_por_celda": estados
                })

        datos_comparacion[planilla] = {
            "codigo_cm": codigo_cm,
            "headers": headers,
            "filas": filas_tabla
        }

    return datos_comparacion


def exportar_json_y_formatear(datos_comparacion, data_dir,
                              formatear_script, ruta_xlsx_salida,
                              python_exe):
    datos_por_tabla = {}
    codigos = []
    nombres = []

    for nombre_planilla, bloque in datos_comparacion.items():
        codigo_cm = bloque['codigo_cm']
        datos_por_tabla[codigo_cm] = {
            'headers': bloque['headers'],
            'filas': bloque['filas']
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

    CREATE_NO_WINDOW = 0x08000000
    subprocess.check_call(
        [python_exe, formatear_script, json_path, ruta_xlsx_salida],
        stderr=subprocess.STDOUT,
        creationflags=CREATE_NO_WINDOW
    )


# ---------------- UI -----------------


def mostrar_ui(datos_comparacion, on_exportar):
    root = tk.Tk()
    root.title("Planilla vs Modelo")
    root.geometry("1200x650")

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

    lbl_titulo = ttk.Label(frame_main, text="Tabla comparativa")
    lbl_titulo.pack(anchor=tk.W, padx=5)

    frame_center = ttk.Frame(frame_main)
    frame_center.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

    tree = crear_treeview_tabla(frame_center)

    frame_bottom = ttk.Frame(root)
    frame_bottom.pack(fill=tk.X, padx=5, pady=5, side=tk.BOTTOM)

    frame_leyenda = ttk.Frame(frame_bottom)
    frame_leyenda.pack(side=tk.LEFT, padx=5)

    for estado, color in COLOR_ESTADOS.items():
        if estado == 'ok':
            texto = 'OK / Coincide'
        elif estado == 'falta_modelo':
            texto = 'Falta en modelo'
        elif estado == 'falta_excel':
            texto = 'Falta en Excel'
        else:
            texto = 'Diferencia de valor'

        lbl_color = tk.Label(frame_leyenda, width=2, background=color)
        lbl_color.pack(side=tk.LEFT, padx=2)

        lbl_texto = ttk.Label(frame_leyenda, text=texto)
        lbl_texto.pack(side=tk.LEFT, padx=4)

    frame_botones = ttk.Frame(frame_bottom)
    frame_botones.pack(side=tk.RIGHT)

    def actualizar_tabla(nombre):
        bloque = datos_comparacion.get(nombre)
        if not bloque:
            limpiar_treeview(tree)
            return
        headers = bloque.get('headers', [])
        filas = bloque.get('filas', [])
        poblar_treeview(tree, headers, filas)

    def on_sel_planilla(event=None):
        nombre = selected_nombre.get()
        if nombre:
            actualizar_tabla(nombre)

    combo_tablas.bind("<<ComboboxSelected>>", on_sel_planilla)

    def cmd_exportar():
        on_exportar()
        messagebox.showinfo(
            "Exportar",
            "Se generará el archivo .xlsx con colores."
        )
        root.destroy()

    def cmd_salir():
        root.destroy()

    btn_exportar = ttk.Button(frame_botones, text="Exportar", command=cmd_exportar)
    btn_exportar.pack(side=tk.LEFT, padx=5)

    btn_salir = ttk.Button(frame_botones, text="Salir", command=cmd_salir)
    btn_salir.pack(side=tk.LEFT, padx=5)

    sizegrip = ttk.Sizegrip(root)
    sizegrip.pack(side=tk.BOTTOM, anchor=tk.SE)

    if nombres_disponibles:
        selected_nombre.set(nombres_disponibles[0])
        actualizar_tabla(nombres_disponibles[0])

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
        tree.tag_configure(estado, background=color)
    tree.tag_configure('fila_header', background='#D9D9D9')

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
    if "CodIntBIM" not in columnas:
        columnas.insert(0, "CodIntBIM")

    tree["columns"] = columnas

    for h in columnas:
        tree.heading(h, text=h)
        tree.column(h, width=160, stretch=False, anchor="w")

    for fila in filas_tabla:
        valores = fila['valores']
        if len(valores) == len(headers):
            pass
        elif len(valores) + 1 == len(columnas) and "CodIntBIM" in columnas:
            cod = fila.get("CodIntBIM", "")
            valores = [cod] + valores

        estados = fila['estado_por_celda']
        if 'difiere' in estados:
            tag = 'difiere'
        elif 'falta_modelo' in estados:
            tag = 'falta_modelo'
        elif 'falta_excel' in estados:
            tag = 'falta_excel'
        else:
            tag = 'ok'

        tree.insert('', 'end', values=valores, tags=(tag,))


if __name__ == '__main__':
    if len(sys.argv) < 9:
        print(
            "Uso: ui_comparacion.py "
            "<script_json> <csv_codigos> <data_dir> "
            "<formatear_script> <ruta_xlsx_salida> "
            "<python_exe> <modelo_json> <headers_json_old>",
            file=sys.stderr
        )
        sys.exit(1)

    script_json = sys.argv[1]
    csv_codigos = sys.argv[2]
    data_dir = sys.argv[3]
    formatear_script = sys.argv[4]
    ruta_xlsx_salida = sys.argv[5]
    python_exe = sys.argv[6]
    modelo_json = sys.argv[7]
    headers_json_old = sys.argv[8]

    cfg = leer_script_json(script_json)
    codigos_planillas = cfg.get('codigos_planillas', {})

    filas_csv = leer_csv_codigos(csv_codigos)
    datos_excel_planilla = construir_excel_por_planilla(
        filas_csv,
        codigos_planillas
    )

    datos_modelo_cm = leer_modelo_por_cm(modelo_json)
    datos_comparacion = construir_tabla_comparativa(
        datos_excel_planilla,
        datos_modelo_cm
    )

    def on_export():
        exportar_json_y_formatear(
            datos_comparacion,
            data_dir,
            formatear_script,
            ruta_xlsx_salida,
            python_exe
        )

    mostrar_ui(datos_comparacion, on_export)
