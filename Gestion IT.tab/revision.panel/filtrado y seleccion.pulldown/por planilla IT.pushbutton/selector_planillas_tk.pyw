# selector_planillas_tk.pyw
# -*- coding: utf-8 -*-

import os
import sys
import json
import tkinter as tk
from tkinter import ttk, messagebox

def cargar_planillas_desde_json(path_json):
    if not os.path.exists(path_json):
        return []
    try:
        with open(path_json, "r", encoding="utf-8") as f:
            data = json.load(f)
        return list(data.get("codigos_planillas", {}).keys())
    except Exception:
        return []

def cargar_planillas_desde_meta(path_meta):
    # Este módulo no ve Revit; aquí solo usarás lo que le pases por JSON
    if not os.path.exists(path_meta):
        return []
    try:
        with open(path_meta, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data.get("planillas_doc", [])
    except Exception:
        return []

def main():
    if len(sys.argv) < 3:
        messagebox.showerror(
            "Error",
            "Se esperaban 2 argumentos: ruta_meta_entrada, ruta_salida_seleccion."
        )
        return

    ruta_meta_entrada = sys.argv[1]
    ruta_salida = sys.argv[2]

    # meta_entrada: { "planillas_doc": [...], "ruta_json": "..." }
    if not os.path.exists(ruta_meta_entrada):
        messagebox.showerror("Error", "No se encontró archivo meta:\n{}".format(ruta_meta_entrada))
        return

    with open(ruta_meta_entrada, "r", encoding="utf-8") as f:
        meta = json.load(f)

    planillas_doc = meta.get("planillas_doc", []) or []
    ruta_json = meta.get("ruta_json", "") or ""
    planillas_json = cargar_planillas_desde_json(ruta_json)

    nombres_combinados = sorted(set(planillas_doc) | set(planillas_json))
    if not nombres_combinados:
        messagebox.showinfo("Aviso", "No se encontraron planillas.")
        return

    root = tk.Tk()
    root.title("Selecciona una planilla")
    root.geometry("400x300")
    root.minsize(350, 250)

    root.columnconfigure(0, weight=1)
    root.rowconfigure(1, weight=1)

    # Filtro
    lbl_filtro = ttk.Label(root, text="Filtrar planillas:")
    lbl_filtro.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 0))

    var_filtro = tk.StringVar()

    entry_filtro = ttk.Entry(root, textvariable=var_filtro)
    entry_filtro.grid(row=0, column=0, sticky="ew", padx=10, pady=(30, 5))

    # Listbox de planillas
    frame_lista = ttk.Frame(root)
    frame_lista.grid(row=1, column=0, sticky="nsew", padx=10, pady=5)
    frame_lista.columnconfigure(0, weight=1)
    frame_lista.rowconfigure(0, weight=1)

    listbox = tk.Listbox(frame_lista, selectmode="browse")
    listbox.grid(row=0, column=0, sticky="nsew")

    scrollbar = ttk.Scrollbar(frame_lista, orient="vertical", command=listbox.yview)
    scrollbar.grid(row=0, column=1, sticky="ns")
    listbox.config(yscrollcommand=scrollbar.set)

    # Cargar lista inicial
    def actualizar_lista(*args):
        filtro = var_filtro.get().strip().lower()
        listbox.delete(0, tk.END)
        for nombre in nombres_combinados:
            if filtro in nombre.lower():
                listbox.insert(tk.END, nombre)

    var_filtro.trace_add("write", actualizar_lista)
    actualizar_lista()

    # Botones
    frame_botones = ttk.Frame(root)
    frame_botones.grid(row=2, column=0, sticky="ew", padx=10, pady=10)
    frame_botones.columnconfigure(0, weight=1)
    frame_botones.columnconfigure(1, weight=1)

    seleccion = {"nombre": None}

    def on_aceptar():
        sel = listbox.curselection()
        if not sel:
            messagebox.showinfo("Aviso", "Por favor seleccione una planilla.")
            return
        nombre = listbox.get(sel[0])
        seleccion["nombre"] = nombre
        # Guardar en JSON de salida
        salida = {"selected_planilla": nombre}
        with open(ruta_salida, "w", encoding="utf-8") as f:
            json.dump(salida, f, indent=2, ensure_ascii=False)
        root.destroy()

    def on_cancelar():
        root.destroy()

    btn_aceptar = ttk.Button(frame_botones, text="Aceptar", command=on_aceptar)
    btn_aceptar.grid(row=0, column=0, sticky="ew", padx=5)

    btn_cancelar = ttk.Button(frame_botones, text="Cancelar", command=on_cancelar)
    btn_cancelar.grid(row=0, column=1, sticky="ew", padx=5)

    entry_filtro.focus_set()
    root.mainloop()

if __name__ == "__main__":
    main()
