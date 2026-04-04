# -*- coding: utf-8 -*-
"""
Selector de planillas en Tkinter.
Lee script.json (codigos_planillas + reemplazos_de_nombres) y
permite elegir una planilla por alias, con filtro por texto.
Escribe un JSON meta con:
- NombrePlanillaOriginal
- NombrePlanillaAlias
- CodigoPlanilla
en la ruta pasada por argumento (planilla_meta_tmp.json).
"""

import os
import sys
import json
import traceback
import tkinter as tk
from tkinter import ttk, messagebox

DATA_DIR = os.path.join(
    os.path.expanduser("~"),
    r"AppData\Roaming\MyPyRevitExtention\PyRevitIT.extension\data"
)

SCRIPT_JSON_PATH = os.path.join(DATA_DIR, "script.json")

# Primer argumento: ruta de salida del meta
if len(sys.argv) > 1:
    PLANILLA_META_PATH = sys.argv[1]
else:
    PLANILLA_META_PATH = os.path.join(os.path.dirname(__file__), "planilla_meta_tmp.json")


def cargar_script_json():
    if not os.path.exists(SCRIPT_JSON_PATH):
        messagebox.showerror(
            "Error script.json",
            "No se encontró script.json en:\n{}".format(SCRIPT_JSON_PATH)
        )
        return {}
    try:
        with open(SCRIPT_JSON_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        messagebox.showerror(
            "Error script.json",
            "Error leyendo script.json:\n{}".format(traceback.format_exc())
        )
        return {}


def construir_lista_planillas(cfg):
    codigos    = cfg.get("codigos_planillas", {}) or {}
    reemplazos = cfg.get("reemplazos_de_nombres", {}) or {}
    items = []
    for key, codigo in codigos.items():
        alias = reemplazos.get(key, key)
        items.append((alias, key, codigo))
    items.sort(key=lambda x: (x[0] or "").lower())
    return items


def guardar_meta(nombre_orig, nombre_alias, codigo):
    try:
        data = {
            "NombrePlanillaOriginal": nombre_orig,
            "NombrePlanillaAlias":    nombre_alias,
            "CodigoPlanilla":         codigo,
        }
        folder = os.path.dirname(PLANILLA_META_PATH)
        if folder and not os.path.exists(folder):
            os.makedirs(folder)
        with open(PLANILLA_META_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
    except Exception:
        messagebox.showerror(
            "Error",
            "Error guardando meta de planilla:\n{}".format(traceback.format_exc())
        )


def main():
    cfg = cargar_script_json()
    if not cfg:
        return

    planillas = construir_lista_planillas(cfg)
    if not planillas:
        messagebox.showinfo(
            "Sin planillas",
            "No se encontraron planillas en script.json."
        )
        return

    root = tk.Tk()
    root.title("Selector de planilla")
    root.geometry("400x400")
    root.minsize(300, 300)

    filt_var = tk.StringVar()
    entry = ttk.Entry(root, textvariable=filt_var)
    entry.pack(fill="x", padx=5, pady=5)

    listbox = tk.Listbox(root, selectmode="single")
    listbox.pack(fill="both", expand=True, padx=5, pady=5)

    current_items = list(planillas)

    def refrescar_lista(*args):
        texto = filt_var.get().strip().lower()
        listbox.delete(0, tk.END)
        del current_items[:]
        for alias, nombre_orig, codigo in planillas:
            if not texto or texto in (alias or "").lower():
                idx = len(current_items)
                current_items.append((alias, nombre_orig, codigo))
                listbox.insert(tk.END, alias)

    filt_var.trace_add("write", refrescar_lista)

    def aceptar():
        sel = listbox.curselection()
        if not sel:
            messagebox.showinfo("Información", "Seleccione una planilla.")
            return
        idx = sel[0]
        alias, nombre_orig, codigo = current_items[idx]
        if not codigo:
            messagebox.showerror(
                "Error",
                "La planilla seleccionada no tiene código en script.json."
            )
            return
        guardar_meta(nombre_orig, alias, codigo)
        root.destroy()

    btn = ttk.Button(root, text="Aceptar", command=aceptar)
    btn.pack(pady=5)

    refrescar_lista()
    entry.focus()
    root.mainloop()


if __name__ == "__main__":
    main()
