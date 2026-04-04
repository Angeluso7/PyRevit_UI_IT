# -*- coding: utf-8 -*-
"""
Visor solo lectura de parámetros por elemento (vista vertical).

Uso:
pythonw.exe visor_elemento.pyw ruta\repo_elemento_tmp.json
"""

import sys
import os
import json
import traceback
import tkinter as tk
from tkinter import ttk, messagebox


def cargar_json(path):
    try:
        if not os.path.exists(path):
            messagebox.showerror(
                "Error",
                "No se encontró el archivo JSON:\n{}".format(path)
            )
            return None
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        messagebox.showerror(
            "Error",
            "Error cargando JSON:\n{}".format(traceback.format_exc())
        )
        return None


def main():
    if len(sys.argv) < 2:
        messagebox.showerror(
            "Error",
            "No se recibió la ruta del JSON temporal."
        )
        return

    json_path = sys.argv[1]
    data = cargar_json(json_path)
    if not data:
        return

    headers = data.get("Headers", []) or []
    row = data.get("Row", {}) or {}
    elem_id = data.get("ElementId", "")
    planilla = data.get("Planilla", "")
    archivo = data.get("Archivo", "")

    root = tk.Tk()
    titulo = "Elemento {} - Planilla '{}'".format(elem_id, planilla or "")
    root.title(titulo)
    root.geometry("600x500")
    root.minsize(400, 300)

    # Frame principal
    main_frame = ttk.Frame(root)
    main_frame.pack(fill="both", expand=True, padx=5, pady=5)

    # Info superior
    info = "Archivo: {}\nElementId: {}".format(archivo, elem_id)
    lbl_info = ttk.Label(main_frame, text=info, anchor="w", justify="left")
    lbl_info.pack(fill="x", pady=(0, 5))

    # Treeview vertical: 2 columnas (Parámetro, Valor)
    columns = ("Parametro", "Valor")
    tree = ttk.Treeview(
        main_frame,
        columns=columns,
        show="headings"
    )

    tree.heading("Parametro", text="Parámetro")
    tree.heading("Valor", text="Valor")

    tree.column("Parametro", width=200, anchor="w", stretch=True)
    tree.column("Valor", width=300, anchor="w", stretch=True)

    # Insertar una fila por cada encabezado
    for h in headers:
        valor = row.get(h, "")
        tree.insert("", "end", values=(h, valor))

    tree.pack(fill="both", expand=True, side="left")

    # Scrollbars
    vsb = ttk.Scrollbar(main_frame, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=vsb.set)
    vsb.pack(fill="y", side="right")

    # Botón Cerrar
    btn_frame = ttk.Frame(root)
    btn_frame.pack(fill="x", pady=5)

    btn_close = ttk.Button(btn_frame, text="Cerrar", command=root.destroy)
    btn_close.pack(side="right", padx=5)

    root.mainloop()


if __name__ == "__main__":
    main()
