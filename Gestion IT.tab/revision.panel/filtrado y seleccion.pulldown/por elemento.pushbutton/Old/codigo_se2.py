# -*- coding: utf-8 -*-
# Visor solo lectura de parámetros para un elemento (2 columnas: Parámetro / Valor)

import tkinter as tk
from tkinter import ttk, messagebox
import json
import os
import sys
import traceback


# Primer argumento: ruta del JSON temporal de elemento
if len(sys.argv) > 1:
    JSON_PATH = sys.argv[1]
else:
    messagebox.showerror("Error", "No se recibió ruta al JSON del elemento.")
    sys.exit(1)


def cargar_json(ruta):
    try:
        if not os.path.exists(ruta):
            messagebox.showinfo("Información", "Archivo no encontrado:\n{}".format(ruta))
            return None
        with open(ruta, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        messagebox.showerror(
            "Error",
            "Error cargando JSON:\n{}".format(traceback.format_exc())
        )
        return None


class ReadOnlyTreeview(ttk.Treeview):
    def __init__(self, master, row_dict, **kwargs):
        # Dos columnas fijas: Parametro / Valor
        headers = ["Parámetro", "Valor"]
        super(ReadOnlyTreeview, self).__init__(
            master,
            columns=headers,
            show="headings",
            **kwargs
        )

        # Definir encabezados
        self.heading("Parámetro", text="Parámetro")
        self.heading("Valor", text="Valor")

        self.column("Parámetro", width=250, anchor="w", stretch=True)
        self.column("Valor", width=500, anchor="w", stretch=True)

        # Insertar filas: una por cada par clave/valor
        for param_name, param_value in row_dict.items():
            self.insert("", "end", values=(param_name, param_value))

        # Bloquear edición
        self.bind("<Double-1>", lambda e: "break")
        self.bind("<Button-1>", self._on_click)

    def _on_click(self, event):
        region = self.identify("region", event.x, event.y)
        if region == "separator":
            return "break"


def main():
    data = cargar_json(JSON_PATH) or {}
    # Row es un diccionario {encabezado: valor}
    row = data.get("Row", {})
    elem_id = data.get("ElementId", "")
    archivo = data.get("Archivo", "")

    if not row:
        messagebox.showinfo("Información", "Sin datos para mostrar.")
        return

    root = tk.Tk()

    titulo = "Datos elemento {}".format(elem_id)
    if archivo:
        titulo += " | {}".format(os.path.basename(archivo))
    root.title(titulo)

    root.geometry("900x400")
    root.minsize(500, 250)

    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)

    frame = ttk.Frame(root)
    frame.grid(row=0, column=0, sticky="nsew")
    frame.columnconfigure(0, weight=1)
    frame.rowconfigure(0, weight=1)

    tree = ReadOnlyTreeview(frame, row)
    tree.grid(row=0, column=0, sticky="nsew")

    vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
    vsb.grid(row=0, column=1, sticky="ns")

    hsb = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
    hsb.grid(row=1, column=0, sticky="ew")

    tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

    btn_frame = ttk.Frame(root)
    btn_frame.grid(row=1, column=0, sticky="ew", pady=5, padx=5)
    btn_frame.columnconfigure(0, weight=1)

    btn_cerrar = ttk.Button(btn_frame, text="Cerrar", command=root.destroy)
    btn_cerrar.grid(row=0, column=0, sticky="ew", padx=5)

    root.mainloop()


if __name__ == "__main__":
    main()
