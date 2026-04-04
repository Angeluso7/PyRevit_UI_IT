# -*- coding: utf-8 -*-
"""
Ventana externa de comprobación de códigos.

Uso:
    python comprobacion_codigos.py <ruta_json>

El JSON debe tener la forma:
{
  "registros": [
    {"CodIntBIM": "CM33CP-TMKA1-0003", "Situacion": "Código encontrado"},
    {"CodIntBIM": "CM13SE-TPIA1-0005", "Situacion": "Código no encontrado"},
    ...
  ]
}
"""

import sys
import os
import json
import tkinter as tk
from tkinter import ttk, filedialog, messagebox


def cargar_registros(json_path):
    try:
        if not os.path.exists(json_path):
            messagebox.showerror(
                "Error",
                "No se encontró el archivo JSON:\n{}".format(json_path)
            )
            return []
        with open(json_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data.get("registros", [])
    except Exception as e:
        messagebox.showerror(
            "Error",
            "Error leyendo JSON de comprobación:\n{}".format(e)
        )
        return []


def mostrar_ventana(registros):
    root = tk.Tk()
    root.title("Comprobación de Códigos")
    root.geometry("600x400")
    root.minsize(400, 300)

    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)

    frame = ttk.Frame(root)
    frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
    frame.columnconfigure(0, weight=1)
    frame.rowconfigure(0, weight=1)

    tree = ttk.Treeview(
        frame,
        columns=("CodIntBIM", "Situacion"),
        show="headings",
        selectmode="browse"
    )

    tree.heading("CodIntBIM", text="CodIntBIM")
    tree.heading("Situacion", text="Situación")

    tree.column("CodIntBIM", width=220, anchor="w", stretch=True)
    tree.column("Situacion", width=320, anchor="w", stretch=True)

    for reg in registros:
        cod = reg.get("CodIntBIM", "")
        sit = reg.get("Situacion", "")
        tree.insert("", "end", values=(cod, sit))

    vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
    hsb = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

    tree.grid(row=0, column=0, sticky="nsew")
    vsb.grid(row=0, column=1, sticky="ns")
    hsb.grid(row=1, column=0, sticky="ew")

    btn_frame = ttk.Frame(root)
    btn_frame.grid(row=1, column=0, sticky="e", padx=10, pady=(0, 10))

    def exportar():
        if not registros:
            messagebox.showinfo("Información", "No hay datos para exportar.")
            return

        filename = filedialog.asksaveasfilename(
            title="Exportar comprobación de códigos",
            defaultextension=".txt",
            filetypes=[("Archivo de texto", "*.txt"), ("Todos los archivos", "*.*")]
        )
        if not filename:
            return

        try:
            with open(filename, "w", encoding="utf-8") as f:
                f.write("CodIntBIM;Situación\n")
                for reg in registros:
                    cod = reg.get("CodIntBIM", "")
                    sit = reg.get("Situacion", "")
                    linea = "{};{}\n".format(cod, sit)
                    f.write(linea)
            messagebox.showinfo(
                "Exportación completada",
                "Archivo exportado en:\n{}".format(filename)
            )
        except Exception as e:
            messagebox.showerror(
                "Error de exportación",
                "No se pudo exportar el archivo:\n{}".format(e)
            )

    def aceptar():
        root.destroy()

    btn_exportar = ttk.Button(btn_frame, text="Exportar", command=exportar)
    btn_aceptar = ttk.Button(btn_frame, text="Aceptar", command=aceptar)

    btn_exportar.grid(row=0, column=0, padx=(0, 5))
    btn_aceptar.grid(row=0, column=1)

    root.mainloop()


def main():
    if len(sys.argv) < 2:
        messagebox.showerror(
            "Error",
            "Uso: comprobacion_codigos.py <ruta_json>"
        )
        return

    json_path = sys.argv[1]
    registros = cargar_registros(json_path)
    if not registros:
        messagebox.showinfo(
            "Información",
            "No hay registros para mostrar."
        )
        return

    mostrar_ventana(registros)


if __name__ == "__main__":
    main()
