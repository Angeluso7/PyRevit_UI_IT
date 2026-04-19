# -*- coding: utf-8 -*-
"""
Visor externo CPython — Comprobación de Códigos CodIntBIM.
Uso: pythonw.exe comprobacion_codigos.py <ruta_json>

JSON esperado:
{
  "registros": [
    {"CodIntBIM": "CM33CP-TMKA1-0003", "Situacion": "Código encontrado"},
    {"CodIntBIM": "CM13SE-TPIA1-0005", "Situacion": "Código no encontrado"}
  ]
}
"""

import sys
import os
import json
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# ── Paleta estándar del proyecto ──────────────────────────────────────────────
BG        = "#1E1E2E"
SURFACE   = "#2A2A3E"
SURFACE2  = "#32324A"
ACCENT    = "#4F7EFF"
TEXT      = "#E0E0F0"
TEXT_MUTED= "#8888AA"
ROW_ODD   = "#252538"
ROW_EVEN  = "#2A2A3E"
BTN_BG    = "#3A3A5A"
BTN_HOVER = "#4F7EFF"
ERROR_COL = "#FF6B6B"
OK_COL    = "#6BFF9E"
# ─────────────────────────────────────────────────────────────────────────────


def cargar_registros(json_path):
    try:
        if not os.path.exists(json_path):
            messagebox.showerror("Error",
                "No se encontró el archivo JSON:\n{}".format(json_path))
            return []
        with open(json_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data.get("registros", [])
    except Exception as e:
        messagebox.showerror("Error",
            "Error leyendo JSON de comprobación:\n{}".format(e))
        return []


def _btn(parent, text, cmd, color=BTN_BG):
    b = tk.Button(
        parent, text=text, command=cmd,
        bg=color, fg=TEXT, activebackground=BTN_HOVER, activeforeground="#FFFFFF",
        relief="flat", padx=14, pady=6,
        font=("Segoe UI", 9), cursor="hand2", bd=0
    )
    b.bind("<Enter>", lambda e: b.config(bg=BTN_HOVER))
    b.bind("<Leave>", lambda e: b.config(bg=color))
    return b


def mostrar_ventana(registros):
    root = tk.Tk()
    root.title("Comprobación de Códigos — CodIntBIM")
    root.geometry("660x440")
    root.minsize(480, 320)
    root.configure(bg=BG)

    # ── Cabecera ──────────────────────────────────────────────────────────────
    hdr = tk.Frame(root, bg=SURFACE, pady=10)
    hdr.pack(fill="x")
    tk.Label(hdr, text="  Comprobación de Códigos CodIntBIM",
             bg=SURFACE, fg=TEXT,
             font=("Segoe UI", 12, "bold")).pack(side="left")

    total   = len(registros)
    ok_cnt  = sum(1 for r in registros if "no encontrado" not in r.get("Situacion","").lower())
    err_cnt = total - ok_cnt
    stats   = "  ✔ {}   ✘ {}   Total {}  ".format(ok_cnt, err_cnt, total)
    tk.Label(hdr, text=stats, bg=SURFACE, fg=TEXT_MUTED,
             font=("Segoe UI", 9)).pack(side="right")

    # ── Treeview ──────────────────────────────────────────────────────────────
    tree_frame = tk.Frame(root, bg=BG)
    tree_frame.pack(fill="both", expand=True, padx=12, pady=(10, 4))
    tree_frame.columnconfigure(0, weight=1)
    tree_frame.rowconfigure(0, weight=1)

    style = ttk.Style()
    style.theme_use("clam")
    style.configure("Dark.Treeview",
        background=SURFACE, foreground=TEXT,
        fieldbackground=SURFACE, rowheight=24,
        font=("Segoe UI", 9))
    style.configure("Dark.Treeview.Heading",
        background=ACCENT, foreground="#FFFFFF",
        font=("Segoe UI", 9, "bold"), relief="flat")
    style.map("Dark.Treeview",
        background=[("selected", ACCENT)],
        foreground=[("selected", "#FFFFFF")])

    tree = ttk.Treeview(
        tree_frame,
        columns=("CodIntBIM", "Situacion"),
        show="headings",
        style="Dark.Treeview",
        selectmode="browse"
    )
    tree.heading("CodIntBIM",  text="CodIntBIM",  anchor="w")
    tree.heading("Situacion",  text="Situación",  anchor="w")
    tree.column("CodIntBIM",   width=240, anchor="w", stretch=True)
    tree.column("Situacion",   width=360, anchor="w", stretch=True)

    tree.tag_configure("odd",  background=ROW_ODD)
    tree.tag_configure("even", background=ROW_EVEN)
    tree.tag_configure("err",  foreground=ERROR_COL)
    tree.tag_configure("ok",   foreground=OK_COL)

    for i, reg in enumerate(registros):
        cod = reg.get("CodIntBIM", "")
        sit = reg.get("Situacion", "")
        row_tag  = "odd" if i % 2 == 0 else "even"
        sit_tag  = "err" if "no encontrado" in sit.lower() else "ok"
        tree.insert("", "end", values=(cod, sit), tags=(row_tag, sit_tag))

    vsb = ttk.Scrollbar(tree_frame, orient="vertical",   command=tree.yview)
    hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
    tree.grid(row=0, column=0, sticky="nsew")
    vsb.grid(row=0, column=1, sticky="ns")
    hsb.grid(row=1, column=0, sticky="ew")

    # ── Footer ────────────────────────────────────────────────────────────────
    foot = tk.Frame(root, bg=SURFACE, pady=8)
    foot.pack(fill="x", side="bottom")

    def exportar():
        if not registros:
            messagebox.showinfo("Información", "No hay datos para exportar.")
            return
        filename = filedialog.asksaveasfilename(
            title="Exportar comprobación",
            defaultextension=".txt",
            filetypes=[("Texto", "*.txt"), ("Todos", "*.*")]
        )
        if not filename:
            return
        try:
            with open(filename, "w", encoding="utf-8") as f:
                f.write("CodIntBIM;Situación\n")
                for reg in registros:
                    f.write("{};{}\n".format(
                        reg.get("CodIntBIM", ""), reg.get("Situacion", "")))
            messagebox.showinfo("Exportación completada",
                "Archivo guardado en:\n{}".format(filename))
        except Exception as e:
            messagebox.showerror("Error", "No se pudo exportar:\n{}".format(e))

    _btn(foot, "Exportar", exportar).pack(side="right", padx=(4, 12))
    _btn(foot, "Cerrar",   root.destroy, color="#3A2A2A").pack(side="right", padx=4)

    tk.Label(foot,
        text="  {} registros cargados".format(total),
        bg=SURFACE, fg=TEXT_MUTED, font=("Segoe UI", 8)
    ).pack(side="left", padx=12)

    root.mainloop()


def main():
    if len(sys.argv) < 2:
        messagebox.showerror("Error", "Uso: comprobacion_codigos.py <ruta_json>")
        return
    registros = cargar_registros(sys.argv[1])
    if not registros:
        messagebox.showinfo("Información", "No hay registros para mostrar.")
        return
    mostrar_ventana(registros)


if __name__ == "__main__":
    main()
