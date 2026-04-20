# -*- coding: utf-8 -*-
"""
visor_elemento.pyw  v2.0
Visor solo lectura de parametros por elemento.
Vista vertical: 2 columnas (Parametro | Valor), una fila por campo.
Tema dark completo.

Uso:
    pythonw.exe visor_elemento.pyw ruta\repo_elemento_tmp.json
"""

import sys
import os
import json
import traceback
import tkinter as tk
from tkinter import ttk, messagebox

# ── Paleta dark ───────────────────────────────────────────────────────────────
C_BG      = "#1E1E2E"
C_SURFACE = "#2A2A3E"
C_HDR_BG  = "#16213E"
C_ACCENT  = "#4F7EFF"
C_TEXT    = "#E0E0F0"
C_MUTED   = "#9898B0"
C_BORDER  = "#3A3A55"
C_ODD     = "#252538"
C_EVEN    = "#2A2A3E"
C_SEL_BG  = "#3A4F8C"
C_SEL_FG  = "#FFFFFF"

FT_TITLE = ("Segoe UI", 11, "bold")
FT_INFO  = ("Segoe UI",  9)
FT_HDR   = ("Segoe UI",  9, "bold")
FT_CELL  = ("Segoe UI",  9)
FT_BTN   = ("Segoe UI",  9, "bold")

# ── Estilos dark ──────────────────────────────────────────────────────────────
def apply_dark(root):
    s = ttk.Style(root)
    s.theme_use("clam")
    root.configure(bg=C_BG)
    s.configure("TFrame",         background=C_BG)
    s.configure("Header.TFrame",  background=C_HDR_BG)
    s.configure("Title.TLabel",   background=C_HDR_BG, foreground=C_TEXT,   font=FT_TITLE)
    s.configure("Info.TLabel",    background=C_HDR_BG, foreground=C_MUTED,  font=FT_INFO)
    s.configure("TLabel",         background=C_BG,     foreground=C_TEXT,   font=FT_CELL)
    s.configure("Treeview",
                background=C_SURFACE, foreground=C_TEXT,
                fieldbackground=C_SURFACE, rowheight=24, font=FT_CELL, borderwidth=0)
    s.configure("Treeview.Heading",
                background=C_HDR_BG, foreground=C_ACCENT,
                font=FT_HDR, relief="flat")
    s.map("Treeview",
          background=[("selected", C_SEL_BG)],
          foreground=[("selected", C_SEL_FG)])
    s.map("Treeview.Heading",
          background=[("active", C_BORDER)])
    s.configure("Vertical.TScrollbar",
                background=C_BORDER, troughcolor=C_SURFACE,
                arrowcolor=C_MUTED, borderwidth=0)
    s.configure("Horizontal.TScrollbar",
                background=C_BORDER, troughcolor=C_SURFACE,
                arrowcolor=C_MUTED, borderwidth=0)
    s.configure("TButton",
                background=C_ACCENT, foreground="#FFFFFF",
                font=FT_BTN, borderwidth=0, padding=(10, 4))
    s.map("TButton",
          background=[("active", "#3A6AEF"), ("pressed", "#2A5ACF")])
    s.configure("Footer.TFrame", background=C_HDR_BG)
    s.configure("Footer.TLabel", background=C_HDR_BG, foreground=C_MUTED, font=FT_INFO)
    s.configure("Sep.TSeparator", background=C_BORDER)

# ── Utilidades ─────────────────────────────────────────────────────────────────
def cargar_json(path):
    try:
        if not os.path.exists(path):
            messagebox.showerror(
                "Error",
                "No se encontro el archivo JSON:\n{}".format(path))
            return None
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        messagebox.showerror(
            "Error",
            "Error cargando JSON:\n{}".format(traceback.format_exc()))
        return None

def _basename(ruta):
    return os.path.basename(ruta) if ruta else ""

# ── Main ────────────────────────────────────────────────────────────────────────
def main():
    if len(sys.argv) < 2:
        messagebox.showerror(
            "Error", "No se recibio la ruta del JSON temporal.")
        return

    data = cargar_json(sys.argv[1])
    if not data:
        return

    headers  = data.get("Headers",   []) or []
    row      = data.get("Row",       {}) or {}
    elem_id  = data.get("ElementId", "")
    planilla = data.get("Planilla",  "")
    archivo  = data.get("Archivo",   "")
    codint   = data.get("CodIntBIM", "")

    # ── Ventana principal ───────────────────────────────────────────────
    root = tk.Tk()
    apply_dark(root)
    root.title(u"Elemento {} — {}".format(elem_id, planilla or u"Sin planilla"))
    root.geometry("640x560")
    root.minsize(420, 320)
    root.resizable(True, True)
    root.columnconfigure(0, weight=1)
    root.rowconfigure(2, weight=1)

    # ── Cabecera ──────────────────────────────────────────────────────────────
    hdr = ttk.Frame(root, padding=(12, 8), style="Header.TFrame")
    hdr.grid(row=0, column=0, sticky="ew")

    ttk.Label(hdr,
              text=u"Planilla: {}".format(planilla or u"\u2014"),
              style="Title.TLabel"
    ).pack(anchor="w")
    ttk.Label(hdr,
              text=u"ID: {}   CodIntBIM: {}   Archivo: {}".format(
                  elem_id, codint or u"\u2014", _basename(archivo) or archivo),
              style="Info.TLabel"
    ).pack(anchor="w", pady=(2, 0))

    ttk.Separator(root, orient="horizontal").grid(row=1, column=0, sticky="ew")

    # ── Tabla vertical: Parametro | Valor ───────────────────────────────
    frame_tbl = ttk.Frame(root)
    frame_tbl.grid(row=2, column=0, sticky="nsew", padx=6, pady=(4, 0))
    frame_tbl.columnconfigure(0, weight=1)
    frame_tbl.rowconfigure(0, weight=1)

    columns = ("Parametro", "Valor")
    tree = ttk.Treeview(
        frame_tbl,
        columns=columns,
        show="headings",
        selectmode="browse"
    )

    tree.heading("Parametro", text=u"Par\xe1metro", anchor="w")
    tree.heading("Valor",     text=u"Valor",      anchor="w")
    tree.column("Parametro",  width=220, minwidth=120, anchor="w", stretch=False)
    tree.column("Valor",      width=360, minwidth=120, anchor="w", stretch=True)

    # Filas alternadas
    tree.tag_configure("odd",  background=C_ODD)
    tree.tag_configure("even", background=C_EVEN)
    tree.tag_configure("filled", foreground=C_TEXT)
    tree.tag_configure("empty", foreground=C_MUTED)

    con_valor = 0
    for idx, h in enumerate(headers):
        valor = row.get(h, "")
        tag_row = "odd" if idx % 2 == 0 else "even"
        tag_fill = "filled" if valor else "empty"
        tree.insert("", "end", values=(h, valor), tags=(tag_row, tag_fill))
        if valor:
            con_valor += 1

    vsb = ttk.Scrollbar(frame_tbl, orient="vertical",   command=tree.yview)
    hsb = ttk.Scrollbar(frame_tbl, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
    tree.grid(row=0, column=0, sticky="nsew")
    vsb.grid(row=0, column=1, sticky="ns")
    hsb.grid(row=1, column=0, sticky="ew")

    # ── Footer ───────────────────────────────────────────────────────────────
    ttk.Separator(root, orient="horizontal").grid(row=3, column=0, sticky="ew")
    footer = ttk.Frame(root, padding=(10, 5), style="Footer.TFrame")
    footer.grid(row=4, column=0, sticky="ew")
    footer.columnconfigure(0, weight=1)

    ttk.Label(
        footer,
        text=u"{} / {} par\xe1metros con valor".format(con_valor, len(headers)),
        style="Footer.TLabel"
    ).grid(row=0, column=0, sticky="w")

    ttk.Button(
        footer, text="Cerrar", command=root.destroy
    ).grid(row=0, column=1, sticky="e", padx=(10, 0))

    ttk.Sizegrip(root).grid(row=5, column=0, sticky="se")
    root.mainloop()


if __name__ == "__main__":
    main()
