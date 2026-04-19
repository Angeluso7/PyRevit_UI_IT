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


# --------------------------------------------------
# Paleta de colores (coherente con visor Por Selección)
# --------------------------------------------------
COLOR_BG        = "#1E1E2E"
COLOR_SURFACE   = "#2A2A3E"
COLOR_HEADER_BG = "#16213E"
COLOR_ACCENT    = "#4F7EFF"
COLOR_TEXT      = "#E0E0F0"
COLOR_TEXT_MUTED= "#9898B0"
COLOR_BORDER    = "#3A3A55"
COLOR_ROW_ODD   = "#252538"
COLOR_ROW_EVEN  = "#2A2A3E"
COLOR_SEL_BG    = "#3A4F8C"
COLOR_SEL_FG    = "#FFFFFF"
COLOR_BTN_BG    = "#4F7EFF"
COLOR_BTN_FG    = "#FFFFFF"
COLOR_BTN_HOV   = "#3A6AEF"

FONT_TITLE  = ("Segoe UI", 11, "bold")
FONT_INFO   = ("Segoe UI", 9)
FONT_HEADER = ("Segoe UI", 9, "bold")
FONT_CELL   = ("Segoe UI", 9)
FONT_BTN    = ("Segoe UI", 9, "bold")


# --------------------------------------------------
# Utilidades
# --------------------------------------------------
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


def _nombre_archivo(ruta):
    """Extrae solo el nombre del archivo de una ruta larga."""
    return os.path.basename(ruta) if ruta else ""


# --------------------------------------------------
# Ventana principal
# --------------------------------------------------
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

    headers   = data.get("Headers",   []) or []
    row       = data.get("Row",        {}) or {}
    elem_id   = data.get("ElementId",  "")
    planilla  = data.get("Planilla",   "")
    archivo   = data.get("Archivo",    "")
    nom_arch  = _nombre_archivo(archivo)

    # ── Ventana ──────────────────────────────────
    root = tk.Tk()
    root.title("Visor Elemento — {}".format(planilla or "Sin planilla"))
    root.geometry("680x540")
    root.minsize(480, 360)
    root.configure(bg=COLOR_BG)
    root.resizable(True, True)

    # ── Estilo ttk ───────────────────────────────
    style = ttk.Style(root)
    style.theme_use("clam")

    style.configure("Main.TFrame",       background=COLOR_BG)
    style.configure("Header.TFrame",     background=COLOR_HEADER_BG)
    style.configure("Info.TLabel",
                    background=COLOR_HEADER_BG,
                    foreground=COLOR_TEXT_MUTED,
                    font=FONT_INFO)
    style.configure("Title.TLabel",
                    background=COLOR_HEADER_BG,
                    foreground=COLOR_TEXT,
                    font=FONT_TITLE)
    style.configure("Visor.Treeview",
                    background=COLOR_SURFACE,
                    foreground=COLOR_TEXT,
                    fieldbackground=COLOR_SURFACE,
                    rowheight=24,
                    font=FONT_CELL,
                    borderwidth=0)
    style.configure("Visor.Treeview.Heading",
                    background=COLOR_HEADER_BG,
                    foreground=COLOR_ACCENT,
                    font=FONT_HEADER,
                    relief="flat")
    style.map("Visor.Treeview",
              background=[("selected", COLOR_SEL_BG)],
              foreground=[("selected", COLOR_SEL_FG)])
    style.map("Visor.Treeview.Heading",
              background=[("active", COLOR_BORDER)])
    style.configure("Visor.Vertical.TScrollbar",
                    background=COLOR_BORDER,
                    troughcolor=COLOR_SURFACE,
                    arrowcolor=COLOR_TEXT_MUTED,
                    borderwidth=0)

    # ── Barra superior ───────────────────────────
    header_frame = ttk.Frame(root, style="Header.TFrame", padding=(12, 8))
    header_frame.pack(fill="x", side="top")

    ttk.Label(
        header_frame,
        text=u"Planilla: {}".format(planilla or "—"),
        style="Title.TLabel"
    ).pack(anchor="w")

    ttk.Label(
        header_frame,
        text=u"Elemento ID: {}    Archivo: {}".format(elem_id, nom_arch or archivo),
        style="Info.TLabel"
    ).pack(anchor="w", pady=(2, 0))

    # Separador
    sep = tk.Frame(root, height=1, bg=COLOR_BORDER)
    sep.pack(fill="x")

    # ── Área de tabla ─────────────────────────────
    content_frame = ttk.Frame(root, style="Main.TFrame", padding=(8, 6))
    content_frame.pack(fill="both", expand=True)

    columns = ("Parametro", "Valor")
    tree = ttk.Treeview(
        content_frame,
        columns=columns,
        show="headings",
        style="Visor.Treeview",
        selectmode="browse"
    )

    tree.heading("Parametro", text="Parámetro", anchor="w")
    tree.heading("Valor",     text="Valor",     anchor="w")
    tree.column("Parametro",  width=220, minwidth=140, anchor="w", stretch=True)
    tree.column("Valor",      width=380, minwidth=180, anchor="w", stretch=True)

    # Tags para filas alternadas
    tree.tag_configure("odd",  background=COLOR_ROW_ODD)
    tree.tag_configure("even", background=COLOR_ROW_EVEN)
    tree.tag_configure("empty", foreground=COLOR_TEXT_MUTED)

    for idx, h in enumerate(headers):
        valor = row.get(h, "")
        tag   = "odd" if idx % 2 == 0 else "even"
        if not valor:
            tag = "empty"
        tree.insert("", "end", values=(h, valor if valor else "—"), tags=(tag,))

    vsb = ttk.Scrollbar(
        content_frame,
        orient="vertical",
        command=tree.yview,
        style="Visor.Vertical.TScrollbar"
    )
    tree.configure(yscrollcommand=vsb.set)
    tree.pack(fill="both", expand=True, side="left")
    vsb.pack(fill="y", side="right")

    # ── Contador de parámetros ───────────────────
    total      = len(headers)
    con_valor  = sum(1 for h in headers if row.get(h, ""))
    sep2 = tk.Frame(root, height=1, bg=COLOR_BORDER)
    sep2.pack(fill="x")

    footer_frame = tk.Frame(root, bg=COLOR_HEADER_BG, pady=6)
    footer_frame.pack(fill="x", side="bottom")

    lbl_count = tk.Label(
        footer_frame,
        text=u"  {} de {} parámetros con valor".format(con_valor, total),
        bg=COLOR_HEADER_BG,
        fg=COLOR_TEXT_MUTED,
        font=FONT_INFO
    )
    lbl_count.pack(side="left", padx=8)

    # ── Botón Cerrar ─────────────────────────────
    btn_close = tk.Button(
        footer_frame,
        text="Cerrar",
        bg=COLOR_BTN_BG,
        fg=COLOR_BTN_FG,
        activebackground=COLOR_BTN_HOV,
        activeforeground=COLOR_BTN_FG,
        font=FONT_BTN,
        relief="flat",
        padx=16,
        pady=4,
        cursor="hand2",
        command=root.destroy
    )
    btn_close.pack(side="right", padx=10)

    root.mainloop()


if __name__ == "__main__":
    main()
