# -*- coding: utf-8 -*-
"""
visor_elemento.pyw  v3.0
Visor de parametros de un elemento.
Modos: AGRUPADO (filas unicas) y DETALLADO (una fila por elemento).
Doble clic en fila agrupada (N>1) abre ventana con miembros del grupo.

Requiere (opcional): pip install sv-ttk
Sin sv-ttk usa paleta dark manual.

Uso:
    pythonw.exe visor_elemento.pyw <ruta_repo_elemento_tmp.json>
"""

import sys, os, json, traceback
import tkinter as tk
from tkinter import ttk, messagebox

try:
    import sv_ttk
    _HAS_SV_TTK = True
except ImportError:
    _HAS_SV_TTK = False

# ── Paleta dark manual ───────────────────────────────────────────────────────
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
C_GRP_BG  = "#1A2640"
C_GRP_FG  = "#A0C4FF"

FT_TITLE = ("Segoe UI", 11, "bold")
FT_INFO  = ("Segoe UI",  9)
FT_HDR   = ("Segoe UI",  9, "bold")
FT_CELL  = ("Segoe UI",  9)
FT_BTN   = ("Segoe UI",  9, "bold")

# ── Utilidades ───────────────────────────────────────────────────────────────
def cargar_json(path):
    try:
        if not os.path.exists(path):
            messagebox.showerror("Error",
                "No se encontro el archivo JSON:\n{}".format(path))
            return None
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        messagebox.showerror("Error",
            "Error cargando JSON:\n{}".format(traceback.format_exc()))
        return None

def _basename(ruta):
    return os.path.basename(ruta) if ruta else ""

# ── Estilos manuales ─────────────────────────────────────────────────────────
def _apply_dark(root):
    s = ttk.Style(root)
    s.theme_use("clam")
    s.configure("TFrame",        background=C_BG)
    s.configure("Header.TFrame", background=C_HDR_BG)
    s.configure("Title.TLabel",  background=C_HDR_BG, foreground=C_TEXT,  font=FT_TITLE)
    s.configure("Info.TLabel",   background=C_HDR_BG, foreground=C_MUTED, font=FT_INFO)
    s.configure("TLabel",        background=C_BG,     foreground=C_TEXT,  font=FT_CELL)
    s.configure("Treeview",
                background=C_SURFACE, foreground=C_TEXT,
                fieldbackground=C_SURFACE, rowheight=24, font=FT_CELL, borderwidth=0)
    s.configure("Treeview.Heading",
                background=C_HDR_BG, foreground=C_ACCENT, font=FT_HDR, relief="flat")
    s.map("Treeview",
          background=[("selected", C_SEL_BG)],
          foreground=[("selected", C_SEL_FG)])
    s.map("Treeview.Heading", background=[("active", C_BORDER)])
    s.configure("Vertical.TScrollbar",
                background=C_BORDER, troughcolor=C_SURFACE, arrowcolor=C_MUTED, borderwidth=0)
    s.configure("Horizontal.TScrollbar",
                background=C_BORDER, troughcolor=C_SURFACE, arrowcolor=C_MUTED, borderwidth=0)
    s.configure("TButton",
                background=C_ACCENT, foreground="#FFFFFF", font=FT_BTN,
                borderwidth=0, padding=(8, 4))
    s.map("TButton",
          background=[("active", "#3A6AEF"), ("pressed", "#2A5ACF")])

# ── Logica de agrupacion ─────────────────────────────────────────────────────
def agrupar_filas(headers, all_rows):
    """
    Agrupa all_rows por combinacion de valores en todas las columnas visibles,
    excluyendo campos de identidad unica del elemento.
    Retorna lista de grupos: [{key_vals, count, members}, ...]
    """
    EXCLUDE = {"CodIntBIM", "id", "ID SQL"}
    group_cols   = [h for h in headers if h not in EXCLUDE]
    groups_dict  = {}
    order        = []

    for entry in all_rows:
        row       = entry.get("Row", {})
        key_tuple = tuple(str(row.get(h, "")) for h in group_cols)

        if key_tuple not in groups_dict:
            groups_dict[key_tuple] = {
                "key_vals": {h: str(row.get(h, "")) for h in headers},
                "count"   : 0,
                "members" : [],
            }
            order.append(key_tuple)

        groups_dict[key_tuple]["count"] += 1
        groups_dict[key_tuple]["members"].append(entry)

    return [groups_dict[k] for k in order]

# ── Ventana detalle de miembros ───────────────────────────────────────────────
def mostrar_miembros(parent, grupo, headers):
    top = tk.Toplevel(parent)
    top.title("Miembros del grupo ({} elementos)".format(grupo["count"]))
    top.geometry("800x420")
    top.minsize(500, 300)
    if not _HAS_SV_TTK:
        top.configure(bg=C_BG)
    top.resizable(True, True)
    top.columnconfigure(0, weight=1)
    top.rowconfigure(0, weight=1)

    cols  = ["#", "ElementId", "Archivo"] + headers
    frame = ttk.Frame(top)
    frame.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
    frame.columnconfigure(0, weight=1)
    frame.rowconfigure(0, weight=1)

    tree = ttk.Treeview(frame, columns=cols, show="headings", selectmode="browse")
    tree.heading("#",         text="#",         anchor="center")
    tree.heading("ElementId", text="ElementId", anchor="w")
    tree.heading("Archivo",   text="Archivo",   anchor="w")
    tree.column("#",         width=36,  minwidth=30,  anchor="center", stretch=False)
    tree.column("ElementId", width=100, minwidth=80,  anchor="w",      stretch=False)
    tree.column("Archivo",   width=220, minwidth=100, anchor="w",      stretch=True)
    for h in headers:
        tree.heading(h, text=h, anchor="w")
        tree.column( h, width=110, minwidth=60, anchor="w", stretch=True)

    if not _HAS_SV_TTK:
        tree.tag_configure("odd",  background=C_ODD)
        tree.tag_configure("even", background=C_EVEN)

    for idx, m in enumerate(grupo["members"], 1):
        row  = m.get("Row", {})
        arch = _basename(m.get("Archivo", ""))
        vals = [idx, m.get("ElementId", ""), arch] + [row.get(h, "") for h in headers]
        tag  = () if _HAS_SV_TTK else (("odd" if idx % 2 else "even"),)
        tree.insert("", "end", values=vals, tags=tag)

    vsb = ttk.Scrollbar(frame, orient="vertical",   command=tree.yview)
    hsb = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
    tree.grid(row=0, column=0, sticky="nsew")
    vsb.grid(row=0, column=1, sticky="ns")
    hsb.grid(row=1, column=0, sticky="ew")

    foot = ttk.Frame(top, padding=(6, 4))
    foot.grid(row=1, column=0, sticky="ew")
    foot.columnconfigure(0, weight=1)
    ttk.Button(foot, text="Cerrar", command=top.destroy).grid(
        row=0, column=1, sticky="e")
    ttk.Sizegrip(top).grid(row=2, column=0, sticky="se")

# ── Ventana principal ─────────────────────────────────────────────────────────
def main():
    if len(sys.argv) < 2:
        messagebox.showerror("Error", "No se recibio la ruta del JSON temporal.")
        return

    data = cargar_json(sys.argv[1])
    if not data:
        return

    headers  = data.get("Headers",  []) or []
    row_main = data.get("Row",       {}) or {}
    elem_id  = data.get("ElementId", "")
    planilla = data.get("Planilla",  "")
    archivo  = data.get("Archivo",   "")
    codint   = data.get("CodIntBIM", "")
    all_rows = data.get("AllRows",   []) or []

    if not all_rows:
        all_rows = [{
            "RepoKey"   : data.get("RepoKey", ""),
            "ElementId" : elem_id,
            "Archivo"   : archivo,
            "Row"       : row_main,
        }]

    grupos     = agrupar_filas(headers, all_rows)
    n_grupos   = len(grupos)
    n_elements = len(all_rows)

    # ── Raiz ──────────────────────────────────────────────────────────────
    root = tk.Tk()
    if _HAS_SV_TTK:
        sv_ttk.set_theme("dark")
    else:
        root.configure(bg=C_BG)
        _apply_dark(root)

    root.title("Visor Elemento — {}".format(planilla or "Sin planilla"))
    root.geometry("820x600")
    root.minsize(540, 380)
    root.resizable(True, True)
    root.columnconfigure(0, weight=1)

    # ── Cabecera ──────────────────────────────────────────────────────────
    hdr = ttk.Frame(root, padding=(12, 8))
    if not _HAS_SV_TTK:
        hdr.configure(style="Header.TFrame")
    hdr.grid(row=0, column=0, sticky="ew")

    ttk.Label(hdr, text="Planilla: {}".format(planilla or u"\u2014"),
              font=FT_TITLE,
              **({} if _HAS_SV_TTK else {"style": "Title.TLabel"})
    ).pack(anchor="w")
    ttk.Label(hdr,
              text="ID: {}   CodIntBIM: {}   Archivo: {}".format(
                  elem_id, codint or u"\u2014", _basename(archivo) or archivo),
              font=FT_INFO,
              **({} if _HAS_SV_TTK else {"style": "Info.TLabel"})
    ).pack(anchor="w", pady=(2, 0))

    ttk.Separator(root, orient="horizontal").grid(row=1, column=0, sticky="ew")

    # ── Barra de modo ─────────────────────────────────────────────────────
    toolbar = ttk.Frame(root, padding=(8, 4))
    toolbar.grid(row=2, column=0, sticky="ew")
    toolbar.columnconfigure(3, weight=1)

    modo_var = tk.StringVar(value="agrupado")

    ttk.Label(toolbar, text="Vista:", font=FT_BTN).grid(row=0, column=0, padx=(0, 6))
    ttk.Radiobutton(toolbar, text="Agrupada",  value="agrupado",  variable=modo_var
    ).grid(row=0, column=1, padx=4)
    ttk.Radiobutton(toolbar, text="Detallada", value="detallado", variable=modo_var
    ).grid(row=0, column=2, padx=4)

    lbl_cuenta = ttk.Label(toolbar,
        text="{} grupos / {} elementos".format(n_grupos, n_elements),
        font=FT_INFO)
    lbl_cuenta.grid(row=0, column=3, sticky="e", padx=8)

    ttk.Separator(root, orient="horizontal").grid(row=3, column=0, sticky="ew")

    # ── Frame tabla ───────────────────────────────────────────────────────
    root.rowconfigure(4, weight=1)
    frame_tbl = ttk.Frame(root)
    frame_tbl.grid(row=4, column=0, sticky="nsew", padx=6, pady=4)
    frame_tbl.columnconfigure(0, weight=1)
    frame_tbl.rowconfigure(0, weight=1)

    def rebuild_tree(modo):
        for w in frame_tbl.winfo_children():
            w.destroy()

        if modo == "agrupado":
            cols = ["N"] + headers
        else:
            cols = ["ElementId", "Archivo"] + headers

        tree = ttk.Treeview(frame_tbl, columns=cols, show="headings",
                            selectmode="browse")

        if modo == "agrupado":
            tree.heading("N", text="N", anchor="center")
            tree.column( "N", width=42, minwidth=30, anchor="center", stretch=False)
        else:
            tree.heading("ElementId", text="ElementId", anchor="w")
            tree.heading("Archivo",   text="Archivo",   anchor="w")
            tree.column( "ElementId", width=110, minwidth=80,  anchor="w", stretch=False)
            tree.column( "Archivo",   width=210, minwidth=100, anchor="w", stretch=True)

        for h in headers:
            tree.heading(h, text=h, anchor="w")
            tree.column( h, width=120, minwidth=60, anchor="w", stretch=True)

        if not _HAS_SV_TTK:
            tree.tag_configure("odd",   background=C_ODD)
            tree.tag_configure("even",  background=C_EVEN)
            tree.tag_configure("group", background=C_GRP_BG, foreground=C_GRP_FG)

        if modo == "agrupado":
            for idx, grp in enumerate(grupos):
                vals = [grp["count"]] + [grp["key_vals"].get(h, "") for h in headers]
                if _HAS_SV_TTK:
                    tag = ()
                else:
                    tag = ("group",) if grp["count"] > 1 else \
                          (("odd" if idx % 2 == 0 else "even"),)
                tree.insert("", "end", iid=str(idx), values=vals, tags=tag)

            def on_dbl(event):
                sel = tree.selection()
                if not sel:
                    return
                grp = grupos[int(sel[0])]
                if grp["count"] > 1:
                    mostrar_miembros(root, grp, headers)
            tree.bind("<Double-1>", on_dbl)
        else:
            for idx, entry in enumerate(all_rows):
                row  = entry.get("Row", {})
                arch = _basename(entry.get("Archivo", ""))
                vals = [entry.get("ElementId", ""), arch] + \
                       [row.get(h, "") for h in headers]
                tag  = () if _HAS_SV_TTK else \
                       (("odd" if idx % 2 == 0 else "even"),)
                tree.insert("", "end", values=vals, tags=tag)

        vsb = ttk.Scrollbar(frame_tbl, orient="vertical",   command=tree.yview)
        hsb = ttk.Scrollbar(frame_tbl, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

    rebuild_tree(modo_var.get())

    def on_modo(*_):
        rebuild_tree(modo_var.get())
        n_vis = len(grupos) if modo_var.get() == "agrupado" else n_elements
        lbl_cuenta.config(
            text="{} filas ({} elementos)".format(n_vis, n_elements))

    modo_var.trace_add("write", on_modo)

    # ── Footer ─────────────────────────────────────────────────────────────
    ttk.Separator(root, orient="horizontal").grid(row=5, column=0, sticky="ew")
    footer = ttk.Frame(root, padding=(8, 4))
    footer.grid(row=6, column=0, sticky="ew")
    footer.columnconfigure(0, weight=1)

    con_valor = sum(1 for h in headers if row_main.get(h, ""))
    ttk.Label(footer,
              text="{} / {} parametros con valor".format(con_valor, len(headers)),
              font=FT_INFO).grid(row=0, column=0, sticky="w")
    ttk.Button(footer, text="Cerrar", command=root.destroy).grid(
        row=0, column=1, sticky="e", padx=(8, 0))

    ttk.Sizegrip(root).grid(row=7, column=0, sticky="se")
    root.mainloop()


if __name__ == "__main__":
    main()
