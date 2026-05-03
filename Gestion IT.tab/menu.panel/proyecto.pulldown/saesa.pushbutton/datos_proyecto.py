# -*- coding: utf-8 -*-
"""
datos_proyecto.py — Visor/Editor de datos del proyecto activo (SAESA).
Estilo: dark manual (paleta coherente con el resto de la extensión PyRevit_UI_IT).
No depende de sv-ttk para el tema base; aplica la paleta directamente vía ttk.Style.

Uso (llamado desde script.py):
    python datos_proyecto.py <MASTER_DIR>
"""

import os
import sys
import json
import traceback
import tkinter as tk
from tkinter import ttk, messagebox

# ── Paleta dark (coherente con toda la extensión) ────────────────────────────
PAL = {
    "bg":           "#1e1e2e",   # fondo principal
    "bg_panel":     "#252537",   # paneles / frames internos
    "bg_header":    "#16162a",   # cabecera / barra superior
    "bg_row_even":  "#1e1e2e",   # filas pares del treeview
    "bg_row_odd":   "#252537",   # filas impares del treeview
    "bg_edited":    "#1a3a5c",   # fila con edición pendiente
    "bg_selected":  "#2d5986",   # fila seleccionada
    "fg":           "#d4d4e8",   # texto principal
    "fg_muted":     "#7e7e9a",   # texto secundario / placeholders
    "fg_header":    "#a0c4e8",   # texto de cabecera de columna
    "accent":       "#4e9de0",   # azul eléctrico — acento principal
    "accent2":      "#7ecbae",   # verde menta — acento secundario
    "border":       "#33334a",   # bordes sutiles
    "btn_bg":       "#2d2d44",   # fondo de botones
    "btn_hover":    "#3a3a55",   # hover de botones
    "btn_save":     "#1c5c8a",   # fondo botón Guardar
    "btn_save_fg":  "#d6ecff",   # texto botón Guardar
    "btn_cancel":   "#3a2a2a",   # fondo botón Cancelar
    "btn_cancel_fg":"#ffb3b3",   # texto botón Cancelar
    "separator":    "#33334a",   # separadores
    "font_title":   ("Segoe UI", 13, "bold"),
    "font_sub":     ("Segoe UI", 9),
    "font_header":  ("Segoe UI", 9, "bold"),
    "font_cell":    ("Segoe UI", 9),
    "font_btn":     ("Segoe UI", 9, "bold"),
}

# ── Rutas dinámicas ───────────────────────────────────────────────────────────
_THIS_DIR = os.path.dirname(os.path.abspath(__file__))
_EXT_ROOT = os.path.abspath(os.path.join(_THIS_DIR, '..', '..', '..', '..'))
_MASTER   = os.path.join(_EXT_ROOT, 'data', 'master')
CONFIG_PATH = os.path.join(_MASTER, 'config_proyecto_activo.json')

NON_EDITABLE = {"Archivo", "ElementId", "nombre_archivo", "CodIntBIM"}


def _resolve_data_json_path():
    if len(sys.argv) > 1:
        arg = sys.argv[1].strip('"').strip("'")
        if os.path.isdir(arg):
            return os.path.join(arg, 'datos_tmp.json')
        elif arg.lower().endswith('.json'):
            return arg
        else:
            return os.path.join(arg, 'datos_tmp.json')
    return os.path.join(_EXT_ROOT, 'data', 'temp', 'datos_tmp.json')


DATA_JSON_PATH = _resolve_data_json_path()


# ── Helpers JSON ──────────────────────────────────────────────────────────────
def cargar_json(ruta, default=None, show_err=True):
    try:
        if not os.path.exists(ruta):
            if show_err:
                messagebox.showinfo("Informacion",
                                    "Archivo no encontrado:\n{}".format(ruta))
            return default
        with open(ruta, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        if show_err:
            messagebox.showerror("Error",
                                 "Error cargando JSON:\n{}".format(traceback.format_exc()))
        return default


def guardar_json(ruta, datos):
    try:
        carpeta = os.path.dirname(ruta)
        if carpeta and not os.path.exists(carpeta):
            os.makedirs(carpeta)
        with open(ruta, "w", encoding="utf-8") as f:
            json.dump(datos, f, indent=2, ensure_ascii=False)
        return True
    except Exception:
        messagebox.showerror("Error",
                             "Error guardando JSON:\n{}".format(traceback.format_exc()))
        return False


def get_repo_path():
    if not os.path.exists(CONFIG_PATH):
        messagebox.showerror(
            "Config no encontrada",
            "No se encontro config_proyecto_activo.json en:\n{}".format(CONFIG_PATH))
        return None
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            cfg = json.load(f)
        ruta = (cfg.get("ruta_repositorio_activo") or "").strip()
        if not ruta:
            messagebox.showerror("Config incompleta",
                                 "No se encontro 'ruta_repositorio_activo' en config.")
            return None
        return ruta
    except Exception:
        messagebox.showerror("Error config",
                             "Error leyendo config:\n{}".format(traceback.format_exc()))
        return None


# ── Aplicar estilo dark al ttk.Style ─────────────────────────────────────────
def aplicar_estilo_dark(root):
    style = ttk.Style(root)
    style.theme_use("clam")

    # Frame / Label
    style.configure("TFrame",       background=PAL["bg_panel"])
    style.configure("Header.TFrame",background=PAL["bg_header"])
    style.configure("TLabel",       background=PAL["bg_panel"],
                    foreground=PAL["fg"], font=PAL["font_sub"])
    style.configure("Title.TLabel", background=PAL["bg_header"],
                    foreground=PAL["accent"], font=PAL["font_title"])
    style.configure("Sub.TLabel",   background=PAL["bg_header"],
                    foreground=PAL["fg_muted"], font=PAL["font_sub"])
    style.configure("Tag.TLabel",   background=PAL["bg_header"],
                    foreground=PAL["accent2"], font=("Segoe UI", 8, "bold"))

    # Scrollbars
    style.configure("TScrollbar",
                    background=PAL["bg_panel"],
                    troughcolor=PAL["bg"],
                    bordercolor=PAL["border"],
                    arrowcolor=PAL["fg_muted"])
    style.map("TScrollbar",
              background=[("active", PAL["btn_hover"]),
                          ("!active", PAL["btn_bg"])])

    # Treeview
    style.configure("Treeview",
                    background=PAL["bg_row_even"],
                    fieldbackground=PAL["bg_row_even"],
                    foreground=PAL["fg"],
                    font=PAL["font_cell"],
                    rowheight=26,
                    borderwidth=0,
                    relief="flat")
    style.configure("Treeview.Heading",
                    background=PAL["bg_header"],
                    foreground=PAL["fg_header"],
                    font=PAL["font_header"],
                    relief="flat",
                    borderwidth=0)
    style.map("Treeview",
              background=[("selected", PAL["bg_selected"])],
              foreground=[("selected", "#ffffff")])
    style.map("Treeview.Heading",
              background=[("active", PAL["btn_hover"]),
                          ("!active", PAL["bg_header"])])

    # Buttons
    style.configure("TButton",
                    background=PAL["btn_bg"],
                    foreground=PAL["fg"],
                    font=PAL["font_btn"],
                    borderwidth=1,
                    relief="flat",
                    padding=(10, 6))
    style.map("TButton",
              background=[("active", PAL["btn_hover"]),
                          ("pressed", PAL["border"])],
              foreground=[("active", "#ffffff")])

    style.configure("Save.TButton",
                    background=PAL["btn_save"],
                    foreground=PAL["btn_save_fg"],
                    font=PAL["font_btn"],
                    relief="flat",
                    padding=(10, 6))
    style.map("Save.TButton",
              background=[("active", "#27699e"),
                          ("pressed", PAL["border"])],
              foreground=[("active", "#ffffff")])

    style.configure("Cancel.TButton",
                    background=PAL["btn_cancel"],
                    foreground=PAL["btn_cancel_fg"],
                    font=PAL["font_btn"],
                    relief="flat",
                    padding=(10, 6))
    style.map("Cancel.TButton",
              background=[("active", "#5a3535"),
                          ("pressed", PAL["border"])],
              foreground=[("active", "#ffffff")])

    # Sizegrip
    style.configure("TSizegrip", background=PAL["bg"])


# ── Treeview editable ─────────────────────────────────────────────────────────
class EditableTreeview(ttk.Treeview):
    def __init__(self, master, headers, rows, **kwargs):
        super().__init__(master, columns=headers, show="headings", **kwargs)
        self.headers       = headers
        self.draft_edits   = {}
        self.editing_entry = None

        # Anchos de columna
        col_widths = {"ElementId": 90, "Archivo": 130, "CodIntBIM": 120}
        for h in headers:
            w = col_widths.get(h, 150)
            self.heading(h, text=h)
            self.column(h, width=w, anchor="center", stretch=False)

        # Insertar filas con alternancia de color
        for i, row in enumerate(rows):
            values = [str(row.get(h, "") or "") for h in headers]
            tag = "odd" if i % 2 else "even"
            self.insert("", "end", iid=str(i), values=values, tags=(tag,))

        self.tag_configure("even",   background=PAL["bg_row_even"],
                           foreground=PAL["fg"])
        self.tag_configure("odd",    background=PAL["bg_row_odd"],
                           foreground=PAL["fg"])
        self.tag_configure("edited", background=PAL["bg_edited"],
                           foreground="#d0eaff")

        self.bind("<Double-1>", self.on_double_click)

    def on_double_click(self, event):
        try:
            if self.editing_entry:
                self.editing_entry.destroy()
                self.editing_entry = None

            region = self.identify_region(event.x, event.y)
            if region != "cell":
                return

            col = self.identify_column(event.x)
            row = self.identify_row(event.y)
            if not row or not col:
                return

            col_idx = int(col.replace("#", "")) - 1
            if col_idx < 0 or col_idx >= len(self.headers):
                return
            header = self.headers[col_idx]
            if header in NON_EDITABLE:
                return

            bbox = self.bbox(row, col)
            if not bbox:
                return
            x, y, width, height = bbox

            current_val = self.set(row, header)

            entry = tk.Entry(
                self,
                font=PAL["font_cell"],
                bg=PAL["bg_edited"],
                fg="#d0eaff",
                insertbackground="#d0eaff",
                relief="flat",
                bd=1,
                highlightthickness=1,
                highlightcolor=PAL["accent"],
                highlightbackground=PAL["border"],
            )
            entry.place(x=x, y=y, width=width, height=height)
            entry.insert(0, current_val)
            entry.select_range(0, tk.END)
            entry.focus()
            self.editing_entry = entry

            def _commit(event=None, r=row, h=header):
                if not self.editing_entry:
                    return
                newval = self.editing_entry.get()
                self.editing_entry.destroy()
                self.editing_entry = None
                self.set(r, h, newval)
                self.draft_edits.setdefault(r, {})[h] = newval
                tags = [t for t in self.item(r, "tags")
                        if t not in ("even", "odd")]
                tags.append("edited")
                self.item(r, tags=tags)

            def _cancel(event=None):
                if self.editing_entry:
                    self.editing_entry.destroy()
                    self.editing_entry = None

            entry.bind("<Return>",   _commit)
            entry.bind("<Tab>",      _commit)
            entry.bind("<FocusOut>", _commit)
            entry.bind("<Escape>",   _cancel)

        except Exception:
            messagebox.showerror("Error edicion",
                                 "Error al editar celda:\n{}".format(
                                     traceback.format_exc()))


# ── Barra de estado ───────────────────────────────────────────────────────────
class StatusBar(tk.Frame):
    def __init__(self, master, **kwargs):
        super().__init__(master, bg=PAL["bg_header"],
                         bd=0, relief="flat", **kwargs)
        self._var = tk.StringVar(value="Listo")
        tk.Label(
            self,
            textvariable=self._var,
            bg=PAL["bg_header"],
            fg=PAL["fg_muted"],
            font=("Segoe UI", 8),
            anchor="w",
            padx=8,
        ).pack(side="left", fill="x", expand=True)

    def set(self, msg):
        self._var.set(msg)

    def clear(self):
        self._var.set("Listo")


# ── Ventana principal ─────────────────────────────────────────────────────────
def main():
    meta      = cargar_json(DATA_JSON_PATH) or {}
    headers   = meta.get("Headers", []) or []
    data_rows = meta.get("Rows",    []) or []
    titulo    = meta.get("Titulo",  "SAESA — Datos del Proyecto")
    proyecto  = meta.get("Proyecto", "")
    n_filas   = len(data_rows)

    if not headers:
        messagebox.showerror("Sin datos",
                             "No se encontraron columnas en el JSON.\n{}".format(
                                 DATA_JSON_PATH))
        return

    root = tk.Tk()
    root.title(titulo)
    root.geometry("1060x600")
    root.minsize(700, 420)
    root.configure(bg=PAL["bg"])

    aplicar_estilo_dark(root)

    root.columnconfigure(0, weight=1)
    root.rowconfigure(1, weight=1)

    # ── Cabecera ──────────────────────────────────────────────────────────────
    header_frame = ttk.Frame(root, style="Header.TFrame")
    header_frame.grid(row=0, column=0, sticky="ew")
    header_frame.columnconfigure(1, weight=1)

    # Banda de color izquierda (acento vertical)
    tk.Frame(header_frame, bg=PAL["accent"], width=4).grid(
        row=0, column=0, rowspan=3, sticky="ns")

    ttk.Label(header_frame, text=u"\u26a1  " + titulo,
              style="Title.TLabel").grid(
        row=0, column=1, sticky="w", padx=(12, 0), pady=(10, 2))

    sub_txt = u"Proyecto: {}   \u2022   {} elementos".format(
        proyecto or "—", n_filas)
    ttk.Label(header_frame, text=sub_txt,
              style="Sub.TLabel").grid(
        row=1, column=1, sticky="w", padx=(12, 0), pady=(0, 2))

    ttk.Label(header_frame,
              text=u"Doble clic sobre una celda para editar  \u2014  "
                   u"campos bloqueados: ElementId, Archivo, CodIntBIM",
              style="Tag.TLabel").grid(
        row=2, column=1, sticky="w", padx=(12, 0), pady=(0, 8))

    # Separador horizontal
    tk.Frame(root, bg=PAL["border"], height=1).grid(
        row=0, column=0, sticky="sew")

    # ── Área de tabla ─────────────────────────────────────────────────────────
    table_frame = ttk.Frame(root)
    table_frame.grid(row=1, column=0, sticky="nsew", padx=0, pady=0)
    table_frame.columnconfigure(0, weight=1)
    table_frame.rowconfigure(0, weight=1)

    tree = EditableTreeview(
        table_frame, headers, data_rows, selectmode="browse")
    tree.grid(row=0, column=0, sticky="nsew")

    vsb = ttk.Scrollbar(table_frame, orient="vertical",   command=tree.yview)
    vsb.grid(row=0, column=1, sticky="ns")
    hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=tree.xview)
    hsb.grid(row=1, column=0, sticky="ew")
    tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

    # ── Barra de estado ───────────────────────────────────────────────────────
    status = StatusBar(root)
    status.grid(row=2, column=0, sticky="ew")

    # ── Botones ───────────────────────────────────────────────────────────────
    btn_frame = tk.Frame(root, bg=PAL["bg"], pady=8)
    btn_frame.grid(row=3, column=0, sticky="ew")
    btn_frame.columnconfigure(0, weight=1)
    btn_frame.columnconfigure(1, weight=0)
    btn_frame.columnconfigure(2, weight=0)
    btn_frame.columnconfigure(3, weight=1)

    counter_var = tk.StringVar(value="{} elementos  |  0 cambios pendientes".format(n_filas))
    tk.Label(btn_frame, textvariable=counter_var,
             bg=PAL["bg"], fg=PAL["fg_muted"],
             font=("Segoe UI", 8),
             anchor="w").grid(row=0, column=0, sticky="w", padx=12)

    def _update_counter():
        n_edits = sum(len(v) for v in tree.draft_edits.values())
        counter_var.set(
            "{} elementos  |  {} cambio{} pendiente{}".format(
                n_filas,
                n_edits,
                "s" if n_edits != 1 else "",
                "s" if n_edits != 1 else "",
            )
        )

    def guardar():
        if not tree.draft_edits:
            status.set("Sin cambios para guardar.")
            messagebox.showinfo("Sin cambios", "No hay cambios para guardar.")
            return
        repo_path = get_repo_path()
        if not repo_path:
            return
        bd = cargar_json(repo_path, default={}, show_err=False) or {}
        for row_iid, cambios in tree.draft_edits.items():
            idx = int(row_iid)
            if 0 <= idx < len(data_rows):
                orig   = data_rows[idx]
                archivo = orig.get("Archivo", "")
                elemid  = orig.get("ElementId", "")
                clave   = "{}_{}".format(archivo, elemid)
                entrada = dict(bd.get(clave, {})) if isinstance(bd.get(clave), dict) else {}
                for h, v in cambios.items():
                    entrada[h] = v
                bd[clave] = entrada
        if guardar_json(repo_path, bd):
            status.set(u"Guardado correctamente \u2192 {}".format(repo_path))
            messagebox.showinfo("Guardado",
                                "Cambios guardados en:\n{}".format(repo_path))
            tree.draft_edits.clear()
            _update_counter()
            root.destroy()

    def cancelar():
        if tree.draft_edits:
            if not messagebox.askyesno(
                    "Descartar cambios",
                    "Hay cambios sin guardar. Salir de todas formas?"):
                return
        root.destroy()

    # Observar ediciones para actualizar contador
    orig_double = tree.on_double_click

    def _patched_double(event):
        orig_double(event)
        root.after(200, _update_counter)

    tree.bind("<Double-1>", _patched_double)

    ttk.Button(btn_frame, text=u"  \u2714  Guardar",
               style="Save.TButton", command=guardar).grid(
        row=0, column=1, padx=(0, 6))

    ttk.Button(btn_frame, text=u"  \u2715  Cancelar",
               style="Cancel.TButton", command=cancelar).grid(
        row=0, column=2, padx=(0, 12))

    ttk.Sizegrip(root).grid(row=4, column=0, sticky="se")

    root.mainloop()


if __name__ == "__main__":
    main()
