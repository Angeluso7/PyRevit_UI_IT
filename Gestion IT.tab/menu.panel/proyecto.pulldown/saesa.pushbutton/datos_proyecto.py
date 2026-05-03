# -*- coding: utf-8 -*-
"""
Visor/editor de datos del proyecto activo.
Estilo: sv-ttk dark (Sun Valley theme).
  pip install sv-ttk
"""

import os
import sys
import json
import traceback
import tkinter as tk
from tkinter import ttk, messagebox

try:
    import sv_ttk
    _HAS_SV_TTK = True
except ImportError:
    _HAS_SV_TTK = False

# ── Rutas dinamicas ──────────────────────────────────────────────────────────
_THIS_DIR = os.path.dirname(os.path.abspath(__file__))
_EXT_ROOT = os.path.abspath(os.path.join(_THIS_DIR, '..', '..', '..', '..'))
_MASTER   = os.path.join(_EXT_ROOT, 'data', 'master')
CONFIG_PATH = os.path.join(_MASTER, 'config_proyecto_activo.json')

# ── Resolución de DATA_JSON_PATH ─────────────────────────────────────────────
# script.py pasa MASTER_DIR (carpeta) como sys.argv[1].
# Si el argumento es una carpeta, se construye la ruta al archivo JSON
# esperado dentro de ella (datos_tmp.json).  Si es un archivo .json se
# usa directamente.  Si no hay argumento, se usa la ruta por defecto.
def _resolve_data_json_path():
    if len(sys.argv) > 1:
        arg = sys.argv[1].strip('"').strip("'")
        if os.path.isdir(arg):
            # Se recibió una carpeta (p.ej. MASTER_DIR) — apuntar al JSON
            return os.path.join(arg, 'datos_tmp.json')
        elif arg.lower().endswith('.json'):
            return arg
        else:
            # Ruta desconocida: tratar como carpeta por seguridad
            return os.path.join(arg, 'datos_tmp.json')
    return os.path.join(_EXT_ROOT, 'data', 'temp', 'datos_tmp.json')

DATA_JSON_PATH = _resolve_data_json_path()


def cargar_json(ruta, default=None, show_err=True):
    try:
        if not os.path.exists(ruta):
            if show_err:
                messagebox.showinfo("Informacion", "Archivo no encontrado:\n{}".format(ruta))
            return default
        with open(ruta, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        if show_err:
            messagebox.showerror("Error", "Error cargando JSON:\n{}".format(traceback.format_exc()))
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
        messagebox.showerror("Error", "Error guardando JSON:\n{}".format(traceback.format_exc()))
        return False


def get_repo_path():
    if not os.path.exists(CONFIG_PATH):
        messagebox.showerror("Config no encontrada",
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


NON_EDITABLE = {"Archivo", "ElementId", "nombre_archivo", "CodIntBIM"}


class EditableTreeview(ttk.Treeview):
    def __init__(self, master, headers, rows, **kwargs):
        super().__init__(master, columns=headers, show="headings", **kwargs)
        self.headers      = headers
        self.draft_edits  = {}
        self.editing_entry = None

        col_widths = {"ElementId": 90, "Archivo": 130, "CodIntBIM": 120}
        for h in headers:
            w = col_widths.get(h, 150)
            self.heading(h, text=h)
            self.column(h, width=w, anchor="center", stretch=False)

        for i, row in enumerate(rows):
            values = [str(row.get(h, "") or "") for h in headers]
            self.insert("", "end", iid=str(i), values=values)

        self.tag_configure("edited", background="#1e3a5f")
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

            self.editing_entry = tk.Entry(self)
            self.editing_entry.place(x=x, y=y, width=width, height=height)
            self.editing_entry.insert(0, current_val)
            self.editing_entry.select_range(0, tk.END)
            self.editing_entry.focus()

            def _commit(event=None, r=row, h=header):
                if not self.editing_entry:
                    return
                newval = self.editing_entry.get()
                self.editing_entry.destroy()
                self.editing_entry = None
                self.set(r, h, newval)
                if r not in self.draft_edits:
                    self.draft_edits[r] = {}
                self.draft_edits[r][h] = newval
                tags = list(self.item(r, "tags"))
                if "edited" not in tags:
                    tags.append("edited")
                self.item(r, tags=tags)

            def _cancel(event=None):
                if self.editing_entry:
                    self.editing_entry.destroy()
                    self.editing_entry = None

            self.editing_entry.bind("<Return>",   _commit)
            self.editing_entry.bind("<Tab>",      _commit)
            self.editing_entry.bind("<FocusOut>", _commit)
            self.editing_entry.bind("<Escape>",   _cancel)

        except Exception:
            messagebox.showerror("Error edicion",
                                 "Error al editar celda:\n{}".format(traceback.format_exc()))


def main():
    meta = cargar_json(DATA_JSON_PATH) or {}
    headers   = meta.get("Headers", []) or []
    data_rows = meta.get("Rows", [])    or []
    titulo    = meta.get("Titulo", "Datos proyecto")

    if not headers:
        messagebox.showerror("Sin datos", "No se encontraron columnas en el JSON.")
        return

    root = tk.Tk()

    if _HAS_SV_TTK:
        sv_ttk.set_theme("dark")

    root.title(titulo)
    root.geometry("1000x560")
    root.minsize(600, 380)
    root.columnconfigure(0, weight=1)
    root.rowconfigure(1, weight=1)

    ttk.Label(root, text=titulo, padding=(8, 6)).grid(row=0, column=0, sticky="ew")

    frame = ttk.Frame(root)
    frame.grid(row=1, column=0, sticky="nsew")
    frame.columnconfigure(0, weight=1)
    frame.rowconfigure(0, weight=1)

    tree = EditableTreeview(frame, headers, data_rows, selectmode="browse")
    tree.grid(row=0, column=0, sticky="nsew")

    vsb = ttk.Scrollbar(frame, orient="vertical",   command=tree.yview)
    vsb.grid(row=0, column=1, sticky="ns")
    hsb = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
    hsb.grid(row=1, column=0, sticky="ew")
    tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

    ttk.Sizegrip(root).grid(row=3, column=0, sticky="se")

    def guardar():
        if not tree.draft_edits:
            messagebox.showinfo("Sin cambios", "No hay cambios para guardar.")
            return
        repo_path = get_repo_path()
        if not repo_path:
            return
        bd = cargar_json(repo_path, default={}, show_err=False) or {}
        for row_iid, cambios in tree.draft_edits.items():
            idx = int(row_iid)
            if 0 <= idx < len(data_rows):
                orig = data_rows[idx]
                archivo = orig.get("Archivo", "")
                elemid  = orig.get("ElementId", "")
                clave   = "{}_{}".format(archivo, elemid)
                entrada = dict(bd.get(clave, {})) if isinstance(bd.get(clave), dict) else {}
                for h, v in cambios.items():
                    entrada[h] = v
                bd[clave] = entrada
        if guardar_json(repo_path, bd):
            messagebox.showinfo("Guardado", "Cambios guardados en:\n{}".format(repo_path))
            tree.draft_edits.clear()
            root.destroy()

    def cancelar():
        if tree.draft_edits:
            if not messagebox.askyesno("Descartar cambios",
                                       "Hay cambios sin guardar. Salir de todas formas?"):
                return
        root.destroy()

    btn_frame = ttk.Frame(root)
    btn_frame.grid(row=2, column=0, sticky="ew", pady=6, padx=6)
    btn_frame.columnconfigure(0, weight=1)
    btn_frame.columnconfigure(1, weight=1)

    ttk.Button(btn_frame, text="Guardar", command=guardar).grid(
        row=0, column=0, sticky="ew", padx=(0, 4))
    ttk.Button(btn_frame, text="Cancelar", command=cancelar).grid(
        row=0, column=1, sticky="ew", padx=(4, 0))

    root.mainloop()


if __name__ == "__main__":
    main()
