# -*- coding: utf-8 -*-
"""
mostrar_tabla_tk.pyw  v2.1
Visor/editor de parametros por planilla, agrupado por CodIntBIM.

Estilo: sv-ttk dark (Sun Valley theme).
  pip install sv-ttk

- Una fila por CodIntBIM unico.
- Columna "Cantidad" al inicio.
- <<VARIOS>> cuando hay distintos valores en el grupo.
- Doble clic edita y propaga a todos los elementos del grupo.
- CodIntBIM / Archivo / ElementId / nombre_archivo: solo lectura.
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

MARKER_VARIOS = "<<VARIOS>>"

# ── Rutas dinamicas ──────────────────────────────────────────────────────────
_THIS_DIR = os.path.dirname(os.path.abspath(__file__))
_EXT_ROOT = os.path.abspath(os.path.join(_THIS_DIR, '..', '..', '..', '..'))
_MASTER   = os.path.join(_EXT_ROOT, 'data', 'master')
CONFIG_PATH = os.path.join(_MASTER, 'config_proyecto_activo.json')

if len(sys.argv) > 1:
    PLANILLA_META_PATH = sys.argv[1]
else:
    PLANILLA_META_PATH = os.path.join(_EXT_ROOT, 'data', 'temp', 'planilla_meta_tmp.json')


# ── Utilidades JSON ──────────────────────────────────────────────────────────
def cargar_json(ruta, show_not_found=True, title="Error"):
    try:
        if not ruta:
            return None
        if not os.path.exists(ruta):
            if show_not_found:
                messagebox.showinfo("Informacion", "Archivo no encontrado:\n{}".format(ruta))
            return None
        with open(ruta, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        messagebox.showerror(title, "Error cargando JSON:\n{}".format(traceback.format_exc()))
        return None


def guardar_json(ruta, datos):
    try:
        if not ruta:
            messagebox.showerror("Error", "No se definio ruta valida para guardar.")
            return False
        carpeta = os.path.dirname(ruta)
        if carpeta and not os.path.exists(carpeta):
            os.makedirs(carpeta)
        with open(ruta, "w", encoding="utf-8") as f:
            json.dump(datos, f, indent=2, ensure_ascii=False)
        return True
    except Exception:
        messagebox.showerror("Error", "Error guardando JSON:\n{}".format(traceback.format_exc()))
        return False


def get_repo_path_from_config():
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
                                 "'ruta_repositorio_activo' vacia en config.")
            return None
        return ruta
    except Exception:
        messagebox.showerror("Error config",
                             "Error leyendo config:\n{}".format(traceback.format_exc()))
        return None


# ── Normalizacion ────────────────────────────────────────────────────────────
def _norm_display(val):
    if val in (None, "", " "):
        return "-"
    return str(val)


def _norm_save(val):
    if val in ("-", None, " ", MARKER_VARIOS):
        return ""
    return str(val)


# ── Agrupamiento por CodIntBIM ───────────────────────────────────────────────
def build_groups(data_rows, edit_headers, cods_por_clave, valores_por_clave):
    groups_order  = []
    group_members = {}
    group_values  = {}

    for row in data_rows:
        codint  = (row.get("CodIntBIM") or "").strip()
        archivo = (row.get("Archivo")   or "").strip()
        elemid  = (row.get("ElementId") or "").strip()
        if not codint or not archivo or not elemid:
            continue
        clave = "{}_{}".format(archivo, elemid)

        if codint not in group_members:
            groups_order.append(codint)
            group_members[codint] = []
            group_values[codint]  = {h: set() for h in edit_headers}

        group_members[codint].append(clave)

        vals_src = valores_por_clave.get(clave) or row
        for h in edit_headers:
            v = str(vals_src.get(h, "") or "")
            group_values[codint][h].add(v)

    group_row = {}
    for codint in groups_order:
        row_display = {"CodIntBIM": codint, "Cantidad": str(len(group_members[codint]))}
        for h in edit_headers:
            vals = group_values[codint][h] - {""}
            if len(vals) == 0:
                row_display[h] = "-"
            elif len(vals) == 1:
                row_display[h] = vals.pop()
            else:
                row_display[h] = MARKER_VARIOS
        group_row[codint] = row_display

    return groups_order, group_members, group_row


# ── Treeview editable agrupado ───────────────────────────────────────────────
NON_EDITABLE = {"CodIntBIM", "Archivo", "ElementId", "nombre_archivo", "Cantidad"}


class GroupedEditableTreeview(ttk.Treeview):
    def __init__(self, master, gui_headers, groups_order,
                 group_members, group_row, **kwargs):
        super().__init__(master, columns=gui_headers, show="headings", **kwargs)
        self.gui_headers   = gui_headers
        self.group_members = group_members
        self.group_row     = group_row
        self.draft_edits   = {}
        self.editing_entry = None

        col_widths = {"Cantidad": 70, "CodIntBIM": 130}
        for h in gui_headers:
            w = col_widths.get(h, 150)
            self.heading(h, text=h)
            self.column(h, width=w, anchor="center", stretch=False)

        for codint in groups_order:
            row = group_row[codint]
            values = [_norm_display(row.get(h, "")) for h in gui_headers]
            self.insert("", "end", iid=codint, values=values)

        self.tag_configure("edited", background="#1e3a5f")
        self.bind("<Double-1>", self.on_double_click)

    def on_double_click(self, event):
        try:
            if self.editing_entry is not None:
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
            if col_idx < 0 or col_idx >= len(self.gui_headers):
                return
            header = self.gui_headers[col_idx]
            if header in NON_EDITABLE:
                return

            bbox = self.bbox(row, col)
            if not bbox:
                return
            x, y, width, height = bbox

            current_val = self.set(row, header)
            if current_val == MARKER_VARIOS:
                current_val = ""

            self.editing_entry = tk.Entry(self)
            self.editing_entry.place(x=x, y=y, width=width, height=height)
            self.editing_entry.insert(0, current_val)
            self.editing_entry.select_range(0, tk.END)
            self.editing_entry.focus()

            def _commit(event=None, codint=row, h=header):
                if self.editing_entry is None:
                    return
                newval = self.editing_entry.get()
                self.editing_entry.destroy()
                self.editing_entry = None
                display = newval if newval != "" else "-"
                self.set(codint, h, display)
                if codint not in self.draft_edits:
                    self.draft_edits[codint] = {}
                self.draft_edits[codint][h] = newval
                tags = list(self.item(codint, "tags"))
                if "edited" not in tags:
                    tags.append("edited")
                self.item(codint, tags=tags)

            def _cancel(event=None):
                if self.editing_entry is not None:
                    self.editing_entry.destroy()
                    self.editing_entry = None

            self.editing_entry.bind("<Return>",   _commit)
            self.editing_entry.bind("<Tab>",      _commit)
            self.editing_entry.bind("<FocusOut>", _commit)
            self.editing_entry.bind("<Escape>",   _cancel)

        except Exception:
            messagebox.showerror("Error edicion",
                                 "Error al editar celda:\n{}".format(traceback.format_exc()))


# ── GUI principal ────────────────────────────────────────────────────────────
def main():
    meta = cargar_json(PLANILLA_META_PATH, show_not_found=True, title="Error meta") or {}
    headers           = meta.get("Headers", [])         or []
    codigo_planilla   = meta.get("CodigoPlanilla", "")  or ""
    nombre_planilla   = meta.get("NombrePlanilla", "")  or "Planilla"
    data_path         = meta.get("DataPath", "")        or ""
    cods_por_clave    = meta.get("CodsPorClave", {})    or {}
    valores_por_clave = meta.get("ValoresPorClave", {}) or {}

    if not headers or not codigo_planilla:
        messagebox.showerror("Datos incompletos", "No se encontraron Headers o CodigoPlanilla.")
        return
    if not data_path:
        messagebox.showerror("Datos incompletos", "No se encontro DataPath en el meta.")
        return

    data_rows = cargar_json(data_path, show_not_found=False, title="Error datos") or []
    if not isinstance(data_rows, list):
        data_rows = []
    if not data_rows:
        messagebox.showinfo("Sin datos", "No se encontraron filas en el dataset.")
        return

    IDENTITY     = {"Archivo", "ElementId", "nombre_archivo", "CodIntBIM"}
    edit_headers = [h for h in headers if h not in IDENTITY]

    groups_order, group_members, group_row = build_groups(
        data_rows, edit_headers, cods_por_clave, valores_por_clave)

    if not groups_order:
        messagebox.showinfo("Sin datos", "No se encontraron grupos CodIntBIM validos.")
        return

    gui_headers = ["Cantidad", "CodIntBIM"] + edit_headers

    # ── Ventana ───────────────────────────────────────────────────────────────
    root = tk.Tk()

    if _HAS_SV_TTK:
        sv_ttk.set_theme("dark")

    root.title("Edicion Planilla [agrupado CodIntBIM] - {}".format(nombre_planilla))
    root.geometry("1100x580")
    root.minsize(700, 400)
    root.columnconfigure(0, weight=1)
    root.rowconfigure(1, weight=1)

    lbl_info = ttk.Label(
        root,
        text="Planilla: {}   |   Codigo: {}   |   Grupos unicos: {}   |   Total elementos: {}".format(
            nombre_planilla, codigo_planilla, len(groups_order), len(data_rows)),
        anchor="w", padding=(10, 6))
    lbl_info.grid(row=0, column=0, sticky="ew")

    frame = ttk.Frame(root)
    frame.grid(row=1, column=0, sticky="nsew")
    frame.columnconfigure(0, weight=1)
    frame.rowconfigure(0, weight=1)

    tree = GroupedEditableTreeview(
        frame, gui_headers, groups_order,
        group_members, group_row,
        selectmode="browse")
    tree.grid(row=0, column=0, sticky="nsew")

    vsb = ttk.Scrollbar(frame, orient="vertical",   command=tree.yview)
    vsb.grid(row=0, column=1, sticky="ns")
    hsb = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
    hsb.grid(row=1, column=0, sticky="ew")
    tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

    ttk.Sizegrip(root).grid(row=3, column=0, sticky="se")

    # ── Guardar ───────────────────────────────────────────────────────────────
    def guardar():
        try:
            if not tree.draft_edits:
                messagebox.showinfo("Sin cambios", "No hay cambios para guardar.")
                return

            repo_path = get_repo_path_from_config()
            if not repo_path:
                return

            bd = cargar_json(repo_path, show_not_found=False, title="Error repo") or {}
            if not isinstance(bd, dict):
                bd = {}

            total_elementos = 0
            for codint, cambios in tree.draft_edits.items():
                if not cambios:
                    continue
                claves = tree.group_members.get(codint, [])
                for clave in claves:
                    entrada = dict(bd.get(clave, {})) if isinstance(bd.get(clave), dict) else {}
                    if not entrada.get("CodIntBIM"):
                        entrada["CodIntBIM"] = codint
                    if not entrada.get("Archivo") or not entrada.get("ElementId"):
                        partes = clave.rsplit("_", 1)
                        if len(partes) == 2:
                            entrada["Archivo"]   = partes[0]
                            entrada["ElementId"] = partes[1]
                    vals_oficiales = valores_por_clave.get(clave, {}) or {}
                    for h in edit_headers:
                        if h not in entrada or not entrada[h]:
                            v = vals_oficiales.get(h, "")
                            if v:
                                entrada[h] = v
                    for h, newval in cambios.items():
                        entrada[h] = _norm_save(newval)
                    bd[clave] = entrada
                    total_elementos += 1

            if total_elementos == 0:
                messagebox.showinfo("Sin cambios", "No se generaron cambios efectivos.")
                return

            if guardar_json(repo_path, bd):
                messagebox.showinfo(
                    "Guardado",
                    "Cambios propagados a {} elemento(s) en {} grupo(s).\nRepositorio:\n{}".format(
                        total_elementos, len(tree.draft_edits), repo_path))
                tree.draft_edits.clear()
                root.destroy()

        except Exception:
            messagebox.showerror("Error al guardar", "Error:\n{}".format(traceback.format_exc()))

    def cancelar():
        if tree.draft_edits:
            if not messagebox.askyesno("Descartar cambios",
                                       "Hay cambios sin guardar. Salir de todas formas?"):
                return
        root.destroy()

    # ── Botones ───────────────────────────────────────────────────────────────
    btn_frame = ttk.Frame(root)
    btn_frame.grid(row=2, column=0, sticky="ew", pady=6, padx=6)
    btn_frame.columnconfigure(0, weight=1)
    btn_frame.columnconfigure(1, weight=1)

    ttk.Button(btn_frame, text="Guardar — propagar a todos los elementos del grupo",
               command=guardar).grid(row=0, column=0, sticky="ew", padx=(0, 4))
    ttk.Button(btn_frame, text="Cancelar",
               command=cancelar).grid(row=0, column=1, sticky="ew", padx=(4, 0))

    root.mainloop()


if __name__ == "__main__":
    main()
