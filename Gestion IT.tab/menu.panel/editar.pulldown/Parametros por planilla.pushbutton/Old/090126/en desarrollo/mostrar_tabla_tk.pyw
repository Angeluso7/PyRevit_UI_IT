# -*- coding: utf-8 -*-

import os
import sys
import json
import traceback
import tkinter as tk
from tkinter import ttk, messagebox

DATA_DIR = os.path.join(
    os.path.expanduser("~"),
    r"AppData\Roaming\MyPyRevitExtention\PyRevitIT.extension\data"
)

CONFIG_PATH = os.path.join(DATA_DIR, "config_proyecto_activo.json")

if len(sys.argv) > 1:
    PLANILLA_META_PATH = sys.argv[1]
else:
    PLANILLA_META_PATH = os.path.join(os.path.dirname(__file__), "planilla_meta_tmp.json")


def cargar_json(ruta, show_not_found=True, title="Error"):
    try:
        if not ruta:
            return None
        if not os.path.exists(ruta):
            if show_not_found:
                messagebox.showinfo(
                    "Información",
                    "Archivo no encontrado:\n{}".format(ruta)
                )
            return None
        with open(ruta, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        messagebox.showerror(
            title,
            "Error cargando JSON:\n{}".format(traceback.format_exc())
        )
        return None


def guardar_json(ruta, datos):
    try:
        if not ruta:
            messagebox.showerror(
                "Error",
                "No se definió ruta válida de repositorio para guardar."
            )
            return
        carpeta = os.path.dirname(ruta)
        if carpeta and not os.path.exists(carpeta):
            os.makedirs(carpeta)
        with open(ruta, "w", encoding="utf-8") as f:
            json.dump(datos, f, indent=2, ensure_ascii=False)
    except Exception:
        messagebox.showerror(
            "Error",
            "Error guardando JSON:\n{}".format(traceback.format_exc())
        )


def get_repo_path_from_config():
    if not os.path.exists(CONFIG_PATH):
        messagebox.showerror(
            "Config no encontrada",
            "No se encontró config_proyecto_activo.json en:\n{}\\n\\n"
            "No se puede determinar el repositorio de datos activo."
            .format(CONFIG_PATH)
        )
        return None
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            cfg = json.load(f)
        ruta = (cfg.get("ruta_repositorio_activo") or "").strip()
        if not ruta:
            messagebox.showerror(
                "Config incompleta",
                "En config_proyecto_activo.json no se encontró 'ruta_repositorio_activo' "
                "o está vacía.\\n\\nNo se puede determinar el repositorio de datos activo."
            )
            return None
        return ruta
    except Exception:
        messagebox.showerror(
            "Error config",
            "Error leyendo config_proyecto_activo.json:\\n{}"
            .format(traceback.format_exc())
        )
        return None


def _view_to_repo(val):
    """Convierte '-' a '' al guardar en BD."""
    if val == "-" or val is None:
        return ""
    return val


class EditableTreeview(ttk.Treeview):
    def __init__(self, master, headers, data_rows, **kwargs):
        super(EditableTreeview, self).__init__(
            master, columns=headers, show="headings", **kwargs
        )

        self.headers = headers
        self.data_rows = data_rows

        for h in headers:
            self.heading(h, text=h)
            self.column(h, width=150, anchor="center", stretch=False)

        for i, row in enumerate(data_rows):
            values = []
            for h in headers:
                val = row.get(h, "")
                if val in (None, "", " "):
                    val = "-"
                values.append(val)
            self.insert("", "end", iid=str(i), values=values)

        self.edited_rows = set()
        self.editing_entry = None

        self.bind("<Double-1>", self.on_double_click)

    def on_double_click(self, event):
        try:
            if self.editing_entry:
                self.editing_entry.destroy()
                self.editing_entry = None

            region = self.identify("region", event.x, event.y)
            if region != "cell":
                return

            col = self.identify_column(event.x)
            row = self.identify_row(event.y)
            if not row or not col:
                return

            x, y, width, height = self.bbox(row, col)
            col_idx = int(col.replace("#", "")) - 1
            if col_idx < 0:
                return

            header = self.headers[col_idx]
            # Solo bloquear Archivo, ElementId, nombre_archivo
            if header in ("Archivo", "ElementId", "nombre_archivo"):
                return

            value = self.set(row, header)

            self.editing_entry = tk.Entry(self)
            self.editing_entry.place(x=x, y=y, width=width, height=height)
            self.editing_entry.insert(0, value)
            self.editing_entry.focus()

            def save_edit(event=None):
                new_val = self.editing_entry.get()
                old_val = self.set(row, header)
                if new_val != old_val:
                    self.edited_rows.add(row)
                    self.set(row, header, new_val if new_val != "" else "-")
                self.editing_entry.destroy()
                self.editing_entry = None

            self.editing_entry.bind("<Return>", save_edit)
            self.editing_entry.bind("<FocusOut>", save_edit)

        except Exception:
            messagebox.showerror(
                "Error en edición",
                "Ocurrió un error al editar celda:\\n{}"
                .format(traceback.format_exc())
            )


def main():
    # 1) Meta de planilla
    meta = cargar_json(PLANILLA_META_PATH, show_not_found=True, title="Error meta") or {}
    headers = meta.get("Headers", []) or []
    codigo_planilla = meta.get("CodigoPlanilla", "") or ""
    nombre_planilla = meta.get("NombrePlanilla", "Planilla")
    data_path = meta.get("DataPath", "") or ""

    if not headers or not codigo_planilla:
        messagebox.showerror(
            "Datos incompletos",
            "No se encontraron Headers o CodigoPlanilla en el meta."
        )
        return

    if not data_path:
        messagebox.showerror(
            "Datos incompletos",
            "No se encontró ruta de datos (DataPath) en el meta."
        )
        return

    # 2) Dataset combinado (BD + modelo) desde JSON temporal
    data_rows = cargar_json(data_path, show_not_found=False, title="Error datos") or []
    if not isinstance(data_rows, list):
        data_rows = []

    if not data_rows:
        messagebox.showinfo(
            "Sin datos",
            "No se encontraron filas para la planilla con código '{}' "
            "en el dataset combinado.".format(codigo_planilla)
        )
        return

    # 3) UI
    root = tk.Tk()
    root.title("Edición Planilla '{}'".format(nombre_planilla))
    root.geometry("1000x550")
    root.minsize(600, 350)

    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)

    frame = ttk.Frame(root)
    frame.grid(row=0, column=0, sticky="nsew")
    frame.columnconfigure(0, weight=1)
    frame.rowconfigure(0, weight=1)

    gui_headers = ["Archivo", "ElementId", "nombre_archivo", "CodIntBIM"] + [
        h for h in headers
        if h not in ("Archivo", "ElementId", "nombre_archivo", "CodIntBIM")
    ]

    tree = EditableTreeview(frame, gui_headers, data_rows)
    tree.grid(row=0, column=0, sticky="nsew")

    vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
    vsb.grid(row=0, column=1, sticky="ns")

    hsb = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
    hsb.grid(row=1, column=0, sticky="ew")

    tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

    sizegrip = ttk.Sizegrip(root)
    sizegrip.grid(row=2, column=0, sticky="se")

    # 4) Guardar SOLO elementos modificados en la BD real
    def guardar():
        try:
            repo_path = get_repo_path_from_config()
            if not repo_path:
                return

            repo_datos = cargar_json(repo_path, show_not_found=False, title="Error repo") or {}
            if not isinstance(repo_datos, dict):
                repo_datos = {}

            solo_cambios = {}
            filas_actualizadas = 0

            for iid in tree.get_children():
                if iid not in tree.edited_rows:
                    continue

                fila_gui = {h: tree.set(iid, h) for h in gui_headers}
                archivo = fila_gui.get("Archivo", "") or ""
                elem_id = fila_gui.get("ElementId", "") or ""
                if not archivo or not elem_id:
                    continue

                clave = "{}_{}".format(archivo, elem_id)
                original = repo_datos.get(clave, {})
                datos_completos = dict(original) if isinstance(original, dict) else {}

                datos_completos["Archivo"] = _view_to_repo(archivo)
                datos_completos["ElementId"] = _view_to_repo(elem_id)
                datos_completos["nombre_archivo"] = _view_to_repo(
                    fila_gui.get("nombre_archivo")
                    or (os.path.basename(archivo) if archivo else "")
                )
                datos_completos["CodIntBIM"] = _view_to_repo(
                    fila_gui.get("CodIntBIM", "")
                )

                for h in headers:
                    if h in ("Archivo", "ElementId", "nombre_archivo", "CodIntBIM"):
                        continue
                    datos_completos[h] = _view_to_repo(fila_gui.get(h, ""))

                if datos_completos == original:
                    continue

                solo_cambios[clave] = datos_completos
                filas_actualizadas += 1

            if filas_actualizadas == 0:
                messagebox.showinfo(
                    "Sin cambios",
                    "No se detectaron cambios para guardar en la base de datos."
                )
                return

            bd_actual = cargar_json(repo_path, show_not_found=False, title="Error repo") or {}
            if not isinstance(bd_actual, dict):
                bd_actual = {}

            bd_actual.update(solo_cambios)
            guardar_json(repo_path, bd_actual)

            messagebox.showinfo(
                "Guardado",
                "Se actualizaron {} elemento(s) en:\\n{}"
                .format(filas_actualizadas, repo_path)
            )
            root.destroy()

        except Exception:
            messagebox.showerror(
                "Error al guardar",
                "Error guardando datos:\\n{}".format(traceback.format_exc())
            )

    def cancelar():
        root.destroy()

    btn_frame = ttk.Frame(root)
    btn_frame.grid(row=1, column=0, sticky="ew", pady=5, padx=5)
    btn_frame.columnconfigure(0, weight=1)
    btn_frame.columnconfigure(1, weight=1)

    btn_guardar = ttk.Button(btn_frame, text="Guardar", command=guardar)
    btn_guardar.grid(row=0, column=0, sticky="ew", padx=5)

    btn_cancelar = ttk.Button(btn_frame, text="Cancelar", command=cancelar)
    btn_cancelar.grid(row=0, column=1, sticky="ew", padx=5)

    root.mainloop()


if __name__ == "__main__":
    main()
