# Script CPython para edición de parámetros por planilla
# Usa como repositorio principal la ruta indicada en ruta_repositorio_activo

import tkinter as tk
from tkinter import ttk, messagebox
import json
import os
import sys
import traceback

# Primer argumento: ruta del JSON temporal de planilla (solo metadatos)
if len(sys.argv) > 1:
    REPO_PLANILLA_PATH = sys.argv[1]
else:
    REPO_PLANILLA_PATH = os.path.join(
        os.path.dirname(__file__),
        "repo_planilla_tmp.json"
    )

# Carpeta común de datos
DATA_DIR = os.path.join(
    os.path.expanduser("~"),
    r"AppData\Roaming\MyPyRevitExtention\PyRevitIT.extension\data"
)

CONFIG_PATH = os.path.join(DATA_DIR, "config_proyecto_activo.json")
SCRIPT_JSON_PATH = os.path.join(DATA_DIR, "script.json")


def get_repo_datos_path_from_config():
    """Devuelve la ruta del repositorio activo desde config_proyecto_activo.json."""
    if not os.path.exists(CONFIG_PATH):
        messagebox.showerror(
            "Config no encontrada",
            "No se encontró config_proyecto_activo.json en:\n{}\n\n"
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
                "En config_proyecto_activo.json no se encontró la clave "
                "'ruta_repositorio_activo' o está vacía.\n\n"
                "No se puede determinar el repositorio de datos activo."
            )
            return None
        return ruta
    except Exception:
        messagebox.showerror(
            "Error config",
            "Error leyendo config_proyecto_activo.json:\n{}"
            .format(traceback.format_exc())
        )
        return None


REPOSITORIO_DATOS_PATH = get_repo_datos_path_from_config()


def cargar_json(ruta, show_not_found=True):
    try:
        if not ruta:
            return None
        if not os.path.exists(ruta):
            if show_not_found:
                messagebox.showinfo("Información", "Archivo no encontrado: {}".format(ruta))
            return None
        with open(ruta, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        messagebox.showerror(
            "Error",
            "Error cargando JSON:\n{}".format(traceback.format_exc())
        )
        return None


def guardar_json(ruta, datos):
    try:
        if not ruta:
            messagebox.showerror(
                "Error",
                "No se definió una ruta válida de repositorio para guardar."
            )
            return
        carpeta = os.path.dirname(ruta)
        if not os.path.exists(carpeta):
            os.makedirs(carpeta)
        with open(ruta, "w", encoding="utf-8") as f:
            json.dump(datos, f, indent=2, ensure_ascii=False)
    except Exception:
        messagebox.showerror(
            "Error",
            "Error guardando JSON:\n{}".format(traceback.format_exc())
        )


def obtener_nombre_planilla(codigo_planilla):
    script_json = cargar_json(SCRIPT_JSON_PATH, show_not_found=False) or {}
    codigos = script_json.get("codigos_planillas", {})
    nombres = script_json.get("reemplazos_de_nombres", {})
    nombre = None
    for key, code in codigos.items():
        if code == codigo_planilla:
            nombre = nombres.get(key)
            break
    return nombre or "Nombre desconocido"


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
            values = [row.get(h, "") for h in headers]
            self.insert("", "end", iid=str(i), values=values)

        self.bind("<Double-1>", self.on_double_click)
        self.editing_entry = None
        self.edited_rows = set()

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
            # No permitir editar campos de control si quieres (opcional):
            # if header in ("Archivo", "ElementId", "nombre_archivo"):
            #     return

            value = self.set(row, header)

            self.editing_entry = tk.Entry(self)
            self.editing_entry.place(x=x, y=y, width=width, height=height)
            self.editing_entry.insert(0, value)
            self.editing_entry.focus()

            def save_edit(event):
                new_val = self.editing_entry.get()
                old_val = self.set(row, header)
                if new_val != old_val:
                    self.edited_rows.add(row)
                    self.set(row, header, new_val)
                self.editing_entry.destroy()
                self.editing_entry = None

            def save_on_focus_out(event):
                if self.editing_entry:
                    new_val = self.editing_entry.get()
                    old_val = self.set(row, header)
                    if new_val != old_val:
                        self.edited_rows.add(row)
                        self.set(row, header, new_val)
                    self.editing_entry.destroy()
                    self.editing_entry = None

            self.editing_entry.bind("<Return>", save_edit)
            self.editing_entry.bind("<FocusOut>", save_on_focus_out)

        except Exception:
            messagebox.showerror(
                "Error en edición",
                "Ocurrió un error al editar celda:\n{}".format(
                    traceback.format_exc()
                ),
            )


def main():
    try:
        # 1) Metadatos desde repo_planilla_tmp.json: Headers + CodigoPlanilla
        meta = cargar_json(REPO_PLANILLA_PATH) or {}
        headers = meta.get("Headers", [])
        codigo_planilla = meta.get("CodigoPlanilla", "")

        if not headers or not codigo_planilla:
            messagebox.showerror(
                "Datos incompletos",
                "No se encontraron Headers o CodigoPlanilla en el archivo temporal."
            )
            return

        # 2) Base de datos principal
        repo_datos = cargar_json(REPOSITORIO_DATOS_PATH, show_not_found=False) or {}

        # 3) Construir lista de filas a mostrar desde la base de datos
        data_rows = []

        for clave, registro in repo_datos.items():
            try:
                archivo = registro.get("Archivo", "")
                elem_id = registro.get("ElementId", "")
                if not archivo or not elem_id:
                    continue

                # Filtro por planilla: CodIntBIM contenga el codigo_planilla
                codint = str(registro.get("CodIntBIM", "") or "")
                if codigo_planilla and codigo_planilla not in codint:
                    continue

                fila = {}
                fila["Archivo"] = archivo
                fila["ElementId"] = elem_id
                fila["nombre_archivo"] = registro.get(
                    "nombre_archivo", os.path.basename(archivo) if archivo else ""
                )

                # Respetar orden de headers para el resto
                for h in headers:
                    if h in ("Archivo", "ElementId", "nombre_archivo"):
                        # Ya gestionados explícitamente
                        continue
                    fila[h] = registro.get(h, "")

                data_rows.append(fila)
            except Exception:
                # Si alguna entrada falla, se ignora
                continue

        # Filtrado adicional robusto por CodIntBIM si quieres ordenar:
        data_rows.sort(
            key=lambda x: str(x.get("CodIntBIM", "") or "").lower()
        )

        if not data_rows:
            messagebox.showinfo(
                "Sin datos",
                "No se encontraron filas para la planilla con código '{}' en la base de datos."
                .format(codigo_planilla)
            )
            return

        nombre_planilla = obtener_nombre_planilla(codigo_planilla)

        # 4) Construir interfaz
        root = tk.Tk()
        root.title("Edición Planilla '{}'".format(nombre_planilla))
        root.geometry("900x500")
        root.minsize(400, 300)
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)

        frame = ttk.Frame(root)
        frame.grid(row=0, column=0, sticky="nsew")
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)

        tree = EditableTreeview(frame, headers, data_rows)
        tree.grid(row=0, column=0, sticky="nsew")

        vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        vsb.grid(row=0, column=1, sticky="ns")

        hsb = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        hsb.grid(row=1, column=0, sticky="ew")

        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        sizegrip = ttk.Sizegrip(root)
        sizegrip.grid(row=2, column=0, sticky="se")

        def guardar():
            try:
                if not REPOSITORIO_DATOS_PATH:
                    messagebox.showerror(
                        "Repositorio no definido",
                        "No se definió una ruta válida de repositorio en "
                        "config_proyecto_activo.json.",
                    )
                    return

                nuevos_datos = dict(repo_datos)  # copia superficial

                for iid in tree.get_children():
                    # Reconstruir fila en orden de headers
                    fila_gui = {h: tree.set(iid, h) for h in headers}

                    # Recuperar campos de control
                    archivo = fila_gui.get("Archivo", "") or ""
                    elem_id = fila_gui.get("ElementId", "") or ""

                    if not archivo or not elem_id:
                        continue

                    clave = "{}_{}".format(archivo, elem_id)

                    # Estructura completa que se guardará
                    datos_completos = dict(nuevos_datos.get(clave, {}))

                    # Forzar campos de control
                    datos_completos["Archivo"] = archivo
                    datos_completos["ElementId"] = elem_id
                    datos_completos["nombre_archivo"] = (
                        fila_gui.get("nombre_archivo")
                        or os.path.basename(archivo)
                        if archivo
                        else ""
                    )

                    # Copiar parámetros de headers (evitando duplicar campos de control)
                    for h in headers:
                        if h in ("Archivo", "ElementId", "nombre_archivo"):
                            continue
                        datos_completos[h] = fila_gui.get(h, "")

                    # Marcar opcionalmente id cuando hubo edición
                    if iid in tree.edited_rows:
                        datos_completos["id"] = elem_id

                    nuevos_datos[clave] = datos_completos

                guardar_json(REPOSITORIO_DATOS_PATH, nuevos_datos)

                messagebox.showinfo(
                    "Guardado",
                    "Datos guardados correctamente en el repositorio:\n{}"
                    .format(REPOSITORIO_DATOS_PATH),
                )
                root.destroy()
            except Exception:
                messagebox.showerror(
                    "Error al guardar",
                    "Error guardando datos:\n{}".format(
                        traceback.format_exc()
                    ),
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

    except Exception:
        messagebox.showerror(
            "Error general",
            "Ocurrió un error inesperado:\n{}".format(traceback.format_exc()),
        )


if __name__ == "__main__":
    main()
