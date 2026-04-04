# -*- coding: utf-8 -*-
"""
Gestor de proyectos para pyRevit.

Uso:
    gestor_proyectos.py <DATA_DIR> <DOCS_INFO_JSON>

DOCS_INFO_JSON tiene la forma:
{
  "activo": {"nombre": "...", "unique_id": "...", "path": "..."},
  "links":  [
      {"nombre": "...", "unique_id": "...", "path": "..."},
      ...
  ]
}

DATA_DIR contiene:
- registro_proyectos.json : índice de proyectos
- config_proyecto_activo.json : proyecto seleccionado
- repositorio_datos_<NUP>.json : base de datos por proyecto
"""

import sys
import os
import json
import tkinter as tk
from tkinter import ttk, messagebox


# ----------------- Utilidades JSON -----------------
def cargar_json(ruta, default):
    try:
        if not os.path.exists(ruta):
            return default
        with open(ruta, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return default


def guardar_json(ruta, data):
    try:
        os.makedirs(os.path.dirname(ruta), exist_ok=True)
        with open(ruta, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        messagebox.showerror("Error", "No se pudo guardar JSON:\n{}".format(e))
        return False


# ----------------- Ventana principal -----------------
class GestorProyectosApp:
    def __init__(self, root, data_dir, docs_info):
        self.root = root
        self.data_dir = data_dir

        self.docs_info = docs_info or {}
        self.doc_activo_uid = (self.docs_info.get("activo") or {}).get("unique_id", "")
        self.doc_activo_nombre = (self.docs_info.get("activo") or {}).get("nombre", "")

        self.registro_path = os.path.join(self.data_dir, "registro_proyectos.json")
        self.config_path = os.path.join(self.data_dir, "config_proyecto_activo.json")

        # Cargar datos iniciales
        self.registro = cargar_json(self.registro_path, {})
        self.config = cargar_json(self.config_path, {})

        self.proyecto_seleccionado = tk.StringVar()  # para radio buttons
        self.entry_nup = None
        self.entry_nombre = None
        self.label_activo_valor = None
        self.frame_radios = None

        self.modo_edicion = False       # True cuando se pulsa "Editar"
        self.modo_eliminar = False      # True cuando se pulsa "Eliminar"

        self._construir_ui()
        self._preseleccionar_por_guid()
        self._refrescar_radios()
        self._refrescar_label_activo()

    # ---------- UI ----------
    def _construir_ui(self):
        self.root.title("Gestor de Proyectos")
        self.root.geometry("520x430")
        self.root.minsize(480, 380)

        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.grid(row=0, column=0, sticky="nsew")
        main_frame.columnconfigure(1, weight=1)

        # Fila 0: Proyecto (NUP)
        lbl_nup = ttk.Label(main_frame, text="Decreto:")
        lbl_nup.grid(row=0, column=0, sticky="w", pady=(0, 5))

        self.entry_nup = ttk.Entry(main_frame)
        self.entry_nup.grid(row=0, column=1, sticky="ew", pady=(0, 5))

        # Fila 1: Nombre de Proyecto
        lbl_nombre = ttk.Label(main_frame, text="Nombre de Proyecto:")
        lbl_nombre.grid(row=1, column=0, sticky="w", pady=(0, 5))

        self.entry_nombre = ttk.Entry(main_frame)
        self.entry_nombre.grid(row=1, column=1, sticky="ew", pady=(0, 5))

        # Fila 2: Proyecto activo (solo lectura)
        lbl_activo = ttk.Label(main_frame, text="Proyecto activo:")
        lbl_activo.grid(row=2, column=0, sticky="w", pady=(0, 5))

        self.label_activo_valor = ttk.Label(main_frame, text="-")
        self.label_activo_valor.grid(row=2, column=1, sticky="w", pady=(0, 5))

        # Fila 3: separador
        sep = ttk.Separator(main_frame, orient="horizontal")
        sep.grid(row=3, column=0, columnspan=2, sticky="ew", pady=10)

        # Fila 4+: proyectos registrados (radio buttons)
        lbl_lista = ttk.Label(main_frame, text="Proyectos registrados:")
        lbl_lista.grid(row=4, column=0, columnspan=2, sticky="w")

        self.frame_radios = ttk.Frame(main_frame)
        self.frame_radios.grid(row=5, column=0, columnspan=2, sticky="nsew", pady=(5, 5))
        main_frame.rowconfigure(5, weight=1)
        self.frame_radios.columnconfigure(0, weight=1)

        # Scroll para radios
        canvas = tk.Canvas(self.frame_radios, borderwidth=0, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.frame_radios, orient="vertical", command=canvas.yview)
        self.inner_radios = ttk.Frame(canvas)

        self.inner_radios.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=self.inner_radios, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.frame_radios.rowconfigure(0, weight=1)
        self.frame_radios.columnconfigure(0, weight=1)

        # Botones inferiores
        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=6, column=0, columnspan=2, sticky="e", pady=(10, 0))

        btn_eliminar = ttk.Button(btn_frame, text="Eliminar", command=self.on_eliminar)
        btn_editar = ttk.Button(btn_frame, text="Editar", command=self.on_editar)
        btn_cancelar = ttk.Button(btn_frame, text="Cancelar", command=self.on_cancelar)
        btn_aceptar = ttk.Button(btn_frame, text="Aceptar", command=self.on_aceptar)

        btn_eliminar.grid(row=0, column=0, padx=(0, 5))
        btn_editar.grid(row=0, column=1, padx=(0, 5))
        btn_cancelar.grid(row=0, column=2, padx=(0, 5))
        btn_aceptar.grid(row=0, column=3)

    # ---------- Preselección por GUID ----------
    def _preseleccionar_por_guid(self):
        """
        Si el GUID del archivo activo ya existe dentro de algún proyecto (en su diccionario 'documentos'),
        se preselecciona ese NUP en la variable de radio buttons.
        """
        if not self.doc_activo_uid:
            return

        for nup, info in self.registro.items():
            documentos = info.get("documentos", {})
            for _, uid in documentos.items():
                if uid == self.doc_activo_uid:
                    self.proyecto_seleccionado.set(nup)
                    return

    # ---------- Refrescos UI ----------
    def _refrescar_radios(self):
        for child in self.inner_radios.winfo_children():
            child.destroy()

        claves_ordenadas = sorted(self.registro.keys())

        activo = self.config.get("nup_activo", "")
        if activo and not self.proyecto_seleccionado.get():
            self.proyecto_seleccionado.set(activo)

        for i, nup in enumerate(claves_ordenadas):
            info = self.registro.get(nup, {})
            nombre = info.get("nombre_proyecto", "")
            texto = "{} - {}".format(nup, nombre) if nombre else nup
            rb = ttk.Radiobutton(
                self.inner_radios,
                text=texto,
                variable=self.proyecto_seleccionado,
                value=nup
            )
            rb.grid(row=i, column=0, sticky="w", pady=2)

    def _refrescar_label_activo(self):
        nup_act = self.config.get("nup_activo", "")
        nombre_act = self.config.get("nombre_proyecto", "")
        if nup_act:
            texto = "{} - {}".format(nup_act, nombre_act) if nombre_act else nup_act
        else:
            texto = "(ninguno)"
        self.label_activo_valor.config(text=texto)

    # ---------- Botones ----------
    def on_editar(self):
        self.modo_eliminar = False
        nup_sel = self.proyecto_seleccionado.get()
        if not nup_sel:
            messagebox.showinfo("Información", "Selecciona primero un proyecto de la lista para editar.")
            return

        info = self.registro.get(nup_sel, {})
        self.entry_nup.delete(0, tk.END)
        self.entry_nup.insert(0, info.get("Decreto", ""))

        self.entry_nombre.delete(0, tk.END)
        self.entry_nombre.insert(0, info.get("nombre_proyecto", ""))

        self.modo_edicion = True

    def on_eliminar(self):
        self.modo_edicion = False
        nup_sel = self.proyecto_seleccionado.get()
        if not nup_sel:
            messagebox.showinfo("Información", "Selecciona primero un proyecto de la lista para eliminar.")
            self.modo_eliminar = False
            return

        nup_activo = self.config.get("nup_activo", "")
        if nup_sel == nup_activo:
            messagebox.showwarning(
                "Aviso",
                "No se puede eliminar el proyecto que está activo.\n"
                "Selecciona otro proyecto o cambia el activo antes de eliminar."
            )
            self.modo_eliminar = False
            return

        self.modo_eliminar = True
        info = self.registro.get(nup_sel, {})
        self.entry_nup.delete(0, tk.END)
        self.entry_nup.insert(0, info.get("nup", nup_sel))
        self.entry_nombre.delete(0, tk.END)
        self.entry_nombre.insert(0, info.get("nombre_proyecto", ""))

    def on_cancelar(self):
        self.root.destroy()

    # ---------- Lógica principal Aceptar ----------
    def on_aceptar(self):
        nup = self.entry_nup.get().strip()
        nombre = self.entry_nombre.get().strip()

        # --- 1) Eliminar si corresponde ---
        if self.modo_eliminar:
            nup_sel = self.proyecto_seleccionado.get()
            if not nup_sel:
                messagebox.showinfo("Información", "No hay proyecto seleccionado para eliminar.")
                return

            nup_activo = self.config.get("nup_activo", "")
            if nup_sel == nup_activo:
                messagebox.showwarning(
                    "Aviso",
                    "No se puede eliminar el proyecto que está activo."
                )
                self.modo_eliminar = False
                return

            info = self.registro.get(nup_sel, {})
            ruta_repo = info.get("ruta_repositorio", "")

            if ruta_repo and os.path.exists(ruta_repo):
                try:
                    os.remove(ruta_repo)
                except Exception:
                    pass

            if nup_sel in self.registro:
                del self.registro[nup_sel]

            if not guardar_json(self.registro_path, self.registro):
                return

            nup_activo = self.config.get("nup_activo", "")
            if nup_activo and nup_activo not in self.registro:
                self.config = {}
                guardar_json(self.config_path, self.config)

            self.modo_eliminar = False
            self.modo_edicion = False
            self.proyecto_seleccionado.set("")
            self.entry_nup.delete(0, tk.END)
            self.entry_nombre.delete(0, tk.END)

            self._refrescar_radios()
            self._refrescar_label_activo()
            return

        # --- 2) Crear o actualizar proyecto ---
        if nup:
            nup_original = None
            if self.modo_edicion and self.proyecto_seleccionado.get():
                nup_original = self.proyecto_seleccionado.get()

            # Determinar ruta de repositorio del proyecto
            if nup_original and nup_original in self.registro:
                vieja_info = self.registro[nup_original]
                vieja_ruta = vieja_info.get("ruta_repositorio", "")
                nueva_ruta = vieja_ruta
                if vieja_ruta and os.path.exists(vieja_ruta):
                    base_dir = os.path.dirname(vieja_ruta)
                    nueva_ruta = os.path.join(base_dir, "repositorio_datos_{}.json".format(nup))
                    if vieja_ruta != nueva_ruta:
                        try:
                            os.rename(vieja_ruta, nueva_ruta)
                        except Exception:
                            nueva_ruta = vieja_ruta
            else:
                nueva_ruta = os.path.join(self.data_dir, "repositorio_datos_{}.json".format(nup))
                if not os.path.exists(nueva_ruta):
                    guardar_json(nueva_ruta, {})

            # Construir diccionario de documentos (modelo activo + links)
            documentos = {}
            if self.docs_info:
                activo = self.docs_info.get("activo") or {}
                if activo.get("unique_id"):
                    documentos[activo.get("nombre", "ModeloActivo")] = activo["unique_id"]

                for link in self.docs_info.get("links", []):
                    if link.get("unique_id"):
                        documentos[link.get("nombre", "Link")] = link["unique_id"]

            # Registrar/actualizar proyecto
            self.registro[nup] = {
                "Decreto": nup,
                "nombre_proyecto": nombre,
                "ruta_repositorio": nueva_ruta,
                "documentos": documentos,
            }

            # Si cambió NUP en modo edición, eliminar la clave vieja
            if self.modo_edicion and self.proyecto_seleccionado.get() and self.proyecto_seleccionado.get() != nup:
                viejo_nup = self.proyecto_seleccionado.get()
                if viejo_nup in self.registro:
                    del self.registro[viejo_nup]

            if not guardar_json(self.registro_path, self.registro):
                return

            # Nuevo proyecto pasa a ser activo por defecto
            nup_activo = nup
        else:
            # Si no se indicó NUP, usar radio seleccionado como activo
            nup_activo = self.proyecto_seleccionado.get()

        # --- 3) Verificar GUID antes de activar ---
        if nup_activo and self.doc_activo_uid:
            info_check = self.registro.get(nup_activo, {})
            documentos = info_check.get("documentos", {})
            match = any(uid == self.doc_activo_uid for uid in documentos.values())
            if not match:
                messagebox.showwarning(
                    "Advertencia",
                    "El proyecto seleccionado no corresponde al archivo activo."
                )

        # --- 4) Actualizar proyecto activo en config ---
        if nup_activo and nup_activo in self.registro:
            info = self.registro[nup_activo]
            cfg = {
                "nup_activo": nup_activo,
                "nombre_proyecto": info.get("nombre_proyecto", ""),
                "ruta_repositorio_activo": info.get("ruta_repositorio", "")
            }
            if not guardar_json(self.config_path, cfg):
                return
            self.config = cfg
        else:
            self.config = {}
            guardar_json(self.config_path, self.config)

        self.root.destroy()


# ----------------- main -----------------
def main():
    if len(sys.argv) < 2:
        messagebox.showerror(
            "Error",
            "Uso: gestor_proyectos.py <DATA_DIR> [DOCS_INFO_JSON]"
        )
        return

    data_dir = sys.argv[1]
    docs_info = {}
    if len(sys.argv) >= 3:
        try:
            docs_info = json.loads(sys.argv[2])
        except Exception:
            docs_info = {}

    if not os.path.exists(data_dir):
        try:
            os.makedirs(data_dir)
        except Exception as e:
            messagebox.showerror(
                "Error",
                "No se pudo crear la carpeta de datos:\n{}\n{}".format(data_dir, e)
            )
            return

    root = tk.Tk()
    app = GestorProyectosApp(root, data_dir, docs_info)
    root.mainloop()


if __name__ == "__main__":
    main()
