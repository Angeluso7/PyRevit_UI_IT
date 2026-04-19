# -*- coding: utf-8 -*-
"""
Gestor de proyectos para pyRevit.

Uso:
    gestor_proyectos.py <MASTER_DIR> <DOCS_INFO_JSON>

DOCS_INFO_JSON tiene la forma:
{
  "activo": {"nombre": "...", "unique_id": "..."},
  "links":  [
      {"nombre": "...", "unique_id": "..."},
      ...
  ]
}

MASTER_DIR contiene:
- registro_proyectos.json      : indice de proyectos
- config_proyecto_activo.json  : proyecto seleccionado actualmente

Las BD por proyecto se guardan en:
- data/proyectos/repositorio_datos_<NUP>.json
"""

import sys
import os
import json
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime


# ----------------- Utilidades -----------------
def _now_iso():
    return datetime.now().strftime("%Y-%m-%dT%H:%M:%S")


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


def merge_documentos(existentes, nuevos):
    """
    Combina GUIDs de documentos sin borrar los de sesiones/equipos anteriores.
    Agrega o actualiza, pero NUNCA elimina claves que ya existian.
    """
    resultado = dict(existentes)
    for nombre, uid in nuevos.items():
        if uid:
            resultado[nombre] = uid
    return resultado


# ----------------- Ventana principal -----------------
class GestorProyectosApp:
    def __init__(self, root, master_dir, docs_info):
        self.root       = root
        self.master_dir = master_dir

        # data/ es el padre de master/
        self.data_dir = os.path.dirname(master_dir)

        # data/proyectos/ es donde viven las BD por proyecto
        self.proyectos_dir = os.path.join(self.data_dir, "proyectos")
        if not os.path.exists(self.proyectos_dir):
            os.makedirs(self.proyectos_dir)

        self.docs_info         = docs_info or {}
        self.doc_activo_uid    = (self.docs_info.get("activo") or {}).get("unique_id", "")
        self.doc_activo_nombre = (self.docs_info.get("activo") or {}).get("nombre", "")

        self.registro_path = os.path.join(self.master_dir, "registro_proyectos.json")
        self.config_path   = os.path.join(self.master_dir, "config_proyecto_activo.json")

        self.registro = cargar_json(self.registro_path, {})
        self.config   = cargar_json(self.config_path, {})

        self.proyecto_seleccionado = tk.StringVar()
        self.entry_nup             = None
        self.entry_nombre          = None
        self.label_activo_valor    = None
        self.frame_radios          = None

        self.modo_edicion  = False
        self.modo_eliminar = False

        self._construir_ui()
        self._preseleccionar_por_guid()
        self._refrescar_radios()
        self._refrescar_label_activo()

    # ---------- UI ----------
    def _construir_ui(self):
        self.root.title("Gestor de Proyectos")
        self.root.geometry("540x460")
        self.root.minsize(480, 400)

        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.grid(row=0, column=0, sticky="nsew")
        main_frame.columnconfigure(1, weight=1)

        ttk.Label(main_frame, text="Decreto:").grid(
            row=0, column=0, sticky="w", pady=(0, 5))
        self.entry_nup = ttk.Entry(main_frame)
        self.entry_nup.grid(row=0, column=1, sticky="ew", pady=(0, 5))

        ttk.Label(main_frame, text="Nombre de Proyecto:").grid(
            row=1, column=0, sticky="w", pady=(0, 5))
        self.entry_nombre = ttk.Entry(main_frame)
        self.entry_nombre.grid(row=1, column=1, sticky="ew", pady=(0, 5))

        ttk.Label(main_frame, text="Modelo abierto:").grid(
            row=2, column=0, sticky="w", pady=(0, 5))
        txt_modelo = self.doc_activo_nombre if self.doc_activo_nombre else "(desconocido)"
        ttk.Label(main_frame, text=txt_modelo, foreground="#555").grid(
            row=2, column=1, sticky="w", pady=(0, 5))

        ttk.Label(main_frame, text="Proyecto activo:").grid(
            row=3, column=0, sticky="w", pady=(0, 5))
        self.label_activo_valor = ttk.Label(main_frame, text="-")
        self.label_activo_valor.grid(row=3, column=1, sticky="w", pady=(0, 5))

        ttk.Separator(main_frame, orient="horizontal").grid(
            row=4, column=0, columnspan=2, sticky="ew", pady=10)

        ttk.Label(main_frame, text="Proyectos registrados:").grid(
            row=5, column=0, columnspan=2, sticky="w")

        self.frame_radios = ttk.Frame(main_frame)
        self.frame_radios.grid(row=6, column=0, columnspan=2, sticky="nsew", pady=(5, 5))
        main_frame.rowconfigure(6, weight=1)
        self.frame_radios.columnconfigure(0, weight=1)

        canvas    = tk.Canvas(self.frame_radios, borderwidth=0, highlightthickness=0)
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

        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=7, column=0, columnspan=2, sticky="e", pady=(10, 0))

        ttk.Button(btn_frame, text="Eliminar",  command=self.on_eliminar).grid(row=0, column=0, padx=(0, 5))
        ttk.Button(btn_frame, text="Editar",    command=self.on_editar).grid(row=0, column=1, padx=(0, 5))
        ttk.Button(btn_frame, text="Cancelar",  command=self.on_cancelar).grid(row=0, column=2, padx=(0, 5))
        ttk.Button(btn_frame, text="Aceptar",   command=self.on_aceptar).grid(row=0, column=3)

    # ---------- Preseleccion por GUID ----------
    def _preseleccionar_por_guid(self):
        if not self.doc_activo_uid:
            return
        for nup, info in self.registro.items():
            for uid in info.get("documentos", {}).values():
                if uid == self.doc_activo_uid:
                    self.proyecto_seleccionado.set(nup)
                    return

    # ---------- Refrescos ----------
    def _refrescar_radios(self):
        for child in self.inner_radios.winfo_children():
            child.destroy()

        activo = self.config.get("nup_activo", "")
        if activo and not self.proyecto_seleccionado.get():
            self.proyecto_seleccionado.set(activo)

        for i, nup in enumerate(sorted(self.registro.keys())):
            info   = self.registro.get(nup, {})
            nombre = info.get("nombre_proyecto", "")
            fecha  = info.get("fecha_modificacion", "")
            texto  = "{} - {}".format(nup, nombre) if nombre else nup
            if fecha:
                texto += "  [{}]".format(fecha[:10])
            ttk.Radiobutton(
                self.inner_radios,
                text=texto,
                variable=self.proyecto_seleccionado,
                value=nup
            ).grid(row=i, column=0, sticky="w", pady=2)

    def _refrescar_label_activo(self):
        nup_act    = self.config.get("nup_activo", "")
        nombre_act = self.config.get("nombre_proyecto", "")
        if nup_act:
            texto = "{} - {}".format(nup_act, nombre_act) if nombre_act else nup_act
        else:
            texto = "(ninguno)"
        self.label_activo_valor.config(text=texto)

    # ---------- Helper ruta repositorio ----------
    def _ruta_repositorio(self, nup):
        """Devuelve la ruta canonica en data/proyectos/ para un NUP dado."""
        return os.path.join(self.proyectos_dir, "repositorio_datos_{}.json".format(nup))

    # ---------- Botones ----------
    def on_editar(self):
        self.modo_eliminar = False
        nup_sel = self.proyecto_seleccionado.get()
        if not nup_sel:
            messagebox.showinfo("Informacion", "Selecciona primero un proyecto de la lista para editar.")
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
            messagebox.showinfo("Informacion", "Selecciona primero un proyecto de la lista para eliminar.")
            self.modo_eliminar = False
            return

        nup_activo = self.config.get("nup_activo", "")
        if nup_sel == nup_activo:
            messagebox.showwarning(
                "Aviso",
                "No se puede eliminar el proyecto que esta activo.\n"
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

    # ---------- Logica Aceptar ----------
    def on_aceptar(self):
        nup    = self.entry_nup.get().strip()
        nombre = self.entry_nombre.get().strip()

        # --- 1) Eliminar ---
        if self.modo_eliminar:
            nup_sel = self.proyecto_seleccionado.get()
            if not nup_sel:
                messagebox.showinfo("Informacion", "No hay proyecto seleccionado para eliminar.")
                return

            nup_activo = self.config.get("nup_activo", "")
            if nup_sel == nup_activo:
                messagebox.showwarning("Aviso", "No se puede eliminar el proyecto que esta activo.")
                self.modo_eliminar = False
                return

            info_sel = self.registro.get(nup_sel, {})
            confirmar = messagebox.askyesno(
                "Confirmar eliminacion",
                "Se eliminara el proyecto:\n\n"
                "  Decreto: {}\n"
                "  Nombre:  {}\n\n"
                "Esta accion no se puede deshacer. ¿Continuar?".format(
                    nup_sel,
                    info_sel.get("nombre_proyecto", "")
                )
            )
            if not confirmar:
                self.modo_eliminar = False
                return

            # Buscar el archivo en data/proyectos/ (ruta canonica)
            ruta_repo = self._ruta_repositorio(nup_sel)
            # Fallback: ruta guardada en el registro (puede ser ruta legacy)
            if not os.path.exists(ruta_repo):
                ruta_repo = info_sel.get("ruta_repositorio", "")
            if ruta_repo and os.path.exists(ruta_repo):
                try:
                    os.remove(ruta_repo)
                except Exception as e:
                    messagebox.showwarning(
                        "Aviso",
                        "No se pudo eliminar el archivo de datos:\n{}\n{}".format(ruta_repo, e)
                    )

            if nup_sel in self.registro:
                del self.registro[nup_sel]

            if not guardar_json(self.registro_path, self.registro):
                return

            if self.config.get("nup_activo", "") not in self.registro:
                self.config = {}
                guardar_json(self.config_path, self.config)

            self.modo_eliminar = False
            self.modo_edicion  = False
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

            # Ruta canonica en data/proyectos/
            if nup_original and nup_original in self.registro:
                # Edicion: si cambio el NUP, renombrar el archivo
                vieja_ruta = self._ruta_repositorio(nup_original)
                # Fallback a ruta legacy si el archivo canonico no existe
                if not os.path.exists(vieja_ruta):
                    vieja_ruta = self.registro[nup_original].get("ruta_repositorio", "")
                nueva_ruta = self._ruta_repositorio(nup)
                if vieja_ruta and os.path.exists(vieja_ruta) and vieja_ruta != nueva_ruta:
                    try:
                        os.rename(vieja_ruta, nueva_ruta)
                    except Exception:
                        nueva_ruta = vieja_ruta
                elif not os.path.exists(nueva_ruta):
                    guardar_json(nueva_ruta, {})
            else:
                # Proyecto nuevo: crear BD vacia en data/proyectos/
                nueva_ruta = self._ruta_repositorio(nup)
                if not os.path.exists(nueva_ruta):
                    guardar_json(nueva_ruta, {})

            # Merge de GUIDs
            existentes = {}
            if nup_original and nup_original in self.registro:
                existentes = self.registro[nup_original].get("documentos", {})
            elif nup in self.registro:
                existentes = self.registro[nup].get("documentos", {})

            nuevos_docs = {}
            if self.docs_info:
                activo = self.docs_info.get("activo") or {}
                if activo.get("unique_id"):
                    nuevos_docs[activo.get("nombre", "ModeloActivo")] = activo["unique_id"]
                for link in self.docs_info.get("links", []):
                    if link.get("unique_id"):
                        nuevos_docs[link.get("nombre", "Link")] = link["unique_id"]

            documentos_finales = merge_documentos(existentes, nuevos_docs)

            ahora = _now_iso()
            fecha_creacion = ahora
            if nup_original and nup_original in self.registro:
                fecha_creacion = self.registro[nup_original].get("fecha_creacion", ahora)
            elif nup in self.registro:
                fecha_creacion = self.registro[nup].get("fecha_creacion", ahora)

            self.registro[nup] = {
                "Decreto":            nup,
                "nombre_proyecto":    nombre,
                "ruta_repositorio":   nueva_ruta,
                "documentos":         documentos_finales,
                "fecha_creacion":     fecha_creacion,
                "fecha_modificacion": ahora,
            }

            if self.modo_edicion and nup_original and nup_original != nup:
                if nup_original in self.registro:
                    del self.registro[nup_original]

            if not guardar_json(self.registro_path, self.registro):
                return

            nup_activo = nup
        else:
            nup_activo = self.proyecto_seleccionado.get()

        # --- 3) Verificar GUID antes de activar ---
        if nup_activo and self.doc_activo_uid:
            info_check = self.registro.get(nup_activo, {})
            match = any(uid == self.doc_activo_uid for uid in info_check.get("documentos", {}).values())
            if not match:
                messagebox.showwarning(
                    "Advertencia",
                    "El proyecto '{}' no tiene registrado el modelo activo.\n"
                    "Verifica que estas trabajando en el proyecto correcto.".format(nup_activo)
                )

        # --- 4) Actualizar proyecto activo en config ---
        if nup_activo and nup_activo in self.registro:
            info = self.registro[nup_activo]
            cfg  = {
                "nup_activo":              nup_activo,
                "nombre_proyecto":         info.get("nombre_proyecto", ""),
                "ruta_repositorio_activo": info.get("ruta_repositorio", ""),
                "fecha_activacion":        _now_iso(),
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
            "Uso: gestor_proyectos.py <MASTER_DIR> [DOCS_INFO_JSON]"
        )
        return

    master_dir = sys.argv[1]
    docs_info  = {}
    if len(sys.argv) >= 3:
        try:
            docs_info = json.loads(sys.argv[2])
        except Exception:
            docs_info = {}

    if not os.path.exists(master_dir):
        try:
            os.makedirs(master_dir)
        except Exception as e:
            messagebox.showerror(
                "Error",
                "No se pudo crear la carpeta de datos:\n{}\n{}".format(master_dir, e)
            )
            return

    root = tk.Tk()
    app  = GestorProyectosApp(root, master_dir, docs_info)
    root.mainloop()


if __name__ == "__main__":
    main()
