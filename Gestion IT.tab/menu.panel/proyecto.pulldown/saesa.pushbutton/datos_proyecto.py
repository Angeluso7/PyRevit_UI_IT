# -*- coding: utf-8 -*-
"""
datos_proyecto.py  —  Visor / editor de datos generales del proyecto activo.

Uso (lanzado por script.py via subprocess):
    pythonw.exe datos_proyecto.py <data_dir>

<data_dir>  ruta a data/master/  (provista por script.py usando MASTER_DIR
            de config_utils.py).  Si se omite, se resuelve automáticamente.
"""

import sys
import os
import json
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime


# ── Paleta dark ───────────────────────────────────────────────────────────────
DARK = {
    "bg":           "#1e1e1e",   # fondo ventana / frames
    "surface":      "#2a2a2a",   # fondo de entries, text, listbox
    "surface_alt":  "#252526",   # fondo alternativo (filas pares, etc.)
    "border":       "#3c3c3c",   # bordes sutiles
    "fg":           "#d4d4d4",   # texto principal
    "fg_muted":     "#888888",   # placeholder / texto apagado
    "accent":       "#4f98a3",   # teal primario (botón Aceptar / Editar)
    "accent_hover": "#3a7d88",
    "btn_bg":       "#3a3a3a",   # fondo botones secundarios
    "btn_fg":       "#d4d4d4",
    "disabled_bg":  "#2d2d2d",
    "disabled_fg":  "#5a5a5a",
    "select_bg":    "#094771",   # selección de texto
    "select_fg":    "#ffffff",
}


def _aplicar_tema_dark(root):
    """
    Configura ttk.Style y opciones de tk para un tema oscuro consistente.
    Debe llamarse ANTES de crear cualquier widget.
    """
    root.configure(bg=DARK["bg"])

    style = ttk.Style(root)
    # Usar 'clam' como base — es el más personalizable en Windows
    style.theme_use("clam")

    # ── Frame / LabelFrame ────────────────────────────────────────────────────
    style.configure("TFrame",      background=DARK["bg"])
    style.configure("TLabelframe", background=DARK["bg"],
                    foreground=DARK["fg"], bordercolor=DARK["border"])
    style.configure("TLabelframe.Label", background=DARK["bg"],
                    foreground=DARK["fg"])

    # ── Label ─────────────────────────────────────────────────────────────────
    style.configure("TLabel",
                    background=DARK["bg"],
                    foreground=DARK["fg"])

    # ── Entry ─────────────────────────────────────────────────────────────────
    style.configure("TEntry",
                    fieldbackground=DARK["surface"],
                    foreground=DARK["fg"],
                    insertcolor=DARK["fg"],
                    bordercolor=DARK["border"],
                    lightcolor=DARK["border"],
                    darkcolor=DARK["border"])
    style.map("TEntry",
              fieldbackground=[("disabled", DARK["disabled_bg"]),
                               ("readonly", DARK["disabled_bg"])],
              foreground=[("disabled", DARK["disabled_fg"]),
                          ("readonly", DARK["disabled_fg"])])

    # ── Combobox ──────────────────────────────────────────────────────────────
    style.configure("TCombobox",
                    fieldbackground=DARK["surface"],
                    background=DARK["btn_bg"],
                    foreground=DARK["fg"],
                    arrowcolor=DARK["fg"],
                    bordercolor=DARK["border"],
                    lightcolor=DARK["border"],
                    darkcolor=DARK["border"],
                    selectbackground=DARK["select_bg"],
                    selectforeground=DARK["select_fg"])
    style.map("TCombobox",
              fieldbackground=[("disabled", DARK["disabled_bg"]),
                               ("readonly", DARK["surface"])],
              foreground=[("disabled", DARK["disabled_fg"]),
                          ("readonly", DARK["fg"])],
              selectbackground=[("readonly", DARK["surface"])],
              selectforeground=[("readonly", DARK["fg"])])
    # Desplegable del combobox (tk.Listbox interno)
    root.option_add("*TCombobox*Listbox.background",  DARK["surface"])
    root.option_add("*TCombobox*Listbox.foreground",  DARK["fg"])
    root.option_add("*TCombobox*Listbox.selectBackground", DARK["select_bg"])
    root.option_add("*TCombobox*Listbox.selectForeground", DARK["select_fg"])

    # ── Button ────────────────────────────────────────────────────────────────
    style.configure("TButton",
                    background=DARK["btn_bg"],
                    foreground=DARK["btn_fg"],
                    bordercolor=DARK["border"],
                    lightcolor=DARK["border"],
                    darkcolor=DARK["border"],
                    relief="flat",
                    padding=(8, 4))
    style.map("TButton",
              background=[("active", DARK["border"]),
                          ("disabled", DARK["disabled_bg"])],
              foreground=[("disabled", DARK["disabled_fg"])])

    # Botón de acción primaria (Aceptar)
    style.configure("Accent.TButton",
                    background=DARK["accent"],
                    foreground="#ffffff",
                    bordercolor=DARK["accent"],
                    lightcolor=DARK["accent"],
                    darkcolor=DARK["accent"],
                    relief="flat",
                    padding=(8, 4))
    style.map("Accent.TButton",
              background=[("active", DARK["accent_hover"]),
                          ("disabled", DARK["disabled_bg"])],
              foreground=[("disabled", DARK["disabled_fg"])])

    # ── Scrollbar ─────────────────────────────────────────────────────────────
    style.configure("TScrollbar",
                    background=DARK["btn_bg"],
                    troughcolor=DARK["bg"],
                    bordercolor=DARK["border"],
                    arrowcolor=DARK["fg"],
                    relief="flat")
    style.map("TScrollbar",
              background=[("active", DARK["border"])])


# ── Resolución de data_dir ────────────────────────────────────────────────────
def _resolve_data_dir():
    """
    Devuelve la ruta a data/master/.
    Prioridad: argv[1] → fallback idéntico a config_utils.MASTER_DIR.
    """
    if len(sys.argv) >= 2:
        return sys.argv[1]
    return os.path.normpath(os.path.join(
        os.path.expanduser("~"),
        "AppData", "Roaming", "MyPyRevitExtention",
        "PyRevitIT.extension", "data", "master"
    ))


# ── Utilidades JSON ───────────────────────────────────────────────────────────
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


# ── Ventana principal ─────────────────────────────────────────────────────────
class DatosProyectoApp:
    PLACEHOLDER       = "AAAA/MM/DD"
    PLACEHOLDER_COLOR = DARK["fg_muted"]
    NORMAL_COLOR      = DARK["fg"]

    def __init__(self, root, data_dir):
        self.root     = root
        self.data_dir = data_dir

        self.registro_path = os.path.join(self.data_dir, "registro_proyectos.json")
        self.config_path   = os.path.join(self.data_dir, "config_proyecto_activo.json")

        self.registro   = cargar_json(self.registro_path, {})
        self.config     = cargar_json(self.config_path, {})
        self.nup_activo = self.config.get("nup_activo", "")

        if not self.nup_activo:
            messagebox.showwarning(
                "Aviso",
                "No hay proyecto activo definido.\nConfigura un proyecto activo primero."
            )

        self.info_activo = self.registro.get(self.nup_activo, {})

        self.map_text_to_code = {
            "Subestaciones": "CM01",
            "Patios":        "CM03",
            "Pa\u00f1os":   "CM06",
        }
        self.map_code_to_text = {v: k for k, v in self.map_text_to_code.items()}

        self.label_proyecto_val = None
        self.combo_tipo         = None
        self.entry_fecha_ini    = None
        self.entry_fecha_fin    = None
        self.entry_propietario  = None
        self.combo_estado       = None
        self.entry_comuna       = None
        self.text_descripcion   = None

        self.modo_edicion = False

        self._construir_ui()
        self._cargar_datos_en_campos()
        self._set_campos_editables(False)

    # ── UI ────────────────────────────────────────────────────────────────────
    def _construir_ui(self):
        self.root.title("Datos del Proyecto")
        self.root.geometry("650x540")
        self.root.minsize(600, 500)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        main = ttk.Frame(self.root, padding=10)
        main.grid(row=0, column=0, sticky="nsew")
        for i in range(3):
            main.columnconfigure(i, weight=1)

        fila = 0

        # 1) Proyecto (solo lectura)
        ttk.Label(main, text="Proyecto:").grid(row=fila, column=0, sticky="w", pady=(0, 5))
        self.label_proyecto_val = ttk.Label(main, text="-",
                                            foreground=DARK["accent"])
        self.label_proyecto_val.grid(row=fila, column=1, columnspan=2, sticky="w", pady=(0, 5))
        fila += 1

        # 2) Tipo
        ttk.Label(main, text="Tipo:").grid(row=fila, column=0, sticky="w", pady=(0, 5))
        self.combo_tipo = ttk.Combobox(main, state="readonly")
        self.combo_tipo["values"] = ("Subestaciones", "Patios", "Pa\u00f1os")
        self.combo_tipo.grid(row=fila, column=1, columnspan=2, sticky="ew", pady=(0, 5))
        fila += 1

        # 3) Fecha entrada
        ttk.Label(main, text="Fecha de entrada en operaci\u00f3n:").grid(
            row=fila, column=0, sticky="w", pady=(0, 5)
        )
        self.entry_fecha_ini = ttk.Entry(main)
        self.entry_fecha_ini.grid(row=fila, column=1, sticky="ew", pady=(0, 5))
        ttk.Button(main, text="Seleccionar", command=self._select_fecha_ini).grid(
            row=fila, column=2, sticky="w", padx=(5, 0)
        )
        fila += 1

        # 4) Fecha salida
        ttk.Label(main, text="Fecha salida de operaci\u00f3n:").grid(
            row=fila, column=0, sticky="w", pady=(0, 5)
        )
        self.entry_fecha_fin = ttk.Entry(main)
        self.entry_fecha_fin.grid(row=fila, column=1, sticky="ew", pady=(0, 5))
        ttk.Button(main, text="Seleccionar", command=self._select_fecha_fin).grid(
            row=fila, column=2, sticky="w", padx=(5, 0)
        )
        fila += 1

        # 5) Propietario
        ttk.Label(main, text="Propietario:").grid(row=fila, column=0, sticky="w", pady=(0, 5))
        self.entry_propietario = ttk.Entry(main)
        self.entry_propietario.grid(row=fila, column=1, columnspan=2, sticky="ew", pady=(0, 5))
        fila += 1

        # 6) Estado
        ttk.Label(main, text="Estado:").grid(row=fila, column=0, sticky="w", pady=(0, 5))
        self.combo_estado = ttk.Combobox(main, state="readonly")
        self.combo_estado["values"] = ("En Proyecto", "En Operaci\u00f3n", "Fuera de Operaci\u00f3n")
        self.combo_estado.grid(row=fila, column=1, columnspan=2, sticky="ew", pady=(0, 5))
        fila += 1

        # 7) Comuna
        ttk.Label(main, text="Comuna:").grid(row=fila, column=0, sticky="w", pady=(0, 5))
        self.entry_comuna = ttk.Entry(main)
        self.entry_comuna.grid(row=fila, column=1, columnspan=2, sticky="ew", pady=(0, 5))
        fila += 1

        # 8) Descripción
        ttk.Label(main, text="Descripci\u00f3n:").grid(row=fila, column=0, sticky="nw", pady=(0, 5))
        text_frame = ttk.Frame(main)
        text_frame.grid(row=fila, column=1, columnspan=2, sticky="nsew", pady=(0, 5))
        main.rowconfigure(fila, weight=1)
        text_frame.rowconfigure(0, weight=1)
        text_frame.columnconfigure(0, weight=1)
        self.text_descripcion = tk.Text(
            text_frame, height=4, wrap="word",
            bg=DARK["surface"], fg=DARK["fg"],
            insertbackground=DARK["fg"],
            selectbackground=DARK["select_bg"],
            selectforeground=DARK["select_fg"],
            relief="flat", borderwidth=1,
            highlightbackground=DARK["border"],
            highlightcolor=DARK["accent"],
            highlightthickness=1
        )
        self.text_descripcion.grid(row=0, column=0, sticky="nsew")
        scroll_desc = ttk.Scrollbar(text_frame, orient="vertical",
                                    command=self.text_descripcion.yview)
        scroll_desc.grid(row=0, column=1, sticky="ns")
        self.text_descripcion.configure(yscrollcommand=scroll_desc.set)
        fila += 1

        # Botones
        btn_frame = ttk.Frame(main)
        btn_frame.grid(row=fila, column=0, columnspan=3, sticky="e", pady=(10, 0))
        ttk.Button(btn_frame, text="Editar",   command=self.on_editar).grid(
            row=0, column=0, padx=(0, 5))
        ttk.Button(btn_frame, text="Cancelar", command=self.on_cancelar).grid(
            row=0, column=1, padx=(0, 5))
        ttk.Button(btn_frame, text="Aceptar",  command=self.on_aceptar,
                   style="Accent.TButton").grid(row=0, column=2)

        self._configurar_placeholders_fecha()

    # ── Placeholders de fecha ─────────────────────────────────────────────────
    def _configurar_placeholders_fecha(self):
        for entry in (self.entry_fecha_ini, self.entry_fecha_fin):
            entry.bind("<FocusIn>",  self._on_focus_in_fecha)
            entry.bind("<FocusOut>", self._on_focus_out_fecha)

    def _on_focus_in_fecha(self, event):
        entry = event.widget
        if entry.cget("state") == "disabled":
            return
        if (entry.get() == self.PLACEHOLDER and
                entry.cget("foreground") == self.PLACEHOLDER_COLOR):
            entry.delete(0, tk.END)
            entry.config(foreground=self.NORMAL_COLOR)

    def _on_focus_out_fecha(self, event):
        entry = event.widget
        if entry.cget("state") == "disabled":
            return
        if not entry.get().strip():
            entry.delete(0, tk.END)
            entry.insert(0, self.PLACEHOLDER)
            entry.config(foreground=self.PLACEHOLDER_COLOR)

    def _set_texto_fecha(self, entry, valor):
        entry.config(state="normal")
        entry.delete(0, tk.END)
        if valor:
            entry.insert(0, valor)
            entry.config(foreground=self.NORMAL_COLOR)
        else:
            entry.insert(0, self.PLACEHOLDER)
            entry.config(foreground=self.PLACEHOLDER_COLOR)

    # ── Carga de datos ────────────────────────────────────────────────────────
    def _cargar_datos_en_campos(self):
        self.registro    = cargar_json(self.registro_path, {})
        self.info_activo = self.registro.get(self.nup_activo, {})

        if self.nup_activo:
            nom   = self.info_activo.get("nombre_proyecto", "")
            texto = "{} - {}".format(self.nup_activo, nom) if nom else self.nup_activo
        else:
            texto = "(ninguno)"
        self.label_proyecto_val.config(text=texto)

        tipo_codigo = self.info_activo.get("tipo_codigo", None)
        if tipo_codigo:
            texto_tipo = self.map_code_to_text.get(tipo_codigo, "")
            self.combo_tipo.set(texto_tipo if texto_tipo in self.combo_tipo["values"] else "")
        else:
            self.combo_tipo.set("")

        self._set_texto_fecha(self.entry_fecha_ini, self.info_activo.get("fecha_entrada", ""))
        self._set_texto_fecha(self.entry_fecha_fin,  self.info_activo.get("fecha_salida",  ""))

        self.entry_propietario.config(state="normal")
        self.entry_propietario.delete(0, tk.END)
        self.entry_propietario.insert(0, self.info_activo.get("propietario", ""))

        estado_val = self.info_activo.get("estado", "").strip()
        self.combo_estado.set(estado_val if estado_val in self.combo_estado["values"] else "")

        self.entry_comuna.config(state="normal")
        self.entry_comuna.delete(0, tk.END)
        self.entry_comuna.insert(0, self.info_activo.get("comuna", ""))

        self.text_descripcion.config(state="normal")
        self.text_descripcion.delete("1.0", tk.END)
        desc = self.info_activo.get("descripcion", "")
        if desc:
            self.text_descripcion.insert("1.0", desc)
        self.text_descripcion.config(state="disabled")

    # ── Selectores de fecha ───────────────────────────────────────────────────
    def _select_fecha_ini(self):
        if self.entry_fecha_ini.cget("state") == "disabled":
            return
        fecha = self._dialogo_fecha(self.entry_fecha_ini.get().strip())
        if fecha:
            self._set_texto_fecha(self.entry_fecha_ini, fecha)

    def _select_fecha_fin(self):
        if self.entry_fecha_fin.cget("state") == "disabled":
            return
        fecha = self._dialogo_fecha(self.entry_fecha_fin.get().strip())
        if fecha:
            self._set_texto_fecha(self.entry_fecha_fin, fecha)

    def _dialogo_fecha(self, valor_actual):
        dlg = tk.Toplevel(self.root)
        dlg.title("Seleccionar fecha")
        dlg.configure(bg=DARK["bg"])
        dlg.transient(self.root)
        dlg.grab_set()

        tk.Label(dlg, text="A\u00f1o:", bg=DARK["bg"], fg=DARK["fg"]).grid(
            row=0, column=0, padx=5, pady=5, sticky="e")
        tk.Label(dlg, text="Mes:",  bg=DARK["bg"], fg=DARK["fg"]).grid(
            row=1, column=0, padx=5, pady=5, sticky="e")
        tk.Label(dlg, text="D\u00eda:", bg=DARK["bg"], fg=DARK["fg"]).grid(
            row=2, column=0, padx=5, pady=5, sticky="e")

        year_var  = tk.StringVar()
        month_var = tk.StringVar()
        day_var   = tk.StringVar()

        if valor_actual and valor_actual != self.PLACEHOLDER:
            try:
                y, m, d = valor_actual.split("/")
                year_var.set(y)
                month_var.set(m)
                day_var.set(d)
            except Exception:
                pass

        years  = [str(y) for y in range(1980, 2101)]
        months = ["{:02d}".format(m) for m in range(1, 13)]
        days   = ["{:02d}".format(d) for d in range(1, 32)]

        cb_year  = ttk.Combobox(dlg, textvariable=year_var,  values=years,  width=6, state="readonly")
        cb_month = ttk.Combobox(dlg, textvariable=month_var, values=months, width=4, state="readonly")
        cb_day   = ttk.Combobox(dlg, textvariable=day_var,   values=days,   width=4, state="readonly")

        cb_year.grid( row=0, column=1, padx=5, pady=5)
        cb_month.grid(row=1, column=1, padx=5, pady=5)
        cb_day.grid(  row=2, column=1, padx=5, pady=5)

        resultado = {"fecha": None}

        def aceptar():
            y = year_var.get().strip()
            m = month_var.get().strip()
            d = day_var.get().strip()
            if not (y and m and d):
                messagebox.showwarning("Aviso", "Selecciona a\u00f1o, mes y d\u00eda.")
                return
            try:
                datetime(int(y), int(m), int(d))
            except ValueError:
                messagebox.showerror("Error", "La fecha seleccionada no es v\u00e1lida.")
                return
            resultado["fecha"] = "{}/{}/{}".format(y, m, d)
            dlg.destroy()

        def cancelar():
            resultado["fecha"] = None
            dlg.destroy()

        btn_frame = ttk.Frame(dlg)
        btn_frame.grid(row=3, column=0, columnspan=2, pady=10)
        ttk.Button(btn_frame, text="Aceptar",  command=aceptar,
                   style="Accent.TButton").grid(row=0, column=0, padx=5)
        ttk.Button(btn_frame, text="Cancelar", command=cancelar).grid(row=0, column=1, padx=5)

        dlg.wait_window()
        return resultado["fecha"]

    # ── Modo edición ──────────────────────────────────────────────────────────
    def _set_campos_editables(self, editable):
        state = "normal" if editable else "disabled"
        self.combo_tipo.config(state="readonly" if editable else "disabled")
        self.combo_estado.config(state="readonly" if editable else "disabled")
        for entry in (self.entry_fecha_ini, self.entry_fecha_fin,
                      self.entry_propietario, self.entry_comuna):
            entry.config(state=state)
        self.text_descripcion.config(
            state="normal" if editable else "disabled",
            bg=DARK["surface"] if editable else DARK["disabled_bg"],
            fg=DARK["fg"]      if editable else DARK["disabled_fg"]
        )

    def on_editar(self):
        if not self.nup_activo:
            messagebox.showinfo(
                "Informaci\u00f3n",
                "No hay proyecto activo para editar.\nConfigura primero un proyecto activo."
            )
            return
        self.modo_edicion = True
        self._set_campos_editables(True)

    def on_cancelar(self):
        self.root.destroy()

    def _leer_fecha_entry(self, entry, nombre_campo):
        texto = entry.get().strip()
        if texto == "" or texto == self.PLACEHOLDER:
            return ""
        try:
            dt = datetime.strptime(texto, "%Y/%m/%d")
            return dt.strftime("%Y/%m/%d")
        except ValueError:
            messagebox.showerror(
                "Error",
                "El campo '{}' debe tener el formato AAAA/MM/DD y representar una fecha v\u00e1lida.".format(
                    nombre_campo
                )
            )
            raise

    def on_aceptar(self):
        if not self.modo_edicion:
            self.root.destroy()
            return

        if not self.nup_activo:
            messagebox.showwarning(
                "Aviso",
                "No hay proyecto activo definido.\nNo se puede guardar."
            )
            return

        combo_val   = self.combo_tipo.get().strip()
        tipo_texto  = combo_val
        tipo_codigo = self.map_text_to_code.get(combo_val, "") if combo_val else ""

        try:
            fecha_entrada = self._leer_fecha_entry(self.entry_fecha_ini, "Fecha de entrada en operaci\u00f3n")
            fecha_salida  = self._leer_fecha_entry(self.entry_fecha_fin,  "Fecha salida de operaci\u00f3n")
        except Exception:
            return

        propietario = self.entry_propietario.get().strip()
        estado      = self.combo_estado.get().strip()
        comuna      = self.entry_comuna.get().strip()
        descripcion = self.text_descripcion.get("1.0", tk.END).strip()

        self.registro = cargar_json(self.registro_path, {})
        info = self.registro.get(self.nup_activo, {})
        info["tipo_codigo"]   = tipo_codigo
        info["tipo_texto"]    = tipo_texto
        info["fecha_entrada"] = fecha_entrada
        info["fecha_salida"]  = fecha_salida
        info["propietario"]   = propietario
        info["estado"]        = estado
        info["comuna"]        = comuna
        info["descripcion"]   = descripcion
        self.registro[self.nup_activo] = info

        if not guardar_json(self.registro_path, self.registro):
            return

        messagebox.showinfo("Informaci\u00f3n", "Datos del proyecto guardados correctamente.")
        self.root.destroy()


# ── Entry point ───────────────────────────────────────────────────────────────
def main():
    data_dir = _resolve_data_dir()

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
    _aplicar_tema_dark(root)       # ← tema dark aplicado antes de construir la UI
    DatosProyectoApp(root, data_dir)
    root.mainloop()


if __name__ == "__main__":
    main()
