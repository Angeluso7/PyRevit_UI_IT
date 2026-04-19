# -*- coding: utf-8 -*-
"""
Edicion de datos generales del proyecto activo (SAESA).

Uso:
    datos_proyecto.py <master_dir>

master_dir contiene:
- registro_proyectos.json     : indice de proyectos
- config_proyecto_activo.json : proyecto activo actual
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


# ----------------- Ventana principal -----------------
class DatosProyectoApp:
    PLACEHOLDER       = "AAAA/MM/DD"
    PLACEHOLDER_COLOR = "#888888"
    NORMAL_COLOR      = "#000000"

    def __init__(self, root, master_dir):
        self.root       = root
        self.master_dir = master_dir

        self.registro_path = os.path.join(self.master_dir, "registro_proyectos.json")
        self.config_path   = os.path.join(self.master_dir, "config_proyecto_activo.json")

        self.registro = cargar_json(self.registro_path, {})
        self.config   = cargar_json(self.config_path, {})

        self.nup_activo = self.config.get("nup_activo", "")

        # Guard: proyecto activo no definido
        if not self.nup_activo:
            messagebox.showwarning(
                "Aviso",
                "No hay proyecto activo definido.\n"
                "Configura un proyecto activo en el Coordinador primero."
            )

        # Guard: NUP activo no existe en el registro (proyecto fantasma)
        self.info_activo = self.registro.get(self.nup_activo, {})
        if self.nup_activo and not self.info_activo:
            messagebox.showerror(
                "Error",
                "El proyecto activo '{}' no existe en el registro.\n"
                "Es posible que haya sido eliminado.\n"
                "Vuelve al Coordinador y selecciona un proyecto valido.".format(self.nup_activo)
            )
            self.nup_activo = ""  # bloquear edicion

        # Mapa texto <-> codigo para tipo de proyecto
        self.map_text_to_code = {
            "Subestaciones": "CM01",
            "Patios":        "CM03",
            "Panos":         "CM06",
        }
        self.map_code_to_text = {v: k for k, v in self.map_text_to_code.items()}

        # Widgets
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

    # ---------- UI ----------
    def _construir_ui(self):
        self.root.title("Datos del Proyecto")
        self.root.geometry("650x560")
        self.root.minsize(600, 500)

        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        main = ttk.Frame(self.root, padding=10)
        main.grid(row=0, column=0, sticky="nsew")
        for i in range(3):
            main.columnconfigure(i, weight=1)

        fila = 0

        # Proyecto (solo lectura)
        ttk.Label(main, text="Proyecto:").grid(row=fila, column=0, sticky="w", pady=(0, 5))
        self.label_proyecto_val = ttk.Label(main, text="-")
        self.label_proyecto_val.grid(row=fila, column=1, columnspan=2, sticky="w", pady=(0, 5))
        fila += 1

        # Tipo
        ttk.Label(main, text="Tipo:").grid(row=fila, column=0, sticky="w", pady=(0, 5))
        self.combo_tipo = ttk.Combobox(main, state="readonly")
        self.combo_tipo["values"] = ("Subestaciones", "Patios", "Panos")
        self.combo_tipo.grid(row=fila, column=1, columnspan=2, sticky="ew", pady=(0, 5))
        fila += 1

        # Fecha entrada
        ttk.Label(main, text="Fecha de entrada en operacion:").grid(
            row=fila, column=0, sticky="w", pady=(0, 5))
        self.entry_fecha_ini = ttk.Entry(main)
        self.entry_fecha_ini.grid(row=fila, column=1, sticky="ew", pady=(0, 5))
        ttk.Button(main, text="Seleccionar", command=self._select_fecha_ini).grid(
            row=fila, column=2, sticky="w", padx=(5, 0))
        fila += 1

        # Fecha salida
        ttk.Label(main, text="Fecha salida de operacion:").grid(
            row=fila, column=0, sticky="w", pady=(0, 5))
        self.entry_fecha_fin = ttk.Entry(main)
        self.entry_fecha_fin.grid(row=fila, column=1, sticky="ew", pady=(0, 5))
        ttk.Button(main, text="Seleccionar", command=self._select_fecha_fin).grid(
            row=fila, column=2, sticky="w", padx=(5, 0))
        fila += 1

        # Propietario
        ttk.Label(main, text="Propietario:").grid(row=fila, column=0, sticky="w", pady=(0, 5))
        self.entry_propietario = ttk.Entry(main)
        self.entry_propietario.grid(row=fila, column=1, columnspan=2, sticky="ew", pady=(0, 5))
        fila += 1

        # Estado
        ttk.Label(main, text="Estado:").grid(row=fila, column=0, sticky="w", pady=(0, 5))
        self.combo_estado = ttk.Combobox(main, state="readonly")
        self.combo_estado["values"] = ("En Proyecto", "En Operacion", "Fuera de Operacion")
        self.combo_estado.grid(row=fila, column=1, columnspan=2, sticky="ew", pady=(0, 5))
        fila += 1

        # Comuna
        ttk.Label(main, text="Comuna:").grid(row=fila, column=0, sticky="w", pady=(0, 5))
        self.entry_comuna = ttk.Entry(main)
        self.entry_comuna.grid(row=fila, column=1, columnspan=2, sticky="ew", pady=(0, 5))
        fila += 1

        # Descripcion
        ttk.Label(main, text="Descripcion:").grid(row=fila, column=0, sticky="nw", pady=(0, 5))
        text_frame = ttk.Frame(main)
        text_frame.grid(row=fila, column=1, columnspan=2, sticky="nsew", pady=(0, 5))
        main.rowconfigure(fila, weight=1)
        text_frame.rowconfigure(0, weight=1)
        text_frame.columnconfigure(0, weight=1)
        self.text_descripcion = tk.Text(text_frame, height=4, wrap="word")
        self.text_descripcion.grid(row=0, column=0, sticky="nsew")
        scroll_desc = ttk.Scrollbar(text_frame, orient="vertical", command=self.text_descripcion.yview)
        scroll_desc.grid(row=0, column=1, sticky="ns")
        self.text_descripcion.configure(yscrollcommand=scroll_desc.set)
        fila += 1

        # Ultima modificacion SAESA (solo lectura, informativo)
        ttk.Label(main, text="Ultima modificacion:").grid(row=fila, column=0, sticky="w", pady=(0, 5))
        self.label_modificacion = ttk.Label(main, text="-", foreground="#555")
        self.label_modificacion.grid(row=fila, column=1, columnspan=2, sticky="w", pady=(0, 5))
        fila += 1

        # Botones
        btn_frame = ttk.Frame(main)
        btn_frame.grid(row=fila, column=0, columnspan=3, sticky="e", pady=(10, 0))
        ttk.Button(btn_frame, text="Editar",   command=self.on_editar).grid(row=0, column=0, padx=(0, 5))
        ttk.Button(btn_frame, text="Cancelar", command=self.on_cancelar).grid(row=0, column=1, padx=(0, 5))
        ttk.Button(btn_frame, text="Aceptar",  command=self.on_aceptar).grid(row=0, column=2)

        self._configurar_placeholders_fecha()

    # ---------- Placeholders ----------
    def _configurar_placeholders_fecha(self):
        for entry in (self.entry_fecha_ini, self.entry_fecha_fin):
            entry.bind("<FocusIn>",  self._on_focus_in_fecha)
            entry.bind("<FocusOut>", self._on_focus_out_fecha)

    def _on_focus_in_fecha(self, event):
        entry = event.widget
        if entry.cget("state") == "disabled":
            return
        if entry.get() == self.PLACEHOLDER and entry.cget("foreground") == self.PLACEHOLDER_COLOR:
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

    # ---------- Carga de datos ----------
    def _cargar_datos_en_campos(self):
        self.registro    = cargar_json(self.registro_path, {})
        self.info_activo = self.registro.get(self.nup_activo, {})

        if self.nup_activo:
            nom   = self.info_activo.get("nombre_proyecto", "")
            texto = "{} - {}".format(self.nup_activo, nom) if nom else self.nup_activo
        else:
            texto = "(ninguno)"
        self.label_proyecto_val.config(text=texto)

        # Tipo
        tipo_codigo = self.info_activo.get("tipo_codigo", "")
        texto_tipo  = self.map_code_to_text.get(tipo_codigo, "")
        self.combo_tipo.set(texto_tipo if texto_tipo in self.combo_tipo["values"] else "")

        # Fechas
        self._set_texto_fecha(self.entry_fecha_ini, self.info_activo.get("fecha_entrada", ""))
        self._set_texto_fecha(self.entry_fecha_fin,  self.info_activo.get("fecha_salida",  ""))

        # Propietario
        self.entry_propietario.config(state="normal")
        self.entry_propietario.delete(0, tk.END)
        self.entry_propietario.insert(0, self.info_activo.get("propietario", ""))

        # Estado
        estado_val = self.info_activo.get("estado", "").strip()
        self.combo_estado.set(estado_val if estado_val in self.combo_estado["values"] else "")

        # Comuna
        self.entry_comuna.config(state="normal")
        self.entry_comuna.delete(0, tk.END)
        self.entry_comuna.insert(0, self.info_activo.get("comuna", ""))

        # Descripcion
        self.text_descripcion.config(state="normal")
        self.text_descripcion.delete("1.0", tk.END)
        desc = self.info_activo.get("descripcion", "")
        if desc:
            self.text_descripcion.insert("1.0", desc)
        self.text_descripcion.config(state="disabled")

        # Ultima modificacion SAESA
        ult_mod = self.info_activo.get("fecha_modificacion_saesa", "")
        self.label_modificacion.config(text=ult_mod if ult_mod else "(sin modificaciones previas)")

    # ---------- Selectores de fecha ----------
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
        dlg.transient(self.root)
        dlg.grab_set()

        tk.Label(dlg, text="Anno:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        tk.Label(dlg, text="Mes:").grid(row=1,  column=0, padx=5, pady=5, sticky="e")
        tk.Label(dlg, text="Dia:").grid(row=2,  column=0, padx=5, pady=5, sticky="e")

        year_var  = tk.StringVar()
        month_var = tk.StringVar()
        day_var   = tk.StringVar()

        if valor_actual and valor_actual != self.PLACEHOLDER:
            try:
                y, m, d = valor_actual.split("/")
                year_var.set(y); month_var.set(m); day_var.set(d)
            except Exception:
                pass

        years  = [str(y) for y in range(1980, 2101)]
        months = ["{:02d}".format(m) for m in range(1, 13)]
        days   = ["{:02d}".format(d) for d in range(1, 32)]

        ttk.Combobox(dlg, textvariable=year_var,  values=years,  width=6, state="readonly").grid(row=0, column=1, padx=5, pady=5)
        ttk.Combobox(dlg, textvariable=month_var, values=months, width=4, state="readonly").grid(row=1, column=1, padx=5, pady=5)
        ttk.Combobox(dlg, textvariable=day_var,   values=days,   width=4, state="readonly").grid(row=2, column=1, padx=5, pady=5)

        resultado = {"fecha": None}

        def aceptar():
            y = year_var.get().strip()
            m = month_var.get().strip()
            d = day_var.get().strip()
            if not (y and m and d):
                messagebox.showwarning("Aviso", "Selecciona anno, mes y dia.")
                return
            try:
                datetime(int(y), int(m), int(d))
            except ValueError:
                messagebox.showerror("Error", "La fecha seleccionada no es valida.")
                return
            resultado["fecha"] = "{}/{}/{}".format(y, m, d)
            dlg.destroy()

        def cancelar():
            dlg.destroy()

        btn_f = ttk.Frame(dlg)
        btn_f.grid(row=3, column=0, columnspan=2, pady=10)
        ttk.Button(btn_f, text="Aceptar",  command=aceptar).grid(row=0, column=0, padx=5)
        ttk.Button(btn_f, text="Cancelar", command=cancelar).grid(row=0, column=1, padx=5)

        dlg.wait_window()
        return resultado["fecha"]

    # ---------- Modo edicion ----------
    def _set_campos_editables(self, editable):
        state = "normal" if editable else "disabled"
        self.combo_tipo.config(state="readonly" if editable else "disabled")
        self.combo_estado.config(state="readonly" if editable else "disabled")
        for entry in (self.entry_fecha_ini, self.entry_fecha_fin,
                      self.entry_propietario, self.entry_comuna):
            entry.config(state=state)
        self.text_descripcion.config(state="normal" if editable else "disabled")

    def on_editar(self):
        if not self.nup_activo:
            messagebox.showinfo(
                "Informacion",
                "No hay proyecto activo para editar.\n"
                "Configura primero un proyecto activo en el Coordinador."
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
                "El campo '{}' debe tener el formato AAAA/MM/DD y "
                "representar una fecha valida.".format(nombre_campo)
            )
            raise

    def on_aceptar(self):
        if not self.modo_edicion:
            self.root.destroy()
            return

        if not self.nup_activo:
            messagebox.showwarning("Aviso", "No hay proyecto activo. No se puede guardar.")
            return

        # Tipo
        combo_val  = self.combo_tipo.get().strip()
        tipo_texto  = combo_val
        tipo_codigo = self.map_text_to_code.get(combo_val, "") if combo_val else ""

        # Fechas con validacion de formato
        try:
            fecha_entrada = self._leer_fecha_entry(self.entry_fecha_ini, "Fecha de entrada en operacion")
            fecha_salida  = self._leer_fecha_entry(self.entry_fecha_fin,  "Fecha salida de operacion")
        except Exception:
            return

        # Validacion logica: fecha_entrada <= fecha_salida
        if fecha_entrada and fecha_salida:
            try:
                dt_ini = datetime.strptime(fecha_entrada, "%Y/%m/%d")
                dt_fin = datetime.strptime(fecha_salida,  "%Y/%m/%d")
                if dt_ini > dt_fin:
                    messagebox.showerror(
                        "Error de fechas",
                        "La fecha de entrada en operacion no puede ser posterior "
                        "a la fecha de salida.\n\n"
                        "  Entrada: {}\n"
                        "  Salida:  {}".format(fecha_entrada, fecha_salida)
                    )
                    return
            except Exception:
                pass

        propietario = self.entry_propietario.get().strip()
        estado      = self.combo_estado.get().strip()
        comuna      = self.entry_comuna.get().strip()
        descripcion = self.text_descripcion.get("1.0", tk.END).strip()

        # Recargar registro para no pisar cambios externos
        self.registro    = cargar_json(self.registro_path, {})
        info             = self.registro.get(self.nup_activo, {})

        # Preservar fecha_creacion si ya existe
        fecha_creacion = info.get("fecha_creacion", _now_iso())

        info["tipo_codigo"]              = tipo_codigo
        info["tipo_texto"]               = tipo_texto
        info["fecha_entrada"]            = fecha_entrada
        info["fecha_salida"]             = fecha_salida
        info["propietario"]              = propietario
        info["estado"]                   = estado
        info["comuna"]                   = comuna
        info["descripcion"]              = descripcion
        info["fecha_creacion"]           = fecha_creacion
        info["fecha_modificacion"]       = _now_iso()   # timestamp general del proyecto
        info["fecha_modificacion_saesa"] = _now_iso()   # timestamp especifico SAESA

        self.registro[self.nup_activo] = info

        if not guardar_json(self.registro_path, self.registro):
            return

        messagebox.showinfo("Informacion", "Datos del proyecto guardados correctamente.")
        self.root.destroy()


# ----------------- main -----------------
def main():
    if len(sys.argv) < 2:
        messagebox.showerror(
            "Error",
            "Uso: datos_proyecto.py <master_dir>"
        )
        return

    master_dir = sys.argv[1]
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
    app  = DatosProyectoApp(root, master_dir)
    root.mainloop()


if __name__ == "__main__":
    main()
