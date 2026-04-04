# -*- coding: utf-8 -*-
import sys
import os
import json
import tkinter as tk
from tkinter import ttk, messagebox


def cargar_cache(path_json):
    with open(path_json, "r", encoding="utf-8") as f:
        return json.load(f)


class CodIntBIMSelector(tk.Tk):
    def __init__(self, cache_data):
        super(CodIntBIMSelector, self).__init__()

        self.title("Selección por CodIntBIM")
        self.geometry("600x500")
        self.minsize(600, 400)

        self.cache_data = cache_data or {}
        esp_dict = self.cache_data.get("especialidad", {})

        # solo especialidades que tienen al menos un código de elemento
        self.especialidades = sorted([
            esp for esp, data in esp_dict.items()
            if data.get("codigos_elementos")
        ])

        self.selected_especialidad = None
        self.selected_codigo = None

        self._build_ui()

    def _build_ui(self):
        # fila 0: especialidad
        frame_esp = tk.Frame(self)
        frame_esp.pack(side=tk.TOP, fill=tk.X, padx=10, pady=(10, 5))

        tk.Label(frame_esp, text="Especialidad:").pack(side=tk.LEFT)

        self.var_especialidad = tk.StringVar()
        self.cmb_especialidad = ttk.Combobox(
            frame_esp,
            textvariable=self.var_especialidad,
            state="readonly",
            values=self.especialidades
        )
        self.cmb_especialidad.pack(side=tk.LEFT, padx=(5, 0), fill=tk.X, expand=True)
        self.cmb_especialidad.bind("<<ComboboxSelected>>", self.on_especialidad_changed)

        # fila 1: filtro de códigos
        frame_filter = tk.Frame(self)
        frame_filter.pack(side=tk.TOP, fill=tk.X, padx=10, pady=(10, 5))

        tk.Label(frame_filter, text="Filtrar códigos de elementos:").pack(side=tk.LEFT)

        self.var_filter = tk.StringVar()
        self.entry_filter = tk.Entry(frame_filter, textvariable=self.var_filter)
        self.entry_filter.pack(side=tk.LEFT, padx=(5, 0), fill=tk.X, expand=True)
        self.entry_filter.bind("<KeyRelease>", self.on_filter_changed)

        # fila 2: lista de códigos
        frame_list = tk.Frame(self)
        frame_list.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        self.listbox_codigos = tk.Listbox(frame_list, selectmode=tk.SINGLE)
        self.listbox_codigos.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = tk.Scrollbar(frame_list, orient="vertical", command=self.listbox_codigos.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.listbox_codigos.configure(yscrollcommand=scrollbar.set)

        self.listbox_codigos.bind("<<ListboxSelect>>", self.on_codigo_selected)

        # fila 3: botones
        frame_buttons = tk.Frame(self)
        frame_buttons.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=(0, 10))

        btn_cancel = tk.Button(frame_buttons, text="Cancelar", width=12, command=self.on_cancel)
        btn_cancel.pack(side=tk.RIGHT, padx=(5, 0))

        btn_ok = tk.Button(frame_buttons, text="Aceptar", width=12, command=self.on_accept)
        btn_ok.pack(side=tk.RIGHT, padx=(5, 10))

    # -------- lógica UI -------- #

    def on_especialidad_changed(self, event=None):
        self.selected_especialidad = self.var_especialidad.get()
        self.selected_codigo = None
        self.var_filter.set("")
        self._populate_codigos()

    def on_filter_changed(self, event=None):
        self._populate_codigos()

    def _get_codigos_for_especialidad(self):
        if not self.selected_especialidad:
            return []
        esp_dict = self.cache_data.get("especialidad", {})
        data = esp_dict.get(self.selected_especialidad, {})
        return data.get("codigos_elementos", []) or []

    def _populate_codigos(self):
        self.listbox_codigos.delete(0, tk.END)
        codigos = self._get_codigos_for_especialidad()
        filtro = self.var_filter.get().strip().lower()
        for c in codigos:
            if filtro and filtro not in c.lower():
                continue
            self.listbox_codigos.insert(tk.END, c)

    def on_codigo_selected(self, event=None):
        sel = self.listbox_codigos.curselection()
        if not sel:
            self.selected_codigo = None
            return
        idx = sel[0]
        self.selected_codigo = self.listbox_codigos.get(idx)

    def on_cancel(self):
        self.selected_especialidad = None
        self.selected_codigo = None
        self.destroy()

    def on_accept(self):
        if not self.selected_especialidad or not self.selected_codigo:
            messagebox.showwarning(
                "Aviso",
                "Debe seleccionar una especialidad y un código de elemento."
            )
            return
        self.destroy()


def main():
    if len(sys.argv) < 3:
        print("Uso: ui_tk_codintbim.py <cache_json> <output_json>")
        sys.exit(1)

    cache_json = sys.argv[1]
    output_json = sys.argv[2]

    if not os.path.exists(cache_json):
        print("Error: no existe cache_json: {}".format(cache_json))
        sys.exit(1)

    try:
        cache_data = cargar_cache(cache_json)
    except Exception as ex:
        print("Error al leer cache_json: {}".format(ex))
        sys.exit(1)

    app = CodIntBIMSelector(cache_data)
    app.mainloop()

    result = {
        "especialidad": app.selected_especialidad,
        "codigo_elemento": app.selected_codigo,
    }

    try:
        with open(output_json, "w", encoding="utf-8") as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
    except Exception as ex:
        print("Error al escribir output_json: {}".format(ex))
        sys.exit(1)

    sys.exit(0)


if __name__ == "__main__":
    main()
