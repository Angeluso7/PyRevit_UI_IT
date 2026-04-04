# -*- coding: utf-8 -*-
import sys
import os
import json
import tkinter as tk
from tkinter import ttk, messagebox


def cargar_schedules(path_json):
    """Lee el JSON de entrada con la lista de planillas."""
    with open(path_json, 'r', encoding='utf-8') as f:
        return json.load(f)  # lista de dicts: {"name": ..., "id": ...}


class ScheduleSelectorTk(tk.Tk):
    def __init__(self, schedules):
        super(ScheduleSelectorTk, self).__init__()

        self.title("Seleccionar Tablas de Planificación")
        self.geometry("520x600")
        self.minsize(520, 400)

        self.schedules = schedules
        self.all_names = [s["name"] for s in schedules]
        # nombre -> BooleanVar
        self.checked = {s["name"]: tk.BooleanVar(value=False) for s in schedules}

        # resultado
        self.selected_names = []

        self._build_ui()
        self._populate_list(self.all_names)

    def _build_ui(self):
        # --- fila 0: buscador ---
        frame_search = tk.Frame(self)
        frame_search.pack(side=tk.TOP, fill=tk.X, padx=10, pady=(10, 5))

        lbl = tk.Label(frame_search, text="Buscar:")
        lbl.pack(side=tk.LEFT)

        self.var_search = tk.StringVar()
        entry = tk.Entry(frame_search, textvariable=self.var_search, width=50)
        entry.pack(side=tk.LEFT, padx=(5, 0), fill=tk.X, expand=True)
        entry.bind("<KeyRelease>", self.on_search)

        # --- fila 1: lista con scroll ---
        frame_list = tk.Frame(self)
        frame_list.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        self.canvas = tk.Canvas(frame_list, borderwidth=0)
        self.scrollbar = tk.Scrollbar(frame_list, orient="vertical", command=self.canvas.yview)
        self.inner = tk.Frame(self.canvas)

        self.inner.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # habilitar scroll con rueda de mouse
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

        # --- fila 2: botones ---
        frame_buttons = tk.Frame(self)
        frame_buttons.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=(0, 10))

        frame_left = tk.Frame(frame_buttons)
        frame_left.pack(side=tk.LEFT)

        btn_select_all = tk.Button(frame_left, text="Seleccionar Todo",
                                   width=18, command=self.on_select_all)
        btn_select_all.pack(side=tk.LEFT, padx=(0, 5))

        btn_clear = tk.Button(frame_left, text="Eliminar Selección",
                              width=18, command=self.on_clear)
        btn_clear.pack(side=tk.LEFT)

        frame_right = tk.Frame(frame_buttons)
        frame_right.pack(side=tk.RIGHT)

        btn_export = tk.Button(frame_right, text="Exportar Seleccionados",
                               width=22, command=self.on_export)
        btn_export.pack(side=tk.LEFT, padx=(0, 5))

        btn_cancel = tk.Button(frame_right, text="Cancelar",
                               width=10, command=self.on_cancel)
        btn_cancel.pack(side=tk.LEFT)

    def _populate_list(self, names):
        # limpiar contenido previo
        for child in self.inner.winfo_children():
            child.destroy()

        # reconstruir lista
        for name in names:
            frame = tk.Frame(self.inner)
            frame.pack(anchor="w", fill=tk.X)

            var = self.checked[name]
            cb = tk.Checkbutton(frame, variable=var)
            cb.pack(side=tk.LEFT, padx=(2, 4), pady=2)

            lbl = tk.Label(frame, text=name, anchor="w")
            lbl.pack(side=tk.LEFT, fill=tk.X, expand=True)

    def on_search(self, event=None):
        text = self.var_search.get().strip().lower()
        if not text:
            filtered = self.all_names
        else:
            filtered = [n for n in self.all_names if text in n.lower()]
        self._populate_list(filtered)

    def _on_mousewheel(self, event):
        # Windows: event.delta es múltiplo de 120
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def on_select_all(self):
        for name in self.all_names:
            self.checked[name].set(True)

    def on_clear(self):
        for name in self.all_names:
            self.checked[name].set(False)

    def on_export(self):
        sel = [name for name, var in self.checked.items() if var.get()]
        if not sel:
            messagebox.showwarning("Aviso", "Seleccione al menos una tabla.")
            return
        self.selected_names = sel
        self.destroy()

    def on_cancel(self):
        self.selected_names = []
        self.destroy()


def main():
    if len(sys.argv) < 3:
        print("Uso: select_schedules_ui.py <input_json> <output_json>")
        sys.exit(1)

    input_json = sys.argv[1]
    output_json = sys.argv[2]

    if not os.path.exists(input_json):
        print("Error: no existe input_json: {}".format(input_json))
        sys.exit(1)

    schedules = cargar_schedules(input_json)

    app = ScheduleSelectorTk(schedules)
    app.mainloop()

    # escribir resultado
    result = {"selected_names": app.selected_names}
    with open(output_json, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    sys.exit(0)


if __name__ == "__main__":
    main()
