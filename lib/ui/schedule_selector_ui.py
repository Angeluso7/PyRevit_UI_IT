# -*- coding: utf-8 -*-
"""
schedule_selector_ui.py  (CPython)
Ventana Tkinter dark para seleccionar planillas a exportar.
Recibe: argv[1] = ruta JSON entrada  (lista de strings O lista de dicts {name, id})
        argv[2] = ruta JSON salida   (lista seleccionada o cancelar)
        argv[3] = titulo ventana     (opcional)
Salida JSON: {"opcion": "aceptar"|"cancelar", "seleccion": [nombre, ...]}
"""
import sys
import os
import json
import tkinter as tk
from tkinter import ttk, messagebox

# ── Paleta dark ──────────────────────────────────────────────────────────────
DARK = {
    "bg":            "#1e1e1e",
    "surface":       "#2a2a2a",
    "surface_row":   "#252525",
    "surface_hover": "#303030",
    "border":        "#3c3c3c",
    "fg":            "#d4d4d4",
    "fg_muted":      "#888888",
    "accent":        "#4f98a3",
    "accent_hover":  "#3a7d88",
    "btn_bg":        "#3a3a3a",
    "btn_fg":        "#d4d4d4",
    "disabled_bg":   "#2d2d2d",
    "disabled_fg":   "#5a5a5a",
    "select_bg":     "#094771",
    "select_fg":     "#ffffff",
}


def _normalizar_nombres(data):
    """
    Acepta tanto lista de strings como lista de dicts {"name": ..., "id": ...}.
    Devuelve siempre una lista de strings ordenada alfabeticamente.
    """
    if not data:
        return []
    if isinstance(data[0], dict):
        return sorted(item["name"] for item in data)
    return sorted(str(item) for item in data)


def _aplicar_tema_dark(root):
    root.configure(bg=DARK["bg"])
    style = ttk.Style(root)
    style.theme_use("clam")

    style.configure("TFrame",      background=DARK["bg"])
    style.configure("TLabelframe", background=DARK["bg"],
                    foreground=DARK["fg"], bordercolor=DARK["border"])
    style.configure("TLabelframe.Label", background=DARK["bg"],
                    foreground=DARK["fg"])
    style.configure("TLabel",
                    background=DARK["bg"],
                    foreground=DARK["fg"])
    style.configure("Row.TCheckbutton",
                    background=DARK["surface"],
                    foreground=DARK["fg"],
                    focuscolor=DARK["accent"])
    style.map("Row.TCheckbutton",
              background=[("active", DARK["surface_hover"])])
    style.configure("TEntry",
                    fieldbackground=DARK["surface"],
                    foreground=DARK["fg"],
                    insertcolor=DARK["fg"],
                    bordercolor=DARK["border"],
                    lightcolor=DARK["border"],
                    darkcolor=DARK["border"])
    style.map("TEntry",
              fieldbackground=[("disabled", DARK["disabled_bg"])],
              foreground=[("disabled", DARK["disabled_fg"])])
    style.configure("TScrollbar",
                    background=DARK["btn_bg"],
                    troughcolor=DARK["bg"],
                    bordercolor=DARK["border"],
                    arrowcolor=DARK["fg"],
                    relief="flat")
    style.map("TScrollbar",
              background=[("active", DARK["border"])])
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


class ScheduleSelectorApp(object):
    def __init__(self, root, nombres, out_path, title):
        self.root       = root
        self.nombres    = nombres          # ya es lista de strings ordenada
        self.out_path   = out_path
        self.filter_var = tk.StringVar()
        self._checks    = {n: tk.BooleanVar(value=False) for n in self.nombres}
        self._rows      = []
        self._construir_ui(title)
        self._refrescar_lista()
        self.filter_var.trace_add('write', lambda *a: self._refrescar_lista())
        ttk.Sizegrip(self.root).place(relx=1.0, rely=1.0, anchor=tk.SE)

    # ── UI ────────────────────────────────────────────────────────────────
    def _construir_ui(self, title):
        self.root.title(title)
        self.root.geometry('500x440')
        self.root.minsize(380, 320)

        main = ttk.Frame(self.root, padding=10)
        main.pack(fill='both', expand=True)

        # ── Buscador ──────────────────────────────────────────────────────
        top = ttk.Frame(main)
        top.pack(fill='x', pady=(0, 4))
        ttk.Label(top, text='Buscar:').pack(side='left')
        ttk.Entry(top, textvariable=self.filter_var).pack(
            side='left', fill='x', expand=True, padx=(6, 0))

        # ── Acciones rápidas ───────────────────────────────────────────────
        act = ttk.Frame(main)
        act.pack(fill='x', pady=(0, 6))
        ttk.Button(act, text='Seleccionar todo',
                   command=self._seleccionar_todo).pack(side='left', padx=(0, 4))
        ttk.Button(act, text='Limpiar',
                   command=self._limpiar_todo).pack(side='left')

        # ── Lista con checkboxes (Canvas + scrollbar) ────────────────────
        list_outer = ttk.Frame(main)
        list_outer.pack(fill='both', expand=True)
        list_outer.rowconfigure(0, weight=1)
        list_outer.columnconfigure(0, weight=1)

        self.canvas = tk.Canvas(
            list_outer,
            bg=DARK["surface"],
            highlightthickness=1,
            highlightbackground=DARK["border"],
            bd=0
        )
        self.canvas.grid(row=0, column=0, sticky='nsew')

        vscroll = ttk.Scrollbar(list_outer, orient='vertical',
                                command=self.canvas.yview)
        vscroll.grid(row=0, column=1, sticky='ns')
        self.canvas.configure(yscrollcommand=vscroll.set)

        self.inner = tk.Frame(self.canvas, bg=DARK["surface"])
        self._win_id = self.canvas.create_window(
            (0, 0), window=self.inner, anchor='nw')

        self.inner.bind('<Configure>', self._on_inner_configure)
        self.canvas.bind('<Configure>', self._on_canvas_configure)
        self.canvas.bind('<MouseWheel>',
            lambda e: self.canvas.yview_scroll(-1 * (e.delta // 120), 'units'))

        # ── Contador ────────────────────────────────────────────────────────
        self.lbl_count = ttk.Label(
            main, text='0 planillas seleccionadas',
            foreground=DARK["fg_muted"])
        self.lbl_count.pack(anchor='w', pady=(4, 0))

        # ── Botones ─────────────────────────────────────────────────────────
        frame_btn = ttk.Frame(main)
        frame_btn.pack(fill='x', pady=(8, 0))
        frame_btn.columnconfigure(0, weight=1)
        ttk.Button(frame_btn, text='Cancelar',
                   command=self.on_cancelar).grid(
            row=0, column=1, sticky='e', padx=(0, 6))
        ttk.Button(frame_btn, text='Exportar selección',
                   command=self.on_aceptar,
                   style='Accent.TButton').grid(
            row=0, column=2, sticky='e')

    # ── Canvas helpers ──────────────────────────────────────────────────
    def _on_inner_configure(self, _event=None):
        self.canvas.configure(scrollregion=self.canvas.bbox('all'))

    def _on_canvas_configure(self, event):
        self.canvas.itemconfig(self._win_id, width=event.width)

    # ── Lista de checkboxes ──────────────────────────────────────────────
    def _refrescar_lista(self):
        for widget in self.inner.winfo_children():
            widget.destroy()
        self._rows = []

        filtro = self.filter_var.get().strip().lower()
        for nombre in self.nombres:
            if filtro and filtro not in nombre.lower():
                continue
            var = self._checks[nombre]
            row = tk.Frame(self.inner, bg=DARK["surface"], highlightthickness=0)
            row.pack(fill='x', padx=2, pady=1)
            cb = ttk.Checkbutton(
                row,
                text=nombre,
                variable=var,
                style='Row.TCheckbutton',
                command=self._actualizar_contador
            )
            cb.pack(fill='x', padx=6, pady=2)
            row.bind('<Enter>', lambda e, f=row: f.configure(bg=DARK["surface_hover"]))
            row.bind('<Leave>', lambda e, f=row: f.configure(bg=DARK["surface"]))
            self._rows.append((row, cb))

        self._actualizar_contador()

    def _actualizar_contador(self):
        n = sum(1 for v in self._checks.values() if v.get())
        self.lbl_count.configure(
            text='{} planilla{} seleccionada{}'.format(
                n, 's' if n != 1 else '', 's' if n != 1 else ''))

    def _seleccionar_todo(self):
        filtro = self.filter_var.get().strip().lower()
        for nombre, var in self._checks.items():
            if not filtro or filtro in nombre.lower():
                var.set(True)
        self._actualizar_contador()

    def _limpiar_todo(self):
        for var in self._checks.values():
            var.set(False)
        self._actualizar_contador()

    # ── Acciones ────────────────────────────────────────────────────────────
    def on_cancelar(self):
        self._escribir_salida({'opcion': 'cancelar', 'seleccion': []})
        self.root.destroy()

    def on_aceptar(self):
        seleccion = [n for n, v in self._checks.items() if v.get()]
        if not seleccion:
            messagebox.showwarning('Aviso', 'Marca al menos una planilla.')
            return
        self._escribir_salida({'opcion': 'aceptar', 'seleccion': seleccion})
        self.root.destroy()

    def _escribir_salida(self, data):
        try:
            with open(self.out_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            messagebox.showerror('Error', 'No se pudo escribir salida:\n{}'.format(e))


def main():
    if len(sys.argv) < 3:
        messagebox.showerror('Error',
            'Uso: schedule_selector_ui.py <input.json> <output.json> [titulo]')
        return

    in_path  = sys.argv[1]
    out_path = sys.argv[2]
    title    = sys.argv[3] if len(sys.argv) > 3 else 'Seleccionar planillas'

    if not os.path.exists(in_path):
        messagebox.showerror('Error',
            'No se encontró archivo de entrada:\n{}'.format(in_path))
        return

    try:
        with open(in_path, 'r', encoding='utf-8') as f:
            raw = json.load(f)
    except Exception as e:
        messagebox.showerror('Error',
            'Error leyendo JSON de entrada:\n{}'.format(e))
        return

    # Normalizar: acepta lista de strings O lista de dicts {name, id}
    nombres = _normalizar_nombres(raw)

    root = tk.Tk()
    root.resizable(True, True)
    _aplicar_tema_dark(root)
    ScheduleSelectorApp(root, nombres, out_path, title)
    root.mainloop()


if __name__ == '__main__':
    main()
