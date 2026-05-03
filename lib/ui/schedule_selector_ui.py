# -*- coding: utf-8 -*-
"""
schedule_selector_ui.py  (CPython)
Ventana Tkinter dark para seleccionar planillas a exportar.
Recibe: argv[1] = ruta JSON entrada  (lista de nombres)
        argv[2] = ruta JSON salida   (lista seleccionada o cancelar)
"""
import sys
import os
import json
import tkinter as tk
from tkinter import ttk, messagebox

# ── Paleta dark ──────────────────────────────────────────────────────────────
DARK = {
    "bg":           "#1e1e1e",
    "surface":      "#2a2a2a",
    "border":       "#3c3c3c",
    "fg":           "#d4d4d4",
    "fg_muted":     "#888888",
    "accent":       "#4f98a3",
    "accent_hover": "#3a7d88",
    "btn_bg":       "#3a3a3a",
    "btn_fg":       "#d4d4d4",
    "disabled_bg":  "#2d2d2d",
    "disabled_fg":  "#5a5a5a",
    "select_bg":    "#094771",
    "select_fg":    "#ffffff",
}


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
    # Combobox desplegable
    root.option_add("*TCombobox*Listbox.background",  DARK["surface"])
    root.option_add("*TCombobox*Listbox.foreground",  DARK["fg"])
    root.option_add("*TCombobox*Listbox.selectBackground", DARK["select_bg"])
    root.option_add("*TCombobox*Listbox.selectForeground", DARK["select_fg"])


class ScheduleSelectorApp(object):
    def __init__(self, root, nombres, out_path, title):
        self.root      = root
        self.nombres   = nombres
        self.out_path  = out_path
        self.filter_var = tk.StringVar()
        self._construir_ui(title)
        self._refrescar_lista()
        self.filter_var.trace_add('write', lambda *a: self._refrescar_lista())
        ttk.Sizegrip(self.root).place(relx=1.0, rely=1.0, anchor=tk.SE)

    def _construir_ui(self, title):
        self.root.title(title)
        self.root.geometry('480x400')
        self.root.minsize(360, 300)

        main = ttk.Frame(self.root, padding=10)
        main.pack(fill='both', expand=True)

        # ── Buscador ─────────────────────────────────────────────────────────
        top = ttk.Frame(main)
        top.pack(fill='x', pady=(0, 6))
        ttk.Label(top, text='Buscar:').pack(side='left')
        ttk.Entry(top, textvariable=self.filter_var).pack(
            side='left', fill='x', expand=True, padx=(6, 0))

        # ── Lista con scroll ─────────────────────────────────────────────────
        frame_list = ttk.Frame(main)
        frame_list.pack(fill='both', expand=True)
        frame_list.rowconfigure(0, weight=1)
        frame_list.columnconfigure(0, weight=1)

        self.listbox = tk.Listbox(
            frame_list,
            selectmode=tk.EXTENDED,
            exportselection=False,
            bg=DARK["surface"],
            fg=DARK["fg"],
            selectbackground=DARK["select_bg"],
            selectforeground=DARK["select_fg"],
            highlightbackground=DARK["border"],
            highlightcolor=DARK["accent"],
            highlightthickness=1,
            activestyle='dotbox',
            relief='flat',
        )
        self.listbox.grid(row=0, column=0, sticky='nsew')
        scroll = ttk.Scrollbar(frame_list, orient='vertical',
                               command=self.listbox.yview)
        scroll.grid(row=0, column=1, sticky='ns')
        self.listbox.configure(yscrollcommand=scroll.set)

        # ── Contador de selección ─────────────────────────────────────────────
        self.lbl_count = ttk.Label(main, text='0 planillas seleccionadas',
                                   foreground=DARK["fg_muted"])
        self.lbl_count.pack(anchor='w', pady=(4, 0))
        self.listbox.bind('<<ListboxSelect>>', self._actualizar_contador)

        # ── Botones ───────────────────────────────────────────────────────────
        frame_btn = ttk.Frame(main)
        frame_btn.pack(fill='x', pady=(10, 0))
        frame_btn.columnconfigure(0, weight=1)
        ttk.Button(frame_btn, text='Cancelar',
                   command=self.on_cancelar).grid(
            row=0, column=1, sticky='e', padx=(0, 6))
        ttk.Button(frame_btn, text='Exportar selección',
                   command=self.on_aceptar,
                   style='Accent.TButton').grid(
            row=0, column=2, sticky='e')

    def _refrescar_lista(self):
        filtro = self.filter_var.get().strip().lower()
        self.listbox.delete(0, tk.END)
        for nombre in sorted(self.nombres):
            if not filtro or filtro in nombre.lower():
                self.listbox.insert(tk.END, nombre)
        self._actualizar_contador()

    def _actualizar_contador(self, *_):
        n = len(self.listbox.curselection())
        self.lbl_count.configure(
            text='{} planilla{} seleccionada{}'.format(
                n, 's' if n != 1 else '', 's' if n != 1 else ''))

    def on_cancelar(self):
        self._escribir_salida({'opcion': 'cancelar', 'seleccion': []})
        self.root.destroy()

    def on_aceptar(self):
        indices = self.listbox.curselection()
        if not indices:
            messagebox.showwarning('Aviso',
                                   'Selecciona al menos una planilla.')
            return
        seleccion = [self.listbox.get(i) for i in indices]
        self._escribir_salida({'opcion': 'aceptar', 'seleccion': seleccion})
        self.root.destroy()

    def _escribir_salida(self, data):
        try:
            with open(self.out_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            messagebox.showerror('Error',
                                 'No se pudo escribir salida:\n{}'.format(e))


def main():
    if len(sys.argv) < 3:
        messagebox.showerror('Error',
            'Uso: schedule_selector_ui.py <input.json> <output.json>')
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
            nombres = json.load(f)
    except Exception as e:
        messagebox.showerror('Error',
            'Error leyendo JSON de entrada:\n{}'.format(e))
        return

    root = tk.Tk()
    root.resizable(True, True)
    _aplicar_tema_dark(root)       # ← tema dark antes de construir la UI
    ScheduleSelectorApp(root, nombres, out_path, title)
    root.mainloop()


if __name__ == '__main__':
    main()
