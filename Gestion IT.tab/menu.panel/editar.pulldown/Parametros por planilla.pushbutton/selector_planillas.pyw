# -*- coding: utf-8 -*-
"""
Selector de planillas en Tkinter.
Lee script.json (codigos_planillas + reemplazos_de_nombres) y
permite elegir una planilla por alias, con filtro por texto.
Escribe un JSON meta con:
- NombrePlanillaOriginal
- NombrePlanillaAlias
- CodigoPlanilla
en la ruta pasada por argumento (planilla_meta_tmp.json).

Estilo: sv-ttk dark (Sun Valley theme).
  pip install sv-ttk
"""

import os
import sys
import json
import traceback
import tkinter as tk
from tkinter import ttk, messagebox

try:
    import sv_ttk
    _HAS_SV_TTK = True
except ImportError:
    _HAS_SV_TTK = False

# ── Rutas dinamicas desde __file__ ──────────────────────────────────────────
_THIS_DIR = os.path.dirname(os.path.abspath(__file__))
_EXT_ROOT = os.path.abspath(os.path.join(_THIS_DIR, '..', '..', '..', '..'))
MASTER_DIR = os.path.join(_EXT_ROOT, 'data', 'master')

SCRIPT_JSON_PATH = os.path.join(MASTER_DIR, 'script.json')

if len(sys.argv) > 1:
    PLANILLA_META_PATH = sys.argv[1]
else:
    TEMP_DIR = os.path.join(_EXT_ROOT, 'data', 'temp')
    PLANILLA_META_PATH = os.path.join(TEMP_DIR, 'planilla_meta_tmp.json')


def cargar_script_json():
    if not os.path.exists(SCRIPT_JSON_PATH):
        messagebox.showwarning(
            'script.json no encontrado',
            'No se encontro script.json en:\n{}\n\n'
            'Crea el archivo con la clave "codigos_planillas" para '
            'habilitar el selector.'.format(SCRIPT_JSON_PATH)
        )
        return {}
    try:
        with open(SCRIPT_JSON_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        messagebox.showerror(
            'Error script.json',
            'Error leyendo script.json:\n{}'.format(traceback.format_exc())
        )
        return {}


def construir_lista_planillas(cfg):
    codigos    = cfg.get('codigos_planillas', {}) or {}
    reemplazos = cfg.get('reemplazos_de_nombres', {}) or {}
    items = []
    for key, codigo in codigos.items():
        alias = reemplazos.get(key, key)
        items.append((alias, key, codigo))
    items.sort(key=lambda x: (x[0] or '').lower())
    return items


def guardar_meta(nombre_orig, nombre_alias, codigo):
    try:
        data = {
            'NombrePlanillaOriginal': nombre_orig,
            'NombrePlanillaAlias':    nombre_alias,
            'CodigoPlanilla':         codigo,
        }
        folder = os.path.dirname(PLANILLA_META_PATH)
        if folder and not os.path.exists(folder):
            os.makedirs(folder)
        with open(PLANILLA_META_PATH, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
    except Exception:
        messagebox.showerror(
            'Error',
            'Error guardando meta de planilla:\n{}'.format(traceback.format_exc())
        )


def main():
    cfg = cargar_script_json()
    if not cfg:
        return

    planillas = construir_lista_planillas(cfg)
    if not planillas:
        messagebox.showinfo(
            'Sin planillas',
            'No se encontraron planillas en script.json.\n'
            'Agrega la clave "codigos_planillas" con los codigos correspondientes.'
        )
        return

    root = tk.Tk()

    if _HAS_SV_TTK:
        sv_ttk.set_theme("dark")

    root.title('Selector de planilla')
    root.geometry('420x440')
    root.minsize(320, 320)

    # ── Barra de filtro ──────────────────────────────────────────────────────
    top_frame = ttk.Frame(root, padding=(8, 8, 8, 4))
    top_frame.pack(fill='x')

    ttk.Label(top_frame, text='Buscar:').pack(side='left', padx=(0, 6))
    filt_var = tk.StringVar()
    entry = ttk.Entry(top_frame, textvariable=filt_var)
    entry.pack(side='left', fill='x', expand=True)

    # ── Lista ────────────────────────────────────────────────────────────────
    list_frame = ttk.Frame(root, padding=(8, 0, 8, 0))
    list_frame.pack(fill='both', expand=True)

    scrollbar = ttk.Scrollbar(list_frame, orient='vertical')
    listbox = tk.Listbox(
        list_frame,
        selectmode='single',
        yscrollcommand=scrollbar.set,
        activestyle='none',
        relief='flat',
        borderwidth=0,
        highlightthickness=0,
    )
    scrollbar.config(command=listbox.yview)
    scrollbar.pack(side='right', fill='y')
    listbox.pack(side='left', fill='both', expand=True)

    # ── Botones ──────────────────────────────────────────────────────────────
    btn_frame = ttk.Frame(root, padding=(8, 6, 8, 10))
    btn_frame.pack(fill='x')
    btn_frame.columnconfigure(0, weight=1)
    btn_frame.columnconfigure(1, weight=1)

    current_items = list(planillas)

    def refrescar_lista(*args):
        texto = filt_var.get().strip().lower()
        listbox.delete(0, tk.END)
        del current_items[:]
        for alias, nombre_orig, codigo in planillas:
            if not texto or texto in (alias or '').lower():
                current_items.append((alias, nombre_orig, codigo))
                listbox.insert(tk.END, '  ' + alias)

    filt_var.trace_add('write', refrescar_lista)

    def aceptar():
        sel = listbox.curselection()
        if not sel:
            messagebox.showinfo('Informacion', 'Seleccione una planilla.')
            return
        idx = sel[0]
        alias, nombre_orig, codigo = current_items[idx]
        if not codigo:
            messagebox.showerror(
                'Error',
                'La planilla seleccionada no tiene codigo en script.json.'
            )
            return
        guardar_meta(nombre_orig, alias, codigo)
        root.destroy()

    def cancelar():
        root.destroy()

    ttk.Button(btn_frame, text='Cancelar', command=cancelar).grid(
        row=0, column=0, sticky='ew', padx=(0, 4))
    ttk.Button(btn_frame, text='Aceptar', command=aceptar).grid(
        row=0, column=1, sticky='ew', padx=(4, 0))

    listbox.bind('<Double-Button-1>', lambda e: aceptar())

    refrescar_lista()
    entry.focus()
    root.mainloop()


if __name__ == '__main__':
    main()
