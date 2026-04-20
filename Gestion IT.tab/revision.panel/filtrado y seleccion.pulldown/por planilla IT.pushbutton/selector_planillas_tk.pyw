# selector_planillas_tk.pyw
# -*- coding: utf-8 -*-
"""
Selector externo CPython — tema dark.
Recibe: ruta_meta_entrada  ruta_salida_seleccion
"""

import os
import sys
import json
import tkinter as tk
from tkinter import messagebox

# ── Paleta dark ──────────────────────────────────────────────────────────────
BG          = "#1E1E2E"
SURFACE     = "#2A2A3E"
SURFACE2    = "#313145"
ACCENT      = "#4F7EFF"
ACCENT_HOV  = "#3A6AE8"
TEXT        = "#E0E0F0"
TEXT_MUTED  = "#9090AA"
BORDER      = "#3A3A52"
ENTRY_BG    = "#252535"
SEL_BG      = "#4F7EFF"
SEL_FG      = "#FFFFFF"
BTN_FG      = "#FFFFFF"
SCROLLBAR   = "#3A3A52"

# ── Helpers ───────────────────────────────────────────────────────────────────

def cargar_planillas_desde_json(path_json):
    if not os.path.exists(path_json):
        return []
    try:
        with open(path_json, "r", encoding="utf-8") as f:
            data = json.load(f)
        return list(data.get("codigos_planillas", {}).keys())
    except Exception:
        return []


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    if len(sys.argv) < 3:
        messagebox.showerror(
            "Error",
            "Se esperaban 2 argumentos: ruta_meta_entrada, ruta_salida_seleccion.")
        return

    ruta_meta_entrada = sys.argv[1]
    ruta_salida       = sys.argv[2]

    if not os.path.exists(ruta_meta_entrada):
        messagebox.showerror(
            "Error", "No se encontro archivo meta:\n{}".format(ruta_meta_entrada))
        return

    with open(ruta_meta_entrada, "r", encoding="utf-8") as f:
        meta = json.load(f)

    planillas_doc  = meta.get("planillas_doc", []) or []
    ruta_json      = meta.get("ruta_json", "") or ""
    planillas_json = cargar_planillas_desde_json(ruta_json)

    nombres_combinados = sorted(set(planillas_doc) | set(planillas_json))
    if not nombres_combinados:
        messagebox.showinfo("Aviso", "No se encontraron planillas.")
        return

    # ── Ventana principal ────────────────────────────────────────────────────
    root = tk.Tk()
    root.title("Seleccionar Planilla IT")
    root.geometry("480x420")
    root.minsize(380, 320)
    root.configure(bg=BG)
    root.resizable(True, True)

    # Centrar en pantalla
    root.update_idletasks()
    x = (root.winfo_screenwidth()  - 480) // 2
    y = (root.winfo_screenheight() - 420) // 2
    root.geometry("+{}+{}".format(x, y))

    root.columnconfigure(0, weight=1)
    root.rowconfigure(2, weight=1)

    # ── Titulo ───────────────────────────────────────────────────────────────
    frm_header = tk.Frame(root, bg=SURFACE, pady=12)
    frm_header.grid(row=0, column=0, sticky="ew")
    frm_header.columnconfigure(0, weight=1)

    tk.Label(
        frm_header,
        text="Filtrar por Planilla IT",
        font=("Segoe UI", 13, "bold"),
        fg=TEXT, bg=SURFACE
    ).grid(row=0, column=0, padx=16)

    tk.Label(
        frm_header,
        text="Selecciona una planilla para aplicar el filtro en la vista activa.",
        font=("Segoe UI", 9),
        fg=TEXT_MUTED, bg=SURFACE
    ).grid(row=1, column=0, padx=16)

    # ── Buscador ─────────────────────────────────────────────────────────────
    frm_search = tk.Frame(root, bg=BG, pady=10)
    frm_search.grid(row=1, column=0, sticky="ew", padx=16)
    frm_search.columnconfigure(1, weight=1)

    tk.Label(
        frm_search, text="Buscar:", font=("Segoe UI", 9),
        fg=TEXT_MUTED, bg=BG
    ).grid(row=0, column=0, padx=(0, 8))

    var_filtro = tk.StringVar()
    entry_filtro = tk.Entry(
        frm_search, textvariable=var_filtro,
        font=("Segoe UI", 10),
        bg=ENTRY_BG, fg=TEXT,
        insertbackground=TEXT,
        relief="flat",
        highlightthickness=1,
        highlightbackground=BORDER,
        highlightcolor=ACCENT
    )
    entry_filtro.grid(row=0, column=1, sticky="ew", ipady=5)

    # ── Listbox ──────────────────────────────────────────────────────────────
    frm_lista = tk.Frame(root, bg=BG)
    frm_lista.grid(row=2, column=0, sticky="nsew", padx=16, pady=(0, 6))
    frm_lista.columnconfigure(0, weight=1)
    frm_lista.rowconfigure(0, weight=1)

    # Contador
    var_contador = tk.StringVar()
    lbl_contador = tk.Label(
        frm_lista,
        textvariable=var_contador,
        font=("Segoe UI", 8),
        fg=TEXT_MUTED, bg=BG, anchor="w"
    )
    lbl_contador.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 4))

    listbox = tk.Listbox(
        frm_lista,
        selectmode="browse",
        font=("Segoe UI", 10),
        bg=SURFACE, fg=TEXT,
        selectbackground=SEL_BG,
        selectforeground=SEL_FG,
        activestyle="none",
        relief="flat",
        highlightthickness=1,
        highlightbackground=BORDER,
        highlightcolor=ACCENT,
        borderwidth=0
    )
    listbox.grid(row=1, column=0, sticky="nsew")

    scrollbar = tk.Scrollbar(
        frm_lista, orient="vertical", command=listbox.yview,
        bg=SCROLLBAR, troughcolor=SURFACE, width=10
    )
    scrollbar.grid(row=1, column=1, sticky="ns")
    listbox.config(yscrollcommand=scrollbar.set)

    # Hover alternado
    def on_motion(event):
        idx = listbox.nearest(event.y)
        for i in range(listbox.size()):
            if i in listbox.curselection():
                continue
            listbox.itemconfig(i, bg=SURFACE2 if i == idx else
                               (SURFACE if i % 2 == 0 else "#2E2E42"))

    listbox.bind("<Motion>", on_motion)

    # Actualizar lista con filtro
    def actualizar_lista(*_):
        filtro = var_filtro.get().strip().lower()
        listbox.delete(0, tk.END)
        count = 0
        for i, nombre in enumerate(nombres_combinados):
            if filtro in nombre.lower():
                listbox.insert(tk.END, "  " + nombre)
                listbox.itemconfig(tk.END, bg=SURFACE if i % 2 == 0 else "#2E2E42")
                count += 1
        var_contador.set("{} planilla{} encontrada{}".format(
            count, "s" if count != 1 else "", "s" if count != 1 else ""))

    var_filtro.trace_add("write", actualizar_lista)
    actualizar_lista()

    # ── Divisor ──────────────────────────────────────────────────────────────
    tk.Frame(root, bg=BORDER, height=1).grid(row=3, column=0, sticky="ew")

    # ── Botones ──────────────────────────────────────────────────────────────
    frm_btns = tk.Frame(root, bg=SURFACE, pady=12)
    frm_btns.grid(row=4, column=0, sticky="ew")
    frm_btns.columnconfigure(0, weight=1)
    frm_btns.columnconfigure(1, weight=1)

    seleccion = {"nombre": None}

    def on_aceptar(event=None):
        sel = listbox.curselection()
        if not sel:
            messagebox.showinfo("Aviso", "Por favor selecciona una planilla.")
            return
        nombre = listbox.get(sel[0]).strip()
        seleccion["nombre"] = nombre
        with open(ruta_salida, "w", encoding="utf-8") as f:
            json.dump({"selected_planilla": nombre}, f,
                      indent=2, ensure_ascii=False)
        root.destroy()

    def on_cancelar():
        root.destroy()

    # Doble clic = aceptar
    listbox.bind("<Double-Button-1>", on_aceptar)
    listbox.bind("<Return>", on_aceptar)

    def make_btn(parent, text, cmd, primary=False):
        bg_n = ACCENT      if primary else SURFACE2
        bg_h = ACCENT_HOV  if primary else BORDER
        btn  = tk.Label(
            parent, text=text,
            font=("Segoe UI", 10, "bold" if primary else "normal"),
            fg=BTN_FG, bg=bg_n,
            cursor="hand2", padx=24, pady=8, relief="flat"
        )
        btn.bind("<Button-1>",  lambda e: cmd())
        btn.bind("<Enter>",     lambda e: btn.config(bg=bg_h))
        btn.bind("<Leave>",     lambda e: btn.config(bg=bg_n))
        return btn

    btn_aceptar  = make_btn(frm_btns, "Aplicar filtro", on_aceptar,  primary=True)
    btn_cancelar = make_btn(frm_btns, "Cancelar",        on_cancelar, primary=False)
    btn_aceptar.grid( row=0, column=0, sticky="ew", padx=(16, 8))
    btn_cancelar.grid(row=0, column=1, sticky="ew", padx=(8,  16))

    entry_filtro.focus_set()
    root.mainloop()


if __name__ == "__main__":
    main()
