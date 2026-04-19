# -*- coding: utf-8 -*-
import sys
import os
import json
import tkinter as tk
from tkinter import ttk, messagebox

class CodIntSelectorApp(object):
    def __init__(self, root, datos, out_path):
        self.root = root
        self.datos = datos or {}
        self.out_path = out_path
        self.opcion_var = tk.StringVar(value='by_codint')
        self.asignados_var = tk.StringVar(value='asignados')
        self.filter_var = tk.StringVar()
        self.codint_unicos = sorted({e.get('codintbim', '') for e in self.datos.get('elementos', []) if e.get('codintbim', '')})
        self.listbox = None
        self.rb_asig = None
        self.rb_noasig = None
        self.entry_filt = None
        self._construir_ui()
        self._refrescar_lista()
        self.opcion_var.trace_add('write', self._actualizar_estado_renglones)
        self._actualizar_estado_renglones()
        ttk.Sizegrip(self.root).place(relx=1.0, rely=1.0, anchor=tk.SE)

    def _construir_ui(self):
        self.root.title('Selección por CodIntBIM')
        self.root.geometry('520x360')
        self.root.minsize(400, 260)
        main = ttk.Frame(self.root, padding=10)
        main.pack(fill='both', expand=True)
        frame1 = ttk.LabelFrame(main, text='Renglón 1 - Por CodIntBIM')
        frame1.pack(fill='both', expand=True, pady=(0, 10))
        ttk.Radiobutton(frame1, text='Usar filtro por CodIntBIM', variable=self.opcion_var, value='by_codint').grid(row=0, column=0, columnspan=2, sticky='w', pady=(0, 5))
        ttk.Label(frame1, text='Buscar:').grid(row=1, column=0, sticky='w')
        self.entry_filt = ttk.Entry(frame1, textvariable=self.filter_var, width=30)
        self.entry_filt.grid(row=1, column=1, sticky='ew', padx=(5, 0))
        frame1.columnconfigure(1, weight=1)
        self.filter_var.trace_add('write', lambda *args: self._refrescar_lista())
        frame_list = ttk.Frame(frame1)
        frame_list.grid(row=2, column=0, columnspan=2, sticky='nsew', pady=(5, 0))
        frame1.rowconfigure(2, weight=1)
        frame_list.rowconfigure(0, weight=1)
        frame_list.columnconfigure(0, weight=1)
        self.listbox = tk.Listbox(frame_list, height=6, exportselection=False)
        self.listbox.grid(row=0, column=0, sticky='nsew')
        scroll = ttk.Scrollbar(frame_list, orient='vertical', command=self.listbox.yview)
        scroll.grid(row=0, column=1, sticky='ns')
        self.listbox.configure(yscrollcommand=scroll.set)
        frame2 = ttk.LabelFrame(main, text='Renglón 2 - Asignados / No asignados')
        frame2.pack(fill='x', pady=(0, 10))
        ttk.Radiobutton(frame2, text='Usar modo Asignados/No asignados', variable=self.opcion_var, value='row2').grid(row=0, column=0, columnspan=2, sticky='w', pady=(0, 5))
        self.rb_asig = ttk.Radiobutton(frame2, text='Asignados', variable=self.asignados_var, value='asignados')
        self.rb_noasig = ttk.Radiobutton(frame2, text='No asignados', variable=self.asignados_var, value='no_asignados')
        self.rb_asig.grid(row=1, column=0, sticky='w')
        self.rb_noasig.grid(row=1, column=1, sticky='w')
        frame_btn = ttk.Frame(main)
        frame_btn.pack(fill='x', pady=(10, 0))
        frame_btn.columnconfigure(0, weight=1)
        ttk.Button(frame_btn, text='Cancelar', command=self.on_cancelar).grid(row=0, column=1, sticky='e', padx=(0, 5))
        ttk.Button(frame_btn, text='Aceptar', command=self.on_aceptar).grid(row=0, column=2, sticky='e')

    def _actualizar_estado_renglones(self, *args):
        op = self.opcion_var.get()
        if op == 'by_codint':
            self.entry_filt.configure(state='normal')
            self.listbox.configure(state='normal')
            self.rb_asig.configure(state='disabled')
            self.rb_noasig.configure(state='disabled')
        else:
            self.entry_filt.configure(state='disabled')
            self.listbox.configure(state='disabled')
            self.rb_asig.configure(state='normal')
            self.rb_noasig.configure(state='normal')

    def _refrescar_lista(self):
        filtro = self.filter_var.get().strip().lower()
        self.listbox.delete(0, tk.END)
        for cod in self.codint_unicos:
            if not filtro or filtro in cod.lower():
                self.listbox.insert(tk.END, cod)

    def on_cancelar(self):
        self._escribir_salida({'opcion': 'cancelar'})
        self.root.destroy()

    def on_aceptar(self):
        op = self.opcion_var.get()
        if op == 'by_codint':
            sel = self.listbox.curselection()
            if not sel:
                messagebox.showwarning('Aviso', 'Selecciona un CodIntBIM de la lista.')
                return
            idx = sel[0]
            cod = self.listbox.get(idx)
            self._escribir_salida({'opcion': 'by_codint', 'codintbim': cod})
        else:
            self._escribir_salida({'opcion': self.asignados_var.get()})
        self.root.destroy()

    def _escribir_salida(self, data):
        try:
            with open(self.out_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            messagebox.showerror('Error', 'No se pudo escribir salida:\n{}'.format(e))


def main():
    if len(sys.argv) < 3:
        messagebox.showerror('Error', 'Uso: codint_selector.py <input.json> <output.json>')
        return
    in_path = sys.argv[1]
    out_path = sys.argv[2]
    if not os.path.exists(in_path):
        messagebox.showerror('Error', 'No se encontró archivo de entrada:\n{}'.format(in_path))
        return
    try:
        with open(in_path, 'r', encoding='utf-8') as f:
            datos = json.load(f)
    except Exception as e:
        messagebox.showerror('Error', 'Error leyendo JSON de entrada:\n{}'.format(e))
        return
    root = tk.Tk()
    root.resizable(True, True)
    CodIntSelectorApp(root, datos, out_path)
    root.mainloop()

if __name__ == '__main__':
    main()
