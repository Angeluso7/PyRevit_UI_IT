# coding: utf-8
import os
import sys
import json
import tkinter as tk
from tkinter import ttk
from tkinter import scrolledtext

def main():
    # Parámetros: ruta del HTML o texto plano
    if len(sys.argv) < 2:
        return

    datapath = sys.argv[1]
    if not os.path.exists(datapath):
        return

    with open(datapath, "r", encoding="utf-8") as f:
        contenido = f.read()

    root = tk.Tk()
    root.title("Ayuda - Extensión PyRevit Infotécnica")
    root.geometry("900x700")

    frame = ttk.Frame(root, padding=10)
    frame.pack(fill="both", expand=True)

    txt = scrolledtext.ScrolledText(frame, wrap="word")
    txt.pack(fill="both", expand=True)

    txt.insert("1.0", contenido)
    txt.configure(state="disabled")

    root.mainloop()

if __name__ == "__main__":
    main()
