# -*- coding: utf-8 -*-
"""
schedule_selector.py  (IronPython / pyRevit)
Lanza schedule_selector_ui.py como subproceso CPython con tema dark
y lee el resultado desde un JSON temporal.
"""
import os
import sys
import json
import time
import subprocess
import tempfile


def _get_python_exe():
    """Devuelve la ruta al ejecutable CPython (mismo helper que el resto del proyecto)."""
    try:
        _this = os.path.abspath(__file__)
        _lib  = os.path.dirname(_this)                  # lib/ui -> lib
        _ext  = os.path.dirname(_lib)                   # lib    -> extension root
        sys.path.insert(0, _lib)
        from core.env_config import get_python_exe
        return get_python_exe()
    except Exception:
        # Fallback: buscar python.exe en PATH
        return 'python'


def select_schedules(schedules, title='Seleccionar planillas para exportar'):
    """
    Muestra una ventana Tkinter dark con la lista de planillas.
    Devuelve la lista de objetos Schedule seleccionados, o [] si cancela.
    """
    nombres = sorted([s.Name for s in schedules])

    # ── Archivos temporales ────────────────────────────────────────────────
    tmp_dir  = tempfile.gettempdir()
    in_path  = os.path.join(tmp_dir, 'sch_sel_input.json')
    out_path = os.path.join(tmp_dir, 'sch_sel_output.json')

    # Limpiar output anterior
    if os.path.exists(out_path):
        try:
            os.remove(out_path)
        except Exception:
            pass

    # Escribir input
    try:
        with open(in_path, 'w', encoding='utf-8') as f:
            json.dump(nombres, f, ensure_ascii=False, indent=2)
    except Exception as e:
        from pyrevit import forms
        forms.alert('No se pudo crear archivo temporal:\n{}'.format(e),
                    title='Error')
        return []

    # ── Localizar el script UI ─────────────────────────────────────────────
    _ui_dir    = os.path.dirname(os.path.abspath(__file__))
    ui_script  = os.path.join(_ui_dir, 'schedule_selector_ui.py')

    python_exe = _get_python_exe()

    # ── Lanzar subproceso CPython ──────────────────────────────────────────
    try:
        subprocess.Popen(
            [python_exe, ui_script, in_path, out_path, title],
            creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0)
        )
    except Exception as e:
        from pyrevit import forms
        forms.alert('No se pudo lanzar la ventana de selección:\n{}'.format(e),
                    title='Error')
        return []

    # ── Esperar resultado (hasta 60 s) ─────────────────────────────────────
    salida = None
    for _ in range(600):
        if os.path.exists(out_path):
            try:
                with open(out_path, 'r', encoding='utf-8') as f:
                    salida = json.load(f)
                break
            except ValueError:
                pass  # JSON aún incompleto, reintentar
        time.sleep(0.1)

    if not salida or salida.get('opcion') == 'cancelar':
        return []

    seleccion_nombres = set(salida.get('seleccion', []))
    return [s for s in schedules if s.Name in seleccion_nombres]
