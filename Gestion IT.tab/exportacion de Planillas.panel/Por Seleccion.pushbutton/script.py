# -*- coding: utf-8 -*-
__title__   = "Planillas valorización"
__doc__     = """Version = 1.0
Date    = 15.06.2024
________________________________________________________________
Description:

This is the placeholder for a .pushbutton
You can use it to start your pyRevit Add-In

________________________________________________________________
How-To:

1. [Hold ALT + CLICK] on the button to open its source folder.
You will be able to override this placeholder.

2. Automate Your Boring Work ;)

________________________________________________________________
TODO:
[FEATURE] - Describe Your ToDo Tasks Here
________________________________________________________________
Last Updates:
- [15.06.2024] v1.0 Change Description
- [10.06.2024] v0.5 Change Description
- [05.06.2024] v0.1 Change Description 
________________________________________________________________
Author: Erik Frits"""

# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝
#==================================================
# -*- coding: utf-8 -*-
"""Exportar tablas seleccionadas a CSV en RevitPythonShell y crear planillas XLSX"""
import os
import csv
import re
import clr
import subprocess
import json

from pyrevit import script, forms

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import ViewSchedule, SectionType, FilteredElementCollector

from System.IO import File  # para comprobar existencia de ui.xaml

# ╦  ╦╔═╗╦═╗╦╔═╗╔╗ ╦  ╔═╗╔═╗
# ╚╗╔╝╠═╣╠╦╝║╠═╣╠╩╗║  ║╣ ╚═╗
#  ╚╝ ╩ ╩╩╚═╩╩ ╩╚═╝╩═╝╚═╝╚═╝
#==================================================

# Obtener documento activo
doc = __revit__.ActiveUIDocument.Document

# ╔╦╗╔═╗╦╔╗╔
# ║║║╠═╣║║║║
# ╩ ╩╩ ╩╩╝╚╝
#==================================================

def sanitize_filename(name):
    if isinstance(name, tuple):
        name = name[0]
    if not isinstance(name, str):
        name = str(name)
    return re.sub(r'[\/*?:"<>|]', "_", name[:30])


def get_schedule_data(schedule):
    data = []
    body = schedule.GetTableData().GetSectionData(SectionType.Body)
    n_rows = body.NumberOfRows
    n_cols = body.NumberOfColumns

    for r in range(n_rows):
        row_data = []
        for c in range(n_cols):
            val = schedule.GetCellText(SectionType.Body, r, c)
            row_data.append(val if val else "")
        if any(cell.strip() for cell in row_data):
            data.append(row_data)
    return data


def export_schedule_csv(schedule, folder):
    name = sanitize_filename(schedule.Name)
    filepath = os.path.join(folder, u"{0}.csv".format(name))
    data = get_schedule_data(schedule)
    with open(filepath, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f, delimiter=';', quotechar='"', quoting=csv.QUOTE_MINIMAL)
        for row in data:
            writer.writerow(row)
    return filepath


def generar_planilla_xlsx(csv_path, folder):
    # Carpeta donde está este script (y donde estarán script_2.py y script_3.py)
    try:
        this_folder = os.path.dirname(os.path.abspath(__file__))
    except NameError:
        this_folder = os.getcwd()

    python_exe = r"C:\Users\Zbook HP\AppData\Local\Programs\Python\Python313\python.exe"
    script_2_path = os.path.join(this_folder, "script_2.py")
    script_3_path = os.path.join(this_folder, "script_3.py")

    # Ejecutar script_2.py
    cmd_2 = [python_exe, script_2_path, csv_path]
    CREATE_NO_WINDOW = 0x08000000
    proc2 = subprocess.Popen(
    cmd_2,
    stdout=subprocess.PIPE,
    stderr=subprocess.PIPE,
    cwd=this_folder,
    creationflags=CREATE_NO_WINDOW
    )
    out2, err2 = proc2.communicate()
    out2_str = out2.decode("utf-8").strip()
    err2_str = err2.decode("utf-8").strip()
    if proc2.returncode != 0:
        return u"Error en ejecución script_2.py: {0}".format(err2_str if err2_str else u"Sin detalles de error")
    if u"Error" in out2_str:
        return u"Error generado en script_2.py: {0}".format(out2_str)

    # Nombre del xlsx resultado
    base_name = os.path.splitext(os.path.basename(csv_path))[0]
    xlsx_path = os.path.join(folder, base_name + ".xlsx")

    # Ejecutar script_3.py sobre el xlsx
    cmd_3 = [python_exe, script_3_path, xlsx_path]
    proc3 = subprocess.Popen(
    cmd_3,
    stdout=subprocess.PIPE,
    stderr=subprocess.PIPE,
    cwd=this_folder,
    creationflags=CREATE_NO_WINDOW
    )
    out3, err3 = proc3.communicate()
    out3_str = out3.decode("utf-8").strip()
    err3_str = err3.decode("utf-8").strip()
    if proc3.returncode != 0:
        return u"Error en ejecución script_3.py: {0}".format(err3_str if err3_str else u"Sin detalles de error")
    if u"Error" in out3_str:
        return u"Error generado en script_3.py: {0}".format(out3_str)

    return u"Listo"


class ScheduleSelector(forms.WPFWindow):
    def __init__(self, xaml_path, schedules):
        # Cargar XAML
        forms.WPFWindow.__init__(self, xaml_path)

        # Guardar schedules Revit reales
        self.schedules = schedules
        # Nombre -> bool (seleccionado)
        self.checked_states = {s.Name: False for s in schedules}
        # Lista base de nombres (para filtrar)
        self.all_names = [s.Name for s in schedules]

        # Referencias a controles del XAML
        self.search_box = self.SearchBox
        self.listbox = self.SchedulesList
        self.btn_select_all = self.BtnSelectAll
        self.btn_clear = self.BtnClearSelection
        self.btn_export = self.BtnExport
        self.btn_cancel = self.BtnCancel

        # Conectar eventos
        self.btn_select_all.Click += self.on_select_all
        self.btn_clear.Click += self.on_clear_selection
        self.btn_export.Click += self.on_export
        self.btn_cancel.Click += self.on_cancel
        self.search_box.TextChanged += self.on_search_text_changed

        # Resultado
        self.selected_schedules = []
        
        # <<< AÑADIR ESTA LÍNEA >>>
        self._populate_list(self.all_names)

    def _populate_list(self, names):
        # Limpiar elementos
        self.listbox.Items.Clear()
        for name in names:
            self.listbox.Items.Add(name)
            # Re-aplicar selección según estado global
            if self.checked_states.get(name, False):
                self.listbox.SelectedItems.Add(name)

    def on_search_text_changed(self, sender, args):
        text = self.search_box.Text
        if not text:
            filtered = self.all_names
        else:
            low = text.lower()
            filtered = [n for n in self.all_names if low in n.lower()]
        # Antes de repoblar, sincronizar estados de los visibles actuales con SelectedItems
        self._sync_from_listbox()
        self._populate_list(filtered)

    def _sync_from_listbox(self):
        # Poner en False todos los que están visibles
        for i in range(self.listbox.Items.Count):
            name = self.listbox.Items[i]
            self.checked_states[name] = False
        # Marcar los seleccionados visibles
        for sel in self.listbox.SelectedItems:
            self.checked_states[sel] = True

    def on_select_all(self, sender, args):
        # Seleccionar todos los visibles
        self.listbox.SelectAll()
        # Sincronizar al diccionario
        self._sync_from_listbox()

    def on_clear_selection(self, sender, args):
        # Limpiar selección
        self.listbox.UnselectAll()
        # Poner en False todos los visibles
        for i in range(self.listbox.Items.Count):
            name = self.listbox.Items[i]
            self.checked_states[name] = False

    def on_export(self, sender, args):
        # Sincronizar seleccionados actuales
        self._sync_from_listbox()

        # Verificar que haya al menos uno
        if not any(self.checked_states.values()):
            forms.alert("Seleccione al menos una tabla.", title="Aviso")
            return

        # Construir lista de schedules seleccionados
        self.selected_schedules = [
            s for s in self.schedules if self.checked_states.get(s.Name, False)
        ]

        if not self.selected_schedules:
            forms.alert("No se encontraron tablas seleccionadas.", title="Aviso")
            return

        self.Close()

    def on_cancel(self, sender, args):
        # No selecciona nada
        self.selected_schedules = []
        self.Close()

def seleccionar_planillas_con_tkinter(schedules):
    """Lanza CPython + Tkinter para seleccionar planillas.

    schedules: lista de ViewSchedule Revit.
    Devuelve lista de ViewSchedule seleccionadas.
    """
    # preparar datos para JSON (nombre + id entero)
    data = []
    for s in schedules:
        try:
            sid = s.Id.IntegerValue
        except:
            sid = -1
        data.append({"name": s.Name, "id": sid})

    # carpeta del botón (donde está script.py)
    try:
        this_folder = os.path.dirname(os.path.abspath(__file__))
    except NameError:
        this_folder = os.getcwd()

    # rutas de JSON temporal y script Tkinter
    input_json = os.path.join(this_folder, "schedules_input.json")
    output_json = os.path.join(this_folder, "schedules_output.json")
    select_ui = os.path.join(this_folder, "select_schedules_ui.py")

    # guardar lista de planillas
    with open(input_json, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    # lanzar CPython + Tkinter
    python_exe = r"C:\Users\Zbook HP\AppData\Local\Programs\Python\Python313\python.exe"
    cmd = [python_exe, select_ui, input_json, output_json]
    CREATE_NO_WINDOW = 0x08000000
    proc = subprocess.Popen(
    cmd,
    stdout=subprocess.PIPE,
    stderr=subprocess.PIPE,
    cwd=this_folder,
    creationflags=CREATE_NO_WINDOW
    )
    out, err = proc.communicate()

    if proc.returncode != 0:
        forms.alert(
            u"Error al abrir ventana Tkinter:\n{}\n{}".format(
                out.decode("utf-8", "ignore"), err.decode("utf-8", "ignore")
            ),
            title="Error"
        )
        return []

    if not os.path.exists(output_json):
        forms.alert("No se generó archivo de selección de tablas.", title="Aviso")
        return []

    with open(output_json, "r", encoding="utf-8") as f:
        result = json.load(f)

    selected_names = set(result.get("selected_names", []))

    if not selected_names:
        return []

    # mapear nombres a objetos ViewSchedule
    seleccionadas = [s for s in schedules if s.Name in selected_names]
    return seleccionadas


def main():
    if doc is None:
        forms.alert("No hay documento Revit activo. Abre un proyecto primero.", title="Error")
        return

    # Obtener schedules (no plantillas)
    collector = FilteredElementCollector(doc).OfClass(ViewSchedule)
    schedules = [s for s in collector if not s.IsTemplate]

    if not schedules:
        forms.alert("No se encontraron tablas para exportar.", title="Aviso")
        return

    # Selección mediante Tkinter (CPython)
    seleccionadas = seleccionar_planillas_con_tkinter(schedules)
    if not seleccionadas:
        forms.alert("Operación cancelada o sin selección.", title="Aviso")
        return

    # Carpeta de exportación fija (igual que antes)
    folder = r"C:\Users\Zbook HP\AppData\Roaming\MyPyRevitExtention\PyRevitIT.extension\Gestion IT.tab\Exportacion de Planillas.panel\Por Seleccion.pushbutton\Planillas"
    if not os.path.exists(folder):
        os.makedirs(folder)

    exported_files = []
    status_msgs = []

    for sched in seleccionadas:
        path_csv = export_schedule_csv(sched, folder)
        exported_files.append(path_csv)
        status = generar_planilla_xlsx(os.path.abspath(path_csv), folder)
        status_msgs.append(u"{0} : {1}".format(os.path.basename(path_csv), status))

    forms.alert(
        u"Se exportaron {0} tablas a CSV en:\n{1}\n\n{2}".format(
            len(exported_files),
            folder,
            u"\n".join(status_msgs)
        ),
        title="Exportación completa"
    )


if __name__ == "__main__":
    main()

