# -*- coding: utf-8 -*-
"""
Por Seleccion.pushbutton / script.py  (IronPython)
Exporta planillas Revit seleccionadas a XLSX con formato.
"""
import os
import csv
import re
import sys
import json
import subprocess
import clr

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import ViewSchedule, SectionType, FilteredElementCollector
from pyrevit import forms

# ── lib al path ──────────────────────────────────────────────────────────────
try:
    _btn_dir  = os.path.dirname(os.path.abspath(__file__))
    _ext_root = os.path.normpath(os.path.join(_btn_dir, '..', '..', '..'))
    _lib_path = os.path.join(_ext_root, 'lib')
    if _lib_path not in sys.path:
        sys.path.insert(0, _lib_path)
except Exception:
    pass

from config_utils import TEMP_DIR, MASTER_DIR
from core.env_config import get_python_exe

# ── documento activo ─────────────────────────────────────────────────────────
doc = __revit__.ActiveUIDocument.Document

CREATE_NO_WINDOW = 0x08000000

# ── helpers ───────────────────────────────────────────────────────────────────
def sanitize_filename(name):
    if isinstance(name, tuple):
        name = name[0]
    if not isinstance(name, str):
        name = str(name)
    return re.sub(r'[\\/*?:"<>|]', '_', name[:50])


def get_schedule_data(schedule):
    data = []
    body = schedule.GetTableData().GetSectionData(SectionType.Body)
    for r in range(body.NumberOfRows):
        row = [schedule.GetCellText(SectionType.Body, r, c) or ''
               for c in range(body.NumberOfColumns)]
        if any(cell.strip() for cell in row):
            data.append(row)
    return data


def export_schedule_csv(schedule, folder):
    name = sanitize_filename(schedule.Name)
    filepath = os.path.join(folder, u'{0}.csv'.format(name))
    data = get_schedule_data(schedule)
    with open(filepath, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f, delimiter=';',
                            quotechar='"', quoting=csv.QUOTE_MINIMAL)
        for row in data:
            writer.writerow(row)
    return filepath


def _run_cpython(python_exe, script_path, *args):
    """Ejecuta un script CPython y devuelve (returncode, stdout, stderr)."""
    cmd = [python_exe, script_path] + list(args)
    proc = subprocess.Popen(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        creationflags=CREATE_NO_WINDOW
    )
    out, err = proc.communicate()
    return (proc.returncode,
            out.decode('utf-8', 'ignore').strip(),
            err.decode('utf-8', 'ignore').strip())


def generar_planilla_xlsx(csv_path, xlsx_folder, python_exe, lib_services):
    """CSV → XLSX → post-proceso de formato. Devuelve mensaje de estado."""
    script_csv2xlsx = os.path.join(lib_services, 'csv_to_xlsx.py')
    script_format   = os.path.join(lib_services, 'format_xlsx_schedules.py')

    # Paso 1: CSV → XLSX
    rc, out, err = _run_cpython(python_exe, script_csv2xlsx, csv_path)
    if rc != 0 or 'Error' in out:
        return u'Error csv_to_xlsx: {0}'.format(err or out)

    base     = os.path.splitext(os.path.basename(csv_path))[0]
    xlsx_path = os.path.join(xlsx_folder, base + '.xlsx')

    # Paso 2: formato XLSX
    rc, out, err = _run_cpython(python_exe, script_format, xlsx_path)
    if rc != 0 or 'Error' in out:
        return u'Error format_xlsx: {0}'.format(err or out)

    return u'OK'


# ── selección mediante Tkinter ─────────────────────────────────────────────
def seleccionar_planillas(schedules, python_exe):
    """
    Lanza la UI Tkinter (CPython) para seleccionar planillas.
    Devuelve lista de ViewSchedule seleccionados.
    """
    lib_ui_path = os.path.join(_lib_path, 'ui', 'schedule_selector_ui.py')

    # Datos para JSON de entrada: lista de {name, id}
    data_in = [{'name': s.Name, 'id': s.Id.IntegerValue} for s in schedules]

    input_json  = os.path.join(TEMP_DIR, 'schedules_input.json')
    output_json = os.path.join(TEMP_DIR, 'schedules_output.json')

    # Limpiar output previo
    if os.path.exists(output_json):
        os.remove(output_json)

    with open(input_json, 'w', encoding='utf-8') as f:
        json.dump(data_in, f, ensure_ascii=False, indent=2)

    rc, out, err = _run_cpython(
        python_exe, lib_ui_path,
        input_json, output_json,
        u'Seleccionar planillas a exportar'
    )

    if rc != 0:
        forms.alert(
            u'Error al abrir ventana de selección:\n{0}\n{1}'.format(out, err),
            title=u'Error'
        )
        return []

    if not os.path.exists(output_json):
        return []

    with open(output_json, 'r', encoding='utf-8') as f:
        result = json.load(f)

    if result.get('opcion') != 'aceptar':
        return []

    selected_names = set(result.get('seleccion', []))
    return [s for s in schedules if s.Name in selected_names]


# ── main ──────────────────────────────────────────────────────────────────────
def main():
    if doc is None:
        forms.alert(u'No hay documento Revit activo.', title=u'Error')
        return

    # 1. Obtener Python y rutas de servicios
    python_exe   = get_python_exe()
    lib_services = os.path.join(_lib_path, 'services')

    if not python_exe or not os.path.exists(python_exe):
        forms.alert(
            u'No se encontró el ejecutable de Python CPython.\n'
            u'Configura la ruta en lib/core/env_config.py',
            title=u'Error'
        )
        return

    # 2. Recolectar schedules del modelo
    schedules = [s for s in
                 FilteredElementCollector(doc).OfClass(ViewSchedule)
                 if not s.IsTemplate]

    if not schedules:
        forms.alert(u'No hay planillas en el modelo.', title=u'Aviso')
        return

    # 3. Selección en UI Tkinter
    seleccionadas = seleccionar_planillas(schedules, python_exe)
    if not seleccionadas:
        forms.alert(u'Operación cancelada o sin selección.', title=u'Aviso')
        return

    # 4. Carpeta de exportación
    export_folder = os.path.join(MASTER_DIR, 'Planillas_Exportadas', 'Por_Seleccion')
    if not os.path.exists(export_folder):
        os.makedirs(export_folder)

    # 5. Exportar cada planilla
    resultados = []
    for sched in seleccionadas:
        csv_path = export_schedule_csv(sched, export_folder)
        estado   = generar_planilla_xlsx(
            os.path.abspath(csv_path), export_folder, python_exe, lib_services
        )
        resultados.append(u'{0}: {1}'.format(sched.Name, estado))

    forms.alert(
        u'Exportación completada ({0} planillas):\n\n{1}\n\nCarpeta:\n{2}'.format(
            len(seleccionadas),
            u'\n'.join(resultados),
            export_folder
        ),
        title=u'Por Selección — Completado'
    )


if __name__ == '__main__':
    main()