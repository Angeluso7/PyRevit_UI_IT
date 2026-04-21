# -*- coding: utf-8 -*-
"""
combinar_datos.py  —  IronPython / Revit API
Cruza las filas del TXT temporal con el modelo Revit (host + links)
y actualiza el repositorio activo del proyecto.
Llamado desde script.py con combinar_datos.main(repo_tmp_path).
"""
import os
import sys
import json

import clr
from pyrevit import forms

clr.AddReference("RevitAPI")
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInParameter,
    ViewSchedule, RevitLinkInstance
)
from System.Windows.Forms import (
    Form, ListView, ColumnHeader, Button, DialogResult,
    SaveFileDialog, View as WFView, AnchorStyles,
    FormStartPosition
)
from System.Drawing import Size, Point

# ── Rutas centralizadas ──────────────────────────────────────────────────────
_THIS_DIR = os.path.dirname(os.path.abspath(__file__))
_EXT_ROOT = os.path.normpath(os.path.join(_THIS_DIR, "..", "..", "..", "..", ".."))
_LIB_DIR  = os.path.join(_EXT_ROOT, "lib")
if _LIB_DIR not in sys.path:
    sys.path.insert(0, _LIB_DIR)

from config_utils import get_repo_activo_path, get_script_json_path

# ── Documento activo ────────────────────────────────────────────────────────
try:
    doc = __revit__.ActiveUIDocument.Document
except Exception:
    doc = None


# ── Repositorio ─────────────────────────────────────────────────────────────

def _get_repo_path():
    try:
        return get_repo_activo_path()
    except Exception as e:
        forms.alert(u"{}".format(e), title="Sin proyecto configurado")
        return None


def load_repo(repo_path):
    if not repo_path or not os.path.isfile(repo_path):
        return {}
    try:
        with open(repo_path, "r", encoding="utf-8", errors="replace") as f:
            return json.load(f)
    except Exception as e:
        forms.alert(u"Error leyendo repositorio:\n{}".format(e), title="Error repositorio")
        return {}


def save_repo(repo_path, data):
    if not repo_path:
        return
    try:
        folder = os.path.dirname(repo_path)
        if folder and not os.path.exists(folder):
            os.makedirs(folder)
        with open(repo_path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
    except Exception as e:
        forms.alert(u"Error guardando repositorio:\n{}".format(e), title="Error repositorio")


def load_script_json():
    path = get_script_json_path()
    if not os.path.isfile(path):
        forms.alert(u"No se encontró script.json en:\n{}".format(path), title="Error script.json")
        return None
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        forms.alert(u"Error leyendo script.json:\n{}".format(e), title="Error script.json")
        return None


# ── Upsert por CodIntBIM ─────────────────────────────────────────────────────

def upsert_by_codintbim(repo, fila_dict):
    cod = (fila_dict.get("CodIntBIM") or "").strip()
    if not cod:
        return repo

    coincidentes = [
        k for k, v in repo.items()
        if (v.get("CodIntBIM") or "").strip() == cod
    ]

    if coincidentes:
        for k in coincidentes:
            d = repo[k]
            for campo, valor in fila_dict.items():
                if campo in ("Archivo", "ElementId") and campo in d:
                    continue
                d[campo] = valor
    else:
        archivo  = fila_dict.get("Archivo", "")
        elem_id  = str(fila_dict.get("ElementId", "") or "")
        clave    = u"{}_{}".format(archivo, elem_id) if archivo and elem_id else cod
        repo[clave] = dict(fila_dict)

    return repo


# ── Utilidades Revit ─────────────────────────────────────────────────────────

def get_view_schedules_by_name(document):
    res = {}
    try:
        for vs in FilteredElementCollector(document).OfClass(ViewSchedule):
            if not vs.IsTemplate:
                res[vs.Name] = vs
    except Exception as e:
        forms.alert(u"Error obteniendo planillas:\n{}".format(e), title="Error Revit")
    return res


def get_schedule_param_headers(schedule):
    headers = []
    try:
        for fid in schedule.Definition.GetFieldOrder():
            headers.append(schedule.Definition.GetField(fid).GetName())
    except Exception as e:
        forms.alert(
            u"Error leyendo encabezados de '{}':\n{}".format(schedule.Name, e),
            title="Error encabezados"
        )
    return headers


def get_all_docs_with_links():
    docs = []
    try:
        docs.append((doc, doc.PathName or ""))
        seen = set()
        for li in FilteredElementCollector(doc).OfClass(RevitLinkInstance):
            link_doc = li.GetLinkDocument()
            if link_doc is None:
                continue
            path = link_doc.PathName or ""
            if path not in seen:
                seen.add(path)
                docs.append((link_doc, path))
    except Exception as e:
        forms.alert(u"Error obteniendo documentos linkeados:\n{}".format(e), title="Error links")
    return docs


def find_elements_by_codintbim_all_docs(codintbim):
    resultados = []
    for d, ruta in get_all_docs_with_links():
        try:
            for el in FilteredElementCollector(d).WhereElementIsNotElementType():
                try:
                    p = el.LookupParameter("CodIntBIM")
                    if p and p.AsString() == codintbim:
                        resultados.append((d, ruta, el))
                except Exception:
                    continue
        except Exception:
            continue
    return resultados


# ── Ventana de comprobación (WinForms) ───────────────────────────────────────

class ComprobacionForm(Form):
    def __init__(self, registros):
        self.Text            = u"Comprobación de Códigos"
        self.Size            = Size(660, 460)
        self.MinimumSize     = Size(500, 350)
        self.StartPosition   = FormStartPosition.CenterScreen

        self.listview = ListView()
        self.listview.View          = WFView.Details
        self.listview.FullRowSelect = True
        self.listview.GridLines     = True
        self.listview.Anchor        = (
            AnchorStyles.Top | AnchorStyles.Bottom |
            AnchorStyles.Left | AnchorStyles.Right
        )
        self.listview.Location = Point(10, 10)
        self.listview.Size     = Size(620, 360)

        col_cod      = ColumnHeader(); col_cod.Text = "CodIntBIM";  col_cod.Width = 230
        col_sit      = ColumnHeader(); col_sit.Text = u"Situación"; col_sit.Width = 360
        self.listview.Columns.Add(col_cod)
        self.listview.Columns.Add(col_sit)

        for cod, sit in registros:
            item = self.listview.Items.Add(cod)
            item.SubItems.Add(sit)

        self.btn_exportar        = Button()
        self.btn_exportar.Text   = "Exportar"
        self.btn_exportar.Size   = Size(90, 30)
        self.btn_exportar.Anchor = AnchorStyles.Bottom | AnchorStyles.Right
        self.btn_exportar.Click += self.on_exportar

        self.btn_aceptar        = Button()
        self.btn_aceptar.Text   = "Aceptar"
        self.btn_aceptar.Size   = Size(90, 30)
        self.btn_aceptar.Anchor = AnchorStyles.Bottom | AnchorStyles.Right
        self.btn_aceptar.Click += self.on_aceptar

        self.Resize += self.on_resize
        self.Controls.Add(self.listview)
        self.Controls.Add(self.btn_exportar)
        self.Controls.Add(self.btn_aceptar)
        self.registros = registros
        self._reposicionar_botones()

    def _reposicionar_botones(self):
        self.btn_exportar.Location = Point(self.ClientSize.Width - 200, self.ClientSize.Height - 45)
        self.btn_aceptar.Location  = Point(self.ClientSize.Width - 100, self.ClientSize.Height - 45)

    def on_resize(self, sender, args):
        self.listview.Size = Size(self.ClientSize.Width - 20, self.ClientSize.Height - 80)
        self._reposicionar_botones()

    def on_exportar(self, sender, args):
        if not self.registros:
            forms.alert(u"No hay datos para exportar.", title="Información")
            return
        dlg = SaveFileDialog()
        dlg.Title       = u"Exportar comprobación de códigos"
        dlg.Filter      = "Archivo de texto (*.txt)|*.txt|Todos los archivos (*.*)|*.*"
        dlg.DefaultExt  = "txt"
        dlg.FileName    = "ComprobacionCodigos.txt"
        if dlg.ShowDialog() == DialogResult.OK:
            try:
                with open(dlg.FileName, "w", encoding="utf-8") as f:
                    f.write(u"CodIntBIM;Situación\n")
                    for cod, sit in self.registros:
                        f.write(u"{};{}\n".format(cod, sit))
                forms.alert(u"Exportado en:\n{}".format(dlg.FileName), title=u"Exportación OK")
            except Exception as e:
                forms.alert(u"Error exportando:\n{}".format(e), title="Error exportación")

    def on_aceptar(self, sender, args):
        self.Close()


def mostrar_comprobacion(registros):
    if registros:
        ComprobacionForm(registros).ShowDialog()


# ── Main ─────────────────────────────────────────────────────────────────────

def main(repo_tmp_path):
    if doc is None:
        forms.alert(u"No hay documento activo.", title="Error")
        return

    if not os.path.isfile(repo_tmp_path):
        forms.alert(
            u"No se encontró el archivo temporal:\n{}".format(repo_tmp_path),
            title="Sin datos"
        )
        return

    repo_path = _get_repo_path()
    if not repo_path:
        return

    script_data = load_script_json()
    if script_data is None:
        return

    codigos_planillas = script_data.get("codigos_planillas", {})
    if not codigos_planillas:
        forms.alert(u"'codigos_planillas' vacío en script.json.", title="Error script.json")
        return

    schedules_by_name = get_view_schedules_by_name(doc)
    if not schedules_by_name:
        forms.alert(u"No hay planillas de planificación en el modelo activo.", title="Sin planillas")
        return

    repo = load_repo(repo_path)

    try:
        with open(repo_tmp_path, "r", encoding="utf-8") as f:
            lineas = [ln.strip() for ln in f if ln.strip()]
    except Exception as e:
        forms.alert(u"Error leyendo archivo temporal:\n{}".format(e), title="Error")
        return

    filas_total = filas_con_planilla = elementos_afectados = 0
    registros_codigos = []

    for linea in lineas:
        filas_total += 1
        partes     = linea.split(";")
        codintbim  = partes[0].strip() if partes else ""
        if not codintbim or len(codintbim) < 4:
            continue

        pref4          = codintbim[:4]
        nombre_planilla = None
        for key, val in codigos_planillas.items():
            if str(val).startswith(pref4):
                nombre_planilla = key
                break

        if not nombre_planilla:
            continue

        schedule = schedules_by_name.get(nombre_planilla)
        if schedule is None:
            continue

        filas_con_planilla += 1
        headers = get_schedule_param_headers(schedule)
        if not headers:
            continue

        fila_vals = list(partes)
        if len(fila_vals) < len(headers):
            fila_vals += [""] * (len(headers) - len(fila_vals))
        else:
            fila_vals = fila_vals[:len(headers)]

        header_to_value = dict(zip(headers, fila_vals))
        elems_info      = find_elements_by_codintbim_all_docs(codintbim)

        if not elems_info:
            registros_codigos.append((codintbim, u"Código no encontrado en modelo"))
            continue

        for doc_elem, ruta_archivo, el in elems_info:
            elementos_afectados += 1
            fila_dict = {
                "Archivo":   ruta_archivo,
                "ElementId": str(el.Id.IntegerValue),
                "CodIntBIM": codintbim,
                "nombre_archivo": os.path.basename(ruta_archivo) if ruta_archivo else ""
            }

            try:
                host_p = el.get_Parameter(BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM)
                if host_p:
                    fila_dict[u"Anfitrión"] = host_p.AsValueString() or ""
            except Exception:
                pass

            for hdr in headers:
                val_excel = str(header_to_value.get(hdr, "") or "").strip()
                param_val = ""
                try:
                    p_el = el.LookupParameter(hdr)
                    if p_el:
                        param_val = p_el.AsString() or p_el.AsValueString() or ""
                except Exception:
                    pass
                fila_dict[hdr] = val_excel if val_excel else param_val

            repo = upsert_by_codintbim(repo, fila_dict)

    save_repo(repo_path, repo)

    forms.alert(
        u"Combinación completada.\n\n"
        u"Filas leídas del Excel:          {}\n"
        u"Filas que encontraron planilla:  {}\n"
        u"Elementos afectados:             {}\n\n"
        u"Repositorio:\n{}".format(
            filas_total, filas_con_planilla, elementos_afectados, repo_path
        ),
        title=u"Combinación completada"
    )

    mostrar_comprobacion(registros_codigos)