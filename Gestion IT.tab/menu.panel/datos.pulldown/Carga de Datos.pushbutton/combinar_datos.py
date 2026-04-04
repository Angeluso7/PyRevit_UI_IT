# -*- coding: utf-8 -*-

import os
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
    SaveFileDialog, DockStyle, View as WFView, AnchorStyles,
    FormStartPosition
)

from System.Drawing import Size, Point

doc = __revit__.ActiveUIDocument.Document

# Carpeta común de datos
DATA_DIR = os.path.join(
    os.path.expanduser("~"),
    r"AppData\Roaming\MyPyRevitExtention\PyRevitIT.extension\data"
)

SCRIPT_JSON_PATH = os.path.join(DATA_DIR, "script.json")
CONFIG_PROYECTO_ACTIVO = os.path.join(DATA_DIR, "config_proyecto_activo.json")


# ----------------- Ruta de BD desde config_proyecto_activo -----------------

def get_repo_path_from_config():
    """
    Lee config_proyecto_activo.json, toma 'ruta_repositorio_activo'
    y la devuelve tal cual (ruta completa del repositorio).
    """
    if not os.path.exists(CONFIG_PROYECTO_ACTIVO):
        forms.alert(
            "No se encontró config_proyecto_activo.json en:\n{}".format(CONFIG_PROYECTO_ACTIVO),
            title="Error config"
        )
        return None

    try:
        with open(CONFIG_PROYECTO_ACTIVO, "r", encoding="utf-8") as f:
            cfg = json.load(f)
    except Exception as e:
        forms.alert(
            "Error leyendo config_proyecto_activo.json:\n{}".format(e),
            title="Error config"
        )
        return None

    ruta = (cfg.get("ruta_repositorio_activo") or "").strip()
    if not ruta:
        forms.alert(
            "En config_proyecto_activo.json no se encontró 'ruta_repositorio_activo' o está vacío.",
            title="Config incompleta"
        )
        return None

    return ruta


REPO_PATH = get_repo_path_from_config()


def load_repo():
    """Carga la BD desde el archivo definido en config_proyecto_activo.json."""
    if not REPO_PATH:
        return {}
    try:
        if os.path.exists(REPO_PATH):
            # Tolerar caracteres no UTF-8
            with open(REPO_PATH, "r", encoding="utf-8", errors="replace") as f:
                return json.load(f)
    except Exception as e:
        forms.alert(
            "Error leyendo repositorio de datos:\n{}\nSe usará un repositorio vacío."
            .format(e),
            title="Error repositorio"
        )
    return {}


def save_repo(data):
    """Guarda la BD en el archivo definido en config_proyecto_activo.json."""
    if not REPO_PATH:
        return
    try:
        folder = os.path.dirname(REPO_PATH)
        if folder and not os.path.exists(folder):
            os.makedirs(folder)
        with open(REPO_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
    except Exception as e:
        forms.alert(
            "Error guardando repositorio de datos:\n{}".format(e),
            title="Error repositorio"
        )


def load_script_json():
    """Carga script.json (para codigos_planillas, etc.)."""
    if not os.path.exists(SCRIPT_JSON_PATH):
        forms.alert(
            "No se encontró script.json en:\n{}".format(SCRIPT_JSON_PATH),
            title="Error script.json"
        )
        return None
    try:
        with open(SCRIPT_JSON_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        forms.alert(
            "Error leyendo script.json:\n{}".format(e),
            title="Error script.json"
        )
        return None


# ----------------- Upsert por CodIntBIM -----------------

def upsert_by_codintbim(repo, fila_dict):
    """
    Actualiza o inserta una entrada en repo según CodIntBIM.

    - Si existe alguna entrada con el mismo CodIntBIM, se actualizan sus campos.
      No se sobrescriben Archivo ni ElementId si ya existen en el registro.
    - Si no existe, se crea una nueva entrada con clave Archivo_ElementId
      (o el propio código si faltan esos datos).
    """
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
        archivo = fila_dict.get("Archivo", "")
        elem_id = str(fila_dict.get("ElementId", "") or "")
        clave = u"{}_{}".format(archivo, elem_id) if archivo and elem_id else cod
        repo[clave] = dict(fila_dict)

    return repo


# ----------------- Utilidades Revit -----------------

def get_view_schedules_by_name(document):
    res = {}
    try:
        col = FilteredElementCollector(document).OfClass(ViewSchedule)
        for vs in col:
            if vs.IsTemplate:
                continue
            res[vs.Name] = vs
    except Exception as e:
        forms.alert(
            "Error obteniendo tablas de planificación:\n{}".format(e),
            title="Error Revit"
        )
    return res


def get_schedule_param_headers(schedule):
    """Devuelve lista de nombres de parámetros (no texto gráfico) en orden de columnas."""
    headers = []
    try:
        field_ids = schedule.Definition.GetFieldOrder()
        for fid in field_ids:
            field = schedule.Definition.GetField(fid)
            pname = field.GetName()
            headers.append(pname)
    except Exception as e:
        forms.alert(
            "Error obteniendo encabezados de parámetros para '{}':\n{}"
            .format(schedule.Name, e),
            title="Error encabezados"
        )
    return headers


def get_all_docs_with_links():
    docs = []
    try:
        main_doc = doc
        docs.append((main_doc, main_doc.PathName or ""))
        link_instances = FilteredElementCollector(main_doc).OfClass(RevitLinkInstance)
        seen = set()
        for li in link_instances:
            link_doc = li.GetLinkDocument()
            if link_doc is None:
                continue
            path = link_doc.PathName or ""
            if path in seen:
                continue
            seen.add(path)
            docs.append((link_doc, path))
    except Exception as e:
        forms.alert(
            "Error obteniendo documentos linkeados:\n{}".format(e),
            title="Error links"
        )
    return docs


def find_elements_by_codintbim_all_docs(codintbim):
    resultados = []
    for d, ruta in get_all_docs_with_links():
        try:
            collector = FilteredElementCollector(d).WhereElementIsNotElementType()
            for el in collector:
                try:
                    p = el.LookupParameter("CodIntBIM")
                    if p and p.AsString() == codintbim:
                        resultados.append((d, ruta, el))
                except Exception:
                    continue
        except Exception:
            continue
    return resultados


# ----------------- Ventana ListView de comprobación -----------------

class ComprobacionForm(Form):
    def __init__(self, registros):
        self.Text = "Comprobación de Códigos"
        self.Size = Size(650, 450)
        self.MinimumSize = Size(500, 350)
        self.StartPosition = FormStartPosition.CenterScreen

        self.listview = ListView()
        self.listview.View = WFView.Details
        self.listview.FullRowSelect = True
        self.listview.GridLines = True
        self.listview.Anchor = (
            AnchorStyles.Top | AnchorStyles.Bottom |
            AnchorStyles.Left | AnchorStyles.Right
        )
        self.listview.Location = Point(10, 10)
        self.listview.Size = Size(610, 350)

        col_cod = ColumnHeader()
        col_cod.Text = "CodIntBIM"
        col_cod.Width = 220

        col_sit = ColumnHeader()
        col_sit.Text = "Situación"
        col_sit.Width = 360

        self.listview.Columns.Add(col_cod)
        self.listview.Columns.Add(col_sit)

        for cod, sit in registros:
            item = self.listview.Items.Add(cod)
            item.SubItems.Add(sit)

        self.btn_exportar = Button()
        self.btn_exportar.Text = "Exportar"
        self.btn_exportar.Size = Size(90, 30)
        self.btn_exportar.Anchor = AnchorStyles.Bottom | AnchorStyles.Right
        self.btn_exportar.Location = Point(self.ClientSize.Width - 200, self.ClientSize.Height - 50)
        self.btn_exportar.Click += self.on_exportar

        self.btn_aceptar = Button()
        self.btn_aceptar.Text = "Aceptar"
        self.btn_aceptar.Size = Size(90, 30)
        self.btn_aceptar.Anchor = AnchorStyles.Bottom | AnchorStyles.Right
        self.btn_aceptar.Location = Point(self.ClientSize.Width - 100, self.ClientSize.Height - 50)
        self.btn_aceptar.Click += self.on_aceptar

        self.Resize += self.on_resize

        self.Controls.Add(self.listview)
        self.Controls.Add(self.btn_exportar)
        self.Controls.Add(self.btn_aceptar)

        self.registros = registros

    def on_resize(self, sender, args):
        self.listview.Size = Size(self.ClientSize.Width - 20, self.ClientSize.Height - 80)
        self.btn_exportar.Location = Point(self.ClientSize.Width - 200, self.ClientSize.Height - 50)
        self.btn_aceptar.Location = Point(self.ClientSize.Width - 100, self.ClientSize.Height - 50)

    def on_exportar(self, sender, args):
        if not self.registros:
            forms.alert("No hay datos para exportar.", title="Información")
            return

        dlg = SaveFileDialog()
        dlg.Title = "Exportar comprobación de códigos"
        dlg.Filter = "Archivo de texto (*.txt)|*.txt|Todos los archivos (*.*)|*.*"
        dlg.DefaultExt = "txt"
        dlg.FileName = "ComprobacionCodigos.txt"

        if dlg.ShowDialog() == DialogResult.OK:
            filename = dlg.FileName
            try:
                with open(filename, "w", encoding="utf-8") as f:
                    f.write("CodIntBIM;Situación\n")
                    for cod, sit in self.registros:
                        linea = u"{};{}\n".format(cod, sit)
                        f.write(linea)
                forms.alert(
                    "Archivo exportado en:\n{}".format(filename),
                    title="Exportación completada"
                )
            except Exception as e:
                forms.alert(
                    "No se pudo exportar el archivo:\n{}".format(e),
                    title="Error de exportación"
                )

    def on_aceptar(self, sender, args):
        self.Close()


def mostrar_comprobacion_codigos(registros):
    if not registros:
        return
    form = ComprobacionForm(registros)
    form.ShowDialog()


# ----------------- Main -----------------

def main(repo_tmp_path):
    if doc is None:
        forms.alert("No hay documento activo.", title="Error")
        return

    if not os.path.exists(repo_tmp_path):
        forms.alert(
            "No se encontró el repositorio temporal con las filas de Excel.\n\nRuta:\n{}"
            .format(repo_tmp_path),
            title="Sin datos"
        )
        return

    if not REPO_PATH:
        return

    script_data = load_script_json()
    if script_data is None:
        return

    codigos_planillas = script_data.get("codigos_planillas", {})
    if not codigos_planillas:
        forms.alert(
            "En script.json no se encontró 'codigos_planillas'.",
            title="Error script.json"
        )
        return

    schedules_by_name = get_view_schedules_by_name(doc)
    if not schedules_by_name:
        forms.alert(
            "No se encontraron tablas de planificación en el modelo activo.",
            title="Sin planillas"
        )
        return

    repo = load_repo()

    filas_total = 0
    filas_con_planilla = 0
    elementos_afectados = 0
    registros_codigos = []

    try:
        with open(repo_tmp_path, "r", encoding="utf-8") as f:
            lineas = [ln.strip() for ln in f.readlines() if ln.strip()]
    except Exception as e:
        forms.alert(
            "Error leyendo el repositorio temporal:\n{}".format(e),
            title="Error temporal"
        )
        return

    for linea in lineas:
        filas_total += 1
        partes = linea.split(";")
        if not partes:
            continue

        codintbim = partes[0].strip()
        if not codintbim or len(codintbim) < 4:
            continue

        pref4 = codintbim[:4]
        nombre_planilla = None

        for key, val in codigos_planillas.items():
            try:
                if isinstance(val, str) and val.startswith(pref4):
                    nombre_planilla = key
                    break
            except Exception:
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

        fila_vals = partes
        if len(fila_vals) < len(headers):
            fila_vals = list(fila_vals) + ["" for _ in range(len(headers) - len(fila_vals))]
        else:
            fila_vals = list(fila_vals[:len(headers)])

        header_to_value = {hdr: val for hdr, val in zip(headers, fila_vals)}

        elems_info = find_elements_by_codintbim_all_docs(codintbim)

        if not elems_info:
            registros_codigos.append((codintbim, "Código no encontrado"))
            continue

        for doc_elem, ruta_archivo, el in elems_info:
            elementos_afectados += 1
            elem_id_int = el.Id.IntegerValue
            elem_id_str = str(elem_id_int)

            archivo = ruta_archivo

            fila_dict = {
                "Archivo": archivo,
                "ElementId": elem_id_str,
                "CodIntBIM": codintbim
            }

            try:
                host_param = el.get_Parameter(BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM)
                if host_param:
                    fila_dict["Anfitrión"] = host_param.AsValueString() or ""
            except Exception:
                pass

            try:
                nombre_archivo = os.path.basename(archivo) if archivo else ""
                fila_dict["nombre_archivo"] = nombre_archivo
            except Exception:
                pass

            for hdr in headers:
                val_excel = header_to_value.get(hdr, "")
                if val_excel is None:
                    val_excel_str = ""
                else:
                    val_excel_str = str(val_excel).strip()

                param_val_model = ""
                try:
                    p_el = el.LookupParameter(hdr)
                    if p_el:
                        param_val_model = p_el.AsString() or p_el.AsValueString() or ""
                except Exception:
                    param_val_model = ""

                valor_final = val_excel_str if val_excel_str != "" else param_val_model
                fila_dict[hdr] = valor_final

            repo = upsert_by_codintbim(repo, fila_dict)

    save_repo(repo)

    msg = (
        "Combinación completada.\n\n"
        "Filas leídas del Excel: {}\n"
        "Filas que encontraron planilla: {}\n"
        "Elementos afectados (modelo activo + links): {}\n"
        "Repositorio:\n{}"
    ).format(filas_total, filas_con_planilla, elementos_afectados, REPO_PATH or "(no definido)")

    forms.alert(msg, title="Combinación completada")

    mostrar_comprobacion_codigos(registros_codigos)
