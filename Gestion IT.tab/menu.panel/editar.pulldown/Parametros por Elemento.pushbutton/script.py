# -*- coding: utf-8 -*-
__title__ = "Parametros por elemento"
__doc__ = """Version = 1.1
Date = 19.04.2026
________________________________________________________________
Description:

Editor de parámetros por elemento sobre modelo linkeado,
usando planillas y un repositorio JSON.

________________________________________________________________
Last Updates:
- [19.04.2026] v1.1  Rutas centralizadas via config.paths.
________________________________________________________________
Author: Erik Frits + ajustes Angeluso
"""

# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝
#==================================================
import unicodedata
import clr
import os
import sys
import json
import threading
import time

from Autodesk.Revit.DB import Transaction, ViewSchedule, FilteredElementCollector
from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler
from Autodesk.Revit.UI.Selection import ObjectType

from pyrevit import forms

# ------------------------------------------------------------
# Normalización de texto (acentos, ñ)
# ------------------------------------------------------------

def normalizar_clave(texto):
    """Quita acentos/virgulillas y convierte ñ/Ñ en n/N."""
    if texto is None:
        return None
    if not isinstance(texto, str):
        texto = str(texto)
    nfkd = unicodedata.normalize('NFKD', texto)
    sin_tildes = u"".join(c for c in nfkd if not unicodedata.combining(c))
    sin_enies = sin_tildes.replace(u"ñ", u"n").replace(u"Ñ", u"N")
    return sin_enies


# ------------------------------------------------------------
# Rutas centralizadas desde config.paths
# ------------------------------------------------------------
try:
    _this_dir = os.path.dirname(os.path.abspath(__file__))
except Exception:
    _this_dir = os.getcwd()

# pushbutton(1) -> pulldown(2) -> panel(3) -> tab(4) -> EXT_ROOT
_EXT_ROOT = os.path.abspath(os.path.join(_this_dir, '..', '..', '..', '..'))
_LIB_DIR  = os.path.join(_EXT_ROOT, 'lib')
if _LIB_DIR not in sys.path:
    sys.path.insert(0, _LIB_DIR)

try:
    from config.paths import DATA_DIR, MASTER_DIR, TEMP_DIR, CACHE_DIR, \
                             CONFIG_PROYECTO, REGISTRO_PROYECTOS, \
                             SCRIPT_JSON_PATH_LIB, ensure_runtime_dirs
    ensure_runtime_dirs()
except Exception as _path_err:
    _DATA_DIR            = os.path.join(_EXT_ROOT, 'data')
    DATA_DIR             = _DATA_DIR
    MASTER_DIR           = os.path.join(_DATA_DIR, 'master')
    TEMP_DIR             = os.path.join(_DATA_DIR, 'temp')
    CACHE_DIR            = os.path.join(_DATA_DIR, 'cache')
    CONFIG_PROYECTO      = os.path.join(MASTER_DIR, 'config_proyecto_activo.json')
    REGISTRO_PROYECTOS   = os.path.join(MASTER_DIR, 'registro_proyectos.json')
    SCRIPT_JSON_PATH_LIB = os.path.join(MASTER_DIR, 'script.json')

# Alias locales para compatibilidad con el resto del script
CONFIG_PATH      = CONFIG_PROYECTO
SCRIPT_JSON_PATH = SCRIPT_JSON_PATH_LIB


# ------------------------------------------------------------
# Funciones de repositorio
# ------------------------------------------------------------

def load_active_repo_path():
    if not os.path.exists(CONFIG_PATH):
        raise Exception(
            u"No se encontró config_proyecto_activo.json en:\n{}".format(CONFIG_PATH)
        )
    with open(CONFIG_PATH, "r") as f:
        cfg = json.load(f)

    ruta_repo = (cfg.get("ruta_repositorio_activo") or "").strip()
    if not ruta_repo:
        raise Exception(
            u"config_proyecto_activo.json no tiene 'ruta_repositorio_activo'."
        )
    return ruta_repo


def load_repo():
    try:
        bd_path = load_active_repo_path()
    except Exception as e:
        forms.alert(
            u"Error obteniendo ruta del repositorio:\n{}".format(e),
            title="Error repositorio",
        )
        return {}

    if not os.path.exists(bd_path):
        return {}

    try:
        with open(bd_path, "r") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception as e:
        forms.alert(
            u"Error leyendo repositorio:\n{}\nSe usará uno vacío.".format(e),
            title="Error repositorio",
        )
        return {}


def save_repo(repo_dict):
    try:
        bd_path = load_active_repo_path()
    except Exception as e:
        forms.alert(
            u"Error obteniendo ruta del repositorio para guardar:\n{}".format(e),
            title="Error repositorio",
        )
        return

    try:
        with open(bd_path, "w") as f:
            json.dump(repo_dict, f, indent=4, ensure_ascii=False)
    except Exception as e:
        forms.alert(
            u"Error guardando repositorio:\n{}".format(e),
            title="Error repositorio",
        )


def load_script_json():
    if not os.path.exists(SCRIPT_JSON_PATH):
        forms.alert(
            u"No se encontró script.json en:\n{}".format(SCRIPT_JSON_PATH),
            title="Error script.json",
        )
        return None
    try:
        with open(SCRIPT_JSON_PATH, "r") as f:
            return json.load(f)
    except Exception as e:
        forms.alert(
            u"Error leyendo script.json:\n{}\n\nRevisa sintaxis JSON.".format(e),
            title="Error script.json",
        )
        return None


# ------------------------------------------------------------
# Clases de datos y handler
# ------------------------------------------------------------

class ParamItem(object):
    def __init__(self, name, value, revit_param):
        self.Name = name
        self.Value = value
        self.RevitParam = revit_param


class UpdateParamsHandler(IExternalEventHandler):
    def __init__(self):
        self._editor = None

    def Execute(self, uiapp):
        if self._editor:
            self._editor.on_update_complete(True, u"Actualización completada.")

    def GetName(self):
        return "UpdateParamsHandler"


# ------------------------------------------------------------
# Editor WPF (pyRevit.forms.WPFWindow)
# ------------------------------------------------------------

class ParamsEditor(forms.WPFWindow):
    def __init__(self, linked_doc, element, external_event, handler):
        xaml_path = os.path.join(os.path.dirname(__file__), "ui.xaml")
        forms.WPFWindow.__init__(self, xaml_path)

        self._doc = linked_doc
        self.element = element
        self._external_event = external_event
        self._handler = handler
        if self._handler:
            self._handler._editor = self

        # Datos auxiliares
        self._repo = load_repo()
        self._script_data = load_script_json() or {}

        # Diccionario original de reemplazos y versión normalizada
        self._reemplazos_raw = self._script_data.get("reemplazos_de_nombres", {})
        self._reemplazos = {
            normalizar_clave(k): v
            for k, v in self._reemplazos_raw.items()
        }

        # Colección de parámetros (lista python)
        self.params = []
        self.paramsListView.ItemsSource = self.params

        self._original_values = {}
        self._changes_saved = False
        self._background_updater = None

        # Clave base para este elemento (file + id)
        self._key = self._build_key()

        # Botones
        self.btnSave.Click += self.on_save
        self.btnCancel.Click += self.on_cancel

        # Cargar parámetros según planilla
        try:
            self.load_parameters_from_repo_or_model()
        except Exception as e:
            forms.alert(
                u"Error inicializando editor de parámetros:\n{}".format(e),
                title="Error editor",
            )

    # ---------------- helpers de clave y archivo ----------------

    def _build_key(self):
        file_path = self._doc.PathName or "unknown_file"
        eid = self.element.Id.IntegerValue
        return u"{}_{}".format(file_path, eid)

    def _get_archivo_y_id(self):
        """Devuelve (archivo_completo, element_id_str) según la clave."""
        base_key = self._key.rsplit("_", 1)[0]
        archivo_completo = base_key
        element_id_str = str(self.element.Id.IntegerValue)
        return archivo_completo, element_id_str

    # ---------------- lógica de planilla y encabezados ---------

    def _get_headers_order(self):
        """Obtiene clave de planilla y encabezados en orden, usando CodIntBIM y script.json."""
        try:
            p_cod = self.element.LookupParameter("CodIntBIM")
            cod_val = p_cod.AsString() if p_cod and p_cod.HasValue else ""
        except Exception as e:
            forms.alert(
                u"Error leyendo CodIntBIM:\n{}".format(e),
                title="Error CodIntBIM",
            )
            return None, []

        if not cod_val or len(cod_val) < 4:
            forms.alert(
                u"CodIntBIM vacío o con menos de 4 caracteres.\nValor: '{}'".format(cod_val),
                title="CodIntBIM inválido",
            )
            return None, []

        pref_cod = cod_val[:4]

        script_data = self._script_data or load_script_json() or {}
        codigos_planillas = script_data.get("codigos_planillas", {})
        if not codigos_planillas:
            forms.alert(
                u"En script.json no se encontró 'codigos_planillas' o está vacío.",
                title="Error script.json",
            )
            return None, []

        clave_planilla = None
        try:
            for clave, vals in codigos_planillas.items():
                if isinstance(vals, list) and any(
                    isinstance(v, str) and v.startswith(pref_cod) for v in vals
                ):
                    clave_planilla = clave
                    break
                elif isinstance(vals, str) and vals.startswith(pref_cod):
                    clave_planilla = clave
                    break
                else:
                    vals_str = str(vals)
                    if vals_str.startswith(pref_cod):
                        clave_planilla = clave
                        break
        except Exception as e:
            forms.alert(
                u"Error procesando 'codigos_planillas':\n{}".format(e),
                title="Error script.json",
            )
            return None, []

        if not clave_planilla:
            forms.alert(
                u"No se encontró planilla asociada para el prefijo '{}'.".format(pref_cod),
                title="Planilla no definida",
            )
            return None, []

        try:
            host_doc = __revit__.ActiveUIDocument.Document
            schedules = (
                FilteredElementCollector(host_doc)
                .OfClass(ViewSchedule)
                .ToElements()
            )
            planilla_obj = next(
                (s for s in schedules if not s.IsTemplate and s.Name == clave_planilla),
                None,
            )
        except Exception as e:
            forms.alert(
                u"Error buscando la planilla '{}':\n{}".format(clave_planilla, e),
                title="Error planilla",
            )
            return None, []

        if planilla_obj is None:
            forms.alert(
                u"No se encontró la planilla '{}' en el modelo.".format(clave_planilla),
                title="Planilla no encontrada",
            )
            return None, []

        headers_order = []
        try:
            for fid in planilla_obj.Definition.GetFieldOrder():
                field = planilla_obj.Definition.GetField(fid)
                if field:
                    headers_order.append(field.GetName())
        except Exception as e:
            forms.alert(
                u"Error obteniendo encabezados de '{}':\n{}".format(clave_planilla, e),
                title="Error encabezados",
            )
            return None, []

        if not headers_order:
            forms.alert(
                u"No se obtuvieron encabezados para la planilla '{}'.\n"
                u"Verifica que tenga columnas de datos.".format(clave_planilla),
                title="Sin encabezados",
            )
            return None, []

        return clave_planilla, headers_order

    # ---------------- carga repo -> modelo ----------------------

    def load_parameters_from_repo_or_model(self):
        """Construye la lista de parámetros siguiendo el orden de la planilla."""
        self.params = []
        self.paramsListView.ItemsSource = self.params
        self._original_values = {}

        try:
            archivo_completo, element_id_str = self._get_archivo_y_id()
            repo_all = self._repo
            datos_repo_por_id = None
            for k, v in repo_all.items():
                try:
                    if (
                        isinstance(v, dict)
                        and v.get("Archivo") == archivo_completo
                        and v.get("ElementId") == element_id_str
                    ):
                        datos_repo_por_id = v
                        break
                except Exception:
                    continue

            clave_planilla, headers_order = self._get_headers_order()
            if not headers_order:
                self.statusLabel.Content = u"No se pudieron obtener encabezados."
                return

            parametros = {}
            for p in self.element.Parameters:
                try:
                    if not p.Definition or not p.HasValue:
                        continue
                    val = p.AsString() or p.AsValueString()
                    if not val or not val.strip():
                        continue
                    pname = p.Definition.Name
                    parametros[pname] = (p, val.strip())
                except Exception:
                    continue

            parametros_renombrados = {}
            for k, (p_obj, v) in parametros.items():
                k_norm = normalizar_clave(k)
                nuevo_nombre = self._reemplazos.get(k_norm, k_norm)
                parametros_renombrados[nuevo_nombre] = (p_obj, v)

            for head in headers_order:
                head_norm = normalizar_clave(head)
                p = None
                val_model = ""
                if head_norm in parametros_renombrados:
                    p, val_model = parametros_renombrados[head_norm]
                if datos_repo_por_id and head in datos_repo_por_id:
                    valor_final = datos_repo_por_id.get(head, "")
                else:
                    valor_final = val_model
                self.params.append(ParamItem(head, valor_final, p))
                self._original_values[head] = valor_final

            self.paramsListView.ItemsSource = None
            self.paramsListView.ItemsSource = self.params
            self.statusLabel.Content = u"Datos cargados (planilla '{}').".format(clave_planilla)

        except Exception as e:
            self.statusLabel.Content = u"Error cargando datos."
            forms.alert(
                u"Error en load_parameters_from_repo_or_model:\n{}".format(e),
                title="Error editor",
            )

    # ---------------- utilidades de cambios / guardado ---------

    def _has_changes(self):
        for p in self.params:
            orig = self._original_values.get(p.Name, "")
            if (p.Value or "") != (orig or ""):
                return True
        return False

    def save_params_to_repo(self):
        """Actualiza/crea registro para Archivo + ElementId si hay cambios."""
        if not self._has_changes():
            return False

        try:
            repo = load_repo()
            if not isinstance(repo, dict):
                repo = {}

            archivo_completo, element_id_str = self._get_archivo_y_id()
            nombre_archivo = (
                os.path.basename(archivo_completo) if archivo_completo else ""
            )

            existing_key = None
            for k, v in repo.items():
                if not isinstance(v, dict):
                    continue
                if (
                    v.get("Archivo") == archivo_completo
                    and v.get("ElementId") == element_id_str
                ):
                    existing_key = k
                    break

            if existing_key is None:
                existing_key = self._key
                i = 1
                while existing_key in repo:
                    existing_key = "{}_{}".format(self._key, i)
                    i += 1
                repo[existing_key] = {}

            entry = repo[existing_key]
            entry["Archivo"]        = archivo_completo
            entry["nombre_archivo"] = nombre_archivo
            entry["ElementId"]      = element_id_str

            for p in self.params:
                entry[p.Name] = p.Value

            save_repo(repo)
            self._repo = repo
            self._original_values = {p.Name: p.Value for p in self.params}
            return True

        except Exception as e:
            forms.alert(
                u"Error guardando datos en el repositorio:\n{}".format(e),
                title="Error guardado",
            )
            return False

    # ---------------- eventos de botones -----------------------

    def on_save(self, sender, e):
        try:
            if not self._has_changes():
                self.statusLabel.Content = u"Sin cambios. No se actualizó el repositorio."
                self._changes_saved = False
                self.Close()
                return
            self.statusLabel.Content = u"Guardando datos..."
            ok = self.save_params_to_repo()
            if ok:
                self._changes_saved = True
                self.statusLabel.Content = u"Datos guardados en repositorio."
            else:
                self._changes_saved = False
            self.Close()
        except Exception as e:
            forms.alert(u"Error en on_save:\n{}".format(e), title="Error guardar")

    def on_cancel(self, sender, e):
        try:
            self._changes_saved = False
            self.Close()
        except Exception as e:
            forms.alert(u"Error en on_cancel:\n{}".format(e), title="Error cancelar")

    def on_update_complete(self, success, message):
        try:
            self.statusLabel.Content = message
            if success:
                self._repo = load_repo()
                self.load_parameters_from_repo_or_model()
        except Exception as e:
            forms.alert(
                u"Error en callback de actualización:\n{}".format(e),
                title="Error actualización",
            )


# ------------------------------------------------------------
# main
# ------------------------------------------------------------

def main():
    uidoc = __revit__.ActiveUIDocument
    doc   = uidoc.Document

    try:
        ref = uidoc.Selection.PickObject(
            ObjectType.LinkedElement,
            "Selecciona un elemento en un archivo linkeado",
        )
    except Exception:
        forms.alert("Selección cancelada.")
        return

    try:
        link_instance = doc.GetElement(ref.ElementId)
        linked_doc    = link_instance.GetLinkDocument()
        if linked_doc is None:
            forms.alert("No se pudo obtener el documento linkeado.")
            return
        element = linked_doc.GetElement(ref.LinkedElementId)
        if element is None:
            forms.alert("No se pudo obtener el elemento.")
            return
    except Exception as e:
        forms.alert(
            u"Error accediendo al elemento linkeado:\n{}".format(e),
            title="Error selección",
        )
        return

    handler        = UpdateParamsHandler()
    external_event = ExternalEvent.Create(handler)
    editor         = ParamsEditor(linked_doc, element, external_event, handler)
    editor.ShowDialog()


if __name__ == "__main__":
    main()

#==================================================
#🚫 DELETE BELOW
#from Snippets._customprint import kit_button_clicked  # Import Reusable Function
#kit_button_clicked(btn_name=__title__)  # Display Default Print Message
