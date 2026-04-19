# -*- coding: utf-8 -*-
__title__ = "Parametros por elemento"
__doc__ = """Version = 1.2
Date = 19.04.2026
________________________________________________________________
Description:

Editor de parametros por elemento sobre modelo linkeado,
usando planillas y un repositorio JSON.

Cambios v1.2:
- Clave repo normalizada: basename_sin_ext + "_" + ElementId
  (consistente entre sesiones independiente del path de red)
- Busqueda en repo por (archivo_basename, ElementId) en vez
  de path completo (resuelve el bug de seleccion/no-match)
- Delimitador de clave cambiado a "||" para evitar colisiones
  con "_" dentro de rutas
- ui.xaml con estilo dark
________________________________________________________________
Author: Angeluso
"""

# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝
import unicodedata
import os
import sys
import json

from Autodesk.Revit.DB import ViewSchedule, FilteredElementCollector
from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler
from Autodesk.Revit.UI.Selection import ObjectType

from pyrevit import forms


# ── Normalizacion ────────────────────────────────────────────────────────────
def normalizar_clave(texto):
    """Quita acentos/virgulillas y convierte n/N."""
    if texto is None:
        return None
    if not isinstance(texto, str):
        texto = str(texto)
    nfkd = unicodedata.normalize('NFKD', texto)
    sin_tildes = u"".join(c for c in nfkd if not unicodedata.combining(c))
    return sin_tildes.replace(u"\xf1", u"n").replace(u"\xd1", u"N")


# ── Rutas centralizadas ──────────────────────────────────────────────────────
try:
    _this_dir = os.path.dirname(os.path.abspath(__file__))
except Exception:
    _this_dir = os.getcwd()

_EXT_ROOT = os.path.abspath(os.path.join(_this_dir, '..', '..', '..', '..'))
_LIB_DIR  = os.path.join(_EXT_ROOT, 'lib')
if _LIB_DIR not in sys.path:
    sys.path.insert(0, _LIB_DIR)

try:
    from config.paths import (DATA_DIR, MASTER_DIR, TEMP_DIR, CACHE_DIR,
                               CONFIG_PROYECTO, REGISTRO_PROYECTOS,
                               SCRIPT_JSON_PATH_LIB, ensure_runtime_dirs)
    ensure_runtime_dirs()
except Exception:
    _DATA_DIR            = os.path.join(_EXT_ROOT, 'data')
    DATA_DIR             = _DATA_DIR
    MASTER_DIR           = os.path.join(_DATA_DIR, 'master')
    TEMP_DIR             = os.path.join(_DATA_DIR, 'temp')
    CACHE_DIR            = os.path.join(_DATA_DIR, 'cache')
    CONFIG_PROYECTO      = os.path.join(MASTER_DIR, 'config_proyecto_activo.json')
    REGISTRO_PROYECTOS   = os.path.join(MASTER_DIR, 'registro_proyectos.json')
    SCRIPT_JSON_PATH_LIB = os.path.join(MASTER_DIR, 'script.json')

CONFIG_PATH      = CONFIG_PROYECTO
SCRIPT_JSON_PATH = SCRIPT_JSON_PATH_LIB


# ── Repositorio ──────────────────────────────────────────────────────────────
def load_active_repo_path():
    if not os.path.exists(CONFIG_PATH):
        raise Exception(
            u"No se encontro config_proyecto_activo.json en:\n{}".format(CONFIG_PATH))
    with open(CONFIG_PATH, "r") as f:
        cfg = json.load(f)
    ruta = (cfg.get("ruta_repositorio_activo") or "").strip()
    if not ruta:
        raise Exception(u"config_proyecto_activo.json no tiene 'ruta_repositorio_activo'.")
    return ruta


def load_repo():
    try:
        bd_path = load_active_repo_path()
    except Exception as e:
        forms.alert(u"Error obteniendo ruta del repositorio:\n{}".format(e),
                    title="Error repositorio")
        return {}
    if not os.path.exists(bd_path):
        return {}
    try:
        with open(bd_path, "r") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception as e:
        forms.alert(u"Error leyendo repositorio:\n{}\nSe usara uno vacio.".format(e),
                    title="Error repositorio")
        return {}


def save_repo(repo_dict):
    try:
        bd_path = load_active_repo_path()
    except Exception as e:
        forms.alert(u"Error obteniendo ruta para guardar:\n{}".format(e),
                    title="Error repositorio")
        return
    try:
        with open(bd_path, "w") as f:
            json.dump(repo_dict, f, indent=4, ensure_ascii=False)
    except Exception as e:
        forms.alert(u"Error guardando repositorio:\n{}".format(e),
                    title="Error repositorio")


def load_script_json():
    if not os.path.exists(SCRIPT_JSON_PATH):
        forms.alert(u"No se encontro script.json en:\n{}".format(SCRIPT_JSON_PATH),
                    title="Error script.json")
        return None
    try:
        with open(SCRIPT_JSON_PATH, "r") as f:
            return json.load(f)
    except Exception as e:
        forms.alert(u"Error leyendo script.json:\n{}".format(e),
                    title="Error script.json")
        return None


# ── Helpers de clave repo ────────────────────────────────────────────────────
def _archivo_basename(doc):
    """
    Devuelve solo el nombre de archivo sin extension y sin path.
    Ejemplo: 'C:/work/Proyecto_v2.rvt' -> 'Proyecto_v2'
    Funciona aunque PathName este vacio (usa Title como fallback).
    """
    path = (doc.PathName or "").strip()
    if path:
        base = os.path.splitext(os.path.basename(path))[0]
        if base:
            return base
    # Fallback: Title del documento
    try:
        title = (doc.Title or "").strip()
        if title:
            return os.path.splitext(title)[0]
    except Exception:
        pass
    return "unknown_file"


def _make_repo_key(archivo_basename, element_id_str):
    """
    Clave unica de repositorio: 'basename||ElementId'
    Usar '||' evita colisiones con '_' dentro de nombres de archivo.
    """
    return u"{}||{}".format(archivo_basename, element_id_str)


def _buscar_en_repo(repo, archivo_basename, element_id_str):
    """
    Busca la entrada del repo por (archivo_basename, ElementId).
    Acepta tanto claves con '||' (nuevo formato) como claves legacy con '_'.
    Retorna (key, entry) o (None, None).
    """
    target_key = _make_repo_key(archivo_basename, element_id_str)

    # Busqueda exacta por clave nueva
    if target_key in repo and isinstance(repo[target_key], dict):
        return target_key, repo[target_key]

    # Busqueda por campos Archivo (basename) + ElementId en valores
    for k, v in repo.items():
        if not isinstance(v, dict):
            continue
        v_archivo = v.get("Archivo", "")
        # Comparar basename normalizado
        v_basename = os.path.splitext(os.path.basename(v_archivo))[0] if v_archivo else ""
        if v_basename == archivo_basename and str(v.get("ElementId", "")) == element_id_str:
            return k, v

    return None, None


# ── Clases de datos ──────────────────────────────────────────────────────────
class ParamItem(object):
    def __init__(self, name, value, revit_param):
        self.Name       = name
        self.Value      = value
        self.RevitParam = revit_param


class UpdateParamsHandler(IExternalEventHandler):
    def __init__(self):
        self._editor = None

    def Execute(self, uiapp):
        if self._editor:
            self._editor.on_update_complete(True, u"Actualizacion completada.")

    def GetName(self):
        return "UpdateParamsHandler"


# ── Editor WPF ───────────────────────────────────────────────────────────────
class ParamsEditor(forms.WPFWindow):
    def __init__(self, linked_doc, element, external_event, handler):
        xaml_path = os.path.join(os.path.dirname(__file__), "ui.xaml")
        forms.WPFWindow.__init__(self, xaml_path)

        self._doc            = linked_doc
        self.element         = element
        self._external_event = external_event
        self._handler        = handler
        if self._handler:
            self._handler._editor = self

        # Datos auxiliares
        self._repo         = load_repo()
        self._script_data  = load_script_json() or {}

        self._reemplazos_raw = self._script_data.get("reemplazos_de_nombres", {})
        self._reemplazos = {
            normalizar_clave(k): v
            for k, v in self._reemplazos_raw.items()
        }

        self.params           = []
        self._original_values = {}
        self._changes_saved   = False

        # Clave normalizada para repo
        self._archivo_base = _archivo_basename(self._doc)
        self._element_id   = str(self.element.Id.IntegerValue)
        self._repo_key     = _make_repo_key(self._archivo_base, self._element_id)

        self.paramsListView.ItemsSource = self.params

        self.btnSave.Click   += self.on_save
        self.btnCancel.Click += self.on_cancel

        # Titulo de ventana con info del elemento
        self.Title = u"Editor Parametros  —  {} | ID {}".format(
            self._archivo_base, self._element_id)

        try:
            self.load_parameters_from_repo_or_model()
        except Exception as e:
            forms.alert(u"Error inicializando editor:\n{}".format(e),
                        title="Error editor")

    # ── logica de planilla ────────────────────────────────────────────────────
    def _get_headers_order(self):
        try:
            p_cod  = self.element.LookupParameter("CodIntBIM")
            cod_val = p_cod.AsString() if p_cod and p_cod.HasValue else ""
        except Exception as e:
            forms.alert(u"Error leyendo CodIntBIM:\n{}".format(e), title="Error CodIntBIM")
            return None, []

        if not cod_val or len(cod_val) < 4:
            forms.alert(
                u"CodIntBIM vacio o con menos de 4 caracteres.\nValor: '{}'".format(cod_val),
                title="CodIntBIM invalido")
            return None, []

        pref_cod    = cod_val[:4]
        script_data = self._script_data or load_script_json() or {}
        codigos_planillas = script_data.get("codigos_planillas", {})

        if not codigos_planillas:
            forms.alert(u"script.json no tiene 'codigos_planillas'.",
                        title="Error script.json")
            return None, []

        clave_planilla = None
        for clave, vals in codigos_planillas.items():
            if isinstance(vals, list):
                if any(isinstance(v, str) and v.startswith(pref_cod) for v in vals):
                    clave_planilla = clave
                    break
            elif isinstance(vals, str) and vals.startswith(pref_cod):
                clave_planilla = clave
                break

        if not clave_planilla:
            forms.alert(
                u"No se encontro planilla para el prefijo '{}'.".format(pref_cod),
                title="Planilla no definida")
            return None, []

        try:
            host_doc = __revit__.ActiveUIDocument.Document
            schedules = (FilteredElementCollector(host_doc)
                         .OfClass(ViewSchedule).ToElements())
            planilla_obj = next(
                (s for s in schedules if not s.IsTemplate and s.Name == clave_planilla),
                None)
        except Exception as e:
            forms.alert(u"Error buscando planilla '{}':\n{}".format(clave_planilla, e),
                        title="Error planilla")
            return None, []

        if planilla_obj is None:
            forms.alert(u"No se encontro la planilla '{}' en el modelo.".format(clave_planilla),
                        title="Planilla no encontrada")
            return None, []

        headers_order = []
        try:
            for fid in planilla_obj.Definition.GetFieldOrder():
                field = planilla_obj.Definition.GetField(fid)
                if field:
                    headers_order.append(field.GetName())
        except Exception as e:
            forms.alert(u"Error obteniendo encabezados:\n{}".format(e),
                        title="Error encabezados")
            return None, []

        if not headers_order:
            forms.alert(u"Planilla '{}' sin columnas.".format(clave_planilla),
                        title="Sin encabezados")
            return None, []

        return clave_planilla, headers_order

    # ── carga parametros ──────────────────────────────────────────────────────
    def load_parameters_from_repo_or_model(self):
        self.params = []
        self.paramsListView.ItemsSource = self.params
        self._original_values = {}

        try:
            # Buscar entrada en repo por basename + ElementId (tolerante a cambios de path)
            _, datos_repo = _buscar_en_repo(
                self._repo, self._archivo_base, self._element_id)

            clave_planilla, headers_order = self._get_headers_order()
            if not headers_order:
                self.statusLabel.Content = u"No se pudieron obtener encabezados."
                return

            # Mapear parametros del elemento (modelo)
            parametros = {}
            for p in self.element.Parameters:
                try:
                    if not p.Definition or not p.HasValue:
                        continue
                    val = p.AsString() or p.AsValueString()
                    if not val or not val.strip():
                        continue
                    parametros[p.Definition.Name] = (p, val.strip())
                except Exception:
                    continue

            # Aplicar reemplazos de nombres
            parametros_renombrados = {}
            for k, (p_obj, v) in parametros.items():
                nuevo = self._reemplazos.get(normalizar_clave(k), normalizar_clave(k))
                parametros_renombrados[nuevo] = (p_obj, v)

            # Construir lista respetando el orden de la planilla
            for head in headers_order:
                head_norm  = normalizar_clave(head)
                p_obj      = None
                val_model  = ""
                if head_norm in parametros_renombrados:
                    p_obj, val_model = parametros_renombrados[head_norm]

                # Repo tiene prioridad sobre el modelo
                if datos_repo and head in datos_repo:
                    valor_final = datos_repo.get(head, "")
                else:
                    valor_final = val_model

                self.params.append(ParamItem(head, valor_final, p_obj))
                self._original_values[head] = valor_final

            self.paramsListView.ItemsSource = None
            self.paramsListView.ItemsSource = self.params
            self.statusLabel.Content = u"Datos cargados  —  planilla: '{}'".format(
                clave_planilla)

        except Exception as e:
            self.statusLabel.Content = u"Error cargando datos."
            forms.alert(u"Error en load_parameters_from_repo_or_model:\n{}".format(e),
                        title="Error editor")

    # ── cambios y guardado ────────────────────────────────────────────────────
    def _has_changes(self):
        for p in self.params:
            if (p.Value or "") != (self._original_values.get(p.Name, "") or ""):
                return True
        return False

    def save_params_to_repo(self):
        if not self._has_changes():
            return False
        try:
            repo = load_repo()
            if not isinstance(repo, dict):
                repo = {}

            # Buscar si ya existe una entrada (legacy o nueva clave)
            existing_key, _ = _buscar_en_repo(repo, self._archivo_base, self._element_id)
            if existing_key is None:
                existing_key = self._repo_key

            entry = dict(repo.get(existing_key, {})) if existing_key in repo else {}

            # Guardar siempre el path completo + basename para compatibilidad
            path_completo = (self._doc.PathName or "").strip()
            entry["Archivo"]        = path_completo if path_completo else self._archivo_base
            entry["nombre_archivo"] = self._archivo_base
            entry["ElementId"]      = self._element_id
            entry["CodIntBIM"]      = ""  # se sobreescribira si existe el parametro

            for p in self.params:
                entry[p.Name] = p.Value

            repo[existing_key] = entry
            save_repo(repo)

            self._repo            = repo
            self._original_values = {p.Name: p.Value for p in self.params}
            return True

        except Exception as e:
            forms.alert(u"Error guardando datos:\n{}".format(e), title="Error guardado")
            return False

    # ── eventos ───────────────────────────────────────────────────────────────
    def on_save(self, sender, e):
        try:
            if not self._has_changes():
                self.statusLabel.Content = u"Sin cambios."
                self.Close()
                return
            self.statusLabel.Content = u"Guardando..."
            ok = self.save_params_to_repo()
            if ok:
                self._changes_saved = True
                self.statusLabel.Content = u"Guardado correctamente."
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
            forms.alert(u"Error en callback:\n{}".format(e), title="Error actualizacion")


# ── main ─────────────────────────────────────────────────────────────────────
def main():
    uidoc = __revit__.ActiveUIDocument
    doc   = uidoc.Document

    try:
        ref = uidoc.Selection.PickObject(
            ObjectType.LinkedElement,
            u"Selecciona un elemento en un archivo vinculado",
        )
    except Exception:
        # Usuario cancelo la seleccion
        return

    try:
        link_instance = doc.GetElement(ref.ElementId)
        linked_doc    = link_instance.GetLinkDocument()
        if linked_doc is None:
            forms.alert(u"No se pudo obtener el documento vinculado.\n"
                        u"Verifica que el link este cargado.")
            return
        element = linked_doc.GetElement(ref.LinkedElementId)
        if element is None:
            forms.alert(u"No se pudo obtener el elemento seleccionado.")
            return
    except Exception as e:
        forms.alert(u"Error accediendo al elemento vinculado:\n{}".format(e),
                    title="Error seleccion")
        return

    handler        = UpdateParamsHandler()
    external_event = ExternalEvent.Create(handler)
    editor         = ParamsEditor(linked_doc, element, external_event, handler)
    editor.ShowDialog()


if __name__ == "__main__":
    main()
