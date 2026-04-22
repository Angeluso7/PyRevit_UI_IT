# -*- coding: utf-8 -*-
__title__ = "Parametros por elemento"
__doc__ = """Version = 1.7
Date    = 22.04.2026
________________________________________________________________
Description:

Editor de parametros por elemento sobre modelo linkeado,
usando planillas y un repositorio JSON.

Cambios v1.7:
- Corregido error WPF: no se cierra ParamsEditor antes de ShowDialog().
- Se usa flag _abort_show para evitar mostrar la ventana principal
  cuando solo se insertó CodIntBIM en la BD.
- Se mantiene lectura de CodIntBIM desde BD primero y luego desde Revit.
- Inserción de CodIntBIM sigue creando entrada completa en BD.

________________________________________________________________
Author: Angeluso
"""

import unicodedata
import os
import sys
import json

import clr
clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")

from System.Windows import (
    Window, MessageBox, MessageBoxButton, MessageBoxResult,
    ResizeMode, WindowStartupLocation,
    Thickness, HorizontalAlignment
)
from System.Windows.Controls import (
    StackPanel, Label, TextBox, Button, DockPanel, Dock
)
from System.Windows.Media import Brushes, SolidColorBrush, Color

from Autodesk.Revit.DB import ViewSchedule, FilteredElementCollector
from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler
from Autodesk.Revit.UI.Selection import ObjectType

from pyrevit import forms


# ── Normalizacion ─────────────────────────────────────────────
def normalizar_clave(texto):
    if texto is None:
        return None
    if not isinstance(texto, str):
        texto = str(texto)
    nfkd = unicodedata.normalize('NFKD', texto)
    sin_tildes = u"".join(c for c in nfkd if not unicodedata.combining(c))
    return sin_tildes.replace(u"\xf1", u"n").replace(u"\xd1", u"N")


# ── Rutas centralizadas ───────────────────────────────────────
try:
    _this_dir = os.path.dirname(os.path.abspath(__file__))
except Exception:
    _this_dir = os.getcwd()

_EXT_ROOT = os.path.abspath(os.path.join(_this_dir, '..', '..', '..', '..'))
_LIB_DIR  = os.path.join(_EXT_ROOT, 'lib')
if _LIB_DIR not in sys.path:
    sys.path.insert(0, _LIB_DIR)

try:
    from config.paths import (
        DATA_DIR, MASTER_DIR, TEMP_DIR, CACHE_DIR,
        CONFIG_PROYECTO, REGISTRO_PROYECTOS,
        SCRIPT_JSON_PATH_LIB, ensure_runtime_dirs,
        get_ruta_repositorio
    )
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
    def get_ruta_repositorio(nup):
        nombre = u'repositorio_datos_{}.json'.format(nup)
        return os.path.join(os.path.join(_EXT_ROOT, 'data'), 'proyectos', nombre)

CONFIG_PATH      = CONFIG_PROYECTO
SCRIPT_JSON_PATH = SCRIPT_JSON_PATH_LIB


# ── Repositorio ───────────────────────────────────────────────
def load_active_repo_path():
    if not os.path.exists(CONFIG_PATH):
        raise Exception(
            u"No se encontro config_proyecto_activo.json en:\n{}".format(CONFIG_PATH))
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        cfg = json.load(f)
    nup = (cfg.get("nup_activo") or "").strip()
    if not nup:
        raise Exception(u"config_proyecto_activo.json no tiene 'nup_activo'.")
    ruta = get_ruta_repositorio(nup)
    if not os.path.exists(ruta):
        raise Exception(
            u"No se encontro repositorio del proyecto activo:\n{}".format(ruta))
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
        with open(bd_path, "r", encoding="utf-8") as f:
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
        with open(bd_path, "w", encoding="utf-8") as f:
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
        with open(SCRIPT_JSON_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        forms.alert(u"Error leyendo script.json:\n{}".format(e),
                    title="Error script.json")
        return None


# ── Helpers repo ──────────────────────────────────────────────
def _archivo_basename(doc):
    path = (doc.PathName or "").strip()
    if path:
        base = os.path.splitext(os.path.basename(path))[0]
        if base:
            return base
    try:
        title = (doc.Title or "").strip()
        if title:
            return os.path.splitext(title)[0]
    except Exception:
        pass
    return "unknown_file"


def _make_repo_key(archivo_basename, element_id_str):
    return u"{}||{}".format(archivo_basename, element_id_str)


def _buscar_en_repo(repo, archivo_basename, element_id_str):
    target_key = _make_repo_key(archivo_basename, element_id_str)
    if target_key in repo and isinstance(repo[target_key], dict):
        return target_key, repo[target_key]
    for k, v in repo.items():
        if not isinstance(v, dict):
            continue
        v_archivo  = v.get("Archivo", "")
        v_basename = os.path.splitext(os.path.basename(v_archivo))[0] if v_archivo else ""
        if v_basename == archivo_basename and str(v.get("ElementId", "")) == element_id_str:
            return k, v
    return None, None


def _leer_codintbim(element, datos_repo):
    if datos_repo:
        cod_bd = (datos_repo.get("CodIntBIM") or "").strip()
        if cod_bd:
            return cod_bd
    try:
        p = element.LookupParameter("CodIntBIM")
        if p and p.HasValue:
            val = (p.AsString() or "").strip()
            if val:
                return val
    except Exception:
        pass
    return ""


def _obtener_headers_planilla(pref_cod, script_data):
    codigos_planillas = script_data.get("codigos_planillas", {})
    if not codigos_planillas:
        return None, []

    nombre_planilla = None
    for nombre_plan, lista_codigos in codigos_planillas.items():
        if isinstance(lista_codigos, list):
            for cod in lista_codigos:
                cod_str = str(cod).strip()
                if cod_str[:4].upper() == pref_cod.upper():
                    nombre_planilla = nombre_plan
                    break
        elif isinstance(lista_codigos, str):
            if lista_codigos.strip()[:4].upper() == pref_cod.upper():
                nombre_planilla = nombre_plan
        if nombre_planilla:
            break

    if not nombre_planilla:
        return None, []

    try:
        host_doc  = __revit__.ActiveUIDocument.Document
        schedules = FilteredElementCollector(host_doc).OfClass(ViewSchedule).ToElements()
        planilla  = next(
            (s for s in schedules if not s.IsTemplate and s.Name == nombre_planilla),
            None
        )
    except Exception:
        return nombre_planilla, []

    if planilla is None:
        return nombre_planilla, []

    headers = []
    try:
        for fid in planilla.Definition.GetFieldOrder():
            field = planilla.Definition.GetField(fid)
            if field:
                headers.append(field.GetName())
    except Exception:
        pass

    return nombre_planilla, headers


def _crear_entrada_bd(linked_doc, element, nuevo_cod,
                      archivo_base, element_id_str,
                      headers_planilla, repo):
    existing_key, existing_entry = _buscar_en_repo(repo, archivo_base, element_id_str)
    key_usar = existing_key if existing_key else _make_repo_key(archivo_base, element_id_str)

    entry = dict(existing_entry) if existing_entry else {}

    path_completo           = (linked_doc.PathName or "").strip()
    entry["Archivo"]        = path_completo if path_completo else archivo_base
    entry["nombre_archivo"] = archivo_base
    entry["ElementId"]      = element_id_str
    entry["CodIntBIM"]      = nuevo_cod

    for h in headers_planilla:
        if h not in entry:
            entry[h] = ""

    repo[key_usar] = entry
    save_repo(repo)
    return key_usar, entry


# ── Ventana inserción ─────────────────────────────────────────
class CodIntBIMEditorWindow(Window):
    def __init__(self, linked_doc, element, archivo_base,
                 element_id_str, script_data, valor_actual=""):
        Window.__init__(self)
        self.guardado = False

        self._linked_doc   = linked_doc
        self._element      = element
        self._archivo_base = archivo_base
        self._element_id   = element_id_str
        self._script_data  = script_data

        self.Title                 = u"Insertar CodIntBIM"
        self.Width                 = 440
        self.Height                = 190
        self.ResizeMode            = ResizeMode.CanResize
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.Background            = SolidColorBrush(Color.FromRgb(45, 45, 48))

        panel        = StackPanel()
        panel.Margin = Thickness(16, 14, 16, 14)

        lbl            = Label()
        lbl.Content    = (u"Elemento sin CodIntBIM registrado.\n"
                          u"Ingrese el codigo para registrarlo en la BD:")
        lbl.Foreground = Brushes.LightGray
        lbl.Padding    = Thickness(0, 0, 0, 6)
        panel.Children.Add(lbl)

        self._txt            = TextBox()
        self._txt.Text       = valor_actual or ""
        self._txt.Height     = 28
        self._txt.Padding    = Thickness(6, 4, 6, 4)
        self._txt.FontSize   = 13
        self._txt.Background = SolidColorBrush(Color.FromRgb(30, 30, 30))
        self._txt.Foreground = Brushes.White
        self._txt.CaretBrush = Brushes.White
        self._txt.Margin     = Thickness(0, 0, 0, 14)
        panel.Children.Add(self._txt)

        self._lbl_estado            = Label()
        self._lbl_estado.Content    = ""
        self._lbl_estado.Foreground = Brushes.LightYellow
        self._lbl_estado.Padding    = Thickness(0, 0, 0, 6)
        self._lbl_estado.Height     = 22
        panel.Children.Add(self._lbl_estado)

        btn_panel               = DockPanel()
        btn_panel.LastChildFill = False

        btn_cancel                     = Button()
        btn_cancel.Content             = u"Cancelar"
        btn_cancel.Width               = 90
        btn_cancel.Height              = 28
        btn_cancel.Margin              = Thickness(0, 0, 8, 0)
        btn_cancel.HorizontalAlignment = HorizontalAlignment.Right
        btn_cancel.Background          = SolidColorBrush(Color.FromRgb(62, 62, 66))
        btn_cancel.Foreground          = Brushes.White
        btn_cancel.Click              += self._on_cancelar
        DockPanel.SetDock(btn_cancel, Dock.Right)

        btn_ok                         = Button()
        btn_ok.Content                 = u"Guardar en BD"
        btn_ok.Width                   = 110
        btn_ok.Height                  = 28
        btn_ok.HorizontalAlignment     = HorizontalAlignment.Right
        btn_ok.Background              = SolidColorBrush(Color.FromRgb(0, 122, 204))
        btn_ok.Foreground              = Brushes.White
        btn_ok.Click                  += self._on_guardar
        DockPanel.SetDock(btn_ok, Dock.Right)

        btn_panel.Children.Add(btn_cancel)
        btn_panel.Children.Add(btn_ok)
        panel.Children.Add(btn_panel)

        self.Content  = panel
        self.Loaded  += lambda s, e: self._txt.Focus()

    def _on_guardar(self, sender, e):
        val = (self._txt.Text or "").strip()
        if len(val) < 4:
            self._lbl_estado.Content = u"⚠ El codigo debe tener al menos 4 caracteres."
            return

        pref_cod = val[:4].upper()
        nombre_planilla, headers = _obtener_headers_planilla(pref_cod, self._script_data)

        if not headers:
            self._lbl_estado.Content = (
                u"⚠ No se encontro planilla para '{}'. Verifique script.json.".format(pref_cod))
            return

        try:
            repo = load_repo()
            _crear_entrada_bd(
                self._linked_doc, self._element, val,
                self._archivo_base, self._element_id,
                headers, repo
            )
            self.guardado     = True
            self.DialogResult = True
            self.Close()
        except Exception as ex:
            self._lbl_estado.Content = u"Error al guardar: {}".format(ex)

    def _on_cancelar(self, sender, e):
        self.guardado     = False
        self.DialogResult = False
        self.Close()


# ── Clases de apoyo ───────────────────────────────────────────
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


# ── Editor principal ──────────────────────────────────────────
class ParamsEditor(forms.WPFWindow):
    def __init__(self, linked_doc, element, external_event, handler):
        xaml_path = os.path.join(os.path.dirname(__file__), "ui.xaml")
        forms.WPFWindow.__init__(self, xaml_path)

        self._doc            = linked_doc
        self.element         = element
        self._external_event = external_event
        self._handler        = handler
        self._abort_show     = False

        if self._handler:
            self._handler._editor = self

        self._repo        = load_repo()
        self._script_data = load_script_json() or {}

        self._reemplazos_raw = self._script_data.get("reemplazos_de_nombres", {})
        self._reemplazos     = {
            normalizar_clave(k): v
            for k, v in self._reemplazos_raw.items()
        }

        self.params           = []
        self._original_values = {}
        self._changes_saved   = False

        self._archivo_base = _archivo_basename(self._doc)
        self._element_id   = str(self.element.Id.IntegerValue)
        self._repo_key     = _make_repo_key(self._archivo_base, self._element_id)

        self.paramsListView.ItemsSource = self.params

        self.btnSave.Click   += self.on_save
        self.btnCancel.Click += self.on_cancel

        self.Title = u"Editor Parametros  —  {} | ID {}".format(
            self._archivo_base, self._element_id)

        try:
            self.load_parameters_from_repo_or_model()
        except Exception as e:
            forms.alert(u"Error inicializando editor:\n{}".format(e),
                        title="Error editor")

    def _get_headers_order(self, datos_repo):
        cod_val = _leer_codintbim(self.element, datos_repo)

        if not cod_val or len(cod_val) < 4:
            win = CodIntBIMEditorWindow(
                linked_doc     = self._doc,
                element        = self.element,
                archivo_base   = self._archivo_base,
                element_id_str = self._element_id,
                script_data    = self._script_data,
                valor_actual   = cod_val
            )
            win.ShowDialog()

            self._abort_show = True
            return None, None

        pref_cod = cod_val[:4].upper()
        nombre_planilla, headers_order = _obtener_headers_planilla(
            pref_cod, self._script_data)

        if not headers_order:
            forms.alert(
                u"No se encontro planilla para el prefijo '{}'.\n\n"
                u"Verifique que en script.json → 'codigos_planillas'\n"
                u"exista una lista con codigos que empiecen por '{}'.".format(
                    pref_cod, pref_cod),
                title="Planilla no definida")
            return None, []

        return nombre_planilla, headers_order

    def load_parameters_from_repo_or_model(self):
        self.params = []
        self.paramsListView.ItemsSource = self.params
        self._original_values           = {}

        try:
            _, datos_repo = _buscar_en_repo(
                self._repo, self._archivo_base, self._element_id)

            clave_planilla, headers_order = self._get_headers_order(datos_repo)

            if self._abort_show:
                return

            if not headers_order:
                return

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

            parametros_renombrados = {}
            for k, (p_obj, v) in parametros.items():
                nuevo = self._reemplazos.get(
                    normalizar_clave(k), normalizar_clave(k))
                parametros_renombrados[nuevo] = (p_obj, v)

            for head in headers_order:
                head_norm = normalizar_clave(head)
                p_obj     = None
                val_model = ""
                if head_norm in parametros_renombrados:
                    p_obj, val_model = parametros_renombrados[head_norm]

                if datos_repo and head in datos_repo:
                    valor_final = datos_repo.get(head, "") or ""
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
            forms.alert(
                u"Error en load_parameters_from_repo_or_model:\n{}".format(e),
                title="Error editor")

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

            existing_key, _ = _buscar_en_repo(
                repo, self._archivo_base, self._element_id)
            if existing_key is None:
                existing_key = self._repo_key

            entry = dict(repo.get(existing_key, {})) if existing_key in repo else {}

            path_completo           = (self._doc.PathName or "").strip()
            entry["Archivo"]        = path_completo if path_completo else self._archivo_base
            entry["nombre_archivo"] = self._archivo_base
            entry["ElementId"]      = self._element_id
            entry["CodIntBIM"]      = entry.get("CodIntBIM", "")

            for p in self.params:
                entry[p.Name] = p.Value

            repo[existing_key] = entry
            save_repo(repo)

            self._repo            = repo
            self._original_values = {p.Name: p.Value for p in self.params}
            return True

        except Exception as e:
            forms.alert(u"Error guardando datos:\n{}".format(e),
                        title="Error guardado")
            return False

    def on_save(self, sender, e):
        try:
            if not self._has_changes():
                self.statusLabel.Content = u"Sin cambios."
                self.Close()
                return
            self.statusLabel.Content = u"Guardando..."
            ok = self.save_params_to_repo()
            if ok:
                self._changes_saved      = True
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
            forms.alert(u"Error en callback:\n{}".format(e),
                        title="Error actualizacion")


# ── main ──────────────────────────────────────────────────────
def main():
    uidoc = __revit__.ActiveUIDocument
    doc   = uidoc.Document

    try:
        ref = uidoc.Selection.PickObject(
            ObjectType.LinkedElement,
            u"Selecciona un elemento en un archivo vinculado",
        )
    except Exception:
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

    if getattr(editor, "_abort_show", False):
        return

    editor.ShowDialog()


if __name__ == "__main__":
    main()