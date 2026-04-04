# -*- coding: utf-8 -*-
"""
Script principal para PyRevit (IronPython)
Extrae ElementIds, CodIntBIM y parámetros de modelos linkeados
Genera archivo Excel con estructura definida por usuario

Autor: BIM Automation
Versión: 2.0
"""

import os
import sys
import json
import subprocess
from collections import defaultdict, OrderedDict

# PyRevit imports
from pyrevit import forms
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, LinkElementId, RevitLinkInstance

# Revit context
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# ============================================================
# CONSTANTES Y RUTAS
# ============================================================

DATA_DIR_EXT = os.path.join(
    os.path.expanduser("~"),
    r"AppData\Roaming\MyPyRevitExtention\PyRevitIT.extension\data"
)

SCRIPT_JSON_PATH = os.path.join(DATA_DIR_EXT, 'script.json')
FORMATEAR_SCRIPT = os.path.join(DATA_DIR_EXT, 'formatear_tablas_excel_v2.py')
PYTHON3_EXE = r"C:\Users\Zbook HP\AppData\Local\Programs\Python\Python313\python.exe"

# ============================================================
# FUNCIONES AUXILIARES
# ============================================================

def verificar_rutas():
    """Verifica que existan las rutas necesarias"""
    if not os.path.exists(DATA_DIR_EXT):
        forms.alert(
            "No se encontró la carpeta data de la extensión:\n{}".format(DATA_DIR_EXT),
            title="Error"
        )
        raise SystemExit
    
    if not os.path.exists(SCRIPT_JSON_PATH):
        forms.alert(
            "No se encontró script.json en:\n{}".format(SCRIPT_JSON_PATH),
            title="Error"
        )
        raise SystemExit
    
    if not os.path.exists(FORMATEAR_SCRIPT):
        forms.alert(
            "No se encontró formatear_tablas_excel_v2.py en:\n{}".format(FORMATEAR_SCRIPT),
            title="Error"
        )
        raise SystemExit


def cargar_config():
    """Carga config desde script.json"""
    try:
        with open(SCRIPT_JSON_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        forms.alert("Error al leer script.json:\n{}".format(e), title="Error")
        raise SystemExit


def seleccionar_carpeta_salida():
    """Diálogo para seleccionar carpeta donde guardar el Excel"""
    carpeta = forms.pick_folder()
    if not carpeta:
        return None
    return carpeta


def extraer_modelos_linkeados():
    """
    Extrae datos de todos los modelos linkeados.
    Retorna: dict con estructura:
    {
        'datos_elementos': [  # lista de dicts
            {
                'ElementId': int,
                'CodIntBIM': str,
                'Parámetro1': valor,
                'Parámetro2': valor,
                'Categoría': str,
                'Familia': str,
                'Tipo': str,
                'Nombre_RVT': str,
                ...
            },
            ...
        ],
        'codigos_planillas': {...},  # dict de mapping
    }
    """
    
    datos_elementos = []
    
    # Obtener todos los LinkElementId del documento
    collector = FilteredElementCollector(doc)
    link_instances = collector.OfCategory(BuiltInCategory.OST_RvtLinks).WhereElementIsNotElementType().ToElements()
    
    codigos_planillas = {}
    config = cargar_config()
    if config:
        codigos_planillas = config.get('codigos_planillas', {})
    
    for link_instance in link_instances:
        try:
            linked_doc = link_instance.GetLinkDocument()
            if linked_doc is None:
                continue
            
            nombre_rvt = os.path.basename(linked_doc.PathName) if linked_doc.PathName else "Desconocido"
            
            # Extraer todos los elementos del documento linkeado
            collector_linked = FilteredElementCollector(linked_doc)
            all_elements = collector_linked.WhereElementIsNotElementType().ToElements()
            
            for element in all_elements:
                try:
                    elem_id = element.Id.IntegerValue
                    
                    # Obtener CodIntBIM
                    codint_param = element.LookupParameter("CodIntBIM")
                    codint = ""
                    if codint_param and codint_param.AsString():
                        codint = codint_param.AsString().strip()
                    
                    # Construir dict del elemento
                    elem_dict = {
                        'ElementId': elem_id,
                        'CodIntBIM': codint,
                        'Nombre_RVT': nombre_rvt,
                        'Categoría': element.Category.Name if element.Category else "Sin Categoría",
                        'Familia': "",
                        'Tipo': ""
                    }
                    
                    # Extraer Familia y Tipo si es aplicable
                    if hasattr(element, 'Symbol') and element.Symbol:
                        elem_dict['Familia'] = element.Symbol.FamilyName if hasattr(element.Symbol, 'FamilyName') else ""
                        elem_dict['Tipo'] = element.Symbol.Name if hasattr(element.Symbol, 'Name') else ""
                    
                    # Extraer parámetros adicionales (todos los parámetros del elemento)
                    if hasattr(element, 'Parameters'):
                        for param in element.Parameters:
                            try:
                                param_name = param.Definition.Name
                                
                                # Saltarse parámetros no relevantes
                                if param_name in ['ElementId', 'CodIntBIM', 'Nombre_RVT', 'Categoría', 'Familia', 'Tipo']:
                                    continue
                                
                                param_value = ""
                                if param.StorageType.ToString() == "String":
                                    param_value = param.AsString() or ""
                                elif param.StorageType.ToString() == "Integer":
                                    param_value = str(param.AsInteger()) if param.AsInteger() is not None else ""
                                elif param.StorageType.ToString() == "Double":
                                    param_value = str(param.AsDouble()) if param.AsDouble() is not None else ""
                                elif param.StorageType.ToString() == "ElementId":
                                    param_value = str(param.AsElementId().IntegerValue) if param.AsElementId() else ""
                                
                                if param_value:  # Solo agregar si tiene valor
                                    elem_dict[param_name] = param_value
                            except:
                                pass
                    
                    datos_elementos.append(elem_dict)
                
                except Exception as e:
                    pass  # Continuar con el siguiente elemento
        
        except Exception as e:
            pass  # Continuar con el siguiente link
    
    return {
        'datos_elementos': datos_elementos,
        'codigos_planillas': codigos_planillas
    }


def procesar_datos(datos_raw, codigos_planillas):
    """
    Procesa datos extraídos para generar estructura de tablas por planilla.
    
    Retorna: {
        'tablas_por_prefijo': OrderedDict con tablas ordenadas por prefijo (CM01, CM02, ...),
        'excepciones': lista de elementos sin CodIntBIM o sin parámetro,
        'listado_tablas': dict con valores y claves de codigos_planillas
    }
    """
    
    elementos = datos_raw.get('datos_elementos', [])
    
    # Determinar qué parámetros son "tabla" (están en los primeros 4 chars del CodIntBIM)
    tablas_encontradas = {}  # prefijo -> set de headers reales
    elementos_por_tabla = defaultdict(list)
    excepciones = []
    
    for elem in elementos:
        codint = elem.get('CodIntBIM', '').strip()
        
        if not codint:
            # Elemento sin CodIntBIM
            excepciones.append({
                'elemento': elem,
                'situacion': 'No existe'  # No tiene parámetro CodIntBIM
            })
            continue
        
        if len(codint) < 4:
            excepciones.append({
                'elemento': elem,
                'situacion': 'No Asignado'  # Tiene parámetro pero sin datos válidos
            })
            continue
        
        # Obtener prefijo (primeros 4 caracteres)
        prefijo = codint[:4]
        
        # Buscar tabla correspondiente en codigos_planillas
        nombre_tabla = None
        for pref_key, tabla_name in codigos_planillas.items():
            if pref_key == prefijo:
                nombre_tabla = tabla_name
                break
        
        if not nombre_tabla:
            nombre_tabla = "Sin_mapeo_{}".format(prefijo)
        
        elementos_por_tabla[nombre_tabla].append(elem)
        
        # Registrar headers de esta tabla
        if nombre_tabla not in tablas_encontradas:
            tablas_encontradas[nombre_tabla] = set()
        
        for key in elem.keys():
            if key not in ['ElementId', 'CodIntBIM', 'Nombre_RVT', 'Categoría', 'Familia', 'Tipo']:
                tablas_encontradas[nombre_tabla].add(key)
    
    # Ordenar tablas por prefijo (CM01, CM02, CM03, etc.)
    tablas_ordenadas = OrderedDict()
    for pref in sorted(codigos_planillas.keys()):
        tabla_name = codigos_planillas[pref]
        if tabla_name in elementos_por_tabla:
            tablas_ordenadas[tabla_name] = elementos_por_tabla[tabla_name]
    
    return {
        'elementos_por_tabla': tablas_ordenadas,
        'excepciones': excepciones,
        'listado_tablas': {
            'valores': list(codigos_planillas.values()),
            'claves': list(codigos_planillas.keys())
        }
    }


def generar_json_para_formatear(datos_procesados, ruta_json):
    """
    Genera archivo JSON intermedio con datos procesados
    para que el script de formateo lo lea.
    """
    
    json_data = {
        'elementos_por_tabla': {},
        'excepciones': datos_procesados.get('excepciones', []),
        'listado_tablas': datos_procesados.get('listado_tablas', {})
    }
    
    # Convertir OrderedDict a dict serializable
    for tabla_name, elementos in datos_procesados.get('elementos_por_tabla', {}).items():
        json_data['elementos_por_tabla'][tabla_name] = [
            {k: str(v) if v is not None else "" for k, v in elem.items()}
            for elem in elementos
        ]
    
    try:
        with open(ruta_json, 'w', encoding='utf-8') as f:
            json.dump(json_data, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        forms.alert("Error al generar JSON intermedio:\n{}".format(e), title="Error")
        return False


def ejecutar_formateo(ruta_json, ruta_xlsx_salida):
    """
    Ejecuta el script de formateo con CPython
    """
    try:
        subprocess.check_call(
            [PYTHON3_EXE, FORMATEAR_SCRIPT, ruta_json, ruta_xlsx_salida],
            stderr=subprocess.STDOUT
        )
        return True
    except subprocess.CalledProcessError as e:
        forms.alert("Error en formateo de Excel:\n{}".format(e), title="Error")
        return False


# ============================================================
# FLUJO PRINCIPAL
# ============================================================

def main():
    """Flujo principal del script"""
    
    # Verificar rutas
    verificar_rutas()
    
    # Cargar configuración
    config = cargar_config()
    codigos_planillas = config.get('codigos_planillas', {})
    
    if not codigos_planillas:
        forms.alert(
            "No se encontró diccionario 'codigos_planillas' en script.json",
            title="Error"
        )
        return
    
    # Seleccionar carpeta de salida
    carpeta_salida = seleccionar_carpeta_salida()
    if not carpeta_salida:
        return
    
    # Extraer datos de modelos linkeados
    forms.alert("Extrayendo datos de modelos linkeados...", title="Procesando")
    datos_raw = extraer_modelos_linkeados()
    
    if not datos_raw.get('datos_elementos'):
        forms.alert("No se encontraron elementos en los modelos linkeados", title="Información")
        return
    
    # Procesar datos
    datos_procesados = procesar_datos(datos_raw, codigos_planillas)
    
    # Generar JSON intermedio
    ruta_json_temp = os.path.join(carpeta_salida, '_temp_datos.json')
    if not generar_json_para_formatear(datos_procesados, ruta_json_temp):
        return
    
    # Definir ruta de salida del Excel
    nombre_xlsx = "Exportacion_BIM_{}.xlsx".format(
        __revit__.ActiveUIDocument.Document.Title
    )
    ruta_xlsx = os.path.join(carpeta_salida, nombre_xlsx)
    
    # Ejecutar formateo
    if ejecutar_formateo(ruta_json_temp, ruta_xlsx):
        forms.alert(
            "Archivo generado exitosamente:\n{}".format(ruta_xlsx),
            title="Éxito"
        )
        
        # Limpiar JSON temporal
        try:
            os.remove(ruta_json_temp)
        except:
            pass
    else:
        forms.alert("Error al generar el archivo Excel", title="Error")


if __name__ == '__main__':
    main()
