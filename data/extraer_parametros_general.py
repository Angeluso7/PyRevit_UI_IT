# -*- coding: utf-8 -*-
"""
exportar_parametros_general.py

- Lee script.json
- Toma TODOS los nombres de parámetros que aparezcan en 'reemplazos_encabezados'
- Elimina duplicados
- Genera un TXT 'parametros_general.txt' en la carpeta data
  con un parámetro por línea
"""

import os
import json
from collections import OrderedDict

# Misma ruta data que el resto de tus scripts
DATA_DIR = os.path.join(
    os.path.expanduser("~"),
    r"AppData\\Roaming\\MyPyRevitExtention\\PyRevitIT.extension\\data"
)

SCRIPT_JSON_PATH = os.path.join(DATA_DIR, "script.json")
TXT_SALIDA = os.path.join(DATA_DIR, "parametros_general.txt")


def cargar_script_json(ruta):
    if not os.path.exists(ruta):
        raise IOError("No se encontró script.json en: {}".format(ruta))
    with open(ruta, "r", encoding="utf-8") as f:
        return json.load(f)


def main():
    print("=" * 70)
    print("GENERAR 'parametros_general.txt' DESDE script.json")
    print("=" * 70)

    data_script = cargar_script_json(SCRIPT_JSON_PATH)

    # OJO: nombre correcto de la clave -> 'reemplazos_encabezados'
    reemplazos = data_script.get("reemplazos_encabezados", {})
    if not isinstance(reemplazos, dict):
        raise ValueError("El campo 'reemplazos_encabezados' no es un dict en script.json")

    # Tomamos las CLAVES del diccionario (nombres originales de parámetros)
    parametros = list(OrderedDict((k, None) for k in reemplazos.keys()).keys())
    print("  - Parámetros encontrados en 'reemplazos_encabezados':", len(parametros))

    # Escribir TXT en data
    with open(TXT_SALIDA, "w", encoding="utf-8") as f:
        for p in parametros:
            f.write(p + "\n")

    print("✓ Archivo generado:", TXT_SALIDA)
    print("=" * 70)


if __name__ == "__main__":
    main()
