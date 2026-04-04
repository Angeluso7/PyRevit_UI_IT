# -*- coding: utf-8 -*-
"""
generar_colores_json_desde_txt.py

Lee colores_parametros.txt (con pares "parametro": "#RRGGBB")
y genera colores_parametros.json con formato ARGB "FFRRGGBB"
para usar directamente en formatear_tablas_excel_v2.py.
"""

import os
import json

# Ajusta esta ruta a donde tengas el txt
TXT_PATH = r"D:\SAESA\1.-PROYECTOS\1.- SUBESTACION TRINIDAD_FELIPE OLIVOS\NUBE\colores_parametros.txt"

DATA_DIR = os.path.join(
    os.path.expanduser("~"),
    r"AppData\Roaming\MyPyRevitExtention\PyRevitIT.extension\data"
)
JSON_OUT_PATH = os.path.join(DATA_DIR, "colores_parametros.json")


def hex6_to_argb8(color_hex):
    """
    Convierte "#RRGGBB" o "RRGGBB" a "FFRRGGBB".
    Si ya viene como "AARRGGBB", lo devuelve tal cual.
    """
    if not color_hex:
        return "FFFFFFFF"
    c = color_hex.strip()
    if c.startswith("#"):
        c = c[1:]
    # Ya viene como AARRGGBB
    if len(c) == 8:
        return c.upper()
    # Viene como RRGGBB
    if len(c) == 6:
        return ("FF" + c).upper()
    # Cualquier otro caso, blanco
    return "FFFFFFFF"


def main():
    if not os.path.exists(TXT_PATH):
        print("ERROR: No se encontró el archivo txt:")
        print("  {}".format(TXT_PATH))
        return

    # Leer todo el txt
    with open(TXT_PATH, "r", encoding="utf-8") as f:
        contenido = f.read()

    # Envolver en llaves si no las tiene, para parsear como JSON
    texto = contenido.strip()
    if not texto.startswith("{"):
        texto = "{\n" + texto
    if not texto.endswith("}"):
        texto = texto.rstrip(", \n") + "\n}"

    try:
        data = json.loads(texto)
    except Exception as e:
        print("ERROR al interpretar el txt como JSON:")
        print(e)
        return

    # Convertir todos los colores a formato ARGB
    colores_argb = {}
    for param, color in data.items():
        if not isinstance(param, str):
            param = str(param)
        if not isinstance(color, str):
            color = str(color)
        colores_argb[param] = hex6_to_argb8(color)

    # Asegurar carpeta de salida
    if not os.path.exists(DATA_DIR):
        os.makedirs(DATA_DIR)

    with open(JSON_OUT_PATH, "w", encoding="utf-8") as f:
        json.dump(colores_argb, f, ensure_ascii=False, indent=2)

    print("colores_parametros.json generado en:")
    print("  {}".format(JSON_OUT_PATH))
    print("Total de parámetros:", len(colores_argb))


if __name__ == "__main__":
    main()
