# -*- coding: utf-8 -*-
import sys
import os
import csv
import openpyxl


def leer_hoja_codigo(ruta_xlsm, data_dir):
    """
    Lee hoja 'CODIGO' desde fila 2, sin encabezado.
    Columna 1 = CodIntBIM, resto columnas en el mismo orden que Excel.
    Genera CODIGO.csv en data_dir y devuelve su ruta.
    """
    wb = openpyxl.load_workbook(ruta_xlsm, data_only=True)
    hoja = None
    for nombre in wb.sheetnames:
        if nombre.upper() == 'CODIGO':
            hoja = wb[nombre]
            break
    if hoja is None:
        raise RuntimeError("No se encontró hoja CODIGO.")

    filas = []
    for row in hoja.iter_rows(min_row=2, values_only=True):
        if not row or row[0] in (None, ''):
            continue
        filas.append(list(row))

    if not filas:
        raise RuntimeError("Hoja CODIGO sin filas útiles.")

    max_cols = max(len(r) for r in filas)
    headers = ['CodIntBIM'] + ['Param{}'.format(i) for i in range(2, max_cols + 1)]

    if not os.path.exists(data_dir):
        os.makedirs(data_dir)

    csv_path = os.path.join(data_dir, 'CODIGO.csv')
    with open(csv_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f, delimiter=';')
        writer.writerow(headers)
        for row in filas:
            row = list(row) + [''] * (max_cols - len(row))
            writer.writerow(row[:max_cols])

    print(csv_path)
    return csv_path


if __name__ == '__main__':
    if len(sys.argv) < 3:
        print("Uso: leer_xlsm_codigos.py <ruta_xlsm> <data_dir>", file=sys.stderr)
        sys.exit(1)

    ruta_xlsm = sys.argv[1]
    data_dir = sys.argv[2]

    try:
        leer_hoja_codigo(ruta_xlsm, data_dir)
    except Exception as e:
        print("Error: {}".format(e), file=sys.stderr)
        sys.exit(1)
