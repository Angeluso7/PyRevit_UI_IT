# -*- coding: utf-8 -*-
__title__   = "Restablecer Vista"
__doc__     = """Version = 1.1
Date    = 20.04.2026
________________________________________________________________
Description:

Restablece los filtros de la vista activa:
- Activa y hace visibles TODOS los filtros aplicados a la vista.
- Limpia cualquier override grafico de cada filtro.
- Muestra un resumen de cuantos filtros fueron restablecidos.
________________________________________________________________
Author: Argenis Angel"""

# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝
#==================================================
import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitServices")
clr.AddReference("System.Windows.Forms")

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    ParameterFilterElement,
    Transaction,
    OverrideGraphicSettings,
)
from System.Windows.Forms import MessageBox, MessageBoxButtons, DialogResult

# ╦  ╦╔═╗╦═╗╦╔═╗╔╗ ╦  ╔═╗╔═╗
# ╚╗╔╝╠═╣╠╦╝║╠═╣╠╩╗║  ║╣ ╚═╗
#  ╚╝ ╩ ╩╩╚═╩╩ ╩╚═╝╩═╝╚═╝╚═╝
#==================================================
doc   = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
vista = doc.ActiveView

# ╔╦╗╔═╗╦╔╗╔
# ║║║╠═╣║║║║
# ╩ ╩╩ ╩╩╝╚╝
#==================================================

def main():
    filtros_en_vista = list(vista.GetFilters())

    if not filtros_en_vista:
        MessageBox.Show(
            "La vista activa no tiene filtros aplicados.\n"
            "No hay nada que restablecer.",
            "Restablecer Vista — Aviso")
        return

    # ── Mapear ids → nombres para el resumen ────────────────────────────────
    filtros_dict = {}
    for f in FilteredElementCollector(doc).OfClass(ParameterFilterElement):
        filtros_dict[f.Id] = f.Name

    # ── Confirmar ────────────────────────────────────────────────────────────
    nombres = "\n".join(
        "  • " + filtros_dict.get(fid, str(fid))
        for fid in filtros_en_vista
    )
    respuesta = MessageBox.Show(
        "Se restablecerán {} filtro(s) en la vista activa:\n\n{}\n\n"
        "¿Continuar?".format(len(filtros_en_vista), nombres),
        "Restablecer Vista",
        MessageBoxButtons.YesNo)

    if respuesta != DialogResult.Yes:
        return

    # ── Transacción ──────────────────────────────────────────────────────────
    ogs_limpio = OverrideGraphicSettings()

    with Transaction(doc, "Restablecer filtros de la vista") as t:
        t.Start()
        try:
            for filtro_id in filtros_en_vista:
                vista.SetIsFilterEnabled(filtro_id, True)
                vista.SetFilterVisibility(filtro_id, True)
                vista.SetFilterOverrides(filtro_id, ogs_limpio)
            t.Commit()
        except Exception as ex:
            t.RollBack()
            MessageBox.Show(
                "Error al restablecer los filtros:\n{}".format(ex),
                "Error")
            return

    MessageBox.Show(
        "{} filtro(s) restablecidos correctamente en la vista:\n\n{}".format(
            len(filtros_en_vista), nombres),
        "Restablecer Vista — Listo")


# ──────────────────────────────────────────────────────────────────────────────
main()
