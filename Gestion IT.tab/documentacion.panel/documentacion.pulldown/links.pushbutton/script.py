# -*- coding: utf-8 -*-
__title__   = "Links Infotécnica"
__doc__     = """Version = 1.0
Date    = 15.06.2024
________________________________________________________________
Description:

This is the placeholder for a .pushbutton
You can use it to start your pyRevit Add-In

________________________________________________________________
How-To:

1. [Hold ALT + CLICK] on the button to open its source folder.
You will be able to override this placeholder.

2. Automate Your Boring Work ;)

________________________________________________________________
TODO:
[FEATURE] - Describe Your ToDo Tasks Here
________________________________________________________________
Last Updates:
- [15.06.2024] v1.0 Change Description
- [10.06.2024] v0.5 Change Description
- [05.06.2024] v0.1 Change Description 
________________________________________________________________
Author: Erik Frits"""

# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝
#==================================================
import os
import clr
from pyrevit import forms

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

# ╦  ╦╔═╗╦═╗╦╔═╗╔╗ ╦  ╔═╗╔═╗
# ╚╗╔╝╠═╣╠╦╝║╠═╣╠╩╗║  ║╣ ╚═╗
#  ╚╝ ╩ ╩╩╚═╩╩ ╩╚═╝╩═╝╚═╝╚═╝
#==================================================
#app    = __revit__.Application
#uidoc  = __revit__.ActiveUIDocument
#doc    = __revit__.ActiveUIDocument.Document #type:Document

# ╔╦╗╔═╗╦╔╗╔
# ║║║╠═╣║║║║
# ╩ ╩╩ ╩╩╝╚╝
#==================================================

try:
    THIS_DIR = os.path.dirname(__file__)
except Exception:
    THIS_DIR = os.getcwd()

LINKS_HTML = os.path.join(THIS_DIR, "links_infotecnica.html")

class LinksWindow(forms.WPFWindow):
    def __init__(self):
        xaml_path = os.path.join(THIS_DIR, 'ui.xaml')
        forms.WPFWindow.__init__(self, xaml_path)

        if os.path.exists(LINKS_HTML):
            uri = "file:///" + LINKS_HTML.replace("\\", "/")
            self.LinksBrowser.Navigate(uri)
        else:
            forms.alert(
                u"No se encontró el archivo de links:\n{0}".format(LINKS_HTML),
                title="Links Infotécnica"
            )

def main():
    w = LinksWindow()
    w.ShowDialog()

if __name__ == '__main__':
    main()

#==================================================
#🚫 DELETE BELOW
from Snippets._customprint import kit_button_clicked    # Import Reusable Function from 'lib/Snippets/_customprint.py'
kit_button_clicked(btn_name=__title__)                  # Display Default Print Message
