"""UNO tools for yWriter import/export in LibreOffice. 

Python wrappers for UNO widgets.

Part of the pywlo project.
Copyright (c) 2020, peter88213
For further information see https://github.com/peter88213/pywlo
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import uno
from msgbox import MsgBox
from uno_wrapper.uno_stub import *

# shortcut:
createUnoService = (
    XSCRIPTCONTEXT
    .getComponentContext()
    .getServiceManager()
    .createInstance
)


def FilePicker(path=None, mode=0):
    """
    Read file:  `mode in (0, 6, 7, 8, 9)`
    Write file: `mode in (1, 2, 3, 4, 5, 10)`
    see: (http://api.libreoffice.org/docs/idl/ref/
            namespacecom_1_1sun_1_1star_1_1ui_1_1
            dialogs_1_1TemplateDescription.html)

    See: https://stackoverflow.com/questions/30840736/libreoffice-how-to-create-a-file-dialog-via-python-macro
    """

    filepicker = createUnoService("com.sun.star.ui.dialogs.OfficeFilePicker")

    if path:
        filepicker.setDisplayDirectory(path)

    filepicker.initialize((mode,))
    filepicker.appendFilter("yWriter 7 Files (.yw7)", "*.yw7")
    filepicker.appendFilter("yWriter 6 Files (.yw6)", "*.yw6")

    if filepicker.execute():
        return filepicker.getFiles()[0]


def msgbox(message):
    myBox = MsgBox(XSCRIPTCONTEXT.getComponentContext())
    myBox.addButton('OK')
    myBox.renderFromBoxSize(200)
    myBox.numberOflines = 3
    myBox.show(message, 0, 'PyWriter')
