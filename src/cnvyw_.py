"""Convert yWriter project to odt or csv and vice versa. 

Version @release

Copyright (c) 2020 Peter Triesberger
For further information see https://github.com/peter88213/yw-cnv
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import os

from configparser import ConfigParser
from urllib.parse import unquote
from urllib.parse import quote

from pywriter.converter.universal_file_factory import UniversalFileFactory
from pywriter.odt.odt_proof import OdtProof
from pywriter.odt.odt_manuscript import OdtManuscript
from pywriter.odt.odt_scenedesc import OdtSceneDesc
from pywriter.odt.odt_chapterdesc import OdtChapterDesc
from pywriter.odt.odt_partdesc import OdtPartDesc
from pywriter.csv.csv_scenelist import CsvSceneList
from pywriter.csv.csv_plotlist import CsvPlotList
from pywriter.csv.csv_charlist import CsvCharList
from pywriter.csv.csv_loclist import CsvLocList
from pywriter.csv.csv_itemlist import CsvItemList
from pywriter.odt.odt_characters import OdtCharacters
from pywriter.odt.odt_items import OdtItems
from pywriter.odt.odt_locations import OdtLocations

import uno
from com.sun.star.awt.MessageBoxType import MESSAGEBOX, INFOBOX, WARNINGBOX, ERRORBOX, QUERYBOX

from libreoffice.uno_tools import *
from libreoffice.yw_cnv_uno import YwCnvUno

INI_FILE = 'openyw.ini'


class Converter(YwCnvUno):
    """Converter for yWriter project files.
    Variant with UNO UI.
    """

    def __init__(self, silentMode=False):
        YwCnvUno.__init__(self, silentMode)
        self.fileFactory = UniversalFileFactory()


def open_yw7(suffix, newExt):

    # Set last opened yWriter project as default (if existing).

    scriptLocation = os.path.dirname(__file__)
    inifile = unquote(
        (scriptLocation + '/' + INI_FILE).replace('file:///', ''))
    defaultFile = None
    config = ConfigParser()

    try:
        config.read(inifile)
        ywLastOpen = config.get('FILES', 'yw_last_open')

        if os.path.isfile(ywLastOpen):
            defaultFile = quote('file:///' + ywLastOpen, '/:')

    except:
        pass

    # Ask for yWriter 6 or 7 project to open:

    ywFile = FilePicker(path=defaultFile)

    if ywFile is None:
        return

    sourcePath = unquote(ywFile.replace('file:///', ''))
    ywExt = os.path.splitext(sourcePath)[1]

    if not ywExt in ['.yw6', '.yw7']:
        msgbox('Please choose a yWriter 6/7 project.',
               'Import from yWriter', type_msg=ERRORBOX)
        return

    # Store selected yWriter project as "last opened".

    newFile = ywFile.replace(ywExt, suffix + newExt)
    dirName, filename = os.path.split(newFile)
    lockFile = unquote((dirName + '/.~lock.' + filename +
                        '#').replace('file:///', ''))

    if not config.has_section('FILES'):
        config.add_section('FILES')

    config.set('FILES', 'yw_last_open', unquote(
        ywFile.replace('file:///', '')))

    with open(inifile, 'w') as f:
        config.write(f)

    # Check if import file is already open in LibreOffice:

    if os.path.isfile(lockFile):
        msgbox('Please close "' + filename + '" first.',
               'Import from yWriter', type_msg=ERRORBOX)
        return

    # Open yWriter project and convert data.

    workdir = os.path.dirname(sourcePath)
    os.chdir(workdir)
    converter = Converter()
    converter.run(sourcePath, suffix)

    if converter.success:
        desktop = XSCRIPTCONTEXT.getDesktop()
        doc = desktop.loadComponentFromURL(newFile, "_blank", 0, ())


def import_yw(*args):
    '''Import scenes from yWriter 6/7 to a Writer document
    without chapter and scene markers. 
    '''
    open_yw7('', '.odt')


def proof_yw(*args):
    '''Import scenes from yWriter 6/7 to a Writer document
    with visible chapter and scene markers. 
    '''
    open_yw7(OdtProof.SUFFIX, OdtProof.EXTENSION)


def get_manuscript(*args):
    '''Import scenes from yWriter 6/7 to a Writer document
    with invisible chapter and scene markers. 
    '''
    open_yw7(OdtManuscript.SUFFIX, OdtManuscript.EXTENSION)


def get_partdesc(*args):
    '''Import pard descriptions from yWriter 6/7 to a Writer document
    with invisible chapter and scene markers. 
    '''
    open_yw7(OdtPartDesc.SUFFIX, OdtPartDesc.EXTENSION)


def get_chapterdesc(*args):
    '''Import chapter descriptions from yWriter 6/7 to a Writer document
    with invisible chapter and scene markers. 
    '''
    open_yw7(OdtChapterDesc.SUFFIX, OdtChapterDesc.EXTENSION)


def get_scenedesc(*args):
    '''Import scene descriptions from yWriter 6/7 to a Writer document
    with invisible chapter and scene markers. 
    '''
    open_yw7(OdtSceneDesc.SUFFIX, OdtSceneDesc.EXTENSION)


def get_chardesc(*args):
    '''Import character descriptions from yWriter 6/7 to a Writer document.
    '''
    open_yw7(OdtCharacters.SUFFIX, OdtCharacters.EXTENSION)


def get_locdesc(*args):
    '''Import location descriptions from yWriter 6/7 to a Writer document.
    '''
    open_yw7(OdtLocations.SUFFIX, OdtLocations.EXTENSION)


def get_itemdesc(*args):
    '''Import item descriptions from yWriter 6/7 to a Writer document.
    '''
    open_yw7(OdtItems.SUFFIX, OdtItems.EXTENSION)


def get_scenelist(*args):
    '''Import a scene list from yWriter 6/7 to a Calc document.
    '''
    open_yw7(CsvSceneList.SUFFIX, CsvSceneList.EXTENSION)


def get_plotlist(*args):
    '''Import a plot list from yWriter 6/7 to a Calc document.
    '''
    open_yw7(CsvPlotList.SUFFIX, CsvPlotList.EXTENSION)


def get_charlist(*args):
    '''Import a character list from yWriter 6/7 to a Calc document.
    '''
    open_yw7(CsvCharList.SUFFIX, CsvCharList.EXTENSION)


def get_loclist(*args):
    '''Import a location list from yWriter 6/7 to a Calc document.
    '''
    open_yw7(CsvLocList.SUFFIX, CsvLocList.EXTENSION)


def get_itemlist(*args):
    '''Import an item list from yWriter 6/7 to a Calc document.
    '''
    open_yw7(CsvItemList.SUFFIX, CsvItemList.EXTENSION)


def export_yw(*args):
    '''Export the document to a yWriter 6/7 project.
    '''

    # Get document's filename

    document = XSCRIPTCONTEXT.getDocument().CurrentController.Frame
    # document   = ThisComponent.CurrentController.Frame

    ctx = XSCRIPTCONTEXT.getComponentContext()
    smgr = ctx.getServiceManager()
    dispatcher = smgr.createInstanceWithContext(
        "com.sun.star.frame.DispatchHelper", ctx)
    # dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")

    documentPath = XSCRIPTCONTEXT.getDocument().getURL()
    # documentPath = ThisComponent.getURL()

    from com.sun.star.beans import PropertyValue
    args1 = []
    args1.append(PropertyValue())
    args1.append(PropertyValue())
    # dim args1(1) as new com.sun.star.beans.PropertyValue

    if documentPath.endswith('.odt') or documentPath.endswith('.html'):
        odtPath = documentPath.replace('.html', '.odt')
        htmlPath = documentPath.replace('.odt', '.html')

        # Save document in HTML format

        args1[0].Name = 'URL'
        # args1(0).Name = "URL"
        args1[0].Value = htmlPath
        # args1(0).Value = htmlPath
        args1[1].Name = 'FilterName'
        # args1(1).Name = "FilterName"
        args1[1].Value = 'HTML (StarWriter)'
        # args1(1).Value = "HTML (StarWriter)"
        dispatcher.executeDispatch(document, ".uno:SaveAs", "", 0, args1)
        # dispatcher.executeDispatch(document, ".uno:SaveAs", "", 0, args1())

        # Save document in OpenDocument format

        args1[0].Value = odtPath
        # args1(0).Value = odtPath
        args1[1].Value = 'writer8'
        # args1(1).Value = "writer8"
        dispatcher.executeDispatch(document, ".uno:SaveAs", "", 0, args1)
        # dispatcher.executeDispatch(document, ".uno:SaveAs", "", 0, args1())

        targetPath = unquote(htmlPath.replace('file:///', ''))

    elif documentPath.endswith('.ods') or documentPath.endswith('.csv'):
        odsPath = documentPath.replace('.csv', '.ods')
        csvPath = documentPath.replace('.ods', '.csv')

        # Save document in csv format

        args1[0].Name = 'URL'
        # args1(0).Name = "URL"
        args1[0].Value = csvPath
        # args1(0).Value = csvPath
        args1[1].Name = 'FilterName'
        # args1(1).Name = "FilterName"
        args1[1].Value = 'Text - txt - csv (StarCalc)'
        # args1(1).Value = "Text - txt - csv (StarCalc)"
        dispatcher.executeDispatch(document, ".uno:SaveAs", "", 0, args1)
        # dispatcher.executeDispatch(document, ".uno:SaveAs", "", 0, args1())

        # Save document in OpenDocument format

        args1.append(PropertyValue())

        args1[0].Value = odsPath
        # args1(0).Value = odsPath
        args1[1].Value = 'calc8'
        # args1(1).Value = "calc8"
        args1[2].Name = "FilterOptions"
        # args1(2).Name = "FilterOptions"
        args1[2].Value = "124,34,76,1,,0,false,true,true"
        # args1(2).Value = "124,34,76,1,,0,false,true,true"
        dispatcher.executeDispatch(document, ".uno:SaveAs", "", 0, args1)
        # dispatcher.executeDispatch(document, ".uno:SaveAs", "", 0, args1())

        targetPath = unquote(csvPath.replace('file:///', ''))

    else:
        msgbox('ERROR: File type of "' + os.path.normpath(documentPath) +
               '" not supported.', type_msg=ERRORBOX)

    Converter().run(targetPath, None)
