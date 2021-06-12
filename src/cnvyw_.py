"""Convert yWriter project to odt or ods and vice versa. 

Version @release

Copyright (c) 2021 Peter Triesberger
For further information see https://github.com/peter88213/yw-cnv
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import os

from configparser import ConfigParser
from urllib.parse import unquote
from urllib.parse import quote

from pywriter.odt.odt_proof import OdtProof
from pywriter.odt.odt_manuscript import OdtManuscript
from pywriter.odt.odt_scenedesc import OdtSceneDesc
from pywriter.odt.odt_chapterdesc import OdtChapterDesc
from pywriter.odt.odt_partdesc import OdtPartDesc
from pywriter.odt.odt_xref import OdtXref
from pywriter.ods.ods_scenelist import OdsSceneList
from pywriter.ods.ods_plotlist import OdsPlotList
from pywriter.ods.ods_charlist import OdsCharList
from pywriter.ods.ods_loclist import OdsLocList
from pywriter.ods.ods_itemlist import OdsItemList
from pywriter.odt.odt_characters import OdtCharacters
from pywriter.odt.odt_items import OdtItems
from pywriter.odt.odt_locations import OdtLocations

import uno
from com.sun.star.awt.MessageBoxType import MESSAGEBOX, INFOBOX, WARNINGBOX, ERRORBOX, QUERYBOX
from com.sun.star.beans import PropertyValue

from libreoffice.uno_tools import *
from libreoffice.yw_cnv_uno import YwCnvUno
from libreoffice.ui_uno import UiUno

INI_FILE = 'openyw.ini'


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

    # Ask for yWriter 7 project to open:

    ywFile = FilePicker(path=defaultFile)

    if ywFile is None:
        return

    sourcePath = unquote(ywFile.replace('file:///', ''))
    ywExt = os.path.splitext(sourcePath)[1]

    if not ywExt in ['.yw7']:
        msgbox('Please choose a yWriter 7 project.',
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
    converter = YwCnvUno()
    converter.ui = UiUno('Import from yWriter')
    kwargs = {'suffix': suffix}
    converter.run(sourcePath, **kwargs)

    if converter.newFile:
        desktop = XSCRIPTCONTEXT.getDesktop()
        doc = desktop.loadComponentFromURL(newFile, "_blank", 0, ())


def import_yw():
    '''Import scenes fromyWriter 7 to a Writer document
    without chapter and scene markers. 
    '''
    open_yw7('', '.odt')


def proof_yw():
    '''Import scenes fromyWriter 7 to a Writer document
    with visible chapter and scene markers. 
    '''
    open_yw7(OdtProof.SUFFIX, OdtProof.EXTENSION)


def get_manuscript():
    '''Import scenes fromyWriter 7 to a Writer document
    with invisible chapter and scene markers. 
    '''
    open_yw7(OdtManuscript.SUFFIX, OdtManuscript.EXTENSION)


def get_partdesc():
    '''Import pard descriptions fromyWriter 7 to a Writer document
    with invisible chapter and scene markers. 
    '''
    open_yw7(OdtPartDesc.SUFFIX, OdtPartDesc.EXTENSION)


def get_chapterdesc():
    '''Import chapter descriptions fromyWriter 7 to a Writer document
    with invisible chapter and scene markers. 
    '''
    open_yw7(OdtChapterDesc.SUFFIX, OdtChapterDesc.EXTENSION)


def get_scenedesc():
    '''Import scene descriptions fromyWriter 7 to a Writer document
    with invisible chapter and scene markers. 
    '''
    open_yw7(OdtSceneDesc.SUFFIX, OdtSceneDesc.EXTENSION)


def get_chardesc():
    '''Import character descriptions fromyWriter 7 to a Writer document.
    '''
    open_yw7(OdtCharacters.SUFFIX, OdtCharacters.EXTENSION)


def get_locdesc():
    '''Import location descriptions fromyWriter 7 to a Writer document.
    '''
    open_yw7(OdtLocations.SUFFIX, OdtLocations.EXTENSION)


def get_itemdesc():
    '''Import item descriptions fromyWriter 7 to a Writer document.
    '''
    open_yw7(OdtItems.SUFFIX, OdtItems.EXTENSION)


def get_xref():
    '''Generate cross references fromyWriter 7 to a Writer document.
    '''
    open_yw7(OdtXref.SUFFIX, OdtXref.EXTENSION)


def get_scenelist():
    '''Import a scene list fromyWriter 7 to a Calc document.
    '''
    open_yw7(OdsSceneList.SUFFIX, OdsSceneList.EXTENSION)


def get_plotlist():
    '''Import a plot list fromyWriter 7 to a Calc document.
    '''
    open_yw7(OdsPlotList.SUFFIX, OdsPlotList.EXTENSION)


def get_charlist():
    '''Import a character list fromyWriter 7 to a Calc document.
    '''
    open_yw7(OdsCharList.SUFFIX, OdsCharList.EXTENSION)


def get_loclist():
    '''Import a location list fromyWriter 7 to a Calc document.
    '''
    open_yw7(OdsLocList.SUFFIX, OdsLocList.EXTENSION)


def get_itemlist():
    '''Import an item list fromyWriter 7 to a Calc document.
    '''
    open_yw7(OdsItemList.SUFFIX, OdsItemList.EXTENSION)


def export_yw():
    '''Export the document to ayWriter 7 project.
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

        args1.append(PropertyValue())

        args1[0].Name = 'URL'
        # args1(0).Name = "URL"
        args1[0].Value = csvPath
        # args1(0).Value = csvPath
        args1[1].Name = 'FilterName'
        # args1(1).Name = "FilterName"
        args1[1].Value = 'Text - txt - csv (StarCalc)'
        # args1(1).Value = "Text - txt - csv (StarCalc)"
        args1[2].Name = "FilterOptions"
        # args1(2).Name = "FilterOptions"
        args1[2].Value = "124,34,76,1,,0,false,true,true"
        # args1(2).Value = "124,34,76,1,,0,false,true,true"
        dispatcher.executeDispatch(document, ".uno:SaveAs", "", 0, args1)
        # dispatcher.executeDispatch(document, ".uno:SaveAs", "", 0, args1())

        # Save document in OpenDocument format

        args1[0].Value = odsPath
        # args1(0).Value = odsPath
        args1[1].Value = 'calc8'
        # args1(1).Value = "calc8"
        dispatcher.executeDispatch(document, ".uno:SaveAs", "", 0, args1)
        # dispatcher.executeDispatch(document, ".uno:SaveAs", "", 0, args1())

        targetPath = unquote(csvPath.replace('file:///', ''))

    else:
        msgbox('ERROR: File type of "' + os.path.normpath(documentPath) +
               '" not supported.', type_msg=ERRORBOX)

    converter = YwCnvUno()
    converter.ui = UiUno('Export to yWriter')
    kwargs = {'suffix': None}
    converter.run(targetPath, **kwargs)


def to_blank_lines():
    """Replace scene dividers with blank lines

    This will replace the three-line "* * *" scene dividers
    with single blank lines. The style of the scene-dividing
    lines will be changed from  _Heading 4_  to  _Heading 5_.
    """

    pStyles = XSCRIPTCONTEXT.getDocument().StyleFamilies.getByName('ParagraphStyles')
    # pStyles = ThisComponent.StyleFamilies.getByName("ParagraphStyles")

    document = XSCRIPTCONTEXT.getDocument().CurrentController.Frame
    # document   = ThisComponent.CurrentController.Frame

    ctx = XSCRIPTCONTEXT.getComponentContext()
    smgr = ctx.getServiceManager()
    dispatcher = smgr.createInstanceWithContext(
        "com.sun.star.frame.DispatchHelper", ctx)
    # dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")

    args1 = []

    for i in range(19):
        args1.append(PropertyValue())
    # dim args1(18) as new com.sun.star.beans.PropertyValue

    args1[0].Name = "SearchItem.StyleFamily"
    args1[0].Value = 2
    args1[1].Name = "SearchItem.CellType"
    args1[1].Value = 0
    args1[2].Name = "SearchItem.RowDirection"
    args1[2].Value = True
    args1[3].Name = "SearchItem.AllTables"
    args1[3].Value = False
    args1[4].Name = "SearchItem.Backward"
    args1[4].Value = False
    args1[5].Name = "SearchItem.Pattern"
    args1[5].Value = True
    args1[6].Name = "SearchItem.Content"
    args1[6].Value = False
    args1[7].Name = "SearchItem.AsianOptions"
    args1[7].Value = False
    args1[8].Name = "SearchItem.AlgorithmType"
    args1[8].Value = 0
    args1[9].Name = "SearchItem.SearchFlags"
    args1[9].Value = 65536
    args1[10].Name = "SearchItem.SearchString"
    args1[10].Value = pStyles.getByName("Heading 4").DisplayName
    args1[11].Name = "SearchItem.ReplaceString"
    args1[11].Value = pStyles.getByName("Heading 5").DisplayName
    args1[12].Name = "SearchItem.Locale"
    args1[12].Value = 255
    args1[13].Name = "SearchItem.ChangedChars"
    args1[13].Value = 2
    args1[14].Name = "SearchItem.DeletedChars"
    args1[14].Value = 2
    args1[15].Name = "SearchItem.InsertedChars"
    args1[15].Value = 2
    args1[16].Name = "SearchItem.TransliterateFlags"
    args1[16].Value = 1280
    args1[17].Name = "SearchItem.Command"
    args1[17].Value = 3
    args1[18].Name = "Quiet"
    args1[18].Value = True

    dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args1)

    args2 = []

    for i in range(19):
        args2.append(PropertyValue())
    # dim args2(18) as new com.sun.star.beans.PropertyValue

    args2[0].Name = "SearchItem.StyleFamily"
    args2[0].Value = 2
    args2[1].Name = "SearchItem.CellType"
    args2[1].Value = 0
    args2[2].Name = "SearchItem.RowDirection"
    args2[2].Value = True
    args2[3].Name = "SearchItem.AllTables"
    args2[3].Value = False
    args2[4].Name = "SearchItem.Backward"
    args2[4].Value = False
    args2[5].Name = "SearchItem.Pattern"
    args2[5].Value = True
    args2[6].Name = "SearchItem.Content"
    args2[6].Value = False
    args2[7].Name = "SearchItem.AsianOptions"
    args2[7].Value = False
    args2[8].Name = "SearchItem.AlgorithmType"
    args2[8].Value = 0
    args2[9].Name = "SearchItem.SearchFlags"
    args2[9].Value = 65536
    args2[10].Name = "SearchItem.SearchString"
    args2[10].Value = pStyles.getByName("Heading 5").DisplayName
    args2[11].Name = "SearchItem.ReplaceString"
    args2[11].Value = pStyles.getByName("Heading 5").DisplayName
    args2[12].Name = "SearchItem.Locale"
    args2[12].Value = 255
    args2[13].Name = "SearchItem.ChangedChars"
    args2[13].Value = 2
    args2[14].Name = "SearchItem.DeletedChars"
    args2[14].Value = 2
    args2[15].Name = "SearchItem.InsertedChars"
    args2[15].Value = 2
    args2[16].Name = "SearchItem.TransliterateFlags"
    args2[16].Value = 1280
    args2[17].Name = "SearchItem.Command"
    args2[17].Value = 1
    args2[18].Name = "Quiet"
    args2[18].Value = True

    dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args2)

    args3 = []

    for i in range(19):
        args3.append(PropertyValue())
    # dim args3(18) as new com.sun.star.beans.PropertyValue

    args3[0].Name = "SearchItem.StyleFamily"
    args3[0].Value = 2
    args3[1].Name = "SearchItem.CellType"
    args3[1].Value = 0
    args3[2].Name = "SearchItem.RowDirection"
    args3[2].Value = True
    args3[3].Name = "SearchItem.AllTables"
    args3[3].Value = False
    args3[4].Name = "SearchItem.Backward"
    args3[4].Value = False
    args3[5].Name = "SearchItem.Pattern"
    args3[5].Value = False
    args3[6].Name = "SearchItem.Content"
    args3[6].Value = False
    args3[7].Name = "SearchItem.AsianOptions"
    args3[7].Value = False
    args3[8].Name = "SearchItem.AlgorithmType"
    args3[8].Value = 0
    args3[9].Name = "SearchItem.SearchFlags"
    args3[9].Value = 71680
    args3[10].Name = "SearchItem.SearchString"
    args3[10].Value = "* * *"
    args3[11].Name = "SearchItem.ReplaceString"
    args3[11].Value = ""
    args3[12].Name = "SearchItem.Locale"
    args3[12].Value = 255
    args3[13].Name = "SearchItem.ChangedChars"
    args3[13].Value = 2
    args3[14].Name = "SearchItem.DeletedChars"
    args3[14].Value = 2
    args3[15].Name = "SearchItem.InsertedChars"
    args3[15].Value = 2
    args3[16].Name = "SearchItem.TransliterateFlags"
    args3[16].Value = 1280
    args3[17].Name = "SearchItem.Command"
    args3[17].Value = 3
    args3[18].Name = "Quiet"
    args3[18].Value = True

    dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args3)
