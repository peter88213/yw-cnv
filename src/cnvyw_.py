"""Convert yw7 to odt/ods, or html/csv to yw7. 

Version @release
Requires Python 3.6+
Copyright (c) 2023 Peter Triesberger
For further information see https://github.com/peter88213/yw-cnv
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import uno
from com.sun.star.awt.MessageBoxType import MESSAGEBOX, INFOBOX, WARNINGBOX, ERRORBOX, QUERYBOX
from com.sun.star.beans import PropertyValue
import os
from configparser import ConfigParser
from ywcnvlib.uno_tools import *
from ywcnvlib.yw_cnv_uno import YwCnvUno
from ywcnvlib.ui_uno import UiUno

from pywriter.pywriter_globals import *
from pywriter.odt_w.odt_w_proof import OdtWProof
from pywriter.odt_w.odt_w_manuscript import OdtWManuscript
from pywriter.odt_w.odt_w_scenedesc import OdtWSceneDesc
from pywriter.odt_w.odt_w_chapterdesc import OdtWChapterDesc
from pywriter.odt_w.odt_w_partdesc import OdtWPartDesc
from pywriter.odt_w.odt_w_brief_synopsis import OdtWBriefSynopsis
from pywriter.odt_w.odt_w_characters import OdtWCharacters
from pywriter.odt_w.odt_w_items import OdtWItems
from pywriter.odt_w.odt_w_locations import OdtWLocations
from pywriter.odt_w.odt_w_xref import OdtWXref
from pywriter.odt_w.odt_w_notes import OdtWNotes
from pywriter.odt_w.odt_w_todo import OdtWTodo
from pywriter.ods_w.ods_w_charlist import OdsWCharList
from pywriter.ods_w.ods_w_loclist import OdsWLocList
from pywriter.ods_w.ods_w_itemlist import OdsWItemList
from pywriter.ods_w.ods_w_scenelist import OdsWSceneList

INI_FILE = 'openyw.ini'


def open_yw7(suffix, newExt):
    """Open a yWriter project, create a new document and load it.
    
    Positional arguments:
        suffix -- str: filename suffix of the document to create.
        newExt -- str: file extension of the document to create.   
    """

    # Set last opened yWriter project as default (if existing).
    scriptLocation = os.path.dirname(__file__)
    inifile = uno.fileUrlToSystemPath(f'{scriptLocation}/{INI_FILE}')
    defaultFile = None
    config = ConfigParser()
    try:
        config.read(inifile)
        ywLastOpen = config.get('FILES', 'yw_last_open')
        if os.path.isfile(ywLastOpen):
            defaultFile = uno.systemPathToFileUrl(ywLastOpen)
    except:
        pass

    # Ask for yWriter 7 project to open:
    ywFile = FilePicker(path=defaultFile)
    if ywFile is None:
        return

    sourcePath = uno.fileUrlToSystemPath(ywFile)
    __, ywExt = os.path.splitext(sourcePath)
    if not ywExt in ['.yw7']:
        msgbox(f'{_("Please choose a yWriter 7 project")}.', type_msg=ERRORBOX)
        return

    # Store selected yWriter project as "last opened".
    newFile = ywFile.replace(ywExt, f'{suffix}{newExt}')
    dirName, fileName = os.path.split(newFile)
    thisDir = uno.fileUrlToSystemPath(f'{dirName}/')
    lockFile = f'{thisDir}.~lock.{fileName}#'
    if not config.has_section('FILES'):
        config.add_section('FILES')
    config.set('FILES', 'yw_last_open', uno.fileUrlToSystemPath(ywFile))
    with open(inifile, 'w') as f:
        config.write(f)

    # Check if import file is already open in LibreOffice:
    if os.path.isfile(lockFile):
        msgbox(f'{_("Please close document first")}t: "{fileName}".', type_msg=ERRORBOX)
        return

    # Open yWriter project and convert data.
    workdir = os.path.dirname(sourcePath)
    os.chdir(workdir)
    converter = YwCnvUno()
    converter.ui = UiUno(_('Import from yWriter'))
    kwargs = {'suffix': suffix}
    converter.run(sourcePath, **kwargs)
    if converter.newFile:
        desktop = XSCRIPTCONTEXT.getDesktop()
        desktop.loadComponentFromURL(newFile, "_blank", 0, ())


def import_yw():
    '''Import scenes from yWriter 7 to a Writer document.'''
    open_yw7('', '.odt')


def proof_yw():
    '''Import scenes from yWriter 7 to a Writer document.'''
    open_yw7(OdtWProof.SUFFIX, OdtWProof.EXTENSION)


def get_brf_synopsis():
    '''Import chapter and scene titles from yWriter 7 to a Writer document.'''
    open_yw7(OdtWBriefSynopsis.SUFFIX, OdtWBriefSynopsis.EXTENSION)


def get_manuscript():
    '''Import scenes from yWriter 7 to a Writer document.'''
    open_yw7(OdtWManuscript.SUFFIX, OdtWManuscript.EXTENSION)


def get_partdesc():
    '''Import part descriptions from yWriter 7 to a Writer document.'''
    open_yw7(OdtWPartDesc.SUFFIX, OdtWPartDesc.EXTENSION)


def get_chapterdesc():
    '''Import chapter descriptions from yWriter 7 to a Writer document.'''
    open_yw7(OdtWChapterDesc.SUFFIX, OdtWChapterDesc.EXTENSION)


def get_scenedesc():
    '''Import scene descriptions from yWriter 7 to a Writer document.'''
    open_yw7(OdtWSceneDesc.SUFFIX, OdtWSceneDesc.EXTENSION)


def get_chardesc():
    '''Import character descriptions from yWriter 7 to a Writer document.'''
    open_yw7(OdtWCharacters.SUFFIX, OdtWCharacters.EXTENSION)


def get_locdesc():
    '''Import location descriptions from yWriter 7 to a Writer document.'''
    open_yw7(OdtWLocations.SUFFIX, OdtWLocations.EXTENSION)


def get_itemdesc():
    '''Import item descriptions from yWriter 7 to a Writer document.'''
    open_yw7(OdtWItems.SUFFIX, OdtWItems.EXTENSION)


def get_xref():
    '''Generate cross references from yWriter 7 to a Writer document.'''
    open_yw7(OdtWXref.SUFFIX, OdtWXref.EXTENSION)


def get_scenelist():
    '''Import a scene list from yWriter 7 to a Calc document.'''
    open_yw7(OdsWSceneList.SUFFIX, OdsWSceneList.EXTENSION)


def get_notes():
    '''Import Notes chapters from yWriter 7 to a Writer document.'''
    open_yw7(OdtWNotes.SUFFIX, OdtWNotes.EXTENSION)


def get_todo():
    '''Import Todo chapters from yWriter 7 to a Writer document.'''
    open_yw7(OdtWTodo.SUFFIX, OdtWTodo.EXTENSION)


def get_charlist():
    '''Import a character list from yWriter 7 to a Calc document.'''
    open_yw7(OdsWCharList.SUFFIX, OdsWCharList.EXTENSION)


def get_loclist():
    '''Import a location list from yWriter 7 to a Calc document.'''
    open_yw7(OdsWLocList.SUFFIX, OdsWLocList.EXTENSION)


def get_itemlist():
    '''Import an item list from yWriter 7 to a Calc document.'''
    open_yw7(OdsWItemList.SUFFIX, OdsWItemList.EXTENSION)


def export_yw():
    """Save the document if modified, and call the converter script."""
    thisComponent = XSCRIPTCONTEXT.getDocument()

    if thisComponent.isModified():
        thisComponent.store()

    documentUrl = thisComponent.getURL()
    if documentUrl:
        sourcePath = uno.fileUrlToSystemPath(documentUrl)
    else:
        sourcePath = ''
    converter = YwCnvUno()
    converter.ui = UiUno(_('Export to yWriter'))
    kwargs = {'suffix': None}
    converter.run(sourcePath, **kwargs)


def to_blank_lines():
    """Replace scene dividers with blank lines.

    Replace the three-lines "* * *" scene dividers with single blank lines. 
    Change the style of the scene-dividing paragraphs from  _Heading 4_  to  _Heading 5_.
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

    #--- Save cursor position.
    oViewCursor = XSCRIPTCONTEXT.getDocument().CurrentController.getViewCursor()
    # oViewCursor = ThisComponent.CurrentController().getViewCursor()
    oSaveCursor = XSCRIPTCONTEXT.getDocument().Text.createTextCursorByRange(oViewCursor)
    # oSaveCursor = ThisComponent.Text.createTextCursorByRange(oViewCursor)

    #--- Replace "Heading 4" by "Heading 5".
    args1 = []
    for __ in range(19):
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

    #--- Find all "Heading 5".
    args2 = []
    for __ in range(19):
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

    #--- Delete scene dividers.
    args3 = []
    for __ in range(19):
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

    #--- Reset search options with a dummy search.
    args3[9].Value = 65536
    args3[10].Value = "#"
    args3[17].Value = 1
    dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args3)

    #--- Restore cursor position.
    oViewCursor.gotoRange(oSaveCursor, False)


def indent_paragraphs():
    """Indent paragraphs that start with '> '.

    Select all paragraphs that start with '> ' 
    and change their paragraph style to _Quotations_.
    """
    pStyles = XSCRIPTCONTEXT.getDocument().StyleFamilies.getByName('ParagraphStyles')
    # pStyles = ThisComponent.StyleFamilies.getByName("ParagraphStyles")
    document = XSCRIPTCONTEXT.getDocument().CurrentController.Frame
    # document   = ThisComponent.CurrentController.Frame
    ctx = XSCRIPTCONTEXT.getComponentContext()
    smgr = ctx.getServiceManager()
    dispatcher = smgr.createInstanceWithContext("com.sun.star.frame.DispatchHelper", ctx)
    # dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")

    #--- Save cursor position.
    oViewCursor = XSCRIPTCONTEXT.getDocument().CurrentController.getViewCursor()
    # oViewCursor = ThisComponent.CurrentController().getViewCursor()
    oSaveCursor = XSCRIPTCONTEXT.getDocument().Text.createTextCursorByRange(oViewCursor)
    # oSaveCursor = ThisComponent.Text.createTextCursorByRange(oViewCursor)

    #--- Assign all paragraphs beginning with '> ' the 'Quotations' style.
    args1 = []
    for __ in range(19):
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
    args1[5].Value = False
    args1[6].Name = "SearchItem.Content"
    args1[6].Value = False
    args1[7].Name = "SearchItem.AsianOptions"
    args1[7].Value = False
    args1[8].Name = "SearchItem.AlgorithmType"
    args1[8].Value = 1
    args1[9].Name = "SearchItem.SearchFlags"
    args1[9].Value = 65536
    args1[10].Name = "SearchItem.SearchString"
    args1[10].Value = "^> "
    args1[11].Name = "SearchItem.ReplaceString"
    args1[11].Value = ""
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
    args1[17].Value = 1
    args1[18].Name = "Quiet"
    args1[18].Value = True
    dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args1)
    if is_anything_selected(XSCRIPTCONTEXT.getDocument()):
        args2 = []
        for __ in range(2):
            args2.append(PropertyValue())
        # dim args2(1) as new com.sun.star.beans.PropertyValue
        args2[0].Name = "Template"
        args2[0].Value = pStyles.getByName("Quotations").DisplayName
        args2[1].Name = "Family"
        args2[1].Value = 2
        dispatcher.executeDispatch(document, ".uno:StyleApply", "", 0, args2)

        #--- Delete the markup.
        args1[17].Value = 3
        dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args1)

    #--- Reset search options with a dummy search.
    args1[8].Value = 0
    args1[10].Value = "#"
    args1[17].Value = 1
    dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args1)

    #--- Restore cursor position.
    oViewCursor.gotoRange(oSaveCursor, False)


def replace_bullets():
    """Replace list strokes with bullets.

    Select all paragraphs that start with '- ' 
    and apply a list paragraph style.
    """
    document = XSCRIPTCONTEXT.getDocument().CurrentController.Frame
    # document   = ThisComponent.CurrentController.Frame
    ctx = XSCRIPTCONTEXT.getComponentContext()
    smgr = ctx.getServiceManager()
    dispatcher = smgr.createInstanceWithContext("com.sun.star.frame.DispatchHelper", ctx)
    # dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")

    #--- Save cursor position.
    oViewCursor = XSCRIPTCONTEXT.getDocument().CurrentController.getViewCursor()
    # oViewCursor = ThisComponent.CurrentController().getViewCursor()
    oSaveCursor = XSCRIPTCONTEXT.getDocument().Text.createTextCursorByRange(oViewCursor)
    # oSaveCursor = ThisComponent.Text.createTextCursorByRange(oViewCursor)

    #--- Find all list strokes.
    args1 = []
    for __ in range(19):
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
    args1[5].Value = False
    args1[6].Name = "SearchItem.Content"
    args1[6].Value = False
    args1[7].Name = "SearchItem.AsianOptions"
    args1[7].Value = False
    args1[8].Name = "SearchItem.AlgorithmType"
    args1[8].Value = 1
    args1[9].Name = "SearchItem.SearchFlags"
    args1[9].Value = 65536
    args1[10].Name = "SearchItem.SearchString"
    args1[10].Value = "^- "
    args1[11].Name = "SearchItem.ReplaceString"
    args1[11].Value = ""
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
    args1[17].Value = 1
    args1[18].Name = "Quiet"
    args1[18].Value = True
    dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args1)
    if is_anything_selected(XSCRIPTCONTEXT.getDocument()):
        #--- Apply list bullets to search result.
        args2 = []
        for __ in range(1):
            args2.append(PropertyValue())
        # dim args2(0) as new com.sun.star.beans.PropertyValue
        args2[0].Name = "On"
        args2[0].Value = True
        dispatcher.executeDispatch(document, ".uno:DefaultBullet", "", 0, args2)

        #--- Delete list strokes.
        dispatcher.executeDispatch(document, ".uno:Delete", "", 0, [])
        # dispatcher.executeDispatch(document, ".uno:Delete", "", 0, Array())

    #--- Reset search options with a dummy search.
    args1[8].Value = 0
    args1[10].Value = "#"
    args1[17].Value = 1
    dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args1)

    #--- Restore cursor position.
    oViewCursor.gotoRange(oSaveCursor, False)


def is_anything_selected(oDoc):
    """Return True if anything is selected.
    
    Positional arguments:
        oDoc -- ThisComponent
    
    Code example by Andrew D. Pitonyak
    OpenOffice.org Macros Explained
    OOME Third Edition
    """
    # Assume nothing is selected
    IsAnythingSelected = False
    if oDoc is None:
        return False

    # The current selection in the current controller.
    # If there is no current controller, it returns NULL.
    oSelections = oDoc.getCurrentSelection()
    if oSelections is None:
        return False

    if oSelections.getCount() == 0:
        return False

    if oSelections.getCount() > 1:
        # There is more than one selection so return True
        IsAnythingSelected = True
    else:
        # There is only one selection so obtain the first selection
        oSel = oSelections.getByIndex(0)
        # Create a text cursor that covers the range and then see if it is
        # collapsed.
        oCursor = oDoc.Text.createTextCursorByRange(oSel)
        if not oCursor.isCollapsed():
            IsAnythingSelected = True
    return IsAnythingSelected
