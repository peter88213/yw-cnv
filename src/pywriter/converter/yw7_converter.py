"""Provide a converter class for yWriter 7 universal import and export.

Copyright (c) 2022 Peter Triesberger
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
from pywriter.converter.yw_cnv_ff import YwCnvFf
from pywriter.converter.new_project_factory import NewProjectFactory
from pywriter.yw.yw7_file import Yw7File
from pywriter.odt.odt_proof import OdtProof
from pywriter.odt.odt_manuscript import OdtManuscript
from pywriter.odt.odt_scenedesc import OdtSceneDesc
from pywriter.odt.odt_chapterdesc import OdtChapterDesc
from pywriter.odt.odt_partdesc import OdtPartDesc
from pywriter.odt.odt_brief_synopsis import OdtBriefSynopsis
from pywriter.odt.odt_export import OdtExport
from pywriter.odt.odt_characters import OdtCharacters
from pywriter.odt.odt_items import OdtItems
from pywriter.odt.odt_locations import OdtLocations
from pywriter.odt.odt_xref import OdtXref
from pywriter.odt.odt_notes import OdtNotes
from pywriter.odt.odt_todo import OdtTodo
from pywriter.ods.ods_charlist import OdsCharList
from pywriter.ods.ods_loclist import OdsLocList
from pywriter.ods.ods_itemlist import OdsItemList
from pywriter.ods.ods_scenelist import OdsSceneList
from pywriter.html.html_proof import HtmlProof
from pywriter.html.html_manuscript import HtmlManuscript
from pywriter.html.html_notes import HtmlNotes
from pywriter.html.html_todo import HtmlTodo
from pywriter.html.html_scenedesc import HtmlSceneDesc
from pywriter.html.html_chapterdesc import HtmlChapterDesc
from pywriter.html.html_partdesc import HtmlPartDesc
from pywriter.html.html_characters import HtmlCharacters
from pywriter.html.html_locations import HtmlLocations
from pywriter.html.html_items import HtmlItems
from pywriter.csv.csv_scenelist import CsvSceneList
from pywriter.csv.csv_charlist import CsvCharList
from pywriter.csv.csv_loclist import CsvLocList
from pywriter.csv.csv_itemlist import CsvItemList


class Yw7Converter(YwCnvFf):
    """A converter for universal import and export.

    Support yWriter 7 projects and most of the Novel subclasses 
    that can be read or written by OpenOffice/LibreOffice.

    Overrides the superclass constants EXPORT_SOURCE_CLASSES,
    EXPORT_TARGET_CLASSES, IMPORT_SOURCE_CLASSES, IMPORT_TARGET_CLASSES.

    Class constants:
        CREATE_SOURCE_CLASSES -- list of classes that - additional to HtmlImport
                        and HtmlOutline - can be exported to a new yWriter project.
    """
    EXPORT_SOURCE_CLASSES = [Yw7File]
    EXPORT_TARGET_CLASSES = [OdtExport,
                             OdtProof,
                             OdtManuscript,
                             OdtBriefSynopsis,
                             OdtSceneDesc,
                             OdtChapterDesc,
                             OdtPartDesc,
                             OdtCharacters,
                             OdtItems,
                             OdtLocations,
                             OdsCharList,
                             OdsLocList,
                             OdsItemList,
                             OdsSceneList,
                             OdtXref,
                             OdtNotes,
                             OdtTodo,
                             ]
    IMPORT_SOURCE_CLASSES = [HtmlProof,
                             HtmlManuscript,
                             HtmlSceneDesc,
                             HtmlChapterDesc,
                             HtmlPartDesc,
                             HtmlCharacters,
                             HtmlItems,
                             HtmlLocations,
                             HtmlNotes,
                             HtmlTodo,
                             CsvCharList,
                             CsvLocList,
                             CsvItemList,
                             CsvSceneList,
                             ]
    IMPORT_TARGET_CLASSES = [Yw7File]
    CREATE_SOURCE_CLASSES = []

    def __init__(self):
        """Change the newProjectFactory strategy.
        
        Extends the superclass constructor.
        """
        super().__init__()
        self.newProjectFactory = NewProjectFactory(self.CREATE_SOURCE_CLASSES)
