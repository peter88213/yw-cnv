"""Convert html or csv to yWriter format. 

Input file format: html (with visible or invisible chapter and scene tags).

Version @release

Copyright (c) 2020, peter88213
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import sys
import os

from pywriter.globals import *
from pywriter.html.html_proof import HtmlProof
from pywriter.html.html_manuscript import HtmlManuscript
from pywriter.html.html_scenedesc import HtmlSceneDesc
from pywriter.html.html_chapterdesc import HtmlChapterDesc
from pywriter.html.html_import import HtmlImport
from pywriter.yw.yw_file import YwFile
from pywriter.yw.yw_new_file import YwNewFile
from pywriter.converter.yw_cnv import YwCnv
from pywriter.csv.csv_scenelist import CsvSceneList
from pywriter.csv.csv_plotlist import CsvPlotList
from pywriter.csv.csv_charlist import CsvCharList
from pywriter.csv.csv_loclist import CsvLocList
from pywriter.csv.csv_itemlist import CsvItemList

import uno

from uno_wrapper.uno_tools import *

TAILS = [PROOF_SUFFIX + '.html', MANUSCRIPT_SUFFIX + '.html', SCENEDESC_SUFFIX + '.html',
         CHAPTERDESC_SUFFIX + '.html', PARTDESC_SUFFIX +
         '.html', SCENELIST_SUFFIX + '.csv',
         PLOTLIST_SUFFIX + '.csv', CHARLIST_SUFFIX + '.csv', LOCLIST_SUFFIX + '.csv', ITEMLIST_SUFFIX + '.csv']

YW_EXTENSIONS = ['.yw7', '.yw6', '.yw5']


def delete_tempfile(filePath):
    """If an Office file exists, delete the temporary file."""

    if filePath.endswith('.html'):

        if os.path.isfile(filePath.replace('.html', '.odt')):
            try:
                os.remove(filePath)
            except:
                pass

    elif filePath.endswith('.csv'):

        if os.path.isfile(filePath.replace('.csv', '.ods')):
            try:
                os.remove(filePath)
            except:
                pass


def run(sourcePath):
    sourcePath = sourcePath.replace('file:///', '').replace('%20', ' ')

    ywPath = None

    for tail in TAILS:
        # Determine the document type

        if sourcePath.endswith(tail):

            for ywExtension in YW_EXTENSIONS:
                # Determine the yWriter project file path

                testPath = sourcePath.replace(tail, ywExtension)

                if os.path.isfile(testPath):
                    ywPath = testPath
                    break

            break

    if ywPath:

        if tail == PROOF_SUFFIX + '.html':
            sourceDoc = HtmlProof(sourcePath)

        elif tail == MANUSCRIPT_SUFFIX + '.html':
            sourceDoc = HtmlManuscript(sourcePath)

        elif tail == SCENEDESC_SUFFIX + '.html':
            sourceDoc = HtmlSceneDesc(sourcePath)

        elif tail == CHAPTERDESC_SUFFIX + '.html':
            sourceDoc = HtmlChapterDesc(sourcePath)

        elif tail == PARTDESC_SUFFIX + '.html':
            sourceDoc = HtmlChapterDesc(sourcePath)

        elif tail == SCENELIST_SUFFIX + '.csv':
            sourceDoc = CsvSceneList(sourcePath)

        elif tail == PLOTLIST_SUFFIX + '.csv':
            sourceDoc = CsvPlotList(sourcePath)

        elif tail == CHARLIST_SUFFIX + '.csv':
            sourceDoc = CsvCharList(sourcePath)

        elif tail == LOCLIST_SUFFIX + '.csv':
            sourceDoc = CsvLocList(sourcePath)

        elif tail == ITEMLIST_SUFFIX + '.csv':
            sourceDoc = CsvItemList(sourcePath)

        else:
            return 'ERROR: File format not supported.'

        ywFile = YwFile(ywPath)
        converter = YwCnv()
        message = converter.document_to_yw(sourceDoc, ywFile)

    elif sourcePath.endswith('.html'):
        sourceDoc = HtmlImport(sourcePath)
        ywPath = sourcePath.replace('.html', '.yw7')
        ywFile = YwNewFile(ywPath)
        converter = YwCnv()
        message = converter.document_to_yw(sourceDoc, ywFile)

    else:
        message = 'ERROR: No yWriter project found.'

    if not message.startswith('ERROR'):
        delete_tempfile(sourcePath)

    return message


def export_yw(*args):
    '''Export the document to a yWriter 6/7 project.
    '''
    documentPath = ''

    # Get document's filename

    desktop = XSCRIPTCONTEXT.getDesktop()

    """
    document   = ThisComponent.CurrentController.Frame
    dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
    documentPath = ThisComponent.getURL()
    """
    documentPath = documentPath.lower()

    if documentPath.endswith('.odt') or documentPath.endswith('.html'):
        odtPath = documentPath.replace('.html', '.odt')
        htmlPath = documentPath.replace('.odt', '.html')
        """
        dim args1(1) as new com.sun.star.beans.PropertyValue

        # Save document in HTML format

        args1(0).Name = "URL"
        args1(0).Value = htmlPath
        args1(1).Name = "FilterName"
        args1(1).Value = "HTML (StarWriter)"

        dispatcher.executeDispatch(document, ".uno:SaveAs", "", 0, args1())

        # Save document in OpenDocument format

        args1(0).Value = odtPath
        args1(1).Value = "writer8"
        dispatcher.executeDispatch(document, ".uno:SaveAs", "", 0, args1())
        """
        result = run(htmlPath)

    elif documentPath.endswith('.ods') or documentPath.endswith('.csv'):
        odsPath = documentPath.replace('.csv', '.ods')
        csvPath = documentPath.replace('.ods', '.csv')
        """
        dim args1(1) as new com.sun.star.beans.PropertyValue

        # Save document in csv format

        args1(0).Name = "URL"
        args1(0).Value = csvPath
        args1(1).Name = "FilterName"
        args1(1).Value = "Text - txt - csv (StarCalc)"    
        dispatcher.executeDispatch(document, ".uno:SaveAs", "", 0, args1())

        # Save document in OpenDocument format

        args1(0).Value = odsPath
        args1(1).Value = "calc8"
        dispatcher.executeDispatch(document, ".uno:SaveAs", "", 0, args1())
        """
        result = run(csvPath)

    else:
        result = "ERROR: File type not supported."

    msgbox(result)


if __name__ == '__main__':
    try:
        sourcePath = sys.argv[1]
    except:
        sourcePath = ''
    print(run(sourcePath))
