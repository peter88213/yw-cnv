"""Convert html or csv to yWriter format. 

Input file format: html (with visible or invisible chapter and scene tags).

Depends on the PyWriter library v1.6

Copyright (c) 2020, peter88213
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import sys
import os

from pywriter.globals import (PROOF_HTML, MANUSCRIPT_HTML, SCENEDESC_HTML, CHAPTERDESC_HTML,
                              PARTDESC_HTML, SCENELIST_CSV, PLOTLIST_CSV, CHARLIST_CSV, LOCLIST_CSV, ITEMLIST_CSV)
from pywriter.html.html_proof import HtmlProof
from pywriter.html.html_manuscript import HtmlManuscript
from pywriter.html.html_scenedesc import HtmlSceneDesc
from pywriter.html.html_chapterdesc import HtmlChapterDesc
from pywriter.yw.yw_file import YwFile
from pywriter.converter.yw_cnv import YwCnv
from pywriter.csv.csv_scenelist import CsvSceneList
from pywriter.csv.csv_plotlist import CsvPlotList

TAILS = [PROOF_HTML, MANUSCRIPT_HTML, SCENEDESC_HTML,
         CHAPTERDESC_HTML, PARTDESC_HTML, SCENELIST_CSV,
         PLOTLIST_CSV, CHARLIST_CSV, LOCLIST_CSV, ITEMLIST_CSV]

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

    for tail in TAILS:
        # Determine the document type

        if sourcePath.endswith(tail):

            ywPath = None

            for ywExtension in YW_EXTENSIONS:
                # Determine the yWriter project file path

                testPath = sourcePath.replace(tail, ywExtension)

                if os.path.isfile(testPath):
                    ywPath = testPath
                    break

            break

    if ywPath:

        if tail == PROOF_HTML:
            sourceDoc = HtmlProof(sourcePath)

        elif tail == MANUSCRIPT_HTML:
            sourceDoc = HtmlManuscript(sourcePath)

        elif tail == SCENEDESC_HTML:
            sourceDoc = HtmlSceneDesc(sourcePath)

        elif tail == CHAPTERDESC_HTML:
            sourceDoc = HtmlChapterDesc(sourcePath)

        elif tail == PARTDESC_HTML:
            sourceDoc = HtmlChapterDesc(sourcePath)

        elif tail == SCENELIST_CSV:
            sourceDoc = CsvSceneList(sourcePath)

        elif tail == PLOTLIST_CSV:
            sourceDoc = CsvPlotList(sourcePath)

        else:
            return 'ERROR: File format not supported.'

        ywFile = YwFile(ywPath)
        converter = YwCnv()
        message = converter.document_to_yw(sourceDoc, ywFile)

    else:
        message = 'ERROR: No yWriter project found.'

    if not message.startswith('ERROR'):
        delete_tempfile(sourcePath)

    return message


if __name__ == '__main__':
    try:
        sourcePath = sys.argv[1]
    except:
        sourcePath = ''
    print(run(sourcePath))
