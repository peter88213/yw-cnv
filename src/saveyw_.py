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
from pywriter.converter.yw_cnv import YwCnv
from pywriter.csv.csv_scenelist import CsvSceneList
from pywriter.csv.csv_plotlist import CsvPlotList

TAILS = [PROOF_SUFFIX + '.html', MANUSCRIPT_SUFFIX + '.html', SCENEDESC_SUFFIX + '.html',
         CHAPTERDESC_SUFFIX + '.html', PARTDESC_SUFFIX +
         '.html', '.html', SCENELIST_SUFFIX + '.csv',
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

        elif tail == '.html':
            sourceDoc = HtmlImport(sourcePath)

        elif tail == SCENELIST_SUFFIX + '.csv':
            sourceDoc = CsvSceneList(sourcePath)

        elif tail == PLOTLIST_SUFFIX + '.csv':
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
