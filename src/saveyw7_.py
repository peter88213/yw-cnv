"""Convert html or csv to yWriter format. 

Input file format: html (with visible or invisible chapter and scene tags).

Depends on the PyWriter library v1.6

Copyright (c) 2020 Peter Triesberger.
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import sys
import os

from pywriter.html.html_proof import HtmlProof
from pywriter.html.html_manuscript import HtmlManuscript
from pywriter.html.html_scenedesc import HtmlSceneDesc
from pywriter.html.html_chapterdesc import HtmlChapterDesc
from pywriter.yw.yw_file import YwFile
from pywriter.converter.yw_cnv import YwCnv
from pywriter.csv.csv_scenelist import CsvSceneList
from pywriter.csv.csv_plotlist import CsvPlotList

TAILS = ['_proof.html', '_manuscript.html', '_scenes.html',
         '_chapters.html', '_parts.html', '_scenelist.csv', '_plotlist.csv']


def delete_tempfile(filePath):

    if filePath.endswith('.html'):

        if os.path.isfile(filePath.split('.html')[0] + '.odt'):
            try:
                os.remove(filePath)
            except:
                pass

    elif filePath.endswith('.csv'):

        if os.path.isfile(filePath.split('.csv')[0] + '.ods'):
            try:
                os.remove(filePath)
            except:
                pass


def run(sourcePath):
    sourcePath = sourcePath.replace('file:///', '').replace('%20', ' ')

    # Determine the project file path.

    for tail in TAILS:

        if sourcePath.endswith(tail):
            ywPath = sourcePath.replace(tail, '.yw7')

            if not os.path.isfile(ywPath):
                ywPath = sourcePath.replace(tail, '.yw6')

                if not os.path.isfile(ywPath):
                    ywPath = sourcePath.replace(tail, '.yw5')

                    if not os.path.isfile(ywPath):
                        ywPath = None
                        message = 'ERROR: No yWriter project found.'

            break

    if ywPath:

        if tail == '_proof.html':
            sourceDoc = HtmlProof(sourcePath)

        elif tail == '_manuscript.html':
            sourceDoc = HtmlManuscript(sourcePath)

        elif tail == '_scenes.html':
            sourceDoc = HtmlSceneDesc(sourcePath)

        elif tail == '_chapters.html':
            sourceDoc = HtmlChapterDesc(sourcePath)

        elif tail == '_parts.html':
            sourceDoc = HtmlChapterDesc(sourcePath)

        elif tail == '_scenelist.csv':
            sourceDoc = CsvSceneList(sourcePath)

        elif tail == '_plotlist.csv':
            sourceDoc = CsvPlotList(sourcePath)

        else:
            return 'ERROR: File format not supported.'

        ywFile = YwFile(ywPath)
        converter = YwCnv()
        message = converter.document_to_yw(sourceDoc, ywFile)

    if not message.startswith('ERROR'):
        delete_tempfile(sourcePath)

    return message


if __name__ == '__main__':
    try:
        sourcePath = sys.argv[1]
    except:
        sourcePath = ''
    print(run(sourcePath))
