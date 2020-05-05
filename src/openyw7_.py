"""Convert yWriter project to odt or csv. 

Input file format: yWriter
Output file format: odt (with visible or invisible chapter and scene tags) or csv.

Depends on the PyWriter library v1.6

Copyright (c) 2020 Peter Triesberger.
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import sys
import os

from pywriter.odt.odt_proof import OdtProof
from pywriter.odt.odt_manuscript import OdtManuscript
from pywriter.odt.odt_scenedesc import OdtSceneDesc
from pywriter.odt.odt_chapterdesc import OdtChapterDesc
from pywriter.odt.odt_partdesc import OdtPartDesc
from pywriter.yw.yw_file import YwFile
from pywriter.converter.yw_cnv import YwCnv
from pywriter.csv.csv_scenelist import CsvSceneList
from pywriter.csv.csv_plotlist import CsvPlotList
from pywriter.odt.odt_file import OdtFile


def run(sourcePath, suffix):

    fileName, FileExtension = os.path.splitext(sourcePath)

    if suffix == '_proof':
        ywFile = YwFile(sourcePath)
        targetDoc = OdtProof(
            fileName + suffix + '.odt')

    elif suffix == '_manuscript':
        ywFile = YwFile(sourcePath)
        targetDoc = OdtManuscript(
            fileName + suffix + '.odt')

    elif suffix == '_scenes':
        ywFile = YwFile(sourcePath)
        targetDoc = OdtSceneDesc(
            fileName + suffix + '.odt')

    elif suffix == '_chapters':
        ywFile = YwFile(sourcePath)
        targetDoc = OdtChapterDesc(
            fileName + suffix + '.odt')

    elif suffix == '_parts':
        ywFile = YwFile(sourcePath)
        targetDoc = OdtPartDesc(
            fileName + suffix + '.odt')

    elif suffix == '_scenelist':
        ywFile = YwFile(sourcePath)
        targetDoc = CsvSceneList(fileName + suffix + '.csv')

    elif suffix == '_plotlist':
        ywFile = YwFile(sourcePath)
        targetDoc = CsvPlotList(fileName + suffix + '.csv')

    else:
        ywFile = YwFile(sourcePath)
        targetDoc = OdtFile(fileName + '.odt')

    converter = YwCnv()
    message = converter.yw_to_document(ywFile, targetDoc)

    return message


if __name__ == '__main__':
    try:
        sourcePath = sys.argv[1]
    except:
        sourcePath = ''

    fileName, FileExtension = os.path.splitext(sourcePath)

    if FileExtension in ['.yw5', '.yw6', '.yw7']:
        try:
            suffix = sys.argv[2]
        except:
            suffix = ''

        print(run(sourcePath, suffix))

    else:
        print('ERROR: File is not an yWriter project.')
