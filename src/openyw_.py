"""Convert yWriter project to odt or csv. 

Input file format: yWriter
Output file format: odt (with visible or invisible chapter and scene tags) or csv.

Version @release

Copyright (c) 2020, peter88213
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import sys
import os

from pywriter.globals import *
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

import uno

from uno_wrapper.uno_tools import *
from uno_wrapper.uno_stub import *


def run(sourcePath, suffix):

    fileName, FileExtension = os.path.splitext(sourcePath)

    if suffix == PROOF_SUFFIX:
        targetDoc = OdtProof(fileName + suffix + '.odt')

    elif suffix == MANUSCRIPT_SUFFIX:
        targetDoc = OdtManuscript(fileName + suffix + '.odt')

    elif suffix == SCENEDESC_SUFFIX:
        targetDoc = OdtSceneDesc(fileName + suffix + '.odt')

    elif suffix == CHAPTERDESC_SUFFIX:
        targetDoc = OdtChapterDesc(fileName + suffix + '.odt')

    elif suffix == PARTDESC_SUFFIX:
        targetDoc = OdtPartDesc(fileName + suffix + '.odt')

    elif suffix == SCENELIST_SUFFIX:
        targetDoc = CsvSceneList(fileName + suffix + '.csv')

    elif suffix == PLOTLIST_SUFFIX:
        targetDoc = CsvPlotList(fileName + suffix + '.csv')

    else:
        targetDoc = OdtFile(fileName + '.odt')

    ywFile = YwFile(sourcePath)
    converter = YwCnv()
    message = converter.yw_to_document(ywFile, targetDoc)

    return message


def open_yw7(suffix):
    ywFile = FilePicker()
    sourcePath = ywFile.replace('file:///', '')
    extension = os.path.splitext(sourcePath)[1]

    if not extension in ['.yw6', '.yw7']:
        msgbox('Please choose an yWriter 6/7 project.')
        return

    newFile = ywFile.replace(extension, suffix + '.odt')
    dirName, filename = os.path.split(newFile)
    lockFile = (dirName + '/.~lock.' + filename + '#').replace('file:///', '')

    if os.path.isfile(lockFile):
        msgbox('Please close "' + filename + '" first.')
        return

    workdir = os.path.dirname(sourcePath)
    os.chdir(workdir)
    message = run(sourcePath, suffix)

    if message.startswith('ERROR'):
        msgbox(message)

    else:
        desktop = XSCRIPTCONTEXT.getDesktop()
        doc = desktop.loadComponentFromURL(newFile, "_blank", 0, ())


def import_yw(*args):
    '''Import scenes from yWriter 6/7 to a Writer document
    without chapter and scene markers. 
    '''
    open_yw7('')


def proof_yw(*args):
    '''Import scenes from yWriter 6/7 to a Writer document
    with visible chapter and scene markers. 
    '''
    open_yw7(PROOF_SUFFIX)
