"""Convert yw7 to odt ot csv. 

Input file format: yw7
Output file format: odt (with visible or invisible chapter and scene tags) or csv.

Depends on the PyWriter library v1.5

Copyright (c) 2020 Peter Triesberger.
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import sys

from pywriter.fileop.odt_proof_writer import OdtProofWriter
from pywriter.fileop.odt_manuscript_writer import OdtManuscriptWriter
from pywriter.fileop.odt_scenedesc_writer import OdtSceneDescWriter
from pywriter.fileop.odt_chapterdesc_writer import OdtChapterDescWriter
from pywriter.fileop.odt_partdesc_writer import OdtPartDescWriter
from pywriter.fileop.yw7file import Yw7File
from pywriter.converter.yw7cnv import Yw7Cnv
from pywriter.fileop.scenelist import SceneList
from pywriter.plot.plotlist import PlotList
from pywriter.fileop.odt_file_writer import OdtFileWriter


def run(sourcePath, suffix):

    if suffix == '_proof':
        yw7File = Yw7File(sourcePath)
        targetDoc = OdtProofWriter(
            sourcePath.split('.yw7')[0] + suffix + '.odt')

    elif suffix == '_manuscript':
        yw7File = Yw7File(sourcePath)
        targetDoc = OdtManuscriptWriter(
            sourcePath.split('.yw7')[0] + suffix + '.odt')

    elif suffix == '_scenes':
        yw7File = Yw7File(sourcePath)
        targetDoc = OdtSceneDescWriter(
            sourcePath.split('.yw7')[0] + suffix + '.odt')

    elif suffix == '_chapters':
        yw7File = Yw7File(sourcePath)
        targetDoc = OdtChapterDescWriter(
            sourcePath.split('.yw7')[0] + suffix + '.odt')

    elif suffix == '_parts':
        yw7File = Yw7File(sourcePath)
        targetDoc = OdtPartDescWriter(
            sourcePath.split('.yw7')[0] + suffix + '.odt')

    elif suffix == '_scenelist':
        yw7File = Yw7File(sourcePath)
        targetDoc = SceneList(sourcePath.split('.yw7')[0] + suffix + '.csv')

    elif suffix == '_plotlist':
        yw7File = Yw7File(sourcePath)
        targetDoc = PlotList(sourcePath.split('.yw7')[0] + suffix + '.csv')

    else:
        yw7File = Yw7File(sourcePath)
        targetDoc = OdtFileWriter(sourcePath.split('.yw7')[0] + '.odt')

    converter = Yw7Cnv()
    message = converter.yw7_to_document(yw7File, targetDoc)

    return message


if __name__ == '__main__':
    try:
        sourcePath = sys.argv[1]
    except:
        sourcePath = ''

    if sourcePath.endswith('.yw7'):
        try:
            suffix = sys.argv[2]
        except:
            suffix = ''

        print(run(sourcePath, suffix))

    else:
        print('ERROR: File is not an yWriter 7 project.')
