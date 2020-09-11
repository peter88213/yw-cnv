"""Import and export yWriter data. 

Standalone yWriter file converter with basic error handling 

Copyright (c) 2020 Peter Triesberger
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
from com.sun.star.awt.MessageBoxResults import OK, YES, NO, CANCEL
from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK, BUTTONS_OK_CANCEL, BUTTONS_YES_NO, BUTTONS_YES_NO_CANCEL, BUTTONS_RETRY_CANCEL, BUTTONS_ABORT_IGNORE_RETRY
from com.sun.star.awt.MessageBoxType import MESSAGEBOX, INFOBOX, WARNINGBOX, ERRORBOX, QUERYBOX

from pywriter.converter.yw_cnv import YwCnv
from pywriter.converter.file_factory import FileFactory
from pywriter.yw.yw7_new_file import Yw7NewFile
from uno_wrapper.uno_tools import *


class YwCnvUno(YwCnv):
    """Converter for yWriter project files.
    Variant with UNO UI.
    """

    def __init__(self, sourcePath, suffix=None):
        """Run the converter with a GUI. """

        self.success = False
        fileFactory = FileFactory()

        message, sourceFile, TargetFile = fileFactory.get_file_objects(
            sourcePath, suffix)

        if message.startswith('SUCCESS'):
            self.convert(sourceFile, TargetFile)

        else:
            msgbox(message)

    def convert(self, sourceFile, targetFile):
        """Determine the direction and invoke the converter. """

        # The conversion's direction depends on the sourcePath argument.

        if not sourceFile.file_exists():
            message = 'ERROR: File not found.'

        else:
            if sourceFile.EXTENSION in ['.yw5', '.yw6', '.yw7']:

                message = self.yw_to_document(sourceFile, targetFile)

            elif isinstance(targetFile, Yw7NewFile):

                if targetFile.file_exists():
                    message = 'ERROR: "' + targetFile._filePath + '" already exists.'

                else:
                    message = self.document_to_yw(sourceFile, targetFile)

            else:
                message = self.document_to_yw(sourceFile, targetFile)

            # Visualize the outcome.

            if message.startswith('SUCCESS'):
                self.success = True

            else:
                msgbox(message, type_msg=ERRORBOX)

    def confirm_overwrite(self, filePath):
        result = msgbox('Overwrite existing file "' + filePath + '"?',
                        'WARNING', buttons=BUTTONS_YES_NO, type_msg=WARNINGBOX)

        if result == YES:
            return True

        else:
            return False
