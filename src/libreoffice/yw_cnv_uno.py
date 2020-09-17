"""User interface for the converter: UNO facade

Copyright (c) 2020 Peter Triesberger
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import os

from com.sun.star.awt.MessageBoxResults import OK, YES, NO, CANCEL
from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK, BUTTONS_OK_CANCEL, BUTTONS_YES_NO, BUTTONS_YES_NO_CANCEL, BUTTONS_RETRY_CANCEL, BUTTONS_ABORT_IGNORE_RETRY
from com.sun.star.awt.MessageBoxType import MESSAGEBOX, INFOBOX, WARNINGBOX, ERRORBOX, QUERYBOX

from pywriter.converter.yw_cnv import YwCnv
from pywriter.converter.file_factory import FileFactory
from pywriter.yw.yw7_tree_creator import Yw7TreeCreator
from libreoffice.uno_tools import *
from libreoffice.cnv_ui_uno import CnvUiUno


class YwCnvUno(YwCnv):
    """Converter for yWriter project files.
    Variant with UNO UI.
    """

    def __init__(self, sourcePath, suffix=None, silentMode=False):
        """Run the converter with a GUI. """

        self.silentMode = silentMode
        fileFactory = FileFactory()

        # Initialize the GUI

        self.cnvUi = CnvUiUno()

        # Run the converter.

        self.success = False
        message, sourceFile, TargetFile = fileFactory.get_file_objects(
            sourcePath, suffix)

        if message.startswith('SUCCESS'):
            self.convert(sourceFile, TargetFile)

        else:
            self.cnvUi.set_info_how(message)

    def convert(self, sourceFile, targetFile):
        """Determine the direction and invoke the converter. """
        showBox = False

        # The conversion's direction depends on the sourcePath argument.

        if not sourceFile.file_exists():
            message = 'ERROR: "' + \
                os.path.normpath(sourceFile.filePath) + '" File not found.'

        else:
            if sourceFile.EXTENSION in FileFactory.YW_EXTENSIONS:

                message = YwCnv.convert(self, sourceFile, targetFile)

            elif isinstance(targetFile.ywTreeBuilder, Yw7TreeCreator):

                if targetFile.file_exists():
                    message = 'ERROR: "' + \
                        os.path.normpath(targetFile.filePath) + \
                        '" already exists.'

                else:
                    message = YwCnv.convert(self, sourceFile, targetFile)
                    showBox = True

            else:
                message = YwCnv.convert(self, sourceFile, targetFile)
                showBox = True

            # Visualize the outcome.

            if message.startswith('SUCCESS'):
                self.success = True
                msgType = INFOBOX

            else:
                msgType = ERRORBOX
                showBox = True

            if showBox:
                msgbox(message, type_msg=msgType)

    def confirm_overwrite(self, filePath):
        """ Invoked by the parent if a file already exists. """

        if self.silentMode:
            return True

        else:
            return self.cnvUi.ask_yes_no('Overwrite existing file "' + os.path.normpath(filePath) + '"?')

    def edit(self):
        pass
