"""User interface for the converter: UNO facade

Copyright (c) 2020 Peter Triesberger
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import os

from pywriter.converter.yw_cnv_ui import YwCnvUi
from pywriter.converter.file_factory import FileFactory
from pywriter.converter.ui import Ui
from libreoffice.ui_uno import UiUno


class YwCnvUno(YwCnvUi):
    """Converter for yWriter project files.
    Variant with UNO UI.
    """

    def __init__(self, sourcePath, suffix=None, silentMode=False):
        """Run the converter with a GUI. """

        if silentMode:
            self.UserInterface = Ui('')

        else:
            self.UserInterface = UiUno('yWriter import/export')

        self.fileFactory = FileFactory()

        # Run the converter.

        self.success = False
        self.run_conversion(sourcePath, suffix)

    def export_from_yw(self, sourceFile, targetFile):
        """Method for conversion from yw to other.
        Show only error messages.
        """
        message = self.convert(sourceFile, targetFile)

        if message.startswith('SUCCESS'):
            self.success = True

        else:
            self.UserInterface.set_info_how(message)
