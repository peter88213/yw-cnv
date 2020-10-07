"""User interface for the converter: UNO facade

Copyright (c) 2020 Peter Triesberger
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
from pywriter.converter.yw_cnv_ui import YwCnvUi
from pywriter.converter.universal_file_factory import UniversalFileFactory
from pywriter.converter.ui import Ui
from libreoffice.ui_uno import UiUno


class YwCnvUno(YwCnvUi):
    """Converter for yWriter project files.
    Variant with UNO UI.
    """

    def __init__(self, silentMode=False):
        if silentMode:
            self.userInterface = Ui('')

        else:
            self.userInterface = UiUno('yWriter import/export')

        self.success = False
        self.fileFactory = None

    def run(self, sourcePath, suffix=None):
        YwCnvUi.run(self, sourcePath, suffix)

        if self.success:
            self.delete_tempfile(sourcePath)

    def export_from_yw(self, sourceFile, targetFile):
        """Method for conversion from yw to other.
        Show only error messages.
        """
        message = self.convert(sourceFile, targetFile)

        if message.startswith('SUCCESS'):
            self.success = True

        else:
            self.userInterface.set_info_how(message)
