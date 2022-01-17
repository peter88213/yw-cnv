"""Provide a converter class for universal import and export.

Copyright (c) 2022 Peter Triesberger
For further information see https://github.com/peter88213/yw-cnv
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
from pywriter.converter.yw7_converter import Yw7Converter


class YwCnvUno(Yw7Converter):
    """A converter for universal import and export.
    Support yWriter 7 projects and most of the Novel subclasses 
    that can be read or written by OpenOffice/LibreOffice.
    - No message in case of success when converting from yWriter.
    """

    def export_from_yw(self, sourceFile, targetFile):
        """Method for conversion from yw to other.
        Override the superclass method.
        Show only error messages.
        """
        message = self.convert(sourceFile, targetFile)

        if message.startswith('SUCCESS'):
            self.newFile = targetFile.filePath

        else:
            self.ui.set_info_how(message)
