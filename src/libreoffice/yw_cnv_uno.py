"""User interface for the converter: UNO facade

Copyright (c) 2020 Peter Triesberger
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
from pywriter.converter.universal_converter import UniversalConverter


class YwCnvUno(UniversalConverter):
    """Converter for yWriter project files.
    Variant with UNO UI.
    """

    def export_from_yw(self, sourceFile, targetFile):
        """Method for conversion from yw to other.
        Show only error messages.
        """
        message = self.convert(sourceFile, targetFile)

        if message.startswith('SUCCESS'):
            self.newFile = targetFile.filePath

        else:
            self.ui.set_info_how(message)
