"""Provide a converter class for universal import and export.

Copyright (c) 2022 Peter Triesberger
For further information see https://github.com/peter88213/yw-cnv
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
from pywriter.pywriter_globals import ERROR
from pywriter.converter.yw7_converter import Yw7Converter


class YwCnvUno(Yw7Converter):
    """A converter for universal import and export.
    Support yWriter 7 projects and most of the Novel subclasses 
    that can be read or written by OpenOffice/LibreOffice.
    - No message in case of success when converting from yWriter.
    """

    def export_from_yw(self, source, target):
        """Method for conversion from yw to other.
        Override the superclass method.
        Show only error messages.
        """
        message = self.convert(source, target)

        if message.startswith(ERROR):
            self.ui.set_info_how(message)

        else:
            self.newFile = target.filePath
