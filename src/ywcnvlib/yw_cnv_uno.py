"""Provide a converter class for universal import and export.

Copyright (c) 2022 Peter Triesberger
For further information see https://github.com/peter88213/yw-cnv
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
from pywriter.pywriter_globals import *
from pywriter.converter.yw7_converter import Yw7Converter


class YwCnvUno(Yw7Converter):
    """A converter for universal import and export.
    
    Public methods:
        export_from_yw(sourceFile, targetFile) -- Convert from yWriter project to other file format.

    Support yWriter 7 projects and most of the Novel subclasses 
    that can be read or written by OpenOffice/LibreOffice.
    - No message in case of success when converting from yWriter.
    """

    def export_from_yw(self, source, target):
        """Convert from yWriter project to other file format.

        Positional arguments:
            source -- YwFile subclass instance.
            target -- Any Novel subclass instance.

        Show only error messages.
        Overrides the superclass method.
        """
        try:
            self.convert(source, target)
        except Error as ex:
            self.newFile = None
            self.ui.set_info_how(f'!{str(ex)}')
        else:
            self.newFile = target.filePath
