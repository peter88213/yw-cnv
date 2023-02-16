"""Provide a generic class for ODS file import.

Other ODS file readers inherit from this class.

Copyright (c) 2023 Peter Triesberger
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
from pywriter.pywriter_globals import *
from pywriter.file.file import File
from pywriter.file.file_export import FileExport
from pywriter.ods_r.ods_parser import OdsParser


class OdsReader(File):
    """Generic OpenDocument spreadsheet document reader.

    Public methods:
        read() -- parse the file and get the instance variables.

    """
    EXTENSION = '.ods'
    # overwrites File.EXTENSION
    _SEPARATOR = ','
    # delimits data fields within a record.
    _rowTitles = []

    _DIVIDER = FileExport._DIVIDER

    def __init__(self, filePath, **kwargs):
        """Initialize instance variables.

        Positional arguments:
            filePath -- str: path to the file represented by the File instance.
            
        Optional arguments:
            kwargs -- keyword arguments to be used by subclasses.            
        
        Extends the superclass constructor.
        """
        super().__init__(filePath)
        self._rows = []

    def read(self):
        """Parse the file and get the instance variables.
        
        Parse the ODS file located at filePath, fetching the rows.
        Check the number of fields in each row.
        Raise the "Error" exception in case of error. 
        Overrides the superclass method.
        """
        self._rows = []
        cellsPerRow = len(self._rowTitles)
        reader = OdsParser()
        self._rows = reader.get_rows(self.filePath, cellsPerRow)
        for row in self._rows:
            if len(row) != cellsPerRow:
                print(row)
                print(len(row), cellsPerRow)
                raise Error(f'{_("Wrong table structure")}.')

