"""Provide a class for csv character list import.

Copyright (c) 2022 Peter Triesberger
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import re
from pywriter.pywriter_globals import *
from pywriter.csv.csv_file import CsvFile


class CsvCharList(CsvFile):
    """csv file representation of a yWriter project's characters table. 
    
    Public methods:
        read() -- parse the file and get the instance variables.
    """
    DESCRIPTION = _('Character list')
    SUFFIX = '_charlist'
    _rowTitles = ['ID', 'Name', 'Full name', 'Aka', 'Description', 'Bio', 'Goals', 'Importance', 'Tags', 'Notes']

    def read(self):
        """Parse the file and get the instance variables.
        
        Parse the csv file located at filePath, fetching the Character attributes contained.
        Raise the "Error" exception in case of error. 
        Extends the superclass method.
        """
        super().read()
        for cells in self._rows:
            if 'CrID:' in cells[0]:
                crId = re.search('CrID\:([0-9]+)', cells[0]).group(1)
                self.srtCharacters.append(crId)
                self.characters[crId] = self.CHARACTER_CLASS()
                self.characters[crId].title = cells[1]
                self.characters[crId].fullName = cells[2]
                self.characters[crId].aka = cells[3]
                self.characters[crId].desc = self._convert_to_yw(cells[4])
                self.characters[crId].bio = self._convert_to_yw(cells[5])
                self.characters[crId].goals = self._convert_to_yw(cells[6])
                if self.CHARACTER_CLASS.MAJOR_MARKER in cells[7]:
                    self.characters[crId].isMajor = True
                else:
                    self.characters[crId].isMajor = False
                self.characters[crId].tags = string_to_list(cells[8], divider=self._DIVIDER)
                self.characters[crId].notes = self._convert_to_yw(cells[9])
