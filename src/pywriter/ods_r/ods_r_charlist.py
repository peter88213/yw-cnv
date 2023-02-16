"""Provide a class for ODS character list import.

Copyright (c) 2023 Peter Triesberger
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import re
from pywriter.pywriter_globals import *
from pywriter.model.character import Character
from pywriter.ods_r.ods_reader import OdsReader


class OdsRCharList(OdsReader):
    """ODS character list reader. 
    
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
        self.novel.srtCharacters = []
        for cells in self._rows:
            if 'CrID:' in cells[0]:
                crId = re.search('CrID\:([0-9]+)', cells[0]).group(1)
                self.novel.srtCharacters.append(crId)
                if not crId in self.novel.characters:
                    self.novel.characters[crId] = Character()
                if self.novel.characters[crId].title or cells[1]:
                    self.novel.characters[crId].title = cells[1]
                if self.novel.characters[crId].fullName or cells[2]:
                    self.novel.characters[crId].fullName = cells[2]
                if self.novel.characters[crId].aka or cells[3]:
                    self.novel.characters[crId].aka = cells[3]
                if self.novel.characters[crId].desc or cells[4]:
                    self.novel.characters[crId].desc = self._convert_to_yw(cells[4])
                if self.novel.characters[crId].bio or cells[5]:
                    self.novel.characters[crId].bio = self._convert_to_yw(cells[5])
                if self.novel.characters[crId].goals  or cells[6]:
                    self.novel.characters[crId].goals = self._convert_to_yw(cells[6])
                if Character.MAJOR_MARKER in cells[7]:
                    self.novel.characters[crId].isMajor = True
                else:
                    self.novel.characters[crId].isMajor = False
                if self.novel.characters[crId].tags or cells[8]:
                    self.novel.characters[crId].tags = string_to_list(cells[8], divider=self._DIVIDER)
                if self.novel.characters[crId].notes or cells[9]:
                    self.novel.characters[crId].notes = self._convert_to_yw(cells[9])
