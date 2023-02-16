"""Provide a class for ODS location list import.

Copyright (c) 2023 Peter Triesberger
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import re
from pywriter.pywriter_globals import *
from pywriter.model.world_element import WorldElement
from pywriter.ods_r.ods_reader import OdsReader


class OdsRLocList(OdsReader):
    """ODS location list reader. 
    
    Public methods:
        read() -- parse the file and get the instance variables.
    """
    DESCRIPTION = _('Location list')
    SUFFIX = '_loclist'
    _rowTitles = ['ID', 'Name', 'Description', 'Aka', 'Tags']

    def read(self):
        """Parse the file and get the instance variables.
        
        Parse the csv file located at filePath, fetching the location attributes contained.
        Raise the "Error" exception in case of error. 
        Extends the superclass method.
        """
        super().read()
        self.novel.srtLocations = []
        for cells in self._rows:
            if 'LcID:' in cells[0]:
                lcId = re.search('LcID\:([0-9]+)', cells[0]).group(1)
                self.novel.srtLocations.append(lcId)
                if not lcId in self.novel.locations:
                    self.novel.locations[lcId] = WorldElement()
                if self.novel.locations[lcId].title or cells[1]:
                    self.novel.locations[lcId].title = self._convert_to_yw(cells[1])
                if self.novel.locations[lcId].desc or cells[2]:
                    self.novel.locations[lcId].desc = self._convert_to_yw(cells[2])
                if self.novel.locations[lcId].aka or cells[3]:
                    self.novel.locations[lcId].aka = self._convert_to_yw(cells[3])
                if self.novel.locations[lcId].tags or cells[4]:
                    self.novel.locations[lcId].tags = string_to_list(cells[4], divider=self._DIVIDER)
