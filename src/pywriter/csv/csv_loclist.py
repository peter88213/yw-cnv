"""Provide a class for csv location list import.

Copyright (c) 2022 Peter Triesberger
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import re
from pywriter.pywriter_globals import *
from pywriter.csv.csv_file import CsvFile


class CsvLocList(CsvFile):
    """csv file representation of a yWriter project's locations table. 
    
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
        for cells in self._rows:
            if 'LcID:' in cells[0]:
                lcId = re.search('LcID\:([0-9]+)', cells[0]).group(1)
                self.srtLocations.append(lcId)
                self.locations[lcId] = self.WE_CLASS()
                self.locations[lcId].title = self._convert_to_yw(cells[1])
                self.locations[lcId].desc = self._convert_to_yw(cells[2])
                self.locations[lcId].aka = self._convert_to_yw(cells[3])
                self.locations[lcId].tags = string_to_list(cells[4], divider=self._DIVIDER)
