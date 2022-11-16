"""Provide a class for csv item list import.

Copyright (c) 2022 Peter Triesberger
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import re
from pywriter.pywriter_globals import *
from pywriter.csv.csv_file import CsvFile


class CsvItemList(CsvFile):
    """csv file representation of a yWriter project's items table. 
    
    Public methods:
        read() -- parse the file and get the instance variables.
    """
    DESCRIPTION = _('Item list')
    SUFFIX = '_itemlist'
    _rowTitles = ['ID', 'Name', 'Description', 'Aka', 'Tags']

    def read(self):
        """Parse the file and get the instance variables.
        
        Parse the csv file located at filePath, fetching the item attributes contained.
        Raise the "Error" exception in case of error. 
        Extends the superclass method.
        """
        super().read()
        for cells in self._rows:
            if 'ItID:' in cells[0]:
                itId = re.search('ItID\:([0-9]+)', cells[0]).group(1)
                self.srtItems.append(itId)
                self.items[itId] = self.WE_CLASS()
                self.items[itId].title = self._convert_to_yw(cells[1])
                self.items[itId].desc = self._convert_to_yw(cells[2])
                self.items[itId].aka = self._convert_to_yw(cells[3])
                self.items[itId].tags = string_to_list(cells[4], divider=self._DIVIDER)
