"""Provide a class for csv scene list import.

Copyright (c) 2022 Peter Triesberger
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import re
from pywriter.pywriter_globals import *
from pywriter.csv.csv_file import CsvFile


class CsvSceneList(CsvFile):
    """csv file representation of a yWriter project's scenes table. 
    
    Public methods:
        read() -- parse the file and get the instance variables.
    """
    DESCRIPTION = _('Scene list')
    SUFFIX = '_scenelist'
    _SCENE_RATINGS = ['2', '3', '4', '5', '6', '7', '8', '9', '10']
    # '1' is assigned N/A (empty table cell).
    _rowTitles = ['Scene link', 'Scene title', 'Scene description', 'Tags', 'Scene notes', 'A/R',
                 'Goal', 'Conflict', 'Outcome', 'Scene', 'Words total',
                 '$FieldTitle1', '$FieldTitle2', '$FieldTitle3', '$FieldTitle4',
                 'Word count', 'Letter count', 'Status', 'Characters', 'Locations', 'Items']

    def read(self):
        """Parse the file and get the instance variables.
        
        Parse the csv file located at filePath, fetching the Scene attributes contained.
        Extends the superclass method.
        """
        super().read()
        for cells in self._rows:
            i = 0
            if 'ScID:' in cells[i]:
                scId = re.search('ScID\:([0-9]+)', cells[0]).group(1)
                self.scenes[scId] = self.SCENE_CLASS()
                i += 1
                self.scenes[scId].title = self._convert_to_yw(cells[i])
                i += 1
                self.scenes[scId].desc = self._convert_to_yw(cells[i])
                i += 1
                self.scenes[scId].tags = string_to_list(cells[i], divider=self._DIVIDER)
                i += 1
                self.scenes[scId].notes = self._convert_to_yw(cells[i])
                i += 1
                if self.SCENE_CLASS.REACTION_MARKER.lower() in cells[i].lower():
                    self.scenes[scId].isReactionScene = True
                else:
                    self.scenes[scId].isReactionScene = False
                i += 1
                self.scenes[scId].goal = self._convert_to_yw(cells[i])
                i += 1
                self.scenes[scId].conflict = self._convert_to_yw(cells[i])
                i += 1
                self.scenes[scId].outcome = self._convert_to_yw(cells[i])
                i += 1
                # Don't write back sceneCount
                i += 1
                # Don't write back wordCount
                i += 1

                # Transfer scene ratings; set to 1 if deleted
                if cells[i] in self._SCENE_RATINGS:
                    self.scenes[scId].field1 = self._convert_to_yw(cells[i])
                else:
                    self.scenes[scId].field1 = '1'
                i += 1
                if cells[i] in self._SCENE_RATINGS:
                    self.scenes[scId].field2 = self._convert_to_yw(cells[i])
                else:
                    self.scenes[scId].field2 = '1'
                i += 1
                if cells[i] in self._SCENE_RATINGS:
                    self.scenes[scId].field3 = self._convert_to_yw(cells[i])
                else:
                    self.scenes[scId].field3 = '1'
                i += 1
                if cells[i] in self._SCENE_RATINGS:
                    self.scenes[scId].field4 = self._convert_to_yw(cells[i])
                else:
                    self.scenes[scId].field4 = '1'
                i += 1
                # Don't write back scene words total
                i += 1
                # Don't write back scene letters total
                i += 1
                try:
                    self.scenes[scId].status = self.SCENE_CLASS.STATUS.index(cells[i])
                except ValueError:
                    pass
                    # Scene status remains None and will be ignored when
                    # writing back.
                i += 1
                # Can't write back character IDs, because self.characters is None.
                i += 1
                # Can't write back location IDs, because self.locations is None.
                i += 1
                # Can't write back item IDs, because self.items is None.
