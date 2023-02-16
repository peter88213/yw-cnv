"""Provide a class for ODT invisibly tagged location descriptions import.

Copyright (c) 2023 Peter Triesberger
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import re
from pywriter.pywriter_globals import *
from pywriter.model.world_element import WorldElement
from pywriter.odt_r.odt_reader import OdtReader


class OdtRLocations(OdtReader):
    """ODT location descriptions file reader.

    Import a location sheet with invisibly tagged descriptions.
    """
    DESCRIPTION = _('Location descriptions')
    SUFFIX = '_locations'

    def __init__(self, filePath, **kwargs):
        """Initialize local instance variables for parsing.

        Positional arguments:
            filePath -- str: path to the file represented by the Novel instance.
            
        The ODT parser works like a state machine. 
        The location ID must be saved between the transitions.         
        Extends the superclass constructor.
        """
        super().__init__(filePath)
        self._lcId = None

    def handle_starttag(self, tag, attrs):
        """Identify locations.
        
        Positional arguments:
            tag -- str: name of the tag converted to lower case.
            attrs -- list of (name, value) pairs containing the attributes found inside the tag’s <> brackets.
        
        Overrides the superclass method.
        """
        if tag == 'div':
            if attrs[0][0] == 'id':
                if attrs[0][1].startswith('LcID'):
                    self._lcId = re.search('[0-9]+', attrs[0][1]).group()
                    if not self._lcId in self.novel.locations:
                        self.novel.srtLocations.append(self._lcId)
                        self.novel.locations[self._lcId] = WorldElement()

    def handle_endtag(self, tag):
        """Recognize the end of the location section and save data.
        
        Positional arguments:
            tag -- str: name of the tag converted to lower case.

        Overrides the superclass method.
        """
        if self._lcId is not None:
            if tag == 'div':
                self.novel.locations[self._lcId].desc = ''.join(self._lines).rstrip()
                self._lines = []
                self._lcId = None
            elif tag == 'p':
                self._lines.append('\n')

    def handle_data(self, data):
        """collect data within location sections.
        
        Positional arguments:
            data -- str: text to be stored. 
        
        Overrides the superclass method.
        """
        if self._lcId is not None:
            self._lines.append(data.strip())
