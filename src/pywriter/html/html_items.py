"""Provide a class for html item invisibly tagged descriptions import.

Copyright (c) 2022 Peter Triesberger
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import re
from pywriter.pywriter_globals import *
from pywriter.html.html_file import HtmlFile


class HtmlItems(HtmlFile):
    """HTML item descriptions file representation.

    Import a item sheet with invisibly tagged descriptions.
    """
    DESCRIPTION = _('Item descriptions')
    SUFFIX = '_items'

    def __init__(self, filePath, **kwargs):
        """Initialize local instance variables for parsing.

        Positional arguments:
            filePath -- str: path to the file represented by the Novel instance.
            
        The HTML parser works like a state machine. 
        The item ID must be saved between the transitions.         
        Extends the superclass constructor.
        """
        super().__init__(filePath)
        self._itId = None

    def handle_starttag(self, tag, attrs):
        """Identify items.
        
        Positional arguments:
            tag -- str: name of the tag converted to lower case.
            attrs -- list of (name, value) pairs containing the attributes found inside the tag’s <> brackets.
        
        Overrides the superclass method.
        """
        if tag == 'div':
            if attrs[0][0] == 'id':
                if attrs[0][1].startswith('ItID'):
                    self._itId = re.search('[0-9]+', attrs[0][1]).group()
                    self.srtItems.append(self._itId)
                    self.items[self._itId] = self.WE_CLASS()

    def handle_endtag(self, tag):
        """Recognize the end of the item section and save data.
        
        Positional arguments:
            tag -- str: name of the tag converted to lower case.

        Overrides HTMLparser.handle_endtag() called by the HTML parser to handle the end tag of an element.
        """
        if self._itId is not None:
            if tag == 'div':
                self.items[self._itId].desc = ''.join(self._lines).rstrip()
                self._lines = []
                self._itId = None
            elif tag == 'p':
                self._lines.append('\n')

    def handle_data(self, data):
        """collect data within item sections.

        Positional arguments:
            data -- str: text to be stored. 
        
        Overrides HTMLparser.handle_data() called by the parser to process arbitrary data.
        """
        if self._itId is not None:
            self._lines.append(data.strip())
