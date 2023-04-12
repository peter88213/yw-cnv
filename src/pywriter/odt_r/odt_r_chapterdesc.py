"""Provide a class for html invisibly tagged chapter descriptions import.

Copyright (c) 2023 Peter Triesberger
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
from pywriter.pywriter_globals import *
from pywriter.odt_r.odt_reader import OdtReader


class OdtRChapterDesc(OdtReader):
    """ODT chapter summaries file reader.

    Public methods:
        handle_data -- Collect data within scene sections.
        handle_endtag -- Recognize the paragraph's end.

    Import a brief synopsis with invisibly tagged chapter descriptions.
    """
    DESCRIPTION = _('Chapter descriptions')
    SUFFIX = '_chapters'

    def handle_endtag(self, tag):
        """Recognize the end of the chapter section and save data.
        
        Positional arguments:
            tag: str -- name of the tag converted to lower case.

        Overrides the superclass method.
        """
        if self._chId is not None:
            if tag == 'div':
                self.novel.chapters[self._chId].desc = ''.join(self._lines).rstrip()
                self._lines = []
                self._chId = None
            elif tag == 'p':
                self._lines.append('\n')
            elif tag == 'h1' or tag == 'h2':
                if not self.novel.chapters[self._chId].title:
                    self.novel.chapters[self._chId].title = ''.join(self._lines)
                self._lines = []

    def handle_data(self, data):
        """Collect data within chapter sections.

        Positional arguments:
            data: str -- text to be stored. 
        
        Overrides the superclass method.
        """
        if self._chId is not None:
            self._lines.append(data.strip())
