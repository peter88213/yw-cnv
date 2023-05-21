"""Provide a generic class for ODT file import.

Other ODT file readers inherit from this class.

Copyright (c) 2023 Peter Triesberger
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import re
from pywriter.pywriter_globals import *
from pywriter.model.chapter import Chapter
from pywriter.model.scene import Scene
from pywriter.file.file import File
from pywriter.odt_r.odt_parser import OdtParser


class OdtReader(File, OdtParser):
    """Generic ODT file reader.
    
    Public methods:
        handle comment(data) -- Process inline comments within scene content.
        handle_starttag(tag, attrs) -- Identify scenes and chapters.
        read() -- Parse the file and get the instance variables.
    """
    EXTENSION = '.odt'

    _TYPE = 0

    _COMMENT_START = '/*'
    _COMMENT_END = '*/'
    _SC_TITLE_BRACKET = '~'
    _BULLET = '-'
    _INDENT = '>'

    def __init__(self, filePath, **kwargs):
        """Initialize the ODT parser and local instance variables for parsing.
        
        Positional arguments:
            filePath: str -- path to the file represented by the File instance.
            
        Optional arguments:
            kwargs -- keyword arguments to be used by subclasses.            

        The ODT parser works like a state machine. 
        Scene ID, chapter ID and processed lines must be saved between the transitions.         
        Extends the superclass constructor.
        """
        super().__init__(filePath)
        self._lines = []
        self._scId = None
        self._chId = None
        self._newline = False
        self._language = ''
        self._skip_data = False

    def handle_comment(self, data):
        """Process inline comments within scene content.
        
        Positional arguments:
            data: str -- comment text. 

        Overrides the superclass method.
        """
        if self._scId is not None:
            self._lines.append(f'{self._COMMENT_START}{data}{self._COMMENT_END}')

    def handle_starttag(self, tag, attrs):
        """Identify scenes and chapters.
        
        Positional arguments:
            tag: str -- name of the tag.
            attrs -- list of (name, value) pairs containing the attributes found inside the tag’s <> brackets.
        
        Overrides the superclass method. 
        This method is applicable to ODT files that are divided into chapters and scenes. 
        For differently structured ODT files  do override this method in a subclass.
        """
        if tag == 'div':
            if attrs[0][0] == 'id':
                if attrs[0][1].startswith('ScID'):
                    self._scId = re.search('[0-9]+', attrs[0][1]).group()
                    if not self._scId in self.novel.scenes:
                        self.novel.scenes[self._scId] = Scene()
                        self.novel.chapters[self._chId].srtScenes.append(self._scId)
                    self.novel.scenes[self._scId].scType = self._TYPE
                elif attrs[0][1].startswith('ChID'):
                    self._chId = re.search('[0-9]+', attrs[0][1]).group()
                    if not self._chId in self.novel.chapters:
                        self.novel.chapters[self._chId] = Chapter()
                        self.novel.chapters[self._chId].srtScenes = []
                        self.novel.srtChapters.append(self._chId)
                    self.novel.chapters[self._chId].chType = self._TYPE
        elif tag == 's':
            self._lines.append(' ')

    def read(self):
        OdtParser.feed_file(self, self.filePath)

    def _convert_to_yw(self, text):
        """Convert html formatting tags to yWriter 7 raw markup.
        
        Positional arguments:
            text -- string to convert.
        
        Return a yw7 markup string.
        Overrides the superclass method.
        """
        #--- Put everything in one line.
        text = text.replace('\n', ' ')
        text = text.replace('\r', ' ')
        text = text.replace('\t', ' ')
        while '  ' in text:
            text = text.replace('  ', ' ')

        return text

