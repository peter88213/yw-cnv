"""Provide a generic class for html file import.

Other html file representations inherit from this class.

Copyright (c) 2022 Peter Triesberger
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import re
from html.parser import HTMLParser
from pywriter.pywriter_globals import *
from pywriter.model.novel import Novel
from pywriter.html.html_fop import read_html_file


class HtmlFile(Novel, HTMLParser):
    """Generic HTML file representation.
    
    Public methods:
        handle_starttag -- identify scenes and chapters.
        handle comment --
        read --
    """
    EXTENSION = '.html'
    _COMMENT_START = '/*'
    _COMMENT_END = '*/'
    _SC_TITLE_BRACKET = '~'
    _BULLET = '-'
    _INDENT = '>'

    def __init__(self, filePath, **kwargs):
        """Initialize the HTML parser and local instance variables for parsing.
        
        Positional arguments:
            filePath -- str: path to the file represented by the Novel instance.
            
        Optional arguments:
            kwargs -- keyword arguments to be used by subclasses.            

        The HTML parser works like a state machine. 
        Scene ID, chapter ID and processed lines must be saved between the transitions.         
        Extends the superclass constructor.
        """
        super().__init__(filePath)
        HTMLParser.__init__(self)
        self._lines = []
        self._scId = None
        self._chId = None
        self._newline = False
        self._language = ''
        self._doNothing = False

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

    def _preprocess(self, text):
        """Process HTML text before parsing.
        
        Positional arguments:
            text -- str: HTML text to be processed.
        """
        return self._convert_to_yw(text)

    def _postprocess(self):
        """Process the plain text after parsing.
        
        This is a hook for subclasses.
        """

    def handle_starttag(self, tag, attrs):
        """Identify scenes and chapters.
        
        Positional arguments:
            tag -- str: name of the tag converted to lower case.
            attrs -- list of (name, value) pairs containing the attributes found inside the tag’s <> brackets.
        
        Overrides HTMLparser.handle_starttag() called by the parser to handle the start of a tag. 
        This method is applicable to HTML files that are divided into chapters and scenes. 
        For differently structured HTML files  do override this method in a subclass.
        """
        if tag == 'div':
            if attrs[0][0] == 'id':
                if attrs[0][1].startswith('ScID'):
                    self._scId = re.search('[0-9]+', attrs[0][1]).group()
                    self.scenes[self._scId] = self.SCENE_CLASS()
                    self.chapters[self._chId].srtScenes.append(self._scId)
                elif attrs[0][1].startswith('ChID'):
                    self._chId = re.search('[0-9]+', attrs[0][1]).group()
                    self.chapters[self._chId] = self.CHAPTER_CLASS()
                    self.chapters[self._chId].srtScenes = []
                    self.srtChapters.append(self._chId)

    def handle_comment(self, data):
        """Process inline comments within scene content.
        
        Positional arguments:
            data -- str: comment text. 
        
        Overrides HTMLparser.handle_comment() called by the parser when a comment is encountered.
        """
        if self._scId is not None:
            self._lines.append(f'{self._COMMENT_START}{data}{self._COMMENT_END}')

    def read(self):
        """Parse the file and get the instance variables.
        
        This is a template method for subclasses tailored to the 
        content of the respective HTML file.
        """
        content = read_html_file(self._filePath)
        content = self._preprocess(content)
        self.feed(content)
        self._postprocess()
