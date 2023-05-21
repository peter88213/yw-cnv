"""Provide a class for ODT outline import.

Conventions:
An outline has at least one third level heading.

-   Heading 1 -- New chapter title (beginning a new section).
-   Heading 2 -- New chapter title.
-   Heading 3 -- New scene title.
-   All other text is considered to be chapter/scene description.

Copyright (c) 2023 Peter Triesberger
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
from pywriter.pywriter_globals import *
from pywriter.model.chapter import Chapter
from pywriter.model.scene import Scene
from pywriter.odt_r.odt_reader import OdtReader


class OdtROutline(OdtReader):
    """ODT outline file reader.

    Public methods:
        handle_data -- Collect data within scene sections.
        handle_endtag -- Recognize the paragraph's end.
        handle_starttag -- Recognize the paragraph's beginning.
    
    Import an outline without chapter and scene tags.
    """
    DESCRIPTION = _('Novel outline')
    SUFFIX = ''

    def __init__(self, filePath, **kwargs):
        """Initialize local instance variables for parsing.

        Positional arguments:
            filePath: str -- path to the file represented by the Novel instance.
            
        The ODT parser works like a state machine. 
        Chapter and scene count must be saved between the transitions.         
        Extends the superclass constructor.
        """
        super().__init__(filePath)
        self._chCount = 0
        self._scCount = 0

    def handle_data(self, data):
        """Collect data within scene sections.

        Positional arguments:
            data: str -- text to be stored. 
        
        Overrides the superclass method.
        """
        self._lines.append(data)

    def handle_endtag(self, tag):
        """Recognize the paragraph's end.
        
        Positional arguments:
            tag: str -- name of the tag converted to lower case.

        Overrides the superclass method.
        """
        text = ''.join(self._lines)
        if tag == 'p':
            text = f'{text.strip()}\n'
            self._lines = [text]
            if self._scId is not None:
                self.novel.scenes[self._scId].desc = text
            elif self._chId is not None:
                self.novel.chapters[self._chId].desc = text
        elif tag in ('h1', 'h2'):
            self.novel.chapters[self._chId].title = text.strip()
            self._lines = []
        elif tag == 'h3':
            self.novel.scenes[self._scId].title = text.strip()
            self._lines = []
        elif tag == 'title':
            self.novel.title = text.strip()

    def handle_starttag(self, tag, attrs):
        """Recognize the paragraph's beginning.
        
        Positional arguments:
            tag: str -- name of the tag converted to lower case.
            attrs -- list of (name, value) pairs containing the attributes found inside the tag’s <> brackets.
        
        Overrides the superclass method.
        """
        if tag in ('h1', 'h2'):
            self._scId = None
            self._lines = []
            self._chCount += 1
            self._chId = str(self._chCount)
            self.novel.chapters[self._chId] = Chapter()
            self.novel.chapters[self._chId].srtScenes = []
            self.novel.srtChapters.append(self._chId)
            self.novel.chapters[self._chId].chType = 0
            if tag == 'h1':
                self.novel.chapters[self._chId].chLevel = 1
            else:
                self.novel.chapters[self._chId].chLevel = 0
        elif tag == 'h3':
            self._lines = []
            self._scCount += 1
            self._scId = str(self._scCount)
            self.novel.scenes[self._scId] = Scene()
            self.novel.chapters[self._chId].srtScenes.append(self._scId)
            self.novel.scenes[self._scId].sceneContent = ''
            self.novel.scenes[self._scId].status = Scene.STATUS.index('Outline')
        elif tag == 'div':
            self._scId = None
            self._chId = None
        elif tag == 'meta':
            if attrs[0][1] == 'author':
                self.novel.authorName = attrs[1][1]
            if attrs[0][1] == 'description':
                self.novel.desc = attrs[1][1]
        elif tag == 'title':
            self._lines = []
        elif tag == 'body':
            for attr in attrs:
                if attr[0] == 'language':
                    if attr[1]:
                        self.novel.languageCode = attr[1]
                elif attr[0] == 'country':
                    if attr[1]:
                        self.novel.countryCode = attr[1]
        elif tag == 's':
            self._lines.append(' ')
