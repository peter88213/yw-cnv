"""Provide a class for ODT 'work in progress' import.

Conventions:
A work in progress has no third level heading.

-   Heading 1 -- New chapter title (beginning a new section).
-   Heading 2 -- New chapter title.
-   * * * -- Scene divider (not needed for the first scene in a chapter).
-   Comments right at the scene beginning are considered scene titles.
-   All other text is considered scene content.

Copyright (c) 2023 Peter Triesberger
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
from pywriter.pywriter_globals import *
from pywriter.model.chapter import Chapter
from pywriter.model.scene import Scene
from pywriter.odt_r.odt_r_formatted import OdtRFormatted


class OdtRImport(OdtRFormatted):
    """ODT 'work in progress' file reader.

    Public methods:
        handle_comment -- Process inline comments within scene content.
        handle_data -- Collect data within scene sections.
        handle_endtag -- Recognize the paragraph's end.
        handle_starttag -- Recognize the paragraph's beginning.
        read() -- Parse the file and get the instance variables.

    Import untagged chapters and scenes.
    """
    DESCRIPTION = _('Work in progress')
    SUFFIX = ''
    _SCENE_DIVIDER = '* * *'
    _LOW_WORDCOUNT = 10

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

    def handle_comment(self, data):
        """Process inline comments within scene content.
        
        Positional arguments:
            data: str -- comment text. 
        
        Use marked comments at scene start as scene titles.
        Overrides the superclass method.
        """
        if self._scId is not None:
            if not self._lines:
                # Comment is at scene start
                try:
                    self.novel.scenes[self._scId].title = data.strip()
                except:
                    pass
                return

            self._lines.append(f'{self._COMMENT_START}{data.strip()}{self._COMMENT_END}')

    def handle_data(self, data):
        """Collect data within scene sections.

        Positional arguments:
            data: str -- text to be stored. 
        
        Overrides the superclass method.
        """
        if self._scId is not None and self._SCENE_DIVIDER in data:
            self._scId = None
        else:
            self._lines.append(data)

    def handle_endtag(self, tag):
        """Recognize the paragraph's end.
        
        Positional arguments:
            tag: str -- name of the tag converted to lower case.

        Overrides the superclass method.
        """
        if tag in ('p', 'blockquote'):
            if self._language:
                self._lines.append(f'[/lang={self._language}]')
                self._language = ''
            self._lines.append('\n')
            if self._scId is not None:
                sceneText = ''.join(self._lines).rstrip()
                sceneText = self._cleanup_scene(sceneText)
                self.novel.scenes[self._scId].sceneContent = sceneText
                if self.novel.scenes[self._scId].wordCount < self._LOW_WORDCOUNT:
                    self.novel.scenes[self._scId].status = Scene.STATUS.index('Outline')
                else:
                    self.novel.scenes[self._scId].status = Scene.STATUS.index('Draft')
        elif tag == 'em':
            self._lines.append('[/i]')
        elif tag == 'strong':
            self._lines.append('[/b]')
        elif tag == 'lang':
            if self._language:
                self._lines.append(f'[/lang={self._language}]')
                self._language = ''
        elif tag in ('h1', 'h2'):
            self.novel.chapters[self._chId].title = ''.join(self._lines)
            self._lines = []
        elif tag == 'title':
            self.novel.title = ''.join(self._lines)

    def handle_starttag(self, tag, attrs):
        """Recognize the paragraph's beginning.
        
        Positional arguments:
            tag: str -- name of the tag converted to lower case.
            attrs -- list of (name, value) pairs containing the attributes found inside the tag’s <> brackets.
        
        Overrides the superclass method.
        """
        if tag == 'p':
            if self._scId is None and self._chId is not None:
                self._lines = []
                self._scCount += 1
                self._scId = str(self._scCount)
                self.novel.scenes[self._scId] = Scene()
                self.novel.chapters[self._chId].srtScenes.append(self._scId)
                self.novel.scenes[self._scId].status = '1'
                self.novel.scenes[self._scId].title = f'{_("Scene")} {self._scCount}'
            try:
                if attrs[0][0] == 'lang':
                    self._language = attrs[0][1]
                    if not self._language in self.novel.languages:
                        self.novel.languages.append(self._language)
                    self._lines.append(f'[lang={self._language}]')
            except:
                pass
        elif tag == 'em':
            self._lines.append('[i]')
        elif tag == 'strong':
            self._lines.append('[b]')
        elif tag == 'lang':
            if attrs[0][0] == 'lang':
                self._language = attrs[0][1]
                if not self._language in self.novel.languages:
                    self.novel.languages.append(self._language)
                self._lines.append(f'[lang={self._language}]')
        elif tag in ('h1', 'h2'):
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
        elif tag == 'li':
                self._lines.append(f'{self._BULLET} ')
        elif tag == 'blockquote':
            self._lines.append(f'{self._INDENT} ')
            try:
                if attrs[0][0] == 'lang':
                    self._language = attrs[0][1]
                    if not self._language in self.novel.languages:
                        self.novel.languages.append(self._language)
                    self._lines.append(f'[lang={self._language}]')
            except:
                pass
        elif tag == 's':
            self._lines.append(' ')

    def read(self):
        """Parse the file and get the instance variables.
        
        Initialize the languages list.
        Extends the superclass method.
        """
        self.novel.languages = []
        super().read()

