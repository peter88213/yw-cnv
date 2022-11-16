"""Provide a class for html 'work in progress' import.

Conventions:
A work in progress has no third level heading.

-   Heading 1 -- New chapter title (beginning a new section).
-   Heading 2 -- New chapter title.
-   * * * -- Scene divider (not needed for the first scene in a chapter).
-   Comments right at the scene beginning are considered scene titles.
-   All other text is considered scene content.

Copyright (c) 2022 Peter Triesberger
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
from pywriter.pywriter_globals import *
from pywriter.html.html_formatted import HtmlFormatted


class HtmlImport(HtmlFormatted):
    """HTML 'work in progress' file representation.

    Import untagged chapters and scenes.
    """
    DESCRIPTION = _('Work in progress')
    SUFFIX = ''
    _SCENE_DIVIDER = '* * *'
    _LOW_WORDCOUNT = 10

    def __init__(self, filePath, **kwargs):
        """Initialize local instance variables for parsing.

        Positional arguments:
            filePath -- str: path to the file represented by the Novel instance.
            
        The HTML parser works like a state machine. 
        Chapter and scene count must be saved between the transitions.         
        Extends the superclass constructor.
        """
        super().__init__(filePath)
        self._chCount = 0
        self._scCount = 0

    def handle_starttag(self, tag, attrs):
        """Recognize the paragraph's beginning.
        
        Positional arguments:
            tag -- str: name of the tag converted to lower case.
            attrs -- list of (name, value) pairs containing the attributes found inside the tag’s <> brackets.
        
        Overrides the superclass method.
        """
        if tag == 'p':
            if self._scId is None and self._chId is not None:
                self._lines = []
                self._scCount += 1
                self._scId = str(self._scCount)
                self.scenes[self._scId] = self.SCENE_CLASS()
                self.chapters[self._chId].srtScenes.append(self._scId)
                self.scenes[self._scId].status = '1'
                self.scenes[self._scId].title = f'Scene {self._scCount}'
            try:
                if attrs[0][0] == 'lang':
                    self._language = attrs[0][1]
                    if not self._language in self.languages:
                        self.languages.append(self._language)
                    self._lines.append(f'[lang={self._language}]')
            except:
                pass
        elif tag == 'br':
            self._newline = True
        elif tag == 'em' or tag == 'i':
            self._lines.append('[i]')
        elif tag == 'strong' or tag == 'b':
            self._lines.append('[b]')
        elif tag == 'span':
            if attrs[0][0] == 'lang':
                self._language = attrs[0][1]
                if not self._language in self.languages:
                    self.languages.append(self._language)
                self._lines.append(f'[lang={self._language}]')
        elif tag in ('h1', 'h2'):
            self._scId = None
            self._lines = []
            self._chCount += 1
            self._chId = str(self._chCount)
            self.chapters[self._chId] = self.CHAPTER_CLASS()
            self.chapters[self._chId].srtScenes = []
            self.srtChapters.append(self._chId)
            self.chapters[self._chId].chType = 0
            if tag == 'h1':
                self.chapters[self._chId].chLevel = 1
            else:
                self.chapters[self._chId].chLevel = 0
        elif tag == 'div':
            self._scId = None
            self._chId = None
        elif tag == 'meta':
            if attrs[0][1].lower() == 'author':
                self.authorName = attrs[1][1]
            if attrs[0][1].lower() == 'description':
                self.desc = attrs[1][1]
        elif tag == 'title':
            self._lines = []
        elif tag == 'body':
            for attr in attrs:
                if attr[0] == 'lang':
                    try:
                        lngCode, ctrCode = attr[1].split('-')
                        self.languageCode = lngCode
                        self.countryCode = ctrCode
                    except:
                        pass
                    break
        elif tag == 'li':
                self._lines.append(f'{self._BULLET} ')
        elif tag == 'ul':
                self._doNothing = True
        elif tag == 'blockquote':
            self._lines.append(f'{self._INDENT} ')
            try:
                if attrs[0][0] == 'lang':
                    self._language = attrs[0][1]
                    if not self._language in self.languages:
                        self.languages.append(self._language)
                    self._lines.append(f'[lang={self._language}]')
            except:
                pass

    def handle_endtag(self, tag):
        """Recognize the paragraph's end.
        
        Positional arguments:
            tag -- str: name of the tag converted to lower case.

        Overrides HTMLparser.handle_endtag() called by the HTML parser to handle the end tag of an element.
        """
        if tag in ('p', 'blockquote'):
            if self._language:
                self._lines.append(f'[/lang={self._language}]')
                self._language = ''
            self._lines.append('\n')
            self._newline = True
            if self._scId is not None:
                sceneText = ''.join(self._lines).rstrip()
                sceneText = self._cleanup_scene(sceneText)
                self.scenes[self._scId].sceneContent = sceneText
                if self.scenes[self._scId].wordCount < self._LOW_WORDCOUNT:
                    self.scenes[self._scId].status = self.SCENE_CLASS.STATUS.index('Outline')
                else:
                    self.scenes[self._scId].status = self.SCENE_CLASS.STATUS.index('Draft')
        elif tag == 'em' or tag == 'i':
            self._lines.append('[/i]')
        elif tag == 'strong' or tag == 'b':
            self._lines.append('[/b]')
        elif tag == 'span':
            if self._language:
                self._lines.append(f'[/lang={self._language}]')
                self._language = ''
        elif tag in ('h1', 'h2'):
            self.chapters[self._chId].title = ''.join(self._lines)
            self._lines = []
        elif tag == 'title':
            self.title = ''.join(self._lines)

    def handle_comment(self, data):
        """Process inline comments within scene content.
        
        Positional arguments:
            data -- str: comment text. 
        
        Use marked comments at scene start as scene titles.
        Overrides HTMLparser.handle_comment() called by the parser when a comment is encountered.
        """
        if self._scId is not None:
            if not self._lines:
                # Comment is at scene start
                try:
                    self.scenes[self._scId].title = data.strip()
                except:
                    pass
                return

            self._lines.append(f'{self._COMMENT_START}{data.strip()}{self._COMMENT_END}')

    def handle_data(self, data):
        """Collect data within scene sections.

        Positional arguments:
            data -- str: text to be stored. 
        
        Overrides HTMLparser.handle_data() called by the parser to process arbitrary data.
        """
        if self._doNothing:
            self._doNothing = False
        elif self._scId is not None and self._SCENE_DIVIDER in data:
            self._scId = None
        else:
            if self._newline:
                data = data.rstrip()
                self._newline = False
            self._lines.append(data)
