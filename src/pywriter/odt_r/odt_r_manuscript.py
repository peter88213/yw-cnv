"""Provide a class for ODT invisibly tagged chapters and scenes import.

Copyright (c) 2023 Peter Triesberger
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
from pywriter.pywriter_globals import *
from pywriter.odt_r.odt_r_formatted import OdtRFormatted
from pywriter.model.splitter import Splitter


class OdtRManuscript(OdtRFormatted):
    """ODT manuscript file reader.

    Public methods:
        handle_comment -- Process inline comments within scene content.
        handle_data -- Collect data within scene sections.
        handle_endtag -- Recognize the paragraph's end.
        handle_starttag -- Recognize the paragraph's beginning.

    Import a manuscript with invisibly tagged chapters and scenes.
    """
    DESCRIPTION = _('Editable manuscript')
    SUFFIX = '_manuscript'

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
                pass
            if self._SC_TITLE_BRACKET in data:
                # Comment is marked as a scene title
                try:
                    self.novel.scenes[self._scId].title = data.split(self._SC_TITLE_BRACKET)[1].strip()
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
        if self._skip_data:
            self._skip_data = False
        elif self._scId is not None:
            if not data.isspace():
                self._lines.append(data)
        elif self._chId is not None:
            if self.novel.chapters[self._chId].title is None:
                self.chapters[self._chId].title = data.strip()

    def handle_endtag(self, tag):
        """Recognize the end of the scene section and save data.
        
        Positional arguments:
            tag: str -- name of the tag converted to lower case.

        Overrides the superclass method.
        """
        if self._scId is not None:
            if tag in ('p', 'blockquote'):
                if self._language:
                    self._lines.append(f'[/lang={self._language}]')
                    self._language = ''
                self._lines.append('\n')
            elif tag == 'em':
                self._lines.append('[/i]')
            elif tag == 'strong':
                self._lines.append('[/b]')
            elif tag == 'lang':
                if self._language:
                    self._lines.append(f'[/lang={self._language}]')
                    self._language = ''
            elif tag == 'div':
                text = ''.join(self._lines)
                self.novel.scenes[self._scId].sceneContent = self._cleanup_scene(text).rstrip()
                self._lines = []
                self._scId = None
            elif tag == 'h1':
                self._lines.append('\n')
            elif tag == 'h2':
                self._lines.append('\n')
        elif self._chId is not None:
            if tag == 'div':
                self._chId = None

    def handle_starttag(self, tag, attrs):
        """Identify scenes and chapters.
        
        Positional arguments:
            tag: str -- name of the tag converted to lower case.
            attrs -- list of (name, value) pairs containing the attributes found inside the tag’s <> brackets.
        
        Extends the superclass method by processing inline chapter and scene dividers.
        """
        super().handle_starttag(tag, attrs)
        if self._scId is not None:
            if tag == 'em':
                self._lines.append('[i]')
            elif tag == 'strong':
                self._lines.append('[b]')
            elif tag in ('lang', 'p'):
                try:
                    if attrs[0][0] == 'lang':
                        self._language = attrs[0][1]
                        if not self._language in self.novel.languages:
                            self.novel.languages.append(self._language)
                        self._lines.append(f'[lang={self._language}]')
                except:
                    pass
            elif tag == 'h3':
                self._skip_data = True
                # this is for downward compatibility with "notes" and "todo"
                # documents generated with PyWriter v8 and before.
            elif tag == 'h2':
                self._lines.append(f'{Splitter.CHAPTER_SEPARATOR} ')
            elif tag == 'h1':
                self._lines.append(f'{Splitter.PART_SEPARATOR} ')
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

