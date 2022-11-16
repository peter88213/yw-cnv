"""Provide a class for html invisibly tagged chapters and scenes import.

Copyright (c) 2022 Peter Triesberger
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
from pywriter.pywriter_globals import *
from pywriter.html.html_formatted import HtmlFormatted
from pywriter.model.splitter import Splitter


class HtmlManuscript(HtmlFormatted):
    """HTML manuscript file representation.

    Import a manuscript with invisibly tagged chapters and scenes.
    """
    DESCRIPTION = _('Editable manuscript')
    SUFFIX = '_manuscript'

    def handle_starttag(self, tag, attrs):
        """Identify scenes and chapters.
        
        Positional arguments:
            tag -- str: name of the tag converted to lower case.
            attrs -- list of (name, value) pairs containing the attributes found inside the tag’s <> brackets.
        
        Extends the superclass method by processing inline chapter and scene dividers.
        """
        super().handle_starttag(tag, attrs)
        if self._scId is not None:
            self._getScTitle = False
            if tag == 'em' or tag == 'i':
                self._lines.append('[i]')
            elif tag == 'strong' or tag == 'b':
                self._lines.append('[b]')
            elif tag in ('span', 'p'):
                try:
                    if attrs[0][0] == 'lang':
                        self._language = attrs[0][1]
                        if not self._language in self.languages:
                            self.languages.append(self._language)
                        self._lines.append(f'[lang={self._language}]')
                except:
                    pass
            elif tag == 'h3':
                if self.scenes[self._scId].title is None:
                    self._getScTitle = True
                else:
                    self._lines.append(f'{Splitter.SCENE_SEPARATOR} ')
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
                        if not self._language in self.languages:
                            self.languages.append(self._language)
                        self._lines.append(f'[lang={self._language}]')
                except:
                    pass
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

    def handle_endtag(self, tag):
        """Recognize the end of the scene section and save data.
        
        Positional arguments:
            tag -- str: name of the tag converted to lower case.

        Overrides HTMLparser.handle_endtag() called by the HTML parser to handle the end tag of an element.
        """
        if self._scId is not None:
            if tag in ('p', 'blockquote'):
                if self._language:
                    self._lines.append(f'[/lang={self._language}]')
                    self._language = ''
                self._lines.append('\n')
            elif tag == 'em' or tag == 'i':
                self._lines.append('[/i]')
            elif tag == 'strong' or tag == 'b':
                self._lines.append('[/b]')
            elif tag == 'span':
                if self._language:
                    self._lines.append(f'[/lang={self._language}]')
                    self._language = ''
            elif tag == 'div':
                text = ''.join(self._lines)
                self.scenes[self._scId].sceneContent = self._cleanup_scene(text).rstrip()
                self._lines = []
                self._scId = None
            elif tag == 'h1':
                self._lines.append('\n')
            elif tag == 'h2':
                self._lines.append('\n')
            elif tag == 'h3' and not self._getScTitle:
                self._lines.append('\n')
        elif self._chId is not None:
            if tag == 'div':
                self._chId = None

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
                pass
            if self._SC_TITLE_BRACKET in data:
                # Comment is marked as a scene title
                try:
                    self.scenes[self._scId].title = data.split(self._SC_TITLE_BRACKET)[1].strip()
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
        if self._scId is not None:
            if self._getScTitle:
                self.scenes[self._scId].title = data.strip()
            elif not data.isspace():
                self._lines.append(data)
        elif self._chId is not None:
            if self.chapters[self._chId].title is None:
                self.chapters[self._chId].title = data.strip()
