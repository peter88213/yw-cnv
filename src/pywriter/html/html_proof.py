"""Provide a class for html visibly tagged chapters and scenes import.

Copyright (c) 2022 Peter Triesberger
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import re
from pywriter.pywriter_globals import *
from pywriter.html.html_formatted import HtmlFormatted
from pywriter.model.splitter import Splitter


class HtmlProof(HtmlFormatted):
    """HTML proof reading file representation.

    Import a manuscript with visibly tagged chapters and scenes.
    """
    DESCRIPTION = _('Tagged manuscript for proofing')
    SUFFIX = '_proof'

    def handle_starttag(self, tag, attrs):
        """Recognize the paragraph's beginning.
        
        Positional arguments:
            tag -- str: name of the tag converted to lower case.
            attrs -- list of (name, value) pairs containing the attributes found inside the tag’s <> brackets.
        
        Overrides the superclass method.
        """
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
        elif tag in ('br', 'ul'):
            self._doNothing = True
            # avoid inserting an unwanted blank

    def handle_endtag(self, tag):
        """Recognize the paragraph's end.      
        
        Positional arguments:
            tag -- str: name of the tag converted to lower case.

        Overrides HTMLparser.handle_endtag() called by the HTML parser to handle the end tag of an element.
        """
        if tag in ['p', 'h2', 'h1', 'blockquote']:
            if self._language:
                self._lines.append(f'[/lang={self._language}]')
                self._language = ''
            self._newline = True
        elif tag == 'em' or tag == 'i':
            self._lines.append('[/i]')
        elif tag == 'strong' or tag == 'b':
            self._lines.append('[/b]')
        elif tag == 'span':
            if self._language:
                self._lines.append(f'[/lang={self._language}]')
                self._language = ''

    def handle_data(self, data):
        """Parse the paragraphs and build the document structure.      

        Positional arguments:
            data -- str: text to be parsed. 
        
        Overrides HTMLparser.handle_data() called by the parser to process arbitrary data.
        """
        if self._doNothing:
            self._doNothing = False
        elif '[ScID' in data:
            self._scId = re.search('[0-9]+', data).group()
            self.scenes[self._scId] = self.SCENE_CLASS()
            self.chapters[self._chId].srtScenes.append(self._scId)
            self._lines = []
        elif '[/ScID' in data:
            text = ''.join(self._lines)
            self.scenes[self._scId].sceneContent = self._cleanup_scene(text).strip()
            self._scId = None
        elif '[ChID' in data:
            self._chId = re.search('[0-9]+', data).group()
            self.chapters[self._chId] = self.CHAPTER_CLASS()
            self.srtChapters.append(self._chId)
        elif '[/ChID' in data:
            self._chId = None
        elif self._scId is not None:
            if self._newline:
                self._newline = False
                data = f'{data.rstrip()}\n'
            self._lines.append(data)
