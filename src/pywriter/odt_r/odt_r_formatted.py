"""Provide a base class for ODT documents containing text that is formatted in yWriter.

Copyright (c) 2023 Peter Triesberger
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
from pywriter.odt_r.odt_reader import OdtReader
from pywriter.model.splitter import Splitter


class OdtRFormatted(OdtReader):
    """ODT file reader.
    
    Public methods:
        read() -- Parse the file and get the instance variables.

    Provide methods and data for processing chapters with formatted text.
    """
    _COMMENT_START = '/*'
    _COMMENT_END = '*/'
    _SC_TITLE_BRACKET = '~'
    _BULLET = '-'
    _INDENT = '>'

    def read(self):
        """Parse the file and get the instance variables.
        
        Extends the superclass method.
        """
        self.novel.languages = []
        super().read()

        # Split scenes, if necessary.
        sceneSplitter = Splitter()
        self.scenesSplit = sceneSplitter.split_scenes(self)

    def _cleanup_scene(self, text):
        """Clean up yWriter markup.
        
        Positional arguments:
            text -- string to clean up.
        
        Return a yw7 markup string.
        """
        #--- Remove redundant tags.
        # In contrast to Office Writer, yWriter accepts markup reaching across linebreaks.
        tags = ['i', 'b']
        for language in self.novel.languages:
            tags.append(f'lang={language}')
        for tag in tags:
            text = text.replace(f'[/{tag}][{tag}]', '')
            text = text.replace(f'[/{tag}]\n[{tag}]', '\n')
            text = text.replace(f'[/{tag}]\n> [{tag}]', '\n> ')

        #--- Remove misplaced formatting tags.
        # text = re.sub('\[\/*[b|i]\]', '', text)
        return text

