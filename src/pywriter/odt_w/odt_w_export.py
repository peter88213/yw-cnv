"""Provide a class for ODT chapters and scenes export.

Copyright (c) 2023 Peter Triesberger
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
from pywriter.pywriter_globals import *
from pywriter.odt_w.odt_w_formatted import OdtWFormatted


class OdtWExport(OdtWFormatted):
    """ODT novel file writer.

    Export a non-reimportable manuscript with chapters and scenes.
    """
    DESCRIPTION = _('manuscript')
    _fileHeader = f'''$ContentHeader<text:p text:style-name="Title">$Title</text:p>
<text:p text:style-name="Subtitle">$AuthorName</text:p>
'''

    _partTemplate = '''<text:h text:style-name="Heading_20_1" text:outline-level="1">$Title</text:h>
'''

    _chapterTemplate = '''<text:h text:style-name="Heading_20_2" text:outline-level="2">$Title</text:h>
'''

    _sceneTemplate = ''''<text:p text:style-name="Text_20_body"><office:annotation><dc:creator>$sceneTitle</dc:creator><text:p>~ ${Title} ~</text:p></office:annotation>$SceneContent</text:p>
    '''

    _appendedSceneTemplate = '''<text:p text:style-name="First_20_line_20_indent"><office:annotation>
<dc:creator>$sceneTitle</dc:creator>
<text:p>~ ${Title} ~</text:p>
</office:annotation>$SceneContent</text:p>
'''

    _sceneDivider = '<text:p text:style-name="Heading_20_4">* * *</text:p>\n'
    _fileFooter = OdtWFormatted._CONTENT_XML_FOOTER

    def _convert_from_yw(self, text, quick=False):
        """Return text, converted from yw7 markup to target format.
        
        Positional arguments:
            text -- string to convert.
        
        Optional arguments:
            quick: bool -- if True, apply a conversion mode for one-liners without formatting.
        
        Extends the superclass method.
        """
        if not quick:
            text = self._remove_inline_code(text)
        text = super()._convert_from_yw(text, quick)
        return(text)

    def _get_chapterMapping(self, chId, chapterNumber):
        """Return a mapping dictionary for a chapter section.
        
        Positional arguments:
            chId: str -- chapter ID.
            chapterNumber: int -- chapter number.
        
        Suppress the chapter title if necessary.
        Extends the superclass method.
        """
        chapterMapping = super()._get_chapterMapping(chId, chapterNumber)
        if self.novel.chapters[chId].suppressChapterTitle:
            chapterMapping['Title'] = ''
        return chapterMapping

