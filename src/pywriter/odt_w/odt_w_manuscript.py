"""Provide a class for ODT invisibly tagged chapters and scenes export.

Copyright (c) 2023 Peter Triesberger
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
from pywriter.pywriter_globals import *
from pywriter.odt_w.odt_w_formatted import OdtWFormatted


class OdtWManuscript(OdtWFormatted):
    """ODT manuscript file writer.

    Export a manuscript with invisibly tagged chapters and scenes.
    """
    DESCRIPTION = _('Editable manuscript')
    SUFFIX = '_manuscript'

    _fileHeader = f'''$ContentHeader<text:p text:style-name="Title">$Title</text:p>
<text:p text:style-name="Subtitle">$AuthorName</text:p>
'''

    _partTemplate = '''<text:section text:style-name="Sect1" text:name="ChID:$ID">
<text:h text:style-name="Heading_20_1" text:outline-level="1"><text:a xlink:href="../${ProjectName}_parts.odt#ChID:$ID%7Cregion">$Title</text:a></text:h>
'''

    _chapterTemplate = '''<text:section text:style-name="Sect1" text:name="ChID:$ID">
<text:h text:style-name="Heading_20_2" text:outline-level="2"><text:a xlink:href="../${ProjectName}_chapters.odt#ChID:$ID%7Cregion">$Title</text:a></text:h>
'''

    _sceneTemplate = '''<text:section text:style-name="Sect1" text:name="ScID:$ID">
<text:p text:style-name="Text_20_body"><office:annotation><dc:creator>$sceneTitle</dc:creator><text:p>~ ${Title} ~</text:p><text:p/><text:p><text:a xlink:href="../${ProjectName}_scenes.odt#ScID:$ID%7Cregion">→$Summary</text:a></text:p></office:annotation>$SceneContent</text:p>
</text:section>
'''

    _appendedSceneTemplate = '''<text:section text:style-name="Sect1" text:name="ScID:$ID">
<text:p text:style-name="First_20_line_20_indent"><office:annotation>
<dc:creator>$sceneTitle</dc:creator>
<text:p>~ ${Title} ~</text:p>
<text:p/>
<text:p><text:a xlink:href="../${ProjectName}_scenes.odt#ScID:$ID%7Cregion">→$Summary</text:a></text:p>
</office:annotation>$SceneContent</text:p>
</text:section>
'''

    _sceneDivider = '<text:p text:style-name="Heading_20_4">* * *</text:p>\n'

    _chapterEndTemplate = '''</text:section>
'''

    _fileFooter = OdtWFormatted._CONTENT_XML_FOOTER

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

    def _get_sceneMapping(self, scId, sceneNumber, wordsTotal, lettersTotal):
        """Return a mapping dictionary for a scene section.
        
        Positional arguments:
            scId: str -- scene ID.
            sceneNumber: int -- scene number to be displayed.
            wordsTotal: int -- accumulated wordcount.
            lettersTotal: int -- accumulated lettercount.
        
        Extends the superclass method.
        """
        sceneMapping = super()._get_sceneMapping(scId, sceneNumber, wordsTotal, lettersTotal)
        sceneMapping['Summary'] = _('Summary')
        return sceneMapping

