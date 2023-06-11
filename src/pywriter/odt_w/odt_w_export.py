"""Provide a class for ODT chapters and scenes export.

Copyright (c) 2023 Peter Triesberger
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import re
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

    def _get_replacements(self):
        return [
                ('\n\n', '</text:p>\r<text:p text:style-name="First_20_line_20_indent" />\r<text:p text:style-name="Text_20_body">'),
                ('\n', '</text:p>\r<text:p text:style-name="First_20_line_20_indent">'),
                ('\r', '\n'),
                ('[i]', '<text:span text:style-name="Emphasis">'),
                ('[/i]', '</text:span>'),
                ('[b]', '<text:span text:style-name="Strong_20_Emphasis">'),
                ('[/b]', '</text:span>'),
                ]

    def _get_text(self):

        def replace_note(match):
            noteType = match.group(1)
            self._noteCounter += 1
            self._noteNumber += 1
            noteLabel = f'{self._noteNumber}'
            if noteType.startswith('fn'):
                noteClass = 'footnote'
                noteStyle = 'Footnote'
                if noteType.endswith('*'):
                    self._noteNumber -= 1
                    noteLabel = '*'
            elif noteType.startswith('en'):
                noteClass = 'endnote'
                noteStyle = 'Endnote'
            text = match.group(2).replace('text:style-name="First_20_line_20_indent"', f'text:style-name="{noteStyle}"')
            return f'<text:note text:id="ftn{self._noteCounter}" text:note-class="{noteClass}"><text:note-citation text:label="{noteLabel}">*</text:note-citation><text:note-body><text:p text:style-name="{noteStyle}">{text}</text:p></text:note-body></text:note>'

        text = super()._get_text()
        if text.find('/*') > 0:
            text = text.replace('\r', '@r@').replace('\n', '@n@')
            self._noteCounter = 0
            self._noteNumber = 0
            simpleComment = f'<office:annotation><dc:creator>{self.novel.authorName}</dc:creator><text:p>\\1</text:p></office:annotation>'
            text = re.sub('\/\* @([ef]n\**) (.*?)\*\/', replace_note, text)
            text = re.sub('\/\*(.*?)\*\/', simpleComment, text)
            text = text.replace('@r@', '\r').replace('@n@', '\n')
        return text

