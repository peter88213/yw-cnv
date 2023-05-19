"""Provide a class for ODT visibly tagged chapters and scenes export.

Copyright (c) 2023 Peter Triesberger
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import re
from string import Template
from pywriter.pywriter_globals import *
from pywriter.odt_w.odt_w_formatted import OdtWFormatted


class OdtWProof(OdtWFormatted):
    """ODT proof reading file writer.

    Export a manuscript with visibly tagged chapters and scenes.
    """
    DESCRIPTION = _('Tagged manuscript for proofing')
    SUFFIX = '_proof'

    _fileHeader = f'''$ContentHeader<text:p text:style-name="Title">$Title</text:p>
<text:p text:style-name="Subtitle">$AuthorName</text:p>
'''

    _partTemplate = '''<text:h text:style-name="Heading_20_1" text:outline-level="1">$Title</text:h>
'''

    _chapterTemplate = '''<text:h text:style-name="Heading_20_2" text:outline-level="2">$Title</text:h>
'''

    _sceneTemplate = '''<text:p text:style-name="scene_20_mark">[ScID:$ID]</text:p>
<text:p text:style-name="Text_20_body">$SceneContent</text:p>
<text:p text:style-name="scene_20_mark">[/ScID]</text:p>
'''

    _sceneDivider = '''<text:p text:style-name="Heading_20_4">* * *</text:p>
'''

    _fileFooter = OdtWFormatted._CONTENT_XML_FOOTER

    def _convert_from_yw(self, text, quick=False):
        """Return text, converted from yw7 markup to target format.
        
        Positional arguments:
            text -- string to convert.
        
        Optional arguments:
            quick: bool -- This argument is not used here.
        
        Apply direct formatting instead of the emphasizing character styles.
        Overrides the superclass method.
        """
        if text:
            # Apply XML predefined entities.
            odtReplacements = [
                ('&', '&amp;'),
                ('>', '&gt;'),
                ('<', '&lt;'),
                ("'", '&apos;'),
                ('"', '&quot;'),
                ]
            tags = ['i', 'b']
            odtReplacements.extend([
                ('\n\n', '</text:p>\r<text:p text:style-name="First_20_line_20_indent" />\r<text:p text:style-name="Text_20_body">'),
                ('\n', '</text:p>\r<text:p text:style-name="First_20_line_20_indent">'),
                ('\r', '\n'),
                ('[/i]', '</text:span>'),
                ('[/b]', '</text:span>'),
                ('/*', f'<office:annotation><dc:creator>{self.novel.authorName}</dc:creator><text:p>'),
                ('*/', '</text:p></office:annotation>'),
            ])
            i = 0
            for i, language in enumerate(self.novel.languages, 1):
                tags.append(f'lang={language}')
                odtReplacements.append((f'[lang={language}]', f'<text:span text:style-name="T{i}">'))
                odtReplacements.append((f'[/lang={language}]', '</text:span>'))
            odtReplacements.extend([
                ('[i]', f'<text:span text:style-name="T{i+1}">'),
                ('[b]', f'<text:span text:style-name="T{i+2}">'),
            ])

            #--- Process markup reaching across linebreaks.
            newlines = []
            lines = text.split('\n')
            isOpen = {}
            opening = {}
            closing = {}
            for tag in tags:
                isOpen[tag] = False
                opening[tag] = f'[{tag}]'
                closing[tag] = f'[/{tag}]'
            for line in lines:
                for tag in tags:
                    if isOpen[tag]:
                        if line.startswith('&gt; '):
                            line = f"&gt; {opening[tag]}{line.lstrip('&gt; ')}"
                        else:
                            line = f'{opening[tag]}{line}'
                        isOpen[tag] = False
                    while line.count(opening[tag]) > line.count(closing[tag]):
                        line = f'{line}{closing[tag]}'
                        isOpen[tag] = True
                    while line.count(closing[tag]) > line.count(opening[tag]):
                        line = f'{opening[tag]}{line}'
                    line = line.replace(f'{opening[tag]}{closing[tag]}', '')
                newlines.append(line)
            text = '\n'.join(newlines).rstrip()

            #--- Apply odt formating.
            for yw, od in odtReplacements:
                text = text.replace(yw, od)

            # Remove highlighting, alignment,
            # strikethrough, and underline tags.
            text = re.sub('\[\/*[h|c|r|s|u]\d*\]', '', text)
        else:
            text = ''
        return text

    def _get_fileHeaderMapping(self):
        """Return a mapping dictionary for the project section.
        
        Add "automatic-styles" items to the "content.xml" header, if required.
        
        Extends the superclass method.
        """
        styleMapping = {}
        i = 0
        lines = ['<office:automatic-styles>']
        for i, language in enumerate(self.novel.languages, 1):
            try:
                lngCode, ctrCode = language.split('-')
            except:
                lngCode = 'zxx'
                ctrCode = 'none'
            lines.append(f'''  <style:style style:name="T{i}" style:family="text">
   <style:text-properties fo:language="{lngCode}" fo:country="{ctrCode}" style:language-asian="{lngCode}" style:country-asian="{ctrCode}" style:language-complex="{lngCode}" style:country-complex="{ctrCode}"/>
  </style:style>''')
        lines.append(f'''  <style:style style:name="T{i+1}" style:family="text">
   <style:text-properties fo:font-style="italic" style:font-style-asian="italic" style:font-style-complex="italic"/>
  </style:style>''')
        lines.append(f'''  <style:style style:name="T{i+2}" style:family="text">
   <style:text-properties fo:font-weight="bold" style:font-weight-asian="bold" style:font-weight-complex="bold"/>
  </style:style>''')
        lines.append(' </office:automatic-styles>')
        styleMapping['automaticStyles'] = '\n'.join(lines)
        template = Template(self._CONTENT_XML_HEADER)
        projectTemplateMapping = super()._get_fileHeaderMapping()
        projectTemplateMapping['ContentHeader'] = template.safe_substitute(styleMapping)
        return projectTemplateMapping

    def _get_text(self):
        """Call all processing methods.
        
        Return a string to be written to the output file.
        Overrides the superclass method.
        """
        lines = self._get_fileHeader()
        lines.extend(self._get_chapters())
        lines.append(self._fileFooter)
        text = ''.join(lines)
        return text
