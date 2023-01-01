"""Provide a base class for ODT documents containing text that is formatted in yWriter.

Copyright (c) 2022 Peter Triesberger
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import re
from string import Template
from pywriter.pywriter_globals import *
from pywriter.odt.odt_file import OdtFile


class OdtFormatted(OdtFile):
    """ODT file representation.

    Public methods:
        write() -- Determine the languages used in the document before writing.
    
    Provide methods for processing chapters with formatted text.
    """
    _CONTENT_XML_HEADER = '''<?xml version="1.0" encoding="UTF-8"?>

<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" xmlns:math="http://www.w3.org/1998/Math/MathML" xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" xmlns:ooo="http://openoffice.org/2004/office" xmlns:ooow="http://openoffice.org/2004/writer" xmlns:oooc="http://openoffice.org/2004/calc" xmlns:dom="http://www.w3.org/2001/xml-events" xmlns:xforms="http://www.w3.org/2002/xforms" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:rpt="http://openoffice.org/2005/report" xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2" xmlns:xhtml="http://www.w3.org/1999/xhtml" xmlns:grddl="http://www.w3.org/2003/g/data-view#" xmlns:tableooo="http://openoffice.org/2009/table" xmlns:field="urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0" office:version="1.2">
 <office:scripts/>
 <office:font-face-decls>
  <style:font-face style:name="StarSymbol" svg:font-family="StarSymbol" style:font-charset="x-symbol"/>
  <style:font-face style:name="Consolas" svg:font-family="Consolas" style:font-adornments="Standard" style:font-family-generic="modern" style:font-pitch="fixed"/>
  <style:font-face style:name="Courier New" svg:font-family="&apos;Courier New&apos;" style:font-adornments="Standard" style:font-family-generic="modern" style:font-pitch="fixed"/>
 </office:font-face-decls>
 $automaticStyles
 <office:body>
  <office:text text:use-soft-page-breaks="true">

'''

    def _convert_from_yw(self, text, quick=False):
        """Return text, converted from yw7 markup to target format.
        
        Positional arguments:
            text -- string to convert.
        
        Optional arguments:
            quick -- bool: if True, apply a conversion mode for one-liners without formatting.
        
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
            if not quick:
                tags = ['i', 'b']
                odtReplacements.extend([
                    ('\n\n', '</text:p>\r<text:p text:style-name="First_20_line_20_indent" />\r<text:p text:style-name="Text_20_body">'),
                    ('\n', '</text:p>\r<text:p text:style-name="First_20_line_20_indent">'),
                    ('\r', '\n'),
                    ('[i]', '<text:span text:style-name="Emphasis">'),
                    ('[/i]', '</text:span>'),
                    ('[b]', '<text:span text:style-name="Strong_20_Emphasis">'),
                    ('[/b]', '</text:span>'),
                    ('/*', f'<office:annotation><dc:creator>{self.authorName}</dc:creator><text:p>'),
                    ('*/', '</text:p></office:annotation>'),
                ])
                for i, language in enumerate(self.languages, 1):
                    tags.append(f'lang={language}')
                    odtReplacements.append((f'[lang={language}]', f'<text:span text:style-name="T{i}">'))
                    odtReplacements.append((f'[/lang={language}]', '</text:span>'))

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

    def _get_text(self):
        """Call all processing methods.
        
        Return a string to be written to the output file.
        Overrides the superclass method.
        """
        lines = self._get_fileHeader()
        lines.extend(self._get_chapters())
        lines.append(self._fileFooter)
        text = ''.join(lines)

        # Set style of paragraphs that start with "> " to "Quotations".
        # This is done here to include the scene openings.
        if '&gt; ' in text:
            quotMarks = ('"First_20_line_20_indent">&gt; ',
                         '"Text_20_body">&gt; ',
                         )
            for quotMark in quotMarks:
                text = text.replace(quotMark, '"Quotations">')
            text = re.sub('"Text_20_body"\>(\<office\:annotation\>.+?\<\/office\:annotation\>)\&gt\; ', '"Quotations">\\1', text)
        return text

    def _get_fileHeaderMapping(self):
        """Return a mapping dictionary for the project section.
        
        Add "automatic-styles" items to the "content.xml" header, if required.
        
        Extends the superclass method.
        """
        styleMapping = {}
        if self.languages:
            lines = ['<office:automatic-styles>']
            for i, language in enumerate(self.languages, 1):
                try:
                    lngCode, ctrCode = language.split('-')
                except:
                    lngCode = 'zxx'
                    ctrCode = 'none'
                lines.append(f'''  <style:style style:name="T{i}" style:family="text">
   <style:text-properties fo:language="{lngCode}" fo:country="{ctrCode}" style:language-asian="{lngCode}" style:country-asian="{ctrCode}" style:language-complex="{lngCode}" style:country-complex="{ctrCode}"/>
  </style:style>''')
            lines.append(' </office:automatic-styles>')
            styleMapping['automaticStyles'] = '\n'.join(lines)
        else:
            styleMapping['automaticStyles'] = '<office:automatic-styles/>'
        template = Template(self._CONTENT_XML_HEADER)
        projectTemplateMapping = super()._get_fileHeaderMapping()
        projectTemplateMapping['ContentHeader'] = template.safe_substitute(styleMapping)
        return projectTemplateMapping

    def write(self):
        """Determine the languages used in the document before writing.
        
        Extends the superclass method.
        """
        if self.languages is None:
            self.get_languages()
        return super().write()
