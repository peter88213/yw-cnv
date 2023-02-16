"""Provide a class for ODT invisibly tagged character descriptions export.

Copyright (c) 2023 Peter Triesberger
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
from pywriter.pywriter_globals import *
from pywriter.odt_w.odt_writer import OdtWriter


class OdtWCharacters(OdtWriter):
    """ODT character descriptions file writer.

    Export a character sheet with invisibly tagged descriptions.
    """
    DESCRIPTION = _('Character descriptions')
    SUFFIX = '_characters'

    _fileHeader = f'''{OdtWriter._CONTENT_XML_HEADER}<text:p text:style-name="Title">$Title</text:p>
<text:p text:style-name="Subtitle">$AuthorName</text:p>
'''

    _characterTemplate = f'''<text:h text:style-name="Heading_20_2" text:outline-level="2">$Title$FullName$AKA</text:h>
<text:section text:style-name="Sect1" text:name="CrID:$ID">
<text:h text:style-name="Heading_20_3" text:outline-level="3">{_("Description")}</text:h>
<text:section text:style-name="Sect1" text:name="CrID_desc:$ID">
<text:p text:style-name="Text_20_body">$Desc</text:p>
</text:section>
<text:h text:style-name="Heading_20_3" text:outline-level="3">{_("Bio")}</text:h>
<text:section text:style-name="Sect1" text:name="CrID_bio:$ID">
<text:p text:style-name="Text_20_body">$Bio</text:p>
</text:section>
<text:h text:style-name="Heading_20_3" text:outline-level="3">{_("Goals")}</text:h>
<text:section text:style-name="Sect1" text:name="CrID_goals:$ID">
<text:p text:style-name="Text_20_body">$Goals</text:p>
</text:section>
<text:h text:style-name="Heading_20_3" text:outline-level="3">{_("Notes")}</text:h>
<text:section text:style-name="Sect1" text:name="CrID_notes:$ID">
<text:p text:style-name="Text_20_body">$Notes</text:p>
</text:section>
</text:section>
'''

    _fileFooter = OdtWriter._CONTENT_XML_FOOTER

    def _get_characterMapping(self, crId):
        """Return a mapping dictionary for a character section.
        
        Positional arguments:
            crId -- str: character ID.
        
        Special formatting of alternate and full name. 
        Extends the superclass method.
        """
        characterMapping = OdtWriter._get_characterMapping(self, crId)
        if self.novel.characters[crId].aka:
            characterMapping['AKA'] = f' ("{self.novel.characters[crId].aka}")'
        if self.novel.characters[crId].fullName:
            characterMapping['FullName'] = f'/{self.novel.characters[crId].fullName}'
        return characterMapping
