"""Provide a class for ODT invisibly tagged chapter descriptions export.

Copyright (c) 2023 Peter Triesberger
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
from pywriter.pywriter_globals import *
from pywriter.odt_w.odt_writer import OdtWriter


class OdtWChapterDesc(OdtWriter):
    """ODT chapter summaries file writer.

    Export a synopsis with invisibly tagged chapter descriptions.
    """
    DESCRIPTION = _('Chapter descriptions')
    SUFFIX = '_chapters'

    _fileHeader = f'''{OdtWriter._CONTENT_XML_HEADER}<text:p text:style-name="Title">$Title</text:p>
<text:p text:style-name="Subtitle">$AuthorName</text:p>
'''

    _partTemplate = '''<text:h text:style-name="Heading_20_1" text:outline-level="1"><text:a xlink:href="../${ProjectName}_parts.odt#ChID:$ID%7Cregion">$Title</text:a></text:h>
'''

    _chapterTemplate = '''<text:section text:style-name="Sect1" text:name="ChID:$ID">
<text:h text:style-name="Heading_20_2" text:outline-level="2"><text:a xlink:href="../${ProjectName}_manuscript.odt#ChID:$ID%7Cregion">$Title</text:a></text:h>
<text:p text:style-name="Text_20_body">$Desc</text:p>
</text:section>
'''

    _fileFooter = OdtWriter._CONTENT_XML_FOOTER
