"""Provide a class for ODT invisibly tagged scene descriptions export.

Copyright (c) 2023 Peter Triesberger
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
from pywriter.pywriter_globals import *
from pywriter.odt_w.odt_writer import OdtWriter


class OdtWSceneDesc(OdtWriter):
    """ODT scene summaries file writer.

    Export a full synopsis with invisibly tagged scene descriptions.
    """
    DESCRIPTION = _('Scene descriptions')
    SUFFIX = '_scenes'

    _fileHeader = f'''{OdtWriter._CONTENT_XML_HEADER}<text:p text:style-name="Title">$Title</text:p>
<text:p text:style-name="Subtitle">$AuthorName</text:p>
'''

    _partTemplate = '''<text:section text:style-name="Sect1" text:name="ChID:$ID">
<text:h text:style-name="Heading_20_1" text:outline-level="1"><text:a xlink:href="../${ProjectName}_parts.odt#ChID:$ID%7Cregion">$Title</text:a></text:h>
'''

    _chapterTemplate = '''<text:section text:style-name="Sect1" text:name="ChID:$ID">
<text:h text:style-name="Heading_20_2" text:outline-level="2"><text:a xlink:href="../${ProjectName}_chapters.odt#ChID:$ID%7Cregion">$Title</text:a></text:h>
'''

    _sceneTemplate = '''<text:section text:style-name="Sect1" text:name="ScID:$ID">
<text:p text:style-name="Text_20_body"><office:annotation>
<dc:creator>scene title</dc:creator>
<text:p>~ ${Title} ~</text:p>
<text:p/>
<text:p><text:a xlink:href="../${ProjectName}_manuscript.odt#ScID:$ID%7Cregion">→Manuscript</text:a></text:p>
</office:annotation>$Desc</text:p>
</text:section>
'''

    _appendedSceneTemplate = '''<text:section text:style-name="Sect1" text:name="ScID:$ID">
<text:p text:style-name="First_20_line_20_indent"><office:annotation>
<dc:creator>scene title</dc:creator>
<text:p>~ ${Title} ~</text:p>
<text:p/>
<text:p><text:a xlink:href="../${ProjectName}_manuscript.odt#ScID:$ID%7Cregion">→Manuscript</text:a></text:p>
</office:annotation>$Desc</text:p>
</text:section>
'''

    _sceneDivider = '''<text:p text:style-name="Heading_20_4">* * *</text:p>
'''

    _chapterEndTemplate = '''</text:section>
'''

    _fileFooter = OdtWriter._CONTENT_XML_FOOTER
