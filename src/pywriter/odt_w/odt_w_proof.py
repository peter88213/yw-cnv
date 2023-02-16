"""Provide a class for ODT visibly tagged chapters and scenes export.

Copyright (c) 2023 Peter Triesberger
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
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

    _partTemplate = '''<text:p text:style-name="yWriter_20_mark">[ChID:$ID]</text:p>
<text:h text:style-name="Heading_20_1" text:outline-level="1">$Title</text:h>
'''

    _chapterTemplate = '''<text:p text:style-name="yWriter_20_mark">[ChID:$ID]</text:p>
<text:h text:style-name="Heading_20_2" text:outline-level="2">$Title</text:h>
'''

    _unusedChapterTemplate = '''<text:p text:style-name="yWriter_20_mark_20_unused">[ChID:$ID (Unused)]</text:p>
<text:h text:style-name="Heading_20_2" text:outline-level="2">$Title</text:h>
'''

    _notesChapterTemplate = '''<text:p text:style-name="yWriter_20_mark_20_notes">[ChID:$ID (Notes)]</text:p>
<text:h text:style-name="Heading_20_2" text:outline-level="2">$Title</text:h>
'''

    _todoChapterTemplate = '''<text:p text:style-name="yWriter_20_mark_20_todo">[ChID:$ID (ToDo)]</text:p>
<text:h text:style-name="Heading_20_2" text:outline-level="2">$Title</text:h>
'''

    _sceneTemplate = '''<text:p text:style-name="yWriter_20_mark">[ScID:$ID]</text:p>
<text:p text:style-name="Text_20_body">$SceneContent</text:p>
<text:p text:style-name="yWriter_20_mark">[/ScID]</text:p>
'''

    _unusedSceneTemplate = '''<text:p text:style-name="yWriter_20_mark_20_unused">[ScID:$ID (Unused)]</text:p>
<text:p text:style-name="Text_20_body">$SceneContent</text:p>
<text:p text:style-name="yWriter_20_mark_20_unused">[/ScID (Unused)]</text:p>
'''

    _notesSceneTemplate = '''<text:p text:style-name="yWriter_20_mark_20_notes">[ScID:$ID (Notes)]</text:p>
<text:p text:style-name="Text_20_body">$SceneContent</text:p>
<text:p text:style-name="yWriter_20_mark_20_notes">[/ScID (Notes)]</text:p>
'''

    _todoSceneTemplate = '''<text:p text:style-name="yWriter_20_mark_20_todo">[ScID:$ID (ToDo)]</text:p>
<text:p text:style-name="Text_20_body">$SceneContent</text:p>
<text:p text:style-name="yWriter_20_mark_20_todo">[/ScID (ToDo)]</text:p>
'''

    _sceneDivider = '''<text:p text:style-name="Heading_20_4">* * *</text:p>
'''

    _chapterEndTemplate = '''<text:p text:style-name="yWriter_20_mark">[/ChID]</text:p>
'''

    _unusedChapterEndTemplate = '''<text:p text:style-name="yWriter_20_mark_20_unused">[/ChID (Unused)]</text:p>
'''

    _notesChapterEndTemplate = '''<text:p text:style-name="yWriter_20_mark_20_notes">[/ChID (Notes)]</text:p>
'''

    _todoChapterEndTemplate = '''<text:p text:style-name="yWriter_20_mark_20_todo">[/ChID (ToDo)]</text:p>
'''

    _fileFooter = OdtWFormatted._CONTENT_XML_FOOTER
