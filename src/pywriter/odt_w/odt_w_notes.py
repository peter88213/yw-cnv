"""Provide a class for ODT invisibly tagged "Notes" chapters export.

Copyright (c) 2023 Peter Triesberger
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
from string import Template
from pywriter.pywriter_globals import *
from pywriter.odt_w.odt_w_manuscript import OdtWManuscript


class OdtWNotes(OdtWManuscript):
    """ODT "Notes" chapters file writer.

    Export a manuscript with invisibly tagged chapters and scenes.
    """
    DESCRIPTION = _('Notes chapters')
    SUFFIX = '_notes'

    _partTemplate = ''
    _chapterTemplate = ''

    _notesPartTemplate = '''<text:section text:style-name="Sect1" text:name="ChID:$ID">
<text:h text:style-name="Heading_20_1" text:outline-level="1">$Title</text:h>
'''

    _notesChapterTemplate = '''<text:section text:style-name="Sect1" text:name="ChID:$ID">
<text:h text:style-name="Heading_20_2" text:outline-level="2">$Title</text:h>
'''

    _notesSceneTemplate = '''<text:h text:style-name="Heading_20_3" text:outline-level="3">$Title</text:h>
<text:section text:style-name="Sect1" text:name="ScID:$ID">
<text:p text:style-name="Text_20_body">$SceneContent</text:p>
</text:section>
'''
    _sceneDivider = ''

    _notesChapterEndTemplate = '''</text:section>
'''

    def _get_chapters(self):
        """Process the chapters and nested scenes.
        
        Iterate through the sorted chapter list and apply the templates, 
        substituting placeholders according to the chapter mapping dictionary.
        For each chapter call the processing of its included scenes.
        Skip chapters not accepted by the chapter filter.
        Return a list of strings.
        This is a template method that can be extended or overridden by subclasses.
        """
        lines = []
        if not self._notesChapterEndTemplate:
            return lines

        chapterNumber = 0
        sceneNumber = 0
        wordsTotal = 0
        lettersTotal = 0
        for chId in self.novel.srtChapters:
            dispNumber = 0
            if not self._chapterFilter.accept(self, chId):
                continue

            # The order counts; be aware that "Notes" chapters are always unused.
            doNotExport = False
            template = None
            if self.novel.chapters[chId].chType == 1:
                # Chapter is "Notes" type (implies "unused").
                if self.novel.chapters[chId].chLevel == 1:
                    # Chapter is "Notes Part" type.
                    if self._notesPartTemplate:
                        template = Template(self._notesPartTemplate)
                elif self._notesChapterTemplate:
                    # Chapter is "Notes Chapter" type.
                    template = Template(self._notesChapterTemplate)
                    chapterNumber += 1
                    dispNumber = chapterNumber
                if template is not None:
                    lines.append(template.safe_substitute(self._get_chapterMapping(chId, dispNumber)))

                    #--- Process scenes.
                    sceneLines, sceneNumber, wordsTotal, lettersTotal = self._get_scenes(
                        chId, sceneNumber, wordsTotal, lettersTotal, doNotExport)
                    lines.extend(sceneLines)

                    #--- Process chapter ending.
                    template = Template(self._notesChapterEndTemplate)
                    lines.append(template.safe_substitute(self._get_chapterMapping(chId, dispNumber)))
        return lines
