"""Provide a class for ODT cross reference export.

Copyright (c) 2023 Peter Triesberger
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
from string import Template
from pywriter.pywriter_globals import *
from pywriter.model.cross_references import CrossReferences
from pywriter.odt_w.odt_writer import OdtWriter


class OdtWXref(OdtWriter):
    """OpenDocument xml cross reference file writer."""
    DESCRIPTION = _('Cross reference')
    SUFFIX = '_xref'

    _fileHeader = f'''{OdtWriter._CONTENT_XML_HEADER}<text:p text:style-name="Title">$Title</text:p>
<text:p text:style-name="Subtitle">$AuthorName</text:p>
'''
    _sceneTemplate = '''<text:p text:style-name="yWriter_20_mark">
<text:a xlink:href="../${ProjectName}_manuscript.odt#ScID:$ID%7Cregion">$SceneNumber</text:a> (Ch $Chapter) $Title
</text:p>
'''
    _unusedSceneTemplate = '''<text:p text:style-name="yWriter_20_mark_20_unused">
$SceneNumber (Ch $Chapter) $Title (Unused)
</text:p>
'''
    _notesSceneTemplate = '''<text:p text:style-name="yWriter_20_mark_20_notes">
$SceneNumber (Ch $Chapter) $Title (Notes)
</text:p>
'''
    _todoSceneTemplate = '''<text:p text:style-name="yWriter_20_mark_20_todo">
$SceneNumber (Ch $Chapter) $Title (ToDo)
</text:p>
'''
    _characterTemplate = '''<text:p text:style-name="Text_20_body">
<text:a xlink:href="../${ProjectName}_characters.odt#CrID:$ID%7Cregion">$Title</text:a> $FullName
</text:p>
'''
    _locationTemplate = '''<text:p text:style-name="Text_20_body">
<text:a xlink:href="../${ProjectName}_locations.odt#LcID:$ID%7Cregion">$Title</text:a>
</text:p>
'''
    _itemTemplate = '''<text:p text:style-name="Text_20_body">
<text:a xlink:href="../${ProjectName}_items.odt#ItrID:$ID%7Cregion">$Title</text:a>
</text:p>
'''
    _scnPerChrTemplate = '''<text:h text:style-name="Heading_20_2" text:outline-level="2">Scenes with Character $Title:</text:h>
'''
    _scnPerLocTemplate = '''<text:h text:style-name="Heading_20_2" text:outline-level="2">Scenes with Location $Title:</text:h>
'''
    _scnPerItmTemplate = '''<text:h text:style-name="Heading_20_2" text:outline-level="2">Scenes with Item $Title:</text:h>
'''
    _chrPerTagTemplate = '''<text:h text:style-name="Heading_20_2" text:outline-level="2">Characters tagged $Tag:</text:h>
'''
    _locPerTagTemplate = '''<text:h text:style-name="Heading_20_2" text:outline-level="2">Locations tagged $Tag:</text:h>
'''
    _itmPerTagTemplate = '''<text:h text:style-name="Heading_20_2" text:outline-level="2">Items tagged $Tag:</text:h>
'''
    _scnPerTagtemplate = '''<text:h text:style-name="Heading_20_2" text:outline-level="2">Scenes tagged $Tag:</text:h>
'''
    _fileFooter = OdtWriter._CONTENT_XML_FOOTER

    def __init__(self, filePath, **kwargs):
        """Apply the strategy pattern by delegating the cross reference to an external object.
        
        Extends the superclass constructor.
        """
        super().__init__(filePath)
        self._xr = CrossReferences()

    def _get_sceneMapping(self, scId):
        """Return a mapping dictionary for a scene section.

        Positional arguments:
            scId -- str: scene ID.
        
        Extends the superclass template method.
        """
        sceneNumber = self._xr.srtScenes.index(scId) + 1
        sceneMapping = super()._get_sceneMapping(scId, sceneNumber, 0, 0)
        chapterNumber = self.novel.srtChapters.index(self._xr.chpPerScn[scId]) + 1
        sceneMapping['Chapter'] = str(chapterNumber)
        return sceneMapping

    def _get_tagMapping(self, tag):
        """Return a mapping dictionary for a tags section. 

        Positional arguments:
            tag -- str: a single scene tag.
        """
        tagMapping = dict(
            Tag=tag,
        )
        return tagMapping

    def _get_scenes(self, scenes):
        """Process the scenes.
        
        Positional arguments:
            scenes -- iterable of scene IDs.
        
        Return a list of strings.
        Overrides the superclass method.
        """
        lines = []
        for scId in scenes:
            if self.novel.scenes[scId].scType == 1:
                template = Template(self._notesSceneTemplate)
            elif self.novel.scenes[scId].scType == 2:
                template = Template(self._todoSceneTemplate)
            elif self.novel.scenes[scId].scType == 3:
                template = Template(self._unusedSceneTemplate)
            else:
                template = Template(self._sceneTemplate)
            lines.append(template.safe_substitute(self._get_sceneMapping(scId)))
        return lines

    def _get_sceneTags(self):
        """Process the scene related tags.
        
        Return a list of strings.
        """
        lines = []
        headerTemplate = Template(self._scnPerTagtemplate)
        for tag in self._xr.scnPerTag:
            if self._xr.scnPerTag[tag]:
                lines.append(headerTemplate.safe_substitute(self._get_tagMapping(tag)))
                lines.extend(self._get_scenes(self._xr.scnPerTag[tag]))
        return lines

    def _get_characters(self):
        """Process the scenes per character.
        
        Return a list of strings.
        Overrides the superclass method.
        """
        lines = []
        headerTemplate = Template(self._scnPerChrTemplate)
        for crId in self._xr.scnPerChr:
            if self._xr.scnPerChr[crId]:
                lines.append(headerTemplate.safe_substitute(self._get_characterMapping(crId)))
                lines.extend(self._get_scenes(self._xr.scnPerChr[crId]))
        return lines

    def _get_locations(self):
        """Process the locations.
        
        Return a list of strings.
        Overrides the superclass method.
        """
        lines = []
        headerTemplate = Template(self._scnPerLocTemplate)
        for lcId in self._xr.scnPerLoc:
            if self._xr.scnPerLoc[lcId]:
                lines.append(headerTemplate.safe_substitute(self._get_locationMapping(lcId)))
                lines.extend(self._get_scenes(self._xr.scnPerLoc[lcId]))
        return lines

    def _get_items(self):
        """Process the items.
        
        Return a list of strings.
        Overrides the superclass method.
        """
        lines = []
        headerTemplate = Template(self._scnPerItmTemplate)
        for itId in self._xr.scnPerItm:
            if self._xr.scnPerItm[itId]:
                lines.append(headerTemplate.safe_substitute(self._get_itemMapping(itId)))
                lines.extend(self._get_scenes(self._xr.scnPerItm[itId]))
        return lines

    def _get_characterTags(self):
        """Process the character related tags.
        
        Return a list of strings.
        """
        lines = []
        headerTemplate = Template(self._chrPerTagTemplate)
        template = Template(self._characterTemplate)
        for tag in self._xr.chrPerTag:
            if self._xr.chrPerTag[tag]:
                lines.append(headerTemplate.safe_substitute(self._get_tagMapping(tag)))
                for crId in self._xr.chrPerTag[tag]:
                    lines.append(template.safe_substitute(self._get_characterMapping(crId)))
        return lines

    def _get_locationTags(self):
        """Process the location related tags.
        
        Return a list of strings.
        """
        lines = []
        headerTemplate = Template(self._locPerTagTemplate)
        template = Template(self._locationTemplate)
        for tag in self._xr.locPerTag:
            if self._xr.locPerTag[tag]:
                lines.append(headerTemplate.safe_substitute(self._get_tagMapping(tag)))
                for lcId in self._xr.locPerTag[tag]:
                    lines.append(template.safe_substitute(self._get_locationMapping(lcId)))
        return lines

    def _get_itemTags(self):
        """Process the item related tags.
        
        Return a list of strings.
        """
        lines = []
        headerTemplate = Template(self._itmPerTagTemplate)
        template = Template(self._itemTemplate)
        for tag in self._xr.itmPerTag:
            if self._xr.itmPerTag[tag]:
                lines.append(headerTemplate.safe_substitute(self._get_tagMapping(tag)))
                for itId in self._xr.itmPerTag[tag]:
                    lines.append(template.safe_substitute(self._get_itemMapping(itId)))
        return lines

    def _get_text(self):
        """Call all processing methods.
        
        Return a string to be written to the output file.
        Overrides the superclass method.
        """
        self._xr.generate_xref(self.novel)
        lines = self._get_fileHeader()
        lines.extend(self._get_characters())
        lines.extend(self._get_locations())
        lines.extend(self._get_items())
        lines.extend(self._get_sceneTags())
        lines.extend(self._get_characterTags())
        lines.extend(self._get_locationTags())
        lines.extend(self._get_itemTags())
        lines.append(self._fileFooter)
        return ''.join(lines)
