"""Convert yWriter project to odt or ods and vice versa. 

Version 1.0.1

Copyright (c) 2021 Peter Triesberger
For further information see https://github.com/peter88213/yw-cnv
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import os

from configparser import ConfigParser
from urllib.parse import unquote
from urllib.parse import quote

import re

import zipfile
import locale
from shutil import rmtree
from datetime import datetime
from string import Template

from string import Template




class WorldElement():
    """Story world element representation.
    # xml: <LOCATIONS><LOCATION> or # xml: <ITEMS><ITEM>
    """

    def __init__(self):
        self.title = None
        # str
        # xml: <Title>

        self.image = None
        # str
        # xml: <ImageFile>

        self.desc = None
        # str
        # xml: <Desc>

        self.tags = None
        # list of str
        # xml: <Tags>

        self.aka = None
        # str
        # xml: <AKA>


class Character(WorldElement):
    """yWriter character representation.
    # xml: <CHARACTERS><CHARACTER>
    """

    MAJOR_MARKER = 'Major'
    MINOR_MARKER = 'Minor'

    def __init__(self):
        WorldElement.__init__(self)

        self.notes = None
        # str
        # xml: <Notes>

        self.bio = None
        # str
        # xml: <Bio>

        self.goals = None
        # str
        # xml: <Goals>

        self.fullName = None
        # str
        # xml: <FullName>

        self.isMajor = None
        # bool
        # xml: <Major>


class Scene():
    """yWriter scene representation.
    # xml: <SCENES><SCENE>
    """

    # Emulate an enumeration for the scene status

    STATUS = [None, 'Outline', 'Draft', '1st Edit', '2nd Edit', 'Done']
    ACTION_MARKER = 'A'
    REACTION_MARKER = 'R'

    def __init__(self):
        self.title = None
        # str
        # xml: <Title>

        self.desc = None
        # str
        # xml: <Desc>

        self._sceneContent = None
        # str
        # xml: <SceneContent>
        # Scene text with yW7 raw markup.

        self.rtfFile = None
        # str
        # xml: <RTFFile>
        # Name of the file containing the scene in yWriter 5.

        self.wordCount = 0
        # int # xml: <WordCount>
        # To be updated by the sceneContent setter

        self.letterCount = 0
        # int
        # xml: <LetterCount>
        # To be updated by the sceneContent setter

        self.isUnused = None
        # bool
        # xml: <Unused> -1

        self.isNotesScene = None
        # bool
        # xml: <Fields><Field_SceneType> 1

        self.isTodoScene = None
        # bool
        # xml: <Fields><Field_SceneType> 2

        self.doNotExport = None
        # bool
        # xml: <ExportCondSpecific><ExportWhenRTF>

        self.status = None
        # int # xml: <Status>

        self.sceneNotes = None
        # str
        # xml: <Notes>

        self.tags = None
        # list of str
        # xml: <Tags>

        self.field1 = None
        # str
        # xml: <Field1>

        self.field2 = None
        # str
        # xml: <Field2>

        self.field3 = None
        # str
        # xml: <Field3>

        self.field4 = None
        # str
        # xml: <Field4>

        self.appendToPrev = None
        # bool
        # xml: <AppendToPrev> -1

        self.isReactionScene = None
        # bool
        # xml: <ReactionScene> -1

        self.isSubPlot = None
        # bool
        # xml: <SubPlot> -1

        self.goal = None
        # str
        # xml: <Goal>

        self.conflict = None
        # str
        # xml: <Conflict>

        self.outcome = None
        # str
        # xml: <Outcome>

        self.characters = None
        # list of str
        # xml: <Characters><CharID>

        self.locations = None
        # list of str
        # xml: <Locations><LocID>

        self.items = None
        # list of str
        # xml: <Items><ItemID>

        self.date = None
        # str
        # xml: <SpecificDateMode>-1
        # xml: <SpecificDateTime>1900-06-01 20:38:00

        self.time = None
        # str
        # xml: <SpecificDateMode>-1
        # xml: <SpecificDateTime>1900-06-01 20:38:00

        self.minute = None
        # str
        # xml: <Minute>

        self.hour = None
        # str
        # xml: <Hour>

        self.day = None
        # str
        # xml: <Day>

        self.lastsMinutes = None
        # str
        # xml: <LastsMinutes>

        self.lastsHours = None
        # str
        # xml: <LastsHours>

        self.lastsDays = None
        # str
        # xml: <LastsDays>

    @property
    def sceneContent(self):
        return self._sceneContent

    @sceneContent.setter
    def sceneContent(self, text):
        """Set sceneContent updating word count and letter count."""
        self._sceneContent = text
        text = re.sub('\[.+?\]|\.|\,| -', '', self._sceneContent)
        # Remove yWriter raw markup for word count

        wordList = text.split()
        self.wordCount = len(wordList)

        text = re.sub('\[.+?\]', '', self._sceneContent)
        # Remove yWriter raw markup for letter count

        text = text.replace('\n', '')
        text = text.replace('\r', '')
        self.letterCount = len(text)
from urllib.parse import quote


class Novel():
    """Abstract yWriter project file representation.

    This class represents a file containing a novel with additional 
    attributes and structural information (a full set or a subset
    of the information included in an yWriter project file).

    Public methods: 
        read() -- Parse the file and store selected properties.
        merge(novel) -- Copy required attributes of the novel object.
        write() -- Write selected properties to the file.
        convert_to_yw(text) -- Return text, converted from source format to yw7 markup.
        convert_from_yw(text) -- Return text, converted from yw7 markup to target format.
        file_exists() -- Return True, if the file specified by filePath exists.

    Instance variables:
        title -- str; title
        desc -- str; description
        author -- str; author name
        fieldTitle1 -- str; field title 1
        fieldTitle2 -- str; field title 2
        fieldTitle3 -- str; field title 3
        fieldTitle4 -- str; field title 4
        chapters -- dict; key = chapter ID, value = Chapter instance.
        scenes -- dict; key = scene ID, value = Scene instance.
        srtChapters -- list of str; The novel's sorted chapter IDs. 
        locations -- dict; key = location ID, value = WorldElement instance.
        srtLocations -- list of str; The novel's sorted location IDs. 
        items -- dict; key = item ID, value = WorldElement instance.
        srtItems -- list of str; The novel's sorted item IDs. 
        characters -- dict; key = character ID, value = Character instance.
        srtCharacters -- list of str The novel's sorted character IDs.
        filePath -- str; path to the file represented by the class.   
    """

    DESCRIPTION = 'Novel'
    EXTENSION = None
    SUFFIX = None
    # To be extended by subclass methods.

    def __init__(self, filePath, **kwargs):
        """Define instance variables.

        Positional argument:
            filePath -- string; path to the file represented by the class.
        """
        self.title = None
        # str
        # xml: <PROJECT><Title>

        self.desc = None
        # str
        # xml: <PROJECT><Desc>

        self.author = None
        # str
        # xml: <PROJECT><AuthorName>

        self.fieldTitle1 = None
        # str
        # xml: <PROJECT><FieldTitle1>

        self.fieldTitle2 = None
        # str
        # xml: <PROJECT><FieldTitle2>

        self.fieldTitle3 = None
        # str
        # xml: <PROJECT><FieldTitle3>

        self.fieldTitle4 = None
        # str
        # xml: <PROJECT><FieldTitle4>

        self.chapters = {}
        # dict
        # xml: <CHAPTERS><CHAPTER><ID>
        # key = chapter ID, value = Chapter instance.
        # The order of the elements does not matter (the novel's
        # order of the chapters is defined by srtChapters)

        self.scenes = {}
        # dict
        # xml: <SCENES><SCENE><ID>
        # key = scene ID, value = Scene instance.
        # The order of the elements does not matter (the novel's
        # order of the scenes is defined by the order of the chapters
        # and the order of the scenes within the chapters)

        self.srtChapters = []
        # list of str
        # The novel's chapter IDs. The order of its elements
        # corresponds to the novel's order of the chapters.

        self.locations = {}
        # dict
        # xml: <LOCATIONS>
        # key = location ID, value = WorldElement instance.
        # The order of the elements does not matter.

        self.srtLocations = []
        # list of str
        # The novel's location IDs. The order of its elements
        # corresponds to the XML project file.

        self.items = {}
        # dict
        # xml: <ITEMS>
        # key = item ID, value = WorldElement instance.
        # The order of the elements does not matter.

        self.srtItems = []
        # list of str
        # The novel's item IDs. The order of its elements
        # corresponds to the XML project file.

        self.characters = {}
        # dict
        # xml: <CHARACTERS>
        # key = character ID, value = Character instance.
        # The order of the elements does not matter.

        self.srtCharacters = []
        # list of str
        # The novel's character IDs. The order of its elements
        # corresponds to the XML project file.

        self._filePath = None
        # str
        # Path to the file. The setter only accepts files of a
        # supported type as specified by EXTENSION.

        self._projectName = None
        # str
        # URL-coded file name without suffix and extension.

        self._projectPath = None
        # str
        # URL-coded path to the project directory.

        self.filePath = filePath

    @property
    def filePath(self):
        return self._filePath

    @filePath.setter
    def filePath(self, filePath):
        """Setter for the filePath instance variable.        
        - Format the path string according to Python's requirements. 
        - Accept only filenames with the right suffix and extension.
        """

        if self.SUFFIX is not None:
            suffix = self.SUFFIX

        else:
            suffix = ''

        if filePath.lower().endswith(suffix + self.EXTENSION):
            self._filePath = filePath
            head, tail = os.path.split(os.path.realpath(filePath))
            self.projectPath = quote(head.replace('\\', '/'), '/:')
            self.projectName = quote(tail.replace(
                suffix + self.EXTENSION, ''))

    def read(self):
        """Parse the file and store selected properties.
        Return a message beginning with SUCCESS or ERROR.
        This is a stub to be overridden by subclass methods.
        """
        return 'ERROR: read method is not implemented.'

    def merge(self, novel):
        """Copy required attributes of the novel object.
        Return a message beginning with SUCCESS or ERROR.
        This is a stub to be overridden by subclass methods.
        """
        return 'ERROR: merge method is not implemented.'

    def write(self):
        """Write selected properties to the file.
        Return a message beginning with SUCCESS or ERROR.
        This is a stub to be overridden by subclass methods.
        """
        return 'ERROR: write method is not implemented.'

    def convert_to_yw(self, text):
        """Return text, converted from source format to yw7 markup.
        This is a stub to be overridden by subclass methods.
        """
        return text

    def convert_from_yw(self, text):
        """Return text, converted from yw7 markup to target format.
        This is a stub to be overridden by subclass methods.
        """
        return text

    def file_exists(self):
        """Return True, if the file specified by filePath exists. 
        Otherwise, return False.
        """
        if os.path.isfile(self.filePath):
            return True

        else:
            return False


class FileExport(Novel):
    """Abstract yWriter project file exporter representation.
    This class is generic and contains no conversion algorithm and no templates.
    """
    SUFFIX = ''

    fileHeader = ''
    partTemplate = ''
    chapterTemplate = ''
    notesChapterTemplate = ''
    todoChapterTemplate = ''
    unusedChapterTemplate = ''
    notExportedChapterTemplate = ''
    sceneTemplate = ''
    appendedSceneTemplate = ''
    notesSceneTemplate = ''
    todoSceneTemplate = ''
    unusedSceneTemplate = ''
    notExportedSceneTemplate = ''
    sceneDivider = ''
    chapterEndTemplate = ''
    unusedChapterEndTemplate = ''
    notExportedChapterEndTemplate = ''
    notesChapterEndTemplate = ''
    characterTemplate = ''
    locationTemplate = ''
    itemTemplate = ''
    fileFooter = ''

    def get_string(self, elements):
        """Return a string which is the concatenation of the 
        members of the list of strings "elements", separated by 
        a comma plus a space. The space allows word wrap in 
        spreadsheet cells.
        """
        text = (', ').join(elements)
        return text

    def get_list(self, text):
        """Split a sequence of strings into a list of strings
        using a comma as delimiter. Remove leading and trailing
        spaces, if any.
        """
        elements = []
        tempList = text.split(',')

        for element in tempList:
            elements.append(element.lstrip().rstrip())

        return elements

    def convert_from_yw(self, text):
        """Convert yw7 markup to target format.
        This is a stub to be overridden by subclass methods.
        """

        if text is None:
            text = ''

        return(text)

    def merge(self, novel):
        """Copy required attributes of the novel object.
        Return a message beginning with SUCCESS or ERROR.
        """

        if novel.title is not None:
            self.title = novel.title

        else:
            self.title = ''

        if novel.desc is not None:
            self.desc = novel.desc

        else:
            self.desc = ''

        if novel.author is not None:
            self.author = novel.author

        else:
            self.author = ''

        if novel.fieldTitle1 is not None:
            self.fieldTitle1 = novel.fieldTitle1

        else:
            self.fieldTitle1 = 'Field 1'

        if novel.fieldTitle2 is not None:
            self.fieldTitle2 = novel.fieldTitle2

        else:
            self.fieldTitle2 = 'Field 2'

        if novel.fieldTitle3 is not None:
            self.fieldTitle3 = novel.fieldTitle3

        else:
            self.fieldTitle3 = 'Field 3'

        if novel.fieldTitle4 is not None:
            self.fieldTitle4 = novel.fieldTitle4

        else:
            self.fieldTitle4 = 'Field 4'

        if novel.srtChapters != []:
            self.srtChapters = novel.srtChapters

        if novel.scenes is not None:
            self.scenes = novel.scenes

        if novel.chapters is not None:
            self.chapters = novel.chapters

        if novel.srtCharacters != []:
            self.srtCharacters = novel.srtCharacters
            self.characters = novel.characters

        if novel.srtLocations != []:
            self.srtLocations = novel.srtLocations
            self.locations = novel.locations

        if novel.srtItems != []:
            self.srtItems = novel.srtItems
            self.items = novel.items

        return 'SUCCESS'

    def get_fileHeaderMapping(self):
        """Return a mapping dictionary for the project section. 
        """
        projectTemplateMapping = dict(
            Title=self.title,
            Desc=self.convert_from_yw(self.desc),
            AuthorName=self.author,
            FieldTitle1=self.fieldTitle1,
            FieldTitle2=self.fieldTitle2,
            FieldTitle3=self.fieldTitle3,
            FieldTitle4=self.fieldTitle4,
        )
        return projectTemplateMapping

    def get_chapterMapping(self, chId, chapterNumber):
        """Return a mapping dictionary for a chapter section. 
        """
        chapterMapping = dict(
            ID=chId,
            ChapterNumber=chapterNumber,
            Title=self.chapters[chId].get_title(),
            Desc=self.convert_from_yw(self.chapters[chId].desc),
            ProjectName=self.projectName,
            ProjectPath=self.projectPath,
        )
        return chapterMapping

    def get_sceneMapping(self, scId, sceneNumber, wordsTotal, lettersTotal):
        """Return a mapping dictionary for a scene section. 
        """

        if self.scenes[scId].tags is not None:
            tags = self.get_string(self.scenes[scId].tags)

        else:
            tags = ''

        try:
            # Note: Due to a bug, yWriter scenes might hold invalid
            # viepoint characters
            sChList = []

            for chId in self.scenes[scId].characters:
                sChList.append(self.characters[chId].title)

            sceneChars = self.get_string(sChList)
            viewpointChar = sChList[0]

        except:
            sceneChars = ''
            viewpointChar = ''

        if self.scenes[scId].locations is not None:
            sLcList = []

            for lcId in self.scenes[scId].locations:
                sLcList.append(self.locations[lcId].title)

            sceneLocs = self.get_string(sLcList)

        else:
            sceneLocs = ''

        if self.scenes[scId].items is not None:
            sItList = []

            for itId in self.scenes[scId].items:
                sItList.append(self.items[itId].title)

            sceneItems = self.get_string(sItList)

        else:
            sceneItems = ''

        if self.scenes[scId].isReactionScene:
            reactionScene = Scene.REACTION_MARKER

        else:
            reactionScene = Scene.ACTION_MARKER

        sceneMapping = dict(
            ID=scId,
            SceneNumber=sceneNumber,
            Title=self.scenes[scId].title,
            Desc=self.convert_from_yw(self.scenes[scId].desc),
            WordCount=str(self.scenes[scId].wordCount),
            WordsTotal=wordsTotal,
            LetterCount=str(self.scenes[scId].letterCount),
            LettersTotal=lettersTotal,
            Status=Scene.STATUS[self.scenes[scId].status],
            SceneContent=self.convert_from_yw(
                self.scenes[scId].sceneContent),
            FieldTitle1=self.fieldTitle1,
            FieldTitle2=self.fieldTitle2,
            FieldTitle3=self.fieldTitle3,
            FieldTitle4=self.fieldTitle4,
            Field1=self.scenes[scId].field1,
            Field2=self.scenes[scId].field2,
            Field3=self.scenes[scId].field3,
            Field4=self.scenes[scId].field4,
            Date=self.scenes[scId].date,
            Time=self.scenes[scId].time,
            Day=self.scenes[scId].day,
            Hour=self.scenes[scId].hour,
            Minute=self.scenes[scId].minute,
            LastsDays=self.scenes[scId].lastsDays,
            LastsHours=self.scenes[scId].lastsHours,
            LastsMinutes=self.scenes[scId].lastsMinutes,
            ReactionScene=reactionScene,
            Goal=self.convert_from_yw(self.scenes[scId].goal),
            Conflict=self.convert_from_yw(self.scenes[scId].conflict),
            Outcome=self.convert_from_yw(self.scenes[scId].outcome),
            Tags=tags,
            Characters=sceneChars,
            Viewpoint=viewpointChar,
            Locations=sceneLocs,
            Items=sceneItems,
            Notes=self.convert_from_yw(self.scenes[scId].sceneNotes),
            ProjectName=self.projectName,
            ProjectPath=self.projectPath,
        )
        return sceneMapping

    def get_characterMapping(self, crId):
        """Return a mapping dictionary for a character section. 
        """

        if self.characters[crId].tags is not None:
            tags = self.get_string(self.characters[crId].tags)

        else:
            tags = ''

        if self.characters[crId].isMajor:
            characterStatus = Character.MAJOR_MARKER

        else:
            characterStatus = Character.MINOR_MARKER

        characterMapping = dict(
            ID=crId,
            Title=self.characters[crId].title,
            Desc=self.convert_from_yw(self.characters[crId].desc),
            Tags=tags,
            AKA=FileExport.convert_from_yw(self, self.characters[crId].aka),
            Notes=self.convert_from_yw(self.characters[crId].notes),
            Bio=self.convert_from_yw(self.characters[crId].bio),
            Goals=self.convert_from_yw(self.characters[crId].goals),
            FullName=FileExport.convert_from_yw(
                self, self.characters[crId].fullName),
            Status=characterStatus,
            ProjectName=self.projectName,
            ProjectPath=self.projectPath,
        )
        return characterMapping

    def get_locationMapping(self, lcId):
        """Return a mapping dictionary for a location section. 
        """

        if self.locations[lcId].tags is not None:
            tags = self.get_string(self.locations[lcId].tags)

        else:
            tags = ''

        locationMapping = dict(
            ID=lcId,
            Title=self.locations[lcId].title,
            Desc=self.convert_from_yw(self.locations[lcId].desc),
            Tags=tags,
            AKA=FileExport.convert_from_yw(self, self.locations[lcId].aka),
            ProjectName=self.projectName,
            ProjectPath=self.projectPath,
        )
        return locationMapping

    def get_itemMapping(self, itId):
        """Return a mapping dictionary for an item section. 
        """

        if self.items[itId].tags is not None:
            tags = self.get_string(self.items[itId].tags)

        else:
            tags = ''

        itemMapping = dict(
            ID=itId,
            Title=self.items[itId].title,
            Desc=self.convert_from_yw(self.items[itId].desc),
            Tags=tags,
            AKA=FileExport.convert_from_yw(self, self.items[itId].aka),
            ProjectName=self.projectName,
            ProjectPath=self.projectPath,
        )
        return itemMapping

    def get_fileHeader(self):
        """Process the file header.
        Return a list of strings.
        """
        lines = []
        template = Template(self.fileHeader)
        lines.append(template.safe_substitute(
            self.get_fileHeaderMapping()))
        return lines

    def get_scenes(self, chId, sceneNumber, wordsTotal, lettersTotal, doNotExport):
        """Process the scenes.
        Return a list of strings.
        """
        lines = []
        firstSceneInChapter = True

        for scId in self.chapters[chId].srtScenes:
            wordsTotal += self.scenes[scId].wordCount
            lettersTotal += self.scenes[scId].letterCount

            # The order counts; be aware that "Todo" and "Notes" scenes are
            # always unused.

            if self.scenes[scId].isTodoScene:

                if self.todoSceneTemplate != '':
                    template = Template(self.todoSceneTemplate)

                else:
                    continue

            elif self.scenes[scId].isNotesScene or self.chapters[chId].oldType == 1:
                # Scene is "Notes" (new file format) or "Info" (old file
                # format) scene.

                if self.notesSceneTemplate != '':
                    template = Template(self.notesSceneTemplate)

                else:
                    continue

            elif self.scenes[scId].isUnused or self.chapters[chId].isUnused:

                if self.unusedSceneTemplate != '':
                    template = Template(self.unusedSceneTemplate)

                else:
                    continue

            elif self.scenes[scId].doNotExport or doNotExport:

                if self.notExportedSceneTemplate != '':
                    template = Template(self.notExportedSceneTemplate)

                else:
                    continue

            else:
                sceneNumber += 1

                template = Template(self.sceneTemplate)

                if not firstSceneInChapter and self.scenes[scId].appendToPrev and self.appendedSceneTemplate != '':
                    template = Template(self.appendedSceneTemplate)

            if not (firstSceneInChapter or self.scenes[scId].appendToPrev):
                lines.append(self.sceneDivider)

            lines.append(template.safe_substitute(self.get_sceneMapping(
                scId, sceneNumber, wordsTotal, lettersTotal)))

            firstSceneInChapter = False

        return lines, sceneNumber, wordsTotal, lettersTotal

    def get_chapters(self):
        """Process the chapters and nested scenes.
        Return a list of strings.
        """
        lines = []
        chapterNumber = 0
        sceneNumber = 0
        wordsTotal = 0
        lettersTotal = 0

        for chId in self.srtChapters:

            # The order counts; be aware that "Todo" and "Notes" chapters are
            # always unused.

            # Has the chapter only scenes not to be exported?

            sceneCount = 0
            notExportCount = 0
            doNotExport = False

            for scId in self.chapters[chId].srtScenes:
                sceneCount += 1

                if self.scenes[scId].doNotExport:
                    notExportCount += 1

            if sceneCount > 0 and notExportCount == sceneCount:
                doNotExport = True

            if self.chapters[chId].chType == 2:

                if self.todoChapterTemplate != '':
                    template = Template(self.todoChapterTemplate)

                else:
                    continue

            elif self.chapters[chId].chType == 1 or self.chapters[chId].oldType == 1:
                # Chapter is "Notes" (new file format) or "Info" (old file
                # format) chapter.

                if self.notesChapterTemplate != '':
                    template = Template(self.notesChapterTemplate)

                else:
                    continue

            elif self.chapters[chId].isUnused:

                if self.unusedChapterTemplate != '':
                    template = Template(self.unusedChapterTemplate)

                else:
                    continue

            elif doNotExport:

                if self.notExportedChapterTemplate != '':
                    template = Template(self.notExportedChapterTemplate)

                else:
                    continue

            elif self.chapters[chId].chLevel == 1 and self.partTemplate != '':
                template = Template(self.partTemplate)

            else:
                template = Template(self.chapterTemplate)
                chapterNumber += 1

            lines.append(template.safe_substitute(
                self.get_chapterMapping(chId, chapterNumber)))

            # Process scenes.

            sceneLines, sceneNumber, wordsTotal, lettersTotal = self.get_scenes(
                chId, sceneNumber, wordsTotal, lettersTotal, doNotExport)
            lines.extend(sceneLines)

            # Process chapter ending.

            if self.chapters[chId].chType == 2 and self.todoChapterEndTemplate != '':
                lines.append(self.todoChapterEndTemplate)

            elif self.chapters[chId].chType == 1 or self.chapters[chId].oldType == 1:

                if self.notesChapterEndTemplate != '':
                    lines.append(self.notesChapterEndTemplate)

            elif self.chapters[chId].isUnused and self.unusedChapterEndTemplate != '':
                lines.append(self.unusedChapterEndTemplate)

            elif doNotExport and self.notExportedChapterEndTemplate != '':
                lines.append(self.notExportedChapterEndTemplate)

            elif self.chapterEndTemplate != '':
                lines.append(self.chapterEndTemplate)

        return lines

    def get_characters(self):
        """Process the characters.
        Return a list of strings.
        """
        lines = []
        template = Template(self.characterTemplate)

        for crId in self.srtCharacters:
            lines.append(template.safe_substitute(
                self.get_characterMapping(crId)))

        return lines

    def get_locations(self):
        """Process the locations.
        Return a list of strings.
        """
        lines = []
        template = Template(self.locationTemplate)

        for lcId in self.srtLocations:
            lines.append(template.safe_substitute(
                self.get_locationMapping(lcId)))

        return lines

    def get_items(self):
        """Process the items.
        Return a list of strings.
        """
        lines = []
        template = Template(self.itemTemplate)

        for itId in self.srtItems:
            lines.append(template.safe_substitute(self.get_itemMapping(itId)))

        return lines

    def get_text(self):
        """Assemple the whole text applying the templates.
        Return a string to be written to the output file.
        """
        lines = self.get_fileHeader()
        lines.extend(self.get_chapters())
        lines.extend(self.get_characters())
        lines.extend(self.get_locations())
        lines.extend(self.get_items())
        lines.append(self.fileFooter)
        return ''.join(lines)

    def write(self):
        """Create a template-based output file. 
        Return a message string starting with 'SUCCESS' or 'ERROR'.
        """
        text = self.get_text()

        try:
            with open(self.filePath, 'w', encoding='utf-8') as f:
                f.write(text)

        except:
            return 'ERROR: Cannot write "' + os.path.normpath(self.filePath) + '".'

        return 'SUCCESS: "' + os.path.normpath(self.filePath) + '" written.'


class OdfFile(FileExport):
    """Generic OpenDocument xml file representation.
    """
    TEMPDIR = 'temp_odf'

    ODF_COMPONENTS = []
    _MIMETYPE = ''
    _SETTINGS_XML = ''
    _MANIFEST_XML = ''
    _STYLES_XML = ''
    _META_XML = ''

    def tear_down(self):
        """Delete the temporary directory 
        containing the unpacked ODF directory structure.
        """
        try:
            rmtree(self.TEMPDIR)
        except:
            pass

    def set_up(self):
        """Helper method for ZIP file generation.

        Create a temporary directory containing the internal 
        structure of an ODF file except 'content.xml'.
        """
        self.tear_down()
        os.mkdir(self.TEMPDIR)
        os.mkdir(self.TEMPDIR + '/META-INF')

        # Generate mimetype

        try:
            with open(self.TEMPDIR + '/mimetype', 'w', encoding='utf-8') as f:
                f.write(self._MIMETYPE)
        except:
            return 'ERROR: Cannot write "mimetype"'

        # Generate settings.xml

        try:
            with open(self.TEMPDIR + '/settings.xml', 'w', encoding='utf-8') as f:
                f.write(self._SETTINGS_XML)
        except:
            return 'ERROR: Cannot write "settings.xml"'

        # Generate META-INF\manifest.xml

        try:
            with open(self.TEMPDIR + '/META-INF/manifest.xml', 'w', encoding='utf-8') as f:
                f.write(self._MANIFEST_XML)
        except:
            return 'ERROR: Cannot write "manifest.xml"'

        # Generate styles.xml with system language set as document language

        localeCodes = locale.getdefaultlocale()[0].split('_')

        localeMapping = dict(
            Language=localeCodes[0],
            Country=localeCodes[1],
        )
        template = Template(self._STYLES_XML)
        text = template.safe_substitute(localeMapping)

        try:
            with open(self.TEMPDIR + '/styles.xml', 'w', encoding='utf-8') as f:
                f.write(text)
        except:
            return 'ERROR: Cannot write "styles.xml"'

        # Generate meta.xml with actual document metadata

        dt = datetime.today()

        metaMapping = dict(
            Author=self.author,
            Title=self.title,
            Summary='<![CDATA[' + self.desc + ']]>',
            Date=str(dt.year) + '-' + str(dt.month).rjust(2, '0') +
            '-' + str(dt.day).rjust(2, '0'),
            Time=str(dt.hour).rjust(2, '0') +
            ':' + str(dt.minute).rjust(2, '0') +
            ':' + str(dt.second).rjust(2, '0'),
        )
        template = Template(self._META_XML)
        text = template.safe_substitute(metaMapping)

        try:
            with open(self.TEMPDIR + '/meta.xml', 'w', encoding='utf-8') as f:
                f.write(text)
        except:
            return 'ERROR: Cannot write "meta.xml".'

        return 'SUCCESS: ODF structure generated.'

    def write(self):
        """Extend the super class method, adding ZIP file operations."""

        # Create a temporary directory containing the internal
        # structure of an ODS file except "content.xml".

        message = self.set_up()

        if message.startswith('ERROR'):
            return message

        # Add "content.xml" to the temporary directory.

        filePath = self._filePath

        self._filePath = self.TEMPDIR + '/content.xml'

        message = FileExport.write(self)

        self._filePath = filePath

        if message.startswith('ERROR'):
            return message

        # Pack the contents of the temporary directory
        # into the ODF file.

        workdir = os.getcwd()

        try:
            with zipfile.ZipFile(self.filePath, 'w') as odfTarget:
                os.chdir(self.TEMPDIR)

                for file in self.ODF_COMPONENTS:
                    odfTarget.write(file, compress_type=zipfile.ZIP_DEFLATED)
        except:
            os.chdir(workdir)
            return 'ERROR: Cannot generate "' + os.path.normpath(self.filePath) + '".'

        # Remove temporary data.

        os.chdir(workdir)
        self.tear_down()
        return 'SUCCESS: "' + os.path.normpath(self.filePath) + '" written.'


class OdtFile(OdfFile):
    """Generic OpenDocument text document representation."""

    EXTENSION = '.odt'
    # overwrites Novel.EXTENSION

    ODF_COMPONENTS = ['manifest.rdf', 'META-INF', 'content.xml', 'meta.xml', 'mimetype',
                      'settings.xml', 'styles.xml', 'META-INF/manifest.xml']

    CONTENT_XML_HEADER = '''<?xml version="1.0" encoding="UTF-8"?>

<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" xmlns:math="http://www.w3.org/1998/Math/MathML" xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" xmlns:ooo="http://openoffice.org/2004/office" xmlns:ooow="http://openoffice.org/2004/writer" xmlns:oooc="http://openoffice.org/2004/calc" xmlns:dom="http://www.w3.org/2001/xml-events" xmlns:xforms="http://www.w3.org/2002/xforms" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:rpt="http://openoffice.org/2005/report" xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2" xmlns:xhtml="http://www.w3.org/1999/xhtml" xmlns:grddl="http://www.w3.org/2003/g/data-view#" xmlns:tableooo="http://openoffice.org/2009/table" xmlns:field="urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0" office:version="1.2">
 <office:scripts/>
 <office:font-face-decls>
  <style:font-face style:name="StarSymbol" svg:font-family="StarSymbol" style:font-charset="x-symbol"/>
  <style:font-face style:name="Courier New" svg:font-family="&apos;Courier New&apos;" style:font-adornments="Standard" style:font-family-generic="modern" style:font-pitch="fixed"/>
   </office:font-face-decls>
 <office:automatic-styles>
  <style:style style:name="Sect1" style:family="section">
   <style:section-properties style:editable="false">
    <style:columns fo:column-count="1" fo:column-gap="0cm"/>
   </style:section-properties>
  </style:style>
 </office:automatic-styles>
 <office:body>
  <office:text text:use-soft-page-breaks="true">

'''

    CONTENT_XML_FOOTER = '''  </office:text>
 </office:body>
</office:document-content>
'''

    _META_XML = '''<?xml version="1.0" encoding="utf-8"?>
<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:ooo="http://openoffice.org/2004/office" xmlns:grddl="http://www.w3.org/2003/g/data-view#" office:version="1.2">
  <office:meta>
    <meta:generator>PyWriter</meta:generator>
    <dc:title>$Title</dc:title>
    <dc:description>$Summary</dc:description>
    <dc:subject></dc:subject>
    <meta:keyword></meta:keyword>
    <meta:initial-creator>$Author</meta:initial-creator>
    <dc:creator></dc:creator>
    <meta:creation-date>${Date}T${Time}Z</meta:creation-date>
    <dc:date></dc:date>
  </office:meta>
</office:document-meta>
'''
    _MANIFEST_XML = '''<?xml version="1.0" encoding="utf-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">
  <manifest:file-entry manifest:media-type="application/vnd.oasis.opendocument.text" manifest:full-path="/" />
  <manifest:file-entry manifest:media-type="application/xml" manifest:full-path="content.xml" manifest:version="1.2" />
  <manifest:file-entry manifest:media-type="application/rdf+xml" manifest:full-path="manifest.rdf" manifest:version="1.2" />
  <manifest:file-entry manifest:media-type="application/xml" manifest:full-path="styles.xml" manifest:version="1.2" />
  <manifest:file-entry manifest:media-type="application/xml" manifest:full-path="meta.xml" manifest:version="1.2" />
  <manifest:file-entry manifest:media-type="application/xml" manifest:full-path="settings.xml" manifest:version="1.2" />
</manifest:manifest>    
'''
    _MANIFEST_RDF = '''<?xml version="1.0" encoding="utf-8"?>
<rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#">
  <rdf:Description rdf:about="styles.xml">
    <rdf:type rdf:resource="http://docs.oasis-open.org/ns/office/1.2/meta/odf#StylesFile"/>
  </rdf:Description>
  <rdf:Description rdf:about="">
    <ns0:hasPart xmlns:ns0="http://docs.oasis-open.org/ns/office/1.2/meta/pkg#" rdf:resource="styles.xml"/>
  </rdf:Description>
  <rdf:Description rdf:about="content.xml">
    <rdf:type rdf:resource="http://docs.oasis-open.org/ns/office/1.2/meta/odf#ContentFile"/>
  </rdf:Description>
  <rdf:Description rdf:about="">
    <ns0:hasPart xmlns:ns0="http://docs.oasis-open.org/ns/office/1.2/meta/pkg#" rdf:resource="content.xml"/>
  </rdf:Description>
  <rdf:Description rdf:about="">
    <rdf:type rdf:resource="http://docs.oasis-open.org/ns/office/1.2/meta/pkg#Document"/>
  </rdf:Description>
</rdf:RDF>
'''
    _SETTINGS_XML = '''<?xml version="1.0" encoding="UTF-8"?>

<office:document-settings xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0" xmlns:ooo="http://openoffice.org/2004/office" office:version="1.2">
 <office:settings>
  <config:config-item-set config:name="ooo:view-settings">
   <config:config-item config:name="ViewAreaTop" config:type="int">0</config:config-item>
   <config:config-item config:name="ViewAreaLeft" config:type="int">0</config:config-item>
   <config:config-item config:name="ViewAreaWidth" config:type="int">30508</config:config-item>
   <config:config-item config:name="ViewAreaHeight" config:type="int">27783</config:config-item>
   <config:config-item config:name="ShowRedlineChanges" config:type="boolean">true</config:config-item>
   <config:config-item config:name="InBrowseMode" config:type="boolean">false</config:config-item>
   <config:config-item-map-indexed config:name="Views">
    <config:config-item-map-entry>
     <config:config-item config:name="ViewId" config:type="string">view2</config:config-item>
     <config:config-item config:name="ViewLeft" config:type="int">8079</config:config-item>
     <config:config-item config:name="ViewTop" config:type="int">3501</config:config-item>
     <config:config-item config:name="VisibleLeft" config:type="int">0</config:config-item>
     <config:config-item config:name="VisibleTop" config:type="int">0</config:config-item>
     <config:config-item config:name="VisibleRight" config:type="int">30506</config:config-item>
     <config:config-item config:name="VisibleBottom" config:type="int">27781</config:config-item>
     <config:config-item config:name="ZoomType" config:type="short">0</config:config-item>
     <config:config-item config:name="ViewLayoutColumns" config:type="short">0</config:config-item>
     <config:config-item config:name="ViewLayoutBookMode" config:type="boolean">false</config:config-item>
     <config:config-item config:name="ZoomFactor" config:type="short">100</config:config-item>
     <config:config-item config:name="IsSelectedFrame" config:type="boolean">false</config:config-item>
    </config:config-item-map-entry>
   </config:config-item-map-indexed>
  </config:config-item-set>
  <config:config-item-set config:name="ooo:configuration-settings">
   <config:config-item config:name="AddParaSpacingToTableCells" config:type="boolean">true</config:config-item>
   <config:config-item config:name="PrintPaperFromSetup" config:type="boolean">false</config:config-item>
   <config:config-item config:name="IsKernAsianPunctuation" config:type="boolean">false</config:config-item>
   <config:config-item config:name="PrintReversed" config:type="boolean">false</config:config-item>
   <config:config-item config:name="LinkUpdateMode" config:type="short">1</config:config-item>
   <config:config-item config:name="DoNotCaptureDrawObjsOnPage" config:type="boolean">false</config:config-item>
   <config:config-item config:name="SaveVersionOnClose" config:type="boolean">false</config:config-item>
   <config:config-item config:name="PrintEmptyPages" config:type="boolean">true</config:config-item>
   <config:config-item config:name="PrintSingleJobs" config:type="boolean">false</config:config-item>
   <config:config-item config:name="AllowPrintJobCancel" config:type="boolean">true</config:config-item>
   <config:config-item config:name="AddFrameOffsets" config:type="boolean">false</config:config-item>
   <config:config-item config:name="PrintLeftPages" config:type="boolean">true</config:config-item>
   <config:config-item config:name="PrintTables" config:type="boolean">true</config:config-item>
   <config:config-item config:name="ProtectForm" config:type="boolean">false</config:config-item>
   <config:config-item config:name="ChartAutoUpdate" config:type="boolean">true</config:config-item>
   <config:config-item config:name="PrintControls" config:type="boolean">true</config:config-item>
   <config:config-item config:name="PrinterSetup" config:type="base64Binary">8gT+/0hQIExhc2VySmV0IFAyMDE0AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAASFAgTGFzZXJKZXQgUDIwMTQAAAAAAAAAAAAAAAAAAAAWAAEAGAQAAAAAAAAEAAhSAAAEdAAAM1ROVwIACABIAFAAIABMAGEAcwBlAHIASgBlAHQAIABQADIAMAAxADQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQQDANwANAMPnwAAAQAJAJoLNAgAAAEABwBYAgEAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAU0RETQAGAAAABgAASFAgTGFzZXJKZXQgUDIwMTQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIAAAAAAAAAAAAAAAEAAAAJAAAACQAAAAkAAAAJAAAACQAAAAkAAAAJAAAACQAAAAkAAAAJAAAACQAAAAkAAAAJAAAACQAAAAkAAAAJAAAACQAAAAAAAAABAAAAAQAAABoEAAAAAAAAAAAAAAAAAAAPAAAALQAAAAAAAAAAAAAAAQAAAAAAAAAAAAAAgICAAP8AAAD//wAAAP8AAAD//wAAAP8A/wD/AAAAAAAAAAAAAAAAAAAAAAAoAAAAZAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADeAwAA3gMAAAAAAAAAAAAAAIAAAAAAAAAAAAAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABrjvBgNAMAAAAAAAAAAAAABIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABIAQ09NUEFUX0RVUExFWF9NT0RFCgBEVVBMRVhfT0ZG</config:config-item>
   <config:config-item config:name="CurrentDatabaseDataSource" config:type="string"/>
   <config:config-item config:name="LoadReadonly" config:type="boolean">false</config:config-item>
   <config:config-item config:name="CurrentDatabaseCommand" config:type="string"/>
   <config:config-item config:name="ConsiderTextWrapOnObjPos" config:type="boolean">false</config:config-item>
   <config:config-item config:name="ApplyUserData" config:type="boolean">true</config:config-item>
   <config:config-item config:name="AddParaTableSpacing" config:type="boolean">true</config:config-item>
   <config:config-item config:name="FieldAutoUpdate" config:type="boolean">true</config:config-item>
   <config:config-item config:name="IgnoreFirstLineIndentInNumbering" config:type="boolean">false</config:config-item>
   <config:config-item config:name="TabsRelativeToIndent" config:type="boolean">true</config:config-item>
   <config:config-item config:name="IgnoreTabsAndBlanksForLineCalculation" config:type="boolean">false</config:config-item>
   <config:config-item config:name="PrintAnnotationMode" config:type="short">0</config:config-item>
   <config:config-item config:name="AddParaTableSpacingAtStart" config:type="boolean">true</config:config-item>
   <config:config-item config:name="UseOldPrinterMetrics" config:type="boolean">false</config:config-item>
   <config:config-item config:name="TableRowKeep" config:type="boolean">false</config:config-item>
   <config:config-item config:name="PrinterName" config:type="string">HP LaserJet P2014</config:config-item>
   <config:config-item config:name="PrintFaxName" config:type="string"/>
   <config:config-item config:name="UnxForceZeroExtLeading" config:type="boolean">false</config:config-item>
   <config:config-item config:name="PrintTextPlaceholder" config:type="boolean">false</config:config-item>
   <config:config-item config:name="DoNotJustifyLinesWithManualBreak" config:type="boolean">false</config:config-item>
   <config:config-item config:name="PrintRightPages" config:type="boolean">true</config:config-item>
   <config:config-item config:name="CharacterCompressionType" config:type="short">0</config:config-item>
   <config:config-item config:name="UseFormerTextWrapping" config:type="boolean">false</config:config-item>
   <config:config-item config:name="IsLabelDocument" config:type="boolean">false</config:config-item>
   <config:config-item config:name="AlignTabStopPosition" config:type="boolean">true</config:config-item>
   <config:config-item config:name="PrintHiddenText" config:type="boolean">false</config:config-item>
   <config:config-item config:name="DoNotResetParaAttrsForNumFont" config:type="boolean">false</config:config-item>
   <config:config-item config:name="PrintPageBackground" config:type="boolean">true</config:config-item>
   <config:config-item config:name="CurrentDatabaseCommandType" config:type="int">0</config:config-item>
   <config:config-item config:name="OutlineLevelYieldsNumbering" config:type="boolean">false</config:config-item>
   <config:config-item config:name="PrintProspect" config:type="boolean">false</config:config-item>
   <config:config-item config:name="PrintGraphics" config:type="boolean">true</config:config-item>
   <config:config-item config:name="SaveGlobalDocumentLinks" config:type="boolean">false</config:config-item>
   <config:config-item config:name="PrintProspectRTL" config:type="boolean">false</config:config-item>
   <config:config-item config:name="UseFormerLineSpacing" config:type="boolean">false</config:config-item>
   <config:config-item config:name="AddExternalLeading" config:type="boolean">true</config:config-item>
   <config:config-item config:name="UseFormerObjectPositioning" config:type="boolean">false</config:config-item>
   <config:config-item config:name="RedlineProtectionKey" config:type="base64Binary"/>
   <config:config-item config:name="MathBaselineAlignment" config:type="boolean">false</config:config-item>
   <config:config-item config:name="ClipAsCharacterAnchoredWriterFlyFrames" config:type="boolean">false</config:config-item>
   <config:config-item config:name="UseOldNumbering" config:type="boolean">false</config:config-item>
   <config:config-item config:name="PrintDrawings" config:type="boolean">true</config:config-item>
   <config:config-item config:name="PrinterIndependentLayout" config:type="string">disabled</config:config-item>
   <config:config-item config:name="TabAtLeftIndentForParagraphsInList" config:type="boolean">false</config:config-item>
   <config:config-item config:name="PrintBlackFonts" config:type="boolean">false</config:config-item>
   <config:config-item config:name="UpdateFromTemplate" config:type="boolean">true</config:config-item>
  </config:config-item-set>
 </office:settings>
</office:document-settings>
'''
    _STYLES_XML = '''<?xml version="1.0" encoding="UTF-8"?>

<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" xmlns:math="http://www.w3.org/1998/Math/MathML" xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" xmlns:ooo="http://openoffice.org/2004/office" xmlns:ooow="http://openoffice.org/2004/writer" xmlns:oooc="http://openoffice.org/2004/calc" xmlns:dom="http://www.w3.org/2001/xml-events" xmlns:rpt="http://openoffice.org/2005/report" xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2" xmlns:xhtml="http://www.w3.org/1999/xhtml" xmlns:grddl="http://www.w3.org/2003/g/data-view#" xmlns:tableooo="http://openoffice.org/2009/table" office:version="1.2">
 <office:font-face-decls>
  <style:font-face style:name="StarSymbol" svg:font-family="StarSymbol" style:font-charset="x-symbol"/>
  <style:font-face style:name="Arial" svg:font-family="&apos;Arial&apos;" style:font-family-generic="swiss"/>
  <style:font-face style:name="Courier New" svg:font-family="&apos;Courier New&apos;" style:font-adornments="Standard" style:font-family-generic="modern" style:font-pitch="fixed"/>
  <style:font-face style:name="Segoe UI" svg:font-family="&apos;Segoe UI&apos;" style:font-family-generic="roman" style:font-pitch="variable"/>
 </office:font-face-decls>
 <office:styles>
  <style:default-style style:family="graphic">
   <style:graphic-properties fo:wrap-option="no-wrap" draw:shadow-offset-x="0.3cm" draw:shadow-offset-y="0.3cm" draw:start-line-spacing-horizontal="0.283cm" draw:start-line-spacing-vertical="0.283cm" draw:end-line-spacing-horizontal="0.283cm" draw:end-line-spacing-vertical="0.283cm" style:flow-with-text="true"/>
   <style:paragraph-properties style:text-autospace="ideograph-alpha" style:line-break="strict" style:writing-mode="lr-tb" style:font-independent-line-spacing="false">
    <style:tab-stops/>
   </style:paragraph-properties>
   <style:text-properties fo:color="#000000" fo:font-size="10pt" fo:language="$Language" fo:country="$Country" style:letter-kerning="true" style:font-size-asian="10pt" style:language-asian="zxx" style:country-asian="none" style:font-size-complex="1pt" style:language-complex="zxx" style:country-complex="none"/>
  </style:default-style>
  <style:default-style style:family="paragraph">
   <style:paragraph-properties fo:hyphenation-ladder-count="no-limit" style:text-autospace="ideograph-alpha" style:punctuation-wrap="hanging" style:line-break="strict" style:tab-stop-distance="1.251cm" style:writing-mode="lr-tb"/>
   <style:text-properties fo:color="#000000" style:font-name="Segoe UI" fo:font-size="10pt" fo:language="$Language" fo:country="$Country" style:letter-kerning="true" style:font-name-asian="Segoe UI" style:font-size-asian="10pt" style:language-asian="zxx" style:country-asian="none" style:font-name-complex="Segoe UI" style:font-size-complex="1pt" style:language-complex="zxx" style:country-complex="none" fo:hyphenate="false" fo:hyphenation-remain-char-count="2" fo:hyphenation-push-char-count="2"/>
  </style:default-style>
  <style:default-style style:family="table">
   <style:table-properties table:border-model="separating"/>
  </style:default-style>
  <style:default-style style:family="table-row">
   <style:table-row-properties fo:keep-together="always"/>
  </style:default-style>
  <style:style style:name="Standard" style:family="paragraph" style:class="text" style:master-page-name="">
   <style:paragraph-properties fo:line-height="0.73cm" style:page-number="auto"/>
   <style:text-properties style:font-name="Courier New" fo:font-size="12pt" fo:font-weight="normal"/>
  </style:style>
  <style:style style:name="Heading" style:family="paragraph" style:parent-style-name="Standard" style:next-style-name="Text_20_body" style:class="text" style:master-page-name="">
   <style:paragraph-properties fo:line-height="0.73cm" fo:text-align="center" style:justify-single-word="false" style:page-number="auto" fo:keep-with-next="always">
    <style:tab-stops/>
   </style:paragraph-properties>
  </style:style>
  <style:style style:name="Text_20_body" style:display-name="Text body" style:family="paragraph" style:parent-style-name="Standard" style:next-style-name="First_20_line_20_indent" style:class="text" style:master-page-name="">
   <style:paragraph-properties style:page-number="auto">
    <style:tab-stops/>
   </style:paragraph-properties>
  </style:style>
  <style:style style:name="List" style:family="paragraph" style:parent-style-name="Text_20_body" style:class="list">
   <style:text-properties fo:font-weight="normal"/>
  </style:style>
  <style:style style:name="Caption" style:family="paragraph" style:parent-style-name="Standard" style:class="extra">
   <style:paragraph-properties fo:margin-top="0.212cm" fo:margin-bottom="0.212cm"/>
  </style:style>
  <style:style style:name="Table" style:family="paragraph" style:parent-style-name="Caption" style:class="extra"/>
  <style:style style:name="Index" style:family="paragraph" style:parent-style-name="Standard" style:class="index"/>
  <style:style style:name="Heading_20_1" style:display-name="Heading 1" style:family="paragraph" style:parent-style-name="Heading" style:next-style-name="Text_20_body" style:default-outline-level="1" style:list-style-name="" style:class="text" style:master-page-name="">
   <style:paragraph-properties fo:margin-top="0.73cm" fo:margin-bottom="0.73cm" style:page-number="auto">
    <style:tab-stops/>
   </style:paragraph-properties>
   <style:text-properties fo:text-transform="uppercase" fo:font-weight="bold"/>
  </style:style>
  <style:style style:name="Heading_20_2" style:display-name="Heading 2" style:family="paragraph" style:parent-style-name="Heading" style:next-style-name="Text_20_body" style:default-outline-level="2" style:list-style-name="" style:class="text" style:master-page-name="">
   <style:paragraph-properties fo:margin-top="0.73cm" fo:margin-bottom="0.73cm" style:page-number="auto"/>
   <style:text-properties fo:font-weight="bold"/>
  </style:style>
  <style:style style:name="Heading_20_3" style:display-name="Heading 3" style:family="paragraph" style:parent-style-name="Heading" style:next-style-name="Text_20_body" style:default-outline-level="3" style:class="text" style:master-page-name="">
   <style:paragraph-properties fo:margin-top="0.73cm" fo:margin-bottom="0.73cm" style:page-number="auto"/>
   <style:text-properties fo:font-style="italic"/>
  </style:style>
  <style:style style:name="Heading_20_4" style:display-name="Heading 4" style:family="paragraph" style:parent-style-name="Heading" style:next-style-name="Text_20_body" style:default-outline-level="" style:list-style-name="" style:class="text" style:master-page-name="">
   <style:paragraph-properties fo:margin-top="0.73cm" fo:margin-bottom="0.73cm" style:page-number="auto"/>
  </style:style>
  <style:style style:name="Heading_20_5" style:display-name="Heading 5" style:family="paragraph" style:parent-style-name="Heading" style:next-style-name="Text_20_body" style:default-outline-level="5" style:list-style-name="" style:class="text"/>
  <style:style style:name="Heading_20_6" style:display-name="Heading 6" style:family="paragraph" style:parent-style-name="Heading" style:next-style-name="Text_20_body" style:default-outline-level="6" style:list-style-name="" style:class="text"/>
  <style:style style:name="Quotations" style:family="paragraph" style:parent-style-name="Text_20_body" style:class="html" style:master-page-name="">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="0.499cm" fo:margin-right="0.499cm" fo:margin-top="0cm" fo:margin-bottom="0.499cm" fo:text-indent="0cm" style:auto-text-indent="false" style:page-number="auto"/>
  </style:style>
  <style:style style:name="Preformatted_20_Text" style:display-name="Preformatted Text" style:family="paragraph" style:parent-style-name="Standard" style:class="html">
   <style:paragraph-properties fo:margin-top="0cm" fo:margin-bottom="0cm"/>
  </style:style>
  <style:style style:name="Table_20_Contents" style:display-name="Table Contents" style:family="paragraph" style:parent-style-name="Text_20_body" style:class="extra"/>
  <style:style style:name="Table_20_Heading" style:display-name="Table Heading" style:family="paragraph" style:parent-style-name="Table_20_Contents" style:class="extra">
   <style:paragraph-properties fo:text-align="center" style:justify-single-word="false"/>
   <style:text-properties fo:font-style="italic" fo:font-weight="bold"/>
  </style:style>
  <style:style style:name="Footnote" style:family="paragraph" style:parent-style-name="Standard" style:class="extra">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="0.499cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="-0.499cm" style:auto-text-indent="false"/>
   <style:text-properties fo:font-size="10pt"/>
  </style:style>
  <style:style style:name="Footer" style:family="paragraph" style:parent-style-name="Standard" style:class="extra" style:master-page-name="">
   <style:paragraph-properties fo:text-align="center" style:justify-single-word="false" style:page-number="auto" text:number-lines="false" text:line-number="0">
    <style:tab-stops>
     <style:tab-stop style:position="8.5cm" style:type="center"/>
     <style:tab-stop style:position="17.002cm" style:type="right"/>
    </style:tab-stops>
   </style:paragraph-properties>
   <style:text-properties fo:font-size="11pt"/>
  </style:style>
  <style:style style:name="Horizontal_20_Line" style:display-name="Horizontal Line" style:family="paragraph" style:parent-style-name="Standard" style:next-style-name="Text_20_body" style:class="html">
   <style:paragraph-properties fo:margin-top="0cm" fo:margin-bottom="0.499cm" style:border-line-width-bottom="0.002cm 0.035cm 0.002cm" fo:padding="0cm" fo:border-left="none" fo:border-right="none" fo:border-top="none" fo:border-bottom="0.039cm double #808080" text:number-lines="false" text:line-number="0"/>
   <style:text-properties fo:font-size="6pt"/>
  </style:style>
  <style:style style:name="First_20_line_20_indent" style:display-name="First line indent" style:family="paragraph" style:parent-style-name="Text_20_body" style:class="text" style:master-page-name="">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="0cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0.499cm" style:auto-text-indent="false" style:page-number="auto"/>
  </style:style>
  <style:style style:name="Heading_20_7" style:display-name="Heading 7" style:family="paragraph" style:parent-style-name="Heading" style:next-style-name="Text_20_body" style:default-outline-level="7" style:list-style-name="" style:class="text"/>
  <style:style style:name="Heading_20_8" style:display-name="Heading 8" style:family="paragraph" style:parent-style-name="Heading" style:next-style-name="Text_20_body" style:default-outline-level="8" style:list-style-name="" style:class="text"/>
  <style:style style:name="Heading_20_9" style:display-name="Heading 9" style:family="paragraph" style:parent-style-name="Heading" style:next-style-name="Text_20_body" style:default-outline-level="9" style:list-style-name="" style:class="text"/>
  <style:style style:name="Heading_20_10" style:display-name="Heading 10" style:family="paragraph" style:parent-style-name="Heading" style:next-style-name="Text_20_body" style:default-outline-level="10" style:list-style-name="" style:class="text">
   <style:text-properties fo:font-size="75%" fo:font-weight="bold"/>
  </style:style>
  <style:style style:name="Title" style:family="paragraph" style:parent-style-name="Standard" style:next-style-name="Subtitle" style:class="chapter" style:master-page-name="">
   <style:paragraph-properties fo:margin-top="0cm" fo:margin-bottom="0.73cm" fo:text-align="center" style:justify-single-word="false" style:page-number="auto" fo:background-color="transparent" fo:padding="0cm" fo:border="none" text:number-lines="false" text:line-number="0">
    <style:tab-stops/>
    <style:background-image/>
   </style:paragraph-properties>
   <style:text-properties fo:text-transform="uppercase" fo:letter-spacing="normal" fo:font-weight="normal" style:letter-kerning="false"/>
  </style:style>
  <style:style style:name="Subtitle" style:family="paragraph" style:parent-style-name="Title" style:class="chapter" style:master-page-name="">
   <style:paragraph-properties fo:margin-top="0cm" fo:margin-bottom="0.73cm" style:page-number="auto"/>
   <style:text-properties fo:font-variant="normal" fo:text-transform="none" fo:letter-spacing="normal"/>
  </style:style>
  <style:style style:name="Hanging_20_indent" style:display-name="Hanging indent" style:family="paragraph" style:parent-style-name="Text_20_body" style:class="text">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="1cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="-0.499cm" style:auto-text-indent="false">
    <style:tab-stops>
     <style:tab-stop style:position="0cm"/>
    </style:tab-stops>
   </style:paragraph-properties>
  </style:style>
  <style:style style:name="Text_20_body_20_indent" style:display-name="Text body indent" style:family="paragraph" style:parent-style-name="Text_20_body" style:class="text">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="0.499cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="Salutation" style:family="paragraph" style:parent-style-name="Standard" style:class="text"/>
  <style:style style:name="Signature" style:family="paragraph" style:parent-style-name="Standard" style:class="text"/>
  <style:style style:name="List_20_Indent" style:display-name="List Indent" style:family="paragraph" style:parent-style-name="Text_20_body" style:class="text">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="5.001cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="-4.5cm" style:auto-text-indent="false">
    <style:tab-stops>
     <style:tab-stop style:position="0cm"/>
    </style:tab-stops>
   </style:paragraph-properties>
  </style:style>
  <style:style style:name="Marginalia" style:family="paragraph" style:parent-style-name="Text_20_body" style:class="text">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="4.001cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="Numbering_20_1_20_Start" style:display-name="Numbering 1 Start" style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="0.499cm" fo:margin-right="0cm" fo:margin-top="0.423cm" fo:margin-bottom="0.212cm" fo:text-indent="-0.499cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="Numbering_20_1" style:display-name="Numbering 1" style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="0.499cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0.212cm" fo:text-indent="-0.499cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="Numbering_20_1_20_End" style:display-name="Numbering 1 End" style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="0.499cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0.423cm" fo:text-indent="-0.499cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="Numbering_20_1_20_Cont." style:display-name="Numbering 1 Cont." style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="0.499cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0.212cm" fo:text-indent="0cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="Numbering_20_2_20_Start" style:display-name="Numbering 2 Start" style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="0.998cm" fo:margin-right="0cm" fo:margin-top="0.423cm" fo:margin-bottom="0.212cm" fo:text-indent="-0.499cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="Numbering_20_2" style:display-name="Numbering 2" style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="0.998cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0.212cm" fo:text-indent="-0.499cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="Numbering_20_2_20_End" style:display-name="Numbering 2 End" style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="0.998cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0.423cm" fo:text-indent="-0.499cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="Numbering_20_2_20_Cont." style:display-name="Numbering 2 Cont." style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="0.998cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0.212cm" fo:text-indent="0cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="Numbering_20_3_20_Start" style:display-name="Numbering 3 Start" style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="1.498cm" fo:margin-right="0cm" fo:margin-top="0.423cm" fo:margin-bottom="0.212cm" fo:text-indent="-0.499cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="Numbering_20_3" style:display-name="Numbering 3" style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="1.498cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0.212cm" fo:text-indent="-0.499cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="Numbering_20_3_20_End" style:display-name="Numbering 3 End" style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="1.498cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0.423cm" fo:text-indent="-0.499cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="Numbering_20_3_20_Cont." style:display-name="Numbering 3 Cont." style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="1.498cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0.212cm" fo:text-indent="0cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="Numbering_20_4_20_Start" style:display-name="Numbering 4 Start" style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="1.997cm" fo:margin-right="0cm" fo:margin-top="0.423cm" fo:margin-bottom="0.212cm" fo:text-indent="-0.499cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="Numbering_20_4" style:display-name="Numbering 4" style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="1.997cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0.212cm" fo:text-indent="-0.499cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="Numbering_20_4_20_End" style:display-name="Numbering 4 End" style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="1.997cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0.423cm" fo:text-indent="-0.499cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="Numbering_20_4_20_Cont." style:display-name="Numbering 4 Cont." style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="1.997cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0.212cm" fo:text-indent="0cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="Numbering_20_5_20_Start" style:display-name="Numbering 5 Start" style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="2.496cm" fo:margin-right="0cm" fo:margin-top="0.423cm" fo:margin-bottom="0.212cm" fo:text-indent="-0.499cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="Numbering_20_5" style:display-name="Numbering 5" style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="2.496cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0.212cm" fo:text-indent="-0.499cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="Numbering_20_5_20_End" style:display-name="Numbering 5 End" style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="2.496cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0.423cm" fo:text-indent="-0.499cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="Numbering_20_5_20_Cont." style:display-name="Numbering 5 Cont." style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="2.496cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0.212cm" fo:text-indent="0cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="List_20_1_20_Start" style:display-name="List 1 Start" style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="0.499cm" fo:margin-right="0cm" fo:margin-top="0.423cm" fo:margin-bottom="0.212cm" fo:text-indent="-0.499cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="List_20_1" style:display-name="List 1" style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="0.499cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0.212cm" fo:text-indent="-0.499cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="List_20_1_20_End" style:display-name="List 1 End" style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="0.499cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0.423cm" fo:text-indent="-0.499cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="List_20_1_20_Cont." style:display-name="List 1 Cont." style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="0.499cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0.212cm" fo:text-indent="0cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="List_20_2_20_Start" style:display-name="List 2 Start" style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="0.998cm" fo:margin-right="0cm" fo:margin-top="0.423cm" fo:margin-bottom="0.212cm" fo:text-indent="-0.499cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="List_20_2" style:display-name="List 2" style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="0.998cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0.212cm" fo:text-indent="-0.499cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="List_20_2_20_End" style:display-name="List 2 End" style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="0.998cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0.423cm" fo:text-indent="-0.499cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="List_20_2_20_Cont." style:display-name="List 2 Cont." style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="0.998cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0.212cm" fo:text-indent="0cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="List_20_3_20_Start" style:display-name="List 3 Start" style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="1.498cm" fo:margin-right="0cm" fo:margin-top="0.423cm" fo:margin-bottom="0.212cm" fo:text-indent="-0.499cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="List_20_3" style:display-name="List 3" style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="1.498cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0.212cm" fo:text-indent="-0.499cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="List_20_3_20_End" style:display-name="List 3 End" style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="1.498cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0.423cm" fo:text-indent="-0.499cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="List_20_3_20_Cont." style:display-name="List 3 Cont." style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="1.498cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0.212cm" fo:text-indent="0cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="List_20_4_20_Start" style:display-name="List 4 Start" style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="1.997cm" fo:margin-right="0cm" fo:margin-top="0.423cm" fo:margin-bottom="0.212cm" fo:text-indent="-0.499cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="List_20_4" style:display-name="List 4" style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="1.997cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0.212cm" fo:text-indent="-0.499cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="List_20_4_20_End" style:display-name="List 4 End" style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="1.997cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0.423cm" fo:text-indent="-0.499cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="List_20_4_20_Cont." style:display-name="List 4 Cont." style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="1.997cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0.212cm" fo:text-indent="0cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="List_20_5_20_Start" style:display-name="List 5 Start" style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="2.496cm" fo:margin-right="0cm" fo:margin-top="0.423cm" fo:margin-bottom="0.212cm" fo:text-indent="-0.499cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="List_20_5" style:display-name="List 5" style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="2.496cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0.212cm" fo:text-indent="-0.499cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="List_20_5_20_End" style:display-name="List 5 End" style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="2.496cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0.423cm" fo:text-indent="-0.499cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="List_20_5_20_Cont." style:display-name="List 5 Cont." style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="2.496cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0.212cm" fo:text-indent="0cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="Header" style:family="paragraph" style:parent-style-name="Standard" style:class="extra" style:master-page-name="">
   <style:paragraph-properties fo:text-align="end" style:justify-single-word="false" style:page-number="auto" fo:padding="0.049cm" fo:border-left="none" fo:border-right="none" fo:border-top="none" fo:border-bottom="0.002cm solid #000000" style:shadow="none">
    <style:tab-stops>
     <style:tab-stop style:position="8.5cm" style:type="center"/>
     <style:tab-stop style:position="17.002cm" style:type="right"/>
    </style:tab-stops>
   </style:paragraph-properties>
   <style:text-properties fo:font-variant="normal" fo:text-transform="none" fo:font-style="italic"/>
  </style:style>
  <style:style style:name="Header_20_left" style:display-name="Header left" style:family="paragraph" style:parent-style-name="Standard" style:class="extra">
   <style:paragraph-properties>
    <style:tab-stops>
     <style:tab-stop style:position="8.5cm" style:type="center"/>
     <style:tab-stop style:position="17.002cm" style:type="right"/>
    </style:tab-stops>
   </style:paragraph-properties>
  </style:style>
  <style:style style:name="Header_20_right" style:display-name="Header right" style:family="paragraph" style:parent-style-name="Standard" style:class="extra">
   <style:paragraph-properties>
    <style:tab-stops>
     <style:tab-stop style:position="8.5cm" style:type="center"/>
     <style:tab-stop style:position="17.002cm" style:type="right"/>
    </style:tab-stops>
   </style:paragraph-properties>
  </style:style>
  <style:style style:name="Footer_20_left" style:display-name="Footer left" style:family="paragraph" style:parent-style-name="Standard" style:class="extra">
   <style:paragraph-properties>
    <style:tab-stops>
     <style:tab-stop style:position="8.5cm" style:type="center"/>
     <style:tab-stop style:position="17.002cm" style:type="right"/>
    </style:tab-stops>
   </style:paragraph-properties>
  </style:style>
  <style:style style:name="Footer_20_right" style:display-name="Footer right" style:family="paragraph" style:parent-style-name="Standard" style:class="extra">
   <style:paragraph-properties>
    <style:tab-stops>
     <style:tab-stop style:position="8.5cm" style:type="center"/>
     <style:tab-stop style:position="17.002cm" style:type="right"/>
    </style:tab-stops>
   </style:paragraph-properties>
  </style:style>
  <style:style style:name="Illustration" style:family="paragraph" style:parent-style-name="Caption" style:class="extra"/>
  <style:style style:name="Text" style:family="paragraph" style:parent-style-name="Caption" style:class="extra" style:master-page-name="">
   <style:paragraph-properties fo:margin-top="0.21cm" fo:margin-bottom="0.21cm" style:page-number="auto"/>
  </style:style>
  <style:style style:name="Frame_20_contents" style:display-name="Frame contents" style:family="paragraph" style:parent-style-name="Text_20_body" style:class="extra"/>
  <style:style style:name="Addressee" style:family="paragraph" style:parent-style-name="Standard" style:class="extra">
   <style:paragraph-properties fo:margin-top="0cm" fo:margin-bottom="0.106cm"/>
  </style:style>
  <style:style style:name="Sender" style:family="paragraph" style:parent-style-name="Standard" style:class="extra">
   <style:paragraph-properties fo:margin-top="0cm" fo:margin-bottom="0.106cm" fo:line-height="100%" text:number-lines="false" text:line-number="0"/>
  </style:style>
  <style:style style:name="Endnote" style:family="paragraph" style:parent-style-name="Standard" style:class="extra">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="0.499cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="-0.499cm" style:auto-text-indent="false" text:number-lines="false" text:line-number="0"/>
   <style:text-properties fo:font-size="10pt"/>
  </style:style>
  <style:style style:name="Drawing" style:family="paragraph" style:parent-style-name="Caption" style:class="extra"/>
  <style:style style:name="Index_20_Heading" style:display-name="Index Heading" style:family="paragraph" style:parent-style-name="Heading" style:class="index">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="0cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false"/>
   <style:text-properties fo:font-size="16pt" fo:font-weight="bold"/>
  </style:style>
  <style:style style:name="Index_20_1" style:display-name="Index 1" style:family="paragraph" style:parent-style-name="Index" style:class="index">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="0.499cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="Index_20_2" style:display-name="Index 2" style:family="paragraph" style:parent-style-name="Index" style:class="index">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="0.499cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="Index_20_3" style:display-name="Index 3" style:family="paragraph" style:parent-style-name="Index" style:class="index">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="0cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="Index_20_Separator" style:display-name="Index Separator" style:family="paragraph" style:parent-style-name="Index" style:class="index">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="0cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="Contents_20_Heading" style:display-name="Contents Heading" style:family="paragraph" style:parent-style-name="Heading" style:class="index">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="0cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false"/>
   <style:text-properties fo:font-size="16pt" fo:font-weight="bold"/>
  </style:style>
  <style:style style:name="Contents_20_1" style:display-name="Contents 1" style:family="paragraph" style:parent-style-name="Index" style:class="index">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="0cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false">
    <style:tab-stops>
     <style:tab-stop style:position="17.002cm" style:type="right" style:leader-style="dotted" style:leader-text="."/>
    </style:tab-stops>
   </style:paragraph-properties>
  </style:style>
  <style:style style:name="Contents_20_2" style:display-name="Contents 2" style:family="paragraph" style:parent-style-name="Index" style:class="index">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="0.499cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false">
    <style:tab-stops>
     <style:tab-stop style:position="16.503cm" style:type="right" style:leader-style="dotted" style:leader-text="."/>
    </style:tab-stops>
   </style:paragraph-properties>
  </style:style>
  <style:style style:name="Contents_20_3" style:display-name="Contents 3" style:family="paragraph" style:parent-style-name="Index" style:class="index">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="0.998cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false">
    <style:tab-stops>
     <style:tab-stop style:position="16.004cm" style:type="right" style:leader-style="dotted" style:leader-text="."/>
    </style:tab-stops>
   </style:paragraph-properties>
  </style:style>
  <style:style style:name="Contents_20_4" style:display-name="Contents 4" style:family="paragraph" style:parent-style-name="Index" style:class="index">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="1.498cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false">
    <style:tab-stops>
     <style:tab-stop style:position="15.505cm" style:type="right" style:leader-style="dotted" style:leader-text="."/>
    </style:tab-stops>
   </style:paragraph-properties>
  </style:style>
  <style:style style:name="Contents_20_5" style:display-name="Contents 5" style:family="paragraph" style:parent-style-name="Index" style:class="index">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="1.997cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false">
    <style:tab-stops>
     <style:tab-stop style:position="15.005cm" style:type="right" style:leader-style="dotted" style:leader-text="."/>
    </style:tab-stops>
   </style:paragraph-properties>
  </style:style>
  <style:style style:name="User_20_Index_20_Heading" style:display-name="User Index Heading" style:family="paragraph" style:parent-style-name="Heading" style:class="index">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="0cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false"/>
   <style:text-properties fo:font-size="16pt" fo:font-weight="bold"/>
  </style:style>
  <style:style style:name="User_20_Index_20_1" style:display-name="User Index 1" style:family="paragraph" style:parent-style-name="Index" style:class="index">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="0cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false">
    <style:tab-stops>
     <style:tab-stop style:position="17.002cm" style:type="right" style:leader-style="dotted" style:leader-text="."/>
    </style:tab-stops>
   </style:paragraph-properties>
  </style:style>
  <style:style style:name="User_20_Index_20_2" style:display-name="User Index 2" style:family="paragraph" style:parent-style-name="Index" style:class="index">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="0.499cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false">
    <style:tab-stops>
     <style:tab-stop style:position="16.503cm" style:type="right" style:leader-style="dotted" style:leader-text="."/>
    </style:tab-stops>
   </style:paragraph-properties>
  </style:style>
  <style:style style:name="User_20_Index_20_3" style:display-name="User Index 3" style:family="paragraph" style:parent-style-name="Index" style:class="index">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="0.998cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false">
    <style:tab-stops>
     <style:tab-stop style:position="16.004cm" style:type="right" style:leader-style="dotted" style:leader-text="."/>
    </style:tab-stops>
   </style:paragraph-properties>
  </style:style>
  <style:style style:name="User_20_Index_20_4" style:display-name="User Index 4" style:family="paragraph" style:parent-style-name="Index" style:class="index">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="1.498cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false">
    <style:tab-stops>
     <style:tab-stop style:position="15.505cm" style:type="right" style:leader-style="dotted" style:leader-text="."/>
    </style:tab-stops>
   </style:paragraph-properties>
  </style:style>
  <style:style style:name="User_20_Index_20_5" style:display-name="User Index 5" style:family="paragraph" style:parent-style-name="Index" style:class="index">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="1.997cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false">
    <style:tab-stops>
     <style:tab-stop style:position="15.005cm" style:type="right" style:leader-style="dotted" style:leader-text="."/>
    </style:tab-stops>
   </style:paragraph-properties>
  </style:style>
  <style:style style:name="Contents_20_6" style:display-name="Contents 6" style:family="paragraph" style:parent-style-name="Index" style:class="index">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="2.496cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false">
    <style:tab-stops>
     <style:tab-stop style:position="11.105cm" style:type="right" style:leader-style="dotted" style:leader-text="."/>
    </style:tab-stops>
   </style:paragraph-properties>
  </style:style>
  <style:style style:name="Contents_20_7" style:display-name="Contents 7" style:family="paragraph" style:parent-style-name="Index" style:class="index">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="2.995cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false">
    <style:tab-stops>
     <style:tab-stop style:position="10.606cm" style:type="right" style:leader-style="dotted" style:leader-text="."/>
    </style:tab-stops>
   </style:paragraph-properties>
  </style:style>
  <style:style style:name="Contents_20_8" style:display-name="Contents 8" style:family="paragraph" style:parent-style-name="Index" style:class="index">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="3.494cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false">
    <style:tab-stops>
     <style:tab-stop style:position="10.107cm" style:type="right" style:leader-style="dotted" style:leader-text="."/>
    </style:tab-stops>
   </style:paragraph-properties>
  </style:style>
  <style:style style:name="Contents_20_9" style:display-name="Contents 9" style:family="paragraph" style:parent-style-name="Index" style:class="index">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="3.993cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false">
    <style:tab-stops>
     <style:tab-stop style:position="9.608cm" style:type="right" style:leader-style="dotted" style:leader-text="."/>
    </style:tab-stops>
   </style:paragraph-properties>
  </style:style>
  <style:style style:name="Contents_20_10" style:display-name="Contents 10" style:family="paragraph" style:parent-style-name="Index" style:class="index">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="4.493cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false">
    <style:tab-stops>
     <style:tab-stop style:position="9.109cm" style:type="right" style:leader-style="dotted" style:leader-text="."/>
    </style:tab-stops>
   </style:paragraph-properties>
  </style:style>
  <style:style style:name="Illustration_20_Index_20_Heading" style:display-name="Illustration Index Heading" style:family="paragraph" style:parent-style-name="Heading" style:class="index">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="0cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false" text:number-lines="false" text:line-number="0"/>
   <style:text-properties fo:font-size="16pt" fo:font-weight="bold"/>
  </style:style>
  <style:style style:name="Illustration_20_Index_20_1" style:display-name="Illustration Index 1" style:family="paragraph" style:parent-style-name="Index" style:class="index">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="0cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false">
    <style:tab-stops>
     <style:tab-stop style:position="13.601cm" style:type="right" style:leader-style="dotted" style:leader-text="."/>
    </style:tab-stops>
   </style:paragraph-properties>
  </style:style>
  <style:style style:name="Object_20_index_20_heading" style:display-name="Object index heading" style:family="paragraph" style:parent-style-name="Heading" style:class="index">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="0cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false" text:number-lines="false" text:line-number="0"/>
   <style:text-properties fo:font-size="16pt" fo:font-weight="bold"/>
  </style:style>
  <style:style style:name="Object_20_index_20_1" style:display-name="Object index 1" style:family="paragraph" style:parent-style-name="Index" style:class="index">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="0cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false">
    <style:tab-stops>
     <style:tab-stop style:position="13.601cm" style:type="right" style:leader-style="dotted" style:leader-text="."/>
    </style:tab-stops>
   </style:paragraph-properties>
  </style:style>
  <style:style style:name="Table_20_index_20_heading" style:display-name="Table index heading" style:family="paragraph" style:parent-style-name="Heading" style:class="index">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="0cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false" text:number-lines="false" text:line-number="0"/>
   <style:text-properties fo:font-size="16pt" fo:font-weight="bold"/>
  </style:style>
  <style:style style:name="Table_20_index_20_1" style:display-name="Table index 1" style:family="paragraph" style:parent-style-name="Index" style:class="index">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="0cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false">
    <style:tab-stops>
     <style:tab-stop style:position="13.601cm" style:type="right" style:leader-style="dotted" style:leader-text="."/>
    </style:tab-stops>
   </style:paragraph-properties>
  </style:style>
  <style:style style:name="Bibliography_20_Heading" style:display-name="Bibliography Heading" style:family="paragraph" style:parent-style-name="Heading" style:class="index">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="0cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false" text:number-lines="false" text:line-number="0"/>
   <style:text-properties fo:font-size="16pt" fo:font-weight="bold"/>
  </style:style>
  <style:style style:name="Bibliography_20_1" style:display-name="Bibliography 1" style:family="paragraph" style:parent-style-name="Index" style:class="index">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="0cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false">
    <style:tab-stops>
     <style:tab-stop style:position="13.601cm" style:type="right" style:leader-style="dotted" style:leader-text="."/>
    </style:tab-stops>
   </style:paragraph-properties>
  </style:style>
  <style:style style:name="User_20_Index_20_6" style:display-name="User Index 6" style:family="paragraph" style:parent-style-name="Index" style:class="index">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="2.496cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false">
    <style:tab-stops>
     <style:tab-stop style:position="11.105cm" style:type="right" style:leader-style="dotted" style:leader-text="."/>
    </style:tab-stops>
   </style:paragraph-properties>
  </style:style>
  <style:style style:name="User_20_Index_20_7" style:display-name="User Index 7" style:family="paragraph" style:parent-style-name="Index" style:class="index">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="2.995cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false">
    <style:tab-stops>
     <style:tab-stop style:position="10.606cm" style:type="right" style:leader-style="dotted" style:leader-text="."/>
    </style:tab-stops>
   </style:paragraph-properties>
  </style:style>
  <style:style style:name="User_20_Index_20_8" style:display-name="User Index 8" style:family="paragraph" style:parent-style-name="Index" style:class="index">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="3.494cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false">
    <style:tab-stops>
     <style:tab-stop style:position="10.107cm" style:type="right" style:leader-style="dotted" style:leader-text="."/>
    </style:tab-stops>
   </style:paragraph-properties>
  </style:style>
  <style:style style:name="User_20_Index_20_9" style:display-name="User Index 9" style:family="paragraph" style:parent-style-name="Index" style:class="index">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="3.993cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false">
    <style:tab-stops>
     <style:tab-stop style:position="9.608cm" style:type="right" style:leader-style="dotted" style:leader-text="."/>
    </style:tab-stops>
   </style:paragraph-properties>
  </style:style>
  <style:style style:name="User_20_Index_20_10" style:display-name="User Index 10" style:family="paragraph" style:parent-style-name="Index" style:class="index">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="4.493cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false">
    <style:tab-stops>
     <style:tab-stop style:position="9.109cm" style:type="right" style:leader-style="dotted" style:leader-text="."/>
    </style:tab-stops>
   </style:paragraph-properties>
  </style:style>
  <style:style style:name="List_20_Contents" style:display-name="List Contents" style:family="paragraph" style:parent-style-name="Standard" style:class="html">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="1cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="List_20_Heading" style:display-name="List Heading" style:family="paragraph" style:parent-style-name="Standard" style:next-style-name="List_20_Contents" style:class="html">
   <style:paragraph-properties fo:margin="100%" fo:margin-left="0cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="yWriter_20_mark" style:display-name="yWriter mark" style:family="paragraph" style:parent-style-name="Standard" style:next-style-name="Standard" style:class="text">
   <style:text-properties fo:color="#008000" fo:font-size="10pt"/>
  </style:style>
  <style:style style:name="yWriter_20_mark_20_unused" style:display-name="yWriter mark unused" style:family="paragraph" style:parent-style-name="Standard" style:next-style-name="Standard" style:class="text">
   <style:text-properties fo:color="#808080" fo:font-size="10pt"/>
  </style:style>
  <style:style style:name="yWriter_20_mark_20_notes" style:display-name="yWriter mark notes" style:family="paragraph" style:parent-style-name="Standard" style:next-style-name="Standard" style:class="text">
   <style:text-properties fo:color="#0000FF" fo:font-size="10pt"/>
  </style:style>
  <style:style style:name="yWriter_20_mark_20_todo" style:display-name="yWriter mark todo" style:family="paragraph" style:parent-style-name="Standard" style:next-style-name="Standard" style:class="text">
   <style:text-properties fo:color="#B22222" fo:font-size="10pt"/>
  </style:style>
  <style:style style:name="Numbering_20_Symbols" style:display-name="Numbering Symbols" style:family="text"/>
  <style:style style:name="Bullet_20_Symbols" style:display-name="Bullet Symbols" style:family="text">
   <style:text-properties style:font-name="StarSymbol" fo:font-size="9pt"/>
  </style:style>
  <style:style style:name="Emphasis" style:family="text">
   <style:text-properties fo:font-style="italic" fo:background-color="transparent"/>
  </style:style>
  <style:style style:name="Strong_20_Emphasis" style:display-name="Strong Emphasis" style:family="text">
   <style:text-properties fo:text-transform="uppercase"/>
  </style:style>
  <style:style style:name="Citation" style:family="text">
   <style:text-properties fo:font-style="italic"/>
  </style:style>
  <style:style style:name="Teletype" style:family="text">
   <style:text-properties style:font-name="Cumberland" style:font-name-asian="Cumberland" style:font-name-complex="Cumberland"/>
  </style:style>
  <style:style style:name="Internet_20_link" style:display-name="Internet link" style:family="text">
   <style:text-properties fo:color="#000080" style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color"/>
  </style:style>
  <style:style style:name="Footnote_20_Symbol" style:display-name="Footnote Symbol" style:family="text"/>
  <style:style style:name="Footnote_20_anchor" style:display-name="Footnote anchor" style:family="text">
   <style:text-properties style:text-position="super 58%"/>
  </style:style>
  <style:style style:name="Definition" style:family="text"/>
  <style:style style:name="Line_20_numbering" style:display-name="Line numbering" style:family="text">
   <style:text-properties style:font-name="Courier New" fo:font-size="8pt"/>
  </style:style>
  <style:style style:name="Page_20_Number" style:display-name="Page Number" style:family="text">
   <style:text-properties style:font-name="Courier New" fo:font-size="8pt"/>
  </style:style>
  <style:style style:name="Caption_20_characters" style:display-name="Caption characters" style:family="text"/>
  <style:style style:name="Drop_20_Caps" style:display-name="Drop Caps" style:family="text"/>
  <style:style style:name="Visited_20_Internet_20_Link" style:display-name="Visited Internet Link" style:family="text">
   <style:text-properties fo:color="#800000" style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color"/>
  </style:style>
  <style:style style:name="Placeholder" style:family="text">
   <style:text-properties fo:font-variant="small-caps" fo:color="#008080" style:text-underline-style="dotted" style:text-underline-width="auto" style:text-underline-color="font-color"/>
  </style:style>
  <style:style style:name="Index_20_Link" style:display-name="Index Link" style:family="text"/>
  <style:style style:name="Endnote_20_Symbol" style:display-name="Endnote Symbol" style:family="text"/>
  <style:style style:name="Main_20_index_20_entry" style:display-name="Main index entry" style:family="text">
   <style:text-properties fo:font-weight="bold" style:font-weight-asian="bold" style:font-weight-complex="bold"/>
  </style:style>
  <style:style style:name="Endnote_20_anchor" style:display-name="Endnote anchor" style:family="text">
   <style:text-properties style:text-position="super 58%"/>
  </style:style>
  <style:style style:name="Rubies" style:family="text">
   <style:text-properties fo:font-size="6pt" style:font-size-asian="6pt" style:font-size-complex="6pt"/>
  </style:style>
  <style:style style:name="Source_20_Text" style:display-name="Source Text" style:family="text">
   <style:text-properties style:font-name="Cumberland" style:font-name-asian="Cumberland" style:font-name-complex="Cumberland"/>
  </style:style>
  <style:style style:name="Example" style:family="text">
   <style:text-properties style:font-name="Courier New1"/>
  </style:style>
  <style:style style:name="User_20_Entry" style:display-name="User Entry" style:family="text">
   <style:text-properties style:font-name="Cumberland" style:font-name-asian="Cumberland" style:font-name-complex="Cumberland"/>
  </style:style>
  <style:style style:name="Variable" style:family="text">
   <style:text-properties fo:font-style="italic" style:font-style-asian="italic" style:font-style-complex="italic"/>
  </style:style>
  <style:style style:name="Frame" style:family="graphic">
   <style:graphic-properties text:anchor-type="paragraph" svg:x="0cm" svg:y="0cm" style:wrap="parallel" style:number-wrapped-paragraphs="no-limit" style:wrap-contour="false" style:vertical-pos="top" style:vertical-rel="paragraph-content" style:horizontal-pos="center" style:horizontal-rel="paragraph-content"/>
  </style:style>
  <style:style style:name="Graphics" style:family="graphic">
   <style:graphic-properties text:anchor-type="paragraph" svg:x="0cm" svg:y="0cm" style:wrap="none" style:vertical-pos="top" style:vertical-rel="paragraph" style:horizontal-pos="center" style:horizontal-rel="paragraph"/>
  </style:style>
  <style:style style:name="OLE" style:family="graphic">
   <style:graphic-properties text:anchor-type="paragraph" svg:x="0cm" svg:y="0cm" style:wrap="none" style:vertical-pos="top" style:vertical-rel="paragraph" style:horizontal-pos="center" style:horizontal-rel="paragraph"/>
  </style:style>
  <style:style style:name="Formula" style:family="graphic">
   <style:graphic-properties text:anchor-type="as-char" svg:y="0cm" style:vertical-pos="top" style:vertical-rel="baseline"/>
  </style:style>
  <style:style style:name="Labels" style:family="graphic" style:auto-update="true">
   <style:graphic-properties text:anchor-type="as-char" svg:y="0cm" fo:margin-left="0.201cm" fo:margin-right="0.201cm" style:protect="size position" style:vertical-pos="top" style:vertical-rel="baseline"/>
  </style:style>
  <text:outline-style style:name="Outline">
   <text:outline-level-style text:level="1" style:num-format="">
    <style:list-level-properties text:min-label-distance="0.381cm"/>
   </text:outline-level-style>
   <text:outline-level-style text:level="2" style:num-format="">
    <style:list-level-properties text:min-label-distance="0.381cm"/>
   </text:outline-level-style>
   <text:outline-level-style text:level="3" style:num-format="">
    <style:list-level-properties text:min-label-distance="0.381cm"/>
   </text:outline-level-style>
   <text:outline-level-style text:level="4" style:num-format="">
    <style:list-level-properties text:min-label-distance="0.381cm"/>
   </text:outline-level-style>
   <text:outline-level-style text:level="5" style:num-format="">
    <style:list-level-properties text:min-label-distance="0.381cm"/>
   </text:outline-level-style>
   <text:outline-level-style text:level="6" style:num-format="">
    <style:list-level-properties text:min-label-distance="0.381cm"/>
   </text:outline-level-style>
   <text:outline-level-style text:level="7" style:num-format="">
    <style:list-level-properties text:min-label-distance="0.381cm"/>
   </text:outline-level-style>
   <text:outline-level-style text:level="8" style:num-format="">
    <style:list-level-properties text:min-label-distance="0.381cm"/>
   </text:outline-level-style>
   <text:outline-level-style text:level="9" style:num-format="">
    <style:list-level-properties text:min-label-distance="0.381cm"/>
   </text:outline-level-style>
   <text:outline-level-style text:level="10" style:num-format="">
    <style:list-level-properties text:min-label-distance="0.381cm"/>
   </text:outline-level-style>
  </text:outline-style>
  <text:list-style style:name="Numbering_20_1" style:display-name="Numbering 1">
   <text:list-level-style-number text:level="1" text:style-name="Numbering_20_Symbols" style:num-suffix="." style:num-format="1">
    <style:list-level-properties text:min-label-width="0.499cm"/>
   </text:list-level-style-number>
   <text:list-level-style-number text:level="2" text:style-name="Numbering_20_Symbols" style:num-suffix="." style:num-format="1">
    <style:list-level-properties text:space-before="0.499cm" text:min-label-width="0.499cm"/>
   </text:list-level-style-number>
   <text:list-level-style-number text:level="3" text:style-name="Numbering_20_Symbols" style:num-suffix="." style:num-format="1">
    <style:list-level-properties text:space-before="0.999cm" text:min-label-width="0.499cm"/>
   </text:list-level-style-number>
   <text:list-level-style-number text:level="4" text:style-name="Numbering_20_Symbols" style:num-suffix="." style:num-format="1">
    <style:list-level-properties text:space-before="1.498cm" text:min-label-width="0.499cm"/>
   </text:list-level-style-number>
   <text:list-level-style-number text:level="5" text:style-name="Numbering_20_Symbols" style:num-suffix="." style:num-format="1">
    <style:list-level-properties text:space-before="1.997cm" text:min-label-width="0.499cm"/>
   </text:list-level-style-number>
   <text:list-level-style-number text:level="6" text:style-name="Numbering_20_Symbols" style:num-suffix="." style:num-format="1">
    <style:list-level-properties text:space-before="2.496cm" text:min-label-width="0.499cm"/>
   </text:list-level-style-number>
   <text:list-level-style-number text:level="7" text:style-name="Numbering_20_Symbols" style:num-suffix="." style:num-format="1">
    <style:list-level-properties text:space-before="2.995cm" text:min-label-width="0.499cm"/>
   </text:list-level-style-number>
   <text:list-level-style-number text:level="8" text:style-name="Numbering_20_Symbols" style:num-suffix="." style:num-format="1">
    <style:list-level-properties text:space-before="3.494cm" text:min-label-width="0.499cm"/>
   </text:list-level-style-number>
   <text:list-level-style-number text:level="9" text:style-name="Numbering_20_Symbols" style:num-suffix="." style:num-format="1">
    <style:list-level-properties text:space-before="3.994cm" text:min-label-width="0.499cm"/>
   </text:list-level-style-number>
   <text:list-level-style-number text:level="10" text:style-name="Numbering_20_Symbols" style:num-suffix="." style:num-format="1">
    <style:list-level-properties text:space-before="4.493cm" text:min-label-width="0.499cm"/>
   </text:list-level-style-number>
  </text:list-style>
  <text:list-style style:name="Numbering_20_2" style:display-name="Numbering 2">
   <text:list-level-style-number text:level="1" text:style-name="Numbering_20_Symbols" style:num-format="1">
    <style:list-level-properties text:min-label-width="0.499cm"/>
   </text:list-level-style-number>
   <text:list-level-style-number text:level="2" text:style-name="Numbering_20_Symbols" style:num-format="1" text:start-value="2">
    <style:list-level-properties text:space-before="0.499cm" text:min-label-width="0.499cm"/>
   </text:list-level-style-number>
   <text:list-level-style-number text:level="3" text:style-name="Numbering_20_Symbols" style:num-format="1" text:start-value="3">
    <style:list-level-properties text:space-before="0.998cm" text:min-label-width="1cm"/>
   </text:list-level-style-number>
   <text:list-level-style-number text:level="4" text:style-name="Numbering_20_Symbols" style:num-format="1" text:start-value="4">
    <style:list-level-properties text:space-before="1.998cm" text:min-label-width="1.251cm"/>
   </text:list-level-style-number>
   <text:list-level-style-number text:level="5" text:style-name="Numbering_20_Symbols" style:num-format="1" text:start-value="5">
    <style:list-level-properties text:space-before="3.249cm" text:min-label-width="1.499cm"/>
   </text:list-level-style-number>
   <text:list-level-style-number text:level="6" text:style-name="Numbering_20_Symbols" style:num-format="1" text:start-value="6">
    <style:list-level-properties text:space-before="4.748cm" text:min-label-width="1.801cm"/>
   </text:list-level-style-number>
   <text:list-level-style-number text:level="7" text:style-name="Numbering_20_Symbols" style:num-format="1" text:start-value="7">
    <style:list-level-properties text:space-before="6.549cm" text:min-label-width="2.3cm"/>
   </text:list-level-style-number>
   <text:list-level-style-number text:level="8" text:style-name="Numbering_20_Symbols" style:num-format="1" text:start-value="8">
    <style:list-level-properties text:space-before="8.849cm" text:min-label-width="2.6cm"/>
   </text:list-level-style-number>
   <text:list-level-style-number text:level="9" text:style-name="Numbering_20_Symbols" style:num-format="1" text:start-value="9">
    <style:list-level-properties text:space-before="11.449cm" text:min-label-width="2.801cm"/>
   </text:list-level-style-number>
   <text:list-level-style-number text:level="10" text:style-name="Numbering_20_Symbols" style:num-format="1" text:start-value="10">
    <style:list-level-properties text:space-before="14.25cm" text:min-label-width="3.101cm"/>
   </text:list-level-style-number>
  </text:list-style>
  <text:list-style style:name="Numbering_20_3" style:display-name="Numbering 3">
   <text:list-level-style-number text:level="1" text:style-name="Numbering_20_Symbols" style:num-format="1">
    <style:list-level-properties text:min-label-width="3cm"/>
   </text:list-level-style-number>
   <text:list-level-style-number text:level="2" text:style-name="Numbering_20_Symbols" style:num-format="1" text:start-value="2">
    <style:list-level-properties text:space-before="3.001cm" text:min-label-width="3cm"/>
   </text:list-level-style-number>
   <text:list-level-style-number text:level="3" text:style-name="Numbering_20_Symbols" style:num-format="1" text:start-value="3">
    <style:list-level-properties text:space-before="6.001cm" text:min-label-width="3cm"/>
   </text:list-level-style-number>
   <text:list-level-style-number text:level="4" text:style-name="Numbering_20_Symbols" style:num-format="1" text:start-value="4">
    <style:list-level-properties text:space-before="9.002cm" text:min-label-width="3cm"/>
   </text:list-level-style-number>
   <text:list-level-style-number text:level="5" text:style-name="Numbering_20_Symbols" style:num-format="1" text:start-value="5">
    <style:list-level-properties text:space-before="12.002cm" text:min-label-width="3cm"/>
   </text:list-level-style-number>
   <text:list-level-style-number text:level="6" text:style-name="Numbering_20_Symbols" style:num-format="1" text:start-value="6">
    <style:list-level-properties text:space-before="15.002cm" text:min-label-width="3cm"/>
   </text:list-level-style-number>
   <text:list-level-style-number text:level="7" text:style-name="Numbering_20_Symbols" style:num-format="1" text:start-value="7">
    <style:list-level-properties text:space-before="18.003cm" text:min-label-width="3cm"/>
   </text:list-level-style-number>
   <text:list-level-style-number text:level="8" text:style-name="Numbering_20_Symbols" style:num-format="1" text:start-value="8">
    <style:list-level-properties text:space-before="21.003cm" text:min-label-width="3cm"/>
   </text:list-level-style-number>
   <text:list-level-style-number text:level="9" text:style-name="Numbering_20_Symbols" style:num-format="1" text:start-value="9">
    <style:list-level-properties text:space-before="24.003cm" text:min-label-width="3cm"/>
   </text:list-level-style-number>
   <text:list-level-style-number text:level="10" text:style-name="Numbering_20_Symbols" style:num-format="1" text:start-value="10">
    <style:list-level-properties text:space-before="27.004cm" text:min-label-width="3cm"/>
   </text:list-level-style-number>
  </text:list-style>
  <text:list-style style:name="Numbering_20_4" style:display-name="Numbering 4">
   <text:list-level-style-number text:level="1" text:style-name="Numbering_20_Symbols" style:num-suffix="." style:num-format="I">
    <style:list-level-properties text:min-label-width="0.499cm"/>
   </text:list-level-style-number>
   <text:list-level-style-number text:level="2" text:style-name="Numbering_20_Symbols" style:num-suffix="." style:num-format="I" text:start-value="2">
    <style:list-level-properties text:space-before="0.501cm" text:min-label-width="0.499cm"/>
   </text:list-level-style-number>
   <text:list-level-style-number text:level="3" text:style-name="Numbering_20_Symbols" style:num-suffix="." style:num-format="I" text:start-value="3">
    <style:list-level-properties text:space-before="1cm" text:min-label-width="0.499cm"/>
   </text:list-level-style-number>
   <text:list-level-style-number text:level="4" text:style-name="Numbering_20_Symbols" style:num-suffix="." style:num-format="I" text:start-value="4">
    <style:list-level-properties text:space-before="1.501cm" text:min-label-width="0.499cm"/>
   </text:list-level-style-number>
   <text:list-level-style-number text:level="5" text:style-name="Numbering_20_Symbols" style:num-suffix="." style:num-format="I" text:start-value="5">
    <style:list-level-properties text:space-before="2cm" text:min-label-width="0.499cm"/>
   </text:list-level-style-number>
   <text:list-level-style-number text:level="6" text:style-name="Numbering_20_Symbols" style:num-suffix="." style:num-format="I" text:start-value="6">
    <style:list-level-properties text:space-before="2.501cm" text:min-label-width="0.499cm"/>
   </text:list-level-style-number>
   <text:list-level-style-number text:level="7" text:style-name="Numbering_20_Symbols" style:num-suffix="." style:num-format="I" text:start-value="7">
    <style:list-level-properties text:space-before="3.001cm" text:min-label-width="0.499cm"/>
   </text:list-level-style-number>
   <text:list-level-style-number text:level="8" text:style-name="Numbering_20_Symbols" style:num-suffix="." style:num-format="I" text:start-value="8">
    <style:list-level-properties text:space-before="3.502cm" text:min-label-width="0.499cm"/>
   </text:list-level-style-number>
   <text:list-level-style-number text:level="9" text:style-name="Numbering_20_Symbols" style:num-suffix="." style:num-format="I" text:start-value="9">
    <style:list-level-properties text:space-before="4.001cm" text:min-label-width="0.499cm"/>
   </text:list-level-style-number>
   <text:list-level-style-number text:level="10" text:style-name="Numbering_20_Symbols" style:num-suffix="." style:num-format="I" text:start-value="10">
    <style:list-level-properties text:space-before="4.502cm" text:min-label-width="0.499cm"/>
   </text:list-level-style-number>
  </text:list-style>
  <text:list-style style:name="Numbering_20_5" style:display-name="Numbering 5">
   <text:list-level-style-number text:level="1" text:style-name="Numbering_20_Symbols" style:num-suffix="." style:num-format="1">
    <style:list-level-properties text:min-label-width="0.4cm"/>
   </text:list-level-style-number>
   <text:list-level-style-number text:level="2" text:style-name="Numbering_20_Symbols" style:num-suffix="." style:num-format="1" text:start-value="2" text:display-levels="2">
    <style:list-level-properties text:space-before="0.45cm" text:min-label-width="0.651cm"/>
   </text:list-level-style-number>
   <text:list-level-style-number text:level="3" text:style-name="Numbering_20_Symbols" style:num-suffix=")" style:num-format="a" text:start-value="3">
    <style:list-level-properties text:space-before="1.1cm" text:min-label-width="0.45cm"/>
   </text:list-level-style-number>
   <text:list-level-style-bullet text:level="4" text:style-name="Numbering_20_Symbols" text:bullet-char="•">
    <style:list-level-properties text:space-before="1.605cm" text:min-label-width="0.395cm"/>
    <style:text-properties fo:font-family="StarSymbol"/>
   </text:list-level-style-bullet>
   <text:list-level-style-bullet text:level="5" text:style-name="Numbering_20_Symbols" text:bullet-char="•">
    <style:list-level-properties text:space-before="2cm" text:min-label-width="0.395cm"/>
    <style:text-properties fo:font-family="StarSymbol"/>
   </text:list-level-style-bullet>
   <text:list-level-style-bullet text:level="6" text:style-name="Numbering_20_Symbols" text:bullet-char="•">
    <style:list-level-properties text:space-before="2.395cm" text:min-label-width="0.395cm"/>
    <style:text-properties fo:font-family="StarSymbol"/>
   </text:list-level-style-bullet>
   <text:list-level-style-bullet text:level="7" text:style-name="Numbering_20_Symbols" text:bullet-char="•">
    <style:list-level-properties text:space-before="2.791cm" text:min-label-width="0.395cm"/>
    <style:text-properties fo:font-family="StarSymbol"/>
   </text:list-level-style-bullet>
   <text:list-level-style-bullet text:level="8" text:style-name="Numbering_20_Symbols" text:bullet-char="•">
    <style:list-level-properties text:space-before="3.186cm" text:min-label-width="0.395cm"/>
    <style:text-properties fo:font-family="StarSymbol"/>
   </text:list-level-style-bullet>
   <text:list-level-style-bullet text:level="9" text:style-name="Numbering_20_Symbols" text:bullet-char="•">
    <style:list-level-properties text:space-before="3.581cm" text:min-label-width="0.395cm"/>
    <style:text-properties fo:font-family="StarSymbol"/>
   </text:list-level-style-bullet>
   <text:list-level-style-bullet text:level="10" text:style-name="Numbering_20_Symbols" text:bullet-char="•">
    <style:list-level-properties text:space-before="3.976cm" text:min-label-width="0.395cm"/>
    <style:text-properties fo:font-family="StarSymbol"/>
   </text:list-level-style-bullet>
  </text:list-style>
  <text:list-style style:name="List_20_1" style:display-name="List 1">
   <text:list-level-style-bullet text:level="1" text:style-name="Numbering_20_Symbols" text:bullet-char="•">
    <style:list-level-properties text:min-label-width="0.395cm"/>
    <style:text-properties style:font-name="StarSymbol"/>
   </text:list-level-style-bullet>
   <text:list-level-style-bullet text:level="2" text:style-name="Numbering_20_Symbols" text:bullet-char="•">
    <style:list-level-properties text:space-before="0.395cm" text:min-label-width="0.395cm"/>
    <style:text-properties style:font-name="StarSymbol"/>
   </text:list-level-style-bullet>
   <text:list-level-style-bullet text:level="3" text:style-name="Numbering_20_Symbols" text:bullet-char="•">
    <style:list-level-properties text:space-before="0.79cm" text:min-label-width="0.395cm"/>
    <style:text-properties style:font-name="StarSymbol"/>
   </text:list-level-style-bullet>
   <text:list-level-style-bullet text:level="4" text:style-name="Numbering_20_Symbols" text:bullet-char="•">
    <style:list-level-properties text:space-before="1.185cm" text:min-label-width="0.395cm"/>
    <style:text-properties style:font-name="StarSymbol"/>
   </text:list-level-style-bullet>
   <text:list-level-style-bullet text:level="5" text:style-name="Numbering_20_Symbols" text:bullet-char="•">
    <style:list-level-properties text:space-before="1.581cm" text:min-label-width="0.395cm"/>
    <style:text-properties style:font-name="StarSymbol"/>
   </text:list-level-style-bullet>
   <text:list-level-style-bullet text:level="6" text:style-name="Numbering_20_Symbols" text:bullet-char="•">
    <style:list-level-properties text:space-before="1.976cm" text:min-label-width="0.395cm"/>
    <style:text-properties style:font-name="StarSymbol"/>
   </text:list-level-style-bullet>
   <text:list-level-style-bullet text:level="7" text:style-name="Numbering_20_Symbols" text:bullet-char="•">
    <style:list-level-properties text:space-before="2.371cm" text:min-label-width="0.395cm"/>
    <style:text-properties style:font-name="StarSymbol"/>
   </text:list-level-style-bullet>
   <text:list-level-style-bullet text:level="8" text:style-name="Numbering_20_Symbols" text:bullet-char="•">
    <style:list-level-properties text:space-before="2.766cm" text:min-label-width="0.395cm"/>
    <style:text-properties style:font-name="StarSymbol"/>
   </text:list-level-style-bullet>
   <text:list-level-style-bullet text:level="9" text:style-name="Numbering_20_Symbols" text:bullet-char="•">
    <style:list-level-properties text:space-before="3.161cm" text:min-label-width="0.395cm"/>
    <style:text-properties style:font-name="StarSymbol"/>
   </text:list-level-style-bullet>
   <text:list-level-style-bullet text:level="10" text:style-name="Numbering_20_Symbols" text:bullet-char="•">
    <style:list-level-properties text:space-before="3.556cm" text:min-label-width="0.395cm"/>
    <style:text-properties style:font-name="StarSymbol"/>
   </text:list-level-style-bullet>
  </text:list-style>
  <text:list-style style:name="List_20_2" style:display-name="List 2">
   <text:list-level-style-bullet text:level="1" text:style-name="Numbering_20_Symbols" text:bullet-char="–">
    <style:list-level-properties text:list-level-position-and-space-mode="label-alignment">
     <style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="0.3cm" fo:text-indent="-0.3cm" fo:margin-left="0.3cm"/>
    </style:list-level-properties>
    <style:text-properties fo:font-family="OpenSymbol"/>
   </text:list-level-style-bullet>
   <text:list-level-style-bullet text:level="2" text:style-name="Numbering_20_Symbols" text:bullet-char="–">
    <style:list-level-properties text:list-level-position-and-space-mode="label-alignment">
     <style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="0.6cm" fo:text-indent="-0.3cm" fo:margin-left="0.6cm"/>
    </style:list-level-properties>
    <style:text-properties fo:font-family="OpenSymbol"/>
   </text:list-level-style-bullet>
   <text:list-level-style-bullet text:level="3" text:style-name="Numbering_20_Symbols" text:bullet-char="–">
    <style:list-level-properties text:list-level-position-and-space-mode="label-alignment">
     <style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="0.9cm" fo:text-indent="-0.3cm" fo:margin-left="0.9cm"/>
    </style:list-level-properties>
    <style:text-properties fo:font-family="OpenSymbol"/>
   </text:list-level-style-bullet>
   <text:list-level-style-bullet text:level="4" text:style-name="Numbering_20_Symbols" text:bullet-char="–">
    <style:list-level-properties text:list-level-position-and-space-mode="label-alignment">
     <style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="1.199cm" fo:text-indent="-0.3cm" fo:margin-left="1.199cm"/>
    </style:list-level-properties>
    <style:text-properties fo:font-family="OpenSymbol"/>
   </text:list-level-style-bullet>
   <text:list-level-style-bullet text:level="5" text:style-name="Numbering_20_Symbols" text:bullet-char="–">
    <style:list-level-properties text:list-level-position-and-space-mode="label-alignment">
     <style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="1.499cm" fo:text-indent="-0.3cm" fo:margin-left="1.499cm"/>
    </style:list-level-properties>
    <style:text-properties fo:font-family="OpenSymbol"/>
   </text:list-level-style-bullet>
   <text:list-level-style-bullet text:level="6" text:style-name="Numbering_20_Symbols" text:bullet-char="–">
    <style:list-level-properties text:list-level-position-and-space-mode="label-alignment">
     <style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="1.799cm" fo:text-indent="-0.3cm" fo:margin-left="1.799cm"/>
    </style:list-level-properties>
    <style:text-properties fo:font-family="OpenSymbol"/>
   </text:list-level-style-bullet>
   <text:list-level-style-bullet text:level="7" text:style-name="Numbering_20_Symbols" text:bullet-char="–">
    <style:list-level-properties text:list-level-position-and-space-mode="label-alignment">
     <style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="2.101cm" fo:text-indent="-0.3cm" fo:margin-left="2.101cm"/>
    </style:list-level-properties>
    <style:text-properties fo:font-family="OpenSymbol"/>
   </text:list-level-style-bullet>
   <text:list-level-style-bullet text:level="8" text:style-name="Numbering_20_Symbols" text:bullet-char="–">
    <style:list-level-properties text:list-level-position-and-space-mode="label-alignment">
     <style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="2.401cm" fo:text-indent="-0.3cm" fo:margin-left="2.401cm"/>
    </style:list-level-properties>
    <style:text-properties fo:font-family="OpenSymbol"/>
   </text:list-level-style-bullet>
   <text:list-level-style-bullet text:level="9" text:style-name="Numbering_20_Symbols" text:bullet-char="–">
    <style:list-level-properties text:list-level-position-and-space-mode="label-alignment">
     <style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="2.701cm" fo:text-indent="-0.3cm" fo:margin-left="2.701cm"/>
    </style:list-level-properties>
    <style:text-properties fo:font-family="OpenSymbol"/>
   </text:list-level-style-bullet>
   <text:list-level-style-bullet text:level="10" text:style-name="Numbering_20_Symbols" text:bullet-char="–">
    <style:list-level-properties text:list-level-position-and-space-mode="label-alignment">
     <style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="3cm" fo:text-indent="-0.3cm" fo:margin-left="3cm"/>
    </style:list-level-properties>
    <style:text-properties fo:font-family="OpenSymbol"/>
   </text:list-level-style-bullet>
  </text:list-style>
  <text:list-style style:name="List_20_3" style:display-name="List 3">
   <text:list-level-style-bullet text:level="1" text:style-name="Numbering_20_Symbols" text:bullet-char="☑">
    <style:list-level-properties text:list-level-position-and-space-mode="label-alignment">
     <style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="0.395cm" fo:text-indent="-0.395cm" fo:margin-left="0.395cm"/>
    </style:list-level-properties>
    <style:text-properties fo:font-family="OpenSymbol"/>
   </text:list-level-style-bullet>
   <text:list-level-style-bullet text:level="2" text:style-name="Numbering_20_Symbols" text:bullet-char="□">
    <style:list-level-properties text:list-level-position-and-space-mode="label-alignment">
     <style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="0.79cm" fo:text-indent="-0.395cm" fo:margin-left="0.79cm"/>
    </style:list-level-properties>
    <style:text-properties fo:font-family="OpenSymbol"/>
   </text:list-level-style-bullet>
   <text:list-level-style-bullet text:level="3" text:style-name="Numbering_20_Symbols" text:bullet-char="☑">
    <style:list-level-properties text:list-level-position-and-space-mode="label-alignment">
     <style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="0.395cm" fo:text-indent="-0.395cm" fo:margin-left="0.395cm"/>
    </style:list-level-properties>
    <style:text-properties fo:font-family="OpenSymbol"/>
   </text:list-level-style-bullet>
   <text:list-level-style-bullet text:level="4" text:style-name="Numbering_20_Symbols" text:bullet-char="□">
    <style:list-level-properties text:list-level-position-and-space-mode="label-alignment">
     <style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="0.79cm" fo:text-indent="-0.395cm" fo:margin-left="0.79cm"/>
    </style:list-level-properties>
    <style:text-properties fo:font-family="OpenSymbol"/>
   </text:list-level-style-bullet>
   <text:list-level-style-bullet text:level="5" text:style-name="Numbering_20_Symbols" text:bullet-char="☑">
    <style:list-level-properties text:list-level-position-and-space-mode="label-alignment">
     <style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="0.395cm" fo:text-indent="-0.395cm" fo:margin-left="0.395cm"/>
    </style:list-level-properties>
    <style:text-properties fo:font-family="OpenSymbol"/>
   </text:list-level-style-bullet>
   <text:list-level-style-bullet text:level="6" text:style-name="Numbering_20_Symbols" text:bullet-char="□">
    <style:list-level-properties text:list-level-position-and-space-mode="label-alignment">
     <style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="0.79cm" fo:text-indent="-0.395cm" fo:margin-left="0.79cm"/>
    </style:list-level-properties>
    <style:text-properties fo:font-family="OpenSymbol"/>
   </text:list-level-style-bullet>
   <text:list-level-style-bullet text:level="7" text:style-name="Numbering_20_Symbols" text:bullet-char="☑">
    <style:list-level-properties text:list-level-position-and-space-mode="label-alignment">
     <style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="0.395cm" fo:text-indent="-0.395cm" fo:margin-left="0.395cm"/>
    </style:list-level-properties>
    <style:text-properties fo:font-family="OpenSymbol"/>
   </text:list-level-style-bullet>
   <text:list-level-style-bullet text:level="8" text:style-name="Numbering_20_Symbols" text:bullet-char="□">
    <style:list-level-properties text:list-level-position-and-space-mode="label-alignment">
     <style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="0.79cm" fo:text-indent="-0.395cm" fo:margin-left="0.79cm"/>
    </style:list-level-properties>
    <style:text-properties fo:font-family="OpenSymbol"/>
   </text:list-level-style-bullet>
   <text:list-level-style-bullet text:level="9" text:style-name="Numbering_20_Symbols" text:bullet-char="☑">
    <style:list-level-properties text:list-level-position-and-space-mode="label-alignment">
     <style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="0.395cm" fo:text-indent="-0.395cm" fo:margin-left="0.395cm"/>
    </style:list-level-properties>
    <style:text-properties fo:font-family="OpenSymbol"/>
   </text:list-level-style-bullet>
   <text:list-level-style-bullet text:level="10" text:style-name="Numbering_20_Symbols" text:bullet-char="□">
    <style:list-level-properties text:list-level-position-and-space-mode="label-alignment">
     <style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="0.79cm" fo:text-indent="-0.395cm" fo:margin-left="0.79cm"/>
    </style:list-level-properties>
    <style:text-properties fo:font-family="OpenSymbol"/>
   </text:list-level-style-bullet>
  </text:list-style>
  <text:list-style style:name="List_20_4" style:display-name="List 4">
   <text:list-level-style-bullet text:level="1" text:style-name="Numbering_20_Symbols" text:bullet-char="➢">
    <style:list-level-properties text:min-label-width="0.4cm"/>
    <style:text-properties fo:font-family="StarSymbol"/>
   </text:list-level-style-bullet>
   <text:list-level-style-bullet text:level="2" text:style-name="Numbering_20_Symbols" text:bullet-char="">
    <style:list-level-properties text:space-before="0.401cm" text:min-label-width="0.4cm"/>
    <style:text-properties fo:font-family="StarSymbol"/>
   </text:list-level-style-bullet>
   <text:list-level-style-bullet text:level="3" text:style-name="Numbering_20_Symbols" text:bullet-char="">
    <style:list-level-properties text:space-before="0.799cm" text:min-label-width="0.4cm"/>
    <style:text-properties fo:font-family="StarSymbol"/>
   </text:list-level-style-bullet>
   <text:list-level-style-bullet text:level="4" text:style-name="Numbering_20_Symbols" text:bullet-char="">
    <style:list-level-properties text:space-before="1.2cm" text:min-label-width="0.4cm"/>
    <style:text-properties fo:font-family="StarSymbol"/>
   </text:list-level-style-bullet>
   <text:list-level-style-bullet text:level="5" text:style-name="Numbering_20_Symbols" text:bullet-char="">
    <style:list-level-properties text:space-before="1.6cm" text:min-label-width="0.4cm"/>
    <style:text-properties fo:font-family="StarSymbol"/>
   </text:list-level-style-bullet>
   <text:list-level-style-bullet text:level="6" text:style-name="Numbering_20_Symbols" text:bullet-char="">
    <style:list-level-properties text:space-before="2.001cm" text:min-label-width="0.4cm"/>
    <style:text-properties fo:font-family="StarSymbol"/>
   </text:list-level-style-bullet>
   <text:list-level-style-bullet text:level="7" text:style-name="Numbering_20_Symbols" text:bullet-char="">
    <style:list-level-properties text:space-before="2.399cm" text:min-label-width="0.4cm"/>
    <style:text-properties fo:font-family="StarSymbol"/>
   </text:list-level-style-bullet>
   <text:list-level-style-bullet text:level="8" text:style-name="Numbering_20_Symbols" text:bullet-char="">
    <style:list-level-properties text:space-before="2.8cm" text:min-label-width="0.4cm"/>
    <style:text-properties fo:font-family="StarSymbol"/>
   </text:list-level-style-bullet>
   <text:list-level-style-bullet text:level="9" text:style-name="Numbering_20_Symbols" text:bullet-char="">
    <style:list-level-properties text:space-before="3.2cm" text:min-label-width="0.4cm"/>
    <style:text-properties fo:font-family="StarSymbol"/>
   </text:list-level-style-bullet>
   <text:list-level-style-bullet text:level="10" text:style-name="Numbering_20_Symbols" text:bullet-char="">
    <style:list-level-properties text:space-before="3.601cm" text:min-label-width="0.4cm"/>
    <style:text-properties fo:font-family="StarSymbol"/>
   </text:list-level-style-bullet>
  </text:list-style>
  <text:list-style style:name="List_20_5" style:display-name="List 5">
   <text:list-level-style-bullet text:level="1" text:style-name="Numbering_20_Symbols" text:bullet-char="✗">
    <style:list-level-properties text:min-label-width="0.4cm"/>
    <style:text-properties fo:font-family="StarSymbol"/>
   </text:list-level-style-bullet>
   <text:list-level-style-bullet text:level="2" text:style-name="Numbering_20_Symbols" text:bullet-char="✗">
    <style:list-level-properties text:space-before="0.401cm" text:min-label-width="0.4cm"/>
    <style:text-properties fo:font-family="StarSymbol"/>
   </text:list-level-style-bullet>
   <text:list-level-style-bullet text:level="3" text:style-name="Numbering_20_Symbols" text:bullet-char="✗">
    <style:list-level-properties text:space-before="0.799cm" text:min-label-width="0.4cm"/>
    <style:text-properties fo:font-family="StarSymbol"/>
   </text:list-level-style-bullet>
   <text:list-level-style-bullet text:level="4" text:style-name="Numbering_20_Symbols" text:bullet-char="✗">
    <style:list-level-properties text:space-before="1.2cm" text:min-label-width="0.4cm"/>
    <style:text-properties fo:font-family="StarSymbol"/>
   </text:list-level-style-bullet>
   <text:list-level-style-bullet text:level="5" text:style-name="Numbering_20_Symbols" text:bullet-char="✗">
    <style:list-level-properties text:space-before="1.6cm" text:min-label-width="0.4cm"/>
    <style:text-properties fo:font-family="StarSymbol"/>
   </text:list-level-style-bullet>
   <text:list-level-style-bullet text:level="6" text:style-name="Numbering_20_Symbols" text:bullet-char="✗">
    <style:list-level-properties text:space-before="2.001cm" text:min-label-width="0.4cm"/>
    <style:text-properties fo:font-family="StarSymbol"/>
   </text:list-level-style-bullet>
   <text:list-level-style-bullet text:level="7" text:style-name="Numbering_20_Symbols" text:bullet-char="✗">
    <style:list-level-properties text:space-before="2.399cm" text:min-label-width="0.4cm"/>
    <style:text-properties fo:font-family="StarSymbol"/>
   </text:list-level-style-bullet>
   <text:list-level-style-bullet text:level="8" text:style-name="Numbering_20_Symbols" text:bullet-char="✗">
    <style:list-level-properties text:space-before="2.8cm" text:min-label-width="0.4cm"/>
    <style:text-properties fo:font-family="StarSymbol"/>
   </text:list-level-style-bullet>
   <text:list-level-style-bullet text:level="9" text:style-name="Numbering_20_Symbols" text:bullet-char="✗">
    <style:list-level-properties text:space-before="3.2cm" text:min-label-width="0.4cm"/>
    <style:text-properties fo:font-family="StarSymbol"/>
   </text:list-level-style-bullet>
   <text:list-level-style-bullet text:level="10" text:style-name="Numbering_20_Symbols" text:bullet-char="✗">
    <style:list-level-properties text:space-before="3.601cm" text:min-label-width="0.4cm"/>
    <style:text-properties fo:font-family="StarSymbol"/>
   </text:list-level-style-bullet>
  </text:list-style>
  <text:notes-configuration text:note-class="footnote" text:citation-style-name="Footnote_20_Symbol" text:citation-body-style-name="Footnote_20_anchor" style:num-format="1" text:start-value="0" text:footnotes-position="page" text:start-numbering-at="document"/>
  <text:notes-configuration text:note-class="endnote" style:num-format="i" text:start-value="0"/>
  <text:linenumbering-configuration text:number-lines="false" text:offset="0.499cm" style:num-format="1" text:number-position="left" text:increment="5"/>
 </office:styles>
 <office:automatic-styles>
  <style:page-layout style:name="Mpm1">
   <style:page-layout-properties fo:page-width="21.001cm" fo:page-height="29.7cm" style:num-format="1" style:paper-tray-name="[From printer settings]" style:print-orientation="portrait" fo:margin-top="3.2cm" fo:margin-bottom="2.499cm" fo:margin-left="2.701cm" fo:margin-right="3cm" style:writing-mode="lr-tb" style:footnote-max-height="0cm">
    <style:columns fo:column-count="1" fo:column-gap="0cm"/>
    <style:footnote-sep style:width="0.018cm" style:distance-before-sep="0.101cm" style:distance-after-sep="0.101cm" style:adjustment="left" style:rel-width="25%" style:color="#000000"/>
   </style:page-layout-properties>
   <style:header-style/>
   <style:footer-style>
    <style:header-footer-properties fo:min-height="1.699cm" fo:margin-left="0cm" fo:margin-right="0cm" fo:margin-top="1.199cm" style:shadow="none" style:dynamic-spacing="false"/>
   </style:footer-style>
  </style:page-layout>
  <style:page-layout style:name="Mpm2">
   <style:page-layout-properties fo:page-width="21.001cm" fo:page-height="29.7cm" style:num-format="1" style:print-orientation="portrait" fo:margin-top="2cm" fo:margin-bottom="2cm" fo:margin-left="2.499cm" fo:margin-right="2.499cm" style:shadow="none" fo:background-color="transparent" style:writing-mode="lr-tb" style:footnote-max-height="0cm">
    <style:background-image/>
    <style:columns fo:column-count="1" fo:column-gap="0cm"/>
    <style:footnote-sep style:width="0.018cm" style:distance-before-sep="0.101cm" style:distance-after-sep="0.101cm" style:adjustment="left" style:rel-width="25%" style:color="#000000"/>
   </style:page-layout-properties>
   <style:header-style/>
   <style:footer-style/>
  </style:page-layout>
  <style:page-layout style:name="Mpm3" style:page-usage="left">
   <style:page-layout-properties fo:page-width="21.001cm" fo:page-height="29.7cm" style:num-format="1" style:print-orientation="portrait" fo:margin-top="2cm" fo:margin-bottom="1cm" fo:margin-left="2.499cm" fo:margin-right="4.5cm" style:writing-mode="lr-tb" style:footnote-max-height="0cm">
    <style:footnote-sep style:width="0.018cm" style:distance-before-sep="0.101cm" style:distance-after-sep="0.101cm" style:adjustment="left" style:rel-width="25%" style:color="#000000"/>
   </style:page-layout-properties>
   <style:header-style/>
   <style:footer-style/>
  </style:page-layout>
  <style:page-layout style:name="Mpm4" style:page-usage="right">
   <style:page-layout-properties fo:page-width="21.001cm" fo:page-height="29.7cm" style:num-format="1" style:print-orientation="portrait" fo:margin-top="2cm" fo:margin-bottom="1cm" fo:margin-left="2.499cm" fo:margin-right="4.5cm" style:writing-mode="lr-tb" style:footnote-max-height="0cm">
    <style:footnote-sep style:width="0.018cm" style:distance-before-sep="0.101cm" style:distance-after-sep="0.101cm" style:adjustment="left" style:rel-width="25%" style:color="#000000"/>
   </style:page-layout-properties>
   <style:header-style/>
   <style:footer-style/>
  </style:page-layout>
  <style:page-layout style:name="Mpm5">
   <style:page-layout-properties fo:page-width="22.721cm" fo:page-height="11.4cm" style:num-format="1" style:print-orientation="landscape" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:margin-left="0cm" fo:margin-right="0cm" style:writing-mode="lr-tb" style:footnote-max-height="0cm">
    <style:footnote-sep style:width="0.018cm" style:distance-before-sep="0.101cm" style:distance-after-sep="0.101cm" style:adjustment="left" style:rel-width="25%" style:color="#000000"/>
   </style:page-layout-properties>
   <style:header-style/>
   <style:footer-style/>
  </style:page-layout>
  <style:page-layout style:name="Mpm6">
   <style:page-layout-properties fo:page-width="14.801cm" fo:page-height="21.001cm" style:num-format="1" style:print-orientation="portrait" fo:margin-top="2cm" fo:margin-bottom="2cm" fo:margin-left="2cm" fo:margin-right="2cm" style:writing-mode="lr-tb" style:footnote-max-height="0cm">
    <style:footnote-sep style:width="0.018cm" style:distance-before-sep="0.101cm" style:distance-after-sep="0.101cm" style:adjustment="left" style:rel-width="25%" style:color="#000000"/>
   </style:page-layout-properties>
   <style:header-style/>
   <style:footer-style/>
  </style:page-layout>
  <style:page-layout style:name="Mpm7">
   <style:page-layout-properties fo:page-width="20.999cm" fo:page-height="29.699cm" style:num-format="1" style:print-orientation="portrait" fo:margin-top="2cm" fo:margin-bottom="2cm" fo:margin-left="2cm" fo:margin-right="2cm" style:writing-mode="lr-tb" style:footnote-max-height="0cm">
    <style:footnote-sep style:adjustment="left" style:rel-width="25%" style:color="#000000"/>
   </style:page-layout-properties>
   <style:header-style/>
   <style:footer-style/>
  </style:page-layout>
 </office:automatic-styles>
 <office:master-styles>
  <style:master-page style:name="Standard" style:page-layout-name="Mpm1">
   <style:footer>
    <text:p text:style-name="Footer"><text:page-number text:select-page="current">14</text:page-number></text:p>
   </style:footer>
  </style:master-page>
  <style:master-page style:name="First_20_Page" style:display-name="First Page" style:page-layout-name="Mpm2" style:next-style-name="Standard"/>
  <style:master-page style:name="Left_20_Page" style:display-name="Left Page" style:page-layout-name="Mpm3"/>
  <style:master-page style:name="Right_20_Page" style:display-name="Right Page" style:page-layout-name="Mpm4"/>
  <style:master-page style:name="Envelope" style:page-layout-name="Mpm5"/>
  <style:master-page style:name="Index" style:page-layout-name="Mpm6" style:next-style-name="Standard"/>
  <style:master-page style:name="Endnote" style:page-layout-name="Mpm7"/>
 </office:master-styles>
</office:document-styles>
'''
    _MIMETYPE = 'application/vnd.oasis.opendocument.text'

    def set_up(self):
        """Create a temporary directory containing the internal 
        structure of an ODT file except 'content.xml'.
        """

        # Generate the common ODF components.

        message = OdfFile.set_up(self)

        if message.startswith('ERROR'):
            return message

        # Generate manifest.rdf

        try:
            with open(self.TEMPDIR + '/manifest.rdf', 'w', encoding='utf-8') as f:
                f.write(self._MANIFEST_RDF)
        except:
            return 'ERROR: Cannot write "manifest.rdf"'

        return 'SUCCESS: ODT structure generated.'

    def convert_from_yw(self, text):
        """Convert yw7 raw markup to odt. Return an xml string.
        """

        ODT_REPLACEMENTS = [
            ['&', '&amp;'],
            ['>', '&gt;'],
            ['<', '&lt;'],
            ['\n', '</text:p>\n<text:p text:style-name="First_20_line_20_indent">'],
            ['[i]', '<text:span text:style-name="Emphasis">'],
            ['[/i]', '</text:span>'],
            ['[b]', '<text:span text:style-name="Strong_20_Emphasis">'],
            ['[/b]', '</text:span>'],
            ['/*', '<office:annotation><dc:creator>' +
                self.author + '</dc:creator><text:p>'],
            ['*/', '</text:p></office:annotation>'],
        ]

        try:

            # process italics and bold markup reaching across linebreaks

            italics = False
            bold = False
            newlines = []
            lines = text.split('\n')
            for line in lines:
                if italics:
                    line = '[i]' + line
                    italics = False

                while line.count('[i]') > line.count('[/i]'):
                    line += '[/i]'
                    italics = True

                while line.count('[/i]') > line.count('[i]'):
                    line = '[i]' + line

                line = line.replace('[i][/i]', '')

                if bold:
                    line = '[b]' + line
                    bold = False

                while line.count('[b]') > line.count('[/b]'):
                    line += '[/b]'
                    bold = True

                while line.count('[/b]') > line.count('[b]'):
                    line = '[b]' + line

                line = line.replace('[b][/b]', '')

                newlines.append(line)

            text = '\n'.join(newlines).rstrip()

            for r in ODT_REPLACEMENTS:
                text = text.replace(r[0], r[1])

            # Remove highlighting, alignment,
            # strikethrough, and underline tags.

            text = re.sub('\[\/*[h|c|r|s|u]\d*\]', '', text)

        except AttributeError:
            text = ''

        return text


class OdtProof(OdtFile):
    """ODT proof reading file representation.

    Export a manuscript with visibly tagged chapters and scenes.
    """

    DESCRIPTION = 'Tagged manuscript for proofing'
    SUFFIX = '_proof'

    fileHeader = OdtFile.CONTENT_XML_HEADER + '''<text:p text:style-name="Title">$Title</text:p>
<text:p text:style-name="Subtitle">$AuthorName</text:p>
'''

    partTemplate = '''<text:p text:style-name="yWriter_20_mark">[ChID:$ID]</text:p>
<text:h text:style-name="Heading_20_1" text:outline-level="1">$Title</text:h>
'''

    chapterTemplate = '''<text:p text:style-name="yWriter_20_mark">[ChID:$ID]</text:p>
<text:h text:style-name="Heading_20_2" text:outline-level="2">$Title</text:h>
'''

    unusedChapterTemplate = '''<text:p text:style-name="yWriter_20_mark_20_unused">[ChID:$ID (Unused)]</text:p>
<text:h text:style-name="Heading_20_2" text:outline-level="2">$Title</text:h>
'''

    notesChapterTemplate = '''<text:p text:style-name="yWriter_20_mark_20_notes">[ChID:$ID (Notes)]</text:p>
<text:h text:style-name="Heading_20_2" text:outline-level="2">$Title</text:h>
'''

    todoChapterTemplate = '''<text:p text:style-name="yWriter_20_mark_20_todo">[ChID:$ID (ToDo)]</text:p>
<text:h text:style-name="Heading_20_2" text:outline-level="2">$Title</text:h>
'''

    sceneTemplate = '''<text:p text:style-name="yWriter_20_mark">[ScID:$ID]</text:p>
<text:p text:style-name="Text_20_body">$SceneContent</text:p>
<text:p text:style-name="yWriter_20_mark">[/ScID]</text:p>
'''

    unusedSceneTemplate = '''<text:p text:style-name="yWriter_20_mark_20_unused">[ScID:$ID (Unused)]</text:p>
<text:p text:style-name="Text_20_body">$SceneContent</text:p>
<text:p text:style-name="yWriter_20_mark_20_unused">[/ScID (Unused)]</text:p>
'''

    notesSceneTemplate = '''<text:p text:style-name="yWriter_20_mark_20_notes">[ScID:$ID (Notes)]</text:p>
<text:p text:style-name="Text_20_body">$SceneContent</text:p>
<text:p text:style-name="yWriter_20_mark_20_notes">[/ScID (Notes)]</text:p>
'''

    todoSceneTemplate = '''<text:p text:style-name="yWriter_20_mark_20_todo">[ScID:$ID (ToDo)]</text:p>
<text:p text:style-name="Text_20_body">$SceneContent</text:p>
<text:p text:style-name="yWriter_20_mark_20_todo">[/ScID (ToDo)]</text:p>
'''

    sceneDivider = '''<text:p text:style-name="Heading_20_4">* * *</text:p>
'''

    chapterEndTemplate = '''<text:p text:style-name="yWriter_20_mark">[/ChID]</text:p>
'''

    unusedChapterEndTemplate = '''<text:p text:style-name="yWriter_20_mark_20_unused">[/ChID (Unused)]</text:p>
'''

    notesChapterEndTemplate = '''<text:p text:style-name="yWriter_20_mark_20_notes">[/ChID (Notes)]</text:p>
'''

    todoChapterEndTemplate = '''<text:p text:style-name="yWriter_20_mark_20_todo">[/ChID (ToDo)]</text:p>
'''

    fileFooter = OdtFile.CONTENT_XML_FOOTER


class OdtManuscript(OdtFile):
    """ODT manuscript file representation.

    Export a manuscript with invisibly tagged chapters and scenes.
    """

    DESCRIPTION = 'Editable manuscript'
    SUFFIX = '_manuscript'

    fileHeader = OdtFile.CONTENT_XML_HEADER + '''<text:p text:style-name="Title">$Title</text:p>
<text:p text:style-name="Subtitle">$AuthorName</text:p>
'''

    partTemplate = '''<text:section text:style-name="Sect1" text:name="ChID:$ID">
<text:h text:style-name="Heading_20_1" text:outline-level="1"><text:a xlink:href="../${ProjectName}_parts.odt#ChID:$ID%7Cregion">$Title</text:a></text:h>
'''

    chapterTemplate = '''<text:section text:style-name="Sect1" text:name="ChID:$ID">
<text:h text:style-name="Heading_20_2" text:outline-level="2"><text:a xlink:href="../${ProjectName}_chapters.odt#ChID:$ID%7Cregion">$Title</text:a></text:h>
'''

    sceneTemplate = '''<text:section text:style-name="Sect1" text:name="ScID:$ID">
<text:p text:style-name="Text_20_body"><office:annotation>
<dc:creator>scene title</dc:creator>
<text:p>- $Title</text:p>
<text:p/>
<text:p><text:a xlink:href="../${ProjectName}_scenes.odt#ScID:$ID%7Cregion">→Summary</text:a> -</text:p>
</office:annotation>$SceneContent</text:p>
</text:section>
'''

    appendedSceneTemplate = '''<text:section text:style-name="Sect1" text:name="ScID:$ID">
<text:p text:style-name="First_20_line_20_indent"><office:annotation>
<dc:creator>scene title</dc:creator>
<text:p>- $Title</text:p>
<text:p/>
<text:p><text:a xlink:href="../${ProjectName}_scenes.odt#ScID:$ID%7Cregion">→Summary</text:a> -</text:p>
</office:annotation>$SceneContent</text:p>
</text:section>
'''

    sceneDivider = '<text:p ></text:p>'
    #sceneDivider = '<text:p text:style-name="Heading_20_4">* * *</text:p>'

    chapterEndTemplate = '''</text:section>
'''

    fileFooter = OdtFile.CONTENT_XML_FOOTER

    def get_chapterMapping(self, chId, chapterNumber):
        """Return a mapping dictionary for a chapter section. 
        """
        chapterMapping = OdtFile.get_chapterMapping(self, chId, chapterNumber)

        if self.chapters[chId].suppressChapterTitle:
            chapterMapping['Title'] = ''

        return chapterMapping


class OdtSceneDesc(OdtFile):
    """ODT scene summaries file representation.

    Export a full synopsis with invisibly tagged scene descriptions.
    """

    DESCRIPTION = 'Scene descriptions'
    SUFFIX = '_scenes'

    fileHeader = OdtFile.CONTENT_XML_HEADER + '''<text:p text:style-name="Title">$Title</text:p>
<text:p text:style-name="Subtitle">$AuthorName</text:p>
'''

    partTemplate = '''<text:section text:style-name="Sect1" text:name="ChID:$ID">
<text:h text:style-name="Heading_20_1" text:outline-level="1"><text:a xlink:href="../${ProjectName}_parts.odt#ChID:$ID%7Cregion">$Title</text:a></text:h>
'''

    chapterTemplate = '''<text:section text:style-name="Sect1" text:name="ChID:$ID">
<text:h text:style-name="Heading_20_2" text:outline-level="2"><text:a xlink:href="../${ProjectName}_chapters.odt#ChID:$ID%7Cregion">$Title</text:a></text:h>
'''

    sceneTemplate = '''<text:section text:style-name="Sect1" text:name="ScID:$ID">
<text:p text:style-name="Text_20_body"><office:annotation>
<dc:creator>scene title</dc:creator>
<text:p>- $Title</text:p>
<text:p/>
<text:p><text:a xlink:href="../${ProjectName}_manuscript.odt#ScID:$ID%7Cregion">→Manuscript</text:a> -</text:p>
</office:annotation>$Desc</text:p>
</text:section>
'''

    appendedSceneTemplate = '''<text:section text:style-name="Sect1" text:name="ScID:$ID">
<text:p text:style-name="First_20_line_20_indent"><office:annotation>
<dc:creator>scene title</dc:creator>
<text:p>- $Title</text:p>
<text:p/>
<text:p><text:a xlink:href="../${ProjectName}_manuscript.odt#ScID:$ID%7Cregion">→Manuscript</text:a> -</text:p>
</office:annotation>$Desc</text:p>
</text:section>
'''

    sceneDivider = '''<text:p text:style-name="Heading_20_4">* * *</text:p>
'''

    chapterEndTemplate = '''</text:section>
'''

    fileFooter = OdtFile.CONTENT_XML_FOOTER


class OdtChapterDesc(OdtFile):
    """ODT chapter summaries file representation.

    Export a brief synopsis with invisibly tagged chapter descriptions.
    """

    DESCRIPTION = 'Chapter descriptions'
    SUFFIX = '_chapters'

    fileHeader = OdtFile.CONTENT_XML_HEADER + '''<text:p text:style-name="Title">$Title</text:p>
<text:p text:style-name="Subtitle">$AuthorName</text:p>
'''

    partTemplate = '''<text:h text:style-name="Heading_20_1" text:outline-level="1"><text:a xlink:href="../${ProjectName}_parts.odt#ChID:$ID%7Cregion">$Title</text:a></text:h>
'''

    chapterTemplate = '''<text:h text:style-name="Heading_20_2" text:outline-level="2"><text:a xlink:href="../${ProjectName}_manuscript.odt#ChID:$ID%7Cregion">$Title</text:a></text:h>
<text:section text:style-name="Sect1" text:name="ChID:$ID">
<text:p text:style-name="Text_20_body">$Desc</text:p>
</text:section>
'''

    fileFooter = OdtFile.CONTENT_XML_FOOTER


class OdtPartDesc(OdtFile):
    """ODT part summaries file representation.

    Export a very brief synopsis with invisibly tagged part descriptions.
    """

    DESCRIPTION = 'Part descriptions'
    SUFFIX = '_parts'

    fileHeader = OdtFile.CONTENT_XML_HEADER + '''<text:p text:style-name="Title">$Title</text:p>
<text:p text:style-name="Subtitle">$AuthorName</text:p>
'''

    partTemplate = '''<text:h text:style-name="Heading_20_1" text:outline-level="1"><text:a xlink:href="../${ProjectName}_manuscript.odt#ChID:$ID%7Cregion">$Title</text:a></text:h>
<text:section text:style-name="Sect1" text:name="ChID:$ID">
<text:p text:style-name="Text_20_body">$Desc</text:p>
</text:section>
'''

    fileFooter = OdtFile.CONTENT_XML_FOOTER
from string import Template



class CrossReferences():
    """Create dictionaries containing a novel's cross references:

    - Characters per tag
    - Locations per tag
    - Items per tag
    - Scenes per character
    - Scenes per location
    - Scenes per item
    - Scenes per tag
    """

    def __init__(self):
        # Cross reference dictionaries:

        self.scnPerChr = {}
        # dict
        # key = character ID, value: list of scene IDs
        # Scenes per character

        self.scnPerLoc = {}
        # dict
        # key = location ID, value: list of scene IDs
        # Scenes per location

        self.scnPerItm = {}
        # dict
        # key = item ID, value: list of scene IDs
        # Scenes per item

        self.scnPerTag = {}
        # dict
        # key = tag, value: list of scene IDs
        # Scenes per tag

        self.chrPerTag = {}
        # dict
        # key = tag, value: list of character IDs
        # Characters per tag

        self.locPerTag = {}
        # dict
        # key = tag, value: list of location IDs
        # Locations per tag

        self.itmPerTag = {}
        # dict
        # key = tag, value: list of item IDs
        # Items per tag

        self.chpPerScn = {}
        # dict
        # key = scene ID, value: chapter ID
        # Chapter to which the scene belongs

        self.srtScenes = None
        # list of str
        # scene IDs in the overall order

    def generate_xref(self, novel):
        """Generate cross references."""
        self.scnPerChr = {}
        self.scnPerLoc = {}
        self.scnPerItm = {}
        self.scnPerTag = {}
        self.chrPerTag = {}
        self.locPerTag = {}
        self.itmPerTag = {}
        self.chpPerScn = {}
        self.srtScenes = []

        # Characters per tag:

        for crId in novel.srtCharacters:
            self.scnPerChr[crId] = []

            if novel.characters[crId].tags:

                for tag in novel.characters[crId].tags:

                    if not tag in self.chrPerTag:
                        self.chrPerTag[tag] = []

                    self.chrPerTag[tag].append(crId)

        # Locations per tag:

        for lcId in novel.srtLocations:
            self.scnPerLoc[lcId] = []

            if novel.locations[lcId].tags:

                for tag in novel.locations[lcId].tags:

                    if not tag in self.locPerTag:
                        self.locPerTag[tag] = []

                    self.locPerTag[tag].append(lcId)

        # Items per tag:

        for itId in novel.srtItems:
            self.scnPerItm[itId] = []

            if novel.items[itId].tags:

                for tag in novel.items[itId].tags:

                    if not tag in self.itmPerTag:
                        self.itmPerTag[tag] = []

                    self.itmPerTag[tag].append(itId)

        for chId in novel.srtChapters:

            for scId in novel.chapters[chId].srtScenes:
                self.srtScenes.append(scId)
                self.chpPerScn[scId] = chId

                # Scenes per character:

                if novel.scenes[scId].characters:

                    for crId in novel.scenes[scId].characters:
                        self.scnPerChr[crId].append(scId)

                # Scenes per location:

                if novel.scenes[scId].locations:

                    for lcId in novel.scenes[scId].locations:
                        self.scnPerLoc[lcId].append(scId)

                # Scenes per item:

                if novel.scenes[scId].items:

                    for itId in novel.scenes[scId].items:
                        self.scnPerItm[itId].append(scId)

                # Scenes per tag:

                if novel.scenes[scId].tags:

                    for tag in novel.scenes[scId].tags:

                        if not tag in self.scnPerTag:
                            self.scnPerTag[tag] = []

                        self.scnPerTag[tag].append(scId)


class OdtXref(OdtFile):
    """OpenDocument xml cross reference file representation."""

    DESCRIPTION = 'Cross reference'
    SUFFIX = '_xref'

    fileHeader = OdtFile.CONTENT_XML_HEADER + '''<text:p text:style-name="Title">$Title</text:p>
<text:p text:style-name="Subtitle">$AuthorName</text:p>
'''
    sceneTemplate = '''<text:p text:style-name="yWriter_20_mark">
<text:a xlink:href="../${ProjectName}_manuscript.odt#ScID:$ID%7Cregion">$SceneNumber</text:a> (Ch $Chapter) $Title
</text:p>
'''
    unusedSceneTemplate = '''<text:p text:style-name="yWriter_20_mark_20_unused">
$SceneNumber (Ch $Chapter) $Title (Unused)
</text:p>
'''
    notesSceneTemplate = '''<text:p text:style-name="yWriter_20_mark_20_notes">
$SceneNumber (Ch $Chapter) $Title (Notes)
</text:p>
'''
    todoSceneTemplate = '''<text:p text:style-name="yWriter_20_mark_20_todo">
$SceneNumber (Ch $Chapter) $Title (ToDo)
</text:p>
'''
    characterTemplate = '''<text:p text:style-name="Text_20_body">
<text:a xlink:href="../${ProjectName}_characters.odt#CrID:$ID%7Cregion">$Title</text:a> $FullName
</text:p>
'''
    locationTemplate = '''<text:p text:style-name="Text_20_body">
<text:a xlink:href="../${ProjectName}_locations.odt#LcID:$ID%7Cregion">$Title</text:a>
</text:p>
'''
    itemTemplate = '''<text:p text:style-name="Text_20_body">
<text:a xlink:href="../${ProjectName}_items.odt#ItrID:$ID%7Cregion">$Title</text:a>
</text:p>
'''
    scnPerChrTemplate = '''<text:h text:style-name="Heading_20_2" text:outline-level="2">Scenes with Character $Title:</text:h>
'''
    scnPerLocTemplate = '''<text:h text:style-name="Heading_20_2" text:outline-level="2">Scenes with Location $Title:</text:h>
'''
    scnPerItmTemplate = '''<text:h text:style-name="Heading_20_2" text:outline-level="2">Scenes with Item $Title:</text:h>
'''
    chrPerTagTemplate = '''<text:h text:style-name="Heading_20_2" text:outline-level="2">Characters tagged $Tag:</text:h>
'''
    locPerTagTemplate = '''<text:h text:style-name="Heading_20_2" text:outline-level="2">Locations tagged $Tag:</text:h>
'''
    itmPerTagTemplate = '''<text:h text:style-name="Heading_20_2" text:outline-level="2">Items tagged $Tag:</text:h>
'''
    scnPerTagtemplate = '''<text:h text:style-name="Heading_20_2" text:outline-level="2">Scenes tagged $Tag:</text:h>
'''
    fileFooter = OdtFile.CONTENT_XML_FOOTER

    def __init__(self, filePath, **kwargs):
        """Apply the strategy pattern 
        by delegating the cross reference to an external object.
        """
        OdtFile.__init__(self, filePath)
        self.xr = CrossReferences()

    def get_sceneMapping(self, scId):
        """Add the chapter number to the original mapping dictionary.
        """
        sceneNumber = self.xr.srtScenes.index(scId) + 1
        sceneMapping = OdtFile.get_sceneMapping(self, scId, sceneNumber, 0, 0)
        chapterNumber = self.srtChapters.index(self.xr.chpPerScn[scId]) + 1
        sceneMapping['Chapter'] = str(chapterNumber)
        return sceneMapping

    def get_tagMapping(self, tag):
        """Return a mapping dictionary for a tags section. 
        """
        tagMapping = dict(
            Tag=tag,
        )
        return tagMapping

    def get_scenes(self, scenes):
        """Process the scenes.
        Return a list of strings.
        """
        lines = []

        for scId in scenes:

            if self.scenes[scId].isNotesScene:
                template = Template(self.notesSceneTemplate)

            elif self.scenes[scId].isTodoScene:
                template = Template(self.todoSceneTemplate)

            elif self.scenes[scId].isUnused:
                template = Template(self.unusedSceneTemplate)

            else:
                template = Template(self.sceneTemplate)

            lines.append(template.safe_substitute(
                self.get_sceneMapping(scId)))

        return lines

    def get_sceneTags(self):
        """Process the scene related tags.
        Return a list of strings.
        """
        lines = []
        headerTemplate = Template(self.scnPerTagtemplate)
        template = Template(self.sceneTemplate)

        for tag in self.xr.scnPerTag:

            if self.xr.scnPerTag[tag] != []:
                lines.append(headerTemplate.safe_substitute(
                    self.get_tagMapping(tag)))
                lines.extend(self.get_scenes(self.xr.scnPerTag[tag]))

        return lines

    def get_characters(self):
        """Process the scenes per character.
        Return a list of strings.
        """
        lines = []
        headerTemplate = Template(self.scnPerChrTemplate)

        for crId in self.xr.scnPerChr:

            if self.xr.scnPerChr[crId] != []:
                lines.append(headerTemplate.safe_substitute(
                    self.get_characterMapping(crId)))
                lines.extend(self.get_scenes(self.xr.scnPerChr[crId]))

        return lines

    def get_locations(self):
        """Process the locations.
        Return a list of strings.
        """
        lines = []
        headerTemplate = Template(self.scnPerLocTemplate)

        for lcId in self.xr.scnPerLoc:

            if self.xr.scnPerLoc[lcId] != []:
                lines.append(headerTemplate.safe_substitute(
                    self.get_locationMapping(lcId)))
                lines.extend(self.get_scenes(self.xr.scnPerLoc[lcId]))

        return lines

    def get_items(self):
        """Process the items.
        Return a list of strings.
        """
        lines = []
        headerTemplate = Template(self.scnPerItmTemplate)

        for itId in self.xr.scnPerItm:

            if self.xr.scnPerItm[itId] != []:
                lines.append(headerTemplate.safe_substitute(
                    self.get_itemMapping(itId)))
                lines.extend(self.get_scenes(self.xr.scnPerItm[itId]))

        return lines

    def get_characterTags(self):
        """Process the character related tags.
        Return a list of strings.
        """
        lines = []
        headerTemplate = Template(self.chrPerTagTemplate)
        template = Template(self.characterTemplate)

        for tag in self.xr.chrPerTag:

            if self.xr.chrPerTag[tag] != []:
                lines.append(headerTemplate.safe_substitute(
                    self.get_tagMapping(tag)))

                for crId in self.xr.chrPerTag[tag]:
                    lines.append(template.safe_substitute(
                        self.get_characterMapping(crId)))

        return lines

    def get_locationTags(self):
        """Process the location related tags.
        Return a list of strings.
        """
        lines = []
        headerTemplate = Template(self.locPerTagTemplate)
        template = Template(self.locationTemplate)

        for tag in self.xr.locPerTag:

            if self.xr.locPerTag[tag]:
                lines.append(headerTemplate.safe_substitute(
                    self.get_tagMapping(tag)))

                for lcId in self.xr.locPerTag[tag]:
                    lines.append(template.safe_substitute(
                        self.get_locationMapping(lcId)))

        return lines

    def get_itemTags(self):
        """Process the item related tags.
        Return a list of strings.
        """
        lines = []
        headerTemplate = Template(self.itmPerTagTemplate)
        template = Template(self.itemTemplate)

        for tag in self.xr.itmPerTag:

            if self.xr.itmPerTag[tag] != []:
                lines.append(headerTemplate.safe_substitute(
                    self.get_tagMapping(tag)))

                for itId in self.xr.itmPerTag[tag]:
                    lines.append(template.safe_substitute(
                        self.get_itemMapping(itId)))

        return lines

    def get_text(self):
        """Apply the template method pattern
        by overwriting a method called during the file export process.
        """
        self.xr.generate_xref(self)

        lines = self.get_fileHeader()
        lines.extend(self.get_characters())
        lines.extend(self.get_locations())
        lines.extend(self.get_items())
        lines.extend(self.get_sceneTags())
        lines.extend(self.get_characterTags())
        lines.extend(self.get_locationTags())
        lines.extend(self.get_itemTags())
        lines.append(self.fileFooter)
        return ''.join(lines)


class OdsFile(OdfFile):
    """ODS project file representation."""

    EXTENSION = '.ods'
    # overwrites Novel.EXTENSION

    ODF_COMPONENTS = ['META-INF', 'content.xml', 'meta.xml', 'mimetype',
                      'settings.xml', 'styles.xml', 'META-INF/manifest.xml']

    # Column width:
    # co1 2.000cm
    # co2 3.000cm
    # co3 4.000cm
    # co4 8.000cm

    CONTENT_XML_HEADER = '''<?xml version="1.0" encoding="UTF-8"?>

<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" xmlns:math="http://www.w3.org/1998/Math/MathML" xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" xmlns:ooo="http://openoffice.org/2004/office" xmlns:ooow="http://openoffice.org/2004/writer" xmlns:oooc="http://openoffice.org/2004/calc" xmlns:dom="http://www.w3.org/2001/xml-events" xmlns:xforms="http://www.w3.org/2002/xforms" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:rpt="http://openoffice.org/2005/report" xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2" xmlns:xhtml="http://www.w3.org/1999/xhtml" xmlns:grddl="http://www.w3.org/2003/g/data-view#" xmlns:tableooo="http://openoffice.org/2009/table" xmlns:field="urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0" office:version="1.2">
 <office:scripts/>
 <office:font-face-decls>
  <style:font-face style:name="Arial" svg:font-family="Arial" style:font-family-generic="swiss" style:font-pitch="variable"/>
  <style:font-face style:name="Segoe UI" svg:font-family="&apos;Segoe UI&apos;" style:font-adornments="Standard" style:font-family-generic="swiss" style:font-pitch="variable"/>
  <style:font-face style:name="Arial Unicode MS" svg:font-family="&apos;Arial Unicode MS&apos;" style:font-family-generic="system" style:font-pitch="variable"/>
  <style:font-face style:name="Microsoft YaHei" svg:font-family="&apos;Microsoft YaHei&apos;" style:font-family-generic="system" style:font-pitch="variable"/>
  <style:font-face style:name="Tahoma" svg:font-family="Tahoma" style:font-family-generic="system" style:font-pitch="variable"/>
 </office:font-face-decls>
 <office:automatic-styles>
  <style:style style:name="co1" style:family="table-column">
   <style:table-column-properties fo:break-before="auto" style:column-width="2.000cm"/>
  </style:style>
  <style:style style:name="co2" style:family="table-column">
   <style:table-column-properties fo:break-before="auto" style:column-width="3.000cm"/>
  </style:style>
  <style:style style:name="co3" style:family="table-column">
   <style:table-column-properties fo:break-before="auto" style:column-width="4.000cm"/>
  </style:style>
  <style:style style:name="co4" style:family="table-column">
   <style:table-column-properties fo:break-before="auto" style:column-width="8.000cm"/>
  </style:style>
  <style:style style:name="ro1" style:family="table-row">
   <style:table-row-properties style:row-height="1.157cm" fo:break-before="auto" style:use-optimal-row-height="true"/>
  </style:style>
  <style:style style:name="ro2" style:family="table-row">
   <style:table-row-properties style:row-height="2.053cm" fo:break-before="auto" style:use-optimal-row-height="true"/>
  </style:style>
  <style:style style:name="ta1" style:family="table" style:master-page-name="Default">
   <style:table-properties table:display="true" style:writing-mode="lr-tb"/>
  </style:style>
 </office:automatic-styles>
 <office:body>
  <office:spreadsheet>
   <table:table table:name="'''

    CONTENT_XML_FOOTER = '''   </table:table>
  </office:spreadsheet>
 </office:body>
</office:document-content>
'''

    _META_XML = '''<?xml version="1.0" encoding="utf-8"?>
<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:ooo="http://openoffice.org/2004/office" xmlns:grddl="http://www.w3.org/2003/g/data-view#" office:version="1.2">
  <office:meta>
    <meta:generator>PyWriter</meta:generator>
    <dc:title>$Title</dc:title>
    <dc:description>$Summary</dc:description>
    <dc:subject></dc:subject>
    <meta:keyword></meta:keyword>
    <meta:initial-creator>$Author</meta:initial-creator>
    <dc:creator></dc:creator>
    <meta:creation-date>${Date}T${Time}Z</meta:creation-date>
    <dc:date></dc:date>
  </office:meta>
</office:document-meta>
'''
    _MANIFEST_XML = '''<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">
 <manifest:file-entry manifest:media-type="application/vnd.oasis.opendocument.spreadsheet" manifest:version="1.2" manifest:full-path="/"/>
 <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="content.xml"/>
 <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="styles.xml"/>
 <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="meta.xml"/>
 <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="settings.xml"/>
</manifest:manifest>    
'''
    _SETTINGS_XML = '''<?xml version="1.0" encoding="UTF-8"?>

<office:document-settings xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0" xmlns:ooo="http://openoffice.org/2004/office" office:version="1.2">
 <office:settings>
  <config:config-item-set config:name="ooo:view-settings">
   <config:config-item config:name="VisibleAreaTop" config:type="int">0</config:config-item>
   <config:config-item config:name="VisibleAreaLeft" config:type="int">0</config:config-item>
   <config:config-item config:name="VisibleAreaWidth" config:type="int">44972</config:config-item>
   <config:config-item config:name="VisibleAreaHeight" config:type="int">18999</config:config-item>
   <config:config-item-map-indexed config:name="Views">
    <config:config-item-map-entry>
     <config:config-item config:name="ViewId" config:type="string">view1</config:config-item>
     <config:config-item-map-named config:name="Tables">
      <config:config-item-map-entry config:name="Tabelle1">
       <config:config-item config:name="CursorPositionX" config:type="int">5</config:config-item>
       <config:config-item config:name="CursorPositionY" config:type="int">1</config:config-item>
       <config:config-item config:name="HorizontalSplitMode" config:type="short">0</config:config-item>
       <config:config-item config:name="VerticalSplitMode" config:type="short">0</config:config-item>
       <config:config-item config:name="HorizontalSplitPosition" config:type="int">0</config:config-item>
       <config:config-item config:name="VerticalSplitPosition" config:type="int">0</config:config-item>
       <config:config-item config:name="ActiveSplitRange" config:type="short">2</config:config-item>
       <config:config-item config:name="PositionLeft" config:type="int">0</config:config-item>
       <config:config-item config:name="PositionRight" config:type="int">0</config:config-item>
       <config:config-item config:name="PositionTop" config:type="int">0</config:config-item>
       <config:config-item config:name="PositionBottom" config:type="int">0</config:config-item>
       <config:config-item config:name="ZoomType" config:type="short">0</config:config-item>
       <config:config-item config:name="ZoomValue" config:type="int">100</config:config-item>
       <config:config-item config:name="PageViewZoomValue" config:type="int">60</config:config-item>
      </config:config-item-map-entry>
     </config:config-item-map-named>
     <config:config-item config:name="ActiveTable" config:type="string">Tabelle1</config:config-item>
     <config:config-item config:name="HorizontalScrollbarWidth" config:type="int">270</config:config-item>
     <config:config-item config:name="ZoomType" config:type="short">0</config:config-item>
     <config:config-item config:name="ZoomValue" config:type="int">100</config:config-item>
     <config:config-item config:name="PageViewZoomValue" config:type="int">60</config:config-item>
     <config:config-item config:name="ShowPageBreakPreview" config:type="boolean">false</config:config-item>
     <config:config-item config:name="ShowZeroValues" config:type="boolean">true</config:config-item>
     <config:config-item config:name="ShowNotes" config:type="boolean">true</config:config-item>
     <config:config-item config:name="ShowGrid" config:type="boolean">true</config:config-item>
     <config:config-item config:name="GridColor" config:type="long">12632256</config:config-item>
     <config:config-item config:name="ShowPageBreaks" config:type="boolean">true</config:config-item>
     <config:config-item config:name="HasColumnRowHeaders" config:type="boolean">true</config:config-item>
     <config:config-item config:name="HasSheetTabs" config:type="boolean">true</config:config-item>
     <config:config-item config:name="IsOutlineSymbolsSet" config:type="boolean">true</config:config-item>
     <config:config-item config:name="IsSnapToRaster" config:type="boolean">false</config:config-item>
     <config:config-item config:name="RasterIsVisible" config:type="boolean">false</config:config-item>
     <config:config-item config:name="RasterResolutionX" config:type="int">1000</config:config-item>
     <config:config-item config:name="RasterResolutionY" config:type="int">1000</config:config-item>
     <config:config-item config:name="RasterSubdivisionX" config:type="int">1</config:config-item>
     <config:config-item config:name="RasterSubdivisionY" config:type="int">1</config:config-item>
     <config:config-item config:name="IsRasterAxisSynchronized" config:type="boolean">true</config:config-item>
    </config:config-item-map-entry>
   </config:config-item-map-indexed>
  </config:config-item-set>
  <config:config-item-set config:name="ooo:configuration-settings">
   <config:config-item config:name="IsKernAsianPunctuation" config:type="boolean">false</config:config-item>
   <config:config-item config:name="IsRasterAxisSynchronized" config:type="boolean">true</config:config-item>
   <config:config-item config:name="LinkUpdateMode" config:type="short">3</config:config-item>
   <config:config-item config:name="SaveVersionOnClose" config:type="boolean">false</config:config-item>
   <config:config-item config:name="AllowPrintJobCancel" config:type="boolean">true</config:config-item>
   <config:config-item config:name="HasSheetTabs" config:type="boolean">true</config:config-item>
   <config:config-item config:name="ShowPageBreaks" config:type="boolean">true</config:config-item>
   <config:config-item config:name="RasterResolutionX" config:type="int">1000</config:config-item>
   <config:config-item config:name="PrinterSetup" config:type="base64Binary"/>
   <config:config-item config:name="RasterResolutionY" config:type="int">1000</config:config-item>
   <config:config-item config:name="LoadReadonly" config:type="boolean">false</config:config-item>
   <config:config-item config:name="RasterSubdivisionX" config:type="int">1</config:config-item>
   <config:config-item config:name="ShowNotes" config:type="boolean">true</config:config-item>
   <config:config-item config:name="ShowZeroValues" config:type="boolean">true</config:config-item>
   <config:config-item config:name="RasterSubdivisionY" config:type="int">1</config:config-item>
   <config:config-item config:name="ApplyUserData" config:type="boolean">true</config:config-item>
   <config:config-item config:name="GridColor" config:type="long">12632256</config:config-item>
   <config:config-item config:name="RasterIsVisible" config:type="boolean">false</config:config-item>
   <config:config-item config:name="IsSnapToRaster" config:type="boolean">false</config:config-item>
   <config:config-item config:name="PrinterName" config:type="string"/>
   <config:config-item config:name="ShowGrid" config:type="boolean">true</config:config-item>
   <config:config-item config:name="CharacterCompressionType" config:type="short">0</config:config-item>
   <config:config-item-map-indexed config:name="ForbiddenCharacters">
    <config:config-item-map-entry>
     <config:config-item config:name="Language" config:type="string">$Language</config:config-item>
     <config:config-item config:name="Country" config:type="string">$Country</config:config-item>
     <config:config-item config:name="Variant" config:type="string"/>
     <config:config-item config:name="BeginLine" config:type="string"/>
     <config:config-item config:name="EndLine" config:type="string"/>
    </config:config-item-map-entry>
   </config:config-item-map-indexed>
   <config:config-item config:name="IsOutlineSymbolsSet" config:type="boolean">true</config:config-item>
   <config:config-item config:name="AutoCalculate" config:type="boolean">true</config:config-item>
   <config:config-item config:name="IsDocumentShared" config:type="boolean">false</config:config-item>
   <config:config-item config:name="UpdateFromTemplate" config:type="boolean">true</config:config-item>
  </config:config-item-set>
 </office:settings>
</office:document-settings>
'''
    _STYLES_XML = '''<?xml version="1.0" encoding="UTF-8"?>

<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" xmlns:math="http://www.w3.org/1998/Math/MathML" xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" xmlns:ooo="http://openoffice.org/2004/office" xmlns:ooow="http://openoffice.org/2004/writer" xmlns:oooc="http://openoffice.org/2004/calc" xmlns:dom="http://www.w3.org/2001/xml-events" xmlns:rpt="http://openoffice.org/2005/report" xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2" xmlns:xhtml="http://www.w3.org/1999/xhtml" xmlns:grddl="http://www.w3.org/2003/g/data-view#" xmlns:tableooo="http://openoffice.org/2009/table" office:version="1.2">
 <office:font-face-decls>
  <style:font-face style:name="Arial" svg:font-family="Arial" style:font-family-generic="swiss" style:font-pitch="variable"/>
  <style:font-face style:name="Segoe UI" svg:font-family="&apos;Segoe UI&apos;" style:font-adornments="Standard" style:font-family-generic="swiss" style:font-pitch="variable"/>
  <style:font-face style:name="Arial Unicode MS" svg:font-family="&apos;Arial Unicode MS&apos;" style:font-family-generic="system" style:font-pitch="variable"/>
  <style:font-face style:name="Microsoft YaHei" svg:font-family="&apos;Microsoft YaHei&apos;" style:font-family-generic="system" style:font-pitch="variable"/>
  <style:font-face style:name="Tahoma" svg:font-family="Tahoma" style:font-family-generic="system" style:font-pitch="variable"/>
 </office:font-face-decls>
 <office:styles>
  <style:default-style style:family="table-cell">
   <style:paragraph-properties style:tab-stop-distance="1.25cm"/>
   <style:text-properties style:font-name="Arial" fo:language="de" fo:country="DE" style:font-name-asian="Arial Unicode MS" style:language-asian="zh" style:country-asian="CN" style:font-name-complex="Tahoma" style:language-complex="hi" style:country-complex="IN"/>
  </style:default-style>
  <number:number-style style:name="N0">
   <number:number number:min-integer-digits="1"/>
  </number:number-style>
  <style:style style:name="Default" style:family="table-cell">
   <style:table-cell-properties style:text-align-source="fix" style:repeat-content="false" fo:background-color="transparent" fo:wrap-option="wrap" fo:padding="0.136cm" style:vertical-align="top"/>
   <style:paragraph-properties fo:text-align="start"/>
   <style:text-properties style:font-name="Segoe UI" style:font-name-asian="Microsoft YaHei" style:font-name-complex="Arial Unicode MS"/>
  </style:style>
  <style:style style:name="Result" style:family="table-cell" style:parent-style-name="Default">
   <style:text-properties fo:font-style="italic" style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color" fo:font-weight="bold"/>
  </style:style>
  <style:style style:name="Result2" style:family="table-cell" style:parent-style-name="Result"/>
  <style:style style:name="Heading" style:family="table-cell" style:parent-style-name="Default">
   <style:table-cell-properties fo:background-color="#cfe7f5" style:text-align-source="fix" style:repeat-content="false"/>
   <style:paragraph-properties fo:text-align="start"/>
   <style:text-properties fo:font-weight="bold"/>
  </style:style>
  <style:style style:name="Heading1" style:family="table-cell" style:parent-style-name="Heading">
   <style:table-cell-properties style:rotation-angle="90"/>
  </style:style>
 </office:styles>
 <office:automatic-styles>
  <style:page-layout style:name="Mpm1">
   <style:page-layout-properties style:writing-mode="lr-tb"/>
   <style:header-style>
    <style:header-footer-properties fo:min-height="0.751cm" fo:margin-left="0cm" fo:margin-right="0cm" fo:margin-bottom="0.25cm"/>
   </style:header-style>
   <style:footer-style>
    <style:header-footer-properties fo:min-height="0.751cm" fo:margin-left="0cm" fo:margin-right="0cm" fo:margin-top="0.25cm"/>
   </style:footer-style>
  </style:page-layout>
  <style:page-layout style:name="Mpm2">
   <style:page-layout-properties style:writing-mode="lr-tb"/>
   <style:header-style>
    <style:header-footer-properties fo:min-height="0.751cm" fo:margin-left="0cm" fo:margin-right="0cm" fo:margin-bottom="0.25cm" fo:border="0.088cm solid #000000" fo:padding="0.018cm" fo:background-color="#c0c0c0">
     <style:background-image/>
    </style:header-footer-properties>
   </style:header-style>
   <style:footer-style>
    <style:header-footer-properties fo:min-height="0.751cm" fo:margin-left="0cm" fo:margin-right="0cm" fo:margin-top="0.25cm" fo:border="0.088cm solid #000000" fo:padding="0.018cm" fo:background-color="#c0c0c0">
     <style:background-image/>
    </style:header-footer-properties>
   </style:footer-style>
  </style:page-layout>
 </office:automatic-styles>
 <office:master-styles>
  <style:master-page style:name="Default" style:page-layout-name="Mpm1">
   <style:header>
    <text:p><text:sheet-name>???</text:sheet-name></text:p>
   </style:header>
   <style:header-left style:display="false"/>
   <style:footer>
    <text:p>Seite <text:page-number>1</text:page-number></text:p>
   </style:footer>
   <style:footer-left style:display="false"/>
  </style:master-page>
  <style:master-page style:name="Report" style:page-layout-name="Mpm2">
   <style:header>
    <style:region-left>
     <text:p><text:sheet-name>???</text:sheet-name> (<text:title>???</text:title>)</text:p>
    </style:region-left>
    <style:region-right>
     <text:p><text:date style:data-style-name="N2" text:date-value="2021-03-15">15.03.2021</text:date>, <text:time>15:34:40</text:time></text:p>
    </style:region-right>
   </style:header>
   <style:header-left style:display="false"/>
   <style:footer>
    <text:p>Seite <text:page-number>1</text:page-number> / <text:page-count>99</text:page-count></text:p>
   </style:footer>
   <style:footer-left style:display="false"/>
  </style:master-page>
 </office:master-styles>
</office:document-styles>
'''
    _MIMETYPE = 'application/vnd.oasis.opendocument.spreadsheet'

    def convert_from_yw(self, text):
        """Convert yw7 raw markup to ods. Return an xml string.
        """

        ODS_REPLACEMENTS = [
            ['&', '&amp;'],  # must be first!
            ['"', '&quot;'],
            ["'", '&apos;'],
            ['>', '&gt;'],
            ['<', '&lt;'],
            ['\n', '</text:p>\n<text:p>'],
        ]

        try:
            text = text.rstrip()

            for r in ODS_REPLACEMENTS:
                text = text.replace(r[0], r[1])

        except AttributeError:
            text = ''

        return text


class OdsSceneList(OdsFile):
    """ODS scene list representation."""

    DESCRIPTION = 'Scene list'
    SUFFIX = '_scenelist'

    _SCENE_RATINGS = ['2', '3', '4', '5', '6', '7', '8', '9', '10']
    # '1' is assigned N/A (empty table cell).

    # Column width:
    # co1 2.000cm
    # co2 3.000cm
    # co3 4.000cm
    # co4 8.000cm

    # Header structure:
    # Scene link
    # Scene title
    # Scene description
    # Tags
    # Scene notes
    # A/R
    # Goal
    # Conflict
    # Outcome
    # Scene
    # Words total
    # $FieldTitle1
    # $FieldTitle2
    # $FieldTitle3
    # $FieldTitle4
    # Word count
    # Letter count
    # Status
    # Characters
    # Locations
    # Items

    fileHeader = OdsFile.CONTENT_XML_HEADER + DESCRIPTION + '''" table:style-name="ta1" table:print="false">
    <table:table-column table:style-name="co1" table:default-cell-style-name="Default"/>
    <table:table-column table:style-name="co3" table:default-cell-style-name="Default"/>
    <table:table-column table:style-name="co4" table:default-cell-style-name="Default"/>
    <table:table-column table:style-name="co3" table:default-cell-style-name="Default"/>
    <table:table-column table:style-name="co4" table:default-cell-style-name="Default"/>
    <table:table-column table:style-name="co1" table:default-cell-style-name="Default"/>
    <table:table-column table:style-name="co4" table:default-cell-style-name="Default"/>
    <table:table-column table:style-name="co4" table:default-cell-style-name="Default"/>
    <table:table-column table:style-name="co4" table:default-cell-style-name="Default"/>
    <table:table-column table:style-name="co1" table:default-cell-style-name="Default"/>
    <table:table-column table:style-name="co1" table:default-cell-style-name="Default"/>
    <table:table-column table:style-name="co1" table:default-cell-style-name="Default"/>
    <table:table-column table:style-name="co1" table:default-cell-style-name="Default"/>
    <table:table-column table:style-name="co1" table:default-cell-style-name="Default"/>
    <table:table-column table:style-name="co1" table:default-cell-style-name="Default"/>
    <table:table-column table:style-name="co1" table:default-cell-style-name="Default"/>
    <table:table-column table:style-name="co1" table:default-cell-style-name="Default"/>
    <table:table-column table:style-name="co2" table:default-cell-style-name="Default"/>
    <table:table-column table:style-name="co2" table:default-cell-style-name="Default"/>
    <table:table-column table:style-name="co3" table:default-cell-style-name="Default"/>
    <table:table-column table:style-name="co3" table:default-cell-style-name="Default"/>
    <table:table-column table:style-name="co3" table:default-cell-style-name="Default"/>
    <table:table-row table:style-name="ro1">
     <table:table-cell table:style-name="Heading" office:value-type="string">
      <text:p>Scene link</text:p>
     </table:table-cell>
     <table:table-cell table:style-name="Heading" office:value-type="string">
      <text:p>Scene title</text:p>
     </table:table-cell>
     <table:table-cell table:style-name="Heading" office:value-type="string">
      <text:p>Scene description</text:p>
     </table:table-cell>
     <table:table-cell table:style-name="Heading" office:value-type="string">
      <text:p>Tags</text:p>
     </table:table-cell>
     <table:table-cell table:style-name="Heading" office:value-type="string">
      <text:p>Scene notes</text:p>
     </table:table-cell>
     <table:table-cell table:style-name="Heading" office:value-type="string">
      <text:p>A/R</text:p>
     </table:table-cell>
     <table:table-cell table:style-name="Heading" office:value-type="string">
      <text:p>Goal</text:p>
     </table:table-cell>
     <table:table-cell table:style-name="Heading" office:value-type="string">
      <text:p>Conflict</text:p>
     </table:table-cell>
     <table:table-cell table:style-name="Heading" office:value-type="string">
      <text:p>Outcome</text:p>
     </table:table-cell>
     <table:table-cell table:style-name="Heading" office:value-type="string">
      <text:p>Scene</text:p>
     </table:table-cell>
     <table:table-cell table:style-name="Heading" office:value-type="string">
      <text:p>Words total</text:p>
     </table:table-cell>
     <table:table-cell table:style-name="Heading" office:value-type="string">
      <text:p>$FieldTitle1</text:p>
     </table:table-cell>
     <table:table-cell table:style-name="Heading" office:value-type="string">
      <text:p>$FieldTitle2</text:p>
     </table:table-cell>
     <table:table-cell table:style-name="Heading" office:value-type="string">
      <text:p>$FieldTitle3</text:p>
     </table:table-cell>
     <table:table-cell table:style-name="Heading" office:value-type="string">
      <text:p>$FieldTitle4</text:p>
     </table:table-cell>
     <table:table-cell table:style-name="Heading" office:value-type="string">
      <text:p>Word count</text:p>
     </table:table-cell>
     <table:table-cell table:style-name="Heading" office:value-type="string">
      <text:p>Letter count</text:p>
     </table:table-cell>
     <table:table-cell table:style-name="Heading" office:value-type="string">
      <text:p>Status</text:p>
     </table:table-cell>
     <table:table-cell table:style-name="Heading" office:value-type="string">
      <text:p>Characters</text:p>
     </table:table-cell>
     <table:table-cell table:style-name="Heading" office:value-type="string">
      <text:p>Locations</text:p>
     </table:table-cell>
     <table:table-cell table:style-name="Heading" office:value-type="string">
      <text:p>Items</text:p>
     </table:table-cell>
     <table:table-cell table:style-name="Heading" table:number-columns-repeated="1003"/>
    </table:table-row>

'''

    sceneTemplate = '''   <table:table-row table:style-name="ro2">
     <table:table-cell table:formula="of:=HYPERLINK(&quot;file:///$ProjectPath/${ProjectName}_manuscript.odt#ScID:$ID%7Cregion&quot;;&quot;ScID:$ID&quot;)" office:value-type="string" office:string-value="ScID:$ID">
      <text:p>ScID:$ID</text:p>
     </table:table-cell>
     <table:table-cell office:value-type="string">
      <text:p>$Title</text:p>
     </table:table-cell>
     <table:table-cell office:value-type="string">
      <text:p>$Desc</text:p>
     </table:table-cell>
     <table:table-cell office:value-type="string">
      <text:p>$Tags</text:p>
     </table:table-cell>
     <table:table-cell office:value-type="string">
      <text:p>$Notes</text:p>
     </table:table-cell>
     <table:table-cell office:value-type="string">
      <text:p>$ReactionScene</text:p>
     </table:table-cell>
     <table:table-cell office:value-type="string">
      <text:p>$Goal</text:p>
     </table:table-cell>
     <table:table-cell office:value-type="string">
      <text:p>$Conflict</text:p>
     </table:table-cell>
     <table:table-cell office:value-type="string">
      <text:p>$Outcome</text:p>
     </table:table-cell>
     <table:table-cell office:value-type="float" office:value="$SceneNumber">
      <text:p>$SceneNumber</text:p>
     </table:table-cell>
     <table:table-cell office:value-type="float" office:value="$WordsTotal">
      <text:p>$WordsTotal</text:p>
     </table:table-cell>
     <table:table-cell office:value-type="float" office:value="$Field1">
      <text:p>$Field1</text:p>
     </table:table-cell>
     <table:table-cell office:value-type="float" office:value="$Field2">
      <text:p>$Field2</text:p>
     </table:table-cell>
     <table:table-cell office:value-type="float" office:value="$Field3">
      <text:p>$Field3</text:p>
     </table:table-cell>
     <table:table-cell office:value-type="float" office:value="$Field4">
      <text:p>$Field4</text:p>
     </table:table-cell>
     <table:table-cell office:value-type="float" office:value="$WordCount">
      <text:p>$WordCount</text:p>
     </table:table-cell>
     <table:table-cell office:value-type="float" office:value="$LetterCount">
      <text:p>$LetterCount</text:p>
     </table:table-cell>
     <table:table-cell office:value-type="string">
      <text:p>$Status</text:p>
     </table:table-cell>
     <table:table-cell office:value-type="string">
      <text:p>$Characters</text:p>
     </table:table-cell>
     <table:table-cell office:value-type="string">
      <text:p>$Locations</text:p>
     </table:table-cell>
     <table:table-cell>
      <text:p>$Items</text:p>
     </table:table-cell>
    </table:table-row>

'''

    fileFooter = OdsFile.CONTENT_XML_FOOTER

    def get_sceneMapping(self, scId, sceneNumber, wordsTotal, lettersTotal):
        """Return a mapping dictionary for a scene section. 
        """
        sceneMapping = OdsFile.get_sceneMapping(
            self, scId, sceneNumber, wordsTotal, lettersTotal)

        if self.scenes[scId].field1 == '1':
            sceneMapping['Field1'] = ''

        if self.scenes[scId].field2 == '1':
            sceneMapping['Field2'] = ''

        if self.scenes[scId].field3 == '1':
            sceneMapping['Field3'] = ''

        if self.scenes[scId].field4 == '1':
            sceneMapping['Field4'] = ''

        return sceneMapping


class OdsPlotList(OdsFile):
    """ODS plot list representation with plot related metadata."""

    DESCRIPTION = 'Plot list'
    SUFFIX = '_plotlist'

    _STORYLINE_MARKER = 'story'
    # Field names containing this string (case insensitive)
    # are associated to storylines

    _SCENE_RATINGS = ['2', '3', '4', '5', '6', '7', '8', '9', '10']
    # '1' is assigned N/A (empty table cell).

    _NOT_APPLICABLE = 'N/A'
    # Scene field column header for fields not being assigned to a storyline

    _CHAR_STATE = ['', 'N/A', 'unhappy', 'dissatisfied',
                   'vague', 'satisfied', 'happy', '', '', '', '']

    # Column width:
    # co1 2.000cm
    # co2 3.000cm
    # co3 4.000cm
    # co4 8.000cm

    # Header structure:
    # ID
    # Plot section
    # Plot event
    # Plot event title
    # Details
    # Scene
    # Words total
    # $FieldTitle1
    # $FieldTitle2
    # $FieldTitle3
    # $FieldTitle4

    fileHeader = OdsFile.CONTENT_XML_HEADER + DESCRIPTION + '''" table:style-name="ta1" table:print="false">
    <table:table-column table:style-name="co1" table:default-cell-style-name="Default"/>
    <table:table-column table:style-name="co3" table:default-cell-style-name="Default"/>
    <table:table-column table:style-name="co3" table:default-cell-style-name="Default"/>
    <table:table-column table:style-name="co3" table:default-cell-style-name="Default"/>
    <table:table-column table:style-name="co4" table:default-cell-style-name="Default"/>
    <table:table-column table:style-name="co1" table:default-cell-style-name="Default"/>
    <table:table-column table:style-name="co1" table:default-cell-style-name="Default"/>
    <table:table-column table:style-name="co1" table:default-cell-style-name="Default"/>
    <table:table-column table:style-name="co1" table:default-cell-style-name="Default"/>
    <table:table-column table:style-name="co1" table:default-cell-style-name="Default"/>
    <table:table-column table:style-name="co1" table:default-cell-style-name="Default"/>
    <table:table-row table:style-name="ro1">
     <table:table-cell table:style-name="Heading" office:value-type="string">
      <text:p>ID</text:p>
     </table:table-cell>
     <table:table-cell table:style-name="Heading" office:value-type="string">
      <text:p>Plot section</text:p>
     </table:table-cell>
     <table:table-cell table:style-name="Heading" office:value-type="string">
      <text:p>Plot event</text:p>
     </table:table-cell>
     <table:table-cell table:style-name="Heading" office:value-type="string">
      <text:p>Plot event title</text:p>
     </table:table-cell>
     <table:table-cell table:style-name="Heading" office:value-type="string">
      <text:p>Details</text:p>
     </table:table-cell>
     <table:table-cell table:style-name="Heading" office:value-type="string">
      <text:p>Scene</text:p>
     </table:table-cell>
     <table:table-cell table:style-name="Heading" office:value-type="string">
      <text:p>Words total</text:p>
     </table:table-cell>
     <table:table-cell table:style-name="Heading" office:value-type="string">
      <text:p>$FieldTitle1</text:p>
     </table:table-cell>
     <table:table-cell table:style-name="Heading" office:value-type="string">
      <text:p>$FieldTitle2</text:p>
     </table:table-cell>
     <table:table-cell table:style-name="Heading" office:value-type="string">
      <text:p>$FieldTitle3</text:p>
     </table:table-cell>
     <table:table-cell table:style-name="Heading" office:value-type="string">
      <text:p>$FieldTitle4</text:p>
     </table:table-cell>
     <table:table-cell table:style-name="Heading" table:number-columns-repeated="1003"/>
    </table:table-row>

'''

    notesChapterTemplate = '''   <table:table-row table:style-name="ro2">
     <table:table-cell office:value-type="string">
      <text:p>ChID:$ID</text:p>
     </table:table-cell>
     <table:table-cell office:value-type="string">
      <text:p>$Title</text:p>
     </table:table-cell>
     <table:table-cell office:value-type="string">
     </table:table-cell>
     <table:table-cell office:value-type="string">
     </table:table-cell>
     <table:table-cell office:value-type="string">
      <text:p>$Desc</text:p>
     </table:table-cell>
     <table:table-cell office:value-type="string">
     </table:table-cell>
     <table:table-cell office:value-type="string">
     </table:table-cell>
     <table:table-cell office:value-type="string">
     </table:table-cell>
     <table:table-cell office:value-type="string">
     </table:table-cell>
     <table:table-cell office:value-type="string">
     </table:table-cell>
     <table:table-cell office:value-type="string">
     </table:table-cell>
    </table:table-row>

'''
    sceneTemplate = '''   <table:table-row table:style-name="ro2">
     <table:table-cell table:formula="of:=HYPERLINK(&quot;file:///$ProjectPath/${ProjectName}_manuscript.odt#ScID:$ID%7Cregion&quot;;&quot;ScID:$ID&quot;)" office:value-type="string" office:string-value="ScID:$ID">
      <text:p>ScID:$ID</text:p>
     </table:table-cell>
     <table:table-cell office:value-type="string">
     </table:table-cell>
     <table:table-cell office:value-type="string">
      <text:p>$Tags</text:p>
     </table:table-cell>
     <table:table-cell office:value-type="string">
      <text:p>$Title</text:p>
     </table:table-cell>
     <table:table-cell office:value-type="string">
      <text:p>$Notes</text:p>
     </table:table-cell>
     <table:table-cell office:value-type="string">
      <text:p>$SceneNumber</text:p>
     </table:table-cell>
     <table:table-cell office:value-type="float" office:value="$WordsTotal">
      <text:p>$WordsTotal</text:p>
     </table:table-cell>
     <table:table-cell office:value-type=$Field1
     </table:table-cell>
     <table:table-cell office:value-type=$Field2
     </table:table-cell>
     <table:table-cell office:value-type=$Field3
     </table:table-cell>
     <table:table-cell office:value-type=$Field4
     </table:table-cell>
    </table:table-row>

'''

    fileFooter = OdsFile.CONTENT_XML_FOOTER

    def get_fileHeaderMapping(self):
        """Return a mapping dictionary for the project section. 
        """
        projectTemplateMapping = OdsFile.get_fileHeaderMapping(self)

        charList = []

        for crId in self.srtCharacters:
            charList.append(self.characters[crId].title)
            # Collect character names to identify storylines

        if self.fieldTitle1 in charList or self._STORYLINE_MARKER in self.fieldTitle1.lower():
            self.arc1 = True

        else:
            self.arc1 = False
            projectTemplateMapping['FieldTitle1'] = self._NOT_APPLICABLE

        if self.fieldTitle2 in charList or self._STORYLINE_MARKER in self.fieldTitle2.lower():
            self.arc2 = True

        else:
            self.arc2 = False
            projectTemplateMapping['FieldTitle2'] = self._NOT_APPLICABLE

        if self.fieldTitle3 in charList or self._STORYLINE_MARKER in self.fieldTitle3.lower():
            self.arc3 = True

        else:
            self.arc3 = False
            projectTemplateMapping['FieldTitle3'] = self._NOT_APPLICABLE

        if self.fieldTitle4 in charList or self._STORYLINE_MARKER in self.fieldTitle4.lower():
            self.arc4 = True

        else:
            self.arc4 = False
            projectTemplateMapping['FieldTitle4'] = self._NOT_APPLICABLE

        return projectTemplateMapping

    def get_sceneMapping(self, scId, sceneNumber, wordsTotal, lettersTotal):
        """Return a mapping dictionary for a scene section. 
        """
        sceneMapping = OdsFile.get_sceneMapping(
            self, scId, sceneNumber, wordsTotal, lettersTotal)

        # Suppress display if the field doesn't represent a storyline,
        # or if the field's value equals 1

        if self.scenes[scId].field1 == '1' or not self.arc1:
            sceneMapping['Field1'] = '"string">\n'

        else:
            sceneMapping['Field1'] = '"float" office:value="' + sceneMapping['Field1'] + \
                '">\n      <text:p>' + sceneMapping['Field1'] + '</text:p>'

        if self.scenes[scId].field2 == '1' or not self.arc2:
            sceneMapping['Field2'] = '"string">\n'

        else:
            sceneMapping['Field2'] = '"float" office:value="' + sceneMapping['Field2'] + \
                '">\n      <text:p>' + sceneMapping['Field2'] + '</text:p>'

        if self.scenes[scId].field3 == '1' or not self.arc3:
            sceneMapping['Field3'] = '"string">\n'

        else:
            sceneMapping['Field3'] = '"float" office:value="' + sceneMapping['Field3'] + \
                '">\n      <text:p>' + sceneMapping['Field3'] + '</text:p>'

        if self.scenes[scId].field4 == '1' or not self.arc4:
            sceneMapping['Field4'] = '"string">\n'

        else:
            sceneMapping['Field4'] = '"float" office:value="' + sceneMapping['Field4'] + \
                '">\n      <text:p>' + sceneMapping['Field4'] + '</text:p>'

        return sceneMapping


class OdsCharList(OdsFile):
    """ODS character list representation."""

    DESCRIPTION = 'Character list'
    SUFFIX = '_charlist'

    fileHeader = OdsFile.CONTENT_XML_HEADER + DESCRIPTION + '''" table:style-name="ta1" table:print="false">
    <table:table-column table:style-name="co1" table:default-cell-style-name="Default"/>
    <table:table-column table:style-name="co2" table:default-cell-style-name="Default"/>
    <table:table-column table:style-name="co3" table:default-cell-style-name="Default"/>
    <table:table-column table:style-name="co2" table:default-cell-style-name="Default"/>
    <table:table-column table:style-name="co4" table:number-columns-repeated="3" table:default-cell-style-name="Default"/>
    <table:table-column table:style-name="co2" table:default-cell-style-name="Default"/>
    <table:table-column table:style-name="co3" table:default-cell-style-name="Default"/>
    <table:table-column table:style-name="co4" table:default-cell-style-name="Default"/>
    <table:table-column table:style-name="co1" table:number-columns-repeated="1014" table:default-cell-style-name="Default"/>
     <table:table-row table:style-name="ro1">
     <table:table-cell table:style-name="Heading" office:value-type="string">
      <text:p>ID</text:p>
     </table:table-cell>
     <table:table-cell table:style-name="Heading" office:value-type="string">
      <text:p>Name</text:p>
     </table:table-cell>
     <table:table-cell table:style-name="Heading" office:value-type="string">
      <text:p>Full name</text:p>
     </table:table-cell>
     <table:table-cell table:style-name="Heading" office:value-type="string">
      <text:p>Aka</text:p>
     </table:table-cell>
     <table:table-cell table:style-name="Heading" office:value-type="string">
      <text:p>Description</text:p>
     </table:table-cell>
     <table:table-cell table:style-name="Heading" office:value-type="string">
      <text:p>Bio</text:p>
     </table:table-cell>
     <table:table-cell table:style-name="Heading" office:value-type="string">
      <text:p>Goals</text:p>
     </table:table-cell>
     <table:table-cell table:style-name="Heading" office:value-type="string">
      <text:p>Importance</text:p>
     </table:table-cell>
     <table:table-cell table:style-name="Heading" office:value-type="string">
      <text:p>Tags</text:p>
     </table:table-cell>
     <table:table-cell table:style-name="Heading" office:value-type="string">
      <text:p>Notes</text:p>
     </table:table-cell>
     <table:table-cell table:style-name="Heading" table:number-columns-repeated="1014"/>
    </table:table-row>

'''

    characterTemplate = '''   <table:table-row table:style-name="ro2">
     <table:table-cell office:value-type="string">
      <text:p>CrID:$ID</text:p>
     </table:table-cell>
     <table:table-cell office:value-type="string">
      <text:p>$Title</text:p>
     </table:table-cell>
     <table:table-cell office:value-type="string">
      <text:p>$FullName</text:p>
     </table:table-cell>
     <table:table-cell office:value-type="string">
      <text:p>$AKA</text:p>
     </table:table-cell>
     <table:table-cell office:value-type="string">
      <text:p>$Desc</text:p>
     </table:table-cell>
     <table:table-cell office:value-type="string">
      <text:p>$Bio</text:p>
     </table:table-cell>
     <table:table-cell office:value-type="string">
      <text:p>$Goals</text:p>
     </table:table-cell>
     <table:table-cell office:value-type="string">
      <text:p>$Status</text:p>
     </table:table-cell>
     <table:table-cell office:value-type="string">
      <text:p>$Tags</text:p>
     </table:table-cell>
     <table:table-cell office:value-type="string">
      <text:p>$Notes</text:p>
     </table:table-cell>
     <table:table-cell table:number-columns-repeated="1014"/>
    </table:table-row>

'''

    fileFooter = OdsFile.CONTENT_XML_FOOTER


class OdsLocList(OdsFile):
    """ODS location list representation."""

    DESCRIPTION = 'Location list'
    SUFFIX = '_loclist'

    fileHeader = OdsFile.CONTENT_XML_HEADER + DESCRIPTION + '''" table:style-name="ta1" table:print="false">
    <table:table-column table:style-name="co1" table:default-cell-style-name="Default"/>
    <table:table-column table:style-name="co3" table:default-cell-style-name="Default"/>
    <table:table-column table:style-name="co4" table:default-cell-style-name="Default"/>
    <table:table-column table:style-name="co2" table:default-cell-style-name="Default"/>
    <table:table-column table:style-name="co3" table:default-cell-style-name="Default"/>
    <table:table-column table:style-name="co1" table:number-columns-repeated="1014" table:default-cell-style-name="Default"/>
     <table:table-row table:style-name="ro1">
     <table:table-cell table:style-name="Heading" office:value-type="string">
      <text:p>ID</text:p>
     </table:table-cell>
     <table:table-cell table:style-name="Heading" office:value-type="string">
      <text:p>Name</text:p>
     </table:table-cell>
     <table:table-cell table:style-name="Heading" office:value-type="string">
      <text:p>Description</text:p>
     </table:table-cell>
     <table:table-cell table:style-name="Heading" office:value-type="string">
      <text:p>Aka</text:p>
     </table:table-cell>
     <table:table-cell table:style-name="Heading" office:value-type="string">
      <text:p>Tags</text:p>
     </table:table-cell>
     <table:table-cell table:style-name="Heading" table:number-columns-repeated="1014"/>
    </table:table-row>

'''

    locationTemplate = '''   <table:table-row table:style-name="ro2">
     <table:table-cell office:value-type="string">
      <text:p>LcID:$ID</text:p>
     </table:table-cell>
     <table:table-cell office:value-type="string">
      <text:p>$Title</text:p>
     </table:table-cell>
     <table:table-cell office:value-type="string">
      <text:p>$Desc</text:p>
     </table:table-cell>
     <table:table-cell office:value-type="string">
      <text:p>$AKA</text:p>
     </table:table-cell>
     <table:table-cell office:value-type="string">
      <text:p>$Tags</text:p>
     </table:table-cell>
     <table:table-cell table:number-columns-repeated="1014"/>
    </table:table-row>

'''
    fileFooter = OdsFile.CONTENT_XML_FOOTER


class OdsItemList(OdsFile):
    """ODS item list representation."""

    DESCRIPTION = 'Item list'
    SUFFIX = '_itemlist'

    fileHeader = OdsFile.CONTENT_XML_HEADER + DESCRIPTION + '''" table:style-name="ta1" table:print="false">
    <table:table-column table:style-name="co1" table:default-cell-style-name="Default"/>
    <table:table-column table:style-name="co3" table:default-cell-style-name="Default"/>
    <table:table-column table:style-name="co4" table:default-cell-style-name="Default"/>
    <table:table-column table:style-name="co2" table:default-cell-style-name="Default"/>
    <table:table-column table:style-name="co3" table:default-cell-style-name="Default"/>
    <table:table-column table:style-name="co1" table:number-columns-repeated="1014" table:default-cell-style-name="Default"/>
     <table:table-row table:style-name="ro1">
     <table:table-cell table:style-name="Heading" office:value-type="string">
      <text:p>ID</text:p>
     </table:table-cell>
     <table:table-cell table:style-name="Heading" office:value-type="string">
      <text:p>Name</text:p>
     </table:table-cell>
     <table:table-cell table:style-name="Heading" office:value-type="string">
      <text:p>Description</text:p>
     </table:table-cell>
     <table:table-cell table:style-name="Heading" office:value-type="string">
      <text:p>Aka</text:p>
     </table:table-cell>
     <table:table-cell table:style-name="Heading" office:value-type="string">
      <text:p>Tags</text:p>
     </table:table-cell>
     <table:table-cell table:style-name="Heading" table:number-columns-repeated="1014"/>
    </table:table-row>

'''

    itemTemplate = '''   <table:table-row table:style-name="ro2">
     <table:table-cell office:value-type="string">
      <text:p>ItID:$ID</text:p>
     </table:table-cell>
     <table:table-cell office:value-type="string">
      <text:p>$Title</text:p>
     </table:table-cell>
     <table:table-cell office:value-type="string">
      <text:p>$Desc</text:p>
     </table:table-cell>
     <table:table-cell office:value-type="string">
      <text:p>$AKA</text:p>
     </table:table-cell>
     <table:table-cell office:value-type="string">
      <text:p>$Tags</text:p>
     </table:table-cell>
     <table:table-cell table:number-columns-repeated="1014"/>
    </table:table-row>

'''

    fileFooter = OdsFile.CONTENT_XML_FOOTER


class OdtCharacters(OdtFile):
    """ODT character descriptions file representation.

    Export a character sheet with invisibly tagged descriptions.
    """

    DESCRIPTION = 'Character descriptions'
    SUFFIX = '_characters'

    fileHeader = OdtFile.CONTENT_XML_HEADER + '''<text:p text:style-name="Title">$Title</text:p>
<text:p text:style-name="Subtitle">$AuthorName</text:p>
'''

    characterTemplate = '''<text:h text:style-name="Heading_20_2" text:outline-level="2">$Title$FullName$AKA</text:h>
<text:section text:style-name="Sect1" text:name="CrID:$ID">
<text:h text:style-name="Heading_20_3" text:outline-level="3">Description</text:h>
<text:section text:style-name="Sect1" text:name="CrID_desc:$ID">
<text:p text:style-name="Text_20_body">$Desc</text:p>
</text:section>
<text:h text:style-name="Heading_20_3" text:outline-level="3">Bio</text:h>
<text:section text:style-name="Sect1" text:name="CrID_bio:$ID">
<text:p text:style-name="Text_20_body">$Bio</text:p>
</text:section>
<text:h text:style-name="Heading_20_3" text:outline-level="3">Goals</text:h>
<text:section text:style-name="Sect1" text:name="CrID_goals:$ID">
<text:p text:style-name="Text_20_body">$Goals</text:p>
</text:section>
</text:section>
'''

    fileFooter = OdtFile.CONTENT_XML_FOOTER

    def get_characterMapping(self, crId):
        """Return a mapping dictionary for a character section. 
        """
        characterMapping = OdtFile.get_characterMapping(self, crId)

        if self.characters[crId].aka:
            characterMapping['AKA'] = ' ("' + self.characters[crId].aka + '")'

        if self.characters[crId].fullName:
            characterMapping['FullName'] = '/' + self.characters[crId].fullName

        return characterMapping


class OdtItems(OdtFile):
    """ODT item descriptions file representation.

    Export a item sheet with invisibly tagged descriptions.
    """

    DESCRIPTION = 'Item descriptions'
    SUFFIX = '_items'

    fileHeader = OdtFile.CONTENT_XML_HEADER + '''<text:p text:style-name="Title">$Title</text:p>
<text:p text:style-name="Subtitle">$AuthorName</text:p>
'''

    itemTemplate = '''<text:h text:style-name="Heading_20_2" text:outline-level="2">$Title$AKA</text:h>
<text:section text:style-name="Sect1" text:name="ItID:$ID">
<text:p text:style-name="Text_20_body">$Desc</text:p>
</text:section>
'''

    fileFooter = OdtFile.CONTENT_XML_FOOTER

    def get_itemMapping(self, itId):
        """Return a mapping dictionary for an item section. 
        """
        itemMapping = OdtFile.get_itemMapping(self, itId)

        if self.items[itId].aka:
            itemMapping['AKA'] = ' ("' + self.items[itId].aka + '")'

        return itemMapping


class OdtLocations(OdtFile):
    """ODT location descriptions file representation.

    Export a location sheet with invisibly tagged descriptions.
    """

    DESCRIPTION = 'Location descriptions'
    SUFFIX = '_locations'

    fileHeader = OdtFile.CONTENT_XML_HEADER + '''<text:p text:style-name="Title">$Title</text:p>
<text:p text:style-name="Subtitle">$AuthorName</text:p>
'''

    locationTemplate = '''<text:h text:style-name="Heading_20_2" text:outline-level="2">$Title$AKA</text:h>
<text:section text:style-name="Sect1" text:name="LcID:$ID">
<text:p text:style-name="Text_20_body">$Desc</text:p>
</text:section>
'''

    fileFooter = OdtFile.CONTENT_XML_FOOTER

    def get_locationMapping(self, lcId):
        """Return a mapping dictionary for a location section. 
        """
        locationMapping = OdtFile.get_locationMapping(self, lcId)

        if self.locations[lcId].aka:
            locationMapping['AKA'] = ' ("' + self.locations[lcId].aka + '")'

        return locationMapping

import uno
from com.sun.star.awt.MessageBoxType import MESSAGEBOX, INFOBOX, WARNINGBOX, ERRORBOX, QUERYBOX

from com.sun.star.awt.MessageBoxType import MESSAGEBOX, INFOBOX, WARNINGBOX, ERRORBOX, QUERYBOX
from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK, BUTTONS_OK_CANCEL, BUTTONS_YES_NO, BUTTONS_YES_NO_CANCEL, BUTTONS_RETRY_CANCEL, BUTTONS_ABORT_IGNORE_RETRY

CTX = uno.getComponentContext()
SM = CTX.getServiceManager()


def create_instance(name, with_context=False):
    if with_context:
        instance = SM.createInstanceWithContext(name, CTX)
    else:
        instance = SM.createInstance(name)
    return instance


def msgbox(message, title='yWriter import/export', buttons=BUTTONS_OK, type_msg=INFOBOX):
    """ Create message box
        type_msg: MESSAGEBOX, INFOBOX, WARNINGBOX, ERRORBOX, QUERYBOX

        MSG_BUTTONS: BUTTONS_OK, BUTTONS_OK_CANCEL, BUTTONS_YES_NO, 
        BUTTONS_YES_NO_CANCEL, BUTTONS_RETRY_CANCEL, BUTTONS_ABORT_IGNORE_RETRY

        MSG_RESULTS: OK, YES, NO, CANCEL

        http://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1awt_1_1XMessageBoxFactory.html
    """
    toolkit = create_instance('com.sun.star.awt.Toolkit')
    parent = toolkit.getDesktopWindow()
    mb = toolkit.createMessageBox(
        parent, type_msg, buttons, title, str(message))
    return mb.execute()


class Stub():

    def dummy(self):
        pass


def FilePicker(path=None, mode=0):
    """
    Read file:  `mode in (0, 6, 7, 8, 9)`
    Write file: `mode in (1, 2, 3, 4, 5, 10)`
    see: (http://api.libreoffice.org/docs/idl/ref/
            namespacecom_1_1sun_1_1star_1_1ui_1_1
            dialogs_1_1TemplateDescription.html)

    See: https://stackoverflow.com/questions/30840736/libreoffice-how-to-create-a-file-dialog-via-python-macro
    """
    # shortcut:
    createUnoService = (
        XSCRIPTCONTEXT
        .getComponentContext()
        .getServiceManager()
        .createInstance
    )

    filepicker = createUnoService("com.sun.star.ui.dialogs.OfficeFilePicker")

    if path:
        filepicker.setDisplayDirectory(path)

    filepicker.initialize((mode,))
    filepicker.appendFilter("yWriter 7 Files (.yw7)", "*.yw7")
    filepicker.appendFilter("yWriter 6 Files (.yw6)", "*.yw6")

    if filepicker.execute():
        return filepicker.getFiles()[0]


import sys



class Ui():
    """Base class for UI facades, implementing a 'silent mode'.
    """

    def __init__(self, title):
        """Initialize text buffers for messaging.
        """
        self.infoWhatText = ''
        self.infoHowText = ''

    def ask_yes_no(self, text):
        """The application may use a subclass  
        for confirmation requests.    
        """
        return True

    def set_info_what(self, message):
        """What's the converter going to do?"""
        self.infoWhatText = message

    def set_info_how(self, message):
        """How's the converter doing?"""
        self.infoHowText = message

    def start(self):
        """To be overridden by subclasses requiring
        special action to launch the user interaction.
        """


class YwCnv():
    """Base class for Novel file conversion.

    Public methods:
        convert(sourceFile, targetFile) -- Convert sourceFile into targetFile.
    """

    def convert(self, sourceFile, targetFile):
        """Convert sourceFile into targetFile and return a message.

        Positional arguments:
            sourceFile, targetFile -- Novel subclass instances.

        1. Make the source object read the source file.
        2. Make the target object merge the source object's instance variables.
        3. Make the target object write the target file.
        Return a message beginning with SUCCESS or ERROR.

        Error handling:
        - Check if sourceFile and targetFile are correctly initialized.
        - Ask for permission to overwrite targetFile.
        - Pass the error messages of the called methods of sourceFile and targetFile.
        - The success message comes from targetFile.write(), if called.       
        """

        # Initial error handling.

        if sourceFile.filePath is None:
            return 'ERROR: Source "' + os.path.normpath(sourceFile.filePath) + '" is not of the supported type.'

        if not sourceFile.file_exists():
            return 'ERROR: "' + os.path.normpath(sourceFile.filePath) + '" not found.'

        if targetFile.filePath is None:
            return 'ERROR: Target "' + os.path.normpath(targetFile.filePath) + '" is not of the supported type.'

        if targetFile.file_exists() and not self.confirm_overwrite(targetFile.filePath):
            return 'ERROR: Action canceled by user.'

        # Make the source object read the source file.

        message = sourceFile.read()

        if message.startswith('ERROR'):
            return message

        # Make the target object merge the source object's instance variables.

        message = targetFile.merge(sourceFile)

        if message.startswith('ERROR'):
            return message

        # Make the source object write the target file.

        return targetFile.write()

    def confirm_overwrite(self, fileName):
        """Return boolean permission to overwrite the target file.
        This is a stub to be overridden by subclass methods.
        """
        return True



class FileFactory:
    """Base class for conversion object factory classes.

    Public methods:
        make_file_objects(self, sourcePath, **kwargs) -- return conversion objects.

    This class emulates a "FileFactory" Interface.
    Instances can be used as stubs for factories instantiated at runtime.
    """

    def __init__(self, fileClasses=[]):
        """Write the parameter to a private instance variable.

        Positional arguments:
            fileClasses -- list of classes from which an instance can be returned.
        """
        self.fileClasses = fileClasses

    def make_file_objects(self, sourcePath, **kwargs):
        """A factory method stub.

        Positional arguments:
            sourcePath -- string; path to the source file to convert.

        Optional arguments:
            suffix -- string; an indicator for the target file type.

        Return a tuple with three elements:
        - A message string starting with 'ERROR'
        - sourceFile: None
        - targetFile: None

        Factory method to be overridden by subclasses.
        Subclasses return a tuple with three elements:
        - A message string starting with 'SUCCESS' or 'ERROR'
        - sourceFile: a Novel subclass instance
        - targetFile: a Novel subclass instance
        """
        return 'ERROR: File type of "' + os.path.normpath(sourcePath) + '" not supported.', None, None



class ExportSourceFactory(FileFactory):
    """A factory class that instantiates a yWriter object to read."""

    def make_file_objects(self, sourcePath, **kwargs):
        """Instantiate a source object for conversion from a yWriter project.
        Override the superclass method.

        Positional arguments:
            sourcePath -- string; path to the source file to convert.

        Return a tuple with three elements:
        - A message string starting with 'SUCCESS' or 'ERROR'
        - sourceFile: a YwFile subclass instance, or None in case of error
        - targetFile: None
        """
        fileName, fileExtension = os.path.splitext(sourcePath)

        for fileClass in self.fileClasses:

            if fileClass.EXTENSION == fileExtension:
                sourceFile = fileClass(sourcePath, **kwargs)
                return 'SUCCESS', sourceFile, None

        return 'ERROR: File type of "' + os.path.normpath(sourcePath) + '" not supported.', None, None



class ExportTargetFactory(FileFactory):
    """A factory class that instantiates a document object to write."""

    def make_file_objects(self, sourcePath, **kwargs):
        """Instantiate a target object for conversion from a yWriter project.
        Override the superclass method.

        Positional arguments:
            sourcePath -- string; path to the source file to convert.

        Optional arguments:
            suffix -- string; an indicator for the target file type.

        Return a tuple with three elements:
        - A message string starting with 'SUCCESS' or 'ERROR'
        - sourceFile: None
        - targetFile: a FileExport subclass instance, or None in case of error 
        """
        fileName, fileExtension = os.path.splitext(sourcePath)
        suffix = kwargs['suffix']

        for fileClass in self.fileClasses:

            if fileClass.SUFFIX == suffix:

                if suffix is None:
                    suffix = ''

                targetFile = fileClass(
                    fileName + suffix + fileClass.EXTENSION, **kwargs)
                return 'SUCCESS', None, targetFile

        return 'ERROR: File type of "' + os.path.normpath(sourcePath) + '" not supported.', None, None


class ImportSourceFactory(FileFactory):
    """A factory class that instantiates a documente object to read."""

    def make_file_objects(self, sourcePath, **kwargs):
        """Instantiate a source object for conversion to a yWriter project.       
        Override the superclass method.

        Positional arguments:
            sourcePath -- string; path to the source file to convert.

        Return a tuple with three elements:
        - A message string starting with 'SUCCESS' or 'ERROR'
        - sourceFile: a Novel subclass instance, or None in case of error
        - targetFile: None
        """

        for fileClass in self.fileClasses:

            if fileClass.SUFFIX is not None:

                if sourcePath.endswith(fileClass.SUFFIX + fileClass.EXTENSION):
                    sourceFile = fileClass(sourcePath, **kwargs)
                    return 'SUCCESS', sourceFile, None

        return 'ERROR: This document is not meant to be written back.', None, None



class ImportTargetFactory(FileFactory):
    """A factory class that instantiates a yWriter object to write."""

    def make_file_objects(self, sourcePath, **kwargs):
        """Instantiate a target object for conversion to a yWriter project.
        Override the superclass method.

        Positional arguments:
            sourcePath -- string; path to the source file to convert.

        Optional arguments:
            suffix -- string; an indicator for the source file type.

        Return a tuple with three elements:
        - A message string starting with 'SUCCESS' or 'ERROR'
        - sourceFile: None
        - targetFile: a YwFile subclass instance, or None in case of error

        """
        fileName, fileExtension = os.path.splitext(sourcePath)
        sourceSuffix = kwargs['suffix']

        if sourceSuffix:
            ywPathBasis = fileName.split(sourceSuffix)[0]

        else:
            ywPathBasis = fileName

        # Look for an existing yWriter project to rewrite.

        for fileClass in self.fileClasses:

            if os.path.isfile(ywPathBasis + fileClass.EXTENSION):
                targetFile = fileClass(
                    ywPathBasis + fileClass.EXTENSION, **kwargs)
                return 'SUCCESS', None, targetFile

        return 'ERROR: No yWriter project to write.', None, None


class YwCnvUi(YwCnv):
    """Base class for Novel file conversion with user interface.

    Public methods:
        run(sourcePath, suffix) -- Create source and target objects and run conversion.

    Class constants:
        EXPORT_SOURCE_CLASSES -- List of YwFile subclasses from which can be exported.
        EXPORT_TARGET_CLASSES -- List of FileExport subclasses to which export is possible.
        IMPORT_SOURCE_CLASSES -- List of Novel subclasses from which can be imported.
        IMPORT_TARGET_CLASSES -- List of YwFile subclasses to which import is possible.

    All lists are empty and meant to be overridden by subclasses.

    Instance variables:
        ui -- Ui (can be overridden e.g. by subclasses).
        exportSourceFactory -- ExportSourceFactory.
        exportTargetFactory -- ExportTargetFactory.
        importSourceFactory -- ImportSourceFactory.
        importTargetFactory -- ImportTargetFactory.
        newProjectFactory -- FileFactory (a stub to be overridden by subclasses).
        newFile -- string; path to the target file in case of success.   
    """

    EXPORT_SOURCE_CLASSES = []
    EXPORT_TARGET_CLASSES = []
    IMPORT_SOURCE_CLASSES = []
    IMPORT_TARGET_CLASSES = []

    def __init__(self):
        """Define instance variables."""
        self.ui = Ui('')
        # Per default, 'silent mode' is active.

        self.exportSourceFactory = ExportSourceFactory(
            self.EXPORT_SOURCE_CLASSES)
        self.exportTargetFactory = ExportTargetFactory(
            self.EXPORT_TARGET_CLASSES)
        self.importSourceFactory = ImportSourceFactory(
            self.IMPORT_SOURCE_CLASSES)
        self.importTargetFactory = ImportTargetFactory(
            self.IMPORT_TARGET_CLASSES)
        self.newProjectFactory = FileFactory()

        self.newFile = None
        # Also indicates successful conversion.

    def run(self, sourcePath, **kwargs):
        """Create source and target objects and run conversion.

        sourcePath -- str; the source file path.
        suffix -- str; target file name suffix. 

        This is a template method that calls primitive operations by case.
        """
        if not os.path.isfile(sourcePath):
            self.ui.set_info_how(
                'ERROR: File "' + os.path.normpath(sourcePath) + '" not found.')
            return

        message, sourceFile, dummy = self.exportSourceFactory.make_file_objects(
            sourcePath, **kwargs)

        if message.startswith('SUCCESS'):
            # The source file is a yWriter project.

            message, dummy, targetFile = self.exportTargetFactory.make_file_objects(
                sourcePath, **kwargs)

            if message.startswith('SUCCESS'):
                self.export_from_yw(sourceFile, targetFile)

            else:
                self.ui.set_info_how(message)

        else:
            # The source file is not a yWriter project.

            message, sourceFile, dummy = self.importSourceFactory.make_file_objects(
                sourcePath, **kwargs)

            if message.startswith('SUCCESS'):
                kwargs['suffix'] = sourceFile.SUFFIX
                message, dummy, targetFile = self.importTargetFactory.make_file_objects(
                    sourcePath, **kwargs)

                if message.startswith('SUCCESS'):
                    self.import_to_yw(sourceFile, targetFile)

                else:
                    self.ui.set_info_how(message)

            else:
                # A new yWriter project might be required.

                message, sourceFile, targetFile = self.newProjectFactory.make_file_objects(
                    sourcePath, **kwargs)

                if message.startswith('SUCCESS'):
                    self.create_yw7(sourceFile, targetFile)

                else:
                    self.ui.set_info_how(message)

    def export_from_yw(self, sourceFile, targetFile):
        """Convert from yWriter project to other file format.

        sourceFile -- YwFile subclass instance.
        targetFile -- Any Novel subclass instance.

        This is a primitive operation of the run() template method.

        1. Send specific information about the conversion to the UI.
        2. Convert sourceFile into targetFile.
        3. Pass the message to the UI.
        4. Save the new file pathname.

        Error handling:
        - If the conversion fails, newFile is set to None.
        """

        # Send specific information about the conversion to the UI.

        self.ui.set_info_what('Input: ' + sourceFile.DESCRIPTION + ' "' + os.path.normpath(
            sourceFile.filePath) + '"\nOutput: ' + targetFile.DESCRIPTION + ' "' + os.path.normpath(targetFile.filePath) + '"')

        # Convert sourceFile into targetFile.

        message = self.convert(sourceFile, targetFile)

        # Pass the message to the UI.

        self.ui.set_info_how(message)

        # Save the new file pathname.

        if message.startswith('SUCCESS'):
            self.newFile = targetFile.filePath

        else:
            self.newFile = None

    def create_yw7(self, sourceFile, targetFile):
        """Create targetFile from sourceFile.

        sourceFile -- Any Novel subclass instance.
        targetFile -- YwFile subclass instance.

        This is a primitive operation of the run() template method.

        1. Send specific information about the conversion to the UI.
        2. Convert sourceFile into targetFile.
        3. Pass the message to the UI.
        4. Save the new file pathname.

        Error handling:
        - Tf targetFile already exists as a file, the conversion is cancelled,
          an error message is sent to the UI.
        - If the conversion fails, newFile is set to None.
        """

        # Send specific information about the conversion to the UI.

        self.ui.set_info_what(
            'Create a yWriter project file from ' + sourceFile.DESCRIPTION + '\nNew project: "' + os.path.normpath(targetFile.filePath) + '"')

        if targetFile.file_exists():
            self.ui.set_info_how(
                'ERROR: "' + os.path.normpath(targetFile.filePath) + '" already exists.')

        else:
            # Convert sourceFile into targetFile.

            message = self.convert(sourceFile, targetFile)

            # Pass the message to the UI.

            self.ui.set_info_how(message)

            # Save the new file pathname.

            if message.startswith('SUCCESS'):
                self.newFile = targetFile.filePath

            else:
                self.newFile = None

    def import_to_yw(self, sourceFile, targetFile):
        """Convert from any file format to yWriter project.

        sourceFile -- Any Novel subclass instance.
        targetFile -- YwFile subclass instance.

        This is a primitive operation of the run() template method.

        1. Send specific information about the conversion to the UI.
        2. Convert sourceFile into targetFile.
        3. Pass the message to the UI.
        4. Delete the temporay file, if exists.
        5. Save the new file pathname.

        Error handling:
        - If the conversion fails, newFile is set to None.
        """

        # Send specific information about the conversion to the UI.

        self.ui.set_info_what('Input: ' + sourceFile.DESCRIPTION + ' "' + os.path.normpath(
            sourceFile.filePath) + '"\nOutput: ' + targetFile.DESCRIPTION + ' "' + os.path.normpath(targetFile.filePath) + '"')

        # Convert sourceFile into targetFile.

        message = self.convert(sourceFile, targetFile)

        # Pass the message to the UI.

        self.ui.set_info_how(message)

        # Delete the temporay file, if exists.

        self.delete_tempfile(sourceFile.filePath)

        # Save the new file pathname.

        if message.startswith('SUCCESS'):
            self.newFile = targetFile.filePath

        else:
            self.newFile = None

    def confirm_overwrite(self, filePath):
        """Return boolean permission to overwrite the target file, overriding the superclass method."""
        return self.ui.ask_yes_no('Overwrite existing file "' + os.path.normpath(filePath) + '"?')

    def delete_tempfile(self, filePath):
        """Delete filePath if it is a temporary file no longer needed."""

        if filePath.endswith('.html'):
            # Might it be a temporary text document?

            if os.path.isfile(filePath.replace('.html', '.odt')):
                # Does a corresponding Office document exist?

                try:
                    os.remove(filePath)

                except:
                    pass

        elif filePath.endswith('.csv'):
            # Might it be a temporary spreadsheet document?

            if os.path.isfile(filePath.replace('.csv', '.ods')):
                # Does a corresponding Office document exist?

                try:
                    os.remove(filePath)

                except:
                    pass

    def open_newFile(self):
        """Open the converted file for editing and exit the converter script."""
        os.startfile(self.newFile)
        sys.exit(0)








class Chapter():
    """yWriter chapter representation.
    # xml: <CHAPTERS><CHAPTER>
    """

    chapterTitlePrefix = "Chapter "
    # str
    # Can be changed at runtime for non-English projects.

    def __init__(self):
        self.title = None
        # str
        # xml: <Title>

        self.desc = None
        # str
        # xml: <Desc>

        self.chLevel = None
        # int
        # xml: <SectionStart>
        # 0 = chapter level
        # 1 = section level ("this chapter begins a section")

        self.oldType = None
        # int
        # xml: <Type>
        # 0 = chapter type (marked "Chapter")
        # 1 = other type (marked "Other")

        self.chType = None
        # int
        # xml: <ChapterType>
        # 0 = Normal
        # 1 = Notes
        # 2 = Todo

        self.isUnused = None
        # bool
        # xml: <Unused> -1

        self.suppressChapterTitle = None
        # bool
        # xml: <Fields><Field_SuppressChapterTitle> 1
        # True: Chapter heading not to be displayed in written document.
        # False: Chapter heading to be displayed in written document.

        self.isTrash = None
        # bool
        # xml: <Fields><Field_IsTrash> 1
        # True: This chapter is the yw7 project's "trash bin".
        # False: This chapter is not a "trash bin".

        self.suppressChapterBreak = None
        # bool
        # xml: <Fields><Field_SuppressChapterBreak> 0

        self.srtScenes = []
        # list of str
        # xml: <Scenes><ScID>
        # The chapter's scene IDs. The order of its elements
        # corresponds to the chapter's order of the scenes.

    def get_title(self):
        """Fix auto-chapter titles if necessary 
        """
        text = self.title

        if text:
            text = text.replace('Chapter ', self.chapterTitlePrefix)

        return text


class YwFile(Novel):
    """Abstract yWriter xml project file representation.
    """

    def strip_spaces(self, elements):
        """remove leading and trailing spaces from the elements
        of a list of strings.
        """
        stripped = []

        for element in elements:
            stripped.append(element.lstrip().rstrip())

        return stripped

    def read(self):
        """Parse the yWriter xml file located at filePath, fetching the Novel attributes.
        Return a message beginning with SUCCESS or ERROR.
        """

        if self.is_locked():
            return 'ERROR: yWriter seems to be open. Please close first.'

        message = self.ywTreeReader.read_element_tree(self)

        if message.startswith('ERROR'):
            return message

        root = self._tree.getroot()

        # Read locations from the xml element tree.

        for loc in root.iter('LOCATION'):
            lcId = loc.find('ID').text
            self.srtLocations.append(lcId)
            self.locations[lcId] = WorldElement()

            if loc.find('Title') is not None:
                self.locations[lcId].title = loc.find('Title').text

            if loc.find('ImageFile') is not None:
                self.locations[lcId].image = loc.find('ImageFile').text

            if loc.find('Desc') is not None:
                self.locations[lcId].desc = loc.find('Desc').text

            if loc.find('AKA') is not None:
                self.locations[lcId].aka = loc.find('AKA').text

            if loc.find('Tags') is not None:

                if loc.find('Tags').text is not None:
                    tags = loc.find('Tags').text.split(';')
                    self.locations[lcId].tags = self.strip_spaces(tags)

        # Read items from the xml element tree.

        for itm in root.iter('ITEM'):
            itId = itm.find('ID').text
            self.srtItems.append(itId)
            self.items[itId] = WorldElement()

            if itm.find('Title') is not None:
                self.items[itId].title = itm.find('Title').text

            if itm.find('ImageFile') is not None:
                self.items[itId].image = itm.find('ImageFile').text

            if itm.find('Desc') is not None:
                self.items[itId].desc = itm.find('Desc').text

            if itm.find('AKA') is not None:
                self.items[itId].aka = itm.find('AKA').text

            if itm.find('Tags') is not None:

                if itm.find('Tags').text is not None:
                    tags = itm.find('Tags').text.split(';')
                    self.items[itId].tags = self.strip_spaces(tags)

        # Read characters from the xml element tree.

        for crt in root.iter('CHARACTER'):
            crId = crt.find('ID').text
            self.srtCharacters.append(crId)
            self.characters[crId] = Character()

            if crt.find('Title') is not None:
                self.characters[crId].title = crt.find('Title').text

            if crt.find('ImageFile') is not None:
                self.characters[crId].image = crt.find('ImageFile').text

            if crt.find('Desc') is not None:
                self.characters[crId].desc = crt.find('Desc').text

            if crt.find('AKA') is not None:
                self.characters[crId].aka = crt.find('AKA').text

            if crt.find('Tags') is not None:

                if crt.find('Tags').text is not None:
                    tags = crt.find('Tags').text.split(';')
                    self.characters[crId].tags = self.strip_spaces(tags)

            if crt.find('Notes') is not None:
                self.characters[crId].notes = crt.find('Notes').text

            if crt.find('Bio') is not None:
                self.characters[crId].bio = crt.find('Bio').text

            if crt.find('Goals') is not None:
                self.characters[crId].goals = crt.find('Goals').text

            if crt.find('FullName') is not None:
                self.characters[crId].fullName = crt.find('FullName').text

            if crt.find('Major') is not None:
                self.characters[crId].isMajor = True

            else:
                self.characters[crId].isMajor = False

        # Read attributes at novel level from the xml element tree.

        prj = root.find('PROJECT')

        if prj.find('Title') is not None:
            self.title = prj.find('Title').text

        if prj.find('AuthorName') is not None:
            self.author = prj.find('AuthorName').text

        if prj.find('Desc') is not None:
            self.desc = prj.find('Desc').text

        if prj.find('FieldTitle1') is not None:
            self.fieldTitle1 = prj.find('FieldTitle1').text

        if prj.find('FieldTitle2') is not None:
            self.fieldTitle2 = prj.find('FieldTitle2').text

        if prj.find('FieldTitle3') is not None:
            self.fieldTitle3 = prj.find('FieldTitle3').text

        if prj.find('FieldTitle4') is not None:
            self.fieldTitle4 = prj.find('FieldTitle4').text

        # Read attributes at chapter level from the xml element tree.

        for chp in root.iter('CHAPTER'):
            chId = chp.find('ID').text
            self.chapters[chId] = Chapter()
            self.srtChapters.append(chId)

            if chp.find('Title') is not None:
                self.chapters[chId].title = chp.find('Title').text

            if chp.find('Desc') is not None:
                self.chapters[chId].desc = chp.find('Desc').text

            if chp.find('SectionStart') is not None:
                self.chapters[chId].chLevel = 1

            else:
                self.chapters[chId].chLevel = 0

            if chp.find('Type') is not None:
                self.chapters[chId].oldType = int(chp.find('Type').text)

            if chp.find('ChapterType') is not None:
                self.chapters[chId].chType = int(chp.find('ChapterType').text)

            if chp.find('Unused') is not None:
                self.chapters[chId].isUnused = True

            else:
                self.chapters[chId].isUnused = False

            self.chapters[chId].suppressChapterTitle = False

            if self.chapters[chId].title is not None:

                if self.chapters[chId].title.startswith('@'):
                    self.chapters[chId].suppressChapterTitle = True

            for chFields in chp.findall('Fields'):

                if chFields.find('Field_SuppressChapterTitle') is not None:

                    if chFields.find('Field_SuppressChapterTitle').text == '1':
                        self.chapters[chId].suppressChapterTitle = True

                if chFields.find('Field_IsTrash') is not None:

                    if chFields.find('Field_IsTrash').text == '1':
                        self.chapters[chId].isTrash = True

                    else:
                        self.chapters[chId].isTrash = False

                if chFields.find('Field_SuppressChapterBreak') is not None:

                    if chFields.find('Field_SuppressChapterBreak').text == '1':
                        self.chapters[chId].suppressChapterBreak = True

                    else:
                        self.chapters[chId].suppressChapterBreak = False

                else:
                    self.chapters[chId].suppressChapterBreak = False

            self.chapters[chId].srtScenes = []

            if chp.find('Scenes') is not None:

                if not self.chapters[chId].isTrash:

                    for scn in chp.find('Scenes').findall('ScID'):
                        scId = scn.text
                        self.chapters[chId].srtScenes.append(scId)

        # Read attributes at scene level from the xml element tree.

        for scn in root.iter('SCENE'):
            scId = scn.find('ID').text
            self.scenes[scId] = Scene()

            if scn.find('Title') is not None:
                self.scenes[scId].title = scn.find('Title').text

            if scn.find('Desc') is not None:
                self.scenes[scId].desc = scn.find('Desc').text

            if scn.find('RTFFile') is not None:
                self.scenes[scId].rtfFile = scn.find('RTFFile').text

            # This is relevant for yW5 files with no SceneContent:

            if scn.find('WordCount') is not None:
                self.scenes[scId].wordCount = int(
                    scn.find('WordCount').text)

            if scn.find('LetterCount') is not None:
                self.scenes[scId].letterCount = int(
                    scn.find('LetterCount').text)

            if scn.find('SceneContent') is not None:
                sceneContent = scn.find('SceneContent').text

                if sceneContent is not None:
                    self.scenes[scId].sceneContent = sceneContent

            if scn.find('Unused') is not None:
                self.scenes[scId].isUnused = True

            else:
                self.scenes[scId].isUnused = False

            for scFields in scn.findall('Fields'):

                if scFields.find('Field_SceneType') is not None:

                    if scFields.find('Field_SceneType').text == '1':
                        self.scenes[scId].isNotesScene = True

                    if scFields.find('Field_SceneType').text == '2':
                        self.scenes[scId].isTodoScene = True

            if scn.find('ExportCondSpecific') is None:
                self.scenes[scId].doNotExport = False

            elif scn.find('ExportWhenRTF') is not None:
                self.scenes[scId].doNotExport = False

            else:
                self.scenes[scId].doNotExport = True

            if scn.find('Status') is not None:
                self.scenes[scId].status = int(scn.find('Status').text)

            if scn.find('Notes') is not None:
                self.scenes[scId].sceneNotes = scn.find('Notes').text

            if scn.find('Tags') is not None:

                if scn.find('Tags').text is not None:
                    tags = scn.find('Tags').text.split(';')
                    self.scenes[scId].tags = self.strip_spaces(tags)

            if scn.find('Field1') is not None:
                self.scenes[scId].field1 = scn.find('Field1').text

            if scn.find('Field2') is not None:
                self.scenes[scId].field2 = scn.find('Field2').text

            if scn.find('Field3') is not None:
                self.scenes[scId].field3 = scn.find('Field3').text

            if scn.find('Field4') is not None:
                self.scenes[scId].field4 = scn.find('Field4').text

            if scn.find('AppendToPrev') is not None:
                self.scenes[scId].appendToPrev = True

            else:
                self.scenes[scId].appendToPrev = False

            if scn.find('SpecificDateTime') is not None:
                dateTime = scn.find('SpecificDateTime').text.split(' ')

                for dt in dateTime:

                    if '-' in dt:
                        self.scenes[scId].date = dt

                    elif ':' in dt:
                        self.scenes[scId].time = dt

            else:
                if scn.find('Day') is not None:
                    self.scenes[scId].day = scn.find('Day').text

                if scn.find('Hour') is not None:
                    self.scenes[scId].hour = scn.find('Hour').text

                if scn.find('Minute') is not None:
                    self.scenes[scId].minute = scn.find('Minute').text

            if scn.find('LastsDays') is not None:
                self.scenes[scId].lastsDays = scn.find('LastsDays').text

            if scn.find('LastsHours') is not None:
                self.scenes[scId].lastsHours = scn.find('LastsHours').text

            if scn.find('LastsMinutes') is not None:
                self.scenes[scId].lastsMinutes = scn.find('LastsMinutes').text

            if scn.find('ReactionScene') is not None:
                self.scenes[scId].isReactionScene = True

            else:
                self.scenes[scId].isReactionScene = False

            if scn.find('SubPlot') is not None:
                self.scenes[scId].isSubPlot = True

            else:
                self.scenes[scId].isSubPlot = False

            if scn.find('Goal') is not None:
                self.scenes[scId].goal = scn.find('Goal').text

            if scn.find('Conflict') is not None:
                self.scenes[scId].conflict = scn.find('Conflict').text

            if scn.find('Outcome') is not None:
                self.scenes[scId].outcome = scn.find('Outcome').text

            if scn.find('Characters') is not None:
                for crId in scn.find('Characters').iter('CharID'):

                    if self.scenes[scId].characters is None:
                        self.scenes[scId].characters = []

                    self.scenes[scId].characters.append(crId.text)

            if scn.find('Locations') is not None:
                for lcId in scn.find('Locations').iter('LocID'):

                    if self.scenes[scId].locations is None:
                        self.scenes[scId].locations = []

                    self.scenes[scId].locations.append(lcId.text)

            if scn.find('Items') is not None:
                for itId in scn.find('Items').iter('ItemID'):

                    if self.scenes[scId].items is None:
                        self.scenes[scId].items = []

                    self.scenes[scId].items.append(itId.text)

        return 'SUCCESS: ' + str(len(self.scenes)) + ' Scenes read from "' + os.path.normpath(self.filePath) + '".'

    def merge(self, novel):
        """Copy required attributes of the novel object.
        Return a message beginning with SUCCESS or ERROR.
        """

        if self.file_exists():
            message = self.read()
            # initialize data

            if message.startswith('ERROR'):
                return message

        return self.ywProjectMerger.merge_projects(self, novel)

    def write(self):
        """Open the yWriter xml file located at filePath and 
        replace a set of attributes not being None.
        Return a message beginning with SUCCESS or ERROR.
        """

        if self.is_locked():
            return 'ERROR: yWriter seems to be open. Please close first.'

        message = self.ywTreeBuilder.build_element_tree(self)

        if message.startswith('ERROR'):
            return message

        message = self.ywTreeWriter.write_element_tree(self)

        if message.startswith('ERROR'):
            return message

        return self.ywPostprocessor.postprocess_xml_file(self.filePath)

    def is_locked(self):
        """Test whether a .lock file placed by yWriter exists.
        """
        if os.path.isfile(self.filePath + '.lock'):
            return True

        else:
            return False

import xml.etree.ElementTree as ET


class YwTreeBuilder():
    """Build yWriter project xml tree."""

    def build_element_tree(self, ywProject):
        """Modify the yWriter project attributes of an existing xml element tree.
        Return a message beginning with SUCCESS or ERROR.
        To be overridden by file format specific subclasses.
        """
        root = ywProject._tree.getroot()

        # Write locations to the xml element tree.

        locations = root.find('LOCATIONS')
        sortOrder = 0

        # Remove LOCATION entries in order to rewrite
        # the LOCATIONS section in a modified sort order.

        for loc in locations.findall('LOCATION'):
            locations.remove(loc)

        for lcId in ywProject.srtLocations:
            loc = ET.SubElement(locations, 'LOCATION')
            ET.SubElement(loc, 'ID').text = lcId

            if ywProject.locations[lcId].title is not None:
                ET.SubElement(
                    loc, 'Title').text = ywProject.locations[lcId].title

            if ywProject.locations[lcId].image is not None:
                ET.SubElement(
                    loc, 'ImageFile').text = ywProject.locations[lcId].image

            if ywProject.locations[lcId].desc is not None:
                ET.SubElement(
                    loc, 'Desc').text = ywProject.locations[lcId].desc

            if ywProject.locations[lcId].aka is not None:
                ET.SubElement(loc, 'AKA').text = ywProject.locations[lcId].aka

            if ywProject.locations[lcId].tags is not None:
                ET.SubElement(loc, 'Tags').text = ';'.join(
                    ywProject.locations[lcId].tags)

            sortOrder += 1
            ET.SubElement(loc, 'SortOrder').text = str(sortOrder)

        # Write items to the xml element tree.

        items = root.find('ITEMS')
        sortOrder = 0

        # Remove ITEM entries in order to rewrite
        # the ITEMS section in a modified sort order.

        for itm in items.findall('ITEM'):
            items.remove(itm)

        for itId in ywProject.srtItems:
            itm = ET.SubElement(items, 'ITEM')
            ET.SubElement(itm, 'ID').text = itId

            if ywProject.items[itId].title is not None:
                ET.SubElement(itm, 'Title').text = ywProject.items[itId].title

            if ywProject.items[itId].image is not None:
                ET.SubElement(
                    itm, 'ImageFile').text = ywProject.items[itId].image

            if ywProject.items[itId].desc is not None:
                ET.SubElement(itm, 'Desc').text = ywProject.items[itId].desc

            if ywProject.items[itId].aka is not None:
                ET.SubElement(itm, 'AKA').text = ywProject.items[itId].aka

            if ywProject.items[itId].tags is not None:
                ET.SubElement(itm, 'Tags').text = ';'.join(
                    ywProject.items[itId].tags)

            sortOrder += 1
            ET.SubElement(itm, 'SortOrder').text = str(sortOrder)

        # Write characters to the xml element tree.

        characters = root.find('CHARACTERS')
        sortOrder = 0

        # Remove CHARACTER entries in order to rewrite
        # the CHARACTERS section in a modified sort order.

        for crt in characters.findall('CHARACTER'):
            characters.remove(crt)

        for crId in ywProject.srtCharacters:
            crt = ET.SubElement(characters, 'CHARACTER')
            ET.SubElement(crt, 'ID').text = crId

            if ywProject.characters[crId].title is not None:
                ET.SubElement(
                    crt, 'Title').text = ywProject.characters[crId].title

            if ywProject.characters[crId].desc is not None:
                ET.SubElement(
                    crt, 'Desc').text = ywProject.characters[crId].desc

            if ywProject.characters[crId].image is not None:
                ET.SubElement(
                    crt, 'ImageFile').text = ywProject.characters[crId].image

            sortOrder += 1
            ET.SubElement(crt, 'SortOrder').text = str(sortOrder)

            if ywProject.characters[crId].notes is not None:
                ET.SubElement(
                    crt, 'Notes').text = ywProject.characters[crId].notes

            if ywProject.characters[crId].aka is not None:
                ET.SubElement(crt, 'AKA').text = ywProject.characters[crId].aka

            if ywProject.characters[crId].tags is not None:
                ET.SubElement(crt, 'Tags').text = ';'.join(
                    ywProject.characters[crId].tags)

            if ywProject.characters[crId].bio is not None:
                ET.SubElement(crt, 'Bio').text = ywProject.characters[crId].bio

            if ywProject.characters[crId].goals is not None:
                ET.SubElement(
                    crt, 'Goals').text = ywProject.characters[crId].goals

            if ywProject.characters[crId].fullName is not None:
                ET.SubElement(
                    crt, 'FullName').text = ywProject.characters[crId].fullName

            if ywProject.characters[crId].isMajor:
                ET.SubElement(crt, 'Major').text = '-1'

        # Write attributes at novel level to the xml element tree.

        prj = root.find('PROJECT')
        prj.find('Title').text = ywProject.title

        if ywProject.desc is not None:

            if prj.find('Desc') is None:
                ET.SubElement(prj, 'Desc').text = ywProject.desc

            else:
                prj.find('Desc').text = ywProject.desc

        if ywProject.author is not None:

            if prj.find('AuthorName') is None:
                ET.SubElement(prj, 'AuthorName').text = ywProject.author

            else:
                prj.find('AuthorName').text = ywProject.author

        prj.find('FieldTitle1').text = ywProject.fieldTitle1
        prj.find('FieldTitle2').text = ywProject.fieldTitle2
        prj.find('FieldTitle3').text = ywProject.fieldTitle3
        prj.find('FieldTitle4').text = ywProject.fieldTitle4

        # Write attributes at chapter level to the xml element tree.

        for chp in root.iter('CHAPTER'):
            chId = chp.find('ID').text

            if chId in ywProject.chapters:
                chp.find('Title').text = ywProject.chapters[chId].title

                if ywProject.chapters[chId].desc is not None:

                    if chp.find('Desc') is None:
                        ET.SubElement(
                            chp, 'Desc').text = ywProject.chapters[chId].desc

                    else:
                        chp.find('Desc').text = ywProject.chapters[chId].desc

                levelInfo = chp.find('SectionStart')

                if levelInfo is not None:

                    if ywProject.chapters[chId].chLevel == 0:
                        chp.remove(levelInfo)

                chp.find('Type').text = str(ywProject.chapters[chId].oldType)

                if ywProject.chapters[chId].chType is not None:

                    if chp.find('ChapterType') is not None:
                        chp.find('ChapterType').text = str(
                            ywProject.chapters[chId].chType)
                    else:
                        ET.SubElement(chp, 'ChapterType').text = str(
                            ywProject.chapters[chId].chType)

                if ywProject.chapters[chId].isUnused:

                    if chp.find('Unused') is None:
                        ET.SubElement(chp, 'Unused').text = '-1'

                elif chp.find('Unused') is not None:
                    chp.remove(chp.find('Unused'))

        # Write attributes at scene level to the xml element tree.

        for scn in root.iter('SCENE'):
            scId = scn.find('ID').text

            if scId in ywProject.scenes:

                if ywProject.scenes[scId].title is not None:
                    scn.find('Title').text = ywProject.scenes[scId].title

                if ywProject.scenes[scId].desc is not None:

                    if scn.find('Desc') is None:
                        ET.SubElement(
                            scn, 'Desc').text = ywProject.scenes[scId].desc

                    else:
                        scn.find('Desc').text = ywProject.scenes[scId].desc

                # Scene content is written in subclasses.

                if ywProject.scenes[scId].isUnused:

                    if scn.find('Unused') is None:
                        ET.SubElement(scn, 'Unused').text = '-1'

                elif scn.find('Unused') is not None:
                    scn.remove(scn.find('Unused'))

                if ywProject.scenes[scId].isNotesScene:

                    if scn.find('Fields') is None:
                        scFields = ET.SubElement(scn, 'Fields')

                    else:
                        scFields = scn.find('Fields')

                    if scFields.find('Field_SceneType') is None:
                        ET.SubElement(scFields, 'Field_SceneType').text = '1'

                elif scn.find('Fields') is not None:
                    scFields = scn.find('Fields')

                    if scFields.find('Field_SceneType') is not None:

                        if scFields.find('Field_SceneType').text == '1':
                            scFields.remove(scFields.find('Field_SceneType'))

                if ywProject.scenes[scId].isTodoScene:

                    if scn.find('Fields') is None:
                        scFields = ET.SubElement(scn, 'Fields')

                    else:
                        scFields = scn.find('Fields')

                    if scFields.find('Field_SceneType') is None:
                        ET.SubElement(scFields, 'Field_SceneType').text = '2'

                elif scn.find('Fields') is not None:
                    scFields = scn.find('Fields')

                    if scFields.find('Field_SceneType') is not None:

                        if scFields.find('Field_SceneType').text == '2':
                            scFields.remove(scFields.find('Field_SceneType'))

                if ywProject.scenes[scId].status is not None:
                    scn.find('Status').text = str(
                        ywProject.scenes[scId].status)

                if ywProject.scenes[scId].sceneNotes is not None:

                    if scn.find('Notes') is None:
                        ET.SubElement(
                            scn, 'Notes').text = ywProject.scenes[scId].sceneNotes

                    else:
                        scn.find(
                            'Notes').text = ywProject.scenes[scId].sceneNotes

                if ywProject.scenes[scId].tags is not None:

                    if scn.find('Tags') is None:
                        ET.SubElement(scn, 'Tags').text = ';'.join(
                            ywProject.scenes[scId].tags)

                    else:
                        scn.find('Tags').text = ';'.join(
                            ywProject.scenes[scId].tags)

                if ywProject.scenes[scId].field1 is not None:

                    if scn.find('Field1') is None:
                        ET.SubElement(
                            scn, 'Field1').text = ywProject.scenes[scId].field1

                    else:
                        scn.find('Field1').text = ywProject.scenes[scId].field1

                if ywProject.scenes[scId].field2 is not None:

                    if scn.find('Field2') is None:
                        ET.SubElement(
                            scn, 'Field2').text = ywProject.scenes[scId].field2

                    else:
                        scn.find('Field2').text = ywProject.scenes[scId].field2

                if ywProject.scenes[scId].field3 is not None:

                    if scn.find('Field3') is None:
                        ET.SubElement(
                            scn, 'Field3').text = ywProject.scenes[scId].field3

                    else:
                        scn.find('Field3').text = ywProject.scenes[scId].field3

                if ywProject.scenes[scId].field4 is not None:

                    if scn.find('Field4') is None:
                        ET.SubElement(
                            scn, 'Field4').text = ywProject.scenes[scId].field4

                    else:
                        scn.find('Field4').text = ywProject.scenes[scId].field4

                if ywProject.scenes[scId].appendToPrev:

                    if scn.find('AppendToPrev') is None:
                        ET.SubElement(scn, 'AppendToPrev').text = '-1'

                elif scn.find('AppendToPrev') is not None:
                    scn.remove(scn.find('AppendToPrev'))

                # Date/time information

                if (ywProject.scenes[scId].date is not None) and (ywProject.scenes[scId].time is not None):
                    dateTime = ywProject.scenes[scId].date + \
                        ' ' + ywProject.scenes[scId].time

                    if scn.find('SpecificDateTime') is not None:
                        scn.find('SpecificDateTime').text = dateTime

                    else:
                        ET.SubElement(scn, 'SpecificDateTime').text = dateTime
                        ET.SubElement(scn, 'SpecificDateMode').text = '-1'

                        if scn.find('Day') is not None:
                            scn.remove(scn.find('Day'))

                        if scn.find('Hour') is not None:
                            scn.remove(scn.find('Hour'))

                        if scn.find('Minute') is not None:
                            scn.remove(scn.find('Minute'))

                elif (ywProject.scenes[scId].day is not None) or (ywProject.scenes[scId].hour is not None) or (ywProject.scenes[scId].minute is not None):

                    if scn.find('SpecificDateTime') is not None:
                        scn.remove(scn.find('SpecificDateTime'))

                    if scn.find('SpecificDateMode') is not None:
                        scn.remove(scn.find('SpecificDateMode'))

                    if ywProject.scenes[scId].day is not None:

                        if scn.find('Day') is not None:
                            scn.find('Day').text = ywProject.scenes[scId].day

                        else:
                            ET.SubElement(
                                scn, 'Day').text = ywProject.scenes[scId].day

                    if ywProject.scenes[scId].hour is not None:

                        if scn.find('Hour') is not None:
                            scn.find('Hour').text = ywProject.scenes[scId].hour

                        else:
                            ET.SubElement(
                                scn, 'Hour').text = ywProject.scenes[scId].hour

                    if ywProject.scenes[scId].minute is not None:

                        if scn.find('Minute') is not None:
                            scn.find(
                                'Minute').text = ywProject.scenes[scId].minute

                        else:
                            ET.SubElement(
                                scn, 'Minute').text = ywProject.scenes[scId].minute

                if ywProject.scenes[scId].lastsDays is not None:

                    if scn.find('LastsDays') is not None:
                        scn.find(
                            'LastsDays').text = ywProject.scenes[scId].lastsDays

                    else:
                        ET.SubElement(
                            scn, 'LastsDays').text = ywProject.scenes[scId].lastsDays

                if ywProject.scenes[scId].lastsHours is not None:

                    if scn.find('LastsHours') is not None:
                        scn.find(
                            'LastsHours').text = ywProject.scenes[scId].lastsHours

                    else:
                        ET.SubElement(
                            scn, 'LastsHours').text = ywProject.scenes[scId].lastsHours

                if ywProject.scenes[scId].lastsMinutes is not None:

                    if scn.find('LastsMinutes') is not None:
                        scn.find(
                            'LastsMinutes').text = ywProject.scenes[scId].lastsMinutes

                    else:
                        ET.SubElement(
                            scn, 'LastsMinutes').text = ywProject.scenes[scId].lastsMinutes

                # Plot related information

                if ywProject.scenes[scId].isReactionScene:

                    if scn.find('ReactionScene') is None:
                        ET.SubElement(scn, 'ReactionScene').text = '-1'

                elif scn.find('ReactionScene') is not None:
                    scn.remove(scn.find('ReactionScene'))

                if ywProject.scenes[scId].isSubPlot:

                    if scn.find('SubPlot') is None:
                        ET.SubElement(scn, 'SubPlot').text = '-1'

                elif scn.find('SubPlot') is not None:
                    scn.remove(scn.find('SubPlot'))

                if ywProject.scenes[scId].goal is not None:

                    if scn.find('Goal') is None:
                        ET.SubElement(
                            scn, 'Goal').text = ywProject.scenes[scId].goal

                    else:
                        scn.find('Goal').text = ywProject.scenes[scId].goal

                if ywProject.scenes[scId].conflict is not None:

                    if scn.find('Conflict') is None:
                        ET.SubElement(
                            scn, 'Conflict').text = ywProject.scenes[scId].conflict

                    else:
                        scn.find(
                            'Conflict').text = ywProject.scenes[scId].conflict

                if ywProject.scenes[scId].outcome is not None:

                    if scn.find('Outcome') is None:
                        ET.SubElement(
                            scn, 'Outcome').text = ywProject.scenes[scId].outcome

                    else:
                        scn.find(
                            'Outcome').text = ywProject.scenes[scId].outcome

                if ywProject.scenes[scId].characters is not None:
                    characters = scn.find('Characters')

                    for oldCrId in characters.findall('CharID'):
                        characters.remove(oldCrId)

                    for crId in ywProject.scenes[scId].characters:
                        ET.SubElement(characters, 'CharID').text = crId

                if ywProject.scenes[scId].locations is not None:
                    locations = scn.find('Locations')

                    for oldLcId in locations.findall('LocID'):
                        locations.remove(oldLcId)

                    for lcId in ywProject.scenes[scId].locations:
                        ET.SubElement(locations, 'LocID').text = lcId

                if ywProject.scenes[scId].items is not None:
                    items = scn.find('Items')

                    for oldItId in items.findall('ItemID'):
                        items.remove(oldItId)

                    for itId in ywProject.scenes[scId].items:
                        ET.SubElement(items, 'ItemID').text = itId

        self.indent_xml(root)
        ywProject._tree = ET.ElementTree(root)

        return 'SUCCESS'

    def indent_xml(self, elem, level=0):
        """xml pretty printer

        Kudos to to Fredrik Lundh. 
        Source: http://effbot.org/zone/element-lib.htm#prettyprint
        """
        i = "\n" + level * "  "

        if len(elem):

            if not elem.text or not elem.text.strip():
                elem.text = i + "  "

            if not elem.tail or not elem.tail.strip():
                elem.tail = i

            for elem in elem:
                self.indent_xml(elem, level + 1)

            if not elem.tail or not elem.tail.strip():
                elem.tail = i

        else:
            if level and (not elem.tail or not elem.tail.strip()):
                elem.tail = i


class Yw7TreeBuilder(YwTreeBuilder):
    """Build yWriter 7 project xml tree."""

    def build_element_tree(self, ywProject):
        """Modify the yWriter project attributes of an existing xml element tree.
        Return a message beginning with SUCCESS or ERROR.
        """

        root = ywProject._tree.getroot()

        for scn in root.iter('SCENE'):
            scId = scn.find('ID').text

            if ywProject.scenes[scId].sceneContent is not None:
                scn.find(
                    'SceneContent').text = ywProject.scenes[scId].sceneContent
                scn.find('WordCount').text = str(
                    ywProject.scenes[scId].wordCount)
                scn.find('LetterCount').text = str(
                    ywProject.scenes[scId].letterCount)

            try:
                scn.remove(scn.find('RTFFile'))

            except:
                pass

        root.tag = 'YWRITER7'
        root.find('PROJECT').find('Ver').text = '7'
        ywProject._tree = ET.ElementTree(root)

        return YwTreeBuilder.build_element_tree(self, ywProject)

from abc import ABC
from abc import abstractmethod


class YwTreeReader(ABC):
    """Read yWriter xml project file."""

    @abstractmethod
    def read_element_tree(self, ywFile):
        """Parse the yWriter xml file located at filePath, fetching the Novel attributes.
        Return a message beginning with SUCCESS or ERROR.
        To be overridden by file format specific subclasses.
        """


class Utf8TreeReader(YwTreeReader):
    """Read utf-8 encoded yWriter xml project file."""

    def read_element_tree(self, ywFile):
        """Parse the yWriter xml file located at filePath, fetching the Novel attributes.
        Return a message beginning with SUCCESS or ERROR.
        """

        try:
            ywFile._tree = ET.parse(ywFile.filePath)

        except:
            return 'ERROR: Can not process "' + os.path.normpath(ywFile.filePath) + '".'

        return 'SUCCESS: XML element tree read in.'



class YwProjectMerger():
    """Merge the attributes of two yWriter project structures."""

    def merge_projects(self, target, source):
        """Overwrite existing target attributes with source attributes.
        Return a message beginning with SUCCESS or ERROR.
        Create target attributes, if not existing, but return ERROR.
        """

        # Merge and re-order locations.

        if source.srtLocations != []:
            target.srtLocations = source.srtLocations
            temploc = target.locations
            target.locations = {}

            for lcId in source.srtLocations:

                # Build a new target.locations dictionary sorted like the
                # source

                target.locations[lcId] = WorldElement()

                if not lcId in temploc:
                    # A new location has been added
                    temploc[lcId] = WorldElement()

                if source.locations[lcId].title:
                    # avoids deleting the title, if it is empty by accident
                    target.locations[lcId].title = source.locations[lcId].title

                else:
                    target.locations[lcId].title = temploc[lcId].title

                if source.locations[lcId].image is not None:
                    target.locations[lcId].image = source.locations[lcId].image

                else:
                    target.locations[lcId].desc = temploc[lcId].desc

                if source.locations[lcId].desc is not None:
                    target.locations[lcId].desc = source.locations[lcId].desc

                else:
                    target.locations[lcId].desc = temploc[lcId].desc

                if source.locations[lcId].aka is not None:
                    target.locations[lcId].aka = source.locations[lcId].aka

                else:
                    target.locations[lcId].aka = temploc[lcId].aka

                if source.locations[lcId].tags is not None:
                    target.locations[lcId].tags = source.locations[lcId].tags

                else:
                    target.locations[lcId].tags = temploc[lcId].tags

        # Merge and re-order items.

        if source.srtItems != []:
            target.srtItems = source.srtItems
            tempitm = target.items
            target.items = {}

            for itId in source.srtItems:

                # Build a new target.items dictionary sorted like the
                # source

                target.items[itId] = WorldElement()

                if not itId in tempitm:
                    # A new item has been added
                    tempitm[itId] = WorldElement()

                if source.items[itId].title:
                    # avoids deleting the title, if it is empty by accident
                    target.items[itId].title = source.items[itId].title

                else:
                    target.items[itId].title = tempitm[itId].title

                if source.items[itId].image is not None:
                    target.items[itId].image = source.items[itId].image

                else:
                    target.items[itId].image = tempitm[itId].image

                if source.items[itId].desc is not None:
                    target.items[itId].desc = source.items[itId].desc

                else:
                    target.items[itId].desc = tempitm[itId].desc

                if source.items[itId].aka is not None:
                    target.items[itId].aka = source.items[itId].aka

                else:
                    target.items[itId].aka = tempitm[itId].aka

                if source.items[itId].tags is not None:
                    target.items[itId].tags = source.items[itId].tags

                else:
                    target.items[itId].tags = tempitm[itId].tags

        # Merge and re-order characters.

        if source.srtCharacters != []:
            target.srtCharacters = source.srtCharacters
            tempchr = target.characters
            target.characters = {}

            for crId in source.srtCharacters:

                # Build a new target.characters dictionary sorted like the
                # source

                target.characters[crId] = Character()

                if not crId in tempchr:
                    # A new character has been added
                    tempchr[crId] = Character()

                if source.characters[crId].title:
                    # avoids deleting the title, if it is empty by accident
                    target.characters[crId].title = source.characters[crId].title

                else:
                    target.characters[crId].title = tempchr[crId].title

                if source.characters[crId].image is not None:
                    target.characters[crId].image = source.characters[crId].image

                else:
                    target.characters[crId].image = tempchr[crId].image

                if source.characters[crId].desc is not None:
                    target.characters[crId].desc = source.characters[crId].desc

                else:
                    target.characters[crId].desc = tempchr[crId].desc

                if source.characters[crId].aka is not None:
                    target.characters[crId].aka = source.characters[crId].aka

                else:
                    target.characters[crId].aka = tempchr[crId].aka

                if source.characters[crId].tags is not None:
                    target.characters[crId].tags = source.characters[crId].tags

                else:
                    target.characters[crId].tags = tempchr[crId].tags

                if source.characters[crId].notes is not None:
                    target.characters[crId].notes = source.characters[crId].notes

                else:
                    target.characters[crId].notes = tempchr[crId].notes

                if source.characters[crId].bio is not None:
                    target.characters[crId].bio = source.characters[crId].bio

                else:
                    target.characters[crId].bio = tempchr[crId].bio

                if source.characters[crId].goals is not None:
                    target.characters[crId].goals = source.characters[crId].goals

                else:
                    target.characters[crId].goals = tempchr[crId].goals

                if source.characters[crId].fullName is not None:
                    target.characters[crId].fullName = source.characters[crId].fullName

                else:
                    target.characters[crId].fullName = tempchr[crId].fullName

                if source.characters[crId].isMajor is not None:
                    target.characters[crId].isMajor = source.characters[crId].isMajor

                else:
                    target.characters[crId].isMajor = tempchr[crId].isMajor

        # Merge scenes.

        mismatchCount = 0

        for scId in source.scenes:

            if not scId in target.scenes:
                target.scenes[scId] = Scene()
                mismatchCount += 1

            if source.scenes[scId].title:
                # avoids deleting the title, if it is empty by accident
                target.scenes[scId].title = source.scenes[scId].title

            if source.scenes[scId].desc is not None:
                target.scenes[scId].desc = source.scenes[scId].desc

            if source.scenes[scId].sceneContent is not None:
                target.scenes[scId].sceneContent = source.scenes[scId].sceneContent

            if source.scenes[scId].isUnused is not None:
                target.scenes[scId].isUnused = source.scenes[scId].isUnused

            if source.scenes[scId].isNotesScene is not None:
                target.scenes[scId].isNotesScene = source.scenes[scId].isNotesScene

            if source.scenes[scId].isTodoScene is not None:
                target.scenes[scId].isTodoScene = source.scenes[scId].isTodoScene

            if source.scenes[scId].status is not None:
                target.scenes[scId].status = source.scenes[scId].status

            if source.scenes[scId].sceneNotes is not None:
                target.scenes[scId].sceneNotes = source.scenes[scId].sceneNotes

            if source.scenes[scId].tags is not None:
                target.scenes[scId].tags = source.scenes[scId].tags

            if source.scenes[scId].field1 is not None:
                target.scenes[scId].field1 = source.scenes[scId].field1

            if source.scenes[scId].field2 is not None:
                target.scenes[scId].field2 = source.scenes[scId].field2

            if source.scenes[scId].field3 is not None:
                target.scenes[scId].field3 = source.scenes[scId].field3

            if source.scenes[scId].field4 is not None:
                target.scenes[scId].field4 = source.scenes[scId].field4

            if source.scenes[scId].appendToPrev is not None:
                target.scenes[scId].appendToPrev = source.scenes[scId].appendToPrev

            if source.scenes[scId].date is not None:
                target.scenes[scId].date = source.scenes[scId].date

            if source.scenes[scId].time is not None:
                target.scenes[scId].time = source.scenes[scId].time

            if source.scenes[scId].minute is not None:
                target.scenes[scId].minute = source.scenes[scId].minute

            if source.scenes[scId].hour is not None:
                target.scenes[scId].hour = source.scenes[scId].hour

            if source.scenes[scId].day is not None:
                target.scenes[scId].day = source.scenes[scId].day

            if source.scenes[scId].lastsMinutes is not None:
                target.scenes[scId].lastsMinutes = source.scenes[scId].lastsMinutes

            if source.scenes[scId].lastsHours is not None:
                target.scenes[scId].lastsHours = source.scenes[scId].lastsHours

            if source.scenes[scId].lastsDays is not None:
                target.scenes[scId].lastsDays = source.scenes[scId].lastsDays

            if source.scenes[scId].isReactionScene is not None:
                target.scenes[scId].isReactionScene = source.scenes[scId].isReactionScene

            if source.scenes[scId].isSubPlot is not None:
                target.scenes[scId].isSubPlot = source.scenes[scId].isSubPlot

            if source.scenes[scId].goal is not None:
                target.scenes[scId].goal = source.scenes[scId].goal

            if source.scenes[scId].conflict is not None:
                target.scenes[scId].conflict = source.scenes[scId].conflict

            if source.scenes[scId].outcome is not None:
                target.scenes[scId].outcome = source.scenes[scId].outcome

            if source.scenes[scId].characters is not None:
                target.scenes[scId].characters = []

                for crId in source.scenes[scId].characters:

                    if crId in target.characters:
                        target.scenes[scId].characters.append(crId)

            if source.scenes[scId].locations is not None:
                target.scenes[scId].locations = []

                for lcId in source.scenes[scId].locations:

                    if lcId in target.locations:
                        target.scenes[scId].locations.append(lcId)

            if source.scenes[scId].items is not None:
                target.scenes[scId].items = []

                for itId in source.scenes[scId].items:

                    if itId in target.items:
                        target.scenes[scId].append(itId)

        # Merge chapters.

        scenesAssigned = []

        for chId in source.chapters:

            if not chId in target.chapters:
                target.chapters[chId] = Chapter()
                mismatchCount += 1

            if source.chapters[chId].title:
                # avoids deleting the title, if it is empty by accident
                target.chapters[chId].title = source.chapters[chId].title

            if source.chapters[chId].desc is not None:
                target.chapters[chId].desc = source.chapters[chId].desc

            if source.chapters[chId].chLevel is not None:
                target.chapters[chId].chLevel = source.chapters[chId].chLevel

            if source.chapters[chId].oldType is not None:
                target.chapters[chId].oldType = source.chapters[chId].oldType

            if source.chapters[chId].chType is not None:
                target.chapters[chId].chType = source.chapters[chId].chType

            if source.chapters[chId].isUnused is not None:
                target.chapters[chId].isUnused = source.chapters[chId].isUnused

            if source.chapters[chId].suppressChapterTitle is not None:
                target.chapters[chId].suppressChapterTitle = source.chapters[chId].suppressChapterTitle

            if source.chapters[chId].suppressChapterBreak is not None:
                target.chapters[chId].suppressChapterBreak = source.chapters[chId].suppressChapterBreak

            if source.chapters[chId].isTrash is not None:
                target.chapters[chId].isTrash = source.chapters[chId].isTrash

            if source.chapters[chId].srtScenes is not None:
                target.chapters[chId].srtScenes = []

                for scId in source.chapters[chId].srtScenes:

                    if (scId in target.scenes) and not (scId in scenesAssigned):
                        target.chapters[chId].srtScenes.append(scId)
                        scenesAssigned.append(scId)

        # Merge attributes at novel level.

        if source.title:
            # avoids deleting the title, if it is empty by accident
            target.title = source.title

        if source.desc is not None:
            target.desc = source.desc

        if source.author is not None:
            target.author = source.author

        if source.fieldTitle1 is not None:
            target.fieldTitle1 = source.fieldTitle1

        if source.fieldTitle2 is not None:
            target.fieldTitle2 = source.fieldTitle2

        if source.fieldTitle3 is not None:
            target.fieldTitle3 = source.fieldTitle3

        if source.fieldTitle4 is not None:
            target.fieldTitle4 = source.fieldTitle4

        if source.srtChapters != []:
            target.srtChapters = []

            for chId in source.srtChapters:
                target.srtChapters.append(chId)

        if mismatchCount > 0:
            return 'ERROR: Project structure mismatch.'

        else:
            return 'SUCCESS'

from abc import ABC
from abc import abstractmethod


class YwTreeWriter(ABC):
    """Write yWriter 7 xml project file."""

    @abstractmethod
    def write_element_tree(self, ywProject):
        """Write back the xml element tree to a yWriter xml file located at filePath.
        Return a message beginning with SUCCESS or ERROR.
        To be overridden by file format specific subclasses.
        """


class Utf8TreeWriter(YwTreeWriter):
    """Write utf-8 encoded yWriter project file."""

    def write_element_tree(self, ywProject):
        """Write back the xml element tree to a yWriter xml file located at filePath.
        Return a message beginning with SUCCESS or ERROR.
        """

        try:
            ywProject._tree.write(
                ywProject.filePath, xml_declaration=False, encoding='utf-8')

        except(PermissionError):
            return 'ERROR: "' + os.path.normpath(ywProject.filePath) + '" is write protected.'

        return 'SUCCESS'

from abc import ABC
from abc import abstractmethod
from html import unescape


class YwPostprocessor(ABC):

    @abstractmethod
    def postprocess_xml_file(self, ywFile):
        '''Postprocess the xml file created by ElementTree:
        Put a header on top, insert the missing CDATA tags,
        and replace xml entities by plain text.
        Return a message beginning with SUCCESS or ERROR.
        To be overridden by file format specific subclasses.
        '''

    def format_xml(self, text):
        '''Postprocess the xml file created by ElementTree:
           Insert the missing CDATA tags,
           and replace xml entities by plain text.
        '''

        cdataTags = ['Title', 'AuthorName', 'Bio', 'Desc',
                     'FieldTitle1', 'FieldTitle2', 'FieldTitle3',
                     'FieldTitle4', 'LaTeXHeaderFile', 'Tags',
                     'AKA', 'ImageFile', 'FullName', 'Goals',
                     'Notes', 'RTFFile', 'SceneContent',
                     'Outcome', 'Goal', 'Conflict']
        # Names of yWriter xml elements containing CDATA.
        # ElementTree.write omits CDATA tags, so they have to be inserted
        # afterwards.

        lines = text.split('\n')
        newlines = []

        for line in lines:

            for tag in cdataTags:
                line = re.sub('\<' + tag + '\>', '<' +
                              tag + '><![CDATA[', line)
                line = re.sub('\<\/' + tag + '\>',
                              ']]></' + tag + '>', line)

            newlines.append(line)

        text = '\n'.join(newlines)
        text = text.replace('[CDATA[ \n', '[CDATA[')
        text = text.replace('\n]]', ']]')
        text = unescape(text)

        return text


class Utf8Postprocessor(YwPostprocessor):
    """Postprocess ANSI encoded yWriter project."""

    def postprocess_xml_file(self, filePath):
        '''Postprocess the xml file created by ElementTree:
        Put a header on top, insert the missing CDATA tags,
        and replace xml entities by plain text.
        Return a message beginning with SUCCESS or ERROR.
        '''

        with open(filePath, 'r', encoding='utf-8') as f:
            text = f.read()

        text = self.format_xml(text)
        text = '<?xml version="1.0" encoding="utf-8"?>\n' + text

        try:

            with open(filePath, 'w', encoding='utf-8') as f:
                f.write(text)

        except:
            return 'ERROR: Can not write "' + os.path.normpath(filePath) + '".'

        return 'SUCCESS: "' + os.path.normpath(filePath) + '" written.'


class Yw7File(YwFile):
    """yWriter 7 project file representation."""

    DESCRIPTION = 'yWriter 7 project'
    EXTENSION = '.yw7'

    def __init__(self, filePath, **kwargs):
        YwFile.__init__(self, filePath)
        self.ywTreeReader = Utf8TreeReader()
        self.ywProjectMerger = YwProjectMerger()
        self.ywTreeBuilder = Yw7TreeBuilder()
        self.ywTreeWriter = Utf8TreeWriter()
        self.ywPostprocessor = Utf8Postprocessor()



class Yw7TreeCreator(YwTreeBuilder):
    """Create a new yWriter 7 project xml tree."""

    def build_element_tree(self, ywProject):
        """Put the yWriter project attributes to a new xml element tree.
        Return a message beginning with SUCCESS or ERROR.
        """

        root = ET.Element('YWRITER7')

        # Write attributes at novel level to the xml element tree.

        prj = ET.SubElement(root, 'PROJECT')
        ET.SubElement(prj, 'Ver').text = '7'

        if ywProject.title is not None:
            ET.SubElement(prj, 'Title').text = ywProject.title

        if ywProject.desc is not None:
            ET.SubElement(prj, 'Desc').text = ywProject.desc

        if ywProject.author is not None:
            ET.SubElement(prj, 'AuthorName').text = ywProject.author

        if ywProject.fieldTitle1 is not None:
            ET.SubElement(prj, 'FieldTitle1').text = ywProject.fieldTitle1

        if ywProject.fieldTitle2 is not None:
            ET.SubElement(prj, 'FieldTitle2').text = ywProject.fieldTitle2

        if ywProject.fieldTitle3 is not None:
            ET.SubElement(prj, 'FieldTitle3').text = ywProject.fieldTitle3

        if ywProject.fieldTitle4 is not None:
            ET.SubElement(prj, 'FieldTitle4').text = ywProject.fieldTitle4

        # Write locations to the xml element tree.

        locations = ET.SubElement(root, 'LOCATIONS')

        for lcId in ywProject.srtLocations:
            loc = ET.SubElement(locations, 'LOCATION')
            ET.SubElement(loc, 'ID').text = lcId

            if ywProject.locations[lcId].title is not None:
                ET.SubElement(
                    loc, 'Title').text = ywProject.locations[lcId].title

            if ywProject.locations[lcId].desc is not None:
                ET.SubElement(
                    loc, 'Desc').text = ywProject.locations[lcId].desc

            if ywProject.locations[lcId].aka is not None:
                ET.SubElement(loc, 'AKA').text = ywProject.locations[lcId].aka

            if ywProject.locations[lcId].tags is not None:
                ET.SubElement(loc, 'Tags').text = ';'.join(
                    ywProject.locations[lcId].tags)

        # Write items to the xml element tree.

        items = ET.SubElement(root, 'ITEMS')

        for itId in ywProject.srtItems:
            itm = ET.SubElement(items, 'ITEM')
            ET.SubElement(itm, 'ID').text = itId

            if ywProject.items[itId].title is not None:
                ET.SubElement(itm, 'Title').text = ywProject.items[itId].title

            if ywProject.items[itId].desc is not None:
                ET.SubElement(itm, 'Desc').text = ywProject.items[itId].desc

            if ywProject.items[itId].aka is not None:
                ET.SubElement(itm, 'AKA').text = ywProject.items[itId].aka

            if ywProject.items[itId].tags is not None:
                ET.SubElement(itm, 'Tags').text = ';'.join(
                    ywProject.items[itId].tags)

        # Write characters to the xml element tree.

        characters = ET.SubElement(root, 'CHARACTERS')

        for crId in ywProject.srtCharacters:
            crt = ET.SubElement(characters, 'CHARACTER')
            ET.SubElement(crt, 'ID').text = crId

            if ywProject.characters[crId].title is not None:
                ET.SubElement(
                    crt, 'Title').text = ywProject.characters[crId].title

            if ywProject.characters[crId].desc is not None:
                ET.SubElement(
                    crt, 'Desc').text = ywProject.characters[crId].desc

            if ywProject.characters[crId].aka is not None:
                ET.SubElement(crt, 'AKA').text = ywProject.characters[crId].aka

            if ywProject.characters[crId].tags is not None:
                ET.SubElement(crt, 'Tags').text = ';'.join(
                    ywProject.characters[crId].tags)

            if ywProject.characters[crId].notes is not None:
                ET.SubElement(
                    crt, 'Notes').text = ywProject.characters[crId].notes

            if ywProject.characters[crId].bio is not None:
                ET.SubElement(crt, 'Bio').text = ywProject.characters[crId].bio

            if ywProject.characters[crId].goals is not None:
                ET.SubElement(
                    crt, 'Goals').text = ywProject.characters[crId].goals

            if ywProject.characters[crId].fullName is not None:
                ET.SubElement(
                    crt, 'FullName').text = ywProject.characters[crId].fullName

            if ywProject.characters[crId].isMajor:
                ET.SubElement(crt, 'Major').text = '-1'

        # Write attributes at scene level to the xml element tree.

        scenes = ET.SubElement(root, 'SCENES')

        for scId in ywProject.scenes:
            scn = ET.SubElement(scenes, 'SCENE')
            ET.SubElement(scn, 'ID').text = scId

            if ywProject.scenes[scId].title is not None:
                ET.SubElement(scn, 'Title').text = ywProject.scenes[scId].title

            for chId in ywProject.chapters:

                if scId in ywProject.chapters[chId].srtScenes:
                    ET.SubElement(scn, 'BelongsToChID').text = chId
                    break

            if ywProject.scenes[scId].desc is not None:
                ET.SubElement(scn, 'Desc').text = ywProject.scenes[scId].desc

            if ywProject.scenes[scId].sceneContent is not None:
                ET.SubElement(scn,
                              'SceneContent').text = ywProject.scenes[scId].sceneContent
                ET.SubElement(scn, 'WordCount').text = str(
                    ywProject.scenes[scId].wordCount)
                ET.SubElement(scn, 'LetterCount').text = str(
                    ywProject.scenes[scId].letterCount)

            if ywProject.scenes[scId].isUnused:
                ET.SubElement(scn, 'Unused').text = '-1'

            scFields = ET.SubElement(scn, 'Fields')

            if ywProject.scenes[scId].isNotesScene:
                ET.SubElement(scFields, 'Field_SceneType').text = '1'

            elif ywProject.scenes[scId].isTodoScene:
                ET.SubElement(scFields, 'Field_SceneType').text = '2'

            if ywProject.scenes[scId].status is not None:
                ET.SubElement(scn, 'Status').text = str(
                    ywProject.scenes[scId].status)

            if ywProject.scenes[scId].sceneNotes is not None:
                ET.SubElement(
                    scn, 'Notes').text = ywProject.scenes[scId].sceneNotes

            if ywProject.scenes[scId].tags is not None:
                ET.SubElement(scn, 'Tags').text = ';'.join(
                    ywProject.scenes[scId].tags)

            if ywProject.scenes[scId].field1 is not None:
                ET.SubElement(
                    scn, 'Field1').text = ywProject.scenes[scId].field1

            if ywProject.scenes[scId].field2 is not None:
                ET.SubElement(
                    scn, 'Field2').text = ywProject.scenes[scId].field2

            if ywProject.scenes[scId].field3 is not None:
                ET.SubElement(
                    scn, 'Field3').text = ywProject.scenes[scId].field3

            if ywProject.scenes[scId].field4 is not None:
                ET.SubElement(
                    scn, 'Field4').text = ywProject.scenes[scId].field4

            if ywProject.scenes[scId].appendToPrev:
                ET.SubElement(scn, 'AppendToPrev').text = '-1'

            # Date/time information

            if (ywProject.scenes[scId].date is not None) and (ywProject.scenes[scId].time is not None):
                dateTime = ' '.join(
                    ywProject.scenes[scId].date, ywProject.scenes[scId].time)
                ET.SubElement(scn, 'SpecificDateTime').text = dateTime
                ET.SubElement(scn, 'SpecificDateMode').text = '-1'

            elif (ywProject.scenes[scId].day is not None) or (ywProject.scenes[scId].hour is not None) or (ywProject.scenes[scId].minute is not None):

                if ywProject.scenes[scId].day is not None:
                    ET.SubElement(scn, 'Day').text = ywProject.scenes[scId].day

                if ywProject.scenes[scId].hour is not None:
                    ET.SubElement(
                        scn, 'Hour').text = ywProject.scenes[scId].hour

                if ywProject.scenes[scId].minute is not None:
                    ET.SubElement(
                        scn, 'Minute').text = ywProject.scenes[scId].minute

            if ywProject.scenes[scId].lastsDays is not None:
                ET.SubElement(
                    scn, 'LastsDays').text = ywProject.scenes[scId].lastsDays

            if ywProject.scenes[scId].lastsHours is not None:
                ET.SubElement(
                    scn, 'LastsHours').text = ywProject.scenes[scId].lastsHours

            if ywProject.scenes[scId].lastsMinutes is not None:
                ET.SubElement(
                    scn, 'LastsMinutes').text = ywProject.scenes[scId].lastsMinutes

            # Plot related information

            if ywProject.scenes[scId].isReactionScene:
                ET.SubElement(scn, 'ReactionScene').text = '-1'

            if ywProject.scenes[scId].isSubPlot:
                ET.SubElement(scn, 'SubPlot').text = '-1'

            if ywProject.scenes[scId].goal is not None:
                ET.SubElement(scn, 'Goal').text = ywProject.scenes[scId].goal

            if ywProject.scenes[scId].conflict is not None:
                ET.SubElement(
                    scn, 'Conflict').text = ywProject.scenes[scId].conflict

            if ywProject.scenes[scId].outcome is not None:
                ET.SubElement(
                    scn, 'Outcome').text = ywProject.scenes[scId].outcome

            if ywProject.scenes[scId].characters is not None:
                scCharacters = ET.SubElement(scn, 'Characters')

                for crId in ywProject.scenes[scId].characters:
                    ET.SubElement(scCharacters, 'CharID').text = crId

            if ywProject.scenes[scId].locations is not None:
                scLocations = ET.SubElement(scn, 'Locations')

                for lcId in ywProject.scenes[scId].locations:
                    ET.SubElement(scLocations, 'LocID').text = lcId

            if ywProject.scenes[scId].items is not None:
                scItems = ET.SubElement(scn, 'Items')

                for itId in ywProject.scenes[scId].items:
                    ET.SubElement(scItems, 'ItemID').text = itId

        # Write attributes at chapter level to the xml element tree.

        chapters = ET.SubElement(root, 'CHAPTERS')

        sortOrder = 0

        for chId in ywProject.srtChapters:
            sortOrder += 1
            chp = ET.SubElement(chapters, 'CHAPTER')
            ET.SubElement(chp, 'ID').text = chId
            ET.SubElement(chp, 'SortOrder').text = str(sortOrder)

            if ywProject.chapters[chId].title is not None:
                ET.SubElement(
                    chp, 'Title').text = ywProject.chapters[chId].title

            if ywProject.chapters[chId].desc is not None:
                ET.SubElement(chp, 'Desc').text = ywProject.chapters[chId].desc

            if ywProject.chapters[chId].chLevel == 1:
                ET.SubElement(chp, 'SectionStart').text = '-1'

            if ywProject.chapters[chId].oldType is not None:
                ET.SubElement(chp, 'Type').text = str(
                    ywProject.chapters[chId].oldType)

            if ywProject.chapters[chId].chType is not None:
                ET.SubElement(chp, 'ChapterType').text = str(
                    ywProject.chapters[chId].chType)

            if ywProject.chapters[chId].isUnused:
                ET.SubElement(chp, 'Unused').text = '-1'

            sortSc = ET.SubElement(chp, 'Scenes')

            for scId in ywProject.chapters[chId].srtScenes:
                ET.SubElement(sortSc, 'ScID').text = scId

            chFields = ET.SubElement(chp, 'Fields')

            if ywProject.chapters[chId].title is not None:

                if ywProject.chapters[chId].title.startswith('@'):
                    ywProject.chapters[chId].suppressChapterTitle = True

            if ywProject.chapters[chId].suppressChapterTitle:
                ET.SubElement(
                    chFields, 'Field_SuppressChapterTitle').text = '1'

        self.indent_xml(root)
        ywProject._tree = ET.ElementTree(root)

        return 'SUCCESS'



class YwProjectCreator(YwProjectMerger):
    """Extend the super class by disabling its project structure check."""

    def merge_projects(self, target, source):
        """Create target attributes with source attributes.
        Return a message beginning with SUCCESS, even if the source and 
        target project structures are inconsistent. Thus the source
        project can be merged with an empty target, creating a new project.
        """
        YwProjectMerger.merge_projects(self, target, source)
        return 'SUCCESS'


class Yw7NewFile(Yw7File):
    """yWriter 7 new project file representation."""

    def __init__(self, filePath, **kwargs):
        """Extends the superclass constructor."""
        Yw7File.__init__(self, filePath)

        self.ywProjectMerger = YwProjectCreator()
        self.ywTreeBuilder = Yw7TreeCreator()


from html.parser import HTMLParser




def read_html_file(filePath):
    """Open a html file being encoded utf-8 or ANSI.
    Return a tuple:
    [0] = Message beginning with SUCCESS or ERROR.
    [1] = The file content in a single string. 
    """
    try:
        with open(filePath, 'r', encoding='utf-8') as f:
            text = (f.read())
    except:
        # HTML files exported by a word processor may be ANSI encoded.
        try:
            with open(filePath, 'r') as f:
                text = (f.read())

        except(FileNotFoundError):
            return ('ERROR: "' + os.path.normpath(filePath) + '" not found.', None)

    return ('SUCCESS', text)




class HtmlFile(Novel, HTMLParser):
    """Generic HTML file representation."""

    EXTENSION = '.html'

    def __init__(self, filePath, **kwargs):
        Novel.__init__(self, filePath)
        HTMLParser.__init__(self)
        self._lines = []
        self._scId = None
        self._chId = None

    def convert_to_yw(self, text):
        """Convert html tags to yWriter 6/7 raw markup. 
        Return a yw6/7 markup string.
        """

        # Clean up polluted HTML code.

        text = re.sub('</*font.*?>', '', text)
        text = re.sub('</*span.*?>', '', text)
        text = re.sub('</*FONT.*?>', '', text)
        text = re.sub('</*SPAN.*?>', '', text)

        # Put everything in one line.

        text = text.replace('\n', ' ')
        text = text.replace('\r', ' ')
        text = text.replace('\t', ' ')

        while '  ' in text:
            text = text.replace('  ', ' ').rstrip().lstrip()

        # Replace HTML tags by yWriter markup.

        text = text.replace('<i>', '[i]')
        text = text.replace('<I>', '[i]')
        text = text.replace('</i>', '[/i]')
        text = text.replace('</I>', '[/i]')
        text = text.replace('</em>', '[/i]')
        text = text.replace('</EM>', '[/i]')
        text = text.replace('<b>', '[b]')
        text = text.replace('<B>', '[b]')
        text = text.replace('</b>', '[/b]')
        text = text.replace('</B>', '[/b]')
        text = text.replace('</strong>', '[/b]')
        text = text.replace('</STRONG>', '[/b]')
        text = re.sub('<em.*?>', '[i]', text)
        text = re.sub('<EM.*?>', '[i]', text)
        text = re.sub('<strong.*?>', '[b]', text)
        text = re.sub('<STRONG.*?>', '[b]', text)

        # Remove orphaned tags.

        text = text.replace('[/b][b]', '')
        text = text.replace('[/i][i]', '')
        text = text.replace('[/b][b]', '')

        # Remove scene title annotations.

        text = re.sub('\<\!-- - .*? - -->', '', text)

        # Convert author's comments

        text = text.replace('<!--', '/*')
        text = text.replace('-->', '*/')

        return text

    def preprocess(self, text):
        """Clean up the HTML code and strip yWriter 6/7 raw markup. 
        This prevents accidentally applied formatting from being 
        transferred to the yWriter metadata. If rich text is 
        applicable, such as in scenes, overwrite this method 
        in a subclass) 
        """
        text = self.convert_to_yw(text)

        # Remove misplaced formatting tags.

        text = re.sub('\[\/*[b|i]\]', '', text)
        return text

    def postprocess(self):
        """Process the plain text after parsing.
        This is a hook for subclasses.
        """

    def handle_starttag(self, tag, attrs):
        """Identify scenes and chapters.
        Overwrites HTMLparser.handle_starttag().
        This method is applicable to HTML files that are divided into 
        chapters and scenes. For differently structured HTML files 
        overwrite this method in a subclass.
        """
        if tag == 'div':

            if attrs[0][0] == 'id':

                if attrs[0][1].startswith('ScID'):
                    self._scId = re.search('[0-9]+', attrs[0][1]).group()
                    self.scenes[self._scId] = Scene()
                    self.chapters[self._chId].srtScenes.append(self._scId)

                elif attrs[0][1].startswith('ChID'):
                    self._chId = re.search('[0-9]+', attrs[0][1]).group()
                    self.chapters[self._chId] = Chapter()
                    self.chapters[self._chId].srtScenes = []
                    self.srtChapters.append(self._chId)

    def read(self):
        """Read and parse a html file, fetching the Novel attributes.
        Return a message beginning with SUCCESS or ERROR.
        This is a template method for subclasses tailored to the 
        content of the respective HTML file.
        """
        result = read_html_file(self.filePath)

        if result[0].startswith('ERROR'):
            return (result[0])

        text = self.preprocess(result[1])
        self.feed(text)
        self.postprocess()

        return 'SUCCESS'


class HtmlImport(HtmlFile):
    """HTML 'work in progress' file representation.

    Import untagged chapters and scenes.
    """

    DESCRIPTION = 'Work in progress'
    SUFFIX = ''

    _SCENE_DIVIDER = '* * *'
    _LOW_WORDCOUNT = 10
    _COMMENT_START = '/*'
    _COMMENT_END = '*/'

    def __init__(self, filePath, **kwargs):
        HtmlFile.__init__(self, filePath)
        self._chCount = 0
        self._scCount = 0

    def preprocess(self, text):
        """Process the html text before parsing.
        """
        return self.convert_to_yw(text)

    def handle_starttag(self, tag, attrs):

        if tag in ('h1', 'h2'):
            self._scId = None
            self._lines = []
            self._chCount += 1
            self._chId = str(self._chCount)
            self.chapters[self._chId] = Chapter()
            self.chapters[self._chId].srtScenes = []
            self.srtChapters.append(self._chId)
            self.chapters[self._chId].oldType = '0'

            if tag == 'h1':
                self.chapters[self._chId].chLevel = 1

            else:
                self.chapters[self._chId].chLevel = 0

        elif tag == 'p':

            if self._scId is None and self._chId is not None:
                self._lines = []
                self._scCount += 1
                self._scId = str(self._scCount)
                self.scenes[self._scId] = Scene()
                self.chapters[self._chId].srtScenes.append(self._scId)
                self.scenes[self._scId].status = '1'
                self.scenes[self._scId].title = 'Scene ' + str(self._scCount)

        elif tag == 'div':
            self._scId = None
            self._chId = None

        elif tag == 'meta':

            if attrs[0][1].lower() == 'author':
                self.author = attrs[1][1]

            if attrs[0][1].lower() == 'description':
                self.desc = attrs[1][1]

        elif tag == 'title':
            self._lines = []

    def handle_endtag(self, tag):

        if tag == 'p':
            self._lines.append('\n')

            if self._scId is not None:
                self.scenes[self._scId].sceneContent = ''.join(self._lines)

                if self.scenes[self._scId].wordCount < self._LOW_WORDCOUNT:
                    self.scenes[self._scId].status = Scene.STATUS.index(
                        'Outline')

                else:
                    self.scenes[self._scId].status = Scene.STATUS.index(
                        'Draft')

        elif tag in ('h1', 'h2'):
            self.chapters[self._chId].title = ''.join(self._lines)
            self._lines = []

        elif tag == 'title':
            self.title = ''.join(self._lines)

    def handle_data(self, data):
        """Collect data within scene sections.
        Overwrites HTMLparser.handle_data().
        """
        if self._scId is not None and self._SCENE_DIVIDER in data:
            self._scId = None

        else:
            data = data.rstrip().lstrip()

            # Convert prefixed comment into scene title.

            if self._lines == [] and data.startswith(self._COMMENT_START):

                try:
                    scTitle, scText = data.split(
                        sep=self._COMMENT_END, maxsplit=1)
                    self.scenes[self._scId].title = scTitle.lstrip(
                        self._COMMENT_START).lstrip('- ')
                    data = scText

                except:
                    pass

            self._lines.append(data)



class HtmlOutline(HtmlFile):
    """HTML outline file representation.

    Import an outline without chapter and scene tags.
    """

    DESCRIPTION = 'Novel outline'
    SUFFIX = ''

    def __init__(self, filePath, **kwargs):
        HtmlFile.__init__(self, filePath)
        self._chCount = 0
        self._scCount = 0

    def handle_starttag(self, tag, attrs):

        if tag in ('h1', 'h2'):
            self._scId = None
            self._lines = []
            self._chCount += 1
            self._chId = str(self._chCount)
            self.chapters[self._chId] = Chapter()
            self.chapters[self._chId].srtScenes = []
            self.srtChapters.append(self._chId)
            self.chapters[self._chId].oldType = '0'

            if tag == 'h1':
                self.chapters[self._chId].chLevel = 1

            else:
                self.chapters[self._chId].chLevel = 0

        elif tag == 'h3':
            self._lines = []
            self._scCount += 1
            self._scId = str(self._scCount)
            self.scenes[self._scId] = Scene()
            self.chapters[self._chId].srtScenes.append(self._scId)
            self.scenes[self._scId].sceneContent = ''
            self.scenes[self._scId].status = Scene.STATUS.index('Outline')

        elif tag == 'div':
            self._scId = None
            self._chId = None

        elif tag == 'meta':

            if attrs[0][1].lower() == 'author':
                self.author = attrs[1][1]

            if attrs[0][1].lower() == 'description':
                self.desc = attrs[1][1]

        elif tag == 'title':
            self._lines = []

    def handle_endtag(self, tag):

        if tag == 'p':
            self._lines.append('\n')

            if self._scId is not None:
                self.scenes[self._scId].desc = ''.join(self._lines)

            elif self._chId is not None:
                self.chapters[self._chId].desc = ''.join(self._lines)

        elif tag in ('h1', 'h2'):
            self.chapters[self._chId].title = ''.join(self._lines)
            self._lines = []

        elif tag == 'h3':
            self.scenes[self._scId].title = ''.join(self._lines)
            self._lines = []

        elif tag == 'title':
            self.title = ''.join(self._lines)

    def handle_data(self, data):
        """Collect data within scene sections.
        Overwrites HTMLparser.handle_data().
        """
        self._lines.append(data.rstrip().lstrip())



class NewProjectFactory(FileFactory):
    """A factory class that instantiates a document object to read, 
    and a new yWriter project.

    Class constant:
        DO_NOT_IMPORT -- list of suffixes from file classes not meant to be imported.    
    """

    DO_NOT_IMPORT = ['_xref']

    def make_file_objects(self, sourcePath, **kwargs):
        """Instantiate a source and a target object for creation of a new yWriter project.
        Override the superclass method.

        Positional arguments:
            sourcePath -- string; path to the source file to convert.

        Return a tuple with three elements:
        - A message string starting with 'SUCCESS' or 'ERROR'
        - sourceFile: a Novel subclass instance
        - targetFile: a Novel subclass instance
        """
        if not self.canImport(sourcePath):
            return 'ERROR: This document is not meant to be written back.', None, None

        fileName, fileExtension = os.path.splitext(sourcePath)
        targetFile = Yw7NewFile(fileName + Yw7NewFile.EXTENSION, **kwargs)

        if sourcePath.endswith('.html'):

            # The source file might be an outline or a "work in progress".

            result = read_html_file(sourcePath)

            if result[0].startswith('SUCCESS'):

                if "<h3" in result[1].lower():
                    sourceFile = HtmlOutline(sourcePath, **kwargs)

                else:
                    sourceFile = HtmlImport(sourcePath, **kwargs)

                return 'SUCCESS', sourceFile, targetFile

            else:
                return 'ERROR: Cannot read "' + os.path.normpath(sourcePath) + '".', None, None

        else:
            for fileClass in self.fileClasses:

                if fileClass.SUFFIX is not None:

                    if sourcePath.endswith(fileClass.SUFFIX + fileClass.EXTENSION):
                        sourceFile = fileClass(sourcePath, **kwargs)
                        return 'SUCCESS', sourceFile, targetFile

            return 'ERROR: File type of  "' + os.path.normpath(sourcePath) + '" not supported.', None, None

    def canImport(self, sourcePath):
        """Return True, if the file located at sourcepath is of an importable type.
        Otherwise, return False.
        """
        fileName, fileExtension = os.path.splitext(sourcePath)

        for suffix in self.DO_NOT_IMPORT:

            if fileName.endswith(suffix):
                return False

        return True





class Yw6TreeBuilder(YwTreeBuilder):
    """Build yWriter 6 project xml tree."""

    def build_element_tree(self, ywProject):
        """Modify the yWriter project attributes of an existing xml element tree.
        Return a message beginning with SUCCESS or ERROR.
        """

        root = ywProject._tree.getroot()

        for scn in root.iter('SCENE'):
            scId = scn.find('ID').text

            if ywProject.scenes[scId].sceneContent is not None:
                scn.find(
                    'SceneContent').text = ywProject.scenes[scId].sceneContent
                scn.find('WordCount').text = str(
                    ywProject.scenes[scId].wordCount)
                scn.find('LetterCount').text = str(
                    ywProject.scenes[scId].letterCount)

        root.tag = 'YWRITER6'
        root.find('PROJECT').find('Ver').text = '5'
        ywProject._tree = ET.ElementTree(root)

        return YwTreeBuilder.build_element_tree(self, ywProject)


class Yw6File(YwFile):
    """yWriter 6 project file representation."""

    DESCRIPTION = 'yWriter 6 project'
    EXTENSION = '.yw6'

    def __init__(self, filePath, **kwargs):
        YwFile.__init__(self, filePath)
        self.ywTreeReader = Utf8TreeReader()
        self.ywProjectMerger = YwProjectMerger()
        self.ywTreeBuilder = Yw6TreeBuilder()
        self.ywTreeWriter = Utf8TreeWriter()
        self.ywPostprocessor = Utf8Postprocessor()





class Yw5TreeBuilder(YwTreeBuilder):
    """Build yWriter 5 project xml tree."""

    def convert_to_rtf(self, text):
        """Convert yw6/7 raw markup to rtf. 
        Return a rtf encoded string.
        """

        RTF_HEADER = '{\\rtf1\\ansi\\deff0\\nouicompat{\\fonttbl{\\f0\\fnil\\fcharset0 Courier New;}}{\\*\\generator PyWriter}\\viewkind4\\uc1 \\pard\\sa0\\sl240\\slmult1\\f0\\fs24\\lang9 '
        RTF_FOOTER = ' }'

        RTF_REPLACEMENTS = [
            ['\n\n', '\\line\\par '],
            ['\n', '\\par '],
            ['[i]', '{\\i '],
            ['[/i]', '}'],
            ['[b]', '{\\b '],
            ['[/b]', '}'],
            ['–', '--'],
            ['—', '--'],
            ['„', '\\u8222?'],
            ['‚', '\\u8218?'],
            ['‘', '\\lquote '],
            ['’', '\\rquote '],
            ['“', '\\ldblquote '],
            ['”', '\\rdblquote '],
            ['\u202f', '\\~'],
            ['»', '\\u0187?'],
            ['«', '\\u0171?'],
            ['›', '\\u8250?'],
            ['‹', '\\u8249?'],
            ['…', '\\u8230?'],
        ]

        try:

            for r in RTF_REPLACEMENTS:
                text = text.replace(r[0], r[1])

        except AttributeError:
            text = ''

        return RTF_HEADER + text + RTF_FOOTER

    def build_element_tree(self, ywProject):
        """Modify the yWriter project attributes of an existing xml element tree.
        Write scene contents to RTF files.
        Return a message beginning with SUCCESS or ERROR.
        """
        rtfDir = os.path.dirname(ywProject.filePath)

        if rtfDir == '':
            rtfDir = './RTF5'

        else:
            rtfDir += '/RTF5'

        for chId in ywProject.chapters:

            if ywProject.chapters[chId].oldType == 1:
                ywProject.chapters[chId].isUnused = False

        root = ywProject._tree.getroot()

        for scn in root.iter('SCENE'):
            scId = scn.find('ID').text

            try:
                scn.remove(scn.find('SceneContent'))

            except:
                pass

            if ywProject.scenes[scId].rtfFile is not None:

                if scn.find('RTFFile') is None:
                    ET.SubElement(
                        scn, 'RTFFile').text = ywProject.scenes[scId].rtfFile

                if ywProject.scenes[scId].sceneContent is not None:
                    rtfPath = rtfDir + '/' + ywProject.scenes[scId].rtfFile

                    rtfScene = self.convert_to_rtf(
                        ywProject.scenes[scId].sceneContent)

                    try:

                        with open(rtfPath, 'w') as f:
                            f.write(rtfScene)

                    except:

                        return 'ERROR: Can not write scene file "' + rtfPath + '".'

        root.tag = 'YWRITER5'
        root.find('PROJECT').find('Ver').text = '5'
        ywProject._tree = ET.ElementTree(root)

        return YwTreeBuilder.build_element_tree(self, ywProject)



class AnsiTreeReader(YwTreeReader):
    """Read ANSI encoded yWriter xml project file."""

    def read_element_tree(self, ywFile):
        """Parse the yWriter xml file located at filePath, fetching the Novel attributes.
        Return a message beginning with SUCCESS or ERROR.
        """

        _TEMPFILE = '._tempfile.xml'

        try:

            with open(ywFile.filePath, 'r') as f:
                project = f.readlines()

            project[0] = project[0].replace('<?xml version="1.0" encoding="iso-8859-1"?>',
                                            '<?xml version="1.0" encoding="cp1252"?>')

            with open(_TEMPFILE, 'w') as f:
                f.writelines(project)

            ywFile._tree = ET.parse(_TEMPFILE)
            os.remove(_TEMPFILE)

        except:
            return 'ERROR: Can not process "' + os.path.normpath(ywFile.filePath) + '".'

        return 'SUCCESS: XML element tree read in.'



class AnsiTreeWriter(YwTreeWriter):
    """Write ANSI encoded yWriter project file."""

    def write_element_tree(self, ywProject):
        """Write back the xml element tree to a yWriter xml file located at filePath.
        Return a message beginning with SUCCESS or ERROR.
        """

        try:
            ywProject._tree.write(
                ywProject.filePath, xml_declaration=False, encoding='iso-8859-1')

        except(PermissionError):
            return 'ERROR: "' + os.path.normpath(ywProject.filePath) + '" is write protected.'

        return 'SUCCESS'



class AnsiPostprocessor(YwPostprocessor):
    """Postprocess ANSI encoded yWriter project."""

    def postprocess_xml_file(self, filePath):
        '''Postprocess the xml file created by ElementTree:
        Put a header on top, insert the missing CDATA tags,
        and replace xml entities by plain text.
        Return a message beginning with SUCCESS or ERROR.
        '''

        with open(filePath, 'r') as f:
            text = f.read()

        text = self.format_xml(text)
        text = '<?xml version="1.0" encoding="iso-8859-1"?>\n' + text

        try:

            with open(filePath, 'w') as f:
                f.write(text)

        except:
            return 'ERROR: Can not write "' + os.path.normpath(filePath) + '".'

        return 'SUCCESS: "' + os.path.normpath(filePath) + '" written.'


class Yw5File(YwFile):
    """yWriter 5 project file representation."""

    DESCRIPTION = 'yWriter 5 project'
    EXTENSION = '.yw5'

    def __init__(self, filePath, **kwargs):
        """Extends the super class constructor."""
        YwFile.__init__(self, filePath)

        self.ywTreeReader = AnsiTreeReader()
        self.ywProjectMerger = YwProjectMerger()
        self.ywTreeBuilder = Yw5TreeBuilder()
        self.ywTreeWriter = AnsiTreeWriter()
        self.ywPostprocessor = AnsiPostprocessor()



class OdtExport(OdtFile):
    """ODT novel file representation.

    Export a non-reimportable manuscript with chapters and scenes.
    """
    fileHeader = OdtFile.CONTENT_XML_HEADER + '''<text:p text:style-name="Title">$Title</text:p>
<text:p text:style-name="Subtitle">$AuthorName</text:p>
'''

    partTemplate = '''<text:h text:style-name="Heading_20_1" text:outline-level="1">$Title</text:h>
'''

    chapterTemplate = '''<text:h text:style-name="Heading_20_2" text:outline-level="2">$Title</text:h>
'''

    sceneTemplate = '''<text:p text:style-name="Text_20_body"><office:annotation>
<dc:creator>scene title</dc:creator>
<text:p>- $Title</text:p>
</office:annotation>$SceneContent</text:p>
'''

    appendedSceneTemplate = '''<text:p text:style-name="First_20_line_20_indent"><office:annotation>
<dc:creator>scene title</dc:creator>
<text:p>- $Title</text:p>
</office:annotation>$SceneContent</text:p>
'''

    sceneDivider = '<text:p ></text:p>'
    #sceneDivider = '<text:p text:style-name="Heading_20_4">* * *</text:p>'

    fileFooter = OdtFile.CONTENT_XML_FOOTER

    def get_chapterMapping(self, chId, chapterNumber):
        """Return a mapping dictionary for a chapter section. 
        """
        chapterMapping = OdtFile.get_chapterMapping(self, chId, chapterNumber)

        if self.chapters[chId].suppressChapterTitle:
            chapterMapping['Title'] = ''

        return chapterMapping






class HtmlProof(HtmlFile):
    """HTML proof reading file representation.

    Import a manuscript with visibly tagged chapters and scenes.
    """

    DESCRIPTION = 'Tagged manuscript for proofing'
    SUFFIX = '_proof'

    def __init__(self, filePath, **kwargs):
        HtmlFile.__init__(self, filePath)
        self._collectText = False

    def preprocess(self, text):
        """Process the html text before parsing.
        """
        return self.convert_to_yw(text)

    def postprocess(self):
        """Parse the converted text to identify chapters and scenes.
        """
        sceneText = []
        scId = ''
        chId = ''
        inScene = False

        for line in self._lines:

            if '[ScID' in line:
                scId = re.search('[0-9]+', line).group()
                self.scenes[scId] = Scene()
                self.chapters[chId].srtScenes.append(scId)
                inScene = True

            elif '[/ScID' in line:
                self.scenes[scId].sceneContent = '\n'.join(sceneText)
                sceneText = []
                inScene = False

            elif '[ChID' in line:
                chId = re.search('[0-9]+', line).group()
                self.chapters[chId] = Chapter()
                self.srtChapters.append(chId)

            elif '[/ChID' in line:
                pass

            elif inScene:
                sceneText.append(line)

    def handle_starttag(self, tag, attrs):
        """Recognize the paragraph's beginning.
        Overwrites HTMLparser.handle_endtag().
        """
        if tag == 'p':
            self._collectText = True

    def handle_endtag(self, tag):
        """Recognize the paragraph's end.
        Overwrites HTMLparser.handle_endtag().
        """
        if tag == 'p':
            self._collectText = False

    def handle_data(self, data):
        """Copy the scene paragraphs.
        Overwrites HTMLparser.handle_data().
        """
        if self._collectText:
            self._lines.append(data)



class HtmlManuscript(HtmlFile):
    """HTML manuscript file representation.

    Import a manuscript with invisibly tagged chapters and scenes.
    """

    DESCRIPTION = 'Editable manuscript'
    SUFFIX = '_manuscript'

    def preprocess(self, text):
        """Process the html text before parsing.
        """
        return self.convert_to_yw(text)

    def handle_endtag(self, tag):
        """Recognize the end of the scene section and save data.
        Overwrites HTMLparser.handle_endtag().
        """
        if self._scId is not None:

            if tag == 'div':
                self.scenes[self._scId].sceneContent = ''.join(self._lines)
                self._lines = []
                self._scId = None

            elif tag == 'p':
                self._lines.append('\n')

        elif self._chId is not None:

            if tag == 'div':
                self._chId = None

    def handle_data(self, data):
        """Collect data within scene sections.
        Overwrites HTMLparser.handle_data().
        """
        if self._scId is not None:
            self._lines.append(data.rstrip().lstrip())



class HtmlSceneDesc(HtmlFile):
    """HTML scene summaries file representation.

    Import a full synopsis with invisibly tagged scene descriptions.
    """

    DESCRIPTION = 'Scene descriptions'
    SUFFIX = '_scenes'

    def handle_endtag(self, tag):
        """Recognize the end of the scene section and save data.
        Overwrites HTMLparser.handle_endtag().
        """
        if self._scId is not None:

            if tag == 'div':
                self.scenes[self._scId].desc = ''.join(self._lines)
                self._lines = []
                self._scId = None

            elif tag == 'p':
                self._lines.append('\n')

        elif self._chId is not None:

            if tag == 'div':
                self._chId = None

    def handle_data(self, data):
        """Collect data within scene sections.
        Overwrites HTMLparser.handle_data().
        """
        if self._scId is not None:
            self._lines.append(data.rstrip().lstrip())



class HtmlChapterDesc(HtmlFile):
    """HTML chapter summaries file representation.

    Import a brief synopsis with invisibly tagged chapter descriptions.
    """

    DESCRIPTION = 'Chapter descriptions'
    SUFFIX = '_chapters'

    def handle_endtag(self, tag):
        """Recognize the end of the chapter section and save data.
        Overwrites HTMLparser.handle_endtag().
        """
        if self._chId is not None:

            if tag == 'div':
                self.chapters[self._chId].desc = ''.join(self._lines)
                self._lines = []
                self._chId = None

            elif tag == 'p':
                self._lines.append('\n')

    def handle_data(self, data):
        """collect data within chapter sections.
        Overwrites HTMLparser.handle_data().
        """
        if self._chId is not None:
            self._lines.append(data.rstrip().lstrip())



class HtmlPartDesc(HtmlChapterDesc):
    """HTML part summaries file representation.

    Import a very brief synopsis with invisibly tagged part descriptions.
    """

    DESCRIPTION = 'Part descriptions'
    SUFFIX = '_parts'




class HtmlCharacters(HtmlFile):
    """HTML character descriptions file representation.

    Import a character sheet with invisibly tagged descriptions.
    """

    DESCRIPTION = 'Character descriptions'
    SUFFIX = '_characters'

    def __init__(self, filePath, **kwargs):
        HtmlFile.__init__(self, filePath)
        self._crId = None
        self._section = None

    def handle_starttag(self, tag, attrs):
        """Identify characters with subsections.
        Overwrites HTMLparser.handle_starttag()
        """
        if tag == 'div':

            if attrs[0][0] == 'id':

                if attrs[0][1].startswith('CrID_desc'):
                    self._crId = re.search('[0-9]+', attrs[0][1]).group()
                    self.srtCharacters.append(self._crId)
                    self.characters[self._crId] = Character()
                    self._section = 'desc'

                elif attrs[0][1].startswith('CrID_bio'):
                    self._section = 'bio'

                elif attrs[0][1].startswith('CrID_goals'):
                    self._section = 'goals'

    def handle_endtag(self, tag):
        """Recognize the end of the character section and save data.
        Overwrites HTMLparser.handle_endtag().
        """
        if self._crId is not None:

            if tag == 'div':

                if self._section == 'desc':
                    self.characters[self._crId].desc = ''.join(self._lines)
                    self._lines = []
                    self._section = None

                elif self._section == 'bio':
                    self.characters[self._crId].bio = ''.join(self._lines)
                    self._lines = []
                    self._section = None

                elif self._section == 'goals':
                    self.characters[self._crId].goals = ''.join(self._lines)
                    self._lines = []
                    self._section = None

            elif tag == 'p':
                self._lines.append('\n')

    def handle_data(self, data):
        """collect data within character sections.
        Overwrites HTMLparser.handle_data().
        """
        if self._section is not None:
            self._lines.append(data.rstrip().lstrip())




class HtmlLocations(HtmlFile):
    """HTML location descriptions file representation.

    Import a location sheet with invisibly tagged descriptions.
    """

    DESCRIPTION = 'Location descriptions'
    SUFFIX = '_locations'

    def __init__(self, filePath, **kwargs):
        HtmlFile.__init__(self, filePath)
        self._lcId = None

    def handle_starttag(self, tag, attrs):
        """Identify locations.
        Overwrites HTMLparser.handle_starttag()
        """
        if tag == 'div':

            if attrs[0][0] == 'id':

                if attrs[0][1].startswith('LcID'):
                    self._lcId = re.search('[0-9]+', attrs[0][1]).group()
                    self.srtLocations.append(self._lcId)
                    self.locations[self._lcId] = WorldElement()

    def handle_endtag(self, tag):
        """Recognize the end of the location section and save data.
        Overwrites HTMLparser.handle_endtag().
        """
        if self._lcId is not None:

            if tag == 'div':
                self.locations[self._lcId].desc = ''.join(self._lines)
                self._lines = []
                self._lcId = None

            elif tag == 'p':
                self._lines.append('\n')

    def handle_data(self, data):
        """collect data within location sections.
        Overwrites HTMLparser.handle_data().
        """
        if self._lcId is not None:
            self._lines.append(data.rstrip().lstrip())



class HtmlItems(HtmlFile):
    """HTML item descriptions file representation.

    Import a item sheet with invisibly tagged descriptions.
    """

    DESCRIPTION = 'Item descriptions'
    SUFFIX = '_items'

    def __init__(self, filePath, **kwargs):
        HtmlFile.__init__(self, filePath)
        self._itId = None

    def handle_starttag(self, tag, attrs):
        """Identify items.
        Overwrites HTMLparser.handle_starttag()
        """
        if tag == 'div':

            if attrs[0][0] == 'id':

                if attrs[0][1].startswith('ItID'):
                    self._itId = re.search('[0-9]+', attrs[0][1]).group()
                    self.srtItems.append(self._itId)
                    self.items[self._itId] = WorldElement()

    def handle_endtag(self, tag):
        """Recognize the end of the item section and save data.
        Overwrites HTMLparser.handle_endtag().
        """
        if self._itId is not None:

            if tag == 'div':
                self.items[self._itId].desc = ''.join(self._lines)
                self._lines = []
                self._itId = None

            elif tag == 'p':
                self._lines.append('\n')

    def handle_data(self, data):
        """collect data within item sections.
        Overwrites HTMLparser.handle_data().
        """
        if self._itId is not None:
            self._lines.append(data.rstrip().lstrip())


import csv



class CsvFile(FileExport):
    """csv file representation.

    - Records are separated by line breaks.
    - Data fields are delimited by the _SEPARATOR character.
    """

    EXTENSION = '.csv'
    # overwrites Novel.EXTENSION

    _SEPARATOR = '|'
    # delimits data fields within a record.

    CSV_REPLACEMENTS = []

    def convert_from_yw(self, text):
        """Convert line breaks."""

        try:
            text = text.rstrip()

            for r in self.CSV_REPLACEMENTS:
                text = text.replace(r[0], r[1])

        except AttributeError:
            text = ''

        return text

    def convert_to_yw(self, text):
        """Convert line breaks."""

        try:

            for r in self.CSV_REPLACEMENTS:
                text = text.replace(r[1], r[0])

        except AttributeError:
            text = ''

        return text

    def read(self):
        """Parse the csv file located at filePath, fetching the rows.
        Check the number of fields in each row.
        Return a message beginning with SUCCESS or ERROR.
        """
        self.rows = []
        cellsPerRow = len(self.fileHeader.split(self._SEPARATOR))

        try:
            with open(self.filePath, newline='', encoding='utf-8') as f:
                reader = csv.reader(f, delimiter=self._SEPARATOR)

                for row in reader:
                    # Each row read from the csv file is returned
                    # as a list of strings

                    if len(row) != cellsPerRow:
                        return 'ERROR: Wrong csv structure.'

                    self.rows.append(row)

        except(FileNotFoundError):
            return 'ERROR: "' + os.path.normpath(self.filePath) + '" not found.'

        except:
            return 'ERROR: Can not parse "' + os.path.normpath(self.filePath) + '".'

        return 'SUCCESS'


class CsvSceneList(CsvFile):
    """csv file representation of an yWriter project's scenes table. 

    Represents a csv file with a record per scene.
    - Records are separated by line breaks.
    - Data fields are delimited by the _SEPARATOR character.
    """

    DESCRIPTION = 'Scene list'
    SUFFIX = '_scenelist'

    _SCENE_RATINGS = ['2', '3', '4', '5', '6', '7', '8', '9', '10']
    # '1' is assigned N/A (empty table cell).

    fileHeader = '''Scene link|''' +\
        '''Scene title|Scene description|Tags|Scene notes|''' +\
        '''A/R|Goal|Conflict|Outcome|''' +\
        '''Scene|Words total|$FieldTitle1|$FieldTitle2|$FieldTitle3|$FieldTitle4|''' +\
        '''Word count|Letter count|Status|''' +\
        '''Characters|Locations|Items
'''

    sceneTemplate = '''=HYPERLINK("file:///$ProjectPath/${ProjectName}_manuscript.odt#ScID:$ID%7Cregion";"ScID:$ID")|''' +\
        '''$Title|"$Desc"|$Tags|"$Notes"|''' +\
        '''$ReactionScene|"$Goal"|"$Conflict"|"$Outcome"|''' +\
        '''$SceneNumber|$WordsTotal|$Field1|$Field2|$Field3|$Field4|''' +\
        '''$WordCount|$LetterCount|$Status|''' +\
        '''$Characters|$Locations|$Items
'''

    def get_sceneMapping(self, scId, sceneNumber, wordsTotal, lettersTotal):
        """Return a mapping dictionary for a scene section. 
        """
        sceneMapping = CsvFile.get_sceneMapping(
            self, scId, sceneNumber, wordsTotal, lettersTotal)

        if self.scenes[scId].field1 == '1':
            sceneMapping['Field1'] = ''

        if self.scenes[scId].field2 == '1':
            sceneMapping['Field2'] = ''

        if self.scenes[scId].field3 == '1':
            sceneMapping['Field3'] = ''

        if self.scenes[scId].field4 == '1':
            sceneMapping['Field4'] = ''

        return sceneMapping

    def read(self):
        """Parse the csv file located at filePath, 
        fetching the Scene attributes contained.
        Return a message beginning with SUCCESS or ERROR.
        """
        message = CsvFile.read(self)

        if message.startswith('ERROR'):
            return message

        for cells in self.rows:
            i = 0

            if 'ScID:' in cells[i]:
                scId = re.search('ScID\:([0-9]+)', cells[0]).group(1)
                self.scenes[scId] = Scene()
                i += 1
                self.scenes[scId].title = cells[i]
                i += 1
                self.scenes[scId].desc = self.convert_to_yw(cells[i])
                i += 1
                self.scenes[scId].tags = self.get_list(cells[i])
                i += 1
                self.scenes[scId].sceneNotes = self.convert_to_yw(cells[i])
                i += 1

                if Scene.REACTION_MARKER.lower() in cells[i].lower():
                    self.scenes[scId].isReactionScene = True

                else:
                    self.scenes[scId].isReactionScene = False

                i += 1
                self.scenes[scId].goal = cells[i]
                i += 1
                self.scenes[scId].conflict = cells[i]
                i += 1
                self.scenes[scId].outcome = cells[i]
                i += 1
                # Don't write back sceneCount
                i += 1
                # Don't write back wordCount
                i += 1

                # Transfer scene ratings; set to 1 if deleted

                if cells[i] in self._SCENE_RATINGS:
                    self.scenes[scId].field1 = cells[i]

                else:
                    self.scenes[scId].field1 = '1'

                i += 1

                if cells[i] in self._SCENE_RATINGS:
                    self.scenes[scId].field2 = cells[i]

                else:
                    self.scenes[scId].field2 = '1'

                i += 1

                if cells[i] in self._SCENE_RATINGS:
                    self.scenes[scId].field3 = cells[i]

                else:
                    self.scenes[scId].field3 = '1'

                i += 1

                if cells[i] in self._SCENE_RATINGS:
                    self.scenes[scId].field4 = cells[i]

                else:
                    self.scenes[scId].field4 = '1'

                i += 1
                # Don't write back scene words total
                i += 1
                # Don't write back scene letters total
                i += 1

                try:
                    self.scenes[scId].status = Scene.STATUS.index(cells[i])

                except ValueError:
                    pass
                    # Scene status remains None and will be ignored when
                    # writing back.

                i += 1
                ''' Cannot write back character IDs, because self.characters is None
                charaNames = self.get_list(cells[i])
                self.scenes[scId].characters = []

                for charaName in charaNames:

                    for id, name in self.characters.items():

                        if name == charaName:
                            self.scenes[scId].characters.append(id)
                '''
                i += 1
                ''' Cannot write back location IDs, because self.locations is None
                locaNames = self.get_list(cells[i])
                self.scenes[scId].locations = []

                for locaName in locaNames:

                    for id, name in self.locations.items():

                        if name == locaName:
                            self.scenes[scId].locations.append(id)
                '''
                i += 1
                ''' Cannot write back item IDs, because self.items is None
                itemNames = self.get_list(cells[i])
                self.scenes[scId].items = []

                for itemName in itemNames:

                    for id, name in self.items.items():

                        if name == itemName:
                            self.scenes[scId].items.append(id)
                '''

        return 'SUCCESS: Data read from "' + os.path.normpath(self.filePath) + '".'




class CsvPlotList(CsvFile):
    """csv file representation of an yWriter project's scenes table. 

    Represents a csv file with a record per scene.
    - Records are separated by line breaks.
    - Data fields are delimited by the _SEPARATOR character.
    """

    DESCRIPTION = 'Plot list'
    SUFFIX = '_plotlist'

    _SEPARATOR = '|'     # delimits data fields within a record.
    _LINEBREAK = '\t'    # substitutes embedded line breaks.

    _STORYLINE_MARKER = 'story'
    # Field names containing this string (case insensitive)
    # are associated to storylines

    _SCENE_RATINGS = ['2', '3', '4', '5', '6', '7', '8', '9', '10']
    # '1' is assigned N/A (empty table cell).

    _NOT_APPLICABLE = 'N/A'
    # Scene field column header for fields not being assigned to a storyline

    _CHAR_STATE = ['', 'N/A', 'unhappy', 'dissatisfied',
                   'vague', 'satisfied', 'happy', '', '', '', '']

    fileHeader = '''ID|''' +\
        '''Plot section|Plot event|Plot event title|Details|''' +\
        '''Scene|Words total|$FieldTitle1|$FieldTitle2|$FieldTitle3|$FieldTitle4
'''

    notesChapterTemplate = '''ChID:$ID|$Title|||$Desc||||||
'''

    sceneTemplate = '''=HYPERLINK("file:///$ProjectPath/${ProjectName}_manuscript.odt#ScID:$ID%7Cregion";"ScID:$ID")|''' +\
        '''|$Tags|$Title|"$Notes"|''' +\
        '''$SceneNumber|$WordsTotal|$Field1|$Field2|$Field3|$Field4
'''

    def get_fileHeaderMapping(self):
        """Return a mapping dictionary for the project section. 
        """
        projectTemplateMapping = CsvFile.get_fileHeaderMapping(self)

        charList = []

        for crId in self.srtCharacters:
            charList.append(self.characters[crId].title)
            # Collect character names to identify storylines

        if self.fieldTitle1 in charList or self._STORYLINE_MARKER in self.fieldTitle1.lower():
            self.arc1 = True

        else:
            self.arc1 = False
            projectTemplateMapping['FieldTitle1'] = self._NOT_APPLICABLE

        if self.fieldTitle2 in charList or self._STORYLINE_MARKER in self.fieldTitle2.lower():
            self.arc2 = True

        else:
            self.arc2 = False
            projectTemplateMapping['FieldTitle2'] = self._NOT_APPLICABLE

        if self.fieldTitle3 in charList or self._STORYLINE_MARKER in self.fieldTitle3.lower():
            self.arc3 = True

        else:
            self.arc3 = False
            projectTemplateMapping['FieldTitle3'] = self._NOT_APPLICABLE

        if self.fieldTitle4 in charList or self._STORYLINE_MARKER in self.fieldTitle4.lower():
            self.arc4 = True

        else:
            self.arc4 = False
            projectTemplateMapping['FieldTitle4'] = self._NOT_APPLICABLE

        return projectTemplateMapping

    def get_sceneMapping(self, scId, sceneNumber, wordsTotal, lettersTotal):
        """Return a mapping dictionary for a scene section. 
        """
        sceneMapping = CsvFile.get_sceneMapping(
            self, scId, sceneNumber, wordsTotal, lettersTotal)

        if self.scenes[scId].field1 == '1' or not self.arc1:
            sceneMapping['Field1'] = ''

        if self.scenes[scId].field2 == '1' or not self.arc2:
            sceneMapping['Field2'] = ''

        if self.scenes[scId].field3 == '1' or not self.arc3:
            sceneMapping['Field3'] = ''

        if self.scenes[scId].field4 == '1' or not self.arc4:
            sceneMapping['Field4'] = ''

        return sceneMapping

    def read(self):
        """Parse the csv file located at filePath, fetching 
        the Scene attributes contained.
        Return a message beginning with SUCCESS or ERROR.
        """
        message = CsvFile.read(self)

        if message.startswith('ERROR'):
            return message

        tableHeader = self.rows[0]

        for cells in self.rows:

            if 'ChID:' in cells[0]:
                chId = re.search('ChID\:([0-9]+)', cells[0]).group(1)
                self.chapters[chId] = Chapter()
                self.chapters[chId].title = cells[1]
                self.chapters[chId].desc = self.convert_to_yw(cells[4])

            if 'ScID:' in cells[0]:
                scId = re.search('ScID\:([0-9]+)', cells[0]).group(1)
                self.scenes[scId] = Scene()
                self.scenes[scId].tags = self.get_list(cells[2])
                self.scenes[scId].title = cells[3]
                self.scenes[scId].sceneNotes = self.convert_to_yw(cells[4])

                i = 5
                # Don't write back sceneCount
                i += 1
                # Don't write back wordCount
                i += 1

                # Transfer scene ratings; set to 1 if deleted

                if cells[i] in self._SCENE_RATINGS:
                    self.scenes[scId].field1 = cells[i]

                elif tableHeader[i] != self._NOT_APPLICABLE:
                    self.scenes[scId].field1 = '1'

                i += 1

                if cells[i] in self._SCENE_RATINGS:
                    self.scenes[scId].field2 = cells[i]

                elif tableHeader[i] != self._NOT_APPLICABLE:
                    self.scenes[scId].field2 = '1'

                i += 1

                if cells[i] in self._SCENE_RATINGS:
                    self.scenes[scId].field3 = cells[i]

                elif tableHeader[i] != self._NOT_APPLICABLE:
                    self.scenes[scId].field3 = '1'

                i += 1

                if cells[i] in self._SCENE_RATINGS:
                    self.scenes[scId].field4 = cells[i]

                elif tableHeader[i] != self._NOT_APPLICABLE:
                    self.scenes[scId].field4 = '1'

        return 'SUCCESS: Data read from "' + os.path.normpath(self.filePath) + '".'




class CsvCharList(CsvFile):
    """csv file representation of an yWriter project's characters table. 

    Represents a csv file with a record per character.
    - Records are separated by line breaks.
    - Data fields are delimited by the _SEPARATOR character.
    """

    DESCRIPTION = 'Character list'
    SUFFIX = '_charlist'

    fileHeader = '''ID|Name|Full name|Aka|Description|Bio|Goals|Importance|Tags|Notes
'''

    characterTemplate = '''CrID:$ID|$Title|$FullName|$AKA|"$Desc"|"$Bio"|"$Goals"|$Status|$Tags|"$Notes"
'''

    def read(self):
        """Parse the csv file located at filePath, 
        fetching the Character attributes contained.
        Return a message beginning with SUCCESS or ERROR.
        """
        message = CsvFile.read(self)

        if message.startswith('ERROR'):
            return message

        for cells in self.rows:

            if 'CrID:' in cells[0]:
                crId = re.search('CrID\:([0-9]+)', cells[0]).group(1)
                self.srtCharacters.append(crId)
                self.characters[crId] = Character()
                self.characters[crId].title = cells[1]
                self.characters[crId].fullName = cells[2]
                self.characters[crId].aka = cells[3]
                self.characters[crId].desc = self.convert_to_yw(cells[4])
                self.characters[crId].bio = cells[5]
                self.characters[crId].goals = cells[6]

                if Character.MAJOR_MARKER in cells[7]:
                    self.characters[crId].isMajor = True

                else:
                    self.characters[crId].isMajor = False

                self.characters[crId].tags = self.get_list(cells[8])
                self.characters[crId].notes = self.convert_to_yw(cells[9])

        return 'SUCCESS: Data read from "' + os.path.normpath(self.filePath) + '".'

    def merge(self, novel):
        """Copy required attributes of the novel object.
        Return a message beginning with SUCCESS or ERROR.
        """
        self.srtCharacters = novel.srtCharacters
        self.characters = novel.characters
        return 'SUCCESS'




class CsvLocList(CsvFile):
    """csv file representation of an yWriter project's locations table. 

    Represents a csv file with a record per location.
    - Records are separated by line breaks.
    - Data fields are delimited by the _SEPARATOR location.
    """

    DESCRIPTION = 'Location list'
    SUFFIX = '_loclist'

    fileHeader = '''ID|Name|Description|Aka|Tags
'''

    locationTemplate = '''LcID:$ID|$Title|"$Desc"|$AKA|$Tags
'''

    def read(self):
        """Parse the csv file located at filePath, 
        fetching the WorldElement attributes contained.
        Return a message beginning with SUCCESS or ERROR.
        """
        message = CsvFile.read(self)

        if message.startswith('ERROR'):
            return message

        for cells in self.rows:

            if 'LcID:' in cells[0]:
                lcId = re.search('LcID\:([0-9]+)', cells[0]).group(1)
                self.srtLocations.append(lcId)
                self.locations[lcId] = WorldElement()
                self.locations[lcId].title = cells[1]
                self.locations[lcId].desc = self.convert_to_yw(cells[2])
                self.locations[lcId].aka = cells[3]
                self.locations[lcId].tags = self.get_list(cells[4])

        return 'SUCCESS: Data read from "' + os.path.normpath(self.filePath) + '".'

    def merge(self, novel):
        """Copy required attributes of the novel object.
        Return a message beginning with SUCCESS or ERROR.
        """
        self.srtLocations = novel.srtLocations
        self.locations = novel.locations
        return 'SUCCESS'



class CsvItemList(CsvFile):
    """csv file representation of an yWriter project's items table. 

    Represents a csv file with a record per item.
    - Records are separated by line breaks.
    - Data fields are delimited by the _SEPARATOR item.
    """

    DESCRIPTION = 'Item list'
    SUFFIX = '_itemlist'

    fileHeader = '''ID|Name|Description|Aka|Tags
'''

    itemTemplate = '''ItID:$ID|$Title|"$Desc"|$AKA|$Tags
'''

    def read(self):
        """Parse the csv file located at filePath, 
        fetching the WorldElement attributes contained.
        Return a message beginning with SUCCESS or ERROR.
        """
        message = CsvFile.read(self)

        if message.startswith('ERROR'):
            return message

        for cells in self.rows:

            if 'ItID:' in cells[0]:
                itId = re.search('ItID\:([0-9]+)', cells[0]).group(1)
                self.srtItems.append(itId)
                self.items[itId] = WorldElement()
                self.items[itId].title = cells[1]
                self.items[itId].desc = self.convert_to_yw(cells[2])
                self.items[itId].aka = cells[3]
                self.items[itId].tags = self.get_list(cells[4])

        return 'SUCCESS: Data read from "' + os.path.normpath(self.filePath) + '".'

    def merge(self, novel):
        """Copy required attributes of the novel object.
        Return a message beginning with SUCCESS or ERROR.
        """
        self.srtItems = novel.srtItems
        self.items = novel.items
        return 'SUCCESS'


class UniversalConverter(YwCnvUi):
    """A converter for universal import and export.

    Support yWriter 7 projects and most of the Novel subclasses 
    that can be read or written by OpenOffice/LibreOffice.

    Override the superclass constants EXPORT_SOURCE_CLASSES,
    EXPORT_TARGET_CLASSES, IMPORT_SOURCE_CLASSES, IMPORT_TARGET_CLASSES.

    Class constants:
        CREATE_SOURCE_CLASSES -- list of classes that - additional to HtmlImport
                        and HtmlOutline - can be exported to a new yWriter project.
    """
    EXPORT_SOURCE_CLASSES = [Yw7File,
                             Yw6File,
                             Yw5File,
                             ]
    EXPORT_TARGET_CLASSES = [OdtExport,
                             OdtProof,
                             OdtManuscript,
                             OdtSceneDesc,
                             OdtChapterDesc,
                             OdtPartDesc,
                             OdtCharacters,
                             OdtItems,
                             OdtLocations,
                             OdsCharList,
                             OdsLocList,
                             OdsItemList,
                             OdsSceneList,
                             OdsPlotList,
                             OdtXref,
                             ]
    IMPORT_SOURCE_CLASSES = [HtmlProof,
                             HtmlManuscript,
                             HtmlSceneDesc,
                             HtmlChapterDesc,
                             HtmlPartDesc,
                             HtmlCharacters,
                             HtmlItems,
                             HtmlLocations,
                             CsvCharList,
                             CsvLocList,
                             CsvItemList,
                             CsvSceneList,
                             CsvPlotList,
                             ]
    IMPORT_TARGET_CLASSES = [Yw7File,
                             Yw6File,
                             Yw5File,
                             ]
    CREATE_SOURCE_CLASSES = []

    def __init__(self):
        """Extend the superclass constructor.
        Override newProjectFactory.
        """
        YwCnvUi.__init__(self)
        self.newProjectFactory = NewProjectFactory(self.CREATE_SOURCE_CLASSES)


class YwCnvUno(UniversalConverter):
    """Converter for yWriter project files.
    Variant with UNO UI.
    """

    def export_from_yw(self, sourceFile, targetFile):
        """Method for conversion from yw to other.
        Show only error messages.
        """
        message = self.convert(sourceFile, targetFile)

        if message.startswith('SUCCESS'):
            self.newFile = targetFile.filePath

        else:
            self.ui.set_info_how(message)
from com.sun.star.awt.MessageBoxResults import OK, YES, NO, CANCEL
from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK, BUTTONS_OK_CANCEL, BUTTONS_YES_NO, BUTTONS_YES_NO_CANCEL, BUTTONS_RETRY_CANCEL, BUTTONS_ABORT_IGNORE_RETRY
from com.sun.star.awt.MessageBoxType import MESSAGEBOX, INFOBOX, WARNINGBOX, ERRORBOX, QUERYBOX



class UiUno(Ui):
    """UI subclass implementing a LibreOffice UNO facade."""

    def ask_yes_no(self, text):
        result = msgbox(text, buttons=BUTTONS_YES_NO,
                        type_msg=WARNINGBOX)

        if result == YES:
            return True

        else:
            return False

    def set_info_how(self, message):
        """How's the converter doing?"""
        self.infoHowText = message

        if message.startswith('SUCCESS'):
            msgbox(message, type_msg=INFOBOX)

        else:
            msgbox(message, type_msg=ERRORBOX)

INI_FILE = 'openyw.ini'


def open_yw7(suffix, newExt):

    # Set last opened yWriter project as default (if existing).

    scriptLocation = os.path.dirname(__file__)
    inifile = unquote(
        (scriptLocation + '/' + INI_FILE).replace('file:///', ''))
    defaultFile = None
    config = ConfigParser()

    try:
        config.read(inifile)
        ywLastOpen = config.get('FILES', 'yw_last_open')

        if os.path.isfile(ywLastOpen):
            defaultFile = quote('file:///' + ywLastOpen, '/:')

    except:
        pass

    # Ask for yWriter 6 or 7 project to open:

    ywFile = FilePicker(path=defaultFile)

    if ywFile is None:
        return

    sourcePath = unquote(ywFile.replace('file:///', ''))
    ywExt = os.path.splitext(sourcePath)[1]

    if not ywExt in ['.yw6', '.yw7']:
        msgbox('Please choose a yWriter 6/7 project.',
               'Import from yWriter', type_msg=ERRORBOX)
        return

    # Store selected yWriter project as "last opened".

    newFile = ywFile.replace(ywExt, suffix + newExt)
    dirName, filename = os.path.split(newFile)
    lockFile = unquote((dirName + '/.~lock.' + filename +
                        '#').replace('file:///', ''))

    if not config.has_section('FILES'):
        config.add_section('FILES')

    config.set('FILES', 'yw_last_open', unquote(
        ywFile.replace('file:///', '')))

    with open(inifile, 'w') as f:
        config.write(f)

    # Check if import file is already open in LibreOffice:

    if os.path.isfile(lockFile):
        msgbox('Please close "' + filename + '" first.',
               'Import from yWriter', type_msg=ERRORBOX)
        return

    # Open yWriter project and convert data.

    workdir = os.path.dirname(sourcePath)
    os.chdir(workdir)
    converter = YwCnvUno()
    converter.ui = UiUno('Import from yWriter')
    kwargs = {'suffix': suffix}
    converter.run(sourcePath, **kwargs)

    if converter.newFile:
        desktop = XSCRIPTCONTEXT.getDesktop()
        doc = desktop.loadComponentFromURL(newFile, "_blank", 0, ())


def import_yw():
    '''Import scenes from yWriter 6/7 to a Writer document
    without chapter and scene markers. 
    '''
    open_yw7('', '.odt')


def proof_yw():
    '''Import scenes from yWriter 6/7 to a Writer document
    with visible chapter and scene markers. 
    '''
    open_yw7(OdtProof.SUFFIX, OdtProof.EXTENSION)


def get_manuscript():
    '''Import scenes from yWriter 6/7 to a Writer document
    with invisible chapter and scene markers. 
    '''
    open_yw7(OdtManuscript.SUFFIX, OdtManuscript.EXTENSION)


def get_partdesc():
    '''Import pard descriptions from yWriter 6/7 to a Writer document
    with invisible chapter and scene markers. 
    '''
    open_yw7(OdtPartDesc.SUFFIX, OdtPartDesc.EXTENSION)


def get_chapterdesc():
    '''Import chapter descriptions from yWriter 6/7 to a Writer document
    with invisible chapter and scene markers. 
    '''
    open_yw7(OdtChapterDesc.SUFFIX, OdtChapterDesc.EXTENSION)


def get_scenedesc():
    '''Import scene descriptions from yWriter 6/7 to a Writer document
    with invisible chapter and scene markers. 
    '''
    open_yw7(OdtSceneDesc.SUFFIX, OdtSceneDesc.EXTENSION)


def get_chardesc():
    '''Import character descriptions from yWriter 6/7 to a Writer document.
    '''
    open_yw7(OdtCharacters.SUFFIX, OdtCharacters.EXTENSION)


def get_locdesc():
    '''Import location descriptions from yWriter 6/7 to a Writer document.
    '''
    open_yw7(OdtLocations.SUFFIX, OdtLocations.EXTENSION)


def get_itemdesc():
    '''Import item descriptions from yWriter 6/7 to a Writer document.
    '''
    open_yw7(OdtItems.SUFFIX, OdtItems.EXTENSION)


def get_xref():
    '''Generate cross references from yWriter 6/7 to a Writer document.
    '''
    open_yw7(OdtXref.SUFFIX, OdtXref.EXTENSION)


def get_scenelist():
    '''Import a scene list from yWriter 6/7 to a Calc document.
    '''
    open_yw7(OdsSceneList.SUFFIX, OdsSceneList.EXTENSION)


def get_plotlist():
    '''Import a plot list from yWriter 6/7 to a Calc document.
    '''
    open_yw7(OdsPlotList.SUFFIX, OdsPlotList.EXTENSION)


def get_charlist():
    '''Import a character list from yWriter 6/7 to a Calc document.
    '''
    open_yw7(OdsCharList.SUFFIX, OdsCharList.EXTENSION)


def get_loclist():
    '''Import a location list from yWriter 6/7 to a Calc document.
    '''
    open_yw7(OdsLocList.SUFFIX, OdsLocList.EXTENSION)


def get_itemlist():
    '''Import an item list from yWriter 6/7 to a Calc document.
    '''
    open_yw7(OdsItemList.SUFFIX, OdsItemList.EXTENSION)


def export_yw():
    '''Export the document to a yWriter 6/7 project.
    '''

    # Get document's filename

    document = XSCRIPTCONTEXT.getDocument().CurrentController.Frame
    # document   = ThisComponent.CurrentController.Frame

    ctx = XSCRIPTCONTEXT.getComponentContext()
    smgr = ctx.getServiceManager()
    dispatcher = smgr.createInstanceWithContext(
        "com.sun.star.frame.DispatchHelper", ctx)
    # dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")

    documentPath = XSCRIPTCONTEXT.getDocument().getURL()
    # documentPath = ThisComponent.getURL()

    from com.sun.star.beans import PropertyValue
    args1 = []
    args1.append(PropertyValue())
    args1.append(PropertyValue())
    # dim args1(1) as new com.sun.star.beans.PropertyValue

    if documentPath.endswith('.odt') or documentPath.endswith('.html'):
        odtPath = documentPath.replace('.html', '.odt')
        htmlPath = documentPath.replace('.odt', '.html')

        # Save document in HTML format

        args1[0].Name = 'URL'
        # args1(0).Name = "URL"
        args1[0].Value = htmlPath
        # args1(0).Value = htmlPath
        args1[1].Name = 'FilterName'
        # args1(1).Name = "FilterName"
        args1[1].Value = 'HTML (StarWriter)'
        # args1(1).Value = "HTML (StarWriter)"
        dispatcher.executeDispatch(document, ".uno:SaveAs", "", 0, args1)
        # dispatcher.executeDispatch(document, ".uno:SaveAs", "", 0, args1())

        # Save document in OpenDocument format

        args1[0].Value = odtPath
        # args1(0).Value = odtPath
        args1[1].Value = 'writer8'
        # args1(1).Value = "writer8"
        dispatcher.executeDispatch(document, ".uno:SaveAs", "", 0, args1)
        # dispatcher.executeDispatch(document, ".uno:SaveAs", "", 0, args1())

        targetPath = unquote(htmlPath.replace('file:///', ''))

    elif documentPath.endswith('.ods') or documentPath.endswith('.csv'):
        odsPath = documentPath.replace('.csv', '.ods')
        csvPath = documentPath.replace('.ods', '.csv')

        # Save document in csv format

        args1.append(PropertyValue())

        args1[0].Name = 'URL'
        # args1(0).Name = "URL"
        args1[0].Value = csvPath
        # args1(0).Value = csvPath
        args1[1].Name = 'FilterName'
        # args1(1).Name = "FilterName"
        args1[1].Value = 'Text - txt - csv (StarCalc)'
        # args1(1).Value = "Text - txt - csv (StarCalc)"
        args1[2].Name = "FilterOptions"
        # args1(2).Name = "FilterOptions"
        args1[2].Value = "124,34,76,1,,0,false,true,true"
        # args1(2).Value = "124,34,76,1,,0,false,true,true"
        dispatcher.executeDispatch(document, ".uno:SaveAs", "", 0, args1)
        # dispatcher.executeDispatch(document, ".uno:SaveAs", "", 0, args1())

        # Save document in OpenDocument format

        args1[0].Value = odsPath
        # args1(0).Value = odsPath
        args1[1].Value = 'calc8'
        # args1(1).Value = "calc8"
        dispatcher.executeDispatch(document, ".uno:SaveAs", "", 0, args1)
        # dispatcher.executeDispatch(document, ".uno:SaveAs", "", 0, args1())

        targetPath = unquote(csvPath.replace('file:///', ''))

    else:
        msgbox('ERROR: File type of "' + os.path.normpath(documentPath) +
               '" not supported.', type_msg=ERRORBOX)

    converter = YwCnvUno()
    converter.ui = UiUno('Export to yWriter')
    kwargs = {'suffix': None}
    converter.run(targetPath, **kwargs)
