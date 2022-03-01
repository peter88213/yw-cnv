"""Convert yWriter project to odt or ods and vice versa. 

Version 1.14.1
Requires Python 3.6+
Copyright (c) 2022 Peter Triesberger
For further information see https://github.com/peter88213/yw-cnv
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import uno
from com.sun.star.awt.MessageBoxType import MESSAGEBOX, INFOBOX, WARNINGBOX, ERRORBOX, QUERYBOX
from com.sun.star.beans import PropertyValue
import os
from configparser import ConfigParser
import re

ERROR = '!'
import zipfile
import locale
import tempfile
from shutil import rmtree
from datetime import datetime
from string import Template
from string import Template


class WorldElement:
    """Story world element representation (may be location or item).
    
    Public instance variables:
        title -- str: title (name).
        image -- str: image file path.
        desc -- str: description.
        tags -- list of tags.
        aka -- str: alternate name.
    """

    def __init__(self):
        """Initialize instance variables."""
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

    Public instance variables:
        notes -- str: character notes.
        bio -- str: character biography.
        goals -- str: character's goals in the story.
        fullName -- str: full name (the title inherited may be a short name).
        isMajor -- bool: True, if it's a major character.
    """
    MAJOR_MARKER = 'Major'
    MINOR_MARKER = 'Minor'

    def __init__(self):
        """Extends the superclass constructor by adding instance variables."""
        super().__init__()

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


class Scene:
    """yWriter scene representation.
    
    Public instance variables:
        title -- str: scene title.
        desc -- str: scene description in a single string.
        sceneContent -- str: scene content (property with getter and setter).
        rtfFile -- str: RTF file name (yWriter 5).
        wordCount - int: word count (derived; updated by the sceneContent setter).
        letterCount - int: letter count (derived; updated by the sceneContent setter).
        isUnused -- bool: True if the scene is marked "Unused". 
        isNotesScene -- bool: True if the scene type is "Notes".
        isTodoScene -- bool: True if the scene type is "Todo". 
        doNotExport -- bool: True if the scene is not to be exported to RTF.
        status -- int: scene status (Outline/Draft/1st Edit/2nd Edit/Done).
        sceneNotes -- str: scene notes in a single string.
        tags -- list of scene tags. 
        field1 -- int: scene ratings field 1.
        field2 -- int: scene ratings field 2.
        field3 -- int: scene ratings field 3.
        field4 -- int: scene ratings field 4.
        appendToPrev -- bool: if True, append the scene without a divider to the previous scene.
        isReactionScene -- bool: if True, the scene is "reaction". Otherwise, it's "action". 
        isSubPlot -- bool: if True, the scene belongs to a sub-plot. Otherwise it's main plot.  
        goal -- str: the main actor's scene goal. 
        conflict -- str: what hinders the main actor to achieve his goal.
        outcome -- str: what comes out at the end of the scene.
        characters -- list of character IDs related to this scene.
        locations -- list of location IDs related to this scene. 
        items -- list of item IDs related to this scene.
        date -- str: specific start date in ISO format (yyyy-mm-dd).
        time -- str: specific start time in ISO format (hh:mm).
        minute -- str: unspecific start time: minutes.
        hour -- str: unspecific start time: hour.
        day -- str: unspecific start time: day.
        lastsMinutes -- str: scene duration: minutes.
        lastsHours -- str: scene duration: hours.
        lastsDays -- str: scene duration: days. 
        image -- str:  path to an image related to the scene. 
    """
    STATUS = (None, 'Outline', 'Draft', '1st Edit', '2nd Edit', 'Done')
    # Emulate an enumeration for the scene status
    # Since the items are used to replace text,
    # they may contain spaces. This is why Enum cannot be used here.

    ACTION_MARKER = 'A'
    REACTION_MARKER = 'R'
    NULL_DATE = '0001-01-01'
    NULL_TIME = '00:00:00'

    def __init__(self):
        """Initialize instance variables."""
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
        # int
        # xml: <Status>
        # 1 - Outline
        # 2 - Draft
        # 3 - 1st Edit
        # 4 - 2nd Edit
        # 5 - Done
        # See also the STATUS list for conversion.

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

        self.image = None
        # str
        # xml: <ImageFile>

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


class Novel:
    """Abstract yWriter project file representation.

    This class represents a file containing a novel with additional 
    attributes and structural information (a full set or a subset
    of the information included in an yWriter project file).

    Public methods:
        read() -- parse the file and get the instance variables.
        merge(source) -- update instance variables from a source instance.
        write() -- write instance variables to the file.

    Public instance variables:
        title -- str: title.
        desc -- str: description in a single string.
        authorName -- str: author's name.
        author bio -- str: information about the author.
        fieldTitle1 -- str: scene rating field title 1.
        fieldTitle2 -- str: scene rating field title 2.
        fieldTitle3 -- str: scene rating field title 3.
        fieldTitle4 -- str: scene rating field title 4.
        chapters -- dict: (key: ID; value: chapter instance).
        scenes -- dict: (key: ID, value: scene instance).
        srtChapters -- list: the novel's sorted chapter IDs.
        locations -- dict: (key: ID, value: WorldElement instance).
        srtLocations -- list: the novel's sorted location IDs.
        items -- dict: (key: ID, value: WorldElement instance).
        srtItems -- list: the novel's sorted item IDs.
        characters -- dict: (key: ID, value: character instance).
        srtCharacters -- list: the novel's sorted character IDs.
        filePath -- str: path to the file (property with getter and setter). 
    """
    DESCRIPTION = 'Novel'
    EXTENSION = None
    SUFFIX = None
    # To be extended by subclass methods.

    def __init__(self, filePath, **kwargs):
        """Initialize instance variables.

        Positional arguments:
            filePath -- str: path to the file represented by the Novel instance.
            
        Optional arguments:
            kwargs -- keyword arguments to be used by subclasses.            
        """
        self.title = None
        # str
        # xml: <PROJECT><Title>

        self.desc = None
        # str
        # xml: <PROJECT><Desc>

        self.authorName = None
        # str
        # xml: <PROJECT><AuthorName>

        self.authorBio = None
        # str
        # xml: <PROJECT><Bio>

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
        # The order of the elements does not matter (the novel's order of the chapters is defined by srtChapters)

        self.scenes = {}
        # dict
        # xml: <SCENES><SCENE><ID>
        # key = scene ID, value = Scene instance.
        # The order of the elements does not matter (the novel's order of the scenes is defined by
        # the order of the chapters and the order of the scenes within the chapters)

        self.srtChapters = []
        # list of str
        # The novel's chapter IDs. The order of its elements corresponds to the novel's order of the chapters.

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
        # The novel's item IDs. The order of its elements corresponds to the XML project file.

        self.characters = {}
        # dict
        # xml: <CHARACTERS>
        # key = character ID, value = Character instance.
        # The order of the elements does not matter.

        self.srtCharacters = []
        # list of str
        # The novel's character IDs. The order of its elements corresponds to the XML project file.

        self._filePath = None
        # str
        # Path to the file. The setter only accepts files of a supported type as specified by EXTENSION.

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
        if filePath.lower().endswith(f'{suffix}{self.EXTENSION}'.lower()):
            self._filePath = filePath
            head, tail = os.path.split(os.path.realpath(filePath))
            self.projectPath = quote(head.replace('\\', '/'), '/:')
            self.projectName = quote(tail.replace(f'{suffix}{self.EXTENSION}', ''))

    def read(self):
        """Parse the file and get the instance variables.
        
        Return a message beginning with the ERROR constant in case of error.
        This is a stub to be overridden by subclass methods.
        """
        return f'{ERROR}Read method is not implemented.'

    def merge(self, source):
        """Update instance variables from a source instance.
        
        Positional arguments:
            source -- Novel subclass instance to merge.
        
        Return a message beginning with the ERROR constant in case of error.
        This is a stub to be overridden by subclass methods.
        """
        return f'{ERROR}Merge method is not implemented.'

    def write(self):
        """Write instance variables to the file.
        
        Return a message beginning with the ERROR constant in case of error.
        This is a stub to be overridden by subclass methods.
        """
        return f'{ERROR}Write method is not implemented.'

    def _convert_to_yw(self, text):
        """Return text, converted from source format to yw7 markup.
        
        Positional arguments:
            text -- string to convert.
        
        This is a stub to be overridden by subclass methods.
        """
        return text

    def _convert_from_yw(self, text, quick=False):
        """Return text, converted from yw7 markup to target format.
        
        Positional arguments:
            text -- string to convert.
        
        Optional arguments:
            quick -- bool: if True, apply a conversion mode for one-liners without formatting.
        
        This is a stub to be overridden by subclass methods.
        """
        return text


class Filter:
    """Filter an entity (chapter/scene/character/location/item) by filter criteria.
    
    Public methods:
        accept(source, eId) -- check whether an entity matches the filter criteria.
    
    Strategy class, implementing filtering criteria for template-based export.
    This is a stub with no filter criteria specified.
    """

    def accept(self, source, eId):
        """Check whether an entity matches the filter criteria.
        
        Positional arguments:
            source -- Novel instance holding the entity to check.
            eId -- ID of the entity to check.       
        
        Return True if the entity is not to be filtered out.
        This is a stub to be overridden by subclass methods implementing filters.
        """
        return True


class FileExport(Novel):
    """Abstract yWriter project file exporter representation.
    
    Public methods:
        merge(source) -- update instance variables from a source instance.
        write() -- write instance variables to the export file.
    
    This class is generic and contains no conversion algorithm and no templates.
    """
    SUFFIX = ''
    _fileHeader = ''
    _partTemplate = ''
    _chapterTemplate = ''
    _notesChapterTemplate = ''
    _todoChapterTemplate = ''
    _unusedChapterTemplate = ''
    _notExportedChapterTemplate = ''
    _sceneTemplate = ''
    _firstSceneTemplate = ''
    _appendedSceneTemplate = ''
    _notesSceneTemplate = ''
    _todoSceneTemplate = ''
    _unusedSceneTemplate = ''
    _notExportedSceneTemplate = ''
    _sceneDivider = ''
    _chapterEndTemplate = ''
    _unusedChapterEndTemplate = ''
    _notExportedChapterEndTemplate = ''
    _notesChapterEndTemplate = ''
    _todoChapterEndTemplate = ''
    _characterSectionHeading = ''
    _characterTemplate = ''
    _locationSectionHeading = ''
    _locationTemplate = ''
    _itemSectionHeading = ''
    _itemTemplate = ''
    _fileFooter = ''

    def __init__(self, filePath, **kwargs):
        """Initialize filter strategy class instances.
        
        Positional arguments:
            filePath -- str: path to the file represented by the Novel instance.
            
        Optional arguments:
            kwargs -- keyword arguments to be used by subclasses.            

        Extends the superclass constructor.
        """
        super().__init__(filePath, **kwargs)
        self._sceneFilter = Filter()
        self._chapterFilter = Filter()
        self._characterFilter = Filter()
        self._locationFilter = Filter()
        self._itemFilter = Filter()

    def merge(self, source):
        """Update instance variables from a source instance.
        
        Positional arguments:
            source -- Novel subclass instance to merge.
        
        Return a message beginning with the ERROR constant in case of error.
        Overrides the superclass method.
        """
        if source.title is not None:
            self.title = source.title
        else:
            self.title = ''

        if source.desc is not None:
            self.desc = source.desc
        else:
            self.desc = ''

        if source.authorName is not None:
            self.authorName = source.authorName
        else:
            self.authorName = ''

        if source.authorBio is not None:
            self.authorBio = source.authorBio
        else:
            self.authorBio = ''

        if source.fieldTitle1 is not None:
            self.fieldTitle1 = source.fieldTitle1
        else:
            self.fieldTitle1 = 'Field 1'
        
        if source.fieldTitle2 is not None:
            self.fieldTitle2 = source.fieldTitle2
        else:
            self.fieldTitle2 = 'Field 2'
        
        if source.fieldTitle3 is not None:
            self.fieldTitle3 = source.fieldTitle3
        else:
            self.fieldTitle3 = 'Field 3'
        
        if source.fieldTitle4 is not None:
            self.fieldTitle4 = source.fieldTitle4
        else:
            self.fieldTitle4 = 'Field 4'
        
        if source.srtChapters:
            self.srtChapters = source.srtChapters
        
        if source.scenes is not None:
            self.scenes = source.scenes
        
        if source.chapters is not None:
            self.chapters = source.chapters
        
        if source.srtCharacters:
            self.srtCharacters = source.srtCharacters
            self.characters = source.characters
        
        if source.srtLocations:
            self.srtLocations = source.srtLocations
            self.locations = source.locations
        
        if source.srtItems:
            self.srtItems = source.srtItems
            self.items = source.items
        return 'Export data updated from novel.'

    def _get_fileHeaderMapping(self):
        """Return a mapping dictionary for the project section.
        
        This is a template method that can be extended or overridden by subclasses.
        """
        projectTemplateMapping = dict(
            Title=self._convert_from_yw(self.title, True),
            Desc=self._convert_from_yw(self.desc),
            AuthorName=self._convert_from_yw(self.authorName, True),
            AuthorBio=self._convert_from_yw(self.authorBio, True),
            FieldTitle1=self._convert_from_yw(self.fieldTitle1, True),
            FieldTitle2=self._convert_from_yw(self.fieldTitle2, True),
            FieldTitle3=self._convert_from_yw(self.fieldTitle3, True),
            FieldTitle4=self._convert_from_yw(self.fieldTitle4, True),
        )
        return projectTemplateMapping

    def _get_chapterMapping(self, chId, chapterNumber):
        """Return a mapping dictionary for a chapter section.
        
        Positional arguments:
            chId -- str: chapter ID.
            chapterNumber -- int: chapter number.
        
        This is a template method that can be extended or overridden by subclasses.
        """
        if chapterNumber == 0:
            chapterNumber = ''
        
        chapterMapping = dict(
            ID=chId,
            ChapterNumber=chapterNumber,
            Title=self._convert_from_yw(self.chapters[chId].title, True),
            Desc=self._convert_from_yw(self.chapters[chId].desc),
            ProjectName=self._convert_from_yw(self.projectName, True),
            ProjectPath=self.projectPath,
        )
        return chapterMapping

    def _get_sceneMapping(self, scId, sceneNumber, wordsTotal, lettersTotal):
        """Return a mapping dictionary for a scene section.
        
        Positional arguments:
            scId -- str: scene ID.
            sceneNumber -- int: scene number to be displayed.
            wordsTotal -- int: accumulated wordcount.
            lettersTotal -- int: accumulated lettercount.
        
        This is a template method that can be extended or overridden by subclasses.
        """
        
        #--- Create a comma separated tag list.
        if sceneNumber == 0:
            sceneNumber = ''
        if self.scenes[scId].tags is not None:
            tags = self._get_string(self.scenes[scId].tags)
        else:
            tags = ''

        #--- Create a comma separated character list.
        try:
            # Note: Due to a bug, yWriter scenes might hold invalid
            # viepoint characters
            sChList = []
            for chId in self.scenes[scId].characters:
                sChList.append(self.characters[chId].title)
            sceneChars = self._get_string(sChList)
            viewpointChar = sChList[0]
        except:
            sceneChars = ''
            viewpointChar = ''

        #--- Create a comma separated location list.
        if self.scenes[scId].locations is not None:
            sLcList = []
            for lcId in self.scenes[scId].locations:
                sLcList.append(self.locations[lcId].title)
            sceneLocs = self._get_string(sLcList)
        else:
            sceneLocs = ''

        #--- Create a comma separated item list.
        if self.scenes[scId].items is not None:
            sItList = []
            for itId in self.scenes[scId].items:
                sItList.append(self.items[itId].title)
            sceneItems = self._get_string(sItList)
        else:
            sceneItems = ''

        #--- Create A/R marker string.
        if self.scenes[scId].isReactionScene:
            reactionScene = Scene.REACTION_MARKER
        else:
            reactionScene = Scene.ACTION_MARKER

        #--- Create a combined scDate information.
        if self.scenes[scId].date is not None and self.scenes[scId].date != Scene.NULL_DATE:
            scDay = ''
            scDate = self.scenes[scId].date
            cmbDate = self.scenes[scId].date
        else:
            scDate = ''
            if self.scenes[scId].day is not None:
                scDay = self.scenes[scId].day
                cmbDate = f'Day {self.scenes[scId].day}'
            else:
                scDay = ''
                cmbDate = ''

        #--- Create a combined time information.
        if self.scenes[scId].time is not None and self.scenes[scId].date != Scene.NULL_DATE:
            scHour = ''
            scMinute = ''
            scTime = self.scenes[scId].time
            cmbTime = self.scenes[scId].time.rsplit(':', 1)[0]
        else:
            scTime = ''
            if self.scenes[scId].hour or self.scenes[scId].minute:
                if self.scenes[scId].hour:
                    scHour = self.scenes[scId].hour
                else:
                    scHour = '00'
                if self.scenes[scId].minute:
                    scMinute = self.scenes[scId].minute
                else:
                    scMinute = '00'
                cmbTime = f'{scHour.zfill(2)}:{scMinute.zfill(2)}'
            else:
                scHour = ''
                scMinute = ''
                cmbTime = ''

        #--- Create a combined duration information.
        if self.scenes[scId].lastsDays is not None and self.scenes[scId].lastsDays != '0':
            lastsDays = self.scenes[scId].lastsDays
            days = f'{self.scenes[scId].lastsDays}d '
        else:
            lastsDays = ''
            days = ''
        if self.scenes[scId].lastsHours is not None and self.scenes[scId].lastsHours != '0':
            lastsHours = self.scenes[scId].lastsHours
            hours = f'{self.scenes[scId].lastsHours}h '
        else:
            lastsHours = ''
            hours = ''
        if self.scenes[scId].lastsMinutes is not None and self.scenes[scId].lastsMinutes != '0':
            lastsMinutes = self.scenes[scId].lastsMinutes
            minutes = f'{self.scenes[scId].lastsMinutes}min'
        else:
            lastsMinutes = ''
            minutes = ''
        duration = f'{days}{hours}{minutes}'
        
        sceneMapping = dict(
            ID=scId,
            SceneNumber=sceneNumber,
            Title=self._convert_from_yw(self.scenes[scId].title, True),
            Desc=self._convert_from_yw(self.scenes[scId].desc),
            WordCount=str(self.scenes[scId].wordCount),
            WordsTotal=wordsTotal,
            LetterCount=str(self.scenes[scId].letterCount),
            LettersTotal=lettersTotal,
            Status=Scene.STATUS[self.scenes[scId].status],
            SceneContent=self._convert_from_yw(self.scenes[scId].sceneContent),
            FieldTitle1=self._convert_from_yw(self.fieldTitle1, True),
            FieldTitle2=self._convert_from_yw(self.fieldTitle2, True),
            FieldTitle3=self._convert_from_yw(self.fieldTitle3, True),
            FieldTitle4=self._convert_from_yw(self.fieldTitle4, True),
            Field1=self.scenes[scId].field1,
            Field2=self.scenes[scId].field2,
            Field3=self.scenes[scId].field3,
            Field4=self.scenes[scId].field4,
            Date=scDate,
            Time=scTime,
            Day=scDay,
            Hour=scHour,
            Minute=scMinute,
            ScDate=cmbDate,
            ScTime=cmbTime,
            LastsDays=lastsDays,
            LastsHours=lastsHours,
            LastsMinutes=lastsMinutes,
            Duration=duration,
            ReactionScene=reactionScene,
            Goal=self._convert_from_yw(self.scenes[scId].goal),
            Conflict=self._convert_from_yw(self.scenes[scId].conflict),
            Outcome=self._convert_from_yw(self.scenes[scId].outcome),
            Tags=self._convert_from_yw(tags, True),
            Image=self.scenes[scId].image,
            Characters=sceneChars,
            Viewpoint=viewpointChar,
            Locations=sceneLocs,
            Items=sceneItems,
            Notes=self._convert_from_yw(self.scenes[scId].sceneNotes),
            ProjectName=self._convert_from_yw(self.projectName, True),
            ProjectPath=self.projectPath,
        )
        return sceneMapping

    def _get_characterMapping(self, crId):
        """Return a mapping dictionary for a character section.
        
        Positional arguments:
            crId -- str: character ID.
        
        This is a template method that can be extended or overridden by subclasses.
        """
        if self.characters[crId].tags is not None:
            tags = self._get_string(self.characters[crId].tags)
        else:
            tags = ''
        if self.characters[crId].isMajor:
            characterStatus = Character.MAJOR_MARKER
        else:
            characterStatus = Character.MINOR_MARKER
        
        characterMapping = dict(
            ID=crId,
            Title=self._convert_from_yw(self.characters[crId].title, True),
            Desc=self._convert_from_yw(self.characters[crId].desc),
            Tags=self._convert_from_yw(tags),
            Image=self.characters[crId].image,
            AKA=self._convert_from_yw(self.characters[crId].aka, True),
            Notes=self._convert_from_yw(self.characters[crId].notes),
            Bio=self._convert_from_yw(self.characters[crId].bio),
            Goals=self._convert_from_yw(self.characters[crId].goals),
            FullName=self._convert_from_yw(self.characters[crId].fullName, True),
            Status=characterStatus,
            ProjectName=self._convert_from_yw(self.projectName),
            ProjectPath=self.projectPath,
        )
        return characterMapping

    def _get_locationMapping(self, lcId):
        """Return a mapping dictionary for a location section.
        
        Positional arguments:
            lcId -- str: location ID.
        
        This is a template method that can be extended or overridden by subclasses.
        """
        if self.locations[lcId].tags is not None:
            tags = self._get_string(self.locations[lcId].tags)
        else:
            tags = ''
        
        locationMapping = dict(
            ID=lcId,
            Title=self._convert_from_yw(self.locations[lcId].title, True),
            Desc=self._convert_from_yw(self.locations[lcId].desc),
            Tags=self._convert_from_yw(tags, True),
            Image=self.locations[lcId].image,
            AKA=self._convert_from_yw(self.locations[lcId].aka, True),
            ProjectName=self._convert_from_yw(self.projectName, True),
            ProjectPath=self.projectPath,
        )
        return locationMapping

    def _get_itemMapping(self, itId):
        """Return a mapping dictionary for an item section.
        
        Positional arguments:
            itId -- str: item ID.
        
        This is a template method that can be extended or overridden by subclasses.
        """
        if self.items[itId].tags is not None:
            tags = self._get_string(self.items[itId].tags)
        else:
            tags = ''
        
        itemMapping = dict(
            ID=itId,
            Title=self._convert_from_yw(self.items[itId].title, True),
            Desc=self._convert_from_yw(self.items[itId].desc),
            Tags=self._convert_from_yw(tags, True),
            Image=self.items[itId].image,
            AKA=self._convert_from_yw(self.items[itId].aka, True),
            ProjectName=self._convert_from_yw(self.projectName, True),
            ProjectPath=self.projectPath,
        )
        return itemMapping

    def _get_fileHeader(self):
        """Process the file header.
        
        Apply the file header template, substituting placeholders 
        according to the file header mapping dictionary.
        Return a list of strings.
        
        This is a template method that can be extended or overridden by subclasses.
        """
        lines = []
        template = Template(self._fileHeader)
        lines.append(template.safe_substitute(self._get_fileHeaderMapping()))
        return lines

    def _get_scenes(self, chId, sceneNumber, wordsTotal, lettersTotal, doNotExport):
        """Process the scenes.
        
        Positional arguments:
            chId -- str: chapter ID.
            sceneNumber -- int: number of previously processed scenes.
            wordsTotal -- int: accumulated wordcount of the previous scenes.
            lettersTotal -- int: accumulated lettercount of the previous scenes.
            doNotExport -- bool: scene belongs to a chapter that is not to be exported.
        
        Iterate through a sorted scene list and apply the templates, 
        substituting placeholders according to the scene mapping dictionary.
        Skip scenes not accepted by the scene filter.
        
        Return a tuple:
            lines -- list of strings: the lines of the processed scene.
            sceneNumber -- int: number of all processed scenes.
            wordsTotal -- int: accumulated wordcount of all processed scenes.
            lettersTotal -- int: accumulated lettercount of all processed scenes.
        
        This is a template method that can be extended or overridden by subclasses.
        """
        lines = []
        firstSceneInChapter = True
        for scId in self.chapters[chId].srtScenes:
            dispNumber = 0
            if not self._sceneFilter.accept(self, scId):
                continue
            # The order counts; be aware that "Todo" and "Notes" scenes are
            # always unused.
            if self.scenes[scId].isTodoScene:
                if self._todoSceneTemplate:
                    template = Template(self._todoSceneTemplate)
                else:
                    continue

            elif self.scenes[scId].isNotesScene:
                # Scene is "Notes" type.
                if self._notesSceneTemplate:
                    template = Template(self._notesSceneTemplate)
                else:
                    continue

            elif self.scenes[scId].isUnused or self.chapters[chId].isUnused:
                if self._unusedSceneTemplate:
                    template = Template(self._unusedSceneTemplate)
                else:
                    continue

            elif self.chapters[chId].oldType == 1:
                # Scene is "Info" type (old file format).
                if self._notesSceneTemplate:
                    template = Template(self._notesSceneTemplate)
                else:
                    continue

            elif self.scenes[scId].doNotExport or doNotExport:
                if self._notExportedSceneTemplate:
                    template = Template(self._notExportedSceneTemplate)
                else:
                    continue

            else:
                sceneNumber += 1
                dispNumber = sceneNumber
                wordsTotal += self.scenes[scId].wordCount
                lettersTotal += self.scenes[scId].letterCount
                template = Template(self._sceneTemplate)
                if not firstSceneInChapter and self.scenes[scId].appendToPrev and self._appendedSceneTemplate:
                    template = Template(self._appendedSceneTemplate)
            if not (firstSceneInChapter or self.scenes[scId].appendToPrev):
                lines.append(self._sceneDivider)
            if firstSceneInChapter and self._firstSceneTemplate:
                template = Template(self._firstSceneTemplate)
            lines.append(template.safe_substitute(self._get_sceneMapping(
                        scId, dispNumber, wordsTotal, lettersTotal)))
            firstSceneInChapter = False
        return lines, sceneNumber, wordsTotal, lettersTotal

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
        chapterNumber = 0
        sceneNumber = 0
        wordsTotal = 0
        lettersTotal = 0
        for chId in self.srtChapters:
            dispNumber = 0
            if not self._chapterFilter.accept(self, chId):
                continue

            # The order counts; be aware that "Todo" and "Notes" chapters are
            # always unused.
            # Has the chapter only scenes not to be exported?
            sceneCount = 0
            notExportCount = 0
            doNotExport = False
            template = None
            for scId in self.chapters[chId].srtScenes:
                sceneCount += 1
                if self.scenes[scId].doNotExport:
                    notExportCount += 1
            if sceneCount > 0 and notExportCount == sceneCount:
                doNotExport = True
            if self.chapters[chId].chType == 2:
                # Chapter is "ToDo" type (implies "unused").
                if self._todoChapterTemplate:
                    template = Template(self._todoChapterTemplate)
            elif self.chapters[chId].chType == 1:
                # Chapter is "Notes" type (implies "unused").
                if self._notesChapterTemplate:
                    template = Template(self._notesChapterTemplate)
            elif self.chapters[chId].isUnused:
                # Chapter is "really" unused.
                if self._unusedChapterTemplate:
                    template = Template(self._unusedChapterTemplate)
            elif self.chapters[chId].oldType == 1:
                # Chapter is "Info" type (old file format).
                if self._notesChapterTemplate:
                    template = Template(self._notesChapterTemplate)
            elif doNotExport:
                if self._notExportedChapterTemplate:
                    template = Template(self._notExportedChapterTemplate)
            elif self.chapters[chId].chLevel == 1 and self._partTemplate:
                template = Template(self._partTemplate)
            else:
                template = Template(self._chapterTemplate)
                chapterNumber += 1
                dispNumber = chapterNumber
            if template is not None:
                lines.append(template.safe_substitute(self._get_chapterMapping(chId, dispNumber)))

            #--- Process scenes.
            sceneLines, sceneNumber, wordsTotal, lettersTotal = self._get_scenes(
                chId, sceneNumber, wordsTotal, lettersTotal, doNotExport)
            lines.extend(sceneLines)

            #--- Process chapter ending.
            template = None
            if self.chapters[chId].chType == 2:
                if self._todoChapterEndTemplate:
                    template = Template(self._todoChapterEndTemplate)
            elif self.chapters[chId].chType == 1:
                if self._notesChapterEndTemplate:
                    template = Template(self._notesChapterEndTemplate)
            elif self.chapters[chId].isUnused:
                if self._unusedChapterEndTemplate:
                    template = Template(self._unusedChapterEndTemplate)
            elif self.chapters[chId].oldType == 1:
                if self._notesChapterEndTemplate:
                    template = Template(self._notesChapterEndTemplate)
            elif doNotExport:
                if self._notExportedChapterEndTemplate:
                    template = Template(self._notExportedChapterEndTemplate)
            elif self._chapterEndTemplate:
                template = Template(self._chapterEndTemplate)
            if template is not None:
                lines.append(template.safe_substitute(self._get_chapterMapping(chId, dispNumber)))
        return lines

    def _get_characters(self):
        """Process the characters.
        
        Iterate through the sorted character list and apply the template, 
        substituting placeholders according to the character mapping dictionary.
        Skip characters not accepted by the character filter.
        Return a list of strings.
        This is a template method that can be extended or overridden by subclasses.
        """
        if self._characterSectionHeading:
            lines = [self._characterSectionHeading]
        else:
            lines = []
        template = Template(self._characterTemplate)
        for crId in self.srtCharacters:
            if self._characterFilter.accept(self, crId):
                lines.append(template.safe_substitute(self._get_characterMapping(crId)))
        return lines

    def _get_locations(self):
        """Process the locations.
        
        Iterate through the sorted location list and apply the template, 
        substituting placeholders according to the location mapping dictionary.
        Skip locations not accepted by the location filter.
        Return a list of strings.
        This is a template method that can be extended or overridden by subclasses.
        """
        if self._locationSectionHeading:
            lines = [self._locationSectionHeading]
        else:
            lines = []
        template = Template(self._locationTemplate)
        for lcId in self.srtLocations:
            if self._locationFilter.accept(self, lcId):
                lines.append(template.safe_substitute(self._get_locationMapping(lcId)))
        return lines

    def _get_items(self):
        """Process the items. 
        
        Iterate through the sorted item list and apply the template, 
        substituting placeholders according to the item mapping dictionary.
        Skip items not accepted by the item filter.
        Return a list of strings.
        This is a template method that can be extended or overridden by subclasses.
        """
        if self._itemSectionHeading:
            lines = [self._itemSectionHeading]
        else:
            lines = []
        template = Template(self._itemTemplate)
        for itId in self.srtItems:
            if self._itemFilter.accept(self, itId):
                lines.append(template.safe_substitute(self._get_itemMapping(itId)))
        return lines

    def _get_text(self):
        """Call all processing methods.
        
        Return a string to be written to the output file.
        This is a template method that can be extended or overridden by subclasses.
        """
        lines = self._get_fileHeader()
        lines.extend(self._get_chapters())
        lines.extend(self._get_characters())
        lines.extend(self._get_locations())
        lines.extend(self._get_items())
        lines.append(self._fileFooter)
        return ''.join(lines)

    def write(self):
        """Write instance variables to the export file.
        
        Create a template-based output file. 
        Return a message beginning with the ERROR constant in case of error.
        """
        text = self._get_text()
        backedUp = False
        if os.path.isfile(self.filePath):
            try:
                os.replace(self.filePath, f'{self.filePath}.bak')
                backedUp = True            
            except:
                return f'{ERROR}Cannot overwrite "{os.path.normpath(self.filePath)}".'
            
        try:
            with open(self.filePath, 'w', encoding='utf-8') as f:
                f.write(text)
        except:
            if backedUp:
                os.replace(f'{self.filePath}.bak', self.filePath)
            return f'{ERROR}Cannot write "{os.path.normpath(self.filePath)}".'

        return f'"{os.path.normpath(self.filePath)}" written.'

    def _get_string(self, elements):
        """Join strings from a list.
        
        Return a string which is the concatenation of the 
        members of the list of strings "elements", separated by 
        a comma plus a space. The space allows word wrap in 
        spreadsheet cells.
        """
        text = (', ').join(elements)
        return text

    def _convert_from_yw(self, text, quick=False):
        """Return text, converted from yw7 markup to target format.
        
        Positional arguments:
            text -- string to convert.
        
        Optional arguments:
            quick -- bool: if True, apply a conversion mode for one-liners without formatting.
        
        Overrides the superclass method.
        """
        if text is None:
            text = ''
        return(text)


class OdfFile(FileExport):
    """Generic OpenDocument xml file representation.

    Public methods:
        write() -- write instance variables to the export file.
    """
    _ODF_COMPONENTS = []
    _MIMETYPE = ''
    _SETTINGS_XML = ''
    _MANIFEST_XML = ''
    _STYLES_XML = ''
    _META_XML = ''

    def __init__(self, filePath, **kwargs):
        """Create a temporary directory for zipfile generation.
        
        Positional arguments:
            filePath -- str: path to the file represented by the Novel instance.
            
        Optional arguments:
            kwargs -- keyword arguments to be used by subclasses.            

        Extends the superclass constructor,        
        """
        super().__init__(filePath, **kwargs)
        self._tempDir = tempfile.mkdtemp(suffix='.tmp', prefix='odf_')
        self._originalPath = self._filePath

    def __del__(self):
        """Make sure to delete the temporary directory, in case write() has not been called."""
        self._tear_down()

    def _tear_down(self):
        """Delete the temporary directory containing the unpacked ODF directory structure."""
        try:
            rmtree(self._tempDir)
        except:
            pass

    def _set_up(self):
        """Helper method for ZIP file generation.

        Prepare the temporary directory containing the internal structure of an ODF file except 'content.xml'.
        Return a message beginning with the ERROR constant in case of error.
        """
        
        #--- Create and open a temporary directory for the files to zip.
        try:
            self._tear_down()
            os.mkdir(self._tempDir)
            os.mkdir(f'{self._tempDir}/META-INF')
        except:
            return f'{ERROR}Cannot create "{os.path.normpath(self._tempDir)}".'

        #--- Generate mimetype.
        try:
            with open(f'{self._tempDir}/mimetype', 'w', encoding='utf-8') as f:
                f.write(self._MIMETYPE)
        except:
            return f'{ERROR}Cannot write "mimetype"'

        #--- Generate settings.xml.
        try:
            with open(f'{self._tempDir}/settings.xml', 'w', encoding='utf-8') as f:
                f.write(self._SETTINGS_XML)
        except:
            return f'{ERROR}Cannot write "settings.xml"'

        #--- Generate META-INF\manifest.xml.
        try:
            with open(f'{self._tempDir}/META-INF/manifest.xml', 'w', encoding='utf-8') as f:
                f.write(self._MANIFEST_XML)
        except:
            return f'{ERROR}Cannot write "manifest.xml"'

        #--- Generate styles.xml with system language set as document language.
        lng, ctr = locale.getdefaultlocale()[0].split('_')
        localeMapping = dict(
            Language=lng,
            Country=ctr,
        )
        template = Template(self._STYLES_XML)
        text = template.safe_substitute(localeMapping)
        try:
            with open(f'{self._tempDir}/styles.xml', 'w', encoding='utf-8') as f:
                f.write(text)
        except:
            return f'{ERROR}Cannot write "styles.xml"'

        #--- Generate meta.xml with actual document metadata.
        metaMapping = dict(
            Author=self.authorName,
            Title=self.title,
            Summary=f'<![CDATA[{self.desc}]]>',
            Datetime=datetime.today().replace(microsecond=0).isoformat(),
        )
        template = Template(self._META_XML)
        text = template.safe_substitute(metaMapping)
        try:
            with open(f'{self._tempDir}/meta.xml', 'w', encoding='utf-8') as f:
                f.write(text)
        except:
            return f'{ERROR}Cannot write "meta.xml".'

        return 'ODF structure generated.'

    def write(self):
        """Write instance variables to the export file.
        
        Create a template-based output file. 
        Return a message beginning with the ERROR constant in case of error.
        Extends the super class method, adding ZIP file operations.
        """

        #--- Create a temporary directory
        # containing the internal structure of an ODS file except "content.xml".
        message = self._set_up()
        if message.startswith(ERROR):
            return message

        #--- Add "content.xml" to the temporary directory.
        self._originalPath = self._filePath
        self._filePath = f'{self._tempDir}/content.xml'
        message = super().write()
        self._filePath = self._originalPath
        if message.startswith(ERROR):
            return message

        #--- Pack the contents of the temporary directory into the ODF file.
        workdir = os.getcwd()
        backedUp = False
        if os.path.isfile(self.filePath):
            try:
                os.replace(self.filePath, f'{self.filePath}.bak')
                backedUp = True
            except:
                return f'{ERROR}Cannot overwrite "{os.path.normpath(self.filePath)}".'

        try:
            with zipfile.ZipFile(self.filePath, 'w') as odfTarget:
                os.chdir(self._tempDir)
                for file in self._ODF_COMPONENTS:
                    odfTarget.write(file, compress_type=zipfile.ZIP_DEFLATED)
        except:
            if backedUp:
                os.replace(f'{self.filePath}.bak', self.filePath)
            os.chdir(workdir)
            return f'{ERROR}Cannot generate "{os.path.normpath(self.filePath)}".'

        #--- Remove temporary data.
        os.chdir(workdir)
        self._tear_down()
        return f'"{os.path.normpath(self.filePath)}" written.'


class OdtFile(OdfFile):
    """Generic OpenDocument text document representation."""

    EXTENSION = '.odt'
    # overwrites Novel.EXTENSION

    _ODF_COMPONENTS = ['manifest.rdf', 'META-INF', 'content.xml', 'meta.xml', 'mimetype',
                      'settings.xml', 'styles.xml', 'META-INF/manifest.xml']

    _CONTENT_XML_HEADER = '''<?xml version="1.0" encoding="UTF-8"?>

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

    _CONTENT_XML_FOOTER = '''  </office:text>
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
    <meta:creation-date>${Datetime}Z</meta:creation-date>
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

<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" xmlns:math="http://www.w3.org/1998/Math/MathML" xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" xmlns:ooo="http://openoffice.org/2004/office" xmlns:ooow="http://openoffice.org/2004/writer" xmlns:oooc="http://openoffice.org/2004/calc" xmlns:dom="http://www.w3.org/2001/xml-events" xmlns:rpt="http://openoffice.org/2005/report" xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2" xmlns:xhtml="http://www.w3.org/1999/xhtml" xmlns:grddl="http://www.w3.org/2003/g/data-view#" xmlns:tableooo="http://openoffice.org/2009/table" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0">
 <office:font-face-decls>
  <style:font-face style:name="StarSymbol" svg:font-family="StarSymbol" style:font-charset="x-symbol"/>
  <style:font-face style:name="Segoe UI" svg:font-family="&apos;Segoe UI&apos;"/>
  <style:font-face style:name="Courier New" svg:font-family="&apos;Courier New&apos;" style:font-adornments="Standard" style:font-family-generic="modern" style:font-pitch="fixed"/>
  <style:font-face style:name="Cumberland" svg:font-family="Cumberland" style:font-family-generic="modern" style:font-pitch="fixed"/>
  <style:font-face style:name="Courier New1" svg:font-family="&apos;Courier New&apos;" style:font-family-generic="roman" style:font-pitch="fixed"/>
  <style:font-face style:name="ITC Officina Sans Book" svg:font-family="&apos;ITC Officina Sans Book&apos;" style:font-family-generic="swiss" style:font-pitch="variable"/>
 </office:font-face-decls>
 <office:styles>
  <style:default-style style:family="graphic">
   <style:graphic-properties svg:stroke-color="#3465a4" draw:fill-color="#729fcf" fo:wrap-option="no-wrap" draw:shadow-offset-x="0.3cm" draw:shadow-offset-y="0.3cm" draw:start-line-spacing-horizontal="0.283cm" draw:start-line-spacing-vertical="0.283cm" draw:end-line-spacing-horizontal="0.283cm" draw:end-line-spacing-vertical="0.283cm" style:flow-with-text="true"/>
   <style:paragraph-properties style:text-autospace="ideograph-alpha" style:line-break="strict" style:writing-mode="lr-tb" style:font-independent-line-spacing="false">
    <style:tab-stops/>
   </style:paragraph-properties>
   <style:text-properties fo:color="#000000" fo:font-size="10pt" fo:language="$Language" fo:country="$Country" style:font-size-asian="10pt" style:language-asian="zxx" style:country-asian="none" style:font-size-complex="1pt" style:language-complex="zxx" style:country-complex="none"/>
  </style:default-style>
  <style:default-style style:family="paragraph">
   <style:paragraph-properties fo:hyphenation-ladder-count="no-limit" style:text-autospace="ideograph-alpha" style:punctuation-wrap="hanging" style:line-break="strict" style:tab-stop-distance="1.251cm" style:writing-mode="lr-tb"/>
   <style:text-properties fo:color="#000000" style:font-name="Segoe UI" fo:font-size="10pt" fo:language="$Language" fo:country="$Country" style:font-name-asian="Segoe UI" style:font-size-asian="10pt" style:language-asian="zxx" style:country-asian="none" style:font-name-complex="Segoe UI" style:font-size-complex="1pt" style:language-complex="zxx" style:country-complex="none" fo:hyphenate="false" fo:hyphenation-remain-char-count="2" fo:hyphenation-push-char-count="2"/>
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
  <style:style style:name="Text_20_body" style:display-name="Text body" style:family="paragraph" style:parent-style-name="Standard" style:next-style-name="First_20_line_20_indent" style:class="text" style:master-page-name="">
   <style:paragraph-properties style:page-number="auto">
    <style:tab-stops/>
   </style:paragraph-properties>
  </style:style>
  <style:style style:name="First_20_line_20_indent" style:display-name="First line indent" style:family="paragraph" style:parent-style-name="Text_20_body" style:class="text" style:master-page-name="">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="0cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0.499cm" style:auto-text-indent="false" style:page-number="auto"/>
  </style:style>
  <style:style style:name="Hanging_20_indent" style:display-name="Hanging indent" style:family="paragraph" style:parent-style-name="Text_20_body" style:class="text">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="1cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="-0.499cm" style:auto-text-indent="false">
    <style:tab-stops>
     <style:tab-stop style:position="0cm"/>
    </style:tab-stops>
   </style:paragraph-properties>
  </style:style>
  <style:style style:name="Text_20_body_20_indent" style:display-name="Text body indent" style:family="paragraph" style:parent-style-name="Text_20_body" style:class="text">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="0.499cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="Salutation" style:family="paragraph" style:parent-style-name="Standard" style:class="text"/>
  <style:style style:name="Signature" style:family="paragraph" style:parent-style-name="Standard" style:class="text"/>
  <style:style style:name="List_20_Indent" style:display-name="List Indent" style:family="paragraph" style:parent-style-name="Text_20_body" style:class="text">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="5.001cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="-4.5cm" style:auto-text-indent="false">
    <style:tab-stops>
     <style:tab-stop style:position="0cm"/>
    </style:tab-stops>
   </style:paragraph-properties>
  </style:style>
  <style:style style:name="Marginalia" style:family="paragraph" style:parent-style-name="Text_20_body" style:class="text">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="4.001cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="Heading" style:family="paragraph" style:parent-style-name="Standard" style:next-style-name="Text_20_body" style:class="text" style:master-page-name="">
   <style:paragraph-properties fo:line-height="0.73cm" fo:text-align="center" style:justify-single-word="false" style:page-number="auto" fo:keep-with-next="always">
    <style:tab-stops/>
   </style:paragraph-properties>
  </style:style>
  <style:style style:name="Heading_20_1" style:display-name="Heading 1" style:family="paragraph" style:parent-style-name="Heading" style:next-style-name="Text_20_body" style:default-outline-level="1" style:list-style-name="" style:class="text" style:master-page-name="">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin-top="1.461cm" fo:margin-bottom="0.73cm" style:page-number="auto">
    <style:tab-stops/>
   </style:paragraph-properties>
   <style:text-properties fo:text-transform="uppercase" fo:font-weight="bold"/>
  </style:style>
  <style:style style:name="Heading_20_2" style:display-name="Heading 2" style:family="paragraph" style:parent-style-name="Heading" style:next-style-name="Text_20_body" style:default-outline-level="2" style:list-style-name="" style:class="text" style:master-page-name="">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin-top="1.461cm" fo:margin-bottom="0.73cm" style:page-number="auto"/>
   <style:text-properties fo:font-weight="bold"/>
  </style:style>
  <style:style style:name="Heading_20_3" style:display-name="Heading 3" style:family="paragraph" style:parent-style-name="Heading" style:next-style-name="Text_20_body" style:default-outline-level="3" style:list-style-name="" style:class="text" style:master-page-name="">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin-top="0.73cm" fo:margin-bottom="0.73cm" style:page-number="auto"/>
   <style:text-properties fo:font-style="italic"/>
  </style:style>
  <style:style style:name="Heading_20_4" style:display-name="Heading 4" style:family="paragraph" style:parent-style-name="Heading" style:next-style-name="Text_20_body" style:default-outline-level="" style:list-style-name="" style:class="text" style:master-page-name="">
   <style:paragraph-properties fo:margin-top="0.73cm" fo:margin-bottom="0.73cm" style:page-number="auto"/>
  </style:style>
  <style:style style:name="Heading_20_5" style:display-name="Heading 5" style:family="paragraph" style:parent-style-name="Heading" style:next-style-name="Text_20_body" style:default-outline-level="" style:list-style-name="" style:class="text" style:master-page-name="">
   <style:paragraph-properties style:page-number="auto"/>
  </style:style>
  <style:style style:name="Heading_20_6" style:display-name="Heading 6" style:family="paragraph" style:parent-style-name="Heading" style:next-style-name="Text_20_body" style:default-outline-level="" style:list-style-name="" style:class="text"/>
  <style:style style:name="Heading_20_7" style:display-name="Heading 7" style:family="paragraph" style:parent-style-name="Heading" style:next-style-name="Text_20_body" style:default-outline-level="" style:list-style-name="" style:class="text"/>
  <style:style style:name="Heading_20_8" style:display-name="Heading 8" style:family="paragraph" style:parent-style-name="Heading" style:next-style-name="Text_20_body" style:default-outline-level="" style:list-style-name="" style:class="text"/>
  <style:style style:name="Heading_20_9" style:display-name="Heading 9" style:family="paragraph" style:parent-style-name="Heading" style:next-style-name="Text_20_body" style:default-outline-level="" style:list-style-name="" style:class="text"/>
  <style:style style:name="Heading_20_10" style:display-name="Heading 10" style:family="paragraph" style:parent-style-name="Heading" style:next-style-name="Text_20_body" style:default-outline-level="10" style:list-style-name="" style:class="text">
   <style:text-properties fo:font-size="75%" fo:font-weight="bold"/>
  </style:style>
  <style:style style:name="List" style:family="paragraph" style:parent-style-name="Text_20_body" style:class="list"/>
  <style:style style:name="Numbering_20_1_20_Start" style:display-name="Numbering 1 Start" style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="0.499cm" fo:margin-right="0cm" fo:margin-top="0.423cm" fo:margin-bottom="0.212cm" fo:text-indent="-0.499cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="Numbering_20_1" style:display-name="Numbering 1" style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="0.499cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0.212cm" fo:text-indent="-0.499cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="Numbering_20_1_20_End" style:display-name="Numbering 1 End" style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="0.499cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0.423cm" fo:text-indent="-0.499cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="Numbering_20_1_20_Cont." style:display-name="Numbering 1 Cont." style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="0.499cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0.212cm" fo:text-indent="0cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="Numbering_20_2_20_Start" style:display-name="Numbering 2 Start" style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="0.998cm" fo:margin-right="0cm" fo:margin-top="0.423cm" fo:margin-bottom="0.212cm" fo:text-indent="-0.499cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="Numbering_20_2" style:display-name="Numbering 2" style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="0.998cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0.212cm" fo:text-indent="-0.499cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="Numbering_20_2_20_End" style:display-name="Numbering 2 End" style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="0.998cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0.423cm" fo:text-indent="-0.499cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="Numbering_20_2_20_Cont." style:display-name="Numbering 2 Cont." style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="0.998cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0.212cm" fo:text-indent="0cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="Numbering_20_3_20_Start" style:display-name="Numbering 3 Start" style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="1.498cm" fo:margin-right="0cm" fo:margin-top="0.423cm" fo:margin-bottom="0.212cm" fo:text-indent="-0.499cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="Numbering_20_3" style:display-name="Numbering 3" style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="1.498cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0.212cm" fo:text-indent="-0.499cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="Numbering_20_3_20_End" style:display-name="Numbering 3 End" style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="1.498cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0.423cm" fo:text-indent="-0.499cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="Numbering_20_3_20_Cont." style:display-name="Numbering 3 Cont." style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="1.498cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0.212cm" fo:text-indent="0cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="Numbering_20_4_20_Start" style:display-name="Numbering 4 Start" style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="1.997cm" fo:margin-right="0cm" fo:margin-top="0.423cm" fo:margin-bottom="0.212cm" fo:text-indent="-0.499cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="Numbering_20_4" style:display-name="Numbering 4" style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="1.997cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0.212cm" fo:text-indent="-0.499cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="Numbering_20_4_20_End" style:display-name="Numbering 4 End" style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="1.997cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0.423cm" fo:text-indent="-0.499cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="Numbering_20_4_20_Cont." style:display-name="Numbering 4 Cont." style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="1.997cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0.212cm" fo:text-indent="0cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="Numbering_20_5_20_Start" style:display-name="Numbering 5 Start" style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="2.496cm" fo:margin-right="0cm" fo:margin-top="0.423cm" fo:margin-bottom="0.212cm" fo:text-indent="-0.499cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="Numbering_20_5" style:display-name="Numbering 5" style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="2.496cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0.212cm" fo:text-indent="-0.499cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="Numbering_20_5_20_End" style:display-name="Numbering 5 End" style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="2.496cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0.423cm" fo:text-indent="-0.499cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="Numbering_20_5_20_Cont." style:display-name="Numbering 5 Cont." style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="2.496cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0.212cm" fo:text-indent="0cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="List_20_1_20_Start" style:display-name="List 1 Start" style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="0.499cm" fo:margin-right="0cm" fo:margin-top="0.423cm" fo:margin-bottom="0.212cm" fo:text-indent="-0.499cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="List_20_1" style:display-name="List 1" style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="0.499cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0.212cm" fo:text-indent="-0.499cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="List_20_1_20_End" style:display-name="List 1 End" style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="0.499cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0.423cm" fo:text-indent="-0.499cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="List_20_1_20_Cont." style:display-name="List 1 Cont." style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="0.499cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0.212cm" fo:text-indent="0cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="List_20_2_20_Start" style:display-name="List 2 Start" style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="0.998cm" fo:margin-right="0cm" fo:margin-top="0.423cm" fo:margin-bottom="0.212cm" fo:text-indent="-0.499cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="List_20_2" style:display-name="List 2" style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="0.998cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0.212cm" fo:text-indent="-0.499cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="List_20_2_20_End" style:display-name="List 2 End" style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="0.998cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0.423cm" fo:text-indent="-0.499cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="List_20_2_20_Cont." style:display-name="List 2 Cont." style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="0.998cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0.212cm" fo:text-indent="0cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="List_20_3_20_Start" style:display-name="List 3 Start" style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="1.498cm" fo:margin-right="0cm" fo:margin-top="0.423cm" fo:margin-bottom="0.212cm" fo:text-indent="-0.499cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="List_20_3" style:display-name="List 3" style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="1.498cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0.212cm" fo:text-indent="-0.499cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="List_20_3_20_End" style:display-name="List 3 End" style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="1.498cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0.423cm" fo:text-indent="-0.499cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="List_20_3_20_Cont." style:display-name="List 3 Cont." style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="1.498cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0.212cm" fo:text-indent="0cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="List_20_4_20_Start" style:display-name="List 4 Start" style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="1.997cm" fo:margin-right="0cm" fo:margin-top="0.423cm" fo:margin-bottom="0.212cm" fo:text-indent="-0.499cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="List_20_4" style:display-name="List 4" style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="1.997cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0.212cm" fo:text-indent="-0.499cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="List_20_4_20_End" style:display-name="List 4 End" style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="1.997cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0.423cm" fo:text-indent="-0.499cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="List_20_4_20_Cont." style:display-name="List 4 Cont." style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="1.997cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0.212cm" fo:text-indent="0cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="List_20_5_20_Start" style:display-name="List 5 Start" style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="2.496cm" fo:margin-right="0cm" fo:margin-top="0.423cm" fo:margin-bottom="0.212cm" fo:text-indent="-0.499cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="List_20_5" style:display-name="List 5" style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="2.496cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0.212cm" fo:text-indent="-0.499cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="List_20_5_20_End" style:display-name="List 5 End" style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="2.496cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0.423cm" fo:text-indent="-0.499cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="List_20_5_20_Cont." style:display-name="List 5 Cont." style:family="paragraph" style:parent-style-name="List" style:class="list">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="2.496cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0.212cm" fo:text-indent="0cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="Header_20_and_20_Footer" style:display-name="Header and Footer" style:family="paragraph" style:parent-style-name="Standard" style:class="extra">
   <style:paragraph-properties text:number-lines="false" text:line-number="0">
    <style:tab-stops>
     <style:tab-stop style:position="8.5cm" style:type="center"/>
     <style:tab-stop style:position="17cm" style:type="right"/>
    </style:tab-stops>
   </style:paragraph-properties>
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
  <style:style style:name="Footer" style:family="paragraph" style:parent-style-name="Standard" style:class="extra" style:master-page-name="">
   <style:paragraph-properties fo:text-align="center" style:justify-single-word="false" style:page-number="auto" text:number-lines="false" text:line-number="0">
    <style:tab-stops>
     <style:tab-stop style:position="8.5cm" style:type="center"/>
     <style:tab-stop style:position="17.002cm" style:type="right"/>
    </style:tab-stops>
   </style:paragraph-properties>
   <style:text-properties fo:font-size="11pt"/>
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
  <style:style style:name="Table_20_Contents" style:display-name="Table Contents" style:family="paragraph" style:parent-style-name="Text_20_body" style:class="extra"/>
  <style:style style:name="Table_20_Heading" style:display-name="Table Heading" style:family="paragraph" style:parent-style-name="Table_20_Contents" style:class="extra">
   <style:paragraph-properties fo:text-align="center" style:justify-single-word="false"/>
   <style:text-properties fo:font-style="italic" fo:font-weight="bold"/>
  </style:style>
  <style:style style:name="Caption" style:family="paragraph" style:parent-style-name="Standard" style:class="extra">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin-top="0.212cm" fo:margin-bottom="0.212cm"/>
  </style:style>
  <style:style style:name="Illustration" style:family="paragraph" style:parent-style-name="Caption" style:class="extra"/>
  <style:style style:name="Table" style:family="paragraph" style:parent-style-name="Caption" style:class="extra"/>
  <style:style style:name="Text" style:family="paragraph" style:parent-style-name="Caption" style:class="extra" style:master-page-name="">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin-top="0.21cm" fo:margin-bottom="0.21cm" style:page-number="auto"/>
  </style:style>
  <style:style style:name="Frame_20_contents" style:display-name="Frame contents" style:family="paragraph" style:parent-style-name="Text_20_body" style:class="extra"/>
  <style:style style:name="Footnote" style:family="paragraph" style:parent-style-name="Standard" style:class="extra">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="0.499cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="-0.499cm" style:auto-text-indent="false"/>
   <style:text-properties fo:font-size="10pt"/>
  </style:style>
  <style:style style:name="Addressee" style:family="paragraph" style:parent-style-name="Standard" style:class="extra">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin-top="0cm" fo:margin-bottom="0.106cm"/>
  </style:style>
  <style:style style:name="Sender" style:family="paragraph" style:parent-style-name="Standard" style:class="extra">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin-top="0cm" fo:margin-bottom="0.106cm" fo:line-height="100%" text:number-lines="false" text:line-number="0"/>
  </style:style>
  <style:style style:name="Endnote" style:family="paragraph" style:parent-style-name="Standard" style:class="extra">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="0.499cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="-0.499cm" style:auto-text-indent="false" text:number-lines="false" text:line-number="0"/>
   <style:text-properties fo:font-size="10pt"/>
  </style:style>
  <style:style style:name="Drawing" style:family="paragraph" style:parent-style-name="Caption" style:class="extra"/>
  <style:style style:name="Index" style:family="paragraph" style:parent-style-name="Standard" style:class="index"/>
  <style:style style:name="Index_20_Heading" style:display-name="Index Heading" style:family="paragraph" style:parent-style-name="Heading" style:class="index">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="0cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false"/>
   <style:text-properties fo:font-size="16pt" fo:font-weight="bold"/>
  </style:style>
  <style:style style:name="Index_20_1" style:display-name="Index 1" style:family="paragraph" style:parent-style-name="Index" style:class="index">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="0.499cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="Index_20_2" style:display-name="Index 2" style:family="paragraph" style:parent-style-name="Index" style:class="index">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="0.499cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="Index_20_3" style:display-name="Index 3" style:family="paragraph" style:parent-style-name="Index" style:class="index">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="0cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="Index_20_Separator" style:display-name="Index Separator" style:family="paragraph" style:parent-style-name="Index" style:class="index">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="0cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="Contents_20_Heading" style:display-name="Contents Heading" style:family="paragraph" style:parent-style-name="Heading" style:class="index">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="0cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false"/>
   <style:text-properties fo:font-size="16pt" fo:font-weight="bold"/>
  </style:style>
  <style:style style:name="Contents_20_1" style:display-name="Contents 1" style:family="paragraph" style:parent-style-name="Index" style:class="index">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="0cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false">
    <style:tab-stops>
     <style:tab-stop style:position="17.002cm" style:type="right" style:leader-style="dotted" style:leader-text="."/>
    </style:tab-stops>
   </style:paragraph-properties>
  </style:style>
  <style:style style:name="Contents_20_2" style:display-name="Contents 2" style:family="paragraph" style:parent-style-name="Index" style:class="index">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="0.499cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false">
    <style:tab-stops>
     <style:tab-stop style:position="16.503cm" style:type="right" style:leader-style="dotted" style:leader-text="."/>
    </style:tab-stops>
   </style:paragraph-properties>
  </style:style>
  <style:style style:name="Contents_20_3" style:display-name="Contents 3" style:family="paragraph" style:parent-style-name="Index" style:class="index">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="0.998cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false">
    <style:tab-stops>
     <style:tab-stop style:position="16.004cm" style:type="right" style:leader-style="dotted" style:leader-text="."/>
    </style:tab-stops>
   </style:paragraph-properties>
  </style:style>
  <style:style style:name="Contents_20_4" style:display-name="Contents 4" style:family="paragraph" style:parent-style-name="Index" style:class="index">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="1.498cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false">
    <style:tab-stops>
     <style:tab-stop style:position="15.505cm" style:type="right" style:leader-style="dotted" style:leader-text="."/>
    </style:tab-stops>
   </style:paragraph-properties>
  </style:style>
  <style:style style:name="Contents_20_5" style:display-name="Contents 5" style:family="paragraph" style:parent-style-name="Index" style:class="index">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="1.997cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false">
    <style:tab-stops>
     <style:tab-stop style:position="15.005cm" style:type="right" style:leader-style="dotted" style:leader-text="."/>
    </style:tab-stops>
   </style:paragraph-properties>
  </style:style>
  <style:style style:name="User_20_Index_20_Heading" style:display-name="User Index Heading" style:family="paragraph" style:parent-style-name="Heading" style:class="index">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="0cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false"/>
   <style:text-properties fo:font-size="16pt" fo:font-weight="bold"/>
  </style:style>
  <style:style style:name="User_20_Index_20_1" style:display-name="User Index 1" style:family="paragraph" style:parent-style-name="Index" style:class="index">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="0cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false">
    <style:tab-stops>
     <style:tab-stop style:position="17.002cm" style:type="right" style:leader-style="dotted" style:leader-text="."/>
    </style:tab-stops>
   </style:paragraph-properties>
  </style:style>
  <style:style style:name="User_20_Index_20_2" style:display-name="User Index 2" style:family="paragraph" style:parent-style-name="Index" style:class="index">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="0.499cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false">
    <style:tab-stops>
     <style:tab-stop style:position="16.503cm" style:type="right" style:leader-style="dotted" style:leader-text="."/>
    </style:tab-stops>
   </style:paragraph-properties>
  </style:style>
  <style:style style:name="User_20_Index_20_3" style:display-name="User Index 3" style:family="paragraph" style:parent-style-name="Index" style:class="index">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="0.998cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false">
    <style:tab-stops>
     <style:tab-stop style:position="16.004cm" style:type="right" style:leader-style="dotted" style:leader-text="."/>
    </style:tab-stops>
   </style:paragraph-properties>
  </style:style>
  <style:style style:name="User_20_Index_20_4" style:display-name="User Index 4" style:family="paragraph" style:parent-style-name="Index" style:class="index">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="1.498cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false">
    <style:tab-stops>
     <style:tab-stop style:position="15.505cm" style:type="right" style:leader-style="dotted" style:leader-text="."/>
    </style:tab-stops>
   </style:paragraph-properties>
  </style:style>
  <style:style style:name="User_20_Index_20_5" style:display-name="User Index 5" style:family="paragraph" style:parent-style-name="Index" style:class="index">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="1.997cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false">
    <style:tab-stops>
     <style:tab-stop style:position="15.005cm" style:type="right" style:leader-style="dotted" style:leader-text="."/>
    </style:tab-stops>
   </style:paragraph-properties>
  </style:style>
  <style:style style:name="Contents_20_6" style:display-name="Contents 6" style:family="paragraph" style:parent-style-name="Index" style:class="index">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="2.496cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false">
    <style:tab-stops>
     <style:tab-stop style:position="11.105cm" style:type="right" style:leader-style="dotted" style:leader-text="."/>
    </style:tab-stops>
   </style:paragraph-properties>
  </style:style>
  <style:style style:name="Contents_20_7" style:display-name="Contents 7" style:family="paragraph" style:parent-style-name="Index" style:class="index">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="2.995cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false">
    <style:tab-stops>
     <style:tab-stop style:position="10.606cm" style:type="right" style:leader-style="dotted" style:leader-text="."/>
    </style:tab-stops>
   </style:paragraph-properties>
  </style:style>
  <style:style style:name="Contents_20_8" style:display-name="Contents 8" style:family="paragraph" style:parent-style-name="Index" style:class="index">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="3.494cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false">
    <style:tab-stops>
     <style:tab-stop style:position="10.107cm" style:type="right" style:leader-style="dotted" style:leader-text="."/>
    </style:tab-stops>
   </style:paragraph-properties>
  </style:style>
  <style:style style:name="Contents_20_9" style:display-name="Contents 9" style:family="paragraph" style:parent-style-name="Index" style:class="index">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="3.993cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false">
    <style:tab-stops>
     <style:tab-stop style:position="9.608cm" style:type="right" style:leader-style="dotted" style:leader-text="."/>
    </style:tab-stops>
   </style:paragraph-properties>
  </style:style>
  <style:style style:name="Contents_20_10" style:display-name="Contents 10" style:family="paragraph" style:parent-style-name="Index" style:class="index">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="4.493cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false">
    <style:tab-stops>
     <style:tab-stop style:position="9.109cm" style:type="right" style:leader-style="dotted" style:leader-text="."/>
    </style:tab-stops>
   </style:paragraph-properties>
  </style:style>
  <style:style style:name="Illustration_20_Index_20_Heading" style:display-name="Illustration Index Heading" style:family="paragraph" style:parent-style-name="Heading" style:class="index">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="0cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false" text:number-lines="false" text:line-number="0"/>
   <style:text-properties fo:font-size="16pt" fo:font-weight="bold"/>
  </style:style>
  <style:style style:name="Illustration_20_Index_20_1" style:display-name="Illustration Index 1" style:family="paragraph" style:parent-style-name="Index" style:class="index">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="0cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false">
    <style:tab-stops>
     <style:tab-stop style:position="13.601cm" style:type="right" style:leader-style="dotted" style:leader-text="."/>
    </style:tab-stops>
   </style:paragraph-properties>
  </style:style>
  <style:style style:name="Object_20_index_20_heading" style:display-name="Object index heading" style:family="paragraph" style:parent-style-name="Heading" style:class="index">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="0cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false" text:number-lines="false" text:line-number="0"/>
   <style:text-properties fo:font-size="16pt" fo:font-weight="bold"/>
  </style:style>
  <style:style style:name="Object_20_index_20_1" style:display-name="Object index 1" style:family="paragraph" style:parent-style-name="Index" style:class="index">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="0cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false">
    <style:tab-stops>
     <style:tab-stop style:position="13.601cm" style:type="right" style:leader-style="dotted" style:leader-text="."/>
    </style:tab-stops>
   </style:paragraph-properties>
  </style:style>
  <style:style style:name="Table_20_index_20_heading" style:display-name="Table index heading" style:family="paragraph" style:parent-style-name="Heading" style:class="index">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="0cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false" text:number-lines="false" text:line-number="0"/>
   <style:text-properties fo:font-size="16pt" fo:font-weight="bold"/>
  </style:style>
  <style:style style:name="Table_20_index_20_1" style:display-name="Table index 1" style:family="paragraph" style:parent-style-name="Index" style:class="index">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="0cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false">
    <style:tab-stops>
     <style:tab-stop style:position="13.601cm" style:type="right" style:leader-style="dotted" style:leader-text="."/>
    </style:tab-stops>
   </style:paragraph-properties>
  </style:style>
  <style:style style:name="Bibliography_20_Heading" style:display-name="Bibliography Heading" style:family="paragraph" style:parent-style-name="Heading" style:class="index">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="0cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false" text:number-lines="false" text:line-number="0"/>
   <style:text-properties fo:font-size="16pt" fo:font-weight="bold"/>
  </style:style>
  <style:style style:name="Bibliography_20_1" style:display-name="Bibliography 1" style:family="paragraph" style:parent-style-name="Index" style:class="index">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="0cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false">
    <style:tab-stops>
     <style:tab-stop style:position="13.601cm" style:type="right" style:leader-style="dotted" style:leader-text="."/>
    </style:tab-stops>
   </style:paragraph-properties>
  </style:style>
  <style:style style:name="User_20_Index_20_6" style:display-name="User Index 6" style:family="paragraph" style:parent-style-name="Index" style:class="index">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="2.496cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false">
    <style:tab-stops>
     <style:tab-stop style:position="11.105cm" style:type="right" style:leader-style="dotted" style:leader-text="."/>
    </style:tab-stops>
   </style:paragraph-properties>
  </style:style>
  <style:style style:name="User_20_Index_20_7" style:display-name="User Index 7" style:family="paragraph" style:parent-style-name="Index" style:class="index">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="2.995cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false">
    <style:tab-stops>
     <style:tab-stop style:position="10.606cm" style:type="right" style:leader-style="dotted" style:leader-text="."/>
    </style:tab-stops>
   </style:paragraph-properties>
  </style:style>
  <style:style style:name="User_20_Index_20_8" style:display-name="User Index 8" style:family="paragraph" style:parent-style-name="Index" style:class="index">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="3.494cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false">
    <style:tab-stops>
     <style:tab-stop style:position="10.107cm" style:type="right" style:leader-style="dotted" style:leader-text="."/>
    </style:tab-stops>
   </style:paragraph-properties>
  </style:style>
  <style:style style:name="User_20_Index_20_9" style:display-name="User Index 9" style:family="paragraph" style:parent-style-name="Index" style:class="index">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="3.993cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false">
    <style:tab-stops>
     <style:tab-stop style:position="9.608cm" style:type="right" style:leader-style="dotted" style:leader-text="."/>
    </style:tab-stops>
   </style:paragraph-properties>
  </style:style>
  <style:style style:name="User_20_Index_20_10" style:display-name="User Index 10" style:family="paragraph" style:parent-style-name="Index" style:class="index">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="4.493cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false">
    <style:tab-stops>
     <style:tab-stop style:position="9.109cm" style:type="right" style:leader-style="dotted" style:leader-text="."/>
    </style:tab-stops>
   </style:paragraph-properties>
  </style:style>
  <style:style style:name="Title" style:family="paragraph" style:parent-style-name="Standard" style:next-style-name="Subtitle" style:class="chapter" style:master-page-name="">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="0cm" fo:margin-right="0cm" fo:margin-top="0.000cm" fo:margin-bottom="0cm" fo:line-height="200%" fo:text-align="center" style:justify-single-word="false" fo:text-indent="0cm" style:auto-text-indent="false" style:page-number="auto" fo:background-color="transparent" fo:padding="0cm" fo:border="none" text:number-lines="false" text:line-number="0">
    <style:tab-stops/>
    <style:background-image/>
   </style:paragraph-properties>
   <style:text-properties fo:text-transform="uppercase" fo:font-weight="normal" style:letter-kerning="false"/>
  </style:style>
  <style:style style:name="Subtitle" style:family="paragraph" style:parent-style-name="Title" style:class="chapter" style:master-page-name="">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin-top="0cm" fo:margin-bottom="0cm" style:page-number="auto"/>
   <style:text-properties fo:font-variant="normal" fo:text-transform="none" fo:letter-spacing="normal" fo:font-style="italic" fo:font-weight="normal"/>
  </style:style>
  <style:style style:name="Quotations" style:family="paragraph" style:parent-style-name="Text_20_body" style:class="html" style:master-page-name="">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="0.499cm" fo:margin-right="0.499cm" fo:margin-top="0cm" fo:margin-bottom="0.499cm" fo:text-indent="0cm" style:auto-text-indent="false" style:page-number="auto"/>
  </style:style>
  <style:style style:name="Preformatted_20_Text" style:display-name="Preformatted Text" style:family="paragraph" style:parent-style-name="Standard" style:class="html">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin-top="0cm" fo:margin-bottom="0cm"/>
  </style:style>
  <style:style style:name="Horizontal_20_Line" style:display-name="Horizontal Line" style:family="paragraph" style:parent-style-name="Standard" style:next-style-name="Text_20_body" style:class="html">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin-top="0cm" fo:margin-bottom="0.499cm" style:border-line-width-bottom="0.002cm 0.035cm 0.002cm" fo:padding="0cm" fo:border-left="none" fo:border-right="none" fo:border-top="none" fo:border-bottom="0.039cm double #808080" text:number-lines="false" text:line-number="0"/>
   <style:text-properties fo:font-size="6pt"/>
  </style:style>
  <style:style style:name="List_20_Contents" style:display-name="List Contents" style:family="paragraph" style:parent-style-name="Standard" style:class="html">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="1cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false"/>
  </style:style>
  <style:style style:name="List_20_Heading" style:display-name="List Heading" style:family="paragraph" style:parent-style-name="Standard" style:next-style-name="List_20_Contents" style:class="html">
   <style:paragraph-properties loext:contextual-spacing="false" fo:margin="100%" fo:margin-left="0cm" fo:margin-right="0cm" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:text-indent="0cm" style:auto-text-indent="false"/>
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
  <style:style style:name="Footnote_20_Symbol" style:display-name="Footnote Symbol" style:family="text"/>
  <style:style style:name="Page_20_Number" style:display-name="Page Number" style:family="text">
   <style:text-properties style:font-name="ITC Officina Sans Book" fo:font-size="7pt" fo:letter-spacing="0.071cm" fo:font-weight="bold"/>
  </style:style>
  <style:style style:name="Caption_20_characters" style:display-name="Caption characters" style:family="text"/>
  <style:style style:name="Drop_20_Caps" style:display-name="Drop Caps" style:family="text"/>
  <style:style style:name="Numbering_20_Symbols" style:display-name="Numbering Symbols" style:family="text"/>
  <style:style style:name="Bullet_20_Symbols" style:display-name="Bullet Symbols" style:family="text">
   <style:text-properties style:font-name="StarSymbol" fo:font-size="9pt"/>
  </style:style>
  <style:style style:name="Internet_20_link" style:display-name="Internet link" style:family="text">
   <style:text-properties fo:color="#000080" style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color"/>
  </style:style>
  <style:style style:name="Visited_20_Internet_20_Link" style:display-name="Visited Internet Link" style:family="text">
   <style:text-properties fo:color="#800000" style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color"/>
  </style:style>
  <style:style style:name="Placeholder" style:family="text">
   <style:text-properties fo:font-variant="small-caps" fo:color="#008080" style:text-underline-style="dotted" style:text-underline-width="auto" style:text-underline-color="font-color"/>
  </style:style>
  <style:style style:name="Index_20_Link" style:display-name="Index Link" style:family="text"/>
  <style:style style:name="Endnote_20_Symbol" style:display-name="Endnote Symbol" style:family="text"/>
  <style:style style:name="Line_20_numbering" style:display-name="Line numbering" style:family="text">
   <style:text-properties style:font-name="Courier New" fo:font-size="8pt"/>
  </style:style>
  <style:style style:name="Main_20_index_20_entry" style:display-name="Main index entry" style:family="text">
   <style:text-properties fo:font-weight="bold" style:font-weight-asian="bold" style:font-weight-complex="bold"/>
  </style:style>
  <style:style style:name="Footnote_20_anchor" style:display-name="Footnote anchor" style:family="text">
   <style:text-properties style:text-position="super 58%"/>
  </style:style>
  <style:style style:name="Endnote_20_anchor" style:display-name="Endnote anchor" style:family="text">
   <style:text-properties style:text-position="super 58%"/>
  </style:style>
  <style:style style:name="Rubies" style:family="text">
   <style:text-properties fo:font-size="6pt" style:font-size-asian="6pt" style:font-size-complex="6pt"/>
  </style:style>
  <style:style style:name="Emphasis" style:family="text">
   <style:text-properties fo:font-style="italic" fo:background-color="transparent"/>
  </style:style>
  <style:style style:name="Citation" style:family="text">
   <style:text-properties fo:font-style="italic"/>
  </style:style>
  <style:style style:name="Strong_20_Emphasis" style:display-name="Strong Emphasis" style:family="text">
   <style:text-properties fo:text-transform="uppercase"/>
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
  <style:style style:name="Definition" style:family="text"/>
  <style:style style:name="Teletype" style:family="text">
   <style:text-properties style:font-name="Cumberland" style:font-name-asian="Cumberland" style:font-name-complex="Cumberland"/>
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
   <text:outline-level-style text:level="3" text:style-name="Zeichenformat" style:num-format="">
    <style:list-level-properties/>
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
  <text:notes-configuration text:note-class="footnote" text:citation-style-name="Footnote_20_Symbol" text:citation-body-style-name="Footnote_20_anchor" style:num-format="1" text:start-value="0" text:footnotes-position="page" text:start-numbering-at="page"/>
  <text:notes-configuration text:note-class="endnote" text:citation-style-name="Endnote_20_Symbol" text:citation-body-style-name="Endnote_20_anchor" text:master-page-name="Endnote" style:num-format="1" text:start-value="0"/>
  <text:linenumbering-configuration text:style-name="Line_20_numbering" text:number-lines="false" text:offset="0.499cm" style:num-format="1" text:number-position="left" text:increment="5"/>
 </office:styles>
 <office:automatic-styles>
  <style:page-layout style:name="Mpm1">
   <style:page-layout-properties fo:page-width="21.001cm" fo:page-height="29.7cm" style:num-format="1" style:paper-tray-name="[From printer settings]" style:print-orientation="portrait" fo:margin-top="3.2cm" fo:margin-bottom="2.499cm" fo:margin-left="2.701cm" fo:margin-right="3cm" style:writing-mode="lr-tb" style:layout-grid-color="#c0c0c0" style:layout-grid-lines="20" style:layout-grid-base-height="0.706cm" style:layout-grid-ruby-height="0.353cm" style:layout-grid-mode="none" style:layout-grid-ruby-below="false" style:layout-grid-print="false" style:layout-grid-display="false" style:footnote-max-height="0cm">
    <style:columns fo:column-count="1" fo:column-gap="0cm"/>
    <style:footnote-sep style:width="0.018cm" style:distance-before-sep="0.101cm" style:distance-after-sep="0.101cm" style:adjustment="left" style:rel-width="25%" style:color="#000000"/>
   </style:page-layout-properties>
   <style:header-style/>
   <style:footer-style>
    <style:header-footer-properties fo:min-height="1.699cm" fo:margin-left="0cm" fo:margin-right="0cm" fo:margin-top="1.199cm" style:shadow="none" style:dynamic-spacing="false"/>
   </style:footer-style>
  </style:page-layout>
  <style:page-layout style:name="Mpm2">
   <style:page-layout-properties fo:page-width="21.001cm" fo:page-height="29.7cm" style:num-format="1" style:print-orientation="portrait" fo:margin-top="2cm" fo:margin-bottom="2cm" fo:margin-left="2.499cm" fo:margin-right="2.499cm" style:shadow="none" fo:background-color="transparent" style:writing-mode="lr-tb" style:layout-grid-color="#c0c0c0" style:layout-grid-lines="20" style:layout-grid-base-height="0.706cm" style:layout-grid-ruby-height="0.353cm" style:layout-grid-mode="none" style:layout-grid-ruby-below="false" style:layout-grid-print="false" style:layout-grid-display="false" style:footnote-max-height="0cm">
    <style:background-image/>
    <style:columns fo:column-count="1" fo:column-gap="0cm"/>
    <style:footnote-sep style:width="0.018cm" style:distance-before-sep="0.101cm" style:distance-after-sep="0.101cm" style:adjustment="left" style:rel-width="25%" style:color="#000000"/>
   </style:page-layout-properties>
   <style:header-style/>
   <style:footer-style/>
  </style:page-layout>
  <style:page-layout style:name="Mpm3" style:page-usage="left">
   <style:page-layout-properties fo:page-width="21.001cm" fo:page-height="29.7cm" style:num-format="1" style:print-orientation="portrait" fo:margin-top="2cm" fo:margin-bottom="1cm" fo:margin-left="2.499cm" fo:margin-right="4.5cm" style:writing-mode="lr-tb" style:layout-grid-color="#c0c0c0" style:layout-grid-lines="20" style:layout-grid-base-height="0.706cm" style:layout-grid-ruby-height="0.353cm" style:layout-grid-mode="none" style:layout-grid-ruby-below="false" style:layout-grid-print="false" style:layout-grid-display="false" style:footnote-max-height="0cm">
    <style:footnote-sep style:width="0.018cm" style:distance-before-sep="0.101cm" style:distance-after-sep="0.101cm" style:adjustment="left" style:rel-width="25%" style:color="#000000"/>
   </style:page-layout-properties>
   <style:header-style/>
   <style:footer-style/>
  </style:page-layout>
  <style:page-layout style:name="Mpm4" style:page-usage="right">
   <style:page-layout-properties fo:page-width="21.001cm" fo:page-height="29.7cm" style:num-format="1" style:print-orientation="portrait" fo:margin-top="2cm" fo:margin-bottom="1cm" fo:margin-left="2.499cm" fo:margin-right="4.5cm" style:writing-mode="lr-tb" style:layout-grid-color="#c0c0c0" style:layout-grid-lines="20" style:layout-grid-base-height="0.706cm" style:layout-grid-ruby-height="0.353cm" style:layout-grid-mode="none" style:layout-grid-ruby-below="false" style:layout-grid-print="false" style:layout-grid-display="false" style:footnote-max-height="0cm">
    <style:footnote-sep style:width="0.018cm" style:distance-before-sep="0.101cm" style:distance-after-sep="0.101cm" style:adjustment="left" style:rel-width="25%" style:color="#000000"/>
   </style:page-layout-properties>
   <style:header-style/>
   <style:footer-style/>
  </style:page-layout>
  <style:page-layout style:name="Mpm5">
   <style:page-layout-properties fo:page-width="22.721cm" fo:page-height="11.4cm" style:num-format="1" style:print-orientation="landscape" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:margin-left="0cm" fo:margin-right="0cm" style:writing-mode="lr-tb" style:layout-grid-color="#c0c0c0" style:layout-grid-lines="20" style:layout-grid-base-height="0.706cm" style:layout-grid-ruby-height="0.353cm" style:layout-grid-mode="none" style:layout-grid-ruby-below="false" style:layout-grid-print="false" style:layout-grid-display="false" style:footnote-max-height="0cm">
    <style:footnote-sep style:width="0.018cm" style:distance-before-sep="0.101cm" style:distance-after-sep="0.101cm" style:adjustment="left" style:rel-width="25%" style:color="#000000"/>
   </style:page-layout-properties>
   <style:header-style/>
   <style:footer-style/>
  </style:page-layout>
  <style:page-layout style:name="Mpm6">
   <style:page-layout-properties fo:page-width="14.801cm" fo:page-height="21.001cm" style:num-format="1" style:print-orientation="portrait" fo:margin-top="2cm" fo:margin-bottom="2cm" fo:margin-left="2cm" fo:margin-right="2cm" style:writing-mode="lr-tb" style:layout-grid-color="#c0c0c0" style:layout-grid-lines="20" style:layout-grid-base-height="0.706cm" style:layout-grid-ruby-height="0.353cm" style:layout-grid-mode="none" style:layout-grid-ruby-below="false" style:layout-grid-print="false" style:layout-grid-display="false" style:footnote-max-height="0cm">
    <style:footnote-sep style:width="0.018cm" style:distance-before-sep="0.101cm" style:distance-after-sep="0.101cm" style:adjustment="left" style:rel-width="25%" style:color="#000000"/>
   </style:page-layout-properties>
   <style:header-style/>
   <style:footer-style/>
  </style:page-layout>
  <style:page-layout style:name="Mpm7">
   <style:page-layout-properties fo:page-width="20.999cm" fo:page-height="29.699cm" style:num-format="1" style:print-orientation="portrait" fo:margin-top="2cm" fo:margin-bottom="2cm" fo:margin-left="2cm" fo:margin-right="2cm" style:writing-mode="lr-tb" style:layout-grid-color="#c0c0c0" style:layout-grid-lines="20" style:layout-grid-base-height="0.706cm" style:layout-grid-ruby-height="0.353cm" style:layout-grid-mode="none" style:layout-grid-ruby-below="false" style:layout-grid-print="false" style:layout-grid-display="false" style:footnote-max-height="0cm">
    <style:footnote-sep style:adjustment="left" style:rel-width="25%" style:color="#000000"/>
   </style:page-layout-properties>
   <style:header-style/>
   <style:footer-style/>
  </style:page-layout>
 </office:automatic-styles>
 <office:master-styles>
  <style:master-page style:name="Standard" style:page-layout-name="Mpm1">
   <style:footer>
    <text:p text:style-name="Footer"><text:page-number text:select-page="current"/></text:p>
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

    def _set_up(self):
        """Helper method for ZIP file generation.

        Add rdf manifest to the temporary directory containing the internal structure of an ODF file.
        Return a message beginning with the ERROR constant in case of error.
        Extends the superclass method.
        """

        # Generate the common ODF components.

        message = super()._set_up()
        if message.startswith(ERROR):
            return message

        # Generate manifest.rdf

        try:
            with open(f'{self._tempDir}/manifest.rdf', 'w', encoding='utf-8') as f:
                f.write(self._MANIFEST_RDF)
        except:
            return f'{ERROR}Cannot write "manifest.rdf"'

        return 'ODT structure generated.'

    def _convert_from_yw(self, text, quick=False):
        """Return text, converted from yw7 markup to target format.
        
        Positional arguments:
            text -- string to convert.
        
        Optional arguments:
            quick -- bool: if True, apply a conversion mode for one-liners without formatting.
        
        Overrides the superclass method.
        """
        if quick:            
            # Just clean up a one-liner without sophisticated formatting.
            try:
                return text.replace('&', '&amp;').replace('>', '&gt;').replace('<', '&lt;')
            
            except AttributeError:
                return ''

        ODT_REPLACEMENTS = [
            ('&', '&amp;'),
            ('>', '&gt;'),
            ('<', '&lt;'),
            ('\n\n', '</text:p>\r<text:p text:style-name="First_20_line_20_indent" />\r<text:p text:style-name="Text_20_body">'),
            ('\n', '</text:p>\r<text:p text:style-name="First_20_line_20_indent">'),
            ('\r', '\n'),
            ('[i]', '<text:span text:style-name="Emphasis">'),
            ('[/i]', '</text:span>'),
            ('[b]', '<text:span text:style-name="Strong_20_Emphasis">'),
            ('[/b]', '</text:span>'),
            ('/*', f'<office:annotation><dc:creator>{self.authorName}</dc:creator><text:p>'),
            ('*/', '</text:p></office:annotation>'),
        ]
        try:
            # process italics and bold markup reaching across linebreaks

            italics = False
            bold = False
            newlines = []
            lines = text.split('\n')
            for line in lines:
                if italics:
                    line = f'[i]{line}'
                    italics = False
                while line.count('[i]') > line.count('[/i]'):
                    line = f'{line}[/i]'
                    italics = True
                while line.count('[/i]') > line.count('[i]'):
                    line = f'[i]{line}'
                line = line.replace('[i][/i]', '')
                if bold:
                    line = f'[b]{line}'
                    bold = False
                while line.count('[b]') > line.count('[/b]'):
                    line = f'{line}[/b]'
                    bold = True
                while line.count('[/b]') > line.count('[b]'):
                    line = f'[b]{line}'
                line = line.replace('[b][/b]', '')
                newlines.append(line)
            text = '\n'.join(newlines).rstrip()
            for yw, od in ODT_REPLACEMENTS:
                text = text.replace(yw, od)

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

    _fileHeader = f'''{OdtFile._CONTENT_XML_HEADER}<text:p text:style-name="Title">$Title</text:p>
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

    _fileFooter = OdtFile._CONTENT_XML_FOOTER


class OdtManuscript(OdtFile):
    """ODT manuscript file representation.

    Export a manuscript with invisibly tagged chapters and scenes.
    """
    DESCRIPTION = 'Editable manuscript'
    SUFFIX = '_manuscript'

    _fileHeader = f'''{OdtFile._CONTENT_XML_HEADER}<text:p text:style-name="Title">$Title</text:p>
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
<text:p><text:a xlink:href="../${ProjectName}_scenes.odt#ScID:$ID%7Cregion">→Summary</text:a></text:p>
</office:annotation>$SceneContent</text:p>
</text:section>
'''

    _appendedSceneTemplate = '''<text:section text:style-name="Sect1" text:name="ScID:$ID">
<text:p text:style-name="First_20_line_20_indent"><office:annotation>
<dc:creator>scene title</dc:creator>
<text:p>~ ${Title} ~</text:p>
<text:p/>
<text:p><text:a xlink:href="../${ProjectName}_scenes.odt#ScID:$ID%7Cregion">→Summary</text:a></text:p>
</office:annotation>$SceneContent</text:p>
</text:section>
'''

    _sceneDivider = '<text:p text:style-name="Heading_20_4">* * *</text:p>\n'

    _chapterEndTemplate = '''</text:section>
'''

    _fileFooter = OdtFile._CONTENT_XML_FOOTER

    def _get_chapterMapping(self, chId, chapterNumber):
        """Return a mapping dictionary for a chapter section.
        
        Positional arguments:
            chId -- str: chapter ID.
            chapterNumber -- int: chapter number.
        
        Suppress the chapter title if necessary.
        Extends the superclass method.
        """
        chapterMapping = super()._get_chapterMapping(chId, chapterNumber)
        if self.chapters[chId].suppressChapterTitle:
            chapterMapping['Title'] = ''
        return chapterMapping


class OdtSceneDesc(OdtFile):
    """ODT scene summaries file representation.

    Export a full synopsis with invisibly tagged scene descriptions.
    """
    DESCRIPTION = 'Scene descriptions'
    SUFFIX = '_scenes'

    _fileHeader = f'''{OdtFile._CONTENT_XML_HEADER}<text:p text:style-name="Title">$Title</text:p>
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

    _fileFooter = OdtFile._CONTENT_XML_FOOTER


class OdtChapterDesc(OdtFile):
    """ODT chapter summaries file representation.

    Export a synopsis with invisibly tagged chapter descriptions.
    """
    DESCRIPTION = 'Chapter descriptions'
    SUFFIX = '_chapters'

    _fileHeader = f'''{OdtFile._CONTENT_XML_HEADER}<text:p text:style-name="Title">$Title</text:p>
<text:p text:style-name="Subtitle">$AuthorName</text:p>
'''

    _partTemplate = '''<text:h text:style-name="Heading_20_1" text:outline-level="1"><text:a xlink:href="../${ProjectName}_parts.odt#ChID:$ID%7Cregion">$Title</text:a></text:h>
'''

    _chapterTemplate = '''<text:section text:style-name="Sect1" text:name="ChID:$ID">
<text:h text:style-name="Heading_20_2" text:outline-level="2"><text:a xlink:href="../${ProjectName}_manuscript.odt#ChID:$ID%7Cregion">$Title</text:a></text:h>
<text:p text:style-name="Text_20_body">$Desc</text:p>
</text:section>
'''

    _fileFooter = OdtFile._CONTENT_XML_FOOTER


class OdtPartDesc(OdtFile):
    """ODT part summaries file representation.

    Export a synopsis with invisibly tagged part descriptions.
    """
    DESCRIPTION = 'Part descriptions'
    SUFFIX = '_parts'

    _fileHeader = f'''{OdtFile._CONTENT_XML_HEADER}<text:p text:style-name="Title">$Title</text:p>
<text:p text:style-name="Subtitle">$AuthorName</text:p>
'''

    _partTemplate = '''<text:section text:style-name="Sect1" text:name="ChID:$ID">
<text:h text:style-name="Heading_20_1" text:outline-level="1"><text:a xlink:href="../${ProjectName}_manuscript.odt#ChID:$ID%7Cregion">$Title</text:a></text:h>
<text:p text:style-name="Text_20_body">$Desc</text:p>
</text:section>
'''

    _fileFooter = OdtFile._CONTENT_XML_FOOTER


class OdtBriefSynopsis(OdtFile):
    """ODT brief synopsis file representation.

    Export a brief synopsis with chapter titles and scene titles.
    """
    DESCRIPTION = 'Brief synopsis'
    SUFFIX = '_brf_synopsis'

    _fileHeader = f'''{OdtFile._CONTENT_XML_HEADER}<text:p text:style-name="Title">$Title</text:p>
<text:p text:style-name="Subtitle">$AuthorName</text:p>
'''

    _partTemplate = '''<text:h text:style-name="Heading_20_1" text:outline-level="1">$Title</text:h>
'''

    _chapterTemplate = '''<text:h text:style-name="Heading_20_2" text:outline-level="2">$Title</text:h>
'''

    _sceneTemplate = '''<text:p text:style-name="Text_20_body">$Title</text:p>
'''

    _fileFooter = OdtFile._CONTENT_XML_FOOTER
from string import Template


class CrossReferences:
    """Dictionaries containing a novel's cross references.

    Public methods:
        generate_xref(novel) -- Generate cross references for a novel.

    Public instance variables:
        scnPerChr -- scenes per character.
        scnPerLoc -- scenes per location.
        scnPerItm -- scenes per item.
        scnPerTag -- scenes per tag.
        chrPerTag -- characters per tag.
        locPerTag -- locations per tag.
        itmPerTag -- items per tag.
        chpPerScn -- chapters per scene.
        srtScenes -- the novel's sorted scene IDs.
    """

    def __init__(self):
        """Initialize instance variables."""
        
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
        # Scene IDs in the overall order

    def generate_xref(self, novel):
        """Generate cross references for a novel.
        
        Positional argument:
            novel -- Novel instance to process.
        """
        self.scnPerChr = {}
        self.scnPerLoc = {}
        self.scnPerItm = {}
        self.scnPerTag = {}
        self.chrPerTag = {}
        self.locPerTag = {}
        self.itmPerTag = {}
        self.chpPerScn = {}
        self.srtScenes = []

        #--- Characters per tag.
        for crId in novel.srtCharacters:
            self.scnPerChr[crId] = []
            if novel.characters[crId].tags:
                for tag in novel.characters[crId].tags:
                    if not tag in self.chrPerTag:
                        self.chrPerTag[tag] = []
                    self.chrPerTag[tag].append(crId)

        #--- Locations per tag.
        for lcId in novel.srtLocations:
            self.scnPerLoc[lcId] = []
            if novel.locations[lcId].tags:
                for tag in novel.locations[lcId].tags:
                    if not tag in self.locPerTag:
                        self.locPerTag[tag] = []
                    self.locPerTag[tag].append(lcId)

        #--- Items per tag.
        for itId in novel.srtItems:
            self.scnPerItm[itId] = []
            if novel.items[itId].tags:
                for tag in novel.items[itId].tags:
                    if not tag in self.itmPerTag:
                        self.itmPerTag[tag] = []
                    self.itmPerTag[tag].append(itId)
                    
        #--- Process chapters and scenes.
        for chId in novel.srtChapters:

            for scId in novel.chapters[chId].srtScenes:
                self.srtScenes.append(scId)
                self.chpPerScn[scId] = chId

                #--- Scenes per character.
                if novel.scenes[scId].characters:
                    for crId in novel.scenes[scId].characters:
                        self.scnPerChr[crId].append(scId)

                #--- Scenes per location.
                if novel.scenes[scId].locations:
                    for lcId in novel.scenes[scId].locations:
                        self.scnPerLoc[lcId].append(scId)

                #--- Scenes per item.
                if novel.scenes[scId].items:
                    for itId in novel.scenes[scId].items:
                        self.scnPerItm[itId].append(scId)

                #--- Scenes per tag.
                if novel.scenes[scId].tags:
                    for tag in novel.scenes[scId].tags:
                        if not tag in self.scnPerTag:
                            self.scnPerTag[tag] = []
                        self.scnPerTag[tag].append(scId)


class OdtXref(OdtFile):
    """OpenDocument xml cross reference file representation."""
    DESCRIPTION = 'Cross reference'
    SUFFIX = '_xref'

    _fileHeader = f'''{OdtFile._CONTENT_XML_HEADER}<text:p text:style-name="Title">$Title</text:p>
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
    _fileFooter = OdtFile._CONTENT_XML_FOOTER

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
        chapterNumber = self.srtChapters.index(self._xr.chpPerScn[scId]) + 1
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
            if self.scenes[scId].isNotesScene:
                template = Template(self._notesSceneTemplate)
            elif self.scenes[scId].isTodoScene:
                template = Template(self._todoSceneTemplate)
            elif self.scenes[scId].isUnused:
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
        self._xr.generate_xref(self)
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


class OdsFile(OdfFile):
    """Generic OpenDocument spreadsheet document representation."""
    EXTENSION = '.ods'
    _ODF_COMPONENTS = ['META-INF', 'content.xml', 'meta.xml', 'mimetype',
                      'settings.xml', 'styles.xml', 'META-INF/manifest.xml']

    # Column width:
    # co1 2.000cm
    # co2 3.000cm
    # co3 4.000cm
    # co4 8.000cm

    _CONTENT_XML_HEADER = '''<?xml version="1.0" encoding="UTF-8"?>

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

    _CONTENT_XML_FOOTER = '''   </table:table>
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
    <meta:creation-date>${Datetime}Z</meta:creation-date>
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
   <style:text-properties style:font-name="Arial" fo:language="$Language" fo:country="$Country" style:font-name-asian="Arial Unicode MS" style:language-asian="zh" style:country-asian="CN" style:font-name-complex="Tahoma" style:language-complex="hi" style:country-complex="IN"/>
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

    def _convert_from_yw(self, text, quick=False):
        """Return text, converted from yw7 markup to target format.
        
        Positional arguments:
            text -- string to convert.
        
        Optional arguments:
            quick -- bool: if True, apply a conversion mode for one-liners without formatting.
        
        Overrides the superclass method.
        """
        ODS_REPLACEMENTS = [
            ('&', '&amp;'),  # must be first!
            ('"', '&quot;'),
            ("'", '&apos;'),
            ('>', '&gt;'),
            ('<', '&lt;'),
            ('\n', '</text:p>\n<text:p>'),
        ]
        try:
            text = text.rstrip()
            for yw, od in ODS_REPLACEMENTS:
                text = text.replace(yw, od)
        except AttributeError:
            text = ''
        return text


class OdsSceneList(OdsFile):
    """ODS scene list representation."""

    DESCRIPTION = 'Scene list'
    SUFFIX = '_scenelist'

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

    _fileHeader = f'''{OdsFile._CONTENT_XML_HEADER}{DESCRIPTION}" table:style-name="ta1" table:print="false">
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

    _sceneTemplate = '''   <table:table-row table:style-name="ro2">
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

    _fileFooter = OdsFile._CONTENT_XML_FOOTER 

    def _get_sceneMapping(self, scId, sceneNumber, wordsTotal, lettersTotal):
        """Return a mapping dictionary for a scene section.
        
        Positional arguments:
            scId -- str: scene ID.
            sceneNumber -- int: scene number to be displayed.
            wordsTotal -- int: accumulated wordcount.
            lettersTotal -- int: accumulated lettercount.
        
        Scene rating "1" is not applicable.
        Extends the superclass template method.
        """
        sceneMapping = super()._get_sceneMapping(scId, sceneNumber, wordsTotal, lettersTotal)
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
    _NOT_APPLICABLE = 'N/A'
    # Scene field column header for fields not being assigned to a storyline

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

    _fileHeader = f'''{OdsFile._CONTENT_XML_HEADER}{DESCRIPTION}" table:style-name="ta1" table:print="false">
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
      <text:p>Scene title</text:p>
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

    _notesChapterTemplate = '''   <table:table-row table:style-name="ro2">
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
    _sceneTemplate = '''   <table:table-row table:style-name="ro2">
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

    _fileFooter = OdsFile._CONTENT_XML_FOOTER 

    def _get_fileHeaderMapping(self):
        """Return a mapping dictionary for the project section.
        
        Special treatment of scene ratings as storylines.
        Overrides the superclass template method.
        """
        projectTemplateMapping = super()._get_fileHeaderMapping()
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

    def _get_sceneMapping(self, scId, sceneNumber, wordsTotal, lettersTotal):
        """Return a mapping dictionary for a scene section.
        
        Positional arguments:
            scId -- str: scene ID.
            sceneNumber -- int: scene number to be displayed.
            wordsTotal -- int: accumulated wordcount.
            lettersTotal -- int: accumulated lettercount.
        
        Special treatment of scene ratings as storylines.
        Scene rating "1" is not applicable.
        Extends the superclass template method.
        """
        sceneMapping = super()._get_sceneMapping(scId, sceneNumber, wordsTotal, lettersTotal)

        # Suppress display if the field doesn't represent a storyline,
        # or if the field's value equals 1

        if self.scenes[scId].field1 == '1' or not self.arc1:
            sceneMapping['Field1'] = '"string">\n'
        else:
            sceneMapping['Field1'] = f'"float" office:value="{sceneMapping["Field1"]}">\n      <text:p>{sceneMapping["Field1"]}</text:p>'
        if self.scenes[scId].field2 == '1' or not self.arc2:
            sceneMapping['Field2'] = '"string">\n'
        else:
            sceneMapping['Field2'] = f'"float" office:value="{sceneMapping["Field2"]}">\n      <text:p>{sceneMapping["Field2"]}</text:p>'
        if self.scenes[scId].field3 == '1' or not self.arc3:
            sceneMapping['Field3'] = '"string">\n'
        else:
            sceneMapping['Field3'] = f'"float" office:value="{sceneMapping["Field3"]}">\n      <text:p>{sceneMapping["Field3"]}</text:p>'
        if self.scenes[scId].field4 == '1' or not self.arc4:
            sceneMapping['Field4'] = '"string">\n'
        else:
            sceneMapping['Field4'] = f'"float" office:value="{sceneMapping["Field4"]}">\n      <text:p>{sceneMapping["Field4"]}</text:p>'
        return sceneMapping


class OdsCharList(OdsFile):
    """ODS character list representation."""

    DESCRIPTION = 'Character list'
    SUFFIX = '_charlist'
    
    _fileHeader = f'''{OdsFile._CONTENT_XML_HEADER}{DESCRIPTION}" table:style-name="ta1" table:print="false">
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
    _characterTemplate = '''   <table:table-row table:style-name="ro2">
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

    _fileFooter = OdsFile._CONTENT_XML_FOOTER


class OdsLocList(OdsFile):
    """ODS location list representation."""
    DESCRIPTION = 'Location list'
    SUFFIX = '_loclist'

    _fileHeader = f'''{OdsFile._CONTENT_XML_HEADER}{DESCRIPTION}" table:style-name="ta1" table:print="false">
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

    _locationTemplate = '''   <table:table-row table:style-name="ro2">
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
    _fileFooter = OdsFile._CONTENT_XML_FOOTER 


class OdsItemList(OdsFile):
    """ODS item list representation."""

    DESCRIPTION = 'Item list'
    SUFFIX = '_itemlist'

    _fileHeader = f'''{OdsFile._CONTENT_XML_HEADER}{DESCRIPTION}" table:style-name="ta1" table:print="false">
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

    _itemTemplate = '''   <table:table-row table:style-name="ro2">
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

    _fileFooter = OdsFile._CONTENT_XML_FOOTER 


class OdtCharacters(OdtFile):
    """ODT character descriptions file representation.

    Export a character sheet with invisibly tagged descriptions.
    """
    DESCRIPTION = 'Character descriptions'
    SUFFIX = '_characters'

    _fileHeader = f'''{OdtFile._CONTENT_XML_HEADER}<text:p text:style-name="Title">$Title</text:p>
<text:p text:style-name="Subtitle">$AuthorName</text:p>
'''

    _characterTemplate = '''<text:h text:style-name="Heading_20_2" text:outline-level="2">$Title$FullName$AKA</text:h>
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
<text:h text:style-name="Heading_20_3" text:outline-level="3">Notes</text:h>
<text:section text:style-name="Sect1" text:name="CrID_notes:$ID">
<text:p text:style-name="Text_20_body">$Notes</text:p>
</text:section>
</text:section>
'''

    _fileFooter = OdtFile._CONTENT_XML_FOOTER

    def _get_characterMapping(self, crId):
        """Return a mapping dictionary for a character section.
        
        Positional arguments:
            crId -- str: character ID.
        
        Special formatting of alternate and full name. 
        Extends the superclass method.
        """
        characterMapping = OdtFile._get_characterMapping(self, crId)
        if self.characters[crId].aka:
            characterMapping['AKA'] = f' ("{self.characters[crId].aka}")'
        if self.characters[crId].fullName:
            characterMapping['FullName'] = f'/{self.characters[crId].fullName}'
        return characterMapping


class OdtItems(OdtFile):
    """ODT item descriptions file representation.

    Export a item sheet with invisibly tagged descriptions.
    """
    DESCRIPTION = 'Item descriptions'
    SUFFIX = '_items'

    _fileHeader = f'''{OdtFile._CONTENT_XML_HEADER}<text:p text:style-name="Title">$Title</text:p>
<text:p text:style-name="Subtitle">$AuthorName</text:p>
'''

    _itemTemplate = '''<text:h text:style-name="Heading_20_2" text:outline-level="2">$Title$AKA</text:h>
<text:section text:style-name="Sect1" text:name="ItID:$ID">
<text:p text:style-name="Text_20_body">$Desc</text:p>
</text:section>
'''

    _fileFooter = OdtFile._CONTENT_XML_FOOTER

    def _get_itemMapping(self, itId):
        """Return a mapping dictionary for an item section.
        
        Positional arguments:
            itId -- str: item ID.
        
        Special formatting of alternate name. 
        Extends the superclass method.
        """
        itemMapping = super()._get_itemMapping(itId)
        if self.items[itId].aka:
            itemMapping['AKA'] = f' ("{self.items[itId].aka}")'
        return itemMapping


class OdtLocations(OdtFile):
    """ODT location descriptions file representation.

    Export a location sheet with invisibly tagged descriptions.
    """
    DESCRIPTION = 'Location descriptions'
    SUFFIX = '_locations'

    _fileHeader = f'''{OdtFile._CONTENT_XML_HEADER}<text:p text:style-name="Title">$Title</text:p>
<text:p text:style-name="Subtitle">$AuthorName</text:p>
'''

    _locationTemplate = '''<text:h text:style-name="Heading_20_2" text:outline-level="2">$Title$AKA</text:h>
<text:section text:style-name="Sect1" text:name="LcID:$ID">
<text:p text:style-name="Text_20_body">$Desc</text:p>
</text:section>
'''

    _fileFooter = OdtFile._CONTENT_XML_FOOTER

    def _get_locationMapping(self, lcId):
        """Return a mapping dictionary for a location section.
        
        Positional arguments:
            lcId -- str: location ID.
        
        Special formatting of alternate name. 
        Extends the superclass method.
        """
        locationMapping = super()._get_locationMapping(lcId)
        if self.locations[lcId].aka:
            locationMapping['AKA'] = f' ("{self.locations[lcId].aka}")'
        return locationMapping
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
    mb = toolkit.createMessageBox(parent, type_msg, buttons, title, str(message))
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
    if filepicker.execute():
        return filepicker.getFiles()[0]


import sys
import webbrowser


class Ui:
    """Base class for UI facades, implementing a 'silent mode'.
    
    Public methods:
        ask_yes_no(text) -- return True or False.
        set_info_what(message) -- show what the converter is going to do.
        set_info_how(message) -- show how the converter is doing.
        start() -- launch the GUI, if any.
        
    Public instance variables:
        infoWhatText -- buffer for general messages.
        infoHowText -- buffer for error/success messages.
    """

    def __init__(self, title):
        """Initialize text buffers for messaging.
        
        Positional arguments:
            title -- application title.
        """
        self.infoWhatText = ''
        self.infoHowText = ''

    def ask_yes_no(self, text):
        """Return True or False.
        
        Positional arguments:
            text -- question to be asked. 
            
        This is a stub used for "silent mode".
        The application may use a subclass for confirmation requests.    
        """
        return True

    def set_info_what(self, message):
        """Show what the converter is going to do.
        
        Positional arguments:
            message -- message to be buffered. 
        """
        self.infoWhatText = message

    def set_info_how(self, message):
        """Show how the converter is doing.
        
        Positional arguments:
            message -- message to be buffered.
            
        Print the message to stderr, replacing the error marker, if any.
        """
        if message.startswith(ERROR):
            message = f'FAIL: {message.split(ERROR, maxsplit=1)[1].strip()}'
            sys.stderr.write(message)
        self.infoHowText = message

    def start(self):
        """Launch the GUI, if any.
        
        To be overridden by subclasses requiring
        special action to launch the user interaction.
        """


class YwCnv:
    """Base class for Novel file conversion.

    Public methods:
        convert(sourceFile, targetFile) -- Convert sourceFile into targetFile.
    """

    def convert(self, source, target):
        """Convert source into target and return a message.

        Positional arguments:
            source, target -- Novel subclass instances.

        Operation:
        1. Make the source object read the source file.
        2. Make the target object merge the source object's instance variables.
        3. Make the target object write the target file.
        Return a message beginning with the ERROR constant in case of error.

        Error handling:
        - Check if source and target are correctly initialized.
        - Ask for permission to overwrite target.
        - Pass the error messages of the called methods of source and target.
        - The success message comes from target.write(), if called.       
        """
        if source.filePath is None:
            return f'{ERROR}Source "{os.path.normpath(source.filePath)}" is not of the supported type.'

        if not os.path.isfile(source.filePath):
            return f'{ERROR}"{os.path.normpath(source.filePath)}" not found.'

        if target.filePath is None:
            return f'{ERROR}Target "{os.path.normpath(target.filePath)}" is not of the supported type.'

        if os.path.isfile(target.filePath) and not self._confirm_overwrite(target.filePath):
            return f'{ERROR}Action canceled by user.'

        message = source.read()
        if message.startswith(ERROR):
            return message

        message = target.merge(source)
        if message.startswith(ERROR):
            return message

        return target.write()

    def _confirm_overwrite(self, fileName):
        """Return boolean permission to overwrite the target file.
        
        Positional argument:
            fileName -- path to the target file.
        
        This is a stub to be overridden by subclass methods.
        """
        return True


class YwCnvUi(YwCnv):
    """Base class for Novel file conversion with user interface.

    Public methods:
        export_from_yw(sourceFile, targetFile) -- Convert from yWriter project to other file format.
        create_yw(sourceFile, targetFile) -- Create target from source.
        import_to_yw(sourceFile, targetFile) -- Convert from any file format to yWriter project.

    Instance variables:
        ui -- Ui (can be overridden e.g. by subclasses).
        newFile -- str: path to the target file in case of success.   
    """

    def __init__(self):
        """Define instance variables."""
        self.ui = Ui('')
        # Per default, 'silent mode' is active.
        self.newFile = None
        # Also indicates successful conversion.

    def export_from_yw(self, source, target):
        """Convert from yWriter project to other file format.

        Positional arguments:
            source -- YwFile subclass instance.
            target -- Any Novel subclass instance.

        Operation:
        1. Send specific information about the conversion to the UI.
        2. Convert source into target.
        3. Pass the message to the UI.
        4. Save the new file pathname.

        Error handling:
        - If the conversion fails, newFile is set to None.
        """
        self.ui.set_info_what(
            f'Input: {source.DESCRIPTION} "{os.path.normpath(source.filePath)}"\nOutput: {target.DESCRIPTION} "{os.path.normpath(target.filePath)}"')
        message = self.convert(source, target)
        self.ui.set_info_how(message)
        if message.startswith(ERROR):
            self.newFile = None
        else:
            self.newFile = target.filePath

    def create_yw7(self, source, target):
        """Create target from source.

        Positional arguments:
            source -- Any Novel subclass instance.
            target -- YwFile subclass instance.

        Operation:
        1. Send specific information about the conversion to the UI.
        2. Convert source into target.
        3. Pass the message to the UI.
        4. Save the new file pathname.

        Error handling:
        - Tf target already exists as a file, the conversion is cancelled,
          an error message is sent to the UI.
        - If the conversion fails, newFile is set to None.
        """
        self.ui.set_info_what(
            f'Create a yWriter project file from {source.DESCRIPTION}\nNew project: "{os.path.normpath(target.filePath)}"')
        if os.path.isfile(target.filePath):
            self.ui.set_info_how(f'{ERROR}"{os.path.normpath(target.filePath)}" already exists.')
        else:
            message = self.convert(source, target)
            self.ui.set_info_how(message)
            if message.startswith(ERROR):
                self.newFile = None
            else:
                self.newFile = target.filePath

    def import_to_yw(self, source, target):
        """Convert from any file format to yWriter project.

        Positional arguments:
            source -- Any Novel subclass instance.
            target -- YwFile subclass instance.

        Operation:
        1. Send specific information about the conversion to the UI.
        2. Convert source into target.
        3. Pass the message to the UI.
        4. Delete the temporay file, if exists.
        5. Save the new file pathname.

        Error handling:
        - If the conversion fails, newFile is set to None.
        """
        self.ui.set_info_what(
            f'Input: {source.DESCRIPTION} "{os.path.normpath(source.filePath)}"\nOutput: {target.DESCRIPTION} "{os.path.normpath(target.filePath)}"')
        message = self.convert(source, target)
        self.ui.set_info_how(message)
        self._delete_tempfile(source.filePath)
        if message.startswith(ERROR):
            self.newFile = None
        else:
            self.newFile = target.filePath

    def _confirm_overwrite(self, filePath):
        """Return boolean permission to overwrite the target file.
        
        Positional arguments:
            fileName -- path to the target file.
        
        Overrides the superclass method.
        """
        return self.ui.ask_yes_no(f'Overwrite existing file "{os.path.normpath(filePath)}"?')

    def _delete_tempfile(self, filePath):
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

    def _open_newFile(self):
        """Open the converted file for editing and exit the converter script."""
        webbrowser.open(self.newFile)
        sys.exit(0)


class FileFactory:
    """Base class for conversion object factory classes.
    """

    def __init__(self, fileClasses=[]):
        """Write the parameter to a "private" instance variable.

        Optional arguments:
            _fileClasses -- list of classes from which an instance can be returned.
        """
        self._fileClasses = fileClasses


class ExportSourceFactory(FileFactory):
    """A factory class that instantiates a yWriter object to read.

    Public methods:
        make_file_objects(self, sourcePath, **kwargs) -- return conversion objects.
    """

    def make_file_objects(self, sourcePath, **kwargs):
        """Instantiate a source object for conversion from a yWriter project.

        Positional arguments:
            sourcePath -- str: path to the source file to convert.

        Return a tuple with three elements:
        - A message beginning with the ERROR constant in case of error
        - sourceFile: a YwFile subclass instance, or None in case of error
        - targetFile: None
        """
        __, fileExtension = os.path.splitext(sourcePath)
        for fileClass in self._fileClasses:
            if fileClass.EXTENSION == fileExtension:
                sourceFile = fileClass(sourcePath, **kwargs)
                return 'Source object created.', sourceFile, None
            
        return f'{ERROR}File type of "{os.path.normpath(sourcePath)}" not supported.', None, None


class ExportTargetFactory(FileFactory):
    """A factory class that instantiates a document object to write.

    Public methods:
        make_file_objects(self, sourcePath, **kwargs) -- return conversion objects.
    """

    def make_file_objects(self, sourcePath, **kwargs):
        """Instantiate a target object for conversion from a yWriter project.

        Positional arguments:
            sourcePath -- str: path to the source file to convert.

        Optional arguments:
            suffix -- str: an indicator for the target file type.

        Required keyword arguments: 
            suffix -- str: target file name suffix.

        Return a tuple with three elements:
        - A message beginning with the ERROR constant in case of error
        - sourceFile: None
        - targetFile: a FileExport subclass instance, or None in case of error 
        """
        fileName, __ = os.path.splitext(sourcePath)
        suffix = kwargs['suffix']
        for fileClass in self._fileClasses:
            if fileClass.SUFFIX == suffix:
                if suffix is None:
                    suffix = ''
                targetFile = fileClass(f'{fileName}{suffix}{fileClass.EXTENSION}', **kwargs)
                return 'Target object created.', None, targetFile
        
        return f'{ERROR}File type of "{os.path.normpath(sourcePath)}" not supported.', None, None


class ImportSourceFactory(FileFactory):
    """A factory class that instantiates a documente object to read.

    Public methods:
        make_file_objects(self, sourcePath, **kwargs) -- return conversion objects.
    """

    def make_file_objects(self, sourcePath, **kwargs):
        """Instantiate a source object for conversion to a yWriter project.       

        Positional arguments:
            sourcePath -- str: path to the source file to convert.

        Return a tuple with three elements:
        - A message beginning with the ERROR constant in case of error
        - sourceFile: a Novel subclass instance, or None in case of error
        - targetFile: None
        """
        for fileClass in self._fileClasses:
            if fileClass.SUFFIX is not None:
                if sourcePath.endswith(f'{fileClass.SUFFIX }{fileClass.EXTENSION}'):
                    sourceFile = fileClass(sourcePath, **kwargs)
                    return 'Source object created.', sourceFile, None
        return f'{ERROR}This document is not meant to be written back.', None, None


class ImportTargetFactory(FileFactory):
    """A factory class that instantiates a yWriter object to write.

    Public methods:
        make_file_objects(self, sourcePath, **kwargs) -- return conversion objects.
    """

    def make_file_objects(self, sourcePath, **kwargs):
        """Instantiate a target object for conversion to a yWriter project.

        Positional arguments:
            sourcePath -- str: path to the source file to convert.

        Optional arguments:
            suffix -- str: an indicator for the source file type.

        Required keyword arguments: 
            suffix -- str: target file name suffix.

        Return a tuple with three elements:
        - A message beginning with the ERROR constant in case of error
        - sourceFile: None
        - targetFile: a YwFile subclass instance, or None in case of error
        """
        fileName, __ = os.path.splitext(sourcePath)
        sourceSuffix = kwargs['suffix']
        if sourceSuffix:
            ywPathBasis = fileName.split(sourceSuffix)[0]
        else:
            ywPathBasis = fileName

        # Look for an existing yWriter project to rewrite.
        for fileClass in self._fileClasses:
            if os.path.isfile(f'{ywPathBasis}{fileClass.EXTENSION}'):
                targetFile = fileClass(f'{ywPathBasis}{fileClass.EXTENSION}', **kwargs)
                return 'Target object created.', None, targetFile
            
        return f'{ERROR}No yWriter project to write.', None, None


class YwCnvFf(YwCnvUi):
    """Class for Novel file conversion using factory methods 
    to create target and source classes.

    Public methods:
        run(sourcePath, **kwargs) -- Create source and target objects and run conversion.

    Class constants:
        EXPORT_SOURCE_CLASSES -- List of YwFile subclasses from which can be exported.
        EXPORT_TARGET_CLASSES -- List of FileExport subclasses to which export is possible.
        IMPORT_SOURCE_CLASSES -- List of Novel subclasses from which can be imported.
        IMPORT_TARGET_CLASSES -- List of YwFile subclasses to which import is possible.

    All lists are empty and meant to be overridden by subclasses.

    Instance variables:
        exportSourceFactory -- ExportSourceFactory.
        exportTargetFactory -- ExportTargetFactory.
        importSourceFactory -- ImportSourceFactory.
        importTargetFactory -- ImportTargetFactory.
        newProjectFactory -- FileFactory (a stub to be overridden by subclasses).
    """
    EXPORT_SOURCE_CLASSES = []
    EXPORT_TARGET_CLASSES = []
    IMPORT_SOURCE_CLASSES = []
    IMPORT_TARGET_CLASSES = []

    def __init__(self):
        """Create strategy class instances.
        
        Extends the superclass constructor.
        """
        super().__init__()
        self.exportSourceFactory = ExportSourceFactory(self.EXPORT_SOURCE_CLASSES)
        self.exportTargetFactory = ExportTargetFactory(self.EXPORT_TARGET_CLASSES)
        self.importSourceFactory = ImportSourceFactory(self.IMPORT_SOURCE_CLASSES)
        self.importTargetFactory = ImportTargetFactory(self.IMPORT_TARGET_CLASSES)
        self.newProjectFactory = FileFactory()

    def run(self, sourcePath, **kwargs):
        """Create source and target objects and run conversion.

        Positional arguments: 
            sourcePath -- str: the source file path.
        
        Required keyword arguments: 
            suffix -- str: target file name suffix.

        This is a template method that calls superclass methods as primitive operations by case.
        """
        self.newFile = None
        if not os.path.isfile(sourcePath):
            self.ui.set_info_how(f'{ERROR}File "{os.path.normpath(sourcePath)}" not found.')
            return
        
        message, source, __ = self.exportSourceFactory.make_file_objects(sourcePath, **kwargs)
        if message.startswith(ERROR):
            # The source file is not a yWriter project.
            message, source, __ = self.importSourceFactory.make_file_objects(sourcePath, **kwargs)
            if message.startswith(ERROR):
                # A new yWriter project might be required.
                message, source, target = self.newProjectFactory.make_file_objects(sourcePath, **kwargs)
                if message.startswith(ERROR):
                    self.ui.set_info_how(message)
                else:
                    self.create_yw7(source, target)
            else:
                # Try to update an existing yWriter project.
                kwargs['suffix'] = source.SUFFIX
                message, __, target = self.importTargetFactory.make_file_objects(sourcePath, **kwargs)
                if message.startswith(ERROR):
                    self.ui.set_info_how(message)
                else:
                    self.import_to_yw(source, target)
        else:
            # The source file is a yWriter project.
            message, __, target = self.exportTargetFactory.make_file_objects(sourcePath, **kwargs)
            if message.startswith(ERROR):
                self.ui.set_info_how(message)
            else:
                self.export_from_yw(source, target)
from html import unescape
import xml.etree.ElementTree as ET


class Chapter:
    """yWriter chapter representation.
    
    Public instance variables:
        title -- str: chapter title (may be the heading).
        desc -- str: chapter description in a single string.
        chLevel -- int: chapter level (part/chapter).
        oldType -- int: chapter type (Chapter/Other).
        chType -- int: chapter type yWriter 7.0.7.2+ (Normal/Notes/Todo).
        isUnused -- bool: True, if the chapter is marked "Unused".
        suppressChapterTitle -- bool: uppress chapter title when exporting.
        isTrash -- bool: True, if the chapter is the project's trash bin.
        suppressChapterBreak -- bool: Suppress chapter break when exporting.
        srtScenes -- list of str: the chapter's sorted scene IDs.        
    """

    def __init__(self):
        """Initialize instance variables."""
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
        # Applies to projects created by a yWriter version prior to 7.0.7.2.

        self.chType = None
        # int
        # xml: <ChapterType>
        # 0 = Normal
        # 1 = Notes
        # 2 = Todo
        # Applies to projects created by yWriter version 7.0.7.2+.

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



class Splitter:
    """Helper class for scene and chapter splitting.
    
    When importing scenes to yWriter, they may contain manuallyinserted scene and chapter dividers.
    The Splitter class updates a Novel instance by splitting such scenes and creating new chapters and scenes. 
    
    Public methods:
        split_scenes(novel) -- Split scenes by inserted chapter and scene dividers.
        
    Public class constants:
        PART_SEPARATOR -- marker indicating the beginning of a new part, splitting a scene.
        CHAPTER_SEPARATOR -- marker indicating the beginning of a new chapter, splitting a scene.
    """
    PART_SEPARATOR = '# '
    CHAPTER_SEPARATOR = '## '
    _SCENE_SEPARATOR = '* * *'
    _CLIP_TITLE = 20
    # Maximum length of newly generated scene titles.

    def split_scenes(self, novel):
        """Split scenes by inserted chapter and scene dividers.
        
        Update a Novel instance by generating new chapters and scenes 
        if there are dividers within the scene content.
        
        Positional argument: 
            novel -- Novel instance to update.
        """

        def create_chapter(chapterId, title, desc, level):
            """Create a new chapter and add it to the novel.
            
            Positional arguments:
                chapterId -- str: ID of the chapter to create.
                title -- str: title of the chapter to create.
                desc -- str: description of the chapter to create.
                level -- int: chapter level (part/chapter).           
            """
            newChapter = Chapter()
            newChapter.title = title
            newChapter.desc = desc
            newChapter.chLevel = level
            newChapter.chType = 0
            novel.chapters[chapterId] = newChapter

        def create_scene(sceneId, parent, splitCount):
            """Create a new scene and add it to the novel.
            
            Positional arguments:
                sceneId -- str: ID of the scene to create.
                parent -- Scene instance: parent scene.
                splitCount -- int: number of parent's splittings.
            """
            WARNING = ' (!) '

            # Mark metadata of split scenes.
            newScene = Scene()
            if parent.title:
                if len(parent.title) > self._CLIP_TITLE:
                    title = f'{parent.title[:self._CLIP_TITLE]}...'
                else:
                    title = parent.title
                newScene.title = f'{title} Split: {splitCount}'
            else:
                newScene.title = f'New scene Split: {splitCount}'
            if parent.desc and not parent.desc.startswith(WARNING):
                parent.desc = f'{WARNING}{parent.desc}'
            if parent.goal and not parent.goal.startswith(WARNING):
                parent.goal = f'{WARNING}{parent.goal}'
            if parent.conflict and not parent.conflict.startswith(WARNING):
                parent.conflict = f'{WARNING}{parent.conflict}'
            if parent.outcome and not parent.outcome.startswith(WARNING):
                parent.outcome = f'{WARNING}{parent.outcome}'

            # Reset the parent's status to Draft, if not Outline.
            if parent.status > 2:
                parent.status = 2
            newScene.status = parent.status
            newScene.isNotesScene = parent.isNotesScene
            newScene.isUnused = parent.isUnused
            newScene.isTodoScene = parent.isTodoScene
            newScene.date = parent.date
            newScene.time = parent.time
            newScene.day = parent.day
            newScene.hour = parent.hour
            newScene.minute = parent.minute
            newScene.lastsDays = parent.lastsDays
            newScene.lastsHours = parent.lastsHours
            newScene.lastsMinutes = parent.lastsMinutes
            novel.scenes[sceneId] = newScene

        # Get the maximum chapter ID and scene ID.
        chIdMax = 0
        scIdMax = 0
        for chId in novel.srtChapters:
            if int(chId) > chIdMax:
                chIdMax = int(chId)
        for scId in novel.scenes:
            if int(scId) > scIdMax:
                scIdMax = int(scId)
                
        #Process chapters and scenes.
        srtChapters = []
        for chId in novel.srtChapters:
            srtChapters.append(chId)
            chapterId = chId
            srtScenes = []
            for scId in novel.chapters[chId].srtScenes:
                srtScenes.append(scId)
                if not novel.scenes[scId].sceneContent:
                    continue

                sceneId = scId
                lines = novel.scenes[scId].sceneContent.split('\n')
                newLines = []
                inScene = True
                sceneSplitCount = 0

                #Search scene content for dividers.
                for line in lines:
                    if line.startswith(self.PART_SEPARATOR):
                        if inScene:
                            novel.scenes[sceneId].sceneContent = '\n'.join(newLines)
                            newLines = []
                            sceneSplitCount = 0
                            inScene = False
                        novel.chapters[chapterId].srtScenes = srtScenes
                        srtScenes = []
                        chIdMax += 1
                        chapterId = str(chIdMax)
                        create_chapter(chapterId, 'New part', line.replace(self.PART_SEPARATOR, ''), 1)
                        srtChapters.append(chapterId)
                    elif line.startswith(self.CHAPTER_SEPARATOR):
                        if inScene:
                            novel.scenes[sceneId].sceneContent = '\n'.join(newLines)
                            newLines = []
                            sceneSplitCount = 0
                            inScene = False
                        novel.chapters[chapterId].srtScenes = srtScenes
                        srtScenes = []
                        chIdMax += 1
                        chapterId = str(chIdMax)
                        create_chapter(chapterId, 'New chapter', line.replace(self.CHAPTER_SEPARATOR, ''), 0)
                        srtChapters.append(chapterId)
                    elif line.startswith(self._SCENE_SEPARATOR):
                        novel.scenes[sceneId].sceneContent = '\n'.join(newLines)
                        newLines = []
                        sceneSplitCount += 1
                        scIdMax += 1
                        sceneId = str(scIdMax)
                        create_scene(sceneId, novel.scenes[scId], sceneSplitCount)
                        srtScenes.append(sceneId)
                        inScene = True
                    elif not inScene:
                        newLines.append(line)
                        sceneSplitCount += 1
                        scIdMax += 1
                        sceneId = str(scIdMax)
                        create_scene(sceneId, novel.scenes[scId], sceneSplitCount)
                        srtScenes.append(sceneId)
                        inScene = True
                    else:
                        newLines.append(line)
                novel.scenes[sceneId].sceneContent = '\n'.join(newLines)
            novel.chapters[chapterId].srtScenes = srtScenes
        novel.srtChapters = srtChapters


def indent(elem, level=0):
    """xml pretty printer

    Kudos to to Fredrik Lundh. 
    Source: http://effbot.org/zone/element-lib.htm#prettyprint
    """
    i = f'\n{level * "  "}'
    if elem:
        if not elem.text or not elem.text.strip():
            elem.text = f'{i}  '
        if not elem.tail or not elem.tail.strip():
            elem.tail = i
        for elem in elem:
            indent(elem, level + 1)
        if not elem.tail or not elem.tail.strip():
            elem.tail = i
    else:
        if level and (not elem.tail or not elem.tail.strip()):
            elem.tail = i


class Yw7File(Novel):
    """yWriter 7 project file representation.

    Public methods: 
        read() -- parse the yWriter xml file and get the instance variables.
        merge(source) -- update instance variables from a source instance.
        write() -- write instance variables to the yWriter xml file.
        is_locked() -- check whether the yw7 file is locked by yWriter.

    Public instance variables:
        tree -- xml element tree of the yWriter project
    """
    DESCRIPTION = 'yWriter 7 project'
    EXTENSION = '.yw7'
    _CDATA_TAGS = ['Title', 'AuthorName', 'Bio', 'Desc',
                   'FieldTitle1', 'FieldTitle2', 'FieldTitle3',
                   'FieldTitle4', 'LaTeXHeaderFile', 'Tags',
                   'AKA', 'ImageFile', 'FullName', 'Goals',
                   'Notes', 'RTFFile', 'SceneContent',
                   'Outcome', 'Goal', 'Conflict']
    # Names of xml elements containing CDATA.
    # ElementTree.write omits CDATA tags, so they have to be inserted afterwards.

    def __init__(self, filePath, **kwargs):
        """Initialize instance variables.
        
        Positional arguments:
            filePath -- str: path to the yw7 file.
            
        Optional arguments:
            kwargs -- keyword arguments (not used here).            
        
        Extends the superclass constructor.
        """
        super().__init__(filePath)
        self.tree = None

    def read(self):
        """Parse the yWriter xml file and get the instance variables.
        
        Return a message beginning with the ERROR constant in case of error.
        Overrides the superclass method.
        """
        if self.is_locked():
            return f'{ERROR}yWriter seems to be open. Please close first.'
        try:
            self.tree = ET.parse(self.filePath)
        except:
            return f'{ERROR}Can not process "{os.path.normpath(self.filePath)}".'

        root = self.tree.getroot()

        #--- Read locations from the xml element tree.
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
                    self.locations[lcId].tags = self._strip_spaces(tags)

        #--- Read items from the xml element tree.
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
                    self.items[itId].tags = self._strip_spaces(tags)

        #--- Read characters from the xml element tree.
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
                    self.characters[crId].tags = self._strip_spaces(tags)
            
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

        #--- Read attributes at novel level from the xml element tree.
        prj = root.find('PROJECT')
        
        if prj.find('Title') is not None:
            self.title = prj.find('Title').text
        
        if prj.find('AuthorName') is not None:
            self.authorName = prj.find('AuthorName').text
        
        if prj.find('Bio') is not None:
            self.authorBio = prj.find('Bio').text
        
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

        #--- Read attributes at chapter level from the xml element tree.
        self.srtChapters = []
        # This is necessary for re-reading.
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

        #--- Read attributes at scene level from the xml element tree.
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
            self.scenes[scId].isNotesScene = False
            self.scenes[scId].isTodoScene = False
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
                    self.scenes[scId].tags = self._strip_spaces(tags)
            
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
            
            if scn.find('ImageFile') is not None:
                self.scenes[scId].image = scn.find('ImageFile').text
            
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

        # Make sure that ToDo, Notes, and Unused type is inherited from the chapter.
        for chId in self.chapters:
            if self.chapters[chId].chType == 2:
                # Chapter is "ToDo" type.
                for scId in self.chapters[chId].srtScenes:
                    self.scenes[scId].isTodoScene = True
                    self.scenes[scId].isUnused = True
            elif self.chapters[chId].chType == 1:
                # Chapter is "Notes" type.
                for scId in self.chapters[chId].srtScenes:
                    self.scenes[scId].isNotesScene = True
                    self.scenes[scId].isUnused = True
            elif self.chapters[chId].isUnused:
                for scId in self.chapters[chId].srtScenes:
                    self.scenes[scId].isUnused = True
        return 'yWriter project data read in.'

    def merge(self, source):
        """Update instance variables from a source instance.
        
        Positional arguments:
            source -- Novel subclass instance to merge.
        
        Return a message beginning with the ERROR constant in case of error.
        Overrides the superclass method.
        """

        def merge_lists(srcLst, tgtLst):
            """Insert srcLst items to tgtLst, if missing.
            """
            j = 0
            for i in range(len(srcLst)):
                if not srcLst[i] in tgtLst:
                    tgtLst.insert(j, srcLst[i])
                    j += 1
                else:
                    j = tgtLst.index(srcLst[i]) + 1
        if os.path.isfile(self.filePath):
            message = self.read()
            # initialize data
            if message.startswith(ERROR):
                return message

        #--- Merge and re-order locations.
        if source.srtLocations:
            self.srtLocations = source.srtLocations
            temploc = self.locations
            self.locations = {}
            for lcId in source.srtLocations:

                # Build a new self.locations dictionary sorted like the source.
                self.locations[lcId] = WorldElement()
                if not lcId in temploc:
                    # A new location has been added
                    temploc[lcId] = WorldElement()
                if source.locations[lcId].title:
                    # avoids deleting the title, if it is empty by accident
                    self.locations[lcId].title = source.locations[lcId].title
                else:
                    self.locations[lcId].title = temploc[lcId].title
                if source.locations[lcId].image is not None:
                    self.locations[lcId].image = source.locations[lcId].image
                else:
                    self.locations[lcId].desc = temploc[lcId].desc
                if source.locations[lcId].desc is not None:
                    self.locations[lcId].desc = source.locations[lcId].desc
                else:
                    self.locations[lcId].desc = temploc[lcId].desc
                if source.locations[lcId].aka is not None:
                    self.locations[lcId].aka = source.locations[lcId].aka
                else:
                    self.locations[lcId].aka = temploc[lcId].aka
                if source.locations[lcId].tags is not None:
                    self.locations[lcId].tags = source.locations[lcId].tags
                else:
                    self.locations[lcId].tags = temploc[lcId].tags

        #--- Merge and re-order items.
        if source.srtItems:
            self.srtItems = source.srtItems
            tempitm = self.items
            self.items = {}
            for itId in source.srtItems:

                # Build a new self.items dictionary sorted like the source.
                self.items[itId] = WorldElement()
                if not itId in tempitm:
                    # A new item has been added
                    tempitm[itId] = WorldElement()
                if source.items[itId].title:
                    # avoids deleting the title, if it is empty by accident
                    self.items[itId].title = source.items[itId].title
                else:
                    self.items[itId].title = tempitm[itId].title
                if source.items[itId].image is not None:
                    self.items[itId].image = source.items[itId].image
                else:
                    self.items[itId].image = tempitm[itId].image
                if source.items[itId].desc is not None:
                    self.items[itId].desc = source.items[itId].desc
                else:
                    self.items[itId].desc = tempitm[itId].desc
                if source.items[itId].aka is not None:
                    self.items[itId].aka = source.items[itId].aka
                else:
                    self.items[itId].aka = tempitm[itId].aka
                if source.items[itId].tags is not None:
                    self.items[itId].tags = source.items[itId].tags
                else:
                    self.items[itId].tags = tempitm[itId].tags

        #--- Merge and re-order characters.
        if source.srtCharacters:
            self.srtCharacters = source.srtCharacters
            tempchr = self.characters
            self.characters = {}
            for crId in source.srtCharacters:

                # Build a new self.characters dictionary sorted like the source.
                self.characters[crId] = Character()
                if not crId in tempchr:
                    # A new character has been added
                    tempchr[crId] = Character()
                if source.characters[crId].title:
                    # avoids deleting the title, if it is empty by accident
                    self.characters[crId].title = source.characters[crId].title
                else:
                    self.characters[crId].title = tempchr[crId].title
                if source.characters[crId].image is not None:
                    self.characters[crId].image = source.characters[crId].image
                else:
                    self.characters[crId].image = tempchr[crId].image
                if source.characters[crId].desc is not None:
                    self.characters[crId].desc = source.characters[crId].desc
                else:
                    self.characters[crId].desc = tempchr[crId].desc
                if source.characters[crId].aka is not None:
                    self.characters[crId].aka = source.characters[crId].aka
                else:
                    self.characters[crId].aka = tempchr[crId].aka
                if source.characters[crId].tags is not None:
                    self.characters[crId].tags = source.characters[crId].tags
                else:
                    self.characters[crId].tags = tempchr[crId].tags
                if source.characters[crId].notes is not None:
                    self.characters[crId].notes = source.characters[crId].notes
                else:
                    self.characters[crId].notes = tempchr[crId].notes
                if source.characters[crId].bio is not None:
                    self.characters[crId].bio = source.characters[crId].bio
                else:
                    self.characters[crId].bio = tempchr[crId].bio
                if source.characters[crId].goals is not None:
                    self.characters[crId].goals = source.characters[crId].goals
                else:
                    self.characters[crId].goals = tempchr[crId].goals
                if source.characters[crId].fullName is not None:
                    self.characters[crId].fullName = source.characters[crId].fullName
                else:
                    self.characters[crId].fullName = tempchr[crId].fullName
                if source.characters[crId].isMajor is not None:
                    self.characters[crId].isMajor = source.characters[crId].isMajor
                else:
                    self.characters[crId].isMajor = tempchr[crId].isMajor

        #--- Merge scenes.
        sourceHasSceneContent = False
        for scId in source.scenes:
            if not scId in self.scenes:
                self.scenes[scId] = Scene()
            if source.scenes[scId].title:
                # avoids deleting the title, if it is empty by accident
                self.scenes[scId].title = source.scenes[scId].title
            if source.scenes[scId].desc is not None:
                self.scenes[scId].desc = source.scenes[scId].desc
            if source.scenes[scId].sceneContent is not None:
                self.scenes[scId].sceneContent = source.scenes[scId].sceneContent
                sourceHasSceneContent = True
            if source.scenes[scId].isUnused is not None:
                self.scenes[scId].isUnused = source.scenes[scId].isUnused
            if source.scenes[scId].isNotesScene is not None:
                self.scenes[scId].isNotesScene = source.scenes[scId].isNotesScene
            if source.scenes[scId].isTodoScene is not None:
                self.scenes[scId].isTodoScene = source.scenes[scId].isTodoScene
            if source.scenes[scId].status is not None:
                self.scenes[scId].status = source.scenes[scId].status
            if source.scenes[scId].sceneNotes is not None:
                self.scenes[scId].sceneNotes = source.scenes[scId].sceneNotes
            if source.scenes[scId].tags is not None:
                self.scenes[scId].tags = source.scenes[scId].tags
            if source.scenes[scId].field1 is not None:
                self.scenes[scId].field1 = source.scenes[scId].field1
            if source.scenes[scId].field2 is not None:
                self.scenes[scId].field2 = source.scenes[scId].field2
            if source.scenes[scId].field3 is not None:
                self.scenes[scId].field3 = source.scenes[scId].field3
            if source.scenes[scId].field4 is not None:
                self.scenes[scId].field4 = source.scenes[scId].field4
            if source.scenes[scId].appendToPrev is not None:
                self.scenes[scId].appendToPrev = source.scenes[scId].appendToPrev
            if source.scenes[scId].date or source.scenes[scId].time:
                if source.scenes[scId].date is not None:
                    self.scenes[scId].date = source.scenes[scId].date
                if source.scenes[scId].time is not None:
                    self.scenes[scId].time = source.scenes[scId].time
            elif source.scenes[scId].minute or source.scenes[scId].hour or source.scenes[scId].day:
                self.scenes[scId].date = None
                self.scenes[scId].time = None
            if source.scenes[scId].minute is not None:
                self.scenes[scId].minute = source.scenes[scId].minute
            if source.scenes[scId].hour is not None:
                self.scenes[scId].hour = source.scenes[scId].hour
            if source.scenes[scId].day is not None:
                self.scenes[scId].day = source.scenes[scId].day
            if source.scenes[scId].lastsMinutes is not None:
                self.scenes[scId].lastsMinutes = source.scenes[scId].lastsMinutes
            if source.scenes[scId].lastsHours is not None:
                self.scenes[scId].lastsHours = source.scenes[scId].lastsHours
            if source.scenes[scId].lastsDays is not None:
                self.scenes[scId].lastsDays = source.scenes[scId].lastsDays
            if source.scenes[scId].isReactionScene is not None:
                self.scenes[scId].isReactionScene = source.scenes[scId].isReactionScene
            if source.scenes[scId].isSubPlot is not None:
                self.scenes[scId].isSubPlot = source.scenes[scId].isSubPlot
            if source.scenes[scId].goal is not None:
                self.scenes[scId].goal = source.scenes[scId].goal
            if source.scenes[scId].conflict is not None:
                self.scenes[scId].conflict = source.scenes[scId].conflict
            if source.scenes[scId].outcome is not None:
                self.scenes[scId].outcome = source.scenes[scId].outcome
            if source.scenes[scId].characters is not None:
                self.scenes[scId].characters = []
                for crId in source.scenes[scId].characters:
                    if crId in self.characters:
                        self.scenes[scId].characters.append(crId)
            if source.scenes[scId].locations is not None:
                self.scenes[scId].locations = []
                for lcId in source.scenes[scId].locations:
                    if lcId in self.locations:
                        self.scenes[scId].locations.append(lcId)
            if source.scenes[scId].items is not None:
                self.scenes[scId].items = []
                for itId in source.scenes[scId].items:
                    if itId in self.items:
                        self.scenes[scId].items.append(itId)

        #--- Merge chapters.
        for chId in source.chapters:
            if not chId in self.chapters:
                self.chapters[chId] = Chapter()
            if source.chapters[chId].title:
                # avoids deleting the title, if it is empty by accident
                self.chapters[chId].title = source.chapters[chId].title
            if source.chapters[chId].desc is not None:
                self.chapters[chId].desc = source.chapters[chId].desc
            if source.chapters[chId].chLevel is not None:
                self.chapters[chId].chLevel = source.chapters[chId].chLevel
            if source.chapters[chId].oldType is not None:
                self.chapters[chId].oldType = source.chapters[chId].oldType
            if source.chapters[chId].chType is not None:
                self.chapters[chId].chType = source.chapters[chId].chType
            if source.chapters[chId].isUnused is not None:
                self.chapters[chId].isUnused = source.chapters[chId].isUnused
            if source.chapters[chId].suppressChapterTitle is not None:
                self.chapters[chId].suppressChapterTitle = source.chapters[chId].suppressChapterTitle
            if source.chapters[chId].suppressChapterBreak is not None:
                self.chapters[chId].suppressChapterBreak = source.chapters[chId].suppressChapterBreak
            if source.chapters[chId].isTrash is not None:
                self.chapters[chId].isTrash = source.chapters[chId].isTrash

            #--- Merge the chapter's scene list.
            # New scenes may be added.
            # Existing scenes may be moved to another chapter.
            # Deletion of scenes is not considered.
            # The scene's sort order may not change.
            if source.chapters[chId].srtScenes is not None:

                # Remove scenes that have been moved to another chapter from the scene list.
                srtScenes = []
                for scId in self.chapters[chId].srtScenes:
                    if scId in source.chapters[chId].srtScenes or not scId in source.scenes:
                        srtScenes.append(scId)
                        # The scene has not moved to another chapter or isn't imported
                    self.chapters[chId].srtScenes = srtScenes

                # Add new or moved scenes to the scene list.
                merge_lists(source.chapters[chId].srtScenes, self.chapters[chId].srtScenes)

        #--- Merge project attributes.
        if source.title:
            # avoids deleting the title, if it is empty by accident
            self.title = source.title
        if source.desc is not None:
            self.desc = source.desc
        if source.authorName is not None:
            self.authorName = source.authorName
        if source.authorBio is not None:
            self.authorBio = source.authorBio
        if source.fieldTitle1 is not None:
            self.fieldTitle1 = source.fieldTitle1
        if source.fieldTitle2 is not None:
            self.fieldTitle2 = source.fieldTitle2
        if source.fieldTitle3 is not None:
            self.fieldTitle3 = source.fieldTitle3
        if source.fieldTitle4 is not None:
            self.fieldTitle4 = source.fieldTitle4

        # Add new chapters to the chapter list.
        # Deletion of chapters is not considered.
        # The sort order of chapters may not change.
        merge_lists(source.srtChapters, self.srtChapters)

        # Split scenes by inserted part/chapter/scene dividers.
        # This must be done after regular merging
        # in order to avoid creating duplicate IDs.
        if sourceHasSceneContent:
            sceneSplitter = Splitter()
            sceneSplitter.split_scenes(self)
        return 'yWriter project data updated or created.'

    def write(self):
        """Write instance variables to the yWriter xml file.
        
        Open the yWriter xml file located at filePath and replace the instance variables 
        not being None. Create new XML elements if necessary.
        Return a message beginning with the ERROR constant in case of error.
        Overrides the superclass method.
        """

        def build_scene_subtree(xmlScn, prjScn):
            if prjScn.title is not None:
                try:
                    xmlScn.find('Title').text = prjScn.title
                except(AttributeError):
                    ET.SubElement(xmlScn, 'Title').text = prjScn.title
            if xmlScn.find('BelongsToChID') is None:
                for chId in self.chapters:
                    if scId in self.chapters[chId].srtScenes:
                        ET.SubElement(xmlScn, 'BelongsToChID').text = chId
                        break

            if prjScn.desc is not None:
                try:
                    xmlScn.find('Desc').text = prjScn.desc
                except(AttributeError):
                    ET.SubElement(xmlScn, 'Desc').text = prjScn.desc
                       
            if xmlScn.find('SceneContent') is None:
                ET.SubElement(xmlScn, 'SceneContent').text = prjScn.sceneContent

            if xmlScn.find('WordCount') is None:
                ET.SubElement(xmlScn, 'WordCount').text = str(prjScn.wordCount)
            
            if xmlScn.find('LetterCount') is None:
                ET.SubElement(xmlScn, 'LetterCount').text = str(prjScn.letterCount)
            
            if prjScn.isUnused:
                if xmlScn.find('Unused') is None:
                    ET.SubElement(xmlScn, 'Unused').text = '-1'
            elif xmlScn.find('Unused') is not None:
                xmlScn.remove(xmlScn.find('Unused'))
            
            if prjScn.isNotesScene:
                scFields = xmlScn.find('Fields')
                try:
                    scFields.find('Field_SceneType').text = '1'
                except(AttributeError):
                    scFields = ET.SubElement(xmlScn, 'Fields')
                    ET.SubElement(scFields, 'Field_SceneType').text = '1'
            elif xmlScn.find('Fields') is not None:
                scFields = xmlScn.find('Fields')
                if scFields.find('Field_SceneType') is not None:
                    if scFields.find('Field_SceneType').text == '1':
                        scFields.remove(scFields.find('Field_SceneType'))
            
            if prjScn.isTodoScene:
                scFields = xmlScn.find('Fields')
                try:
                    scFields.find('Field_SceneType').text = '2'
                except(AttributeError):
                    scFields = ET.SubElement(xmlScn, 'Fields')
                    ET.SubElement(scFields, 'Field_SceneType').text = '2'
            elif xmlScn.find('Fields') is not None:
                scFields = xmlScn.find('Fields')
                if scFields.find('Field_SceneType') is not None:
                    if scFields.find('Field_SceneType').text == '2':
                        scFields.remove(scFields.find('Field_SceneType'))
            
            if prjScn.status is not None:
                try:
                    xmlScn.find('Status').text = str(prjScn.status)
                except:
                    ET.SubElement(xmlScn, 'Status').text = str(prjScn.status)
            
            if prjScn.sceneNotes is not None:
                try:
                    xmlScn.find('Notes').text = prjScn.sceneNotes
                except(AttributeError):
                    ET.SubElement(xmlScn, 'Notes').text = prjScn.sceneNotes
            
            if prjScn.tags is not None:
                try:
                    xmlScn.find('Tags').text = ';'.join(prjScn.tags)
                except(AttributeError):
                    ET.SubElement(xmlScn, 'Tags').text = ';'.join(prjScn.tags)
            
            if prjScn.field1 is not None:
                try:
                    xmlScn.find('Field1').text = prjScn.field1
                except(AttributeError):
                    ET.SubElement(xmlScn, 'Field1').text = prjScn.field1
            
            if prjScn.field2 is not None:
                try:
                    xmlScn.find('Field2').text = prjScn.field2
                except(AttributeError):
                    ET.SubElement(xmlScn, 'Field2').text = prjScn.field2
            
            if prjScn.field3 is not None:
                try:
                    xmlScn.find('Field3').text = prjScn.field3
                except(AttributeError):
                    ET.SubElement(xmlScn, 'Field3').text = prjScn.field3
            
            if prjScn.field4 is not None:
                try:
                    xmlScn.find('Field4').text = prjScn.field4
                except(AttributeError):
                    ET.SubElement(xmlScn, 'Field4').text = prjScn.field4
            
            if prjScn.appendToPrev:
                if xmlScn.find('AppendToPrev') is None:
                    ET.SubElement(xmlScn, 'AppendToPrev').text = '-1'
            elif xmlScn.find('AppendToPrev') is not None:
                xmlScn.remove(xmlScn.find('AppendToPrev'))

            # Date/time information
            if (prjScn.date is not None) and (prjScn.time is not None):
                dateTime = f'{prjScn.date} {prjScn.time}'
                if xmlScn.find('SpecificDateTime') is not None:
                    xmlScn.find('SpecificDateTime').text = dateTime
                else:
                    ET.SubElement(xmlScn, 'SpecificDateTime').text = dateTime
                    ET.SubElement(xmlScn, 'SpecificDateMode').text = '-1'

                    if xmlScn.find('Day') is not None:
                        xmlScn.remove(xmlScn.find('Day'))
                    
                    if xmlScn.find('Hour') is not None:
                        xmlScn.remove(xmlScn.find('Hour'))
                    
                    if xmlScn.find('Minute') is not None:
                        xmlScn.remove(xmlScn.find('Minute'))
            
            elif (prjScn.day is not None) or (prjScn.hour is not None) or (prjScn.minute is not None):
                
                if xmlScn.find('SpecificDateTime') is not None:
                    xmlScn.remove(xmlScn.find('SpecificDateTime'))
                
                if xmlScn.find('SpecificDateMode') is not None:
                    xmlScn.remove(xmlScn.find('SpecificDateMode'))
                if prjScn.day is not None:
                    try:
                        xmlScn.find('Day').text = prjScn.day
                    except(AttributeError):
                        ET.SubElement(xmlScn, 'Day').text = prjScn.day
                if prjScn.hour is not None:
                    try:
                        xmlScn.find('Hour').text = prjScn.hour
                    except(AttributeError):
                        ET.SubElement(xmlScn, 'Hour').text = prjScn.hour
                if prjScn.minute is not None:
                    try:
                        xmlScn.find('Minute').text = prjScn.minute
                    except(AttributeError):
                        ET.SubElement(xmlScn, 'Minute').text = prjScn.minute
            
            if prjScn.lastsDays is not None:
                try:
                    xmlScn.find('LastsDays').text = prjScn.lastsDays
                except(AttributeError):
                    ET.SubElement(xmlScn, 'LastsDays').text = prjScn.lastsDays
            
            if prjScn.lastsHours is not None:
                try:
                    xmlScn.find('LastsHours').text = prjScn.lastsHours
                except(AttributeError):
                    ET.SubElement(xmlScn, 'LastsHours').text = prjScn.lastsHours
            
            if prjScn.lastsMinutes is not None:
                try:
                    xmlScn.find('LastsMinutes').text = prjScn.lastsMinutes
                except(AttributeError):
                    ET.SubElement(xmlScn, 'LastsMinutes').text = prjScn.lastsMinutes

            # Plot related information
            if prjScn.isReactionScene:
                if xmlScn.find('ReactionScene') is None:
                    ET.SubElement(xmlScn, 'ReactionScene').text = '-1'
            elif xmlScn.find('ReactionScene') is not None:
                xmlScn.remove(xmlScn.find('ReactionScene'))
            
            if prjScn.isSubPlot:
                if xmlScn.find('SubPlot') is None:
                    ET.SubElement(xmlScn, 'SubPlot').text = '-1'
            elif xmlScn.find('SubPlot') is not None:
                xmlScn.remove(xmlScn.find('SubPlot'))
            
            if prjScn.goal is not None:
                try:
                    xmlScn.find('Goal').text = prjScn.goal
                except(AttributeError):
                    ET.SubElement(xmlScn, 'Goal').text = prjScn.goal
            
            if prjScn.conflict is not None:
                try:
                    xmlScn.find('Conflict').text = prjScn.conflict
                except(AttributeError):
                    ET.SubElement(xmlScn, 'Conflict').text = prjScn.conflict
            
            if prjScn.outcome is not None:
                try:
                    xmlScn.find('Outcome').text = prjScn.outcome
                except(AttributeError):
                    ET.SubElement(xmlScn, 'Outcome').text = prjScn.outcome
            
            if prjScn.image is not None:
                try:
                    xmlScn.find('ImageFile').text = prjScn.image
                except(AttributeError):
                    ET.SubElement(xmlScn, 'ImageFile').text = prjScn.image
                    
            # Characters/locations/items           
            if prjScn.characters is not None:
                characters = xmlScn.find('Characters')
                try:
                    for oldCrId in characters.findall('CharID'):
                        characters.remove(oldCrId)
                except(AttributeError):
                    characters = ET.SubElement(xmlScn, 'Characters')
                for crId in prjScn.characters:
                    ET.SubElement(characters, 'CharID').text = crId
            
            if prjScn.locations is not None:
                locations = xmlScn.find('Locations')
                try:
                    for oldLcId in locations.findall('LocID'):
                        locations.remove(oldLcId)
                except(AttributeError):
                    locations = ET.SubElement(xmlScn, 'Locations')
                for lcId in prjScn.locations:
                    ET.SubElement(locations, 'LocID').text = lcId
            
            if prjScn.items is not None:
                items = xmlScn.find('Items')
                try:
                    for oldItId in items.findall('ItemID'):
                        items.remove(oldItId)
                except(AttributeError):
                    items = ET.SubElement(xmlScn, 'Items')
                for itId in prjScn.items:
                    ET.SubElement(items, 'ItemID').text = itId

        def build_chapter_subtree(xmlChp, prjChp, sortOrder):
            try:
                xmlChp.find('SortOrder').text = str(sortOrder)
            except(AttributeError):
                ET.SubElement(xmlChp, 'SortOrder').text = str(sortOrder)
            try:
                xmlChp.find('Title').text = prjChp.title
            except(AttributeError):
                ET.SubElement(xmlChp, 'Title').text = prjChp.title
            
            if prjChp.desc is not None:
                try:
                    xmlChp.find('Desc').text = prjChp.desc
                except(AttributeError):
                    ET.SubElement(xmlChp, 'Desc').text = prjChp.desc
            
            if xmlChp.find('SectionStart') is not None:
                if prjChp.chLevel == 0:
                    xmlChp.remove(xmlChp.find('SectionStart'))
            elif prjChp.chLevel == 1:
                ET.SubElement(xmlChp, 'SectionStart').text = '-1'
            
            if prjChp.oldType is not None:
                try:
                    xmlChp.find('Type').text = str(prjChp.oldType)
                except(AttributeError):
                    ET.SubElement(xmlChp, 'Type').text = str(prjChp.oldType)

            if prjChp.chType is not None:
                try:
                    xmlChp.find('ChapterType').text = str(prjChp.chType)
                except(AttributeError):
                    ET.SubElement(xmlChp, 'ChapterType').text = str(prjChp.chType)

            if prjChp.isUnused:
                if xmlChp.find('Unused') is None:
                    ET.SubElement(xmlChp, 'Unused').text = '-1'
            elif xmlChp.find('Unused') is not None:
                xmlChp.remove(xmlChp.find('Unused'))

            #--- Rebuild the chapter's scene list.

            if prjChp.srtScenes:
                xScnList = xmlChp.find('Scenes')
                if xScnList is not None:
                    xmlChp.remove(xScnList)
                sortSc = ET.SubElement(xmlChp, 'Scenes')
                for scId in prjChp.srtScenes:
                    ET.SubElement(sortSc, 'ScID').text = scId

        def build_location_subtree(xmlLoc, prjLoc, sortOrder):
            ET.SubElement(xmlLoc, 'ID').text = lcId

            if prjLoc.title is not None:
                ET.SubElement(xmlLoc, 'Title').text = prjLoc.title

            if prjLoc.image is not None:
                ET.SubElement(xmlLoc, 'ImageFile').text = prjLoc.image

            if prjLoc.desc is not None:
                ET.SubElement(xmlLoc, 'Desc').text = prjLoc.desc

            if prjLoc.aka is not None:
                ET.SubElement(xmlLoc, 'AKA').text = prjLoc.aka

            if prjLoc.tags is not None:
                ET.SubElement(xmlLoc, 'Tags').text = ';'.join(prjLoc.tags)

            ET.SubElement(xmlLoc, 'SortOrder').text = str(sortOrder)

        def build_item_subtree(xmlItm, prjItm, sortOrder):
            ET.SubElement(xmlItm, 'ID').text = itId

            if prjItm.title is not None:
                ET.SubElement(xmlItm, 'Title').text = prjItm.title

            if prjItm.image is not None:
                ET.SubElement(xmlItm, 'ImageFile').text = prjItm.image

            if prjItm.desc is not None:
                ET.SubElement(xmlItm, 'Desc').text = prjItm.desc

            if prjItm.aka is not None:
                ET.SubElement(xmlItm, 'AKA').text = prjItm.aka

            if prjItm.tags is not None:
                ET.SubElement(xmlItm, 'Tags').text = ';'.join(prjItm.tags)

            ET.SubElement(xmlItm, 'SortOrder').text = str(sortOrder)

        def build_character_subtree(xmlCrt, prjCrt, sortOrder):
            ET.SubElement(xmlCrt, 'ID').text = crId

            if prjCrt.title is not None:
                ET.SubElement(xmlCrt, 'Title').text = prjCrt.title

            if prjCrt.desc is not None:
                ET.SubElement(xmlCrt, 'Desc').text = prjCrt.desc

            if prjCrt.image is not None:
                ET.SubElement(xmlCrt, 'ImageFile').text = prjCrt.image

            ET.SubElement(xmlCrt, 'SortOrder').text = str(sortOrder)

            if prjCrt.notes is not None:
                ET.SubElement(xmlCrt, 'Notes').text = prjCrt.notes

            if prjCrt.aka is not None:
                ET.SubElement(xmlCrt, 'AKA').text = prjCrt.aka

            if prjCrt.tags is not None:
                ET.SubElement(xmlCrt, 'Tags').text = ';'.join(prjCrt.tags)

            if prjCrt.bio is not None:
                ET.SubElement(xmlCrt, 'Bio').text = prjCrt.bio

            if prjCrt.goals is not None:
                ET.SubElement(xmlCrt, 'Goals').text = prjCrt.goals

            if prjCrt.fullName is not None:
                ET.SubElement(xmlCrt, 'FullName').text = prjCrt.fullName

            if prjCrt.isMajor:
                ET.SubElement(xmlCrt, 'Major').text = '-1'

        def build_project_subtree(xmlPrj):
            VER = '7'
            try:
                xmlPrj.find('Ver').text = VER
            except(AttributeError):
                ET.SubElement(xmlPrj, 'Ver').text = VER

            if self.title is not None:
                try:
                    xmlPrj.find('Title').text = self.title
                except(AttributeError):
                    ET.SubElement(xmlPrj, 'Title').text = self.title

            if self.desc is not None:
                try:
                    xmlPrj.find('Desc').text = self.desc
                except(AttributeError):
                    ET.SubElement(xmlPrj, 'Desc').text = self.desc

            if self.authorName is not None:
                try:
                    xmlPrj.find('AuthorName').text = self.authorName
                except(AttributeError):
                    ET.SubElement(xmlPrj, 'AuthorName').text = self.authorName

            if self.authorBio is not None:
                try:
                    xmlPrj.find('Bio').text = self.authorBio
                except(AttributeError):
                    ET.SubElement(xmlPrj, 'Bio').text = self.authorBio

            if self.fieldTitle1 is not None:
                try:
                    xmlPrj.find('FieldTitle1').text = self.fieldTitle1
                except(AttributeError):
                    ET.SubElement(xmlPrj, 'FieldTitle1').text = self.fieldTitle1

            if self.fieldTitle2 is not None:
                try:
                    xmlPrj.find('FieldTitle2').text = self.fieldTitle2
                except(AttributeError):
                    ET.SubElement(xmlPrj, 'FieldTitle2').text = self.fieldTitle2

            if self.fieldTitle3 is not None:
                try:
                    xmlPrj.find('FieldTitle3').text = self.fieldTitle3
                except(AttributeError):
                    ET.SubElement(xmlPrj, 'FieldTitle3').text = self.fieldTitle3

            if self.fieldTitle4 is not None:
                try:
                    xmlPrj.find('FieldTitle4').text = self.fieldTitle4
                except(AttributeError):
                    ET.SubElement(xmlPrj, 'FieldTitle4').text = self.fieldTitle4

        #--- Start write method.
        if self.is_locked():
            return f'{ERROR}yWriter seems to be open. Please close first.'

        TAG = 'YWRITER7'
        xmlScenes = {}
        xmlChapters = {}
        try:
            root = self.tree.getroot()
            xmlPrj = root.find('PROJECT')
            locations = root.find('LOCATIONS')
            items = root.find('ITEMS')
            characters = root.find('CHARACTERS')
            scenes = root.find('SCENES')
            chapters = root.find('CHAPTERS')
        except(AttributeError):
            root = ET.Element(TAG)
            xmlPrj = ET.SubElement(root, 'PROJECT')
            locations = ET.SubElement(root, 'LOCATIONS')
            items = ET.SubElement(root, 'ITEMS')
            characters = ET.SubElement(root, 'CHARACTERS')
            scenes = ET.SubElement(root, 'SCENES')
            chapters = ET.SubElement(root, 'CHAPTERS')

        #--- Process project attributes.

        build_project_subtree(xmlPrj)

        #--- Process locations.

        # Remove LOCATION entries in order to rewrite
        # the LOCATIONS section in a modified sort order.
        for xmlLoc in locations.findall('LOCATION'):
            locations.remove(xmlLoc)

        # Add the new XML location subtrees to the project tree.
        sortOrder = 0
        for lcId in self.srtLocations:
            sortOrder += 1
            xmlLoc = ET.SubElement(locations, 'LOCATION')
            build_location_subtree(xmlLoc, self.locations[lcId], sortOrder)

        #--- Process items.

        # Remove ITEM entries in order to rewrite
        # the ITEMS section in a modified sort order.
        for xmlItm in items.findall('ITEM'):
            items.remove(xmlItm)

        # Add the new XML item subtrees to the project tree.
        sortOrder = 0
        for itId in self.srtItems:
            sortOrder += 1
            xmlItm = ET.SubElement(items, 'ITEM')
            build_item_subtree(xmlItm, self.items[itId], sortOrder)

        #--- Process characters.

        # Remove CHARACTER entries in order to rewrite
        # the CHARACTERS section in a modified sort order.
        for xmlCrt in characters.findall('CHARACTER'):
            characters.remove(xmlCrt)

        # Add the new XML character subtrees to the project tree.
        sortOrder = 0
        for crId in self.srtCharacters:
            sortOrder += 1
            xmlCrt = ET.SubElement(characters, 'CHARACTER')
            build_character_subtree(xmlCrt, self.characters[crId], sortOrder)

        #--- Process scenes.

        # Save the original XML scene subtrees
        # and remove them from the project tree.
        for xmlScn in scenes.findall('SCENE'):
            scId = xmlScn.find('ID').text
            xmlScenes[scId] = xmlScn
            scenes.remove(xmlScn)

        # Add the new XML scene subtrees to the project tree.
        for scId in self.scenes:
            if not scId in xmlScenes:
                xmlScenes[scId] = ET.Element('SCENE')
                ET.SubElement(xmlScenes[scId], 'ID').text = scId
            build_scene_subtree(xmlScenes[scId], self.scenes[scId])
            scenes.append(xmlScenes[scId])

        #--- Process chapters.

        # Save the original XML chapter subtree
        # and remove it from the project tree.
        for xmlChp in chapters.findall('CHAPTER'):
            chId = xmlChp.find('ID').text
            xmlChapters[chId] = xmlChp
            chapters.remove(xmlChp)

        # Add the new XML chapter subtrees to the project tree.
        sortOrder = 0
        for chId in self.srtChapters:
            sortOrder += 1
            if not chId in xmlChapters:
                xmlChapters[chId] = ET.Element('CHAPTER')
                ET.SubElement(xmlChapters[chId], 'ID').text = chId
            build_chapter_subtree(xmlChapters[chId], self.chapters[chId], sortOrder)
            chapters.append(xmlChapters[chId])
        indent(root)

        # Modify the scene contents of an existing xml element tree.
        for scn in root.iter('SCENE'):
            scId = scn.find('ID').text
            if self.scenes[scId].sceneContent is not None:
                scn.find('SceneContent').text = self.scenes[scId].sceneContent
                scn.find('WordCount').text = str(self.scenes[scId].wordCount)
                scn.find('LetterCount').text = str(self.scenes[scId].letterCount)
            try:
                scn.remove(scn.find('RTFFile'))
            except:
                pass
        self.tree = ET.ElementTree(root)
        message = self._write_element_tree(self)
        if message.startswith(ERROR):
            return message

        return self._postprocess_xml_file(self.filePath)

    def is_locked(self):
        """Check whether the yw7 file is locked by yWriter.
        
        Return True if a .lock file placed by yWriter exists.
        Otherwise, return False. 
        """
        return os.path.isfile(f'{self.filePath}.lock')
    
    def _write_element_tree(self, ywProject):
        """Write back the xml element tree to a .yw7 xml file located at filePath.
        
        Return a message beginning with the ERROR constant in case of error.
        """
        if os.path.isfile(ywProject.filePath):
            os.replace(ywProject.filePath, f'{ywProject.filePath}.bak')
            backedUp = True
        else:
            backedUp = False
        try:
            ywProject.tree.write(ywProject.filePath, xml_declaration=False, encoding='utf-8')
        except:
            if backedUp:
                os.replace(f'{ywProject.filePath}.bak', ywProject.filePath)
            return f'{ERROR}Cannot write "{os.path.normpath(ywProject.filePath)}".'

        return 'yWriter XML tree written.'


    def _postprocess_xml_file(self, filePath):
        '''Postprocess an xml file created by ElementTree.
        
        Positional argument:
            filePath -- str: path to xml file.
        
        Read the xml file, put a header on top, insert the missing CDATA tags,
        and replace xml entities by plain text (unescape). Overwrite the .yw7 xml file.
        Return a message beginning with the ERROR constant in case of error.
        
        Note: The path is given as an argument rather than using self.filePath. 
        So this routine can be used for yWriter-generated xml files other than .yw7 as well. 
        '''
        with open(filePath, 'r', encoding='utf-8') as f:
            text = f.read()
        lines = text.split('\n')
        newlines = ['<?xml version="1.0" encoding="utf-8"?>']
        for line in lines:
            for tag in self._CDATA_TAGS:
                line = re.sub(f'\<{tag}\>', f'<{tag}><![CDATA[', line)
                line = re.sub(f'\<\/{tag}\>', f']]></{tag}>', line)
            newlines.append(line)
        text = '\n'.join(newlines)
        text = text.replace('[CDATA[ \n', '[CDATA[')
        text = text.replace('\n]]', ']]')
        text = unescape(text)
        try:
            with open(filePath, 'w', encoding='utf-8') as f:
                f.write(text)
        except:
            return f'{ERROR}Can not write "{os.path.normpath(filePath)}".'

        return f'"{os.path.normpath(filePath)}" written.'

    def _strip_spaces(self, lines):
        """Local helper method.

        Positional argument:
            lines -- list of strings

        Return lines with leading and trailing spaces removed.
        """
        stripped = []
        for line in lines:
            stripped.append(line.strip())
        return stripped

from html.parser import HTMLParser


def read_html_file(filePath):
    """Open a html file being encoded utf-8 or ANSI.
    
    Return a tuple:
    message = Message beginning with the ERROR constant in case of error.
    content = The file content in a single string. None in case of error.
    """
    try:
        with open(filePath, 'r', encoding='utf-8') as f:
            content = f.read()
    except:
        # HTML files exported by a word processor may be ANSI encoded.
        try:
            with open(filePath, 'r') as f:
                content = (f.read())
        except(FileNotFoundError):
            return f'{ERROR}"{os.path.normpath(filePath)}" not found.', None

    return 'HTML data read in.', content


class HtmlFile(Novel, HTMLParser):
    """Generic HTML file representation.
    
    Public methods:
        handle_starttag -- identify scenes and chapters.
        handle comment --
        read --
    """
    EXTENSION = '.html'
    _COMMENT_START = '/*'
    _COMMENT_END = '*/'
    _SC_TITLE_BRACKET = '~'

    def __init__(self, filePath, **kwargs):
        """Initialize the HTML parser and local instance variables for parsing.
        
        Positional arguments:
            filePath -- str: path to the file represented by the Novel instance.
            
        Optional arguments:
            kwargs -- keyword arguments to be used by subclasses.            

        The HTML parser works like a state machine. 
        Scene ID, chapter ID and processed lines must be saved between the transitions.         
        Extends the superclass constructor.
        """
        super().__init__(filePath)
        HTMLParser.__init__(self)
        self._lines = []
        self._scId = None
        self._chId = None

    def _convert_to_yw(self, text):
        """Convert html formatting tags to yWriter 7 raw markup.
        
        Positional arguments:
            text -- string to convert.
        
        Return a yw7 markup string.
        Overrides the superclass method.
        """

        #--- Clean up polluted HTML code.
        text = re.sub('</*font.*?>', '', text)
        text = re.sub('</*span.*?>', '', text)
        text = re.sub('</*FONT.*?>', '', text)
        text = re.sub('</*SPAN.*?>', '', text)

        #--- Put everything in one line.
        text = text.replace('\n', ' ')
        text = text.replace('\r', ' ')
        text = text.replace('\t', ' ')
        while '  ' in text:
            text = text.replace('  ', ' ').strip()

        #--- Replace HTML tags by yWriter markup.
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

        #--- Remove orphaned tags.
        text = text.replace('[/b][b]', '')
        text = text.replace('[/i][i]', '')
        text = text.replace('[/b][b]', '')
        return text

    def _preprocess(self, text):
        """Clean up the HTML code and strip yWriter 7 raw markup.
        
        Positional arguments:
            text -- str: HTML text to be processed.
        
        This prevents accidentally applied formatting from being transferred to the yWriter metadata.
        If rich text is applicable, such as in scenes, overwrite this method in a subclass.
        Return a string.
        """
        text = self._convert_to_yw(text)

        #--- Remove misplaced formatting tags.
        text = re.sub('\[\/*[b|i]\]', '', text)
        return text

    def _postprocess(self):
        """Process the plain text after parsing.
        
        This is a hook for subclasses.
        """

    def handle_starttag(self, tag, attrs):
        """Identify scenes and chapters.
        
        Positional arguments:
            tag -- str: name of the tag converted to lower case.
            attrs -- list of (name, value) pairs containing the attributes found inside the tag’s <> brackets.
        
        Overrides HTMLparser.handle_starttag() called by the parser to handle the start of a tag. 
        This method is applicable to HTML files that are divided into chapters and scenes. 
        For differently structured HTML files  do override this method in a subclass.
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

    def handle_comment(self, data):
        """Process inline comments within scene content.
        
        Positional arguments:
            data -- str: comment text. 
        
        Overrides HTMLparser.handle_comment() called by the parser when a comment is encountered.
        """
        if self._scId is not None: 
            self._lines.append(f'{self._COMMENT_START}{data}{self._COMMENT_END}')

    def read(self):
        """Parse the file and get the instance variables.
        
        Return a message beginning with the ERROR constant in case of error.
        This is a template method for subclasses tailored to the 
        content of the respective HTML file.
        """
        message, content = read_html_file(self._filePath)
        if message.startswith(ERROR):
            return message

        content = self._preprocess(content)
        self.feed(content)
        self._postprocess()
        return 'Created novel structure from HTML data.'


class HtmlImport(HtmlFile):
    """HTML 'work in progress' file representation.

    Import untagged chapters and scenes.
    """
    DESCRIPTION = 'Work in progress'
    SUFFIX = ''
    _SCENE_DIVIDER = '* * *'
    _LOW_WORDCOUNT = 10

    def __init__(self, filePath, **kwargs):
        """Initialize local instance variables for parsing.

        Positional arguments:
            filePath -- str: path to the file represented by the Novel instance.
            
        The HTML parser works like a state machine. 
        Chapter and scene count must be saved between the transitions.         
        Extends the superclass constructor.
        """
        super().__init__(filePath)
        self._chCount = 0
        self._scCount = 0

    def _preprocess(self, text):
        """Process the html text before parsing.
        
        Convert html formatting tags to yWriter 7 raw markup.
        Overrides the superclass method.
        """
        return self._convert_to_yw(text)

    def handle_starttag(self, tag, attrs):
        """Recognize the paragraph's beginning.
        
        Positional arguments:
            tag -- str: name of the tag converted to lower case.
            attrs -- list of (name, value) pairs containing the attributes found inside the tag’s <> brackets.
        
        Overrides the superclass method.
        """
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
                self.scenes[self._scId].title = f'Scene {self._scCount}'
        elif tag == 'div':
            self._scId = None
            self._chId = None
        elif tag == 'meta':
            if attrs[0][1].lower() == 'author':
                self.authorName = attrs[1][1]
            if attrs[0][1].lower() == 'description':
                self.desc = attrs[1][1]
        elif tag == 'title':
            self._lines = []

    def handle_endtag(self, tag):
        """Recognize the paragraph's end.
        
        Positional arguments:
            tag -- str: name of the tag converted to lower case.

        Overrides HTMLparser.handle_endtag() called by the HTML parser to handle the end tag of an element.
        """
        if tag == 'p':
            self._lines.append('\n')
            if self._scId is not None:
                self.scenes[self._scId].sceneContent = ''.join(self._lines)
                if self.scenes[self._scId].wordCount < self._LOW_WORDCOUNT:
                    self.scenes[self._scId].status = Scene.STATUS.index('Outline')
                else:
                    self.scenes[self._scId].status = Scene.STATUS.index('Draft')
        elif tag in ('h1', 'h2'):
            self.chapters[self._chId].title = ''.join(self._lines)
            self._lines = []
        elif tag == 'title':
            self.title = ''.join(self._lines)

    def handle_data(self, data):
        """Collect data within scene sections.

        Positional arguments:
            data -- str: text to be stored. 
        
        Overrides HTMLparser.handle_data() called by the parser when a comment is encountered.
        """
        if self._scId is not None and self._SCENE_DIVIDER in data:
            self._scId = None
        else:
            data = data.strip()

            # Convert prefixed comment into scene title.
            if not self._lines and data.startswith(self._COMMENT_START):
                try:
                    scTitle, scContent = data.split(
                        sep=self._COMMENT_END, maxsplit=1)
                    if self._SC_TITLE_BRACKET in scTitle:
                        scTitle = scTitle.split(self._SC_TITLE_BRACKET)[1]
                    else:
                        scTitle = scTitle.lstrip(self._COMMENT_START)
                    self.scenes[self._scId].title = scTitle.strip()
                    data = scContent
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
        """Initialize local instance variables for parsing.

        Positional arguments:
            filePath -- str: path to the file represented by the Novel instance.
            
        The HTML parser works like a state machine. 
        Chapter and scene count must be saved between the transitions.         
        Extends the superclass constructor.
        """
        super().__init__(filePath)
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
                self.authorName = attrs[1][1]
            if attrs[0][1].lower() == 'description':
                self.desc = attrs[1][1]
        elif tag == 'title':
            self._lines = []

    def handle_endtag(self, tag):
        """Recognize the paragraph's end.
        
        Positional arguments:
            tag -- str: name of the tag converted to lower case.

        Overrides HTMLparser.handle_endtag() called by the HTML parser to handle the end tag of an element.
        """
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

        Positional arguments:
            data -- str: text to be stored. 
        
        Overrides HTMLparser.handle_data() called by the parser when a comment is encountered.
        """
        self._lines.append(data.strip())


class NewProjectFactory(FileFactory):
    """A factory class that instantiates a document object to read, 
    and a new yWriter project.

    Public methods:
        make_file_objects(self, sourcePath, **kwargs) -- return conversion objects.

    Class constant:
        DO_NOT_IMPORT -- list of suffixes from file classes not meant to be imported.    
    """
    DO_NOT_IMPORT = ['_xref', '_brf_synopsis']

    def make_file_objects(self, sourcePath, **kwargs):
        """Instantiate a source and a target object for creation of a new yWriter project.

        Positional arguments:
            sourcePath -- str: path to the source file to convert.

        Return a tuple with three elements:
        - A message beginning with the ERROR constant in case of error
        - sourceFile: a Novel subclass instance
        - targetFile: a Novel subclass instance
        """
        if not self._canImport(sourcePath):
            return f'{ERROR}This document is not meant to be written back.', None, None
        
        fileName, __ = os.path.splitext(sourcePath)
        targetFile = Yw7File(f'{fileName}{Yw7File.EXTENSION}', **kwargs)
        if sourcePath.endswith('.html'):
            # The source file might be an outline or a "work in progress".
            message, content = read_html_file(sourcePath)
            if message.startswith(ERROR):
                return message, None, None
            
            if "<h3" in content.lower():
                sourceFile = HtmlOutline(sourcePath, **kwargs)
            else:
                sourceFile = HtmlImport(sourcePath, **kwargs)
            return 'Source and target objects created.', sourceFile, targetFile
            
        else:
            for fileClass in self._fileClasses:
                if fileClass.SUFFIX is not None:
                    if sourcePath.endswith(f'{fileClass.SUFFIX}{fileClass.EXTENSION}'):
                        sourceFile = fileClass(sourcePath, **kwargs)
                        return 'Source and target objects created.', sourceFile, targetFile
                    
            return f'{ERROR}File type of "{os.path.normpath(sourcePath)}" not supported.', None, None

    def _canImport(self, sourcePath):
        """Check whether the source file can be imported to yWriter.
        
        Positional arguments: 
            sourcePath -- str: path of the file to be ckecked.
        
        Return True, if the file located at sourcepath is of an importable type.
        Otherwise, return False.
        """
        fileName, __ = os.path.splitext(sourcePath)
        for suffix in self.DO_NOT_IMPORT:
            if fileName.endswith(suffix):
                return False
        
        return True


class OdtExport(OdtFile):
    """ODT novel file representation.

    Export a non-reimportable manuscript with chapters and scenes.
    """
    _fileHeader = f'''{OdtFile._CONTENT_XML_HEADER}<text:p text:style-name="Title">$Title</text:p>
<text:p text:style-name="Subtitle">$AuthorName</text:p>
'''

    _partTemplate = '''<text:h text:style-name="Heading_20_1" text:outline-level="1">$Title</text:h>
'''

    _chapterTemplate = '''<text:h text:style-name="Heading_20_2" text:outline-level="2">$Title</text:h>
'''

    _sceneTemplate = '''<text:p text:style-name="Text_20_body"><office:annotation>
<dc:creator>scene title</dc:creator>
<text:p>~ ${Title} ~</text:p>
</office:annotation>$SceneContent</text:p>
'''

    _appendedSceneTemplate = '''<text:p text:style-name="First_20_line_20_indent"><office:annotation>
<dc:creator>scene title</dc:creator>
<text:p>~ ${Title} ~</text:p>
</office:annotation>$SceneContent</text:p>
'''

    _sceneDivider = '<text:p text:style-name="Heading_20_4">* * *</text:p>\n'
    _fileFooter = OdtFile._CONTENT_XML_FOOTER

    def _get_chapterMapping(self, chId, chapterNumber):
        """Return a mapping dictionary for a chapter section.
        
        Positional arguments:
            chId -- str: chapter ID.
            chapterNumber -- int: chapter number.
        
        Suppress the chapter title if necessary.
        Extends the superclass method.
        """
        chapterMapping = super()._get_chapterMapping(chId, chapterNumber)
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
        """Initialize local instance variables for parsing.

        Positional arguments:
            filePath -- str: path to the file represented by the Novel instance.
            
        The HTML parser works like a state machine. 
        A prefix for chapter and scene recognition must be saved between the transitions.         
        Extends the superclass constructor.
        """
        super().__init__(filePath)
        self._prefix = None

    def _preprocess(self, text):
        """Process the html text before parsing.
        
        Convert html formatting tags to yWriter 7 raw markup.
        Overrides the superclass method.
        """
        return self._convert_to_yw(text)

    def _postprocess(self):
        """Parse the converted text to identify chapters and scenes.
        
        Overrides the superclass method.
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
        
        Positional arguments:
            tag -- str: name of the tag converted to lower case.
            attrs -- list of (name, value) pairs containing the attributes found inside the tag’s <> brackets.
        
        Overrides the superclass method.
        """
        if tag == 'p':
            self._prefix = ''
        elif tag == 'h2':
            self._prefix = Splitter.CHAPTER_SEPARATOR
        elif tag == 'h1':
            self._prefix = Splitter.PART_SEPARATOR

    def handle_endtag(self, tag):
        """Recognize the paragraph's end.      
        
        Positional arguments:
            tag -- str: name of the tag converted to lower case.

        Overrides HTMLparser.handle_endtag() called by the HTML parser to handle the end tag of an element.
        """
        if tag in ['p', 'h2', 'h1']:
            self._prefix = None

    def handle_data(self, data):
        """Copy the scene paragraphs.      

        Positional arguments:
            data -- str: text to be stored. 
        
        Overrides HTMLparser.handle_data() called by the parser when a comment is encountered.
        """
        if self._prefix is not None:
            self._lines.append(f'{self._prefix}{data}')


class HtmlManuscript(HtmlFile):
    """HTML manuscript file representation.

    Import a manuscript with invisibly tagged chapters and scenes.
    """
    DESCRIPTION = 'Editable manuscript'
    SUFFIX = '_manuscript'

    def _preprocess(self, text):
        """Process the html text before parsing.
        
        Convert html formatting tags to yWriter 7 raw markup.
        Overrides the superclass method.
        """
        return self._convert_to_yw(text)

    def handle_starttag(self, tag, attrs):
        """Identify scenes and chapters.
        
        Positional arguments:
            tag -- str: name of the tag converted to lower case.
            attrs -- list of (name, value) pairs containing the attributes found inside the tag’s <> brackets.
        
        Extends the superclass method by processing inline chapter and scene dividers.
        """
        super().handle_starttag(tag, attrs)
        if self._scId is not None:
            if tag == 'h1':
                self._lines.append(Splitter.PART_SEPARATOR)
            elif tag == 'h2':
                self._lines.append(Splitter.CHAPTER_SEPARATOR)

    def handle_comment(self, data):
        """Process inline comments within scene content.
        
        Positional arguments:
            data -- str: comment text. 
        
        Use marked comments at scene start as scene titles.
        Overrides the superclass method.
        """
        if self._scId is not None: 
            if not self._lines:
                # Comment is at scene start
                if self._SC_TITLE_BRACKET in data:
                    # Comment is marked as a scene title
                    try:   
                        self.scenes[self._scId].title = data.split(self._SC_TITLE_BRACKET)[1].strip()
                    except:
                        pass
                    return

            self._lines.append(f'{self._COMMENT_START}{data.strip()}{self._COMMENT_END}')
            

    def handle_endtag(self, tag):
        """Recognize the end of the scene section and save data.
        
        Positional arguments:
            tag -- str: name of the tag converted to lower case.

        Overrides HTMLparser.handle_endtag() called by the HTML parser to handle the end tag of an element.
        """
        if self._scId is not None:
            if tag == 'div':
                text = ''.join(self._lines)
                self.scenes[self._scId].sceneContent = text
                self._lines = []
                self._scId = None
            elif tag == 'p':
                self._lines.append('\n')
            elif tag == 'h1':
                self._lines.append('\n')
            elif tag == 'h2':
                self._lines.append('\n')
        elif self._chId is not None:
            if tag == 'div':
                self._chId = None

    def handle_data(self, data):
        """Collect data within scene sections.

        Positional arguments:
            data -- str: text to be stored. 
        
        Overrides HTMLparser.handle_data() called by the parser when a comment is encountered.
        """
        if self._scId is not None:
            if not data.isspace():
                self._lines.append(data)
        elif self._chId is not None:
            if not self.chapters[self._chId].title:
                self.chapters[self._chId].title = data.strip()


class HtmlSceneDesc(HtmlFile):
    """HTML scene summaries file representation.

    Import a full synopsis with invisibly tagged scene descriptions.
    """
    DESCRIPTION = 'Scene descriptions'
    SUFFIX = '_scenes'

    def handle_endtag(self, tag):
        """Recognize the end of the scene section and save data.
        
        Positional arguments:
            tag -- str: name of the tag converted to lower case.

        Overrides HTMLparser.handle_endtag() called by the HTML parser to handle the end tag of an element.
        """
        if self._scId is not None:
            if tag == 'div':
                text = ''.join(self._lines)
                if text.startswith(self._COMMENT_START):
                    try:
                        scTitle, scContent = text.split(
                            sep=self._COMMENT_END, maxsplit=1)
                        if self._SC_TITLE_BRACKET in scTitle:
                            self.scenes[self._scId].title = scTitle.split(
                                self._SC_TITLE_BRACKET)[1].strip()
                        text = scContent
                    except:
                        pass
                self.scenes[self._scId].desc = text
                self._lines = []
                self._scId = None
            elif tag == 'p':
                self._lines.append('\n')
        elif self._chId is not None:
            if tag == 'div':
                self._chId = None

    def handle_data(self, data):
        """Collect data within scene sections.

        Positional arguments:
            data -- str: text to be stored. 
        
        Overrides HTMLparser.handle_data() called by the parser when a comment is encountered.
        """
        if self._scId is not None:
            self._lines.append(data.strip())
        elif self._chId is not None:
            if not self.chapters[self._chId].title:
                self.chapters[self._chId].title = data.strip()


class HtmlChapterDesc(HtmlFile):
    """HTML chapter summaries file representation.

    Import a brief synopsis with invisibly tagged chapter descriptions.
    """
    DESCRIPTION = 'Chapter descriptions'
    SUFFIX = '_chapters'

    def handle_endtag(self, tag):
        """Recognize the end of the chapter section and save data.
        
        Positional arguments:
            tag -- str: name of the tag converted to lower case.

        Overrides HTMLparser.handle_endtag() called by the HTML parser to handle the end tag of an element.
        """
        if self._chId is not None:
            if tag == 'div':
                self.chapters[self._chId].desc = ''.join(self._lines)
                self._lines = []
                self._chId = None
            elif tag == 'p':
                self._lines.append('\n')
            elif tag == 'h1' or tag == 'h2':
                if not self.chapters[self._chId].title:
                    self.chapters[self._chId].title = ''.join(self._lines)
                    self._lines = []

    def handle_data(self, data):
        """Collect data within chapter sections.

        Positional arguments:
            data -- str: text to be stored. 
        
        Overrides HTMLparser.handle_data() called by the parser when a comment is encountered.
        """
        if self._chId is not None:
            self._lines.append(data.strip())


class HtmlPartDesc(HtmlChapterDesc):
    """HTML part summaries file representation.

    Parts are chapters marked in yWriter as beginning of a new section.
    Import a synopsis with invisibly tagged part descriptions.
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
        """Initialize local instance variables for parsing.

        Positional arguments:
            filePath -- str: path to the file represented by the Novel instance.
            
        The HTML parser works like a state machine. 
        Character ID and section title must be saved between the transitions.         
        Extends the superclass constructor.
        """
        HtmlFile.__init__(self, filePath)
        self._crId = None
        self._section = None

    def handle_starttag(self, tag, attrs):
        """Identify characters with subsections.
        
        Positional arguments:
            tag -- str: name of the tag converted to lower case.
            attrs -- list of (name, value) pairs containing the attributes found inside the tag’s <> brackets.
        
        Overrides the superclass method.
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
                elif attrs[0][1].startswith('CrID_notes'):
                    self._section = 'notes'

    def handle_endtag(self, tag):
        """Recognize the end of the character section and save data.
        
        Positional arguments:
            tag -- str: name of the tag converted to lower case.

        Overrides HTMLparser.handle_endtag() called by the HTML parser to handle the end tag of an element.
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
                elif self._section == 'notes':
                    self.characters[self._crId].notes = ''.join(self._lines)
                    self._lines = []
                    self._section = None
            elif tag == 'p':
                self._lines.append('\n')

    def handle_data(self, data):
        """collect data within character sections.

        Positional arguments:
            data -- str: text to be stored. 
        
        Overrides HTMLparser.handle_data() called by the parser when a comment is encountered.
        """
        if self._section is not None:
            self._lines.append(data.strip())


class HtmlLocations(HtmlFile):
    """HTML location descriptions file representation.

    Import a location sheet with invisibly tagged descriptions.
    """
    DESCRIPTION = 'Location descriptions'
    SUFFIX = '_locations'

    def __init__(self, filePath, **kwargs):
        """Initialize local instance variables for parsing.

        Positional arguments:
            filePath -- str: path to the file represented by the Novel instance.
            
        The HTML parser works like a state machine. 
        The location ID must be saved between the transitions.         
        Extends the superclass constructor.
        """
        super().__init__(filePath)
        self._lcId = None

    def handle_starttag(self, tag, attrs):
        """Identify locations.
        
        Positional arguments:
            tag -- str: name of the tag converted to lower case.
            attrs -- list of (name, value) pairs containing the attributes found inside the tag’s <> brackets.
        
        Overrides the superclass method.
        """
        if tag == 'div':
            if attrs[0][0] == 'id':
                if attrs[0][1].startswith('LcID'):
                    self._lcId = re.search('[0-9]+', attrs[0][1]).group()
                    self.srtLocations.append(self._lcId)
                    self.locations[self._lcId] = WorldElement()

    def handle_endtag(self, tag):
        """Recognize the end of the location section and save data.
        
        Positional arguments:
            tag -- str: name of the tag converted to lower case.

        Overrides HTMLparser.handle_endtag() called by the HTML parser to handle the end tag of an element.
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
        
        Positional arguments:
            data -- str: text to be stored. 
        
        Overrides HTMLparser.handle_data() called by the parser when a comment is encountered.
        """
        if self._lcId is not None:
            self._lines.append(data.strip())


class HtmlItems(HtmlFile):
    """HTML item descriptions file representation.

    Import a item sheet with invisibly tagged descriptions.
    """
    DESCRIPTION = 'Item descriptions'
    SUFFIX = '_items'

    def __init__(self, filePath, **kwargs):
        """Initialize local instance variables for parsing.

        Positional arguments:
            filePath -- str: path to the file represented by the Novel instance.
            
        The HTML parser works like a state machine. 
        The item ID must be saved between the transitions.         
        Extends the superclass constructor.
        """
        super().__init__(filePath)
        self._itId = None

    def handle_starttag(self, tag, attrs):
        """Identify items.
        
        Positional arguments:
            tag -- str: name of the tag converted to lower case.
            attrs -- list of (name, value) pairs containing the attributes found inside the tag’s <> brackets.
        
        Overrides the superclass method.
        """
        if tag == 'div':
            if attrs[0][0] == 'id':
                if attrs[0][1].startswith('ItID'):
                    self._itId = re.search('[0-9]+', attrs[0][1]).group()
                    self.srtItems.append(self._itId)
                    self.items[self._itId] = WorldElement()

    def handle_endtag(self, tag):
        """Recognize the end of the item section and save data.
        
        Positional arguments:
            tag -- str: name of the tag converted to lower case.

        Overrides HTMLparser.handle_endtag() called by the HTML parser to handle the end tag of an element.
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

        Positional arguments:
            data -- str: text to be stored. 
        
        Overrides HTMLparser.handle_data() called by the parser when a comment is encountered.
        """
        if self._itId is not None:
            self._lines.append(data.strip())
import csv


class CsvFile(Novel):
    """csv file representation.

    Public methods:
        read() -- parse the file and get the instance variables.

    Convention:
    - Records are separated by line breaks.
    - Data fields are delimited by the _SEPARATOR character.
    """
    EXTENSION = '.csv'
    # overwrites Novel.EXTENSION
    _SEPARATOR = ','
    # delimits data fields within a record.
    _rowTitles = []

    def __init__(self, filePath, **kwargs):
        """Initialize instance variables.

        Positional arguments:
            filePath -- str: path to the file represented by the Novel instance.
            
        Optional arguments:
            kwargs -- keyword arguments to be used by subclasses.            
        
        Extends the superclass constructor.
        """
        super().__init__(filePath)
        self._rows = []
        
        
    def read(self):
        """Parse the file and get the instance variables.
        
        Parse the csv file located at filePath, fetching the rows.
        Check the number of fields in each row.
        Return a message beginning with the ERROR constant in case of error.
        Overrides the superclass method.
        """
        self._rows = []
        cellsPerRow = len(self._rowTitles)
        try:
            with open(self.filePath, newline='', encoding='utf-8') as f:
                reader = csv.reader(f, delimiter=self._SEPARATOR)
                for row in reader:
                    if len(row) != cellsPerRow:
                        return f'{ERROR}Wrong csv structure.'

                    self._rows.append(row)
        except(FileNotFoundError):
            return f'{ERROR}"{os.path.normpath(self.filePath)}" not found.'

        except:
            return f'{ERROR}Can not parse "{os.path.normpath(self.filePath)}".'

        return 'CSV data read in.'

    def _get_list(self, text):
        """Convert a string into a list.
        
        Positional arguments:
            text -- string containing comma-separated substrings.
        
        Split a sequence of comma separated strings into a list of strings.
        Remove leading and trailing spaces, if any.
        Return a list of strings.
        """
        elements = []
        tempList = text.split(',')
        for element in tempList:
            elements.append(element.strip())
        return elements


class CsvSceneList(CsvFile):
    """csv file representation of a yWriter project's scenes table. 
    
    Public methods:
        read() -- parse the file and get the instance variables.
    """
    DESCRIPTION = 'Scene list'
    SUFFIX = '_scenelist'
    _SCENE_RATINGS = ['2', '3', '4', '5', '6', '7', '8', '9', '10']
    # '1' is assigned N/A (empty table cell).
    _rowTitles = ['Scene link', 'Scene title', 'Scene description', 'Tags', 'Scene notes', 'A/R',
                 'Goal', 'Conflict', 'Outcome', 'Scene', 'Words total',
                 '$FieldTitle1', '$FieldTitle2', '$FieldTitle3', '$FieldTitle4',
                 'Word count', 'Letter count', 'Status', 'Characters', 'Locations', 'Items']

    def read(self):
        """Parse the file and get the instance variables.
        
        Parse the csv file located at filePath, fetching the Scene attributes contained.
        Return a message beginning with the ERROR constant in case of error.
        Extends the superclass method.
        """
        message = super().read()
        if message.startswith(ERROR):
            return message

        for cells in self._rows:
            i = 0
            if 'ScID:' in cells[i]:
                scId = re.search('ScID\:([0-9]+)', cells[0]).group(1)
                self.scenes[scId] = Scene()
                i += 1
                self.scenes[scId].title = cells[i]
                i += 1
                self.scenes[scId].desc = self._convert_to_yw(cells[i])
                i += 1
                self.scenes[scId].tags = self._get_list(cells[i])
                i += 1
                self.scenes[scId].sceneNotes = self._convert_to_yw(cells[i])
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
                # Can't write back character IDs, because self.characters is None.
                i += 1
                # Can't write back location IDs, because self.locations is None.
                i += 1
                # Can't write back item IDs, because self.items is None.
        return 'CSV data converted to novel structure.'


class CsvPlotList(CsvFile):
    """csv file representation of a yWriter project's scenes table. 
    
    Public methods:
        read() -- parse the file and get the instance variables.
    """
    DESCRIPTION = 'Plot list'
    SUFFIX = '_plotlist'
    _SCENE_RATINGS = ['2', '3', '4', '5', '6', '7', '8', '9', '10']
    # '1' is assigned N/A (empty table cell).
    _NOT_APPLICABLE = 'N/A'
    # Scene field column header for fields not being assigned to a storyline
    _rowTitles = ['ID', 'Plot section', 'Plot event', 'Scene title', 'Details', 'Scene', 'Words total',
                 '$FieldTitle1', '$FieldTitle2', '$FieldTitle3', '$FieldTitle4']

    def read(self):
        """Parse the file and get the instance variables.
        
        Parse the csv file located at filePath, fetching the Scene attributes contained.
        Return a message beginning with the ERROR constant in case of error.
        Extends the superclass method.
        """
        message = super().read()
        if message.startswith(ERROR):
            return message

        tableHeader = self._rows[0]
        for cells in self._rows:
            if 'ChID:' in cells[0]:
                chId = re.search('ChID\:([0-9]+)', cells[0]).group(1)
                self.chapters[chId] = Chapter()
                self.chapters[chId].title = cells[1]
                self.chapters[chId].desc = self._convert_to_yw(cells[4])
            if 'ScID:' in cells[0]:
                scId = re.search('ScID\:([0-9]+)', cells[0]).group(1)
                self.scenes[scId] = Scene()
                self.scenes[scId].tags = self._get_list(cells[2])
                self.scenes[scId].title = cells[3]
                self.scenes[scId].sceneNotes = self._convert_to_yw(cells[4])
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
        return 'CSV data converted to novel structure.'


class CsvCharList(CsvFile):
    """csv file representation of a yWriter project's characters table. 
    
    Public methods:
        read() -- parse the file and get the instance variables.
    """
    DESCRIPTION = 'Character list'
    SUFFIX = '_charlist'
    _rowTitles = ['ID', 'Name', 'Full name', 'Aka', 'Description', 'Bio', 'Goals', 'Importance', 'Tags', 'Notes']

    def read(self):
        """Parse the file and get the instance variables.
        
        Parse the csv file located at filePath, fetching the Character attributes contained.
        Return a message beginning with the ERROR constant in case of error.
        Extends the superclass method.
        """
        message = super().read()
        if message.startswith(ERROR):
            return message

        for cells in self._rows:
            if 'CrID:' in cells[0]:
                crId = re.search('CrID\:([0-9]+)', cells[0]).group(1)
                self.srtCharacters.append(crId)
                self.characters[crId] = Character()
                self.characters[crId].title = cells[1]
                self.characters[crId].fullName = cells[2]
                self.characters[crId].aka = cells[3]
                self.characters[crId].desc = self._convert_to_yw(cells[4])
                self.characters[crId].bio = cells[5]
                self.characters[crId].goals = cells[6]
                if Character.MAJOR_MARKER in cells[7]:
                    self.characters[crId].isMajor = True
                else:
                    self.characters[crId].isMajor = False
                self.characters[crId].tags = self._get_list(cells[8])
                self.characters[crId].notes = self._convert_to_yw(cells[9])
        return 'Character data read in.'


class CsvLocList(CsvFile):
    """csv file representation of a yWriter project's locations table. 
    
    Public methods:
        read() -- parse the file and get the instance variables.
    """
    DESCRIPTION = 'Location list'
    SUFFIX = '_loclist'
    _rowTitles = ['ID', 'Name', 'Description', 'Aka', 'Tags']

    def read(self):
        """Parse the file and get the instance variables.
        
        Parse the csv file located at filePath, fetching the location attributes contained.
        Return a message beginning with the ERROR constant in case of error.
        Extends the superclass method.
        """
        message = super().read()
        if message.startswith(ERROR):
            return message

        for cells in self._rows:
            if 'LcID:' in cells[0]:
                lcId = re.search('LcID\:([0-9]+)', cells[0]).group(1)
                self.srtLocations.append(lcId)
                self.locations[lcId] = WorldElement()
                self.locations[lcId].title = cells[1]
                self.locations[lcId].desc = self._convert_to_yw(cells[2])
                self.locations[lcId].aka = cells[3]
                self.locations[lcId].tags = self._get_list(cells[4])
        return 'Location data read in.'


class CsvItemList(CsvFile):
    """csv file representation of a yWriter project's items table. 
    
    Public methods:
        read() -- parse the file and get the instance variables.
    """
    DESCRIPTION = 'Item list'
    SUFFIX = '_itemlist'
    _rowTitles = ['ID', 'Name', 'Description', 'Aka', 'Tags']

    def read(self):
        """Parse the file and get the instance variables.
        
        Parse the csv file located at filePath, fetching the item attributes contained.
        Return a message beginning with the ERROR constant in case of error.
        Extends the superclass method.
        """
        message = super().read()
        if message.startswith(ERROR):
            return message

        for cells in self._rows:
            if 'ItID:' in cells[0]:
                itId = re.search('ItID\:([0-9]+)', cells[0]).group(1)
                self.srtItems.append(itId)
                self.items[itId] = WorldElement()
                self.items[itId].title = cells[1]
                self.items[itId].desc = self._convert_to_yw(cells[2])
                self.items[itId].aka = cells[3]
                self.items[itId].tags = self._get_list(cells[4])
        return 'Item data read in.'


class Yw7Converter(YwCnvFf):
    """A converter for universal import and export.

    Support yWriter 7 projects and most of the Novel subclasses 
    that can be read or written by OpenOffice/LibreOffice.

    Overrides the superclass constants EXPORT_SOURCE_CLASSES,
    EXPORT_TARGET_CLASSES, IMPORT_SOURCE_CLASSES, IMPORT_TARGET_CLASSES.

    Class constants:
        CREATE_SOURCE_CLASSES -- list of classes that - additional to HtmlImport
                        and HtmlOutline - can be exported to a new yWriter project.
    """
    EXPORT_SOURCE_CLASSES = [Yw7File]
    EXPORT_TARGET_CLASSES = [OdtExport,
                             OdtProof,
                             OdtManuscript,
                             OdtBriefSynopsis,
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
    IMPORT_TARGET_CLASSES = [Yw7File]
    CREATE_SOURCE_CLASSES = []

    def __init__(self):
        """Change the newProjectFactory strategy.
        
        Extends the superclass constructor.
        """
        super().__init__()
        self.newProjectFactory = NewProjectFactory(self.CREATE_SOURCE_CLASSES)


class YwCnvUno(Yw7Converter):
    """A converter for universal import and export.
    
    Public methods:
        export_from_yw(sourceFile, targetFile) -- Convert from yWriter project to other file format.

    Support yWriter 7 projects and most of the Novel subclasses 
    that can be read or written by OpenOffice/LibreOffice.
    - No message in case of success when converting from yWriter.
    """

    def export_from_yw(self, source, target):
        """Convert from yWriter project to other file format.

        Positional arguments:
            source -- YwFile subclass instance.
            target -- Any Novel subclass instance.

        Show only error messages.
        Overrides the superclass method.
        """
        message = self.convert(source, target)
        if message.startswith(ERROR):
            self.ui.set_info_how(message)
        else:
            self.newFile = target.filePath
from com.sun.star.awt.MessageBoxResults import OK, YES, NO, CANCEL
from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK, BUTTONS_OK_CANCEL, BUTTONS_YES_NO, BUTTONS_YES_NO_CANCEL, BUTTONS_RETRY_CANCEL, BUTTONS_ABORT_IGNORE_RETRY
from com.sun.star.awt.MessageBoxType import MESSAGEBOX, INFOBOX, WARNINGBOX, ERRORBOX, QUERYBOX


class UiUno(Ui):
    """UI subclass implementing a LibreOffice UNO facade."""

    def ask_yes_no(self, text):
        result = msgbox(text, buttons=BUTTONS_YES_NO, type_msg=WARNINGBOX)
        return result == YES

    def set_info_how(self, message):
        """How's the converter doing?"""
        self.infoHowText = message
        if message.startswith(ERROR):
            message = message.split(ERROR, maxsplit=1)[1].strip()
            msgbox(message, type_msg=ERRORBOX)
        else:
            msgbox(message, type_msg=INFOBOX)

INI_FILE = 'openyw.ini'


def open_yw7(suffix, newExt):

    # Set last opened yWriter project as default (if existing).
    scriptLocation = os.path.dirname(__file__)
    inifile = uno.fileUrlToSystemPath(f'{scriptLocation}/{INI_FILE}')
    defaultFile = None
    config = ConfigParser()
    try:
        config.read(inifile)
        ywLastOpen = config.get('FILES', 'yw_last_open')
        if os.path.isfile(ywLastOpen):
            defaultFile = uno.systemPathToFileUrl(ywLastOpen)
    except:
        pass

    # Ask for yWriter 7 project to open:
    ywFile = FilePicker(path=defaultFile)
    if ywFile is None:
        return

    sourcePath = uno.fileUrlToSystemPath(ywFile)
    __, ywExt = os.path.splitext(sourcePath)
    if not ywExt in ['.yw7']:
        msgbox('Please choose a yWriter 7 project.', 'Import from yWriter', type_msg=ERRORBOX)
        return

    # Store selected yWriter project as "last opened".
    newFile = ywFile.replace(ywExt, f'{suffix}{newExt}')
    dirName, fileName = os.path.split(newFile)
    thisDir = uno.fileUrlToSystemPath(f'{dirName}/')
    lockFile = f'{thisDir}.~lock.{fileName}#'
    if not config.has_section('FILES'):
        config.add_section('FILES')
    config.set('FILES', 'yw_last_open', uno.fileUrlToSystemPath(ywFile))
    with open(inifile, 'w') as f:
        config.write(f)

    # Check if import file is already open in LibreOffice:
    if os.path.isfile(lockFile):
        msgbox(f'Please close "{fileName}" first.', 'Import from yWriter', type_msg=ERRORBOX)
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
        desktop.loadComponentFromURL(newFile, "_blank", 0, ())


def import_yw():
    '''Import scenes from yWriter 7 to a Writer document
    without chapter and scene markers. 
    '''
    open_yw7('', '.odt')


def proof_yw():
    '''Import scenes from yWriter 7 to a Writer document
    with visible chapter and scene markers. 
    '''
    open_yw7(OdtProof.SUFFIX, OdtProof.EXTENSION)


def get_brf_synopsis():
    '''Import chapter and scene titles from yWriter 7 to a Writer document. 
    '''
    open_yw7(OdtBriefSynopsis.SUFFIX, OdtBriefSynopsis.EXTENSION)


def get_manuscript():
    '''Import scenes from yWriter 7 to a Writer document
    with invisible chapter and scene markers. 
    '''
    open_yw7(OdtManuscript.SUFFIX, OdtManuscript.EXTENSION)


def get_partdesc():
    '''Import pard descriptions from yWriter 7 to a Writer document
    with invisible chapter and scene markers. 
    '''
    open_yw7(OdtPartDesc.SUFFIX, OdtPartDesc.EXTENSION)


def get_chapterdesc():
    '''Import chapter descriptions from yWriter 7 to a Writer document
    with invisible chapter and scene markers. 
    '''
    open_yw7(OdtChapterDesc.SUFFIX, OdtChapterDesc.EXTENSION)


def get_scenedesc():
    '''Import scene descriptions from yWriter 7 to a Writer document
    with invisible chapter and scene markers. 
    '''
    open_yw7(OdtSceneDesc.SUFFIX, OdtSceneDesc.EXTENSION)


def get_chardesc():
    '''Import character descriptions from yWriter 7 to a Writer document.
    '''
    open_yw7(OdtCharacters.SUFFIX, OdtCharacters.EXTENSION)


def get_locdesc():
    '''Import location descriptions from yWriter 7 to a Writer document.
    '''
    open_yw7(OdtLocations.SUFFIX, OdtLocations.EXTENSION)


def get_itemdesc():
    '''Import item descriptions from yWriter 7 to a Writer document.
    '''
    open_yw7(OdtItems.SUFFIX, OdtItems.EXTENSION)


def get_xref():
    '''Generate cross references from yWriter 7 to a Writer document.
    '''
    open_yw7(OdtXref.SUFFIX, OdtXref.EXTENSION)


def get_scenelist():
    '''Import a scene list from yWriter 7 to a Calc document.
    '''
    open_yw7(OdsSceneList.SUFFIX, OdsSceneList.EXTENSION)


def get_plotlist():
    '''Import a plot list from yWriter 7 to a Calc document.
    '''
    open_yw7(OdsPlotList.SUFFIX, OdsPlotList.EXTENSION)


def get_charlist():
    '''Import a character list from yWriter 7 to a Calc document.
    '''
    open_yw7(OdsCharList.SUFFIX, OdsCharList.EXTENSION)


def get_loclist():
    '''Import a location list from yWriter 7 to a Calc document.
    '''
    open_yw7(OdsLocList.SUFFIX, OdsLocList.EXTENSION)


def get_itemlist():
    '''Import an item list from yWriter 7 to a Calc document.
    '''
    open_yw7(OdsItemList.SUFFIX, OdsItemList.EXTENSION)


def export_yw():
    '''Export the document to ayWriter 7 project.
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

        targetPath = uno.fileUrlToSystemPath(htmlPath)
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
        args1[2].Value = "44,34,76,1,,0,true,true,true"
        # args1(2).Value = "44,34,76,1,,0,true,true,true"
        dispatcher.executeDispatch(document, ".uno:SaveAs", "", 0, args1)
        # dispatcher.executeDispatch(document, ".uno:SaveAs", "", 0, args1())

        # Save document in OpenDocument format
        args1[0].Value = odsPath
        # args1(0).Value = odsPath
        args1[1].Value = 'calc8'
        # args1(1).Value = "calc8"
        dispatcher.executeDispatch(document, ".uno:SaveAs", "", 0, args1)
        # dispatcher.executeDispatch(document, ".uno:SaveAs", "", 0, args1())

        targetPath = uno.fileUrlToSystemPath(csvPath)
    else:
        msgbox(f'File type of "{os.path.normpath(documentPath)}" not supported.', type_msg=ERRORBOX)
        return

    converter = YwCnvUno()
    converter.ui = UiUno('Export to yWriter')
    kwargs = {'suffix': None}
    converter.run(targetPath, **kwargs)


def to_blank_lines():
    """Replace scene dividers with blank lines.

    Replace the three-lines "* * *" scene dividers with single blank lines. 
    Change the style of the scene-dividing paragraphs from  _Heading 4_  to  _Heading 5_.
    """
    pStyles = XSCRIPTCONTEXT.getDocument().StyleFamilies.getByName('ParagraphStyles')
    # pStyles = ThisComponent.StyleFamilies.getByName("ParagraphStyles")
    document = XSCRIPTCONTEXT.getDocument().CurrentController.Frame
    # document   = ThisComponent.CurrentController.Frame
    ctx = XSCRIPTCONTEXT.getComponentContext()
    smgr = ctx.getServiceManager()
    dispatcher = smgr.createInstanceWithContext(
        "com.sun.star.frame.DispatchHelper", ctx)
    # dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
    args1 = []
    for __ in range(19):
        args1.append(PropertyValue())
    # dim args1(18) as new com.sun.star.beans.PropertyValue
    args1[0].Name = "SearchItem.StyleFamily"
    args1[0].Value = 2
    args1[1].Name = "SearchItem.CellType"
    args1[1].Value = 0
    args1[2].Name = "SearchItem.RowDirection"
    args1[2].Value = True
    args1[3].Name = "SearchItem.AllTables"
    args1[3].Value = False
    args1[4].Name = "SearchItem.Backward"
    args1[4].Value = False
    args1[5].Name = "SearchItem.Pattern"
    args1[5].Value = True
    args1[6].Name = "SearchItem.Content"
    args1[6].Value = False
    args1[7].Name = "SearchItem.AsianOptions"
    args1[7].Value = False
    args1[8].Name = "SearchItem.AlgorithmType"
    args1[8].Value = 0
    args1[9].Name = "SearchItem.SearchFlags"
    args1[9].Value = 65536
    args1[10].Name = "SearchItem.SearchString"
    args1[10].Value = pStyles.getByName("Heading 4").DisplayName
    args1[11].Name = "SearchItem.ReplaceString"
    args1[11].Value = pStyles.getByName("Heading 5").DisplayName
    args1[12].Name = "SearchItem.Locale"
    args1[12].Value = 255
    args1[13].Name = "SearchItem.ChangedChars"
    args1[13].Value = 2
    args1[14].Name = "SearchItem.DeletedChars"
    args1[14].Value = 2
    args1[15].Name = "SearchItem.InsertedChars"
    args1[15].Value = 2
    args1[16].Name = "SearchItem.TransliterateFlags"
    args1[16].Value = 1280
    args1[17].Name = "SearchItem.Command"
    args1[17].Value = 3
    args1[18].Name = "Quiet"
    args1[18].Value = True
    dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args1)
    args2 = []
    for __ in range(19):
        args2.append(PropertyValue())
    # dim args2(18) as new com.sun.star.beans.PropertyValue
    args2[0].Name = "SearchItem.StyleFamily"
    args2[0].Value = 2
    args2[1].Name = "SearchItem.CellType"
    args2[1].Value = 0
    args2[2].Name = "SearchItem.RowDirection"
    args2[2].Value = True
    args2[3].Name = "SearchItem.AllTables"
    args2[3].Value = False
    args2[4].Name = "SearchItem.Backward"
    args2[4].Value = False
    args2[5].Name = "SearchItem.Pattern"
    args2[5].Value = True
    args2[6].Name = "SearchItem.Content"
    args2[6].Value = False
    args2[7].Name = "SearchItem.AsianOptions"
    args2[7].Value = False
    args2[8].Name = "SearchItem.AlgorithmType"
    args2[8].Value = 0
    args2[9].Name = "SearchItem.SearchFlags"
    args2[9].Value = 65536
    args2[10].Name = "SearchItem.SearchString"
    args2[10].Value = pStyles.getByName("Heading 5").DisplayName
    args2[11].Name = "SearchItem.ReplaceString"
    args2[11].Value = pStyles.getByName("Heading 5").DisplayName
    args2[12].Name = "SearchItem.Locale"
    args2[12].Value = 255
    args2[13].Name = "SearchItem.ChangedChars"
    args2[13].Value = 2
    args2[14].Name = "SearchItem.DeletedChars"
    args2[14].Value = 2
    args2[15].Name = "SearchItem.InsertedChars"
    args2[15].Value = 2
    args2[16].Name = "SearchItem.TransliterateFlags"
    args2[16].Value = 1280
    args2[17].Name = "SearchItem.Command"
    args2[17].Value = 1
    args2[18].Name = "Quiet"
    args2[18].Value = True
    dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args2)
    args3 = []
    for __ in range(19):
        args3.append(PropertyValue())
    # dim args3(18) as new com.sun.star.beans.PropertyValue
    args3[0].Name = "SearchItem.StyleFamily"
    args3[0].Value = 2
    args3[1].Name = "SearchItem.CellType"
    args3[1].Value = 0
    args3[2].Name = "SearchItem.RowDirection"
    args3[2].Value = True
    args3[3].Name = "SearchItem.AllTables"
    args3[3].Value = False
    args3[4].Name = "SearchItem.Backward"
    args3[4].Value = False
    args3[5].Name = "SearchItem.Pattern"
    args3[5].Value = False
    args3[6].Name = "SearchItem.Content"
    args3[6].Value = False
    args3[7].Name = "SearchItem.AsianOptions"
    args3[7].Value = False
    args3[8].Name = "SearchItem.AlgorithmType"
    args3[8].Value = 0
    args3[9].Name = "SearchItem.SearchFlags"
    args3[9].Value = 71680
    args3[10].Name = "SearchItem.SearchString"
    args3[10].Value = "* * *"
    args3[11].Name = "SearchItem.ReplaceString"
    args3[11].Value = ""
    args3[12].Name = "SearchItem.Locale"
    args3[12].Value = 255
    args3[13].Name = "SearchItem.ChangedChars"
    args3[13].Value = 2
    args3[14].Name = "SearchItem.DeletedChars"
    args3[14].Value = 2
    args3[15].Name = "SearchItem.InsertedChars"
    args3[15].Value = 2
    args3[16].Name = "SearchItem.TransliterateFlags"
    args3[16].Value = 1280
    args3[17].Name = "SearchItem.Command"
    args3[17].Value = 3
    args3[18].Name = "Quiet"
    args3[18].Value = True
    dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args3)


def indent_paragraphs():
    """Indent paragraphs that start with '> '.

    Select all paragraphs that start with '> ' 
    and change their paragraph style to _Text body indent_.
    """
    pStyles = XSCRIPTCONTEXT.getDocument().StyleFamilies.getByName('ParagraphStyles')
    # pStyles = ThisComponent.StyleFamilies.getByName("ParagraphStyles")
    document = XSCRIPTCONTEXT.getDocument().CurrentController.Frame
    # document   = ThisComponent.CurrentController.Frame
    ctx = XSCRIPTCONTEXT.getComponentContext()
    smgr = ctx.getServiceManager()
    dispatcher = smgr.createInstanceWithContext("com.sun.star.frame.DispatchHelper", ctx)
    # dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
    
    # Assign all paragraphs beginning with '> ' the 'Text body indent' style.
    args1 = []
    for __ in range(19):
        args1.append(PropertyValue())
    # dim args1(18) as new com.sun.star.beans.PropertyValue
    args1[0].Name = "SearchItem.StyleFamily"
    args1[0].Value = 2
    args1[1].Name = "SearchItem.CellType"
    args1[1].Value = 0
    args1[2].Name = "SearchItem.RowDirection"
    args1[2].Value = True
    args1[3].Name = "SearchItem.AllTables"
    args1[3].Value = False
    args1[4].Name = "SearchItem.Backward"
    args1[4].Value = False
    args1[5].Name = "SearchItem.Pattern"
    args1[5].Value = False
    args1[6].Name = "SearchItem.Content"
    args1[6].Value = False
    args1[7].Name = "SearchItem.AsianOptions"
    args1[7].Value = False
    args1[8].Name = "SearchItem.AlgorithmType"
    args1[8].Value = 1
    args1[9].Name = "SearchItem.SearchFlags"
    args1[9].Value = 65536
    args1[10].Name = "SearchItem.SearchString"
    args1[10].Value = "^> "
    args1[11].Name = "SearchItem.ReplaceString"
    args1[11].Value = ""
    args1[12].Name = "SearchItem.Locale"
    args1[12].Value = 255
    args1[13].Name = "SearchItem.ChangedChars"
    args1[13].Value = 2
    args1[14].Name = "SearchItem.DeletedChars"
    args1[14].Value = 2
    args1[15].Name = "SearchItem.InsertedChars"
    args1[15].Value = 2
    args1[16].Name = "SearchItem.TransliterateFlags"
    args1[16].Value = 1280
    args1[17].Name = "SearchItem.Command"
    args1[17].Value = 1
    args1[18].Name = "Quiet"
    args1[18].Value = True
    dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args1)
    args2 = []
    for __ in range(2):
        args2.append(PropertyValue())
    # dim args2(18) as new com.sun.star.beans.PropertyValue
    args2[0].Name = "Template"
    args2[0].Value = pStyles.getByName("Text body indent").DisplayName
    args2[1].Name = "Family"
    args2[1].Value = 2
    dispatcher.executeDispatch(document, ".uno:StyleApply", "", 0, args2)

    # Delete the markup.
    args1[17].Value = 3
    dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args1)
