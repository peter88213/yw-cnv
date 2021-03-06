"""Convert yWriter project to odt or csv and vice versa. 

Version 0.33.1

Copyright (c) 2020 Peter Triesberger
For further information see https://github.com/peter88213/yw-cnv
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import os

from configparser import ConfigParser
from urllib.parse import unquote
from urllib.parse import quote




from abc import abstractmethod


class FileFactory():
    """Abstract factory class that instantiates a source file object
    and a target file object for conversion.
    """

    @abstractmethod
    def get_file_objects(self, sourcePath, suffix=None):
        """Abstract method to be overwritten by subclasses.
        Return a tuple with three elements:
        * A message string starting with 'SUCCESS' or 'ERROR'
        * sourceFile: a Novel subclass instance
        * targetFile: a Novel subclass instance
        """




from abc import abstractmethod
from urllib.parse import quote


class Novel():
    """Abstract yWriter project file representation.

    This class represents a file containing a novel with additional 
    attributes and structural information (a full set or a subset
    of the information included in an yWriter project file).
    """

    DESCRIPTION = 'Novel'
    EXTENSION = None
    SUFFIX = None
    # To be extended by file format specific subclasses.

    def __init__(self, filePath):
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
        # key = chapter ID, value = Chapter object.
        # The order of the elements does not matter (the novel's
        # order of the chapters is defined by srtChapters)

        self.scenes = {}
        # dict
        # xml: <SCENES><SCENE><ID>
        # key = scene ID, value = Scene object.
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
        # key = location ID, value = Object.
        # The order of the elements does not matter.

        self.items = {}
        # dict
        # xml: <ITEMS>
        # key = item ID, value = Object.
        # The order of the elements does not matter.

        self.characters = {}
        # dict
        # xml: <CHARACTERS>
        # key = character ID, value = Character object.
        # The order of the elements does not matter.

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
        """Accept only filenames with the right extension. """

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

    @abstractmethod
    def read(self):
        """Parse the file and store selected properties.
        To be overwritten by file format specific subclasses.
        """

    @abstractmethod
    def merge(self, novel):
        """Copy required attributes of the novel object.
        To be overwritten by file format specific subclasses.
        """

    @abstractmethod
    def write(self):
        """Write selected properties to the file.
        To be overwritten by file format specific subclasses.
        """

    @abstractmethod
    def convert_to_yw(self, text):
        """Convert source format to yw7 markup.
        To be overwritten by file format specific subclasses.
        """

    @abstractmethod
    def convert_from_yw(self, text):
        """Convert yw7 markup to target format.
        To be overwritten by file format specific subclasses.
        """

    def file_exists(self):
        """Check whether the file specified by _filePath exists. """
        if os.path.isfile(self._filePath):
            return True

        else:
            return False


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

import re


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



class Object():
    """yWriter object representation.
    # xml: <LOCATIONS><LOCATION> or # xml: <ITEMS><ITEM>
    """

    def __init__(self):
        self.title = None
        # str
        # xml: <Title>

        self.desc = None
        # str
        # xml: <Desc>

        self.tags = None
        # list of str
        # xml: <Tags>

        self.aka = None
        # str
        # xml: <AKA>


class Character(Object):
    """yWriter character representation.
    # xml: <CHARACTERS><CHARACTER>
    """

    MAJOR_MARKER = 'Major'
    MINOR_MARKER = 'Minor'

    def __init__(self):
        Object.__init__(self)

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


class YwFile(Novel):
    """Abstract yWriter xml project file representation.
    To be overwritten by version-specific subclasses. 
    """

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

            self.locations[lcId] = Object()
            self.locations[lcId].title = loc.find('Title').text

            if loc.find('Desc') is not None:
                self.locations[lcId].desc = loc.find('Desc').text

            if loc.find('AKA') is not None:
                self.locations[lcId].aka = loc.find('AKA').text

            if loc.find('Tags') is not None:

                if loc.find('Tags').text is not None:
                    self.locations[lcId].tags = loc.find(
                        'Tags').text.split(';')

        # Read items from the xml element tree.

        for itm in root.iter('ITEM'):
            itId = itm.find('ID').text

            self.items[itId] = Object()
            self.items[itId].title = itm.find('Title').text

            if itm.find('Desc') is not None:
                self.items[itId].desc = itm.find('Desc').text

            if itm.find('AKA') is not None:
                self.items[itId].aka = itm.find('AKA').text

            if itm.find('Tags') is not None:

                if itm.find('Tags').text is not None:
                    self.items[itId].tags = itm.find(
                        'Tags').text.split(';')

        # Read characters from the xml element tree.

        for crt in root.iter('CHARACTER'):
            crId = crt.find('ID').text

            self.characters[crId] = Character()
            self.characters[crId].title = crt.find('Title').text

            if crt.find('Desc') is not None:
                self.characters[crId].desc = crt.find('Desc').text

            if crt.find('AKA') is not None:
                self.characters[crId].aka = crt.find('AKA').text

            if crt.find('Tags') is not None:

                if crt.find('Tags').text is not None:
                    self.characters[crId].tags = crt.find(
                        'Tags').text.split(';')

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
                    self.scenes[scId].tags = scn.find(
                        'Tags').text.split(';')

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

        return 'SUCCESS: ' + str(len(self.scenes)) + ' Scenes read from "' + self._filePath + '".'

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

from abc import abstractmethod


class YwTreeBuilder():
    """Build yWriter project xml tree."""

    @abstractmethod
    def build_element_tree(self, ywProject):
        """Modify the yWriter project attributes of an existing xml element tree.
        Return a message beginning with SUCCESS or ERROR.
        To be overwritten by file format specific subclasses.
        """
        root = ywProject._tree.getroot()

        # Write locations to the xml element tree.

        for loc in root.iter('LOCATION'):
            lcId = loc.find('ID').text

            if lcId in ywProject.locations:

                if ywProject.locations[lcId].title is not None:
                    loc.find('Title').text = ywProject.locations[lcId].title

                if ywProject.locations[lcId].desc is not None:

                    if loc.find('Desc') is None:
                        ET.SubElement(
                            loc, 'Desc').text = ywProject.locations[lcId].desc

                    else:
                        loc.find('Desc').text = ywProject.locations[lcId].desc

                if ywProject.locations[lcId].aka is not None:

                    if loc.find('AKA') is None:
                        ET.SubElement(
                            loc, 'AKA').text = ywProject.locations[lcId].aka

                    else:
                        loc.find('AKA').text = ywProject.locations[lcId].aka

                if ywProject.locations[lcId].tags is not None:

                    if loc.find('Tags') is None:
                        ET.SubElement(loc, 'Tags').text = ';'.join(
                            ywProject.locations[lcId].tags)

                    else:
                        loc.find('Tags').text = ';'.join(
                            ywProject.locations[lcId].tags)

        # Write items to the xml element tree.

        for itm in root.iter('ITEM'):
            itId = itm.find('ID').text

            if itId in ywProject.items:

                if ywProject.items[itId].title is not None:
                    itm.find('Title').text = ywProject.items[itId].title

                if ywProject.items[itId].desc is not None:

                    if itm.find('Desc') is None:
                        ET.SubElement(
                            itm, 'Desc').text = ywProject.items[itId].desc

                    else:
                        itm.find('Desc').text = ywProject.items[itId].desc

                if ywProject.items[itId].aka is not None:

                    if itm.find('AKA') is None:
                        ET.SubElement(
                            itm, 'AKA').text = ywProject.items[itId].aka

                    else:
                        itm.find('AKA').text = ywProject.items[itId].aka

                if ywProject.items[itId].tags is not None:

                    if itm.find('Tags') is None:
                        ET.SubElement(itm, 'Tags').text = ';'.join(
                            ywProject.items[itId].tags)

                    else:
                        itm.find('Tags').text = ';'.join(
                            ywProject.items[itId].tags)

        # Write characters to the xml element tree.

        for crt in root.iter('CHARACTER'):
            crId = crt.find('ID').text

            if crId in ywProject.characters:

                if ywProject.characters[crId].title is not None:
                    crt.find('Title').text = ywProject.characters[crId].title

                if ywProject.characters[crId].desc is not None:

                    if crt.find('Desc') is None:
                        ET.SubElement(
                            crt, 'Desc').text = ywProject.characters[crId].desc

                    else:
                        crt.find('Desc').text = ywProject.characters[crId].desc

                if ywProject.characters[crId].aka is not None:

                    if crt.find('AKA') is None:
                        ET.SubElement(
                            crt, 'AKA').text = ywProject.characters[crId].aka

                    else:
                        crt.find('AKA').text = ywProject.characters[crId].aka

                if ywProject.characters[crId].tags is not None:

                    if crt.find('Tags') is None:
                        ET.SubElement(crt, 'Tags').text = ';'.join(
                            ywProject.characters[crId].tags)

                    else:
                        crt.find('Tags').text = ';'.join(
                            ywProject.characters[crId].tags)

                if ywProject.characters[crId].notes is not None:

                    if crt.find('Notes') is None:
                        ET.SubElement(
                            crt, 'Notes').text = ywProject.characters[crId].notes

                    else:
                        crt.find(
                            'Notes').text = ywProject.characters[crId].notes

                if ywProject.characters[crId].bio is not None:

                    if crt.find('Bio') is None:
                        ET.SubElement(
                            crt, 'Bio').text = ywProject.characters[crId].bio

                    else:
                        crt.find('Bio').text = ywProject.characters[crId].bio

                if ywProject.characters[crId].goals is not None:

                    if crt.find('Goals') is None:
                        ET.SubElement(
                            crt, 'Goals').text = ywProject.characters[crId].goals

                    else:
                        crt.find(
                            'Goals').text = ywProject.characters[crId].goals

                if ywProject.characters[crId].fullName is not None:

                    if crt.find('FullName') is None:
                        ET.SubElement(
                            crt, 'FullName').text = ywProject.characters[crId].fullName

                    else:
                        crt.find(
                            'FullName').text = ywProject.characters[crId].fullName

                majorMarker = crt.find('Major')

                if majorMarker is not None:

                    if not ywProject.characters[crId].isMajor:
                        crt.remove(majorMarker)

                else:
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


from abc import abstractmethod


class YwTreeReader():
    """Read yWriter xml project file."""

    @abstractmethod
    def read_element_tree(self, ywFile):
        """Parse the yWriter xml file located at filePath, fetching the Novel attributes.
        Return a message beginning with SUCCESS or ERROR.
        To be overwritten by file format specific subclasses.
        """


class AnsiTreeReader(YwTreeReader):
    """Read yWriter xml project file."""

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
            return 'ERROR: Can not process "' + os.path.normpath(ywFile._filePath) + '".'

        return 'SUCCESS: XML element tree read in.'



class YwProjectMerger():
    """Merge two yWriter projects.
    """

    def merge_projects(self, target, source):
        """Overwrite existing target attributes with source attributes.
        Create target attributes, if not existing, but return ERROR.
        Return a message beginning with SUCCESS or ERROR.
        """

        mismatchCount = 0

        # Merge locations.

        for lcId in source.locations:

            if not lcId in target.locations:
                target.locations[lcId] = Object()
                mismatchCount += 1

            if source.locations[lcId].title:
                # avoids deleting the title, if it is empty by accident
                target.locations[lcId].title = source.locations[lcId].title

            if source.locations[lcId].desc is not None:
                target.locations[lcId].desc = source.locations[lcId].desc

            if source.locations[lcId].aka is not None:
                target.locations[lcId].aka = source.locations[lcId].aka

            if source.locations[lcId].tags is not None:
                target.locations[lcId].tags = source.locations[lcId].tags

        # Merge items.

        for itId in source.items:

            if not itId in target.items:
                target.items[itId] = Object()
                mismatchCount += 1

            if source.items[itId].title:
                # avoids deleting the title, if it is empty by accident
                target.items[itId].title = source.items[itId].title

            if source.items[itId].desc is not None:
                target.items[itId].desc = source.items[itId].desc

            if source.items[itId].aka is not None:
                target.items[itId].aka = source.items[itId].aka

            if source.items[itId].tags is not None:
                target.items[itId].tags = source.items[itId].tags

        # Merge characters.

        for crId in source.characters:

            if not crId in target.characters:
                target.characters[crId] = Character()
                mismatchCount += 1

            if source.characters[crId].title:
                # avoids deleting the title, if it is empty by accident
                target.characters[crId].title = source.characters[crId].title

            if source.characters[crId].desc is not None:
                target.characters[crId].desc = source.characters[crId].desc

            if source.characters[crId].aka is not None:
                target.characters[crId].aka = source.characters[crId].aka

            if source.characters[crId].tags is not None:
                target.characters[crId].tags = source.characters[crId].tags

            if source.characters[crId].notes is not None:
                target.characters[crId].notes = source.characters[crId].notes

            if source.characters[crId].bio is not None:
                target.characters[crId].bio = source.characters[crId].bio

            if source.characters[crId].goals is not None:
                target.characters[crId].goals = source.characters[crId].goals

            if source.characters[crId].fullName is not None:
                target.characters[crId].fullName = source.characters[crId].fullName

            if source.characters[crId].isMajor is not None:
                target.characters[crId].isMajor = source.characters[crId].isMajor

        # Merge scenes.

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

            if source.scenes[scId].rtfFile is not None:
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


from abc import abstractmethod


class YwTreeWriter():
    """Write yWriter 7 xml project file."""

    @abstractmethod
    def write_element_tree(self, ywProject):
        """Write back the xml element tree to a yWriter xml file located at filePath.
        Return a message beginning with SUCCESS or ERROR.
        To be overwritten by file format specific subclasses.
        """


class AnsiTreeWriter(YwTreeWriter):
    """Write ANSI encoded yWriter project file."""

    def write_element_tree(self, ywProject):
        """Write back the xml element tree to a yWriter xml file located at filePath.
        Return a message beginning with SUCCESS or ERROR.
        """

        try:
            ywProject._tree.write(
                ywProject._filePath, xml_declaration=False, encoding='iso-8859-1')

        except(PermissionError):
            return 'ERROR: "' + os.path.normpath(ywProject._filePath) + '" is write protected.'

        return 'SUCCESS'


from html import unescape
from abc import abstractmethod


class YwPostprocessor():

    @abstractmethod
    def postprocess_xml_file(self, ywFile):
        '''Postprocess the xml file created by ElementTree:
        Put a header on top, insert the missing CDATA tags,
        and replace xml entities by plain text.
        Return a message beginning with SUCCESS or ERROR.
        To be overwritten by file format specific subclasses.
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

    def __init__(self, filePath):
        YwFile.__init__(self, filePath)
        self.ywTreeReader = AnsiTreeReader()
        self.ywProjectMerger = YwProjectMerger()
        self.ywTreeBuilder = Yw5TreeBuilder()
        self.ywTreeWriter = AnsiTreeWriter()
        self.ywPostprocessor = AnsiPostprocessor()




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



class Utf8TreeReader(YwTreeReader):
    """Read yWriter xml project file."""

    def read_element_tree(self, ywFile):
        """Parse the yWriter xml file located at filePath, fetching the Novel attributes.
        Return a message beginning with SUCCESS or ERROR.
        """

        try:
            ywFile._tree = ET.parse(ywFile._filePath)

        except:
            return 'ERROR: Can not process "' + os.path.normpath(ywFile._filePath) + '".'

        return 'SUCCESS: XML element tree read in.'



class Utf8TreeWriter(YwTreeWriter):
    """Write utf-8 encoded yWriter project file."""

    def write_element_tree(self, ywProject):
        """Write back the xml element tree to a yWriter xml file located at filePath.
        Return a message beginning with SUCCESS or ERROR.
        """

        try:
            ywProject._tree.write(
                ywProject._filePath, xml_declaration=False, encoding='utf-8')

        except(PermissionError):
            return 'ERROR: "' + os.path.normpath(ywProject._filePath) + '" is write protected.'

        return 'SUCCESS'



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


class Yw6File(YwFile):
    """yWriter 6 project file representation."""

    DESCRIPTION = 'yWriter 6 project'
    EXTENSION = '.yw6'

    def __init__(self, filePath):
        YwFile.__init__(self, filePath)
        self.ywTreeReader = Utf8TreeReader()
        self.ywProjectMerger = YwProjectMerger()
        self.ywTreeBuilder = Yw6TreeBuilder()
        self.ywTreeWriter = Utf8TreeWriter()
        self.ywPostprocessor = Utf8Postprocessor()




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


class Yw7File(YwFile):
    """yWriter 7 project file representation.
    """

    DESCRIPTION = 'yWriter 7 project'
    EXTENSION = '.yw7'

    def __init__(self, filePath):
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

        for lcId in ywProject.locations:
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

        for itId in ywProject.items:
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

        for crId in ywProject.characters:
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

from shutil import rmtree



class Yw5TreeCreator(Yw5TreeBuilder):
    """Create a new yWriter 7 project xml tree."""

    def build_element_tree(self, ywProject):
        """Put the yWriter project attributes to a new xml element tree.
        Return a message beginning with SUCCESS or ERROR.
        """

        # Copy yw7 file.

        yw7File = os.path.splitext(ywProject.filePath)[0] + '.yw7'

        try:
            with open(yw7File, 'rb') as f:
                project = f.read()

            with open(ywProject.filePath, 'wb') as f:
                f.write(project)

        except:
            return 'ERROR: Can not copy "' + yw7File + ' to ' + ywProject.filePath + '".'

        # Create RTF5 directory.

        rtfDir = os.path.dirname(ywProject.filePath)

        if rtfDir == '':
            rtfDir = './RTF5'

        else:
            rtfDir += '/RTF5'

        try:
            rmtree(rtfDir)

        except:
            pass

        try:
            os.mkdir(rtfDir)

        except:
            return 'ERROR: cannot create scene dir "' + rtfDir + '".'

        # Create RTF file entries.

        sceneCount = 0

        for scId in ywProject.scenes:
            sceneCount += 1
            ywProject.scenes[scId].rtfFile = 'RTF_' + \
                str(sceneCount).zfill(5) + '.rtf'

        # Modify xml tree.

        try:
            ywProject._tree = ET.parse(ywProject._filePath)

        except:
            return 'ERROR: Can not read xml file "' + ywProject._filePath + '".'

        return Yw5TreeBuilder.build_element_tree(self, ywProject)



class YwProjectCreator(YwProjectMerger):
    """Create a new project.
    """

    def merge_projects(self, target, source):
        """Create target attributes with source attributes.
        Return a message beginning with SUCCESS or ERROR.
        """
        YwProjectMerger.merge_projects(self, target, source)
        return 'SUCCESS'

import locale
from shutil import rmtree
from datetime import datetime
from string import Template


class OdtBuilder():
    """Build an ODT zipfile.
    """

    TEMPDIR = 'temp_odt'

    ODT_COMPONENTS = ['manifest.rdf', 'META-INF', 'content.xml', 'meta.xml', 'mimetype',
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

    def tear_down(self):
        """Delete the temporary directory 
        containing the unpacked ODT directory structure.
        """
        try:
            rmtree(self.TEMPDIR)
        except:
            pass

    def set_up(self):
        """Create a temporary directory containing the internal 
        structure of an ODT file except 'content.xml'.
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

        # Generate manifest.rdf

        try:
            with open(self.TEMPDIR + '/manifest.rdf', 'w', encoding='utf-8') as f:
                f.write(self._MANIFEST_RDF)
        except:
            return 'ERROR: Cannot write "manifest.rdf"'

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

        return 'SUCCESS: ODT structure generated.'
import zipfile

from string import Template



class FileExport(Novel):
    """Abstract yWriter project file exporter representation.
    To be overwritten by subclasses providing file type specific 
    markup converters and templates.
    """

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

    def convert_from_yw(self, text):
        """Convert yw7 markup to target format.
        To be overwritten by file format specific subclasses.
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

        if novel.characters is not None:
            self.characters = novel.characters

        if novel.locations is not None:
            self.locations = novel.locations

        if novel.items is not None:
            self.items = novel.items

        return 'SUCCESS'

    def get_projectTemplateMapping(self):
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

        for key in projectTemplateMapping:
            if projectTemplateMapping[key] is None:
                projectTemplateMapping[key] = ''

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

        for key in chapterMapping:
            if chapterMapping[key] is None:
                chapterMapping[key] = ''

        return chapterMapping

    def get_sceneMapping(self, scId, sceneNumber, wordsTotal, lettersTotal):
        """Return a mapping dictionary for a scene section. 
        """

        if self.scenes[scId].tags is not None:
            tags = ', '.join(self.scenes[scId].tags)

        else:
            tags = ''

        try:
            # Note: Due to a bug, yWriter scenes might hold invalid
            # viepoint characters
            sChList = []

            for chId in self.scenes[scId].characters:
                sChList.append(self.characters[chId].title)

            sceneChars = ', '.join(sChList)
            viewpointChar = sChList[0]

        except:
            sceneChars = ''
            viewpointChar = ''

        if self.scenes[scId].locations is not None:
            sLcList = []

            for lcId in self.scenes[scId].locations:
                sLcList.append(self.locations[lcId].title)

            sceneLocs = ', '.join(sLcList)

        else:
            sceneLocs = ''

        if self.scenes[scId].items is not None:
            sItList = []

            for itId in self.scenes[scId].items:
                sItList.append(self.items[itId].title)

            sceneItems = ', '.join(sItList)

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

        for key in sceneMapping:
            if sceneMapping[key] is None:
                sceneMapping[key] = ''

        return sceneMapping

    def get_characterMapping(self, crId):
        """Return a mapping dictionary for a character section. 
        """

        if self.characters[crId].tags is not None:
            tags = ', '.join(self.characters[crId].tags)

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
        )

        for key in characterMapping:
            if characterMapping[key] is None:
                characterMapping[key] = ''

        return characterMapping

    def get_locationMapping(self, lcId):
        """Return a mapping dictionary for a location section. 
        """

        if self.locations[lcId].tags is not None:
            tags = ', '.join(self.locations[lcId].tags)

        else:
            tags = ''

        locationMapping = dict(
            ID=lcId,
            Title=self.locations[lcId].title,
            Desc=self.convert_from_yw(self.locations[lcId].desc),
            Tags=tags,
            AKA=FileExport.convert_from_yw(self, self.locations[lcId].aka),
        )

        for key in locationMapping:
            if locationMapping[key] is None:
                locationMapping[key] = ''

        return locationMapping

    def get_itemMapping(self, itId):
        """Return a mapping dictionary for an item section. 
        """

        if self.items[itId].tags is not None:
            tags = ', '.join(self.items[itId].tags)

        else:
            tags = ''

        itemMapping = dict(
            ID=itId,
            Title=self.items[itId].title,
            Desc=self.convert_from_yw(self.items[itId].desc),
            Tags=tags,
            AKA=FileExport.convert_from_yw(self, self.items[itId].aka),
        )

        for key in itemMapping:
            if itemMapping[key] is None:
                itemMapping[key] = ''

        return itemMapping

    def write(self):
        """Create a template-based output file. 
        Return a message string starting with 'SUCCESS' or 'ERROR'.
        """
        lines = []
        wordsTotal = 0
        lettersTotal = 0
        chapterNumber = 0
        sceneNumber = 0

        template = Template(self.fileHeader)
        lines.append(template.safe_substitute(
            self.get_projectTemplateMapping()))

        for chId in self.srtChapters:

            # The order counts; be aware that "Todo" and "Notes" chapters are
            # always unused.

            # Has the chapter only scenes not to be exported?

            sceneCount = 0
            notExportCount = 0
            doNotExportChapter = False

            for scId in self.chapters[chId].srtScenes:
                sceneCount += 1

                if self.scenes[scId].doNotExport:
                    notExportCount += 1

            if sceneCount > 0 and notExportCount == sceneCount:
                doNotExportChapter = True

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

            elif doNotExportChapter:

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

                elif self.scenes[scId].doNotExport or doNotExportChapter:

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

            if self.chapters[chId].chType == 2 and self.todoChapterEndTemplate != '':
                lines.append(self.todoChapterEndTemplate)

            elif self.chapters[chId].chType == 1 or self.chapters[chId].oldType == 1:

                if self.notesChapterEndTemplate != '':
                    lines.append(self.notesChapterEndTemplate)

            elif self.chapters[chId].isUnused and self.unusedChapterEndTemplate != '':
                lines.append(self.unusedChapterEndTemplate)

            elif doNotExportChapter and self.notExportedChapterEndTemplate != '':
                lines.append(self.notExportedChapterEndTemplate)

            elif self.chapterEndTemplate != '':
                lines.append(self.chapterEndTemplate)

        for crId in self.characters:
            template = Template(self.characterTemplate)
            lines.append(template.safe_substitute(
                self.get_characterMapping(crId)))

        for lcId in self.locations:
            template = Template(self.locationTemplate)
            lines.append(template.safe_substitute(
                self.get_locationMapping(lcId)))

        for itId in self.items:
            template = Template(self.itemTemplate)
            lines.append(template.safe_substitute(self.get_itemMapping(itId)))

        lines.append(self.fileFooter)
        text = ''.join(lines)

        try:
            with open(self.filePath, 'w', encoding='utf-8') as f:
                f.write(text)

        except:
            return 'ERROR: Cannot write "' + os.path.normpath(self.filePath) + '".'

        return 'SUCCESS: "' + os.path.normpath(self.filePath) + '" written.'


class OdtFile(FileExport, OdtBuilder):
    """OpenDocument xml project file representation.
    """

    EXTENSION = '.odt'
    # overwrites Novel.EXTENSION

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

        except AttributeError:
            text = ''

        return text

    def write(self):
        """Generate an odt file from a template.
        Return a message beginning with SUCCESS or ERROR.
        """

        # Create a temporary directory containing the internal
        # structure of an ODT file except "content.xml".

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
        # into the ODT file.

        workdir = os.getcwd()

        try:
            with zipfile.ZipFile(self.filePath, 'w') as odtTarget:
                os.chdir(self.TEMPDIR)

                for file in self.ODT_COMPONENTS:
                    odtTarget.write(file)
        except:
            os.chdir(workdir)
            return 'ERROR: Cannot generate "' + os.path.normpath(self.filePath) + '".'

        # Remove temporary data.

        os.chdir(workdir)
        self.tear_down()
        return 'SUCCESS: "' + os.path.normpath(self.filePath) + '" written.'


class OdtProof(OdtFile):
    """OpenDocument xml proof reading file representation.
    """

    DESCRIPTION = 'Tagged manuscript for proofing'
    SUFFIX = '_proof'

    fileHeader = OdtBuilder.CONTENT_XML_HEADER + '''<text:p text:style-name="Title">$Title</text:p>
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

    fileFooter = OdtBuilder.CONTENT_XML_FOOTER


class OdtManuscript(OdtFile):
    """OpenDocument xml manuscript file representation.
    """

    DESCRIPTION = 'Editable manuscript'
    SUFFIX = '_manuscript'

    fileHeader = OdtBuilder.CONTENT_XML_HEADER + '''<text:p text:style-name="Title">$Title</text:p>
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

    fileFooter = OdtBuilder.CONTENT_XML_FOOTER

    def get_chapterMapping(self, chId, chapterNumber):
        """Return a mapping dictionary for a chapter section. 
        """
        chapterMapping = OdtFile.get_chapterMapping(self, chId, chapterNumber)

        if self.chapters[chId].suppressChapterTitle:
            chapterMapping['Title'] = ''

        return chapterMapping


class OdtSceneDesc(OdtFile):
    """OpenDocument xml scene summaries file representation.
    """

    DESCRIPTION = 'Scene descriptions'
    SUFFIX = '_scenes'

    fileHeader = OdtBuilder.CONTENT_XML_HEADER + '''<text:p text:style-name="Title">$Title</text:p>
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

    fileFooter = OdtBuilder.CONTENT_XML_FOOTER


class OdtChapterDesc(OdtFile):
    """OpenDocument xml chapter summaries file representation.
    """

    DESCRIPTION = 'Chapter descriptions'
    SUFFIX = '_chapters'

    fileHeader = OdtBuilder.CONTENT_XML_HEADER + '''<text:p text:style-name="Title">$Title</text:p>
<text:p text:style-name="Subtitle">$AuthorName</text:p>
'''

    partTemplate = '''<text:h text:style-name="Heading_20_1" text:outline-level="1"><text:a xlink:href="../${ProjectName}_parts.odt#ChID:$ID%7Cregion">$Title</text:a></text:h>
'''

    chapterTemplate = '''<text:h text:style-name="Heading_20_2" text:outline-level="2"><text:a xlink:href="../${ProjectName}_manuscript.odt#ChID:$ID%7Cregion">$Title</text:a></text:h>
<text:section text:style-name="Sect1" text:name="ChID:$ID">
<text:p text:style-name="Text_20_body">$Desc</text:p>
</text:section>
'''

    fileFooter = OdtBuilder.CONTENT_XML_FOOTER


class OdtPartDesc(OdtFile):
    """OpenDocument xml part summaries file representation.
    """

    DESCRIPTION = 'Part descriptions'
    SUFFIX = '_parts'

    fileHeader = OdtBuilder.CONTENT_XML_HEADER + '''<text:p text:style-name="Title">$Title</text:p>
<text:p text:style-name="Subtitle">$AuthorName</text:p>
'''

    partTemplate = '''<text:h text:style-name="Heading_20_1" text:outline-level="1"><text:a xlink:href="../${ProjectName}_manuscript.odt#ChID:$ID%7Cregion">$Title</text:a></text:h>
<text:section text:style-name="Sect1" text:name="ChID:$ID">
<text:p text:style-name="Text_20_body">$Desc</text:p>
</text:section>
'''

    fileFooter = OdtBuilder.CONTENT_XML_FOOTER


class OdtExport(OdtFile):
    """OpenDocument xml project file representation.
    """

    SUFFIX = ''

    fileHeader = OdtBuilder.CONTENT_XML_HEADER + '''<text:p text:style-name="Title">$Title</text:p>
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

    fileFooter = OdtBuilder.CONTENT_XML_FOOTER

    def get_chapterMapping(self, chId, chapterNumber):
        """Return a mapping dictionary for a chapter section. 
        """
        chapterMapping = OdtFile.get_chapterMapping(self, chId, chapterNumber)

        if self.chapters[chId].suppressChapterTitle:
            chapterMapping['Title'] = ''

        return chapterMapping


class OdtCharacters(OdtFile):
    """OpenDocument xml character descriptions file representation.
    """

    DESCRIPTION = 'Character descriptions'
    SUFFIX = '_characters'

    fileHeader = OdtBuilder.CONTENT_XML_HEADER + '''<text:p text:style-name="Title">$Title</text:p>
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

    fileFooter = OdtBuilder.CONTENT_XML_FOOTER

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
    """OpenDocument xml item descriptions file representation.
    """

    DESCRIPTION = 'Item descriptions'
    SUFFIX = '_items'

    fileHeader = OdtBuilder.CONTENT_XML_HEADER + '''<text:p text:style-name="Title">$Title</text:p>
<text:p text:style-name="Subtitle">$AuthorName</text:p>
'''

    itemTemplate = '''<text:h text:style-name="Heading_20_2" text:outline-level="2">$Title$AKA</text:h>
<text:section text:style-name="Sect1" text:name="ItID:$ID">
<text:p text:style-name="Text_20_body">$Desc</text:p>
</text:section>
'''

    fileFooter = OdtBuilder.CONTENT_XML_FOOTER

    def get_itemMapping(self, itId):
        """Return a mapping dictionary for an item section. 
        """
        itemMapping = OdtFile.get_itemMapping(self, itId)

        if self.items[itId].aka:
            itemMapping['AKA'] = ' ("' + self.items[itId].aka + '")'

        return itemMapping


class OdtLocations(OdtFile):
    """OpenDocument xml location descriptions file representation.
    """

    DESCRIPTION = 'Location descriptions'
    SUFFIX = '_locations'

    fileHeader = OdtBuilder.CONTENT_XML_HEADER + '''<text:p text:style-name="Title">$Title</text:p>
<text:p text:style-name="Subtitle">$AuthorName</text:p>
'''

    locationTemplate = '''<text:h text:style-name="Heading_20_2" text:outline-level="2">$Title$AKA</text:h>
<text:section text:style-name="Sect1" text:name="LcID:$ID">
<text:p text:style-name="Text_20_body">$Desc</text:p>
</text:section>
'''

    fileFooter = OdtBuilder.CONTENT_XML_FOOTER

    def get_locationMapping(self, lcId):
        """Return a mapping dictionary for a location section. 
        """
        locationMapping = OdtFile.get_locationMapping(self, lcId)

        if self.locations[lcId].aka:
            locationMapping['AKA'] = ' ("' + self.locations[lcId].aka + '")'

        return locationMapping



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
    """HTML file representation of an yWriter project's part.
    """

    EXTENSION = '.html'

    def __init__(self, filePath):
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
        """Strip yWriter 6/7 raw markup. Return a plain text string."""

        text = self.convert_to_yw(text)
        text = text.replace('[i]', '')
        text = text.replace('[/i]', '')
        text = text.replace('[b]', '')
        text = text.replace('[/b]', '')
        return text

    def postprocess(self):
        """Process the plain text after parsing.
        """

    def handle_starttag(self, tag, attrs):
        """Identify scenes and chapters.
        Overwrites HTMLparser.handle_starttag()
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
        """Read scene content from a html file 
        with chapter and scene sections.
        Return a message beginning with SUCCESS or ERROR. 
        """
        result = read_html_file(self._filePath)

        if result[0].startswith('ERROR'):
            return (result[0])

        text = self.preprocess(result[1])
        self.feed(text)
        self.postprocess()

        return 'SUCCESS'



class HtmlProof(HtmlFile):
    """HTML file representation of an yWriter project's OfficeFile part.

    Represents a html file with visible chapter and scene tags 
    to be read and written by Open/LibreOffice Writer.
    """

    DESCRIPTION = 'Tagged manuscript for proofing'
    SUFFIX = '_proof'

    def __init__(self, filePath):
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
    """HTML file representation of an yWriter project's manuscript part.

    Represents a html file with chapter and scene sections 
    containing scene contents to be read and written by 
    OpenOffice/LibreOffice Writer.
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
    """HTML file representation of an yWriter project's scene summaries.
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
    """HTML file representation of an yWriter project's chapters summaries."""

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
    """HTML file representation of an yWriter project's parts summaries."""

    DESCRIPTION = 'Part descriptions'
    SUFFIX = '_parts'




class HtmlCharacters(HtmlFile):
    """HTML file representation of an yWriter project's character descriptions."""

    DESCRIPTION = 'Character descriptions'
    SUFFIX = '_characters'

    def __init__(self, filePath):
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
    """HTML file representation of an yWriter project's location descriptions."""

    DESCRIPTION = 'Location descriptions'
    SUFFIX = '_locations'

    def __init__(self, filePath):
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
                    self.locations[self._lcId] = Object()

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
    """HTML file representation of an yWriter project's item descriptions."""

    DESCRIPTION = 'Item descriptions'
    SUFFIX = '_items'

    def __init__(self, filePath):
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
                    self.items[self._itId] = Object()

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



class HtmlImport(HtmlFile):
    """HTML file representation of a work in progress to be 
    converted to a new yWriter project yWriter project.
    """

    DESCRIPTION = 'Work in progress'
    SUFFIX = ''

    _SCENE_DIVIDER = '* * *'
    _LOW_WORDCOUNT = 10

    def __init__(self, filePath):
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
                    self.scenes[self._scId].status = 1

                else:
                    self.scenes[self._scId].status = 2

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
            self._lines.append(data.rstrip().lstrip())



class HtmlOutline(HtmlFile):
    """HTML file representation of an yWriter project's OfficeFile part.

    Represents a html file without chapter and scene tags 
    to be written by Open/LibreOffice Writer.
    """

    DESCRIPTION = 'Novel outline'
    SUFFIX = ''

    def __init__(self, filePath):
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
            self.scenes[self._scId].status = '1'

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





class CsvFile(FileExport):
    """csv file representation.
    * Records are separated by line breaks.
    * Data fields are delimited by the _SEPARATOR character.
    """

    EXTENSION = '.csv'
    # overwrites Novel.EXTENSION

    _SEPARATOR = '|'
    # delimits data fields within a record.

    _LINEBREAK = '\t'
    # substitutes embedded line breaks.

    _LIST_SEPARATOR = ','
    # delimits items listed within a data field

    CSV_REPLACEMENTS = [
        ['\n', _LINEBREAK],
    ]

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


class CsvSceneList(CsvFile):
    """csv file representation of an yWriter project's scenes table. 

    Represents a csv file with a record per scene.
    * Records are separated by line breaks.
    * Data fields are delimited by the _SEPARATOR character.
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
        '''$Title|$Desc|$Tags|$Notes|''' +\
        '''$ReactionScene|$Goal|$Conflict|$Outcome|''' +\
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
        try:
            with open(self._filePath, 'r', encoding='utf-8') as f:
                lines = (f.readlines())

        except(FileNotFoundError):
            return 'ERROR: "' + os.path.normpath(self._filePath) + '" not found.'

        cellsInLine = len(self.fileHeader.split(self._SEPARATOR))

        for line in lines:
            cell = line.rstrip().split(self._SEPARATOR)

            if len(cell) != cellsInLine:
                return 'ERROR: Wrong cell structure.'

            i = 0

            if 'ScID:' in cell[i]:
                scId = re.search('ScID\:([0-9]+)', cell[0]).group(1)
                self.scenes[scId] = Scene()
                i += 1
                self.scenes[scId].title = cell[i]
                i += 1
                self.scenes[scId].desc = self.convert_to_yw(cell[i])
                i += 1
                self.scenes[scId].tags = cell[i].split(self._LIST_SEPARATOR)
                i += 1
                self.scenes[scId].sceneNotes = self.convert_to_yw(cell[i])
                i += 1

                if Scene.REACTION_MARKER.lower() in cell[i].lower():
                    self.scenes[scId].isReactionScene = True

                else:
                    self.scenes[scId].isReactionScene = False

                i += 1
                self.scenes[scId].goal = cell[i]
                i += 1
                self.scenes[scId].conflict = cell[i]
                i += 1
                self.scenes[scId].outcome = cell[i]
                i += 1
                # Don't write back sceneCount
                i += 1
                # Don't write back wordCount
                i += 1

                # Transfer scene ratings; set to 1 if deleted

                if cell[i] in self._SCENE_RATINGS:
                    self.scenes[scId].field1 = cell[i]

                else:
                    self.scenes[scId].field1 = '1'

                i += 1

                if cell[i] in self._SCENE_RATINGS:
                    self.scenes[scId].field2 = cell[i]

                else:
                    self.scenes[scId].field2 = '1'

                i += 1

                if cell[i] in self._SCENE_RATINGS:
                    self.scenes[scId].field3 = cell[i]

                else:
                    self.scenes[scId].field3 = '1'

                i += 1

                if cell[i] in self._SCENE_RATINGS:
                    self.scenes[scId].field4 = cell[i]

                else:
                    self.scenes[scId].field4 = '1'

                i += 1
                # Don't write back scene words total
                i += 1
                # Don't write back scene letters total
                i += 1

                try:
                    self.scenes[scId].status = Scene.STATUS.index(cell[i])

                except ValueError:
                    pass
                    # Scene status remains None and will be ignored when
                    # writing back.

                i += 1
                ''' Cannot write back character IDs, because self.characters is None
                charaNames = cell[i].split(self._LIST_SEPARATOR)
                self.scenes[scId].characters = []

                for charaName in charaNames:

                    for id, name in self.characters.items():

                        if name == charaName:
                            self.scenes[scId].characters.append(id)
                '''
                i += 1
                ''' Cannot write back location IDs, because self.locations is None
                locaNames = cell[i].split(self._LIST_SEPARATOR)
                self.scenes[scId].locations = []

                for locaName in locaNames:

                    for id, name in self.locations.items():

                        if name == locaName:
                            self.scenes[scId].locations.append(id)
                '''
                i += 1
                ''' Cannot write back item IDs, because self.items is None
                itemNames = cell[i].split(self._LIST_SEPARATOR)
                self.scenes[scId].items = []

                for itemName in itemNames:

                    for id, name in self.items.items():

                        if name == itemName:
                            self.scenes[scId].items.append(id)
                '''

        return 'SUCCESS: Data read from "' + self._filePath + '".'




class CsvPlotList(CsvFile):
    """csv file representation of an yWriter project's scenes table. 

    Represents a csv file with a record per scene.
    * Records are separated by line breaks.
    * Data fields are delimited by the _SEPARATOR character.
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
        '''|$Tags|$Title|$Notes|''' +\
        '''$SceneNumber|$WordsTotal|$Field1|$Field2|$Field3|$Field4
'''

    def get_projectTemplateMapping(self):
        """Return a mapping dictionary for the project section. 
        """
        projectTemplateMapping = CsvFile.get_projectTemplateMapping(self)

        charList = []

        for crId in self.characters:
            charList.append(self.characters[crId].title)

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
        try:
            with open(self._filePath, 'r', encoding='utf-8') as f:
                lines = (f.readlines())

        except(FileNotFoundError):
            return 'ERROR: "' + os.path.normpath(self._filePath) + '" not found.'

        cellsInLine = len(self.fileHeader.split(self._SEPARATOR))

        tableHeader = lines[0].rstrip().split(self._SEPARATOR)

        for line in lines:
            cell = line.rstrip().split(self._SEPARATOR)

            if len(cell) != cellsInLine:
                return 'ERROR: Wrong cell structure.'

            if 'ChID:' in cell[0]:
                chId = re.search('ChID\:([0-9]+)', cell[0]).group(1)
                self.chapters[chId] = Chapter()
                self.chapters[chId].title = cell[1]
                self.chapters[chId].desc = self.convert_to_yw(cell[4])

            if 'ScID:' in cell[0]:
                scId = re.search('ScID\:([0-9]+)', cell[0]).group(1)
                self.scenes[scId] = Scene()
                self.scenes[scId].tags = cell[2].split(self._LIST_SEPARATOR)
                self.scenes[scId].title = cell[3]
                self.scenes[scId].sceneNotes = self.convert_to_yw(cell[4])

                i = 5
                # Don't write back sceneCount
                i += 1
                # Don't write back wordCount
                i += 1

                # Transfer scene ratings; set to 1 if deleted

                if cell[i] in self._SCENE_RATINGS:
                    self.scenes[scId].field1 = cell[i]

                elif tableHeader[i] != self._NOT_APPLICABLE:
                    self.scenes[scId].field1 = '1'

                i += 1

                if cell[i] in self._SCENE_RATINGS:
                    self.scenes[scId].field2 = cell[i]

                elif tableHeader[i] != self._NOT_APPLICABLE:
                    self.scenes[scId].field2 = '1'

                i += 1

                if cell[i] in self._SCENE_RATINGS:
                    self.scenes[scId].field3 = cell[i]

                elif tableHeader[i] != self._NOT_APPLICABLE:
                    self.scenes[scId].field3 = '1'

                i += 1

                if cell[i] in self._SCENE_RATINGS:
                    self.scenes[scId].field4 = cell[i]

                elif tableHeader[i] != self._NOT_APPLICABLE:
                    self.scenes[scId].field4 = '1'

        return 'SUCCESS: Data read from "' + self._filePath + '".'




class CsvCharList(CsvFile):
    """csv file representation of an yWriter project's characters table. 

    Represents a csv file with a record per character.
    * Records are separated by line breaks.
    * Data fields are delimited by the _SEPARATOR character.
    """

    DESCRIPTION = 'Character list'
    SUFFIX = '_charlist'

    fileHeader = '''ID|Name|Full name|Aka|Description|Bio|Goals|Importance|Tags|Notes
'''

    characterTemplate = '''CrID:$ID|$Title|$FullName|$AKA|$Desc|$Bio|$Goals|$Status|$Tags|$Notes
'''

    def read(self):
        """Parse the csv file located at filePath, 
        fetching the Character attributes contained.
        Return a message beginning with SUCCESS or ERROR.
        """
        try:
            with open(self._filePath, 'r', encoding='utf-8') as f:
                lines = (f.readlines())

        except(FileNotFoundError):
            return 'ERROR: "' + os.path.normpath(self._filePath) + '" not found.'

        if lines[0] != self.fileHeader:
            return 'ERROR: Wrong lines content.'

        cellsInLine = len(self.fileHeader.split(self._SEPARATOR))

        for line in lines:
            cell = line.rstrip().split(self._SEPARATOR)

            if len(cell) != cellsInLine:
                return 'ERROR: Wrong cell structure.'

            if 'CrID:' in cell[0]:
                crId = re.search('CrID\:([0-9]+)', cell[0]).group(1)
                self.characters[crId] = Character()
                self.characters[crId].title = cell[1]
                self.characters[crId].fullName = cell[2]
                self.characters[crId].aka = cell[3]
                self.characters[crId].desc = self.convert_to_yw(cell[4])
                self.characters[crId].bio = cell[5]
                self.characters[crId].goals = cell[6]

                if Character.MAJOR_MARKER in cell[7]:
                    self.characters[crId].isMajor = True

                else:
                    self.characters[crId].isMajor = False

                self.characters[crId].tags = cell[8].split(';')
                self.characters[crId].notes = self.convert_to_yw(cell[9])

        return 'SUCCESS: Data read from "' + self._filePath + '".'

    def merge(self, novel):
        """Copy required attributes of the novel object.
        Return a message beginning with SUCCESS or ERROR.
        """
        self.characters = novel.characters
        return 'SUCCESS'




class CsvLocList(CsvFile):
    """csv file representation of an yWriter project's locations table. 

    Represents a csv file with a record per location.
    * Records are separated by line breaks.
    * Data fields are delimited by the _SEPARATOR location.
    """

    DESCRIPTION = 'Location list'
    SUFFIX = '_loclist'

    fileHeader = '''ID|Name|Description|Aka|Tags
'''

    locationTemplate = '''LcID:$ID|$Title|$Desc|$AKA|$Tags
'''

    def read(self):
        """Parse the csv file located at filePath, 
        fetching the Object attributes contained.
        Return a message beginning with SUCCESS or ERROR.
        """
        try:
            with open(self._filePath, 'r', encoding='utf-8') as f:
                lines = (f.readlines())

        except(FileNotFoundError):
            return 'ERROR: "' + os.path.normpath(self._filePath) + '" not found.'

        if lines[0] != self.fileHeader:
            return 'ERROR: Wrong lines content.'

        cellsInLine = len(self.fileHeader.split(self._SEPARATOR))

        for line in lines:
            cell = line.rstrip().split(self._SEPARATOR)

            if len(cell) != cellsInLine:
                return 'ERROR: Wrong cell structure.'

            if 'LcID:' in cell[0]:
                lcId = re.search('LcID\:([0-9]+)', cell[0]).group(1)
                self.locations[lcId] = Object()
                self.locations[lcId].title = cell[1]
                self.locations[lcId].desc = self.convert_to_yw(cell[2])
                self.locations[lcId].aka = cell[3]
                self.locations[lcId].tags = cell[4].split(';')

        return 'SUCCESS: Data read from "' + self._filePath + '".'

    def merge(self, novel):
        """Copy required attributes of the novel object.
        Return a message beginning with SUCCESS or ERROR.
        """
        self.locations = novel.locations
        return 'SUCCESS'




class CsvItemList(CsvFile):
    """csv file representation of an yWriter project's items table. 

    Represents a csv file with a record per item.
    * Records are separated by line breaks.
    * Data fields are delimited by the _SEPARATOR item.
    """

    DESCRIPTION = 'Item list'
    SUFFIX = '_itemlist'

    fileHeader = '''ID|Name|Description|Aka|Tags
'''

    itemTemplate = '''ItID:$ID|$Title|$Desc|$AKA|$Tags
'''

    def read(self):
        """Parse the csv file located at filePath, 
        fetching the Object attributes contained.
        Return a message beginning with SUCCESS or ERROR.
        """
        try:
            with open(self._filePath, 'r', encoding='utf-8') as f:
                lines = (f.readlines())

        except(FileNotFoundError):
            return 'ERROR: "' + os.path.normpath(self._filePath) + '" not found.'

        if lines[0] != self.fileHeader:
            return 'ERROR: Wrong lines content.'

        cellsInLine = len(self.fileHeader.split(self._SEPARATOR))

        for line in lines:
            cell = line.rstrip().split(self._SEPARATOR)

            if len(cell) != cellsInLine:
                return 'ERROR: Wrong cell structure.'

            if 'ItID:' in cell[0]:
                itId = re.search('ItID\:([0-9]+)', cell[0]).group(1)
                self.items[itId] = Object()
                self.items[itId].title = cell[1]
                self.items[itId].desc = self.convert_to_yw(cell[2])
                self.items[itId].aka = cell[3]
                self.items[itId].tags = cell[4].split(';')

        return 'SUCCESS: Data read from "' + self._filePath + '".'

    def merge(self, novel):
        """Copy required attributes of the novel object.
        Return a message beginning with SUCCESS or ERROR.
        """
        self.items = novel.items
        return 'SUCCESS'



class UniversalFileFactory(FileFactory):
    """A factory class that instantiates a source file object
    and a target file object for conversion.
    All filetypes are covered.
    """

    def get_file_objects(self, sourcePath, suffix=None):
        """Return a tuple with three elements:
        * A message string starting with 'SUCCESS' or 'ERROR'
        * sourceFile: a Novel subclass instance
        * targetFile: a Novel subclass instance
        """
        fileName, fileExtension = os.path.splitext(sourcePath)
        isYwProject = False

        if fileExtension == Yw7File.EXTENSION:
            sourceFile = Yw7File(sourcePath)
            isYwProject = True

        elif fileExtension == Yw5File.EXTENSION:
            sourceFile = Yw5File(sourcePath)
            isYwProject = True

        elif fileExtension == Yw6File.EXTENSION:
            sourceFile = Yw6File(sourcePath)
            isYwProject = True

        if isYwProject:

            # Determine which sort of target is required.

            if suffix is None:
                targetFile = Yw5File(fileName + Yw5File.EXTENSION)
                targetFile.ywTreeBuilder = Yw5TreeCreator()
                targetFile.ywProjectMerger = YwProjectCreator()

            elif suffix == '':
                targetFile = OdtExport(fileName + OdtExport.EXTENSION)

            elif suffix == OdtManuscript.SUFFIX:
                targetFile = OdtManuscript(
                    fileName + suffix + OdtManuscript.EXTENSION)

            elif suffix == OdtProof.SUFFIX:
                targetFile = OdtProof(fileName + suffix + OdtProof.EXTENSION)

            elif suffix == OdtSceneDesc.SUFFIX:
                targetFile = OdtSceneDesc(
                    fileName + suffix + OdtSceneDesc.EXTENSION)

            elif suffix == OdtChapterDesc.SUFFIX:
                targetFile = OdtChapterDesc(
                    fileName + suffix + OdtChapterDesc.EXTENSION)

            elif suffix == OdtPartDesc.SUFFIX:
                targetFile = OdtPartDesc(
                    fileName + suffix + OdtPartDesc.EXTENSION)

            elif suffix == OdtCharacters.SUFFIX:
                targetFile = OdtCharacters(
                    fileName + suffix + OdtCharacters.EXTENSION)

            elif suffix == OdtLocations.SUFFIX:
                targetFile = OdtLocations(
                    fileName + suffix + OdtLocations.EXTENSION)

            elif suffix == OdtItems.SUFFIX:
                targetFile = OdtItems(fileName + suffix + OdtItems.EXTENSION)

            elif suffix == CsvSceneList.SUFFIX:
                targetFile = CsvSceneList(
                    fileName + suffix + CsvSceneList.EXTENSION)

            elif suffix == CsvPlotList.SUFFIX:
                targetFile = CsvPlotList(
                    fileName + suffix + CsvPlotList.EXTENSION)

            elif suffix == CsvCharList.SUFFIX:
                targetFile = CsvCharList(
                    fileName + suffix + CsvCharList.EXTENSION)

            elif suffix == CsvLocList.SUFFIX:
                targetFile = CsvLocList(
                    fileName + suffix + CsvLocList.EXTENSION)

            elif suffix == CsvItemList.SUFFIX:
                targetFile = CsvItemList(
                    fileName + suffix + CsvItemList.EXTENSION)

            else:
                return 'ERROR: File type of "' + os.path.normpath(sourcePath) + '" not supported.', None, None

        else:
            # The source file is not a yWriter project.

            targetFile = None

            if sourcePath.endswith(HtmlProof.SUFFIX + HtmlProof.EXTENSION):
                sourceFile = HtmlProof(sourcePath)

            elif sourcePath.endswith(HtmlManuscript.SUFFIX + HtmlManuscript.EXTENSION):
                sourceFile = HtmlManuscript(sourcePath)

            elif sourcePath.endswith(HtmlSceneDesc.SUFFIX + HtmlSceneDesc.EXTENSION):
                sourceFile = HtmlSceneDesc(sourcePath)

            elif sourcePath.endswith(HtmlChapterDesc.SUFFIX + HtmlChapterDesc.EXTENSION):
                sourceFile = HtmlChapterDesc(sourcePath)

            elif sourcePath.endswith(HtmlPartDesc.SUFFIX + HtmlPartDesc.EXTENSION):
                sourceFile = HtmlPartDesc(sourcePath)

            elif sourcePath.endswith(HtmlCharacters.SUFFIX + HtmlCharacters.EXTENSION):
                sourceFile = HtmlCharacters(sourcePath)

            elif sourcePath.endswith(HtmlLocations.SUFFIX + HtmlLocations.EXTENSION):
                sourceFile = HtmlLocations(sourcePath)

            elif sourcePath.endswith(HtmlItems.SUFFIX + HtmlItems.EXTENSION):
                sourceFile = HtmlItems(sourcePath)

            elif sourcePath.endswith('.html'):

                # Is the source file an outline or a "work in progress"?

                result = read_html_file(sourcePath)

                if 'SUCCESS' in result[0]:
                    targetFile = Yw7File(fileName + Yw7File.EXTENSION)
                    targetFile.ywTreeBuilder = Yw7TreeCreator()
                    targetFile.ywProjectMerger = YwProjectCreator()

                    if "<h3" in result[1].lower():
                        sourceFile = HtmlOutline(sourcePath)

                    else:
                        sourceFile = HtmlImport(sourcePath)

                else:
                    return 'ERROR: Cannot read "' + os.path.normpath(sourcePath) + '".', None, None

            elif sourcePath.endswith(CsvSceneList.SUFFIX + CsvSceneList.EXTENSION):
                sourceFile = CsvSceneList(sourcePath)

            elif sourcePath.endswith(CsvPlotList.SUFFIX + CsvPlotList.EXTENSION):
                sourceFile = CsvPlotList(sourcePath)

            elif sourcePath.endswith(CsvCharList.SUFFIX + CsvCharList.EXTENSION):
                sourceFile = CsvCharList(sourcePath)

            elif sourcePath.endswith(CsvLocList.SUFFIX + CsvLocList.EXTENSION):
                sourceFile = CsvLocList(sourcePath)

            elif sourcePath.endswith(CsvItemList.SUFFIX + CsvItemList.EXTENSION):
                sourceFile = CsvItemList(sourcePath)

            else:
                return 'ERROR: File type of  "' + os.path.normpath(sourcePath) + '" not supported.', None, None

            if targetFile is None:

                ywPathBasis = fileName.split(sourceFile.SUFFIX)[0]

                # Look for an existing yWriter project to rewrite.

                if os.path.isfile(ywPathBasis + Yw7File.EXTENSION):
                    targetFile = Yw7File(ywPathBasis + Yw7File.EXTENSION)

                elif os.path.isfile(ywPathBasis + Yw5File.EXTENSION):
                    targetFile = Yw5File(ywPathBasis + Yw5File.EXTENSION)

                elif os.path.isfile(ywPathBasis + Yw6File.EXTENSION):
                    targetFile = Yw6File(ywPathBasis + Yw6File.EXTENSION)

            if targetFile is None:
                return 'ERROR: No yWriter project to write.', None, None

        return 'SUCCESS', sourceFile, targetFile

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


def msgbox(message, title='LibreOffice', buttons=BUTTONS_OK, type_msg=INFOBOX):
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





class Ui():
    """Superclass for UI facades, implementing a 'silent mode'."""

    def __init__(self, title):
        self.infoWhatText = ''
        self.infoHowText = ''

    def ask_yes_no(self, text):
        return True

    def set_info_what(self, message):
        """What's the converter going to do?"""
        self.infoWhatText = message

    def set_info_how(self, message):
        """How's the converter doing?"""
        self.infoHowText = message

    def finish(self):
        """To be overwritten by subclasses requiring
        special action to finish user interaction.
        """


class YwCnv():
    """Converter for yWriter project files.
    """

    def convert(self, sourceFile, targetFile):
        """Read document file, convert its content to xml, and replace yWriter file.
        Return a message beginning with SUCCESS or ERROR.
        """

        if sourceFile.filePath is None:
            return 'ERROR: "' + os.path.normpath(sourceFile.filePath) + '" is not of the supported type.'

        if not sourceFile.file_exists():
            return 'ERROR: "' + os.path.normpath(sourceFile.filePath) + '" not found.'

        if targetFile.filePath is None:
            return 'ERROR: "' + os.path.normpath(targetFile.filePath) + '" is not of the supported type.'

        if targetFile.file_exists() and not self.confirm_overwrite(targetFile.filePath):
            return 'Program abort by user.'

        message = sourceFile.read()

        if message.startswith('ERROR'):
            return message

        message = targetFile.merge(sourceFile)

        if message.startswith('ERROR'):
            return message

        return targetFile.write()

    def confirm_overwrite(self, fileName):
        """Hook for subclasses with UI."""
        return True


class YwCnvUi(YwCnv):
    """Standalone yWriter converter with a simple tkinter GUI. 
    """

    YW_EXTENSIONS = ['.yw5', '.yw6', '.yw7']

    def __init__(self):
        self.userInterface = Ui('yWriter import/export')
        self.success = False
        self.fileFactory = None

    def run(self, sourcePath, suffix=None):
        """Create source and target objects and run conversion.
        """
        message, sourceFile, targetFile = self.fileFactory.get_file_objects(
            sourcePath, suffix)

        if not message.startswith('SUCCESS'):
            self.userInterface.set_info_how(message)

        elif not sourceFile.file_exists():
            self.userInterface.set_info_how(
                'ERROR: File "' + os.path.normpath(sourceFile.filePath) + '" not found.')

        elif sourceFile.EXTENSION in self.YW_EXTENSIONS:
            self.export_from_yw(sourceFile, targetFile)

        elif isinstance(targetFile.ywTreeBuilder, Yw7TreeCreator):
            self.create_yw7(sourceFile, targetFile)

        else:
            self.import_to_yw(sourceFile, targetFile)

        self.finish(sourcePath)

    def export_from_yw(self, sourceFile, targetFile):
        """Template method for conversion from yw to other.
        """
        self.userInterface.set_info_what('Input: ' + sourceFile.DESCRIPTION + ' "' + os.path.normpath(
            sourceFile.filePath) + '"\nOutput: ' + targetFile.DESCRIPTION + ' "' + os.path.normpath(targetFile.filePath) + '"')
        message = self.convert(sourceFile, targetFile)
        self.userInterface.set_info_how(message)

        if message.startswith('SUCCESS'):
            self.success = True

    def create_yw7(self, sourceFile, targetFile):
        """Template method for creation of a new yw7 project.
        """

        if targetFile.file_exists():
            self.userInterface.set_info_how(
                'ERROR: "' + os.path.normpath(targetFile._filePath) + '" already exists.')

        else:
            self.userInterface.set_info_what(
                'Create a yWriter project file from ' + sourceFile.DESCRIPTION + '\nNew project: "' + os.path.normpath(targetFile.filePath) + '"')
            message = self.convert(sourceFile, targetFile)
            self.userInterface.set_info_how(message)

            if message.startswith('SUCCESS'):
                self.success = True

    def import_to_yw(self, sourceFile, targetFile):
        """Template method for conversion from other to yw.
        """
        self.userInterface.set_info_what('Input: ' + sourceFile.DESCRIPTION + ' "' + os.path.normpath(
            sourceFile.filePath) + '"\nOutput: ' + targetFile.DESCRIPTION + ' "' + os.path.normpath(targetFile.filePath) + '"')
        message = self.convert(sourceFile, targetFile)
        self.userInterface.set_info_how(message)

        if message.startswith('SUCCESS'):
            self.success = True

    def confirm_overwrite(self, filePath):
        """ Invoked by the parent if a file already exists.
        """
        return self.userInterface.ask_yes_no('Overwrite existing file "' + os.path.normpath(filePath) + '"?')

    def delete_tempfile(self, filePath):
        """If an Office file exists, delete the temporary file."""

        if filePath.endswith('.html'):

            if os.path.isfile(filePath.replace('.html', '.odt')):

                try:
                    os.remove(filePath)

                except:
                    pass

        elif filePath.endswith('.csv'):

            if os.path.isfile(filePath.replace('.csv', '.ods')):

                try:
                    os.remove(filePath)

                except:
                    pass

    def finish(self, sourcePath):
        """Hook for actions to take place after the conversion."""
from com.sun.star.awt.MessageBoxResults import OK, YES, NO, CANCEL
from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK, BUTTONS_OK_CANCEL, BUTTONS_YES_NO, BUTTONS_YES_NO_CANCEL, BUTTONS_RETRY_CANCEL, BUTTONS_ABORT_IGNORE_RETRY
from com.sun.star.awt.MessageBoxType import MESSAGEBOX, INFOBOX, WARNINGBOX, ERRORBOX, QUERYBOX



class UiUno(Ui):
    """UI subclass implementing a LibreOffice UNO facade."""

    def ask_yes_no(self, text):
        result = msgbox(text, 'WARNING', buttons=BUTTONS_YES_NO,
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


class YwCnvUno(YwCnvUi):
    """Converter for yWriter project files.
    Variant with UNO UI.
    """

    def __init__(self, silentMode=False):
        if silentMode:
            self.userInterface = Ui('')

        else:
            self.userInterface = UiUno('yWriter import/export')

        self.success = False
        self.fileFactory = None

    def finish(self, sourcePath):

        if self.success:
            self.delete_tempfile(sourcePath)

        self.userInterface.finish()

    def export_from_yw(self, sourceFile, targetFile):
        """Method for conversion from yw to other.
        Show only error messages.
        """
        message = self.convert(sourceFile, targetFile)

        if message.startswith('SUCCESS'):
            self.success = True

        else:
            self.userInterface.set_info_how(message)

INI_FILE = 'openyw.ini'


class Converter(YwCnvUno):
    """Converter for yWriter project files.
    Variant with UNO UI.
    """

    def __init__(self, silentMode=False):
        YwCnvUno.__init__(self, silentMode)
        self.fileFactory = UniversalFileFactory()


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
    converter = Converter()
    converter.run(sourcePath, suffix)

    if converter.success:
        desktop = XSCRIPTCONTEXT.getDesktop()
        doc = desktop.loadComponentFromURL(newFile, "_blank", 0, ())


def import_yw(*args):
    '''Import scenes from yWriter 6/7 to a Writer document
    without chapter and scene markers. 
    '''
    open_yw7('', '.odt')


def proof_yw(*args):
    '''Import scenes from yWriter 6/7 to a Writer document
    with visible chapter and scene markers. 
    '''
    open_yw7(OdtProof.SUFFIX, OdtProof.EXTENSION)


def get_manuscript(*args):
    '''Import scenes from yWriter 6/7 to a Writer document
    with invisible chapter and scene markers. 
    '''
    open_yw7(OdtManuscript.SUFFIX, OdtManuscript.EXTENSION)


def get_partdesc(*args):
    '''Import pard descriptions from yWriter 6/7 to a Writer document
    with invisible chapter and scene markers. 
    '''
    open_yw7(OdtPartDesc.SUFFIX, OdtPartDesc.EXTENSION)


def get_chapterdesc(*args):
    '''Import chapter descriptions from yWriter 6/7 to a Writer document
    with invisible chapter and scene markers. 
    '''
    open_yw7(OdtChapterDesc.SUFFIX, OdtChapterDesc.EXTENSION)


def get_scenedesc(*args):
    '''Import scene descriptions from yWriter 6/7 to a Writer document
    with invisible chapter and scene markers. 
    '''
    open_yw7(OdtSceneDesc.SUFFIX, OdtSceneDesc.EXTENSION)


def get_chardesc(*args):
    '''Import character descriptions from yWriter 6/7 to a Writer document.
    '''
    open_yw7(OdtCharacters.SUFFIX, OdtCharacters.EXTENSION)


def get_locdesc(*args):
    '''Import location descriptions from yWriter 6/7 to a Writer document.
    '''
    open_yw7(OdtLocations.SUFFIX, OdtLocations.EXTENSION)


def get_itemdesc(*args):
    '''Import item descriptions from yWriter 6/7 to a Writer document.
    '''
    open_yw7(OdtItems.SUFFIX, OdtItems.EXTENSION)


def get_scenelist(*args):
    '''Import a scene list from yWriter 6/7 to a Calc document.
    '''
    open_yw7(CsvSceneList.SUFFIX, CsvSceneList.EXTENSION)


def get_plotlist(*args):
    '''Import a plot list from yWriter 6/7 to a Calc document.
    '''
    open_yw7(CsvPlotList.SUFFIX, CsvPlotList.EXTENSION)


def get_charlist(*args):
    '''Import a character list from yWriter 6/7 to a Calc document.
    '''
    open_yw7(CsvCharList.SUFFIX, CsvCharList.EXTENSION)


def get_loclist(*args):
    '''Import a location list from yWriter 6/7 to a Calc document.
    '''
    open_yw7(CsvLocList.SUFFIX, CsvLocList.EXTENSION)


def get_itemlist(*args):
    '''Import an item list from yWriter 6/7 to a Calc document.
    '''
    open_yw7(CsvItemList.SUFFIX, CsvItemList.EXTENSION)


def export_yw(*args):
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

        args1[0].Name = 'URL'
        # args1(0).Name = "URL"
        args1[0].Value = csvPath
        # args1(0).Value = csvPath
        args1[1].Name = 'FilterName'
        # args1(1).Name = "FilterName"
        args1[1].Value = 'Text - txt - csv (StarCalc)'
        # args1(1).Value = "Text - txt - csv (StarCalc)"
        dispatcher.executeDispatch(document, ".uno:SaveAs", "", 0, args1)
        # dispatcher.executeDispatch(document, ".uno:SaveAs", "", 0, args1())

        # Save document in OpenDocument format

        args1.append(PropertyValue())

        args1[0].Value = odsPath
        # args1(0).Value = odsPath
        args1[1].Value = 'calc8'
        # args1(1).Value = "calc8"
        args1[2].Name = "FilterOptions"
        # args1(2).Name = "FilterOptions"
        args1[2].Value = "124,34,76,1,,0,false,true,true"
        # args1(2).Value = "124,34,76,1,,0,false,true,true"
        dispatcher.executeDispatch(document, ".uno:SaveAs", "", 0, args1)
        # dispatcher.executeDispatch(document, ".uno:SaveAs", "", 0, args1())

        targetPath = unquote(csvPath.replace('file:///', ''))

    else:
        msgbox('ERROR: File type of "' + os.path.normpath(documentPath) +
               '" not supported.', type_msg=ERRORBOX)

    Converter().run(targetPath, None)
