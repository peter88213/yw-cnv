"""Convert html or csv to yWriter format. 

Input file format: html (with visible or invisible chapter and scene tags).

Version 0.29.4

Copyright (c) 2020 Peter Triesberger
For further information see https://github.com/peter88213/yw-cnv
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import sys
import os

from urllib.parse import unquote


import re

from html.parser import HTMLParser

from abc import abstractmethod
from urllib.parse import quote


class Novel():
    """Abstract yWriter project file representation.

    This class represents a file containing a novel with additional 
    attributes and structural information (a full set or a subset
    of the information included in an yWriter project file).
    """

    EXTENSION = ''
    SUFFIX = ''
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
        # supported type as specified by _FILE_EXTENSION.

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
        if filePath.lower().endswith(self.SUFFIX + self.EXTENSION):
            self._filePath = filePath
            head, tail = os.path.split(os.path.realpath(filePath))
            self.projectPath = quote(head.replace('\\', '/'), '/:')
            self.projectName = quote(tail.replace(
                self.SUFFIX + self.EXTENSION, ''))

    @abstractmethod
    def read(self):
        """Parse the file and store selected properties.
        To be overwritten by file format specific subclasses.
        """

    @abstractmethod
    def merge(self, novel):
        """Merge selected novel properties.
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

    def get_structure(self):
        """returns a string showing the order of chapters and scenes 
        as a tree. The result can be used to compare two Novel objects 
        by their structure.
        """
        lines = []

        for chId in self.srtChapters:
            lines.append('ChID:' + str(chId) + '\n')

            for scId in self.chapters[chId].srtScenes:
                lines.append('  ScID:' + str(scId) + '\n')

        return ''.join(lines)


class Chapter():
    """yWriter chapter representation.
    # xml: <CHAPTERS><CHAPTER>
    """

    stripChapterFromTitle = False
    # bool
    # True: Remove 'Chapter ' from the chapter title upon import.
    # False: Do not modify the chapter title.

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

        self.chType = None
        # int
        # xml: <Type>
        # 0 = chapter type (marked "Chapter")
        # 1 = other type (marked "Other")

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

        self.doNotExport = None
        # bool
        # xml: <Fields><Field_SuppressChapterBreak> 0

        self.srtScenes = []
        # list of str
        # xml: <Scenes><ScID>
        # The chapter's scene IDs. The order of its elements
        # corresponds to the chapter's order of the scenes.

    def get_title(self):
        """Fix auto-chapter titles for non-English """
        text = self.title

        if self.stripChapterFromTitle:
            text = text.replace('Chapter ', '')

        return text



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
            return ('ERROR: "' + filePath + '" not found.', None)

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

        return 'SUCCESS: ' + str(len(self.scenes)) + ' Scenes read from "' + self._filePath + '".'



class HtmlProof(HtmlFile):
    """HTML file representation of an yWriter project's OfficeFile part.

    Represents a html file with visible chapter and scene tags 
    to be read and written by Open/LibreOffice Writer.
    """

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

    def get_structure(self):
        """This file format has no comparable structure."""



class HtmlSceneDesc(HtmlFile):
    """HTML file representation of an yWriter project's scene summaries.
    """

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

    def get_structure(self):
        """This file format has no comparable structure."""



class HtmlChapterDesc(HtmlFile):
    """HTML file representation of an yWriter project's chapters summaries."""

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

    def get_structure(self):
        """This file format has no comparable structure."""



class HtmlPartDesc(HtmlChapterDesc):
    """HTML file representation of an yWriter project's parts summaries."""

    SUFFIX = '_parts'





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


class HtmlCharacters(HtmlFile):
    """HTML file representation of an yWriter project's character descriptions."""

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

    def get_structure(self):
        """This file format has no comparable structure."""




class HtmlLocations(HtmlFile):
    """HTML file representation of an yWriter project's location descriptions."""

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

    def get_structure(self):
        """This file format has no comparable structure."""



class HtmlItems(HtmlFile):
    """HTML file representation of an yWriter project's item descriptions."""

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

    def get_structure(self):
        """This file format has no comparable structure."""


from string import Template



class FileExport(Novel):
    """Abstract yWriter project file exporter representation.
    """

    fileHeader = ''
    partTemplate = ''
    chapterTemplate = ''
    unusedChapterTemplate = ''
    infoChapterTemplate = ''
    sceneTemplate = ''
    unusedSceneTemplate = ''
    infoSceneTemplate = ''
    sceneDivider = ''
    chapterEndTemplate = ''
    unusedChapterEndTemplate = ''
    infoChapterEndTemplate = ''
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
        """Copy selected novel attributes.
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

    def get_projectTemplateSubst(self):
        return dict(
            Title=self.title,
            Desc=self.convert_from_yw(self.desc),
            AuthorName=self.author,
            FieldTitle1=self.fieldTitle1,
            FieldTitle2=self.fieldTitle2,
            FieldTitle3=self.fieldTitle3,
            FieldTitle4=self.fieldTitle4,
        )

    def get_chapterSubst(self, chId, chapterNumber):
        return dict(
            ID=chId,
            ChapterNumber=chapterNumber,
            Title=self.chapters[chId].title,
            Desc=self.convert_from_yw(self.chapters[chId].desc),
            ProjectName=self.projectName,
            ProjectPath=self.projectPath,
        )

    def get_sceneSubst(self, scId, sceneNumber, wordsTotal, lettersTotal):

        if self.scenes[scId].tags is not None:
            tags = ', '.join(self.scenes[scId].tags)

        else:
            tags = ''

        if self.scenes[scId].characters is not None:
            sChList = []

            for chId in self.scenes[scId].characters:
                sChList.append(self.characters[chId].title)

            sceneChars = ', '.join(sChList)
            viewpointChar = sChList[0]

        else:
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

        return dict(
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

    def get_characterSubst(self, crId):

        if self.characters[crId].tags is not None:
            tags = ', '.join(self.characters[crId].tags)

        else:
            tags = ''

        if self.characters[crId].isMajor:
            characterStatus = Character.MAJOR_MARKER

        else:
            characterStatus = Character.MINOR_MARKER

        return dict(
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

    def get_locationSubst(self, lcId):

        if self.locations[lcId].tags is not None:
            tags = ', '.join(self.locations[lcId].tags)

        else:
            tags = ''

        return dict(
            ID=lcId,
            Title=self.locations[lcId].title,
            Desc=self.convert_from_yw(self.locations[lcId].desc),
            Tags=tags,
            AKA=FileExport.convert_from_yw(self, self.locations[lcId].aka),
        )

    def get_itemSubst(self, itId):

        if self.items[itId].tags is not None:
            tags = ', '.join(self.items[itId].tags)

        else:
            tags = ''

        return dict(
            ID=itId,
            Title=self.items[itId].title,
            Desc=self.convert_from_yw(self.items[itId].desc),
            Tags=tags,
            AKA=FileExport.convert_from_yw(self, self.items[itId].aka),
        )

    def write(self):
        lines = []
        wordsTotal = 0
        lettersTotal = 0
        chapterNumber = 0
        sceneNumber = 0

        template = Template(self.fileHeader)
        lines.append(template.safe_substitute(self.get_projectTemplateSubst()))

        for chId in self.srtChapters:

            if self.chapters[chId].isUnused:

                if self.unusedChapterTemplate != '':
                    template = Template(self.unusedChapterTemplate)

                else:
                    continue

            elif self.chapters[chId].chType != 0:

                if self.infoChapterTemplate != '':
                    template = Template(self.infoChapterTemplate)

                else:
                    continue

            elif self.chapters[chId].chLevel == 1 and self.partTemplate != '':
                template = Template(self.partTemplate)

            else:
                template = Template(self.chapterTemplate)
                chapterNumber += 1

            lines.append(template.safe_substitute(
                self.get_chapterSubst(chId, chapterNumber)))
            firstSceneInChapter = True

            for scId in self.chapters[chId].srtScenes:
                wordsTotal += self.scenes[scId].wordCount
                lettersTotal += self.scenes[scId].letterCount

                if self.scenes[scId].isUnused or self.chapters[chId].isUnused or self.scenes[scId].doNotExport:

                    if self.unusedSceneTemplate != '':
                        template = Template(self.unusedSceneTemplate)

                    else:
                        continue

                elif self.chapters[chId].chType != 0:

                    if self.infoSceneTemplate != '':
                        template = Template(self.infoSceneTemplate)

                    else:
                        continue

                else:
                    sceneNumber += 1
                    template = Template(self.sceneTemplate)

                if not (firstSceneInChapter or self.scenes[scId].appendToPrev):
                    lines.append(self.sceneDivider)

                lines.append(template.safe_substitute(self.get_sceneSubst(
                    scId, sceneNumber, wordsTotal, lettersTotal)))

                firstSceneInChapter = False

            if self.chapters[chId].isUnused and self.unusedChapterEndTemplate != '':
                lines.append(self.unusedChapterEndTemplate)

            elif self.chapters[chId].chType != 0 and self.infoChapterEndTemplate != '':
                lines.append(self.infoChapterEndTemplate)

            else:
                lines.append(self.chapterEndTemplate)

        for crId in self.characters:
            template = Template(self.characterTemplate)
            lines.append(template.safe_substitute(
                self.get_characterSubst(crId)))

        for lcId in self.locations:
            template = Template(self.locationTemplate)
            lines.append(template.safe_substitute(
                self.get_locationSubst(lcId)))

        for itId in self.items:
            template = Template(self.itemTemplate)
            lines.append(template.safe_substitute(self.get_itemSubst(itId)))

        lines.append(self.fileFooter)
        text = ''.join(lines)

        try:
            with open(self.filePath, 'w', encoding='utf-8') as f:
                f.write(text)

        except:
            return 'ERROR: Cannot write "' + self.filePath + '".'

        return 'SUCCESS: Content written to "' + self.filePath + '".'


class CsvFile(FileExport):
    """csv file representation.
    * Records are separated by line breaks.
    * Data fields are delimited by the _SEPARATOR character.
    """

    EXTENSION = '.csv'
    # overwrites Novel._FILE_EXTENSION

    _SEPARATOR = '|'
    # delimits data fields within a record.

    _LINEBREAK = '\t'
    # substitutes embedded line breaks.

    _LIST_SEPARATOR = ','
    # delimits items listed within a data field

    def convert_from_yw(self, text):
        """Convert line breaks."""

        try:
            text = text.rstrip().replace('\n', self._LINEBREAK)

        except AttributeError:
            text = ''

        return text

    def convert_to_yw(self, text):
        """Convert line breaks."""

        try:
            text = text.replace(self._LINEBREAK, '\n')

        except AttributeError:
            text = ''

        return text

    def get_structure(self):
        """This file format has no comparable structure."""
        return None


class CsvSceneList(CsvFile):
    """csv file representation of an yWriter project's scenes table. 

    Represents a csv file with a record per scene.
    * Records are separated by line breaks.
    * Data fields are delimited by the _SEPARATOR character.
    """

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

    def get_sceneSubst(self, scId, sceneNumber, wordsTotal, lettersTotal):
        sceneSubst = CsvFile.get_sceneSubst(
            self, scId, sceneNumber, wordsTotal, lettersTotal)

        if self.scenes[scId].field1 == '1':
            sceneSubst['Field1'] = ''

        if self.scenes[scId].field2 == '1':
            sceneSubst['Field2'] = ''

        if self.scenes[scId].field3 == '1':
            sceneSubst['Field3'] = ''

        if self.scenes[scId].field4 == '1':
            sceneSubst['Field4'] = ''

        return sceneSubst

    def read(self):
        """Parse the csv file located at filePath, 
        fetching the Scene attributes contained.
        Return a message beginning with SUCCESS or ERROR.
        """
        try:
            with open(self._filePath, 'r', encoding='utf-8') as f:
                lines = (f.readlines())

        except(FileNotFoundError):
            return 'ERROR: "' + self._filePath + '" not found.'

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

    EXTENSION = '.csv'
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

    infoChapterTemplate = '''ChID:$ID|$Title|||$Desc||||||
'''

    sceneTemplate = '''=HYPERLINK("file:///$ProjectPath/${ProjectName}_manuscript.odt#ScID:$ID%7Cregion";"ScID:$ID")|''' +\
        '''|$Tags|$Title|$Notes|''' +\
        '''$SceneNumber|$WordsTotal|$Field1|$Field2|$Field3|$Field4
'''

    def get_projectTemplateSubst(self):
        projectTemplateSubst = CsvFile.get_projectTemplateSubst(self)

        charList = []

        for crId in self.characters:
            charList.append(self.characters[crId].title)

        if self.fieldTitle1 in charList or self._STORYLINE_MARKER in self.fieldTitle1.lower():
            self.arc1 = True

        else:
            self.arc1 = False
            projectTemplateSubst['FieldTitle1'] = self._NOT_APPLICABLE

        if self.fieldTitle2 in charList or self._STORYLINE_MARKER in self.fieldTitle2.lower():
            self.arc2 = True

        else:
            self.arc2 = False
            projectTemplateSubst['FieldTitle2'] = self._NOT_APPLICABLE

        if self.fieldTitle3 in charList or self._STORYLINE_MARKER in self.fieldTitle3.lower():
            self.arc3 = True

        else:
            self.arc3 = False
            projectTemplateSubst['FieldTitle3'] = self._NOT_APPLICABLE

        if self.fieldTitle4 in charList or self._STORYLINE_MARKER in self.fieldTitle4.lower():
            self.arc4 = True

        else:
            self.arc4 = False
            projectTemplateSubst['FieldTitle4'] = self._NOT_APPLICABLE

        return projectTemplateSubst

    def get_sceneSubst(self, scId, sceneNumber, wordsTotal, lettersTotal):
        sceneSubst = CsvFile.get_sceneSubst(
            self, scId, sceneNumber, wordsTotal, lettersTotal)

        if self.scenes[scId].field1 == '1' or not self.arc1:
            sceneSubst['Field1'] = ''

        if self.scenes[scId].field2 == '1' or not self.arc1:
            sceneSubst['Field2'] = ''

        if self.scenes[scId].field3 == '1' or not self.arc3:
            sceneSubst['Field3'] = ''

        if self.scenes[scId].field4 == '1' or not self.arc4:
            sceneSubst['Field4'] = ''

        return sceneSubst

    def read(self):
        """Parse the csv file located at filePath, fetching 
        the Scene attributes contained.
        Return a message beginning with SUCCESS or ERROR.
        """
        try:
            with open(self._filePath, 'r', encoding='utf-8') as f:
                lines = (f.readlines())

        except(FileNotFoundError):
            return 'ERROR: "' + self._filePath + '" not found.'

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
            return 'ERROR: "' + self._filePath + '" not found.'

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
        """Copy selected novel attributes.
        """
        self.characters = novel.characters




class CsvLocList(CsvFile):
    """csv file representation of an yWriter project's locations table. 

    Represents a csv file with a record per location.
    * Records are separated by line breaks.
    * Data fields are delimited by the _SEPARATOR location.
    """

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
            return 'ERROR: "' + self._filePath + '" not found.'

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
        """Copy selected novel attributes.
        """
        self.locations = novel.locations




class CsvItemList(CsvFile):
    """csv file representation of an yWriter project's items table. 

    Represents a csv file with a record per item.
    * Records are separated by line breaks.
    * Data fields are delimited by the _SEPARATOR item.
    """

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
            return 'ERROR: "' + self._filePath + '" not found.'

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
        """Copy selected novel attributes.
        """
        self.items = novel.items



class HtmlImport(HtmlFile):
    """HTML file representation of an yWriter project's OfficeFile part.

    Represents a html file without chapter and scene tags 
    to be written by Open/LibreOffice Writer.
    """

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
            self.chapters[self._chId].chType = '0'

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

    def get_structure(self):
        """This file format has no comparable structure."""



class HtmlOutline(HtmlFile):
    """HTML file representation of an yWriter project's OfficeFile part.

    Represents a html file without chapter and scene tags 
    to be written by Open/LibreOffice Writer.
    """

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
            self.chapters[self._chId].chType = '0'

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

    def get_structure(self):
        """This file format has no comparable structure."""

import xml.etree.ElementTree as ET

from html import unescape

EM_DASH = '—'
EN_DASH = '–'
SAFE_DASH = '--'


def replace_unsafe_glyphs(text):
    """Replace glyphs being corrupted by yWriter with safe substitutes. """
    return text.replace(EN_DASH, SAFE_DASH).replace(EM_DASH, SAFE_DASH)


def indent(elem, level=0):
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
            indent(elem, level + 1)

        if not elem.tail or not elem.tail.strip():
            elem.tail = i

    else:
        if level and (not elem.tail or not elem.tail.strip()):
            elem.tail = i


def xml_postprocess(filePath, fileEncoding, cdataTags: list):
    '''Postprocess the xml file created by ElementTree:
       Put a header on top, insert the missing CDATA tags,
       and replace "ampersand" xml entity by plain text.
    '''
    with open(filePath, 'r', encoding=fileEncoding) as f:
        lines = f.readlines()

    newlines = ['<?xml version="1.0" encoding="' + fileEncoding + '"?>\n']

    for line in lines:

        for tag in cdataTags:
            line = re.sub('\<' + tag + '\>', '<' +
                          tag + '><![CDATA[', line)
            line = re.sub('\<\/' + tag + '\>',
                          ']]></' + tag + '>', line)

        newlines.append(line)

    newXml = ''.join(newlines)
    newXml = newXml.replace('[CDATA[ \n', '[CDATA[')
    newXml = newXml.replace('\n]]', ']]')
    newXml = unescape(newXml)

    try:
        with open(filePath, 'w', encoding=fileEncoding) as f:
            f.write(newXml)

    except:
        return 'ERROR: Can not write"' + filePath + '".'

    return 'SUCCESS: "' + filePath + '" written.'




class YwFile(Novel):
    """yWriter xml project file representation."""

    def __init__(self, filePath):
        Novel.__init__(self, filePath)
        self._cdataTags = ['Title', 'AuthorName', 'Bio', 'Desc',
                           'FieldTitle1', 'FieldTitle2', 'FieldTitle3',
                           'FieldTitle4', 'LaTeXHeaderFile', 'Tags',
                           'AKA', 'ImageFile', 'FullName', 'Goals',
                           'Notes', 'RTFFile', 'SceneContent',
                           'Outcome', 'Goal', 'Conflict']
        # Names of yWriter xml elements containing CDATA.
        # ElementTree.write omits CDATA tags, so they have to be inserted
        # afterwards.

    @property
    def filePath(self):
        return self._filePath

    @filePath.setter
    def filePath(self, filePath):
        """Accept only filenames with the correct extension. """

        if filePath.lower().endswith('.yw7'):
            self.EXTENSION = '.yw7'
            self._ENCODING = 'utf-8'
            self._filePath = filePath

        elif filePath.lower().endswith('.yw6'):
            self.EXTENSION = '.yw6'
            self._ENCODING = 'utf-8'
            self._filePath = filePath

        elif filePath.lower().endswith('.yw5'):
            self.EXTENSION = '.yw5'
            self._ENCODING = 'iso-8859-1'
            self._filePath = filePath

    def read(self):
        """Parse the yWriter xml file located at filePath, fetching the Novel attributes.
        Return a message beginning with SUCCESS or ERROR.
        """

        # Complete the list of tags requiring CDATA (if incomplete).

        try:
            with open(self._filePath, 'r', encoding=self._ENCODING) as f:
                xmlData = f.read()

        except(FileNotFoundError):
            return 'ERROR: "' + self._filePath + '" not found.'

        lines = xmlData.split('\n')

        for line in lines:
            tag = re.search('\<(.+?)\>\<\!\[CDATA', line)

            if tag is not None:

                if not (tag.group(1) in self._cdataTags):
                    self._cdataTags.append(tag.group(1))

        # Open the file again to let ElementTree parse its xml structure.

        try:
            self._tree = ET.parse(self._filePath)
            root = self._tree.getroot()

        except:
            return 'ERROR: Can not process "' + self._filePath + '".'

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

            if self.chapters[chId].title.startswith('@'):
                self.chapters[chId].suppressChapterTitle = True

            else:
                self.chapters[chId].suppressChapterTitle = False

            if chp.find('Desc') is not None:
                self.chapters[chId].desc = chp.find('Desc').text

            if chp.find('SectionStart') is not None:
                self.chapters[chId].chLevel = 1

            else:
                self.chapters[chId].chLevel = 0

            if chp.find('Type') is not None:
                self.chapters[chId].chType = int(chp.find('Type').text)

            if chp.find('Unused') is not None:
                self.chapters[chId].isUnused = True

            else:
                self.chapters[chId].isUnused = False

            for fields in chp.findall('Fields'):

                if fields.find('Field_SuppressChapterTitle') is not None:

                    if fields.find('Field_SuppressChapterTitle').text == '1':
                        self.chapters[chId].suppressChapterTitle = True

                if fields.find('Field_IsTrash') is not None:

                    if fields.find('Field_IsTrash').text == '1':
                        self.chapters[chId].isTrash = True

                    else:
                        self.chapters[chId].isTrash = False

                if fields.find('Field_SuppressChapterBreak') is not None:

                    if fields.find('Field_SuppressChapterTitle').text == '0':
                        self.chapters[chId].doNotExport = True

                    else:
                        self.chapters[chId].doNotExport = False

                else:
                    self.chapters[chId].doNotExport = False

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

            if scn.find('SceneContent') is not None:
                sceneContent = scn.find('SceneContent').text

                if sceneContent is not None:
                    self.scenes[scId].sceneContent = sceneContent

            if scn.find('Unused') is not None:
                self.scenes[scId].isUnused = True

            else:
                self.scenes[scId].isUnused = False

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
                self.scenes[scId].date = dateTime[0]
                self.scenes[scId].time = dateTime[1]

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
        """Merge attributes.
        """
        # Merge locations.

        for lcId in novel.locations:

            if not lcId in self.locations:
                self.locations[lcId] = Object()

            if novel.locations[lcId].title:
                # avoids deleting the title, if it is empty by accident
                self.locations[lcId].title = novel.locations[lcId].title

            if novel.locations[lcId].desc is not None:
                self.locations[lcId].desc = novel.locations[lcId].desc

            if novel.locations[lcId].aka is not None:
                self.locations[lcId].aka = novel.locations[lcId].aka

            if novel.locations[lcId].tags is not None:
                self.locations[lcId].tags = novel.locations[lcId].tags

        # Merge items.

        for itId in novel.items:

            if not itId in self.items:
                self.items[itId] = Object()

            if novel.items[itId].title:
                # avoids deleting the title, if it is empty by accident
                self.items[itId].title = novel.items[itId].title

            if novel.items[itId].desc is not None:
                self.items[itId].desc = novel.items[itId].desc

            if novel.items[itId].aka is not None:
                self.items[itId].aka = novel.items[itId].aka

            if novel.items[itId].tags is not None:
                self.items[itId].tags = novel.items[itId].tags

        # Merge characters.

        for crId in novel.characters:

            if not crId in self.characters:
                self.characters[crId] = Character()

            if novel.characters[crId].title:
                # avoids deleting the title, if it is empty by accident
                self.characters[crId].title = novel.characters[crId].title

            if novel.characters[crId].desc is not None:
                self.characters[crId].desc = novel.characters[crId].desc

            if novel.characters[crId].aka is not None:
                self.characters[crId].aka = novel.characters[crId].aka

            if novel.characters[crId].tags is not None:
                self.characters[crId].tags = novel.characters[crId].tags

            if novel.characters[crId].notes is not None:
                self.characters[crId].notes = novel.characters[crId].notes

            if novel.characters[crId].bio is not None:
                self.characters[crId].bio = novel.characters[crId].bio

            if novel.characters[crId].goals is not None:
                self.characters[crId].goals = novel.characters[crId].goals

            if novel.characters[crId].fullName is not None:
                self.characters[crId].fullName = novel.characters[crId].fullName

            if novel.characters[crId].isMajor is not None:
                self.characters[crId].isMajor = novel.characters[crId].isMajor

        # Merge scenes.

        for scId in novel.scenes:

            if not scId in self.scenes:
                self.scenes[scId] = Scene()

            if novel.scenes[scId].title:
                # avoids deleting the title, if it is empty by accident
                self.scenes[scId].title = novel.scenes[scId].title

            if novel.scenes[scId].desc is not None:
                self.scenes[scId].desc = novel.scenes[scId].desc

            if novel.scenes[scId].sceneContent is not None:
                self.scenes[scId].sceneContent = novel.scenes[scId].sceneContent

            if novel.scenes[scId].isUnused is not None:
                self.scenes[scId].isUnused = novel.scenes[scId].isUnused

            if novel.scenes[scId].status is not None:
                self.scenes[scId].status = novel.scenes[scId].status

            if novel.scenes[scId].sceneNotes is not None:
                self.scenes[scId].sceneNotes = novel.scenes[scId].sceneNotes

            if novel.scenes[scId].tags is not None:
                self.scenes[scId].tags = novel.scenes[scId].tags

            if novel.scenes[scId].field1 is not None:
                self.scenes[scId].field1 = novel.scenes[scId].field1

            if novel.scenes[scId].field2 is not None:
                self.scenes[scId].field2 = novel.scenes[scId].field2

            if novel.scenes[scId].field3 is not None:
                self.scenes[scId].field3 = novel.scenes[scId].field3

            if novel.scenes[scId].field4 is not None:
                self.scenes[scId].field4 = novel.scenes[scId].field4

            if novel.scenes[scId].appendToPrev is not None:
                self.scenes[scId].appendToPrev = novel.scenes[scId].appendToPrev

            if novel.scenes[scId].date is not None:
                self.scenes[scId].date = novel.scenes[scId].date

            if novel.scenes[scId].time is not None:
                self.scenes[scId].time = novel.scenes[scId].time

            if novel.scenes[scId].minute is not None:
                self.scenes[scId].minute = novel.scenes[scId].minute

            if novel.scenes[scId].hour is not None:
                self.scenes[scId].hour = novel.scenes[scId].hour

            if novel.scenes[scId].day is not None:
                self.scenes[scId].day = novel.scenes[scId].day

            if novel.scenes[scId].lastsMinutes is not None:
                self.scenes[scId].lastsMinutes = novel.scenes[scId].lastsMinutes

            if novel.scenes[scId].lastsHours is not None:
                self.scenes[scId].lastsHours = novel.scenes[scId].lastsHours

            if novel.scenes[scId].lastsDays is not None:
                self.scenes[scId].lastsDays = novel.scenes[scId].lastsDays

            if novel.scenes[scId].isReactionScene is not None:
                self.scenes[scId].isReactionScene = novel.scenes[scId].isReactionScene

            if novel.scenes[scId].isSubPlot is not None:
                self.scenes[scId].isSubPlot = novel.scenes[scId].isSubPlot

            if novel.scenes[scId].goal is not None:
                self.scenes[scId].goal = novel.scenes[scId].goal

            if novel.scenes[scId].conflict is not None:
                self.scenes[scId].conflict = novel.scenes[scId].conflict

            if novel.scenes[scId].outcome is not None:
                self.scenes[scId].outcome = novel.scenes[scId].outcome

            if novel.scenes[scId].characters is not None:
                self.scenes[scId].characters = []

                for crId in novel.scenes[scId].characters:

                    if crId in self.characters:
                        self.scenes[scId].characters.append(crId)

            if novel.scenes[scId].locations is not None:
                self.scenes[scId].locations = []

                for lcId in novel.scenes[scId].locations:

                    if lcId in self.locations:
                        self.scenes[scId].locations.append(lcId)

            if novel.scenes[scId].items is not None:
                self.scenes[scId].items = []

                for itId in novel.scenes[scId].items:

                    if itId in self.items:
                        self.scenes[scId].append(itId)

        # Merge chapters.

        scenesAssigned = []

        for chId in novel.chapters:

            if not chId in self.chapters:
                self.chapters[chId] = Chapter()

            if novel.chapters[chId].title:
                # avoids deleting the title, if it is empty by accident
                self.chapters[chId].title = novel.chapters[chId].title

            if novel.chapters[chId].desc is not None:
                self.chapters[chId].desc = novel.chapters[chId].desc

            if novel.chapters[chId].chLevel is not None:
                self.chapters[chId].chLevel = novel.chapters[chId].chLevel

            if novel.chapters[chId].chType is not None:
                self.chapters[chId].chType = novel.chapters[chId].chType

            if novel.chapters[chId].isUnused is not None:
                self.chapters[chId].isUnused = novel.chapters[chId].isUnused

            if novel.chapters[chId].suppressChapterTitle is not None:
                self.chapters[chId].suppressChapterTitle = novel.chapters[chId].suppressChapterTitle

            if novel.chapters[chId].isTrash is not None:
                self.chapters[chId].isTrash = novel.chapters[chId].isTrash

            if novel.chapters[chId].srtScenes is not None:
                self.chapters[chId].srtScenes = []

                for scId in novel.chapters[chId].srtScenes:

                    if (scId in self.scenes) and not (scId in scenesAssigned):
                        self.chapters[chId].srtScenes.append(scId)
                        scenesAssigned.append(scId)

        # Merge attributes at novel level.

        if novel.title:
            # avoids deleting the title, if it is empty by accident
            self.title = novel.title

        if novel.desc is not None:
            self.desc = novel.desc

        if novel.author is not None:
            self.author = novel.author

        if novel.fieldTitle1 is not None:
            self.fieldTitle1 = novel.fieldTitle1

        if novel.fieldTitle2 is not None:
            self.fieldTitle2 = novel.fieldTitle2

        if novel.fieldTitle3 is not None:
            self.fieldTitle3 = novel.fieldTitle3

        if novel.fieldTitle4 is not None:
            self.fieldTitle4 = novel.fieldTitle4

        if novel.srtChapters != []:
            self.srtChapters = []

            for chId in novel.srtChapters:

                if chId in self.chapters:
                    self.srtChapters.append(chId)

    def write(self):
        """Open the yWriter xml file located at filePath and 
        replace a set of attributes not being None.
        Return a message beginning with SUCCESS or ERROR.
        """

        root = self._tree.getroot()

        # Write locations to the xml element tree.

        for loc in root.iter('LOCATION'):
            lcId = loc.find('ID').text

            if lcId in self.locations:

                if self.locations[lcId].title is not None:
                    loc.find('Title').text = self.locations[lcId].title

                if self.locations[lcId].desc is not None:

                    if loc.find('Desc') is None:
                        ET.SubElement(
                            loc, 'Desc').text = self.locations[lcId].desc

                    else:
                        loc.find('Desc').text = self.locations[lcId].desc

                if self.locations[lcId].aka is not None:

                    if loc.find('AKA') is None:
                        ET.SubElement(
                            loc, 'AKA').text = self.locations[lcId].aka

                    else:
                        loc.find('AKA').text = self.locations[lcId].aka

                if self.locations[lcId].tags is not None:

                    if loc.find('Tags') is None:
                        ET.SubElement(loc, 'Tags').text = ';'.join(
                            self.locations[lcId].tags)

                    else:
                        loc.find('Tags').text = ';'.join(
                            self.locations[lcId].tags)

        # Write items to the xml element tree.

        for itm in root.iter('ITEM'):
            itId = itm.find('ID').text

            if itId in self.items:

                if self.items[itId].title is not None:
                    itm.find('Title').text = self.items[itId].title

                if self.items[itId].desc is not None:

                    if itm.find('Desc') is None:
                        ET.SubElement(itm, 'Desc').text = self.items[itId].desc

                    else:
                        itm.find('Desc').text = self.items[itId].desc

                if self.items[itId].aka is not None:

                    if itm.find('AKA') is None:
                        ET.SubElement(itm, 'AKA').text = self.items[itId].aka

                    else:
                        itm.find('AKA').text = self.items[itId].aka

                if self.items[itId].tags is not None:

                    if itm.find('Tags') is None:
                        ET.SubElement(itm, 'Tags').text = ';'.join(
                            self.items[itId].tags)

                    else:
                        itm.find('Tags').text = ';'.join(
                            self.items[itId].tags)

        # Write characters to the xml element tree.

        for crt in root.iter('CHARACTER'):
            crId = crt.find('ID').text

            if crId in self.characters:

                if self.characters[crId].title is not None:
                    crt.find('Title').text = self.characters[crId].title

                if self.characters[crId].desc is not None:

                    if crt.find('Desc') is None:
                        ET.SubElement(
                            crt, 'Desc').text = self.characters[crId].desc

                    else:
                        crt.find('Desc').text = self.characters[crId].desc

                if self.characters[crId].aka is not None:

                    if crt.find('AKA') is None:
                        ET.SubElement(
                            crt, 'AKA').text = self.characters[crId].aka

                    else:
                        crt.find('AKA').text = self.characters[crId].aka

                if self.characters[crId].tags is not None:

                    if crt.find('Tags') is None:
                        ET.SubElement(crt, 'Tags').text = ';'.join(
                            self.characters[crId].tags)

                    else:
                        crt.find('Tags').text = ';'.join(
                            self.characters[crId].tags)

                if self.characters[crId].notes is not None:

                    if crt.find('Notes') is None:
                        ET.SubElement(
                            crt, 'Notes').text = self.characters[crId].notes

                    else:
                        crt.find(
                            'Notes').text = self.characters[crId].notes

                if self.characters[crId].bio is not None:

                    if crt.find('Bio') is None:
                        ET.SubElement(
                            crt, 'Bio').text = self.characters[crId].bio

                    else:
                        crt.find('Bio').text = self.characters[crId].bio

                if self.characters[crId].goals is not None:

                    if crt.find('Goals') is None:
                        ET.SubElement(
                            crt, 'Goals').text = self.characters[crId].goals

                    else:
                        crt.find(
                            'Goals').text = self.characters[crId].goals

                if self.characters[crId].fullName is not None:

                    if crt.find('FullName') is None:
                        ET.SubElement(
                            crt, 'FullName').text = self.characters[crId].fullName

                    else:
                        crt.find(
                            'FullName').text = self.characters[crId].fullName

                majorMarker = crt.find('Major')

                if majorMarker is not None:

                    if not self.characters[crId].isMajor:
                        crt.remove(majorMarker)

                else:
                    if self.characters[crId].isMajor:
                        ET.SubElement(crt, 'Major').text = '-1'

        # Write attributes at novel level to the xml element tree.

        prj = root.find('PROJECT')
        prj.find('Title').text = self.title

        if self.desc is not None:

            if prj.find('Desc') is None:
                ET.SubElement(prj, 'Desc').text = self.desc

            else:
                prj.find('Desc').text = self.desc

        if self.author is not None:

            if prj.find('AuthorName') is None:
                ET.SubElement(prj, 'AuthorName').text = self.author

            else:
                prj.find('AuthorName').text = self.author

        prj.find('FieldTitle1').text = self.fieldTitle1
        prj.find('FieldTitle2').text = self.fieldTitle2
        prj.find('FieldTitle3').text = self.fieldTitle3
        prj.find('FieldTitle4').text = self.fieldTitle4

        # Write attributes at chapter level to the xml element tree.

        for chp in root.iter('CHAPTER'):
            chId = chp.find('ID').text

            if chId in self.chapters:
                chp.find('Title').text = self.chapters[chId].title

                if self.chapters[chId].desc is not None:

                    if chp.find('Desc') is None:
                        ET.SubElement(
                            chp, 'Desc').text = self.chapters[chId].desc

                    else:
                        chp.find('Desc').text = self.chapters[chId].desc

                levelInfo = chp.find('SectionStart')

                if levelInfo is not None:

                    if self.chapters[chId].chLevel == 0:
                        chp.remove(levelInfo)

                chp.find('Type').text = str(self.chapters[chId].chType)

                if self.chapters[chId].isUnused:

                    if chp.find('Unused') is None:
                        ET.SubElement(chp, 'Unused').text = '-1'

                elif chp.find('Unused') is not None:
                    chp.remove(chp.find('Unused'))

        # Write attributes at scene level to the xml element tree.

        for scn in root.iter('SCENE'):
            scId = scn.find('ID').text

            if scId in self.scenes:

                if self.scenes[scId].title is not None:
                    scn.find('Title').text = self.scenes[scId].title

                if self.scenes[scId].desc is not None:

                    if scn.find('Desc') is None:
                        ET.SubElement(
                            scn, 'Desc').text = self.scenes[scId].desc

                    else:
                        scn.find('Desc').text = self.scenes[scId].desc

                if self.scenes[scId]._sceneContent is not None:
                    scn.find(
                        'SceneContent').text = replace_unsafe_glyphs(self.scenes[scId]._sceneContent)
                    scn.find('WordCount').text = str(
                        self.scenes[scId].wordCount)
                    scn.find('LetterCount').text = str(
                        self.scenes[scId].letterCount)

                if self.scenes[scId].isUnused:

                    if scn.find('Unused') is None:
                        ET.SubElement(scn, 'Unused').text = '-1'

                elif scn.find('Unused') is not None:
                    scn.remove(scn.find('Unused'))

                if self.scenes[scId].status is not None:
                    scn.find('Status').text = str(self.scenes[scId].status)

                if self.scenes[scId].sceneNotes is not None:

                    if scn.find('Notes') is None:
                        ET.SubElement(
                            scn, 'Notes').text = self.scenes[scId].sceneNotes

                    else:
                        scn.find(
                            'Notes').text = self.scenes[scId].sceneNotes

                if self.scenes[scId].tags is not None:

                    if scn.find('Tags') is None:
                        ET.SubElement(scn, 'Tags').text = ';'.join(
                            self.scenes[scId].tags)

                    else:
                        scn.find('Tags').text = ';'.join(
                            self.scenes[scId].tags)

                if self.scenes[scId].field1 is not None:

                    if scn.find('Field1') is None:
                        ET.SubElement(
                            scn, 'Field1').text = self.scenes[scId].field1

                    else:
                        scn.find('Field1').text = self.scenes[scId].field1

                if self.scenes[scId].field2 is not None:

                    if scn.find('Field2') is None:
                        ET.SubElement(
                            scn, 'Field2').text = self.scenes[scId].field2

                    else:
                        scn.find('Field2').text = self.scenes[scId].field2

                if self.scenes[scId].field3 is not None:

                    if scn.find('Field3') is None:
                        ET.SubElement(
                            scn, 'Field3').text = self.scenes[scId].field3

                    else:
                        scn.find('Field3').text = self.scenes[scId].field3

                if self.scenes[scId].field4 is not None:

                    if scn.find('Field4') is None:
                        ET.SubElement(
                            scn, 'Field4').text = self.scenes[scId].field4

                    else:
                        scn.find('Field4').text = self.scenes[scId].field4

                if self.scenes[scId].appendToPrev:

                    if scn.find('AppendToPrev') is None:
                        ET.SubElement(scn, 'AppendToPrev').text = '-1'

                elif scn.find('AppendToPrev') is not None:
                    scn.remove(scn.find('AppendToPrev'))

                # Date/time information

                if (self.scenes[scId].date is not None) and (self.scenes[scId].time is not None):
                    dateTime = self.scenes[scId].date + \
                        ' ' + self.scenes[scId].time

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

                elif (self.scenes[scId].day is not None) or (self.scenes[scId].hour is not None) or (self.scenes[scId].minute is not None):

                    if scn.find('SpecificDateTime') is not None:
                        scn.remove(scn.find('SpecificDateTime'))

                    if scn.find('SpecificDateMode') is not None:
                        scn.remove(scn.find('SpecificDateMode'))

                    if self.scenes[scId].day is not None:

                        if scn.find('Day') is not None:
                            scn.find('Day').text = self.scenes[scId].day

                        else:
                            ET.SubElement(
                                scn, 'Day').text = self.scenes[scId].day

                    if self.scenes[scId].hour is not None:

                        if scn.find('Hour') is not None:
                            scn.find('Hour').text = self.scenes[scId].hour

                        else:
                            ET.SubElement(
                                scn, 'Hour').text = self.scenes[scId].hour

                    if self.scenes[scId].minute is not None:

                        if scn.find('Minute') is not None:
                            scn.find('Minute').text = self.scenes[scId].minute

                        else:
                            ET.SubElement(
                                scn, 'Minute').text = self.scenes[scId].minute

                if self.scenes[scId].lastsDays is not None:

                    if scn.find('LastsDays') is not None:
                        scn.find(
                            'LastsDays').text = self.scenes[scId].lastsDays

                    else:
                        ET.SubElement(
                            scn, 'LastsDays').text = self.scenes[scId].lastsDays

                if self.scenes[scId].lastsHours is not None:

                    if scn.find('LastsHours') is not None:
                        scn.find(
                            'LastsHours').text = self.scenes[scId].lastsHours

                    else:
                        ET.SubElement(
                            scn, 'LastsHours').text = self.scenes[scId].lastsHours

                if self.scenes[scId].lastsMinutes is not None:

                    if scn.find('LastsMinutes') is not None:
                        scn.find(
                            'LastsMinutes').text = self.scenes[scId].lastsMinutes

                    else:
                        ET.SubElement(
                            scn, 'LastsMinutes').text = self.scenes[scId].lastsMinutes

                # Plot related information

                if self.scenes[scId].isReactionScene:

                    if scn.find('ReactionScene') is None:
                        ET.SubElement(scn, 'ReactionScene').text = '-1'

                elif scn.find('ReactionScene') is not None:
                    scn.remove(scn.find('ReactionScene'))

                if self.scenes[scId].isSubPlot:

                    if scn.find('SubPlot') is None:
                        ET.SubElement(scn, 'SubPlot').text = '-1'

                elif scn.find('SubPlot') is not None:
                    scn.remove(scn.find('SubPlot'))

                if self.scenes[scId].goal is not None:

                    if scn.find('Goal') is None:
                        ET.SubElement(
                            scn, 'Goal').text = self.scenes[scId].goal

                    else:
                        scn.find('Goal').text = self.scenes[scId].goal

                if self.scenes[scId].conflict is not None:

                    if scn.find('Conflict') is None:
                        ET.SubElement(
                            scn, 'Conflict').text = self.scenes[scId].conflict

                    else:
                        scn.find(
                            'Conflict').text = self.scenes[scId].conflict

                if self.scenes[scId].outcome is not None:

                    if scn.find('Outcome') is None:
                        ET.SubElement(
                            scn, 'Outcome').text = self.scenes[scId].outcome

                    else:
                        scn.find(
                            'Outcome').text = self.scenes[scId].outcome

                if self.scenes[scId].characters is not None:
                    characters = scn.find('Characters')

                    for oldCrId in characters.findall('CharID'):
                        characters.remove(oldCrId)

                    for crId in self.scenes[scId].characters:
                        ET.SubElement(characters, 'CharID').text = crId

                if self.scenes[scId].locations is not None:
                    locations = scn.find('Locations')

                    for oldLcId in locations.findall('LocID'):
                        locations.remove(oldLcId)

                    for lcId in self.scenes[scId].locations:
                        ET.SubElement(locations, 'LocID').text = lcId

                if self.scenes[scId].items is not None:
                    items = scn.find('Items')

                    for oldItId in items.findall('ItemID'):
                        items.remove(oldItId)

                    for itId in self.scenes[scId].items:
                        ET.SubElement(items, 'ItemID').text = itId

        # Pretty print the xml tree.

        indent(root)

        # Save the xml tree in a file.

        self._tree = ET.ElementTree(root)

        try:
            self._tree.write(
                self._filePath, xml_declaration=False, encoding=self._ENCODING)

        except(PermissionError):
            return 'ERROR: "' + self._filePath + '" is write protected.'

        # Postprocess the xml file created by ElementTree.

        message = xml_postprocess(
            self._filePath, self._ENCODING, self._cdataTags)

        if message.startswith('ERROR'):
            return message

        return 'SUCCESS: project data written to "' + self._filePath + '".'

    def is_locked(self):
        """Test whether a .lock file placed by yWriter exists.
        """
        if os.path.isfile(self._filePath + '.lock'):
            return True

        else:
            return False




class YwNewFile(YwFile):
    """yWriter xml project file representation."""

    def write(self):
        """Open the yWriter xml file located at filePath and 
        replace a set of attributes not being None.
        Return a message beginning with SUCCESS or ERROR.
        """

        root = ET.Element('YWRITER7')

        # Write attributes at novel level to the xml element tree.

        prj = ET.SubElement(root, 'PROJECT')
        ET.SubElement(prj, 'Ver').text = '7'

        if self.title is not None:
            ET.SubElement(prj, 'Title').text = self.title

        if self.desc is not None:
            ET.SubElement(prj, 'Desc').text = self.desc

        if self.author is not None:
            ET.SubElement(prj, 'AuthorName').text = self.author

        if self.fieldTitle1 is not None:
            ET.SubElement(prj, 'FieldTitle1').text = self.fieldTitle1

        if self.fieldTitle2 is not None:
            ET.SubElement(prj, 'FieldTitle2').text = self.fieldTitle2

        if self.fieldTitle3 is not None:
            ET.SubElement(prj, 'FieldTitle3').text = self.fieldTitle3

        if self.fieldTitle4 is not None:
            ET.SubElement(prj, 'FieldTitle4').text = self.fieldTitle4

        # Write locations to the xml element tree.

        locations = ET.SubElement(root, 'LOCATIONS')

        for lcId in self.locations:
            loc = ET.SubElement(locations, 'LOCATION')
            ET.SubElement(loc, 'ID').text = lcId

            if self.locations[lcId].title is not None:
                ET.SubElement(loc, 'Title').text = self.locations[lcId].title

            if self.locations[lcId].desc is not None:
                ET.SubElement(loc, 'Desc').text = self.locations[lcId].desc

            if self.locations[lcId].aka is not None:
                ET.SubElement(loc, 'AKA').text = self.locations[lcId].aka

            if self.locations[lcId].tags is not None:
                ET.SubElement(loc, 'Tags').text = ';'.join(
                    self.locations[lcId].tags)

        # Write items to the xml element tree.

        items = ET.SubElement(root, 'ITEMS')

        for itId in self.items:
            itm = ET.SubElement(items, 'ITEM')
            ET.SubElement(itm, 'ID').text = itId

            if self.items[itId].title is not None:
                ET.SubElement(itm, 'Title').text = self.items[itId].title

            if self.items[itId].desc is not None:
                ET.SubElement(itm, 'Desc').text = self.items[itId].desc

            if self.items[itId].aka is not None:
                ET.SubElement(itm, 'AKA').text = self.items[itId].aka

            if self.items[itId].tags is not None:
                ET.SubElement(itm, 'Tags').text = ';'.join(
                    self.items[itId].tags)

        # Write characters to the xml element tree.

        characters = ET.SubElement(root, 'CHARACTERS')

        for crId in self.characters:
            crt = ET.SubElement(characters, 'CHARACTER')
            ET.SubElement(crt, 'ID').text = crId

            if self.characters[crId].title is not None:
                ET.SubElement(
                    crt, 'Title').text = self.characters[crId].title

            if self.characters[crId].desc is not None:
                ET.SubElement(
                    crt, 'Desc').text = self.characters[crId].desc

            if self.characters[crId].aka is not None:
                ET.SubElement(crt, 'AKA').text = self.characters[crId].aka

            if self.characters[crId].tags is not None:
                ET.SubElement(crt, 'Tags').text = ';'.join(
                    self.characters[crId].tags)

            if self.characters[crId].notes is not None:
                ET.SubElement(
                    crt, 'Notes').text = self.characters[crId].notes

            if self.characters[crId].bio is not None:
                ET.SubElement(crt, 'Bio').text = self.characters[crId].bio

            if self.characters[crId].goals is not None:
                ET.SubElement(
                    crt, 'Goals').text = self.characters[crId].goals

            if self.characters[crId].fullName is not None:
                ET.SubElement(
                    crt, 'FullName').text = self.characters[crId].fullName

            if self.characters[crId].isMajor:
                ET.SubElement(crt, 'Major').text = '-1'

        # Write attributes at scene level to the xml element tree.

        scenes = ET.SubElement(root, 'SCENES')

        for scId in self.scenes:
            scn = ET.SubElement(scenes, 'SCENE')
            ET.SubElement(scn, 'ID').text = scId

            if self.scenes[scId].title is not None:
                ET.SubElement(scn, 'Title').text = self.scenes[scId].title

            for chId in self.chapters:

                if scId in self.chapters[chId].srtScenes:
                    ET.SubElement(scn, 'BelongsToChID').text = chId
                    break

            if self.scenes[scId].desc is not None:
                ET.SubElement(scn, 'Desc').text = self.scenes[scId].desc

            if self.scenes[scId]._sceneContent is not None:
                ET.SubElement(scn,
                              'SceneContent').text = replace_unsafe_glyphs(self.scenes[scId]._sceneContent)
                ET.SubElement(scn, 'WordCount').text = str(
                    self.scenes[scId].wordCount)
                ET.SubElement(scn, 'LetterCount').text = str(
                    self.scenes[scId].letterCount)

            if self.scenes[scId].isUnused:
                ET.SubElement(scn, 'Unused').text = '-1'

            if self.scenes[scId].status is not None:
                ET.SubElement(scn, 'Status').text = str(
                    self.scenes[scId].status)

            if self.scenes[scId].sceneNotes is not None:
                ET.SubElement(scn, 'Notes').text = self.scenes[scId].sceneNotes

            if self.scenes[scId].tags is not None:
                ET.SubElement(scn, 'Tags').text = ';'.join(
                    self.scenes[scId].tags)

            if self.scenes[scId].field1 is not None:
                ET.SubElement(scn, 'Field1').text = self.scenes[scId].field1

            if self.scenes[scId].field2 is not None:
                ET.SubElement(scn, 'Field2').text = self.scenes[scId].field2

            if self.scenes[scId].field3 is not None:
                ET.SubElement(scn, 'Field3').text = self.scenes[scId].field3

            if self.scenes[scId].field4 is not None:
                ET.SubElement(scn, 'Field4').text = self.scenes[scId].field4

            if self.scenes[scId].appendToPrev:
                ET.SubElement(scn, 'AppendToPrev').text = '-1'

            # Date/time information

            if (self.scenes[scId].date is not None) and (self.scenes[scId].time is not None):
                dateTime = ' '.join(
                    self.scenes[scId].date, self.scenes[scId].time)
                ET.SubElement(scn, 'SpecificDateTime').text = dateTime
                ET.SubElement(scn, 'SpecificDateMode').text = '-1'

            elif (self.scenes[scId].day is not None) or (self.scenes[scId].hour is not None) or (self.scenes[scId].minute is not None):

                if self.scenes[scId].day is not None:
                    ET.SubElement(scn, 'Day').text = self.scenes[scId].day

                if self.scenes[scId].hour is not None:
                    ET.SubElement(scn, 'Hour').text = self.scenes[scId].hour

                if self.scenes[scId].minute is not None:
                    ET.SubElement(
                        scn, 'Minute').text = self.scenes[scId].minute

            if self.scenes[scId].lastsDays is not None:
                ET.SubElement(
                    scn, 'LastsDays').text = self.scenes[scId].lastsDays

            if self.scenes[scId].lastsHours is not None:
                ET.SubElement(
                    scn, 'LastsHours').text = self.scenes[scId].lastsHours

            if self.scenes[scId].lastsMinutes is not None:
                ET.SubElement(
                    scn, 'LastsMinutes').text = self.scenes[scId].lastsMinutes

            # Plot related information

            if self.scenes[scId].isReactionScene:
                ET.SubElement(scn, 'ReactionScene').text = '-1'

            if self.scenes[scId].isSubPlot:
                ET.SubElement(scn, 'SubPlot').text = '-1'

            if self.scenes[scId].goal is not None:
                ET.SubElement(scn, 'Goal').text = self.scenes[scId].goal

            if self.scenes[scId].conflict is not None:
                ET.SubElement(
                    scn, 'Conflict').text = self.scenes[scId].conflict

            if self.scenes[scId].outcome is not None:
                ET.SubElement(scn, 'Outcome').text = self.scenes[scId].outcome

            if self.scenes[scId].characters is not None:
                scCharacters = ET.SubElement(scn, 'Characters')

                for crId in self.scenes[scId].characters:
                    ET.SubElement(scCharacters, 'CharID').text = crId

            if self.scenes[scId].locations is not None:
                scLocations = ET.SubElement(scn, 'Locations')

                for lcId in self.scenes[scId].locations:
                    ET.SubElement(scLocations, 'LocID').text = lcId

            if self.scenes[scId].items is not None:
                scItems = ET.SubElement(scn, 'Items')

                for itId in self.scenes[scId].items:
                    ET.SubElement(scItems, 'ItemID').text = itId

        # Write attributes at chapter level to the xml element tree.

        chapters = ET.SubElement(root, 'CHAPTERS')

        sortOrder = 0

        for chId in self.srtChapters:
            sortOrder += 1
            chp = ET.SubElement(chapters, 'CHAPTER')
            ET.SubElement(chp, 'ID').text = chId
            ET.SubElement(chp, 'SortOrder').text = str(sortOrder)

            if self.chapters[chId].title is not None:
                ET.SubElement(chp, 'Title').text = self.chapters[chId].title

            if self.chapters[chId].desc is not None:
                ET.SubElement(chp, 'Desc').text = self.chapters[chId].desc

            if self.chapters[chId].chLevel == 1:
                ET.SubElement(chp, 'SectionStart').text = '-1'

            if self.chapters[chId].chType is not None:
                ET.SubElement(chp, 'Type').text = str(
                    self.chapters[chId].chType)

            if self.chapters[chId].isUnused:
                ET.SubElement(chp, 'Unused').text = '-1'

            sortSc = ET.SubElement(chp, 'Scenes')

            for scId in self.chapters[chId].srtScenes:
                ET.SubElement(sortSc, 'ScID').text = scId

            chFields = ET.SubElement(chp, 'Fields')

            if self.chapters[chId].title.startswith('@'):
                self.chapters[chId].suppressChapterTitle = True

            if self.chapters[chId].suppressChapterTitle:
                ET.SubElement(
                    chFields, 'Field_SuppressChapterTitle').text = '1'

        # Pretty print the xml tree.

        indent(root)

        # Save the xml tree in a file.

        self._tree = ET.ElementTree(root)

        try:
            self._tree.write(
                self._filePath, xml_declaration=False, encoding=self._ENCODING)

        except(PermissionError):
            return 'ERROR: "' + self._filePath + '" is write protected.'

        # Postprocess the xml file created by ElementTree.

        message = xml_postprocess(
            self._filePath, self._ENCODING, self._cdataTags)

        if message.startswith('ERROR'):
            return message

        return 'SUCCESS: project data written to "' + self._filePath + '".'

import uno
import unohelper

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


def msgbox(message, title='LibreOffice', buttons=BUTTONS_OK, type_msg='infobox'):
    """ Create message box
        type_msg: infobox, warningbox, errorbox, querybox, messbox

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


from com.sun.star.awt.MessageBoxResults import OK, YES, NO, CANCEL



class YwCnv():
    """Converter for yWriter project files.

    # Methods

    yw_to_document : str
        Arguments
            ywFile : YwFile
                an object representing the source file.
            documentFile : Novel
                a Novel subclass instance representing the target file.
        Read yWriter file, parse xml and create a document file.
        Return a message beginning with SUCCESS or ERROR.    

    document_to_yw : str
        Arguments
            documentFile : Novel
                a Novel subclass instance representing the source file.
            ywFile : YwFile
                an object representing the target file.
        Read document file, convert its content to xml, and replace yWriter file.
        Return a message beginning with SUCCESS or ERROR.

    confirm_overwrite : bool
        Arguments
            fileName : str
                Path to the file to be overwritten
        Ask for permission to overwrite the target file.
        Returns True by default.
        This method is to be overwritten by subclasses with an user interface.
    """

    def yw_to_document(self, ywFile, documentFile):
        """Read yWriter file and convert xml to a document file."""
        if ywFile.is_locked():
            return 'ERROR: yWriter seems to be open. Please close first.'

        if ywFile.filePath is None:
            return 'ERROR: "' + ywFile.filePath + '" is not an yWriter project.'

        message = ywFile.read()

        if message.startswith('ERROR'):
            return message

        if documentFile.file_exists():

            if not self.confirm_overwrite(documentFile.filePath):
                return 'Program abort by user.'

        documentFile.merge(ywFile)
        return documentFile.write()

    def document_to_yw(self, documentFile, ywFile):
        """Read document file, convert its content to xml, and replace yWriter file."""
        if ywFile.is_locked():
            return 'ERROR: yWriter seems to be open. Please close first.'

        if ywFile.filePath is None:
            return 'ERROR: "' + ywFile.filePath + '" is not an yWriter project.'

        if ywFile.file_exists() and not self.confirm_overwrite(ywFile.filePath):
            return 'Program abort by user.'

        if documentFile.filePath is None:
            return 'ERROR: Document is not of the supported type.'

        if not documentFile.file_exists():
            return 'ERROR: "' + documentFile.filePath + '" not found.'

        message = documentFile.read()

        if message.startswith('ERROR'):
            return message

        if ywFile.file_exists():
            message = ywFile.read()
            # initialize ywFile data

            if message.startswith('ERROR'):
                return message

        prjStructure = documentFile.get_structure()

        if prjStructure is not None:

            if prjStructure == '':
                return 'ERROR: Source file contains no yWriter project structure information.'

            if prjStructure != ywFile.get_structure():
                return 'ERROR: Structure mismatch - yWriter project not modified.'

        ywFile.merge(documentFile)
        return ywFile.write()

    def confirm_overwrite(self, fileName):
        """To be overwritten by subclasses with UI."""
        return True


class YwCnvUno(YwCnv):
    """Converter for yWriter project files.
    Variant with UNO UI.
    """

    def confirm_overwrite(self, filePath):
        result = msgbox('Overwrite existing file "' + filePath + '"?',
                        'WARNING', BUTTONS_YES_NO, 'warningbox')

        if result == YES:
            return True

        else:
            return False

TAILS = [HtmlProof.SUFFIX + HtmlProof.EXTENSION,
         HtmlManuscript.SUFFIX + HtmlManuscript.EXTENSION,
         HtmlSceneDesc.SUFFIX + HtmlSceneDesc.EXTENSION,
         HtmlChapterDesc.SUFFIX + HtmlChapterDesc.EXTENSION,
         HtmlPartDesc.SUFFIX + HtmlPartDesc.EXTENSION,
         HtmlCharacters.SUFFIX + HtmlCharacters.EXTENSION,
         HtmlLocations.SUFFIX + HtmlLocations.EXTENSION,
         HtmlItems.SUFFIX + HtmlItems.EXTENSION,
         CsvSceneList.SUFFIX + CsvSceneList.EXTENSION,
         CsvPlotList.SUFFIX + CsvPlotList.EXTENSION,
         CsvCharList.SUFFIX + CsvCharList.EXTENSION,
         CsvLocList.SUFFIX + CsvLocList.EXTENSION,
         CsvItemList.SUFFIX + CsvItemList.EXTENSION,
         '.html']

YW_EXTENSIONS = ['.yw7', '.yw6', '.yw5']


def delete_tempfile(filePath):
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


def run(sourcePath):
    sourcePath = unquote(sourcePath.replace('file:///', ''))

    ywPath = None

    for tail in TAILS:
        # Determine the document type

        if sourcePath.endswith(tail):

            for ywExtension in YW_EXTENSIONS:
                # Determine the yWriter project file path

                testPath = sourcePath.replace(tail, ywExtension)

                if os.path.isfile(testPath):
                    ywPath = testPath
                    break

            break

    if ywPath:

        if tail == '.html':
            return 'ERROR: yWriter project already exists.'

        elif tail == HtmlProof.SUFFIX + HtmlProof.EXTENSION:
            sourceDoc = HtmlProof(sourcePath)

        elif tail == HtmlManuscript.SUFFIX + HtmlManuscript.EXTENSION:
            sourceDoc = HtmlManuscript(sourcePath)

        elif tail == HtmlSceneDesc.SUFFIX + HtmlSceneDesc.EXTENSION:
            sourceDoc = HtmlSceneDesc(sourcePath)

        elif tail == HtmlChapterDesc.SUFFIX + HtmlChapterDesc.EXTENSION:
            sourceDoc = HtmlChapterDesc(sourcePath)

        elif tail == HtmlPartDesc.SUFFIX + HtmlPartDesc.EXTENSION:
            sourceDoc = HtmlPartDesc(sourcePath)

        elif tail == HtmlCharacters.SUFFIX + HtmlCharacters.EXTENSION:
            sourceDoc = HtmlCharacters(sourcePath)

        elif tail == HtmlLocations.SUFFIX + HtmlLocations.EXTENSION:
            sourceDoc = HtmlLocations(sourcePath)

        elif tail == HtmlItems.SUFFIX + HtmlItems.EXTENSION:
            sourceDoc = HtmlItems(sourcePath)

        elif tail == CsvSceneList.SUFFIX + CsvSceneList.EXTENSION:
            sourceDoc = CsvSceneList(sourcePath)

        elif tail == CsvPlotList.SUFFIX + CsvPlotList.EXTENSION:
            sourceDoc = CsvPlotList(sourcePath)

        elif tail == CsvCharList.SUFFIX + CsvCharList.EXTENSION:
            sourceDoc = CsvCharList(sourcePath)

        elif tail == CsvLocList.SUFFIX + CsvLocList.EXTENSION:
            sourceDoc = CsvLocList(sourcePath)

        elif tail == CsvItemList.SUFFIX + CsvItemList.EXTENSION:
            sourceDoc = CsvItemList(sourcePath)

        else:
            return 'ERROR: File format not supported.'

        ywFile = YwFile(ywPath)
        converter = YwCnvUno()
        message = converter.document_to_yw(sourceDoc, ywFile)

    elif sourcePath.endswith('.html'):
        result = read_html_file(sourcePath)

        if 'SUCCESS' in result[0]:

            if "<h3" in result[1].lower():
                sourceDoc = HtmlOutline(sourcePath)

            else:
                sourceDoc = HtmlImport(sourcePath)

        ywPath = sourcePath.replace('.html', '.yw7')
        ywFile = YwNewFile(ywPath)
        converter = YwCnvUno()
        message = converter.document_to_yw(sourceDoc, ywFile)

    else:
        message = 'ERROR: No yWriter project found.'

    if not message.startswith('ERROR'):
        delete_tempfile(sourcePath)

    return message


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

        result = run(htmlPath)

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

        result = run(csvPath)

    else:
        result = "ERROR: File type not supported."

    if result.startswith('ERROR'):
        msgType = 'errorbox'

    else:
        msgType = 'infobox'

    msgbox(result, 'Export to yWriter', type_msg=msgType)


if __name__ == '__main__':
    try:
        sourcePath = sys.argv[1]
    except:
        sourcePath = ''
    print(run(sourcePath))
