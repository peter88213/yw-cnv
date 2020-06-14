"""Convert html or csv to yWriter format. 

Input file format: html (with visible or invisible chapter and scene tags).

Version 0.23.0

Copyright (c) 2020, peter88213
For further information see https://github.com/peter88213/PyWriter
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import sys
import os


MANUSCRIPT_SUFFIX = '_manuscript'
PARTDESC_SUFFIX = '_parts'
CHAPTERDESC_SUFFIX = '_chapters'
SCENEDESC_SUFFIX = '_scenes'
PROOF_SUFFIX = '_proof'
SCENELIST_SUFFIX = '_scenelist'
PLOTLIST_SUFFIX = '_plotlist'
CHARLIST_SUFFIX = '_charlist'
LOCLIST_SUFFIX = '_loclist'
ITEMLIST_SUFFIX = '_itemlist'

from html.parser import HTMLParser

from abc import abstractmethod


class Novel():
    """Abstract yWriter project file representation.

    This class represents a file containing a novel with additional 
    attributes and structural information (a full set or a subset
    of the information included in an yWriter project file).
    """

    _FILE_EXTENSION = ''
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

        self.filePath = filePath

    @property
    def filePath(self):
        return self._filePath

    @filePath.setter
    def filePath(self, filePath):
        """Accept only filenames with the right extension. """
        if filePath.lower().endswith(self._FILE_EXTENSION):
            self._filePath = filePath

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
        # xml: <<Fields>Field_SuppressChapterBreak> 0

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

import re


class Scene():
    """yWriter scene representation.
    # xml: <SCENES><SCENE>
    """

    # Emulate an enumeration for the scene status

    STATUS = [None, 'Outline', 'Draft', '1st Edit', '2nd Edit', 'Done']

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

        # xml: <SpecificDateMode>-1</SpecificDateMode>
        # xml: <SpecificDateTime>1900-06-01 20:38:00</SpecificDateTime>

        # xml: <Minute>
        # xml: <Hour>
        # xml: <Day>

        # xml: <LastsMinutes>
        # xml: <LastsHours>
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



def to_yw7(text):
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

    return text


def strip_markup(text):
    """Strip yWriter 6/7 raw markup. Return a plain text string."""
    try:
        text = text.replace('[i]', '')
        text = text.replace('[/i]', '')
        text = text.replace('[b]', '')
        text = text.replace('[/b]', '')

    except:
        pass

    return text


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




class HtmlProof(Novel, HTMLParser):
    """HTML file representation of an yWriter project's OfficeFile part.

    Represents a html file with visible chapter and scene tags 
    to be read and written by Open/LibreOffice Writer.
    """

    _FILE_EXTENSION = 'html'
    # overwrites Novel._FILE_EXTENSION

    def __init__(self, filePath):
        Novel.__init__(self, filePath)
        HTMLParser.__init__(self)
        self._lines = []
        self._collectText = False

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

    def read(self):
        """Read scene content from a html file  
        with visible chapter and scene tags.
        Return a message beginning with SUCCESS or ERROR.
        """
        result = read_html_file(self._filePath)

        if result[0].startswith('ERROR'):
            return (result[0])

        text = to_yw7(result[1])

        # Invoke HTML parser to write the html body as raw text
        # to self._lines.

        self.feed(text)

        # Parse the HTML body to identify chapters and scenes.

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

        return 'SUCCESS: ' + str(len(self.scenes)) + ' Scenes read from "' + self._filePath + '".'




class HtmlManuscript(Novel, HTMLParser):
    """HTML file representation of an yWriter project's manuscript part.

    Represents a html file with chapter and scene sections 
    containing scene contents to be read and written by 
    OpenOffice/LibreOffice Writer.
    """

    _FILE_EXTENSION = 'html'
    # overwrites Novel._FILE_EXTENSION

    def __init__(self, filePath):
        Novel.__init__(self, filePath)
        HTMLParser.__init__(self)
        self._lines = []
        self._scId = None
        self._chId = None
        self._isTitle = False

    def handle_starttag(self, tag, attrs):
        """Recognize the beginning ot the body section.
        Overwrites HTMLparser.handle_starttag()
        """
        if tag == 'div':

            if attrs[0][0] == 'id':

                if attrs[0][1].startswith('ChID'):
                    self._chId = re.search('[0-9]+', attrs[0][1]).group()
                    self.chapters[self._chId] = Chapter()
                    self.chapters[self._chId].srtScenes = []
                    self.srtChapters.append(self._chId)

                elif attrs[0][1].startswith('ScID'):
                    self._scId = re.search('[0-9]+', attrs[0][1]).group()
                    self.scenes[self._scId] = Scene()
                    self.chapters[self._chId].srtScenes.append(self._scId)

        elif tag == 'meta':

            if attrs[0][1].lower() == 'author':
                self.author = attrs[1][1]

            if attrs[0][1].lower() == 'description':
                self.desc = attrs[1][1]

        elif tag == "title":
            self._isTitle = True

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

        if self._isTitle:
            self.title = data
            self._isTitle = False

    def read(self):
        """Read scene content from a html file 
        with chapter and scene sections.
        Return a message beginning with SUCCESS or ERROR. 
        """
        result = read_html_file(self._filePath)

        if result[0].startswith('ERROR'):
            return (result[0])

        text = to_yw7(result[1])

        # Invoke HTML parser.

        self.feed(text)
        return 'SUCCESS: ' + str(len(self.scenes)) + ' Scenes read from "' + self._filePath + '".'

    def get_structure(self):
        """This file format has no comparable structure."""
        return None



class HtmlSceneDesc(HtmlManuscript):
    """HTML file representation of an yWriter project's scene summaries."""

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

    def read(self):
        """Read scene summaries from a html file 
        with chapter and scene sections.
        Return a message beginning with SUCCESS or ERROR. 
        """
        result = read_html_file(self._filePath)

        if result[0].startswith('ERROR'):
            return (result[0])

        text = strip_markup(to_yw7(result[1]))

        # Invoke HTML parser.

        self.feed(text)
        return 'SUCCESS: ' + str(len(self.scenes)) + ' Scenes read from "' + self._filePath + '".'



class HtmlChapterDesc(HtmlSceneDesc):
    """HTML file representation of an yWriter project's chapters summaries."""

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
from html import unescape



class HtmlImport(HtmlManuscript):
    """HTML file representation of an yWriter project's OfficeFile part.

    Represents a html file without chapter and scene tags 
    to be written by Open/LibreOffice Writer.
    """

    def read(self):
        """Parse a HTML file and insert chapter and scene sections.
        Read scene contents.
        Return a message beginning with SUCCESS or ERROR. 
        """
        _TEXT_END_TAGS = ['<div type=footer>', '/body']
        _SCENE_DIVIDER = '* * *'
        _LOW_WORDCOUNT = 10

        _SC_DESC_OPEN = '{_'
        _SC_DESC_CLOSE = '_}'
        _CH_DESC_OPEN = '{#'
        _CH_DESC_CLOSE = '#}'
        _OUTLINE_MODE = '{outline}'

        result = read_html_file(self._filePath)

        if result[0].startswith('ERROR'):
            return (result[0])

        # Insert chapter and scene markers in html text.

        lines = result[1].split('\n')
        newlines = []
        chCount = 0     # overall chapter count
        scCount = 0     # overall scene count

        lastChDone = 0
        lastScDone = 0

        outlineMode = False
        inBody = False
        contentFinished = False
        inSceneSection = False
        inSceneDescription = False
        inChapterDescription = False

        chapterTitles = {}
        sceneTitles = {}
        chapterDescs = {}
        sceneDescs = {}
        chapterLevels = {}

        tagRegEx = re.compile(r'(<!--.*?-->|<[^>]*>)')
        scDesc = ''
        chDesc = ''

        for line in lines:

            if contentFinished:
                break

            line = line.rstrip().lstrip()
            scan = line.lower()

            if _OUTLINE_MODE in scan:
                if not inBody:
                    outlineMode = True

            elif '<h1' in scan or '<h2' in scan:
                inBody = True

                if inSceneDescription or inChapterDescription:
                    return 'ERROR: Wrong description tags in Chapter #' + str(chCount)

                if inSceneSection:

                    # Close the previous scene section.

                    newlines.append('</DIV>')

                    if outlineMode:

                        # Write back scene description.

                        sceneDescs[str(scCount)] = scDesc
                        scDesc = ''

                    inSceneSection = False

                if chCount > 0:

                    # Close the previous chapter section.

                    newlines.append('</DIV>')

                    if outlineMode:

                        # Write back previous chapter description.

                        chapterDescs[str(chCount)] = chDesc
                        chDesc = ''

                chCount += 1

                if '<h1' in scan:
                    # line contains the start of a part heading
                    chapterLevels[str(chCount)] = 1

                else:
                    # line contains the start of a chapter heading
                    chapterLevels[str(chCount)] = 0

                # Get part/chapter title.

                m = re.match('.+?>(.+?)</[h,H][1,2]>', line)

                if m is not None:
                    chapterTitles[str(chCount)] = m.group(1)

                else:
                    chapterTitles[str(chCount)] = 'Chapter' + str(chCount)

                # Open the next chapter section.

                newlines.append('<DIV ID="ChID:' + str(chCount) + '">')

            elif _SCENE_DIVIDER in scan or '<h3' in scan:
                # a new scene begins

                if inSceneDescription or inChapterDescription:
                    return 'ERROR: Wrong description tags in Chapter #' + str(chCount)

                if inSceneSection:

                    # Close the previous scene section.

                    newlines.append('</DIV>')

                    if outlineMode:

                        # write back previous scene description.

                        sceneDescs[str(scCount)] = scDesc
                        scDesc = ''

                scCount += 1

                # Get scene title.

                m = re.match('.+?>(.+?)</[h,H]3>', line)

                if m is not None:
                    sceneTitles[str(scCount)] = m.group(1)

                else:
                    sceneTitles[str(scCount)] = 'Scene ' + str(scCount)

                # Open the next scene section.

                newlines.append('<DIV ID="ScID:' + str(scCount) + '">')
                inSceneSection = True

            elif inBody and '<p' in scan:

                # Process a new paragraph.

                if inChapterDescription or _CH_DESC_OPEN in scan:

                    # Add a tagged chapter description.

                    if chDesc != '':
                        chDesc += '\n'

                    chDesc += unescape(tagRegEx.sub('', line).replace(
                        _CH_DESC_OPEN, '').replace(_CH_DESC_CLOSE, ''))

                    if _CH_DESC_CLOSE in scan:
                        chapterDescs[str(chCount)] = chDesc
                        chDesc = ''
                        inChapterDescription = False

                    else:
                        inChapterDescription = True

                elif chCount > 0 and not outlineMode and not inSceneSection:
                    scCount += 1

                    # Generate scene title.

                    sceneTitles[str(scCount)] = 'Scene ' + str(scCount)

                    # Open a scene section without heading.

                    newlines.append('<DIV ID="ScID:' + str(scCount) + '">')
                    inSceneSection = True

                    if _SC_DESC_OPEN in scan:

                        scDesc += unescape(tagRegEx.sub('', line).replace(
                            _SC_DESC_OPEN, '').replace(_SC_DESC_CLOSE, ''))

                        if _SC_DESC_CLOSE in scan:
                            sceneDescs[str(scCount)] = scDesc
                            scDesc = ''

                        else:
                            inSceneDescription = True

                    elif outlineMode:

                        # Begin a new paragraph in the chapter description.

                        if chDesc != '':
                            chDesc += '\n'

                        chDesc += unescape(tagRegEx.sub('', line))

                    else:
                        newlines.append(line)

                elif inSceneDescription or _SC_DESC_OPEN in scan:

                    # Add a tagged scene description.

                    if scDesc != '':
                        scDesc += '\n'

                    scDesc += unescape(tagRegEx.sub('', line).replace(
                        _SC_DESC_OPEN, '').replace(_SC_DESC_CLOSE, ''))

                    if _SC_DESC_CLOSE in scan:
                        sceneDescs[str(scCount)] = scDesc
                        scDesc = ''
                        inSceneDescription = False

                    else:
                        inSceneDescription = True

                elif outlineMode:

                    if str(scCount) in sceneTitles:

                        # Begin a new paragraph in the scene description.

                        if scDesc != '':
                            scDesc += '\n'

                        scDesc += unescape(tagRegEx.sub('', line))

                    else:

                        # Begin a new paragraph in the chapter description.

                        if chDesc != '':
                            chDesc += '\n'

                        chDesc += unescape(tagRegEx.sub('', line))

                else:
                    newlines.append(line)

            elif inChapterDescription:
                chDesc += '\n' + unescape(tagRegEx.sub('', line).replace(
                    _CH_DESC_OPEN, '').replace(_CH_DESC_CLOSE, ''))

                if _CH_DESC_CLOSE in scan:
                    chapterDescs[str(chCount)] = chDesc
                    chDesc = ''
                    inChapterDescription = False

            elif inSceneDescription:
                scDesc += '\n' + unescape(tagRegEx.sub('', line).replace(
                    _SC_DESC_OPEN, '').replace(_SC_DESC_CLOSE, ''))

                if _SC_DESC_CLOSE in scan:
                    sceneDescs[str(scCount)] = scDesc
                    scDesc = ''
                    inSceneDescription = False

            else:
                for marker in _TEXT_END_TAGS:

                    if marker in scan:

                        # Finish content processing.

                        if outlineMode:

                            # Write back last descriptions.

                            chapterDescs[str(chCount)] = chDesc
                            sceneDescs[str(scCount)] = scDesc

                        if inSceneSection:

                            # Close the last scene section.

                            newlines.append('</DIV>')
                            inSceneSection = False

                        if chCount > 0:

                            # Close the last chapter section.

                            newlines.append('</DIV>')

                        contentFinished = True
                        break

                if not contentFinished:

                    if outlineMode:

                        if str(scCount) in sceneTitles:

                            # Add line to the scene description.

                            scDesc += ' ' + unescape(tagRegEx.sub('', line))

                        else:

                            # Add line to the chapter description.

                            chDesc += ' ' + unescape(tagRegEx.sub('', line))

                    else:
                        newlines.append(line)

        text = '\n'.join(newlines)
        text = to_yw7(text)

        # Invoke HTML parser.

        self.feed(text)

        for scId in self.scenes:
            self.scenes[scId].title = sceneTitles[scId]

            if scId in sceneDescs:
                self.scenes[scId].desc = sceneDescs[scId]

            if self.scenes[scId].wordCount < _LOW_WORDCOUNT:
                self.scenes[scId].status = 1

            else:
                self.scenes[scId].status = 2

        for chId in self.chapters:
            self.chapters[chId].title = chapterTitles[chId]
            self.chapters[chId].chLevel = chapterLevels[chId]
            self.chapters[chId].chType = 0
            self.chapters[chId].suppressChapterTitle = False

            if chId in chapterDescs:
                self.chapters[chId].desc = chapterDescs[chId]

        return 'SUCCESS: ' + str(len(self.scenes)) + ' Scenes read from "' + self._filePath + '".'

import xml.etree.ElementTree as ET




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
            self._FILE_EXTENSION = '.yw7'
            self._ENCODING = 'utf-8'
            self._filePath = filePath

        elif filePath.lower().endswith('.yw6'):
            self._FILE_EXTENSION = '.yw6'
            self._ENCODING = 'utf-8'
            self._filePath = filePath

        elif filePath.lower().endswith('.yw5'):
            self._FILE_EXTENSION = '.yw5'
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
        """Copy selected novel attributes.
        """

        # Merge locations.

        if novel.locations != {}:

            for lcId in novel.locations:

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

        if novel.items != {}:

            for itId in novel.items:

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

        if novel.characters != {}:

            for crId in novel.characters:

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

        '''Do not modify these items yet:
        
        if novel.srtChapters != []:
            self.srtChapters = novel.srtChapters
            
        '''

        # Merge attributes at chapter level.

        if novel.chapters != {}:

            for chId in novel.chapters:

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

                '''Do not modify these items yet:
                
                if novel.chapters[chId].srtScenes != []:
                    self.chapters[chId].srtScenes = novel.chapters[chId].srtScenes

                '''

        # Merge attributes at scene level.

        if novel.scenes != {}:

            for scId in novel.scenes:

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
                            self.scenes[scId].append(crId)

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




class YwNewFile(Novel):
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
            self._FILE_EXTENSION = '.yw7'
            self._ENCODING = 'utf-8'
            self._filePath = filePath

    def merge(self, novel):
        """Copy selected novel attributes.
        """

        # Merge locations.

        if novel.locations != {}:

            for lcId in novel.locations:

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

        if novel.items != {}:

            for itId in novel.items:

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

        if novel.characters != {}:

            for crId in novel.characters:

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

        self.srtChapters = novel.srtChapters

        # Merge attributes at chapter level.

        if novel.chapters != {}:

            for chId in novel.chapters:
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

                if novel.chapters[chId].srtScenes != []:
                    self.chapters[chId].srtScenes = novel.chapters[chId].srtScenes

        # Merge attributes at scene level.

        if novel.scenes != {}:

            for scId in novel.scenes:
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
                            self.scenes[scId].append(crId)

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

    def is_locked(self):
        """Test whether a .lock file placed by yWriter exists.
        """
        if os.path.isfile(self._filePath + '.lock'):
            return True

        else:
            return False


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




class CsvSceneList(Novel):
    """csv file representation of an yWriter project's scenes table. 

    Represents a csv file with a record per scene.
    * Records are separated by line breaks.
    * Data fields are delimited by the _SEPARATOR character.
    """

    _FILE_EXTENSION = 'csv'
    # overwrites Novel._FILE_EXTENSION

    _SEPARATOR = '|'     # delimits data fields within a record.
    _LINEBREAK = '\t'    # substitutes embedded line breaks.

    _SCENE_RATINGS = ['2', '3', '4', '5', '6', '7', '8', '9', '10']
    # '1' is assigned N/A (empty table cell).

    _ACTION_MARKER = 'Action'
    _REACTION_MARKER = 'Reaction'

    _TABLE_HEADER = ('Scene link'
                     + _SEPARATOR
                     + 'Scene title'
                     + _SEPARATOR
                     + 'Scene description'
                     + _SEPARATOR
                     + 'Tags'
                     + _SEPARATOR
                     + 'Scene notes'
                     + _SEPARATOR
                     + 'A/R'
                     + _SEPARATOR
                     + 'Goal'
                     + _SEPARATOR
                     + 'Conflict'
                     + _SEPARATOR
                     + 'Outcome'
                     + _SEPARATOR
                     + 'Scene'
                     + _SEPARATOR
                     + 'Words total'
                     + _SEPARATOR
                     + 'Field 1'
                     + _SEPARATOR
                     + 'Field 2'
                     + _SEPARATOR
                     + 'Field 3'
                     + _SEPARATOR
                     + 'Field 4'
                     + _SEPARATOR
                     + 'Word count'
                     + _SEPARATOR
                     + 'Letter count'
                     + _SEPARATOR
                     + 'Status'
                     + _SEPARATOR
                     + 'Characters'
                     + _SEPARATOR
                     + 'Locations'
                     + _SEPARATOR
                     + 'Items'
                     + '\n')

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

        cellsInLine = len(self._TABLE_HEADER.split(self._SEPARATOR))

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
                self.scenes[scId].desc = cell[i].replace(
                    self._LINEBREAK, '\n')
                i += 1
                self.scenes[scId].tags = cell[i].split(';')
                i += 1
                self.scenes[scId].sceneNotes = cell[i].replace(
                    self._LINEBREAK, '\n')
                i += 1

                if self._REACTION_MARKER.lower() in cell[i].lower():
                    self.scenes[scId].isReactionScene = True

                else:
                    self.scenes[scId].isReactionScene = False

                i += 1
                self.scenes[scId].goal = cell[i].replace(
                    self._LINEBREAK, ' ')
                i += 1
                self.scenes[scId].conflict = cell[i].replace(
                    self._LINEBREAK, ' ')
                i += 1
                self.scenes[scId].outcome = cell[i].replace(
                    self._LINEBREAK, ' ')
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
                charaNames = cell[i].split(';')
                self.scenes[scId].characters = []

                for charaName in charaNames:

                    for id, name in self.characters.items():

                        if name == charaName:
                            self.scenes[scId].characters.append(id)
                '''
                i += 1
                ''' Cannot write back location IDs, because self.locations is None
                locaNames = cell[i].split(';')
                self.scenes[scId].locations = []

                for locaName in locaNames:

                    for id, name in self.locations.items():

                        if name == locaName:
                            self.scenes[scId].locations.append(id)
                '''
                i += 1
                ''' Cannot write back item IDs, because self.items is None
                itemNames = cell[i].split(';')
                self.scenes[scId].items = []

                for itemName in itemNames:

                    for id, name in self.items.items():

                        if name == itemName:
                            self.scenes[scId].items.append(id)
                '''

        return 'SUCCESS: Data read from "' + self._filePath + '".'

    def merge(self, novel):
        """Copy selected novel attributes.
        """

        if novel.srtChapters != []:
            self.srtChapters = novel.srtChapters

        if novel.scenes is not None:
            self.scenes = novel.scenes

        if novel.chapters is not None:
            self.chapters = novel.chapters

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

        self.characters = novel.characters
        self.locations = novel.locations
        self.items = novel.items

    def write(self):
        """Generate a csv file containing a row per scene
        Return a message beginning with SUCCESS or ERROR.
        """
        odtPath = os.path.realpath(self.filePath).replace('\\', '/').replace(
            ' ', '%20').replace(SCENELIST_SUFFIX + '.csv', MANUSCRIPT_SUFFIX + '.odt')

        # first record: the table's column headings

        table = [self._TABLE_HEADER.replace(
            'Field 1', self.fieldTitle1).replace(
            'Field 2', self.fieldTitle2).replace(
            'Field 3', self.fieldTitle3).replace(
            'Field 4', self.fieldTitle4)]

        # Add a record for each used scene in a regular chapter

        sceneCount = 0
        wordCount = 0

        for chId in self.srtChapters:

            if self.chapters[chId].isUnused:
                continue

            if self.chapters[chId].chType != 0:
                continue

            for scId in self.chapters[chId].srtScenes:

                if self.scenes[scId].isUnused:
                    continue

                if self.scenes[scId].doNotExport:
                    continue

                if self.scenes[scId].isReactionScene:
                    pacingType = self._REACTION_MARKER

                else:
                    pacingType = self._ACTION_MARKER

                sceneCount += 1
                wordCount += self.scenes[scId].wordCount

                if self.scenes[scId].desc is None:
                    self.scenes[scId].desc = ''

                if self.scenes[scId].tags is None:
                    self.scenes[scId].tags = ['']

                if self.scenes[scId].sceneNotes is None:
                    self.scenes[scId].sceneNotes = ''

                if self.scenes[scId].isReactionScene is None:
                    self.scenes[scId].isReactionScene = False

                if self.scenes[scId].goal is None:
                    self.scenes[scId].goal = ''

                if self.scenes[scId].conflict is None:
                    self.scenes[scId].conflict = ''

                if self.scenes[scId].outcome is None:
                    self.scenes[scId].outcome = ''

                if self.scenes[scId].field1 is None:
                    self.scenes[scId].field1 = ''

                if self.scenes[scId].field2 is None:
                    self.scenes[scId].field2 = ''

                if self.scenes[scId].field3 is None:
                    self.scenes[scId].field3 = ''

                if self.scenes[scId].field4 is None:
                    self.scenes[scId].field4 = ''

                rating1 = ''
                if self.scenes[scId].field1 != '1':
                    rating1 = self.scenes[scId].field1

                rating2 = ''
                if self.scenes[scId].field2 != '1':
                    rating2 = self.scenes[scId].field2

                rating3 = ''
                if self.scenes[scId].field3 != '1':
                    rating3 = self.scenes[scId].field3

                rating4 = ''
                if self.scenes[scId].field4 != '1':
                    rating4 = self.scenes[scId].field4

                charas = ''

                if self.scenes[scId].characters is not None:

                    for crId in self.scenes[scId].characters:

                        if charas != '':
                            charas += '; '

                        charas += self.characters[crId].title

                locas = ''

                if self.scenes[scId].locations is not None:

                    for lcId in self.scenes[scId].locations:

                        if locas != '':
                            locas += '; '

                        locas += self.locations[lcId].title

                items = ''

                if self.scenes[scId].items is not None:

                    for itId in self.scenes[scId].items:

                        if items != '':
                            items += '; '

                        items += self.items[itId].title

                table.append('=HYPERLINK("file:///'
                             + odtPath + '#ScID:' + scId + '%7Cregion";"ScID:' + scId + '")'
                             + self._SEPARATOR
                             + self.scenes[scId].title
                             + self._SEPARATOR
                             + self.scenes[scId].desc.rstrip().replace('\n', self._LINEBREAK)
                             + self._SEPARATOR
                             + ';'.join(self.scenes[scId].tags)
                             + self._SEPARATOR
                             + self.scenes[scId].sceneNotes.rstrip().replace('\n', self._LINEBREAK)
                             + self._SEPARATOR
                             + pacingType
                             + self._SEPARATOR
                             + self.scenes[scId].goal
                             + self._SEPARATOR
                             + self.scenes[scId].conflict
                             + self._SEPARATOR
                             + self.scenes[scId].outcome
                             + self._SEPARATOR
                             + str(sceneCount)
                             + self._SEPARATOR
                             + str(wordCount)
                             + self._SEPARATOR
                             + rating1
                             + self._SEPARATOR
                             + rating2
                             + self._SEPARATOR
                             + rating3
                             + self._SEPARATOR
                             + rating4
                             + self._SEPARATOR
                             + str(self.scenes[scId].wordCount)
                             + self._SEPARATOR
                             + str(self.scenes[scId].letterCount)
                             + self._SEPARATOR
                             + Scene.STATUS[self.scenes[scId].status]
                             + self._SEPARATOR
                             + charas
                             + self._SEPARATOR
                             + locas
                             + self._SEPARATOR
                             + items
                             + '\n')

        try:
            with open(self._filePath, 'w', encoding='utf-8') as f:
                f.writelines(table)

        except(PermissionError):
            return 'ERROR: ' + self._filePath + '" is write protected.'

        return 'SUCCESS: "' + self._filePath + '" saved.'

    def get_structure(self):
        """This file format has no comparable structure."""
        return None




class CsvPlotList(Novel):
    """csv file representation of an yWriter project's scenes table. 

    Represents a csv file with a record per scene.
    * Records are separated by line breaks.
    * Data fields are delimited by the _SEPARATOR character.
    """

    _FILE_EXTENSION = 'csv'
    # overwrites Novel._FILE_EXTENSION

    _SEPARATOR = '|'     # delimits data fields within a record.
    _LINEBREAK = '\t'    # substitutes embedded line breaks.

    _STORYLINE_MARKER = 'story'
    # Field names containing this string (case insensitive)
    # are associated to storylines

    _SCENE_RATINGS = ['2', '3', '4', '5', '6', '7', '8', '9', '10']
    # '1' is assigned N/A (empty table cell).

    _NOT_APPLICABLE = 'N/A'
    # Scene field column header for fields not being assigned to a storyline

    _TABLE_HEADER = ('ID'
                     + _SEPARATOR
                     + 'Plot section'
                     + _SEPARATOR
                     + 'Plot event'
                     + _SEPARATOR
                     + 'Plot event title'
                     + _SEPARATOR
                     + 'Details'
                     + _SEPARATOR
                     + 'Scene'
                     + _SEPARATOR
                     + 'Words total'
                     + _SEPARATOR
                     + _NOT_APPLICABLE
                     + _SEPARATOR
                     + _NOT_APPLICABLE
                     + _SEPARATOR
                     + _NOT_APPLICABLE
                     + _SEPARATOR
                     + _NOT_APPLICABLE
                     + '\n')

    _CHAR_STATE = ['', 'N/A', 'unhappy', 'dissatisfied',
                   'vague', 'satisfied', 'happy', '', '', '', '']

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

        cellsInLine = len(self._TABLE_HEADER.split(self._SEPARATOR))

        tableHeader = lines[0].rstrip().split(self._SEPARATOR)

        for line in lines:
            cell = line.rstrip().split(self._SEPARATOR)

            if len(cell) != cellsInLine:
                return 'ERROR: Wrong cell structure.'

            if 'ChID:' in cell[0]:
                chId = re.search('ChID\:([0-9]+)', cell[0]).group(1)
                self.chapters[chId] = Chapter()
                self.chapters[chId].title = cell[1]
                self.chapters[chId].desc = cell[4].replace(
                    self._LINEBREAK, '\n')

            if 'ScID:' in cell[0]:
                scId = re.search('ScID\:([0-9]+)', cell[0]).group(1)
                self.scenes[scId] = Scene()
                self.scenes[scId].tags = cell[2].split(';')
                self.scenes[scId].title = cell[3]
                self.scenes[scId].sceneNotes = cell[4].replace(
                    self._LINEBREAK, '\n')

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

    def merge(self, novel):
        """Copy selected novel attributes.
        """

        if novel.srtChapters != []:
            self.srtChapters = novel.srtChapters

        if novel.scenes is not None:
            self.scenes = novel.scenes

        if novel.chapters is not None:
            self.chapters = novel.chapters

        if novel.fieldTitle1 is not None:
            self.fieldTitle1 = novel.fieldTitle1

        else:
            self.fieldTitle1 = self._NOT_APPLICABLE

        if novel.fieldTitle2 is not None:
            self.fieldTitle2 = novel.fieldTitle2

        else:
            self.fieldTitle2 = self._NOT_APPLICABLE

        if novel.fieldTitle3 is not None:
            self.fieldTitle3 = novel.fieldTitle3

        else:
            self.fieldTitle3 = self._NOT_APPLICABLE

        if novel.fieldTitle4 is not None:
            self.fieldTitle4 = novel.fieldTitle4

        else:
            self.fieldTitle4 = self._NOT_APPLICABLE

        self.characters = novel.characters
        self.locations = novel.locations
        self.items = novel.items

    def write(self):
        """Generate a csv file showing the novel's plot structure.
        Return a message beginning with SUCCESS or ERROR.
        """

        odtPath = os.path.realpath(self.filePath).replace('\\', '/').replace(
            ' ', '%20').replace(PLOTLIST_SUFFIX + '.csv', MANUSCRIPT_SUFFIX + '.odt')

        # first record: the table's column headings

        table = [self._TABLE_HEADER]

        # Identify storyline arcs

        charList = []

        for crId in self.characters:
            charList.append(self.characters[crId].title)

        if self.fieldTitle1 in charList or self._STORYLINE_MARKER in self.fieldTitle1.lower():
            table[0] = table[0].replace(self._NOT_APPLICABLE, self.fieldTitle1)
            arc1 = True

        else:
            arc1 = False

        if self.fieldTitle2 in charList or self._STORYLINE_MARKER in self.fieldTitle2.lower():
            table[0] = table[0].replace(self._NOT_APPLICABLE, self.fieldTitle2)
            arc2 = True

        else:
            arc2 = False

        if self.fieldTitle3 in charList or self._STORYLINE_MARKER in self.fieldTitle3.lower():
            table[0] = table[0].replace(self._NOT_APPLICABLE, self.fieldTitle3)
            arc3 = True

        else:
            arc3 = False

        if self.fieldTitle4 in charList or self._STORYLINE_MARKER in self.fieldTitle4.lower():
            table[0] = table[0].replace(self._NOT_APPLICABLE, self.fieldTitle4)
            arc4 = True

        else:
            arc4 = False

        # Add a record for each used scene in a regular chapter
        # and for each chapter marked "Other".

        sceneCount = 0
        wordCount = 0

        for chId in self.srtChapters:

            if self.chapters[chId].isUnused:
                continue

            if self.chapters[chId].chType == 1:
                # Chapter marked "Other" precedes and describes a Plot section.
                # Put chapter description to "details".

                if self.chapters[chId].desc is None:
                    self.chapters[chId].desc = ''

                table.append('ChID:' + chId
                             + self._SEPARATOR
                             + self.chapters[chId].title
                             + self._SEPARATOR
                             + self._SEPARATOR
                             + self._SEPARATOR
                             + self.chapters[chId].desc.rstrip().replace('\n', self._LINEBREAK)
                             + self._SEPARATOR
                             + self._SEPARATOR
                             + self._SEPARATOR
                             + self._SEPARATOR
                             + self._SEPARATOR
                             + self._SEPARATOR
                             + '\n')

            else:
                for scId in self.chapters[chId].srtScenes:

                    if self.scenes[scId].isUnused:
                        continue

                    if self.scenes[scId].doNotExport:
                        continue

                    sceneCount += 1
                    wordCount += self.scenes[scId].wordCount

                    # If the scene contains plot information:
                    # a tag marks the plot event (e.g. inciting event, plot point, climax).
                    # Put scene note text to "details".
                    # Transfer scene ratings > 1 to storyline arc
                    # states.

                    if self.scenes[scId].sceneNotes is None:
                        self.scenes[scId].sceneNotes = ''

                    if self.scenes[scId].tags is None:
                        self.scenes[scId].tags = ['']

                    arcState1 = ''
                    if arc1 and self.scenes[scId].field1 != '1':
                        arcState1 = self.scenes[scId].field1

                    arcState2 = ''
                    if arc2 and self.scenes[scId].field2 != '1':
                        arcState2 = self.scenes[scId].field2

                    arcState3 = ''
                    if arc3 and self.scenes[scId].field3 != '1':
                        arcState3 = self.scenes[scId].field3

                    arcState4 = ''
                    if arc4 and self.scenes[scId].field4 != '1':
                        arcState4 = self.scenes[scId].field4

                    table.append('=HYPERLINK("file:///'
                                 + odtPath + '#ScID:' + scId + '%7Cregion";"ScID:' + scId + '")'
                                 + self._SEPARATOR
                                 + self._SEPARATOR
                                 + ';'.join(self.scenes[scId].tags)
                                 + self._SEPARATOR
                                 + self.scenes[scId].title
                                 + self._SEPARATOR
                                 + self.scenes[scId].sceneNotes.rstrip().replace('\n', self._LINEBREAK)
                                 + self._SEPARATOR
                                 + str(sceneCount)
                                 + self._SEPARATOR
                                 + str(wordCount)
                                 + self._SEPARATOR
                                 + arcState1
                                 + self._SEPARATOR
                                 + arcState2
                                 + self._SEPARATOR
                                 + arcState3
                                 + self._SEPARATOR
                                 + arcState4
                                 + '\n')

        try:
            with open(self._filePath, 'w', encoding='utf-8') as f:
                f.writelines(table)

        except(PermissionError):
            return 'ERROR: ' + self._filePath + '" is write protected.'

        return 'SUCCESS: "' + self._filePath + '" saved.'

    def get_structure(self):
        return None




class CsvCharList(Novel):
    """csv file representation of an yWriter project's characters table. 

    Represents a csv file with a record per character.
    * Records are separated by line breaks.
    * Data fields are delimited by the _SEPARATOR character.
    """

    _FILE_EXTENSION = 'csv'
    # overwrites Novel._FILE_EXTENSION

    _SEPARATOR = '|'     # delimits data fields within a record.
    _LINEBREAK = '\t'    # substitutes embedded line breaks.

    _TABLE_HEADER = ('ID'
                     + _SEPARATOR
                     + 'Name'
                     + _SEPARATOR
                     + 'Full name'
                     + _SEPARATOR
                     + 'Aka'
                     + _SEPARATOR
                     + 'Description'
                     + _SEPARATOR
                     + 'Bio'
                     + _SEPARATOR
                     + 'Goals'
                     + _SEPARATOR
                     + 'Importance'
                     + _SEPARATOR
                     + 'Tags'
                     + _SEPARATOR
                     + 'Notes'
                     + '\n')

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

        if lines[0] != self._TABLE_HEADER:
            return 'ERROR: Wrong lines content.'

        cellsInLine = len(self._TABLE_HEADER.split(self._SEPARATOR))

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
                self.characters[crId].desc = cell[4].replace(
                    self._LINEBREAK, '\n')
                self.characters[crId].bio = cell[5]
                self.characters[crId].goals = cell[6]

                if 'Major' in cell[7]:
                    self.characters[crId].isMajor = True

                else:
                    self.characters[crId].isMajor = False

                self.characters[crId].tags = cell[8].split(';')
                self.characters[crId].notes = cell[9].replace(
                    self._LINEBREAK, '\n')

        return 'SUCCESS: Data read from "' + self._filePath + '".'

    def merge(self, novel):
        """Copy selected novel attributes.
        """

        if novel.characters is not None:
            self.characters = novel.characters

    def write(self):
        """Generate a csv file containing per character:
        - character ID, 
        - character name,
        - character full name,
        - character alternative name, 
        - character description, 
        - character bio,
        - character goals,
        - character importance,
        - character tags,
        - character notes.
        Return a message beginning with SUCCESS or ERROR.
        """

        def importance(isMajor):

            if isMajor:
                return 'Major'

            else:
                return 'Minor'

        # first record: the table's column headings

        table = [self._TABLE_HEADER]

        # Add a record for each character

        for crId in self.characters:

            if self.characters[crId].fullName is None:
                self.characters[crId].fullName = ''

            if self.characters[crId].aka is None:
                self.characters[crId].aka = ''

            if self.characters[crId].desc is None:
                self.characters[crId].desc = ''

            if self.characters[crId].bio is None:
                self.characters[crId].bio = ''

            if self.characters[crId].goals is None:
                self.characters[crId].goals = ''

            if self.characters[crId].isMajor is None:
                self.characters[crId].isMajor = False

            if self.characters[crId].tags is None:
                self.characters[crId].tags = ['']

            if self.characters[crId].notes is None:
                self.characters[crId].notes = ''

            table.append('CrID:' + str(crId)
                         + self._SEPARATOR
                         + self.characters[crId].title
                         + self._SEPARATOR
                         + self.characters[crId].fullName
                         + self._SEPARATOR
                         + self.characters[crId].aka
                         + self._SEPARATOR
                         + self.characters[crId].desc.rstrip().replace('\n', self._LINEBREAK)
                         + self._SEPARATOR
                         + self.characters[crId].bio
                         + self._SEPARATOR
                         + self.characters[crId].goals
                         + self._SEPARATOR
                         + importance(self.characters[crId].isMajor)
                         + self._SEPARATOR
                         + ';'.join(self.characters[crId].tags)
                         + self._SEPARATOR
                         + self.characters[crId].notes.rstrip().replace('\n', self._LINEBREAK)
                         + '\n')

        try:
            with open(self._filePath, 'w', encoding='utf-8') as f:
                f.writelines(table)

        except(PermissionError):
            return 'ERROR: ' + self._filePath + '" is write protected.'

        return 'SUCCESS: "' + self._filePath + '" saved.'

    def get_structure(self):
        """This file format has no comparable structure."""
        return None




class CsvLocList(Novel):
    """csv file representation of an yWriter project's locations table. 

    Represents a csv file with a record per location.
    * Records are separated by line breaks.
    * Data fields are delimited by the _SEPARATOR location.
    """

    _FILE_EXTENSION = 'csv'
    # overwrites Novel._FILE_EXTENSION

    _SEPARATOR = '|'     # delimits data fields within a record.
    _LINEBREAK = '\t'    # substitutes embedded line breaks.

    _TABLE_HEADER = ('ID'
                     + _SEPARATOR
                     + 'Name'
                     + _SEPARATOR
                     + 'Description'
                     + _SEPARATOR
                     + 'Aka'
                     + _SEPARATOR
                     + 'Tags'
                     + '\n')

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

        if lines[0] != self._TABLE_HEADER:
            return 'ERROR: Wrong lines content.'

        cellsInLine = len(self._TABLE_HEADER.split(self._SEPARATOR))

        for line in lines:
            cell = line.rstrip().split(self._SEPARATOR)

            if len(cell) != cellsInLine:
                return 'ERROR: Wrong cell structure.'

            if 'LcID:' in cell[0]:
                lcId = re.search('LcID\:([0-9]+)', cell[0]).group(1)
                self.locations[lcId] = Object()
                self.locations[lcId].title = cell[1]
                self.locations[lcId].desc = cell[2].replace(
                    self._LINEBREAK, '\n')
                self.locations[lcId].aka = cell[3]
                self.locations[lcId].tags = cell[4].split(';')

        return 'SUCCESS: Data read from "' + self._filePath + '".'

    def merge(self, novel):
        """Copy selected novel attributes.
        """

        if novel.locations is not None:
            self.locations = novel.locations

    def write(self):
        """Generate a csv file containing per location:
        - location ID, 
        - location title,
        - location description, 
        - location alternative name, 
        - location tags.
        Return a message beginning with SUCCESS or ERROR.
        """

        # first record: the table's column headings

        table = [self._TABLE_HEADER]

        # Add a record for each location

        for lcId in self.locations:

            if self.locations[lcId].desc is None:
                self.locations[lcId].desc = ''

            if self.locations[lcId].aka is None:
                self.locations[lcId].aka = ''

            if self.locations[lcId].tags is None:
                self.locations[lcId].tags = ['']

            table.append('LcID:' + str(lcId)
                         + self._SEPARATOR
                         + self.locations[lcId].title
                         + self._SEPARATOR
                         + self.locations[lcId].desc.rstrip().replace('\n', self._LINEBREAK)
                         + self._SEPARATOR
                         + self.locations[lcId].aka
                         + self._SEPARATOR
                         + ';'.join(self.locations[lcId].tags)
                         + '\n')

        try:
            with open(self._filePath, 'w', encoding='utf-8') as f:
                f.writelines(table)

        except(PermissionError):
            return 'ERROR: ' + self._filePath + '" is write protected.'

        return 'SUCCESS: "' + self._filePath + '" saved.'

    def get_structure(self):
        """This file format has no comparable structure."""
        return None




class CsvItemList(Novel):
    """csv file representation of an yWriter project's items table. 

    Represents a csv file with a record per item.
    * Records are separated by line breaks.
    * Data fields are delimited by the _SEPARATOR item.
    """

    _FILE_EXTENSION = 'csv'
    # overwrites Novel._FILE_EXTENSION

    _SEPARATOR = '|'     # delimits data fields within a record.
    _LINEBREAK = '\t'    # substitutes embedded line breaks.

    _TABLE_HEADER = ('ID'
                     + _SEPARATOR
                     + 'Name'
                     + _SEPARATOR
                     + 'Description'
                     + _SEPARATOR
                     + 'Aka'
                     + _SEPARATOR
                     + 'Tags'
                     + '\n')

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

        if lines[0] != self._TABLE_HEADER:
            return 'ERROR: Wrong lines content.'

        cellsInLine = len(self._TABLE_HEADER.split(self._SEPARATOR))

        for line in lines:
            cell = line.rstrip().split(self._SEPARATOR)

            if len(cell) != cellsInLine:
                return 'ERROR: Wrong cell structure.'

            if 'ItID:' in cell[0]:
                itId = re.search('ItID\:([0-9]+)', cell[0]).group(1)
                self.items[itId] = Object()
                self.items[itId].title = cell[1]
                self.items[itId].desc = cell[2].replace(
                    self._LINEBREAK, '\n')
                self.items[itId].aka = cell[3]
                self.items[itId].tags = cell[4].split(';')

        return 'SUCCESS: Data read from "' + self._filePath + '".'

    def merge(self, novel):
        """Copy selected novel attributes.
        """

        if novel.items is not None:
            self.items = novel.items

    def write(self):
        """Generate a csv file containing per item:
        - item ID, 
        - item title,
        - item description, 
        - item alternative name, 
        - item tags.
        Return a message beginning with SUCCESS or ERROR.
        """

        # first record: the table's column headings

        table = [self._TABLE_HEADER]

        # Add a record for each item

        for itId in self.items:

            if self.items[itId].desc is None:
                self.items[itId].desc = ''

            if self.items[itId].aka is None:
                self.items[itId].aka = ''

            if self.items[itId].tags is None:
                self.items[itId].tags = ['']

            table.append('ItID:' + str(itId)
                         + self._SEPARATOR
                         + self.items[itId].title
                         + self._SEPARATOR
                         + self.items[itId].desc.rstrip().replace('\n',
                                                                  self._LINEBREAK)
                         + self._SEPARATOR
                         + self.items[itId].aka
                         + self._SEPARATOR
                         + ';'.join(self.items[itId].tags)
                         + '\n')

        try:
            with open(self._filePath, 'w', encoding='utf-8') as f:
                f.writelines(table)

        except(PermissionError):
            return 'ERROR: ' + self._filePath + '" is write protected.'

        return 'SUCCESS: "' + self._filePath + '" saved.'

    def get_structure(self):
        """This file format has no comparable structure."""
        return None

import uno
import unohelper

from msgbox import MsgBox


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


def msgbox(message):
    myBox = MsgBox(XSCRIPTCONTEXT.getComponentContext())
    myBox.addButton('OK')
    myBox.renderFromBoxSize(200)
    myBox.numberOflines = 3
    myBox.show(message, 0, 'PyWriter')



TAILS = [PROOF_SUFFIX + '.html', '.html', MANUSCRIPT_SUFFIX + '.html', SCENEDESC_SUFFIX + '.html',
         CHAPTERDESC_SUFFIX + '.html', PARTDESC_SUFFIX +
         '.html', SCENELIST_SUFFIX + '.csv',
         PLOTLIST_SUFFIX + '.csv', CHARLIST_SUFFIX + '.csv', LOCLIST_SUFFIX + '.csv', ITEMLIST_SUFFIX + '.csv']

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
    sourcePath = sourcePath.replace('file:///', '').replace('%20', ' ')

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

        elif tail == PROOF_SUFFIX + '.html':
            sourceDoc = HtmlProof(sourcePath)

        elif tail == MANUSCRIPT_SUFFIX + '.html':
            sourceDoc = HtmlManuscript(sourcePath)

        elif tail == SCENEDESC_SUFFIX + '.html':
            sourceDoc = HtmlSceneDesc(sourcePath)

        elif tail == CHAPTERDESC_SUFFIX + '.html':
            sourceDoc = HtmlChapterDesc(sourcePath)

        elif tail == PARTDESC_SUFFIX + '.html':
            sourceDoc = HtmlChapterDesc(sourcePath)

        elif tail == SCENELIST_SUFFIX + '.csv':
            sourceDoc = CsvSceneList(sourcePath)

        elif tail == PLOTLIST_SUFFIX + '.csv':
            sourceDoc = CsvPlotList(sourcePath)

        elif tail == CHARLIST_SUFFIX + '.csv':
            sourceDoc = CsvCharList(sourcePath)

        elif tail == LOCLIST_SUFFIX + '.csv':
            sourceDoc = CsvLocList(sourcePath)

        elif tail == ITEMLIST_SUFFIX + '.csv':
            sourceDoc = CsvItemList(sourcePath)

        else:
            return 'ERROR: File format not supported.'

        ywFile = YwFile(ywPath)
        converter = YwCnv()
        message = converter.document_to_yw(sourceDoc, ywFile)

    elif sourcePath.endswith('.html'):
        sourceDoc = HtmlImport(sourcePath)
        ywPath = sourcePath.replace('.html', '.yw7')
        ywFile = YwNewFile(ywPath)
        converter = YwCnv()
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

    documentPath = documentPath.lower()

    if documentPath.endswith('.odt') or documentPath.endswith('.html'):
        odtPath = documentPath.replace('.html', '.odt')
        htmlPath = documentPath.replace('.odt', '.html')

        # Save document in HTML format

        from com.sun.star.beans import PropertyValue
        args1 = []
        args1.append(PropertyValue())
        args1.append(PropertyValue())
        # dim args1(1) as new com.sun.star.beans.PropertyValue

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

    msgbox(result)


if __name__ == '__main__':
    try:
        sourcePath = sys.argv[1]
    except:
        sourcePath = ''
    print(run(sourcePath))
