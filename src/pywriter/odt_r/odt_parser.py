"""Provide a class for parsing ODT documents.

Copyright (c) 2023 Peter Triesberger
For further information see https://github.com/peter88213/
Published under the MIT License (https://opensource.org/licenses/mit-license.php)
"""
import zipfile
from xml import sax
import xml.etree.ElementTree as ET
from pywriter.pywriter_globals import *


class OdtParser(sax.ContentHandler):
    """An ODT document parser, emulating the html.parser.HTMLParser API.
    
    Public methods:
        feed_file(filePath) -- Feed an ODT file to the parser.
    
      Methods overriding xml.sax.ContentHandler methods (not meant to be overridden by subclasses)
        characters -- Receive notification of character data.
        endElement -- Signals the end of an element in non-namespace mode.
        startElement -- Signals the start of an element in non-namespace mode.
      
      HTMLParser-like API
        handle_comment -- Stub for a comment handler to be implemented in a subclass.
        handle_data -- Stub for a data handler to be implemented in a subclass.
        handle_endtag -- Stub for an end tag handler to be implemented in a subclass.
        handle_starttag -- Stub for a start tag handler to be implemented in a subclass.  
    """

    def __init__(self):
        super().__init__()
        self._emTags = ['Emphasis']
        self._strongTags = ['Strong_20_Emphasis']
        self._blockquoteTags = ['Quotations']
        self._languageTags = {}
        self._headingTags = {}
        self._heading = None
        self._paragraph = False
        self._commentParagraphCount = None
        self._blockquote = False
        self._list = False
        self._span = []
        self._style = None

    def feed_file(self, filePath):
        """Feed an ODT file to the parser.
        
        Positional arguments:
            filePath: str -- ODT document path.
        
        First unzip the ODT file located at self.filePath, 
        and get languageCode, countryCode, title, desc, and authorName,        
        Then call the sax parser for content.xml.
        """
        namespaces = dict(
            office='urn:oasis:names:tc:opendocument:xmlns:office:1.0',
            style='urn:oasis:names:tc:opendocument:xmlns:style:1.0',
            fo='urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0',
            dc='http://purl.org/dc/elements/1.1/',
            meta='urn:oasis:names:tc:opendocument:xmlns:meta:1.0'
            )

        try:
            with zipfile.ZipFile(filePath, 'r') as odfFile:
                content = odfFile.read('content.xml')
                styles = odfFile.read('styles.xml')
                try:
                    meta = odfFile.read('meta.xml')
                except KeyError:
                    # meta.xml may be missing in outlines created with e.g. FreeMind
                    meta = None
        except:
            raise Error(f'{_("Cannot read file")}: "{norm_path(filePath)}".')

        #--- Get language and country from 'styles.xml'.
        root = ET.fromstring(styles)
        styles = root.find('office:styles', namespaces)
        for defaultStyle in styles.findall('style:default-style', namespaces):
            if defaultStyle.get(f'{{{namespaces["style"]}}}family') == 'paragraph':
                textProperties = defaultStyle.find('style:text-properties', namespaces)
                lngCode = textProperties.get(f'{{{namespaces["fo"]}}}language')
                ctrCode = textProperties.get(f'{{{namespaces["fo"]}}}country')
                self.handle_starttag('body', [('language', lngCode), ('country', ctrCode)])
                break

        #--- Get title, description, and author from 'meta.xml'.
        if meta:
            root = ET.fromstring(meta)
            meta = root.find('office:meta', namespaces)
            title = meta.find('dc:title', namespaces)
            if title is not None:
                if title.text:
                    self.handle_starttag('title', [()])
                    self.handle_data(title.text)
                    self.handle_endtag('title')
            author = meta.find('meta:initial-creator', namespaces)
            if author is not None:
                if author.text:
                    self.handle_starttag('meta', [('', 'author'), ('', author.text)])
            desc = meta.find('dc:description', namespaces)
            if desc is not None:
                if desc.text:
                    self.handle_starttag('meta', [('', 'description'), ('', desc.text)])

        #--- Parse 'content.xml'.
        sax.parseString(content, self)

    def characters(self, content):
        """Receive notification of character data.
        
        Overrides the xml.sax.ContentHandler method             
        """
        if self._commentParagraphCount is not None:
            if self._commentParagraphCount == 1:
                self._comment = f'{self._comment}{content}'
        elif self._paragraph:
            self.handle_data(content)
        elif self._heading is not None:
            self.handle_data(content)

    def endElement(self, name):
        """Signals the end of an element in non-namespace mode.
        
        Overrides the xml.sax.ContentHandler method     
        """
        if name == 'text:p':
            if self._commentParagraphCount is None:
                while self._span:
                    self.handle_endtag(self._span.pop())
                if self._blockquote:
                    self.handle_endtag('blockquote')
                    self._blockquote = False
                elif self._heading:
                    self.handle_endtag(self._heading)
                    self._heading = None
                else:
                    self.handle_endtag('p')
                self._paragraph = False
        elif name == 'text:span':
            if self._span:
                self.handle_endtag(self._span.pop())
        elif name == 'text:section':
            self.handle_endtag('div')
        elif name == 'office:annotation':
            self.handle_comment(self._comment)
            self._commentParagraphCount = None
        elif name == 'text:h':
            self.handle_endtag(self._heading)
            self._heading = None
        elif name == 'text:list-item':
            self._list = False
        elif name == 'style:style':
            self._style = None

    def startElement(self, name, attrs):
        """Signals the start of an element in non-namespace mode.
        
        Overrides the xml.sax.ContentHandler method             
        """
        xmlAttributes = {}
        for attribute in attrs.items():
            attrKey, attrValue = attribute
            xmlAttributes[attrKey] = attrValue
        style = xmlAttributes.get('text:style-name', '')
        if name == 'text:p':
            param = [()]
            if style in self._languageTags:
                param = [('lang', self._languageTags[style])]
            if self._commentParagraphCount is not None:
                self._commentParagraphCount += 1
            elif style in self._blockquoteTags:
                self.handle_starttag('blockquote', param)
                self._paragraph = True
                self._blockquote = True
            elif style.startswith('Heading'):
                self._heading = f'h{style[-1]}'
                self.handle_starttag(self._heading, [()])
            elif style in self._headingTags:
                self._heading = self._headingTags[style]
                self.handle_starttag(self._heading, [()])
            elif self._list:
                self.handle_starttag('li', [()])
                self._paragraph = True
            else:
                self.handle_starttag('p', param)
                self._paragraph = True
            if style in self._emTags:
                self._span.append('em')
                self.handle_starttag('em', [()])
            if style in self._strongTags:
                self._span.append('strong')
                self.handle_starttag('strong', [()])
        elif name == 'text:span':
            if style in self._emTags:
                self._span.append('em')
                self.handle_starttag('em', [()])
            if style in self._strongTags:
                self._span.append('strong')
                self.handle_starttag('strong', [()])
            if style in self._languageTags:
                self._span.append('lang')
                self.handle_starttag('lang', [('lang', self._languageTags[style])])
        elif name == 'text:section':
            sectionId = xmlAttributes['text:name']
            self.handle_starttag('div', [('id', sectionId)])
        elif name == 'office:annotation':
            self._commentParagraphCount = 0
            self._comment = ''
        elif name == 'text:h':
            try:
                self._heading = f'h{xmlAttributes["text:outline-level"]}'
            except:
                self._heading = f'h{style[-1]}'
            self.handle_starttag(self._heading, [()])
        elif name == 'text:list-item':
            self._list = True
        elif name == 'style:style':
            self._style = xmlAttributes.get('style:name', None)
            styleName = xmlAttributes.get('style:parent-style-name', '')
            if styleName.startswith('Heading'):
                self._headingTags[self._style] = f'h{styleName[-1]}'
            elif styleName == 'Quotations':
                self._blockquoteTags.append(self._style)
        elif name == 'style:text-properties':
            if xmlAttributes.get('fo:font-style', None) == 'italic':
                self._emTags.append(self._style)
            if xmlAttributes.get('fo:font-weight', None) == 'bold':
                self._strongTags.append(self._style)
            if xmlAttributes.get('fo:language', False):
                lngCode = xmlAttributes['fo:language']
                ctrCode = xmlAttributes['fo:country']
                if ctrCode != 'none':
                    locale = f'{lngCode}-{ctrCode}'
                else:
                    locale = lngCode
                self._languageTags[self._style] = locale
        elif name == 'text:s':
            self.handle_starttag('s', [()])

    def handle_comment(self, data):
        """Stub for a comment handler to be implemented in a subclass.
        
        Positional arguments:
            data: str -- comment text. 
        """
        pass

    def handle_data(self, data):
        """Stub for a data handler to be implemented in a subclass.

        Positional arguments:
            data: str -- text to be stored. 
        """
        pass

    def handle_endtag(self, tag):
        """Stub for an end tag handler to be implemented in a subclass.
        
        Positional arguments:
            tag: str -- name of the tag converted to lower case.
        """
        pass

    def handle_starttag(self, tag, attrs):
        """Stub for a start tag handler to be implemented in a subclass.
        
        Positional arguments:
            tag: str -- name of the tag converted to lower case.
            attrs -- list of (name, value) pairs containing the attributes found inside the tag’s <> brackets.
        """
        pass

