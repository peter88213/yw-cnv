# The yw-cnv extension for LibreOffice: Import and export yWriter 6/7 projects 

![Screenshot: Menu in LibreOffice](https://raw.githubusercontent.com/peter88213/yw-cnv/master/docs/Screenshots/lo_menu.png)

## Features

### Import from yw6/7 project 

Generate a "standard manuscript" formatted OpenDocument textfile from a yWriter 6/7 project.

* The document is placed in the same folder as the yWriter project.
* Document's filename: `<yW project name>.odt`.
* Only "used" chapters and scenes will be imported.
* Regular chapter headings are assigned the paragraph style _Heading 2_ .
* Headings of chapters beginning a new section are assigned the paragraph style _Heading 1_ .
* The first paragraph of each scene is formatted as _Text body_ , all further paragraphs as _First line indent_ . 
* Text markup: Bold and italics are supported. Other formatting such as underline, strikethrough, alignment, or colored highlighting are lost.
* yWriter's project variables and global variables are not supported.
* Scene titles are converted to navigable comments. 
* Text passages embraced in slashes and asterisks like `/* this is a comment */` are converted to author's comments.


### Proof reading

#### Load yWriter 6/7 chapters and scenes into an OpenDocument file with chapter and scene markers. 

* Please consider the following conventions:
    * Text markup: Bold and italics are supported. Other formatting such as underline, strikethrough, alignment, or colored highlighting are lost.
    * All chapters and scenes will be exported, whether "used" or "unused".
    
* Back up your yWriter project and close yWriter before.
* The proof read document is placed in the same folder as the yWriter project.
* Document's filename: `<yW project name>_proof.odt`.
* The document contains chapter `[ChID:x]` and scene `[ScID:y]` markers according to yWriter 5 standard.  __Do not touch lines containing the markers__  if you want to be able to reimport the document into yWriter.

#### Write back the proofread scenes to the yWriter 6/7 project file.

* The yWriter 6/7 project to rewrite must exist in the same folder as the document.
* If both yw6 and yw7 project files exist, yw7 is rewritten. 

### World building lists

* Character lists that can be edited in Office Calc and written back to yWriter format.
* Location lists that can be edited in Office Calc and written back to yWriter format.
* Item lists that can be edited in Office Calc and written back to yWriter format.

You may change the sort order of the rows. You may also add or remove rows. New entities must get a unique ID.

### Cross references

An OpenDocument text file containing navigable cross references:

* Scenes per character,
* scenes per location,
* scenes per item,
* scenes per tag, 
* characters per tag, 
* locations per tag, 
* items per tag.

### Create a new yw7 project 

Generate a new yWriter 7 project from a work in progress or an outline.

* The new yWriter project is placed in the same folder as the document.
* yWriter project's filename: `<document name>.yw7`.
* Existing yWriter 7 projects will not be overwritten.


#### Formatting a work in progress

A work in progress has no third level heading.

* _Heading 1_  -->  New chapter title (beginning a new section).
* _Heading 2_  -->  New chapter title.
* `* * *`  -->  Scene divider (not needed for the first scenes in a chapter).
* All other text is considered to be scene content.

#### Formatting an outline

An outline has at least one third level heading.

* _Heading 1_  -->  New chapter title (beginning a new section).
* _Heading 2_  -->  New chapter title.
* _Heading 3_  -->  New scene title.
* All other text is considered to be chapter/scene description.


## Download and install

[Download page](https://github.com/peter88213/yw-cnv/releases/latest)

* Download the `OXT` file.

* Install it using the Libreoffice extension manager.

* After installation (and Office restart) you find a new "yWriter Import/Export" submenu in the "Files" menu.

* If no additional "yWriter Import/Export" submenu shows up in the "Files" menu, please look at the "Tools" > "Extensions" menu.

## Requirements

* LibreOffice 5.4 or more recent. LibreOffice 7 is highly recommended.
* Java Runtime Environment (LibreOffice might need it for macro execution).

## Credits

[yWriter](http://spacejock.com/yWriter7.html) by Simon Haynes.

[OpenOffice Extension Compiler](https://wiki.openoffice.org/wiki/Extensions_Packager#Extension_Compiler) by Bernard Marcelly.

## License

This extension is distributed under the [MIT License](http://www.opensource.org/licenses/mit-license.php).
