# The pywlo extension for LibreOffice: Import and export yWriter 6/7 projects 

![Screenshot: Menu in LibreOffice](https://raw.githubusercontent.com/peter88213/pywlo/master/docs/Screenshots/lo_menu.png)

## Features

### yWriter5 style proof reading

#### Export an yWriter 6/7 project to an OpenDocument file with chapter and scene markers. 

* The proof read document is placed in the same folder as the yWriter project.
* Document's filename: `<yW project name>_proof.odt`

#### Write back the proofread scenes to the yWriter 6/7 project file.

* The yWriter 6/7 project to rewrite must exist in the document folder.
* If both yw6 and yw7 project files exist, yw7 is rewritten. 

### Import from yw6/7 project 

Generate a "standard manuscript" formatted OpenDocument textfile from an yWriter 6/7 project.

* The document is placed in the same folder as the yWriter project.
* Document's filename: `<yW project name>.odt`


### Export to a yw7 project 

Generate a new yWriter 6/7 project from a work in progress or an outline.

* The new yWriter project is placed in the same folder as the document.
* yWriter project's filename: `<document name>.yw7`
* Existing yWriter 7 projects will not be overwritten.


#### Formatting a work in progress or outline

* _Heading 1_  -->  New chapter title (beginning a new section)
* _Heading 2_  -->  New chapter title
* Optional: `{#`Chapter description`#}` 
* Optional: `{_`Scene description`_}`
* Text in paragraphs is considered to be scene content.
* `* * *`  -->  Scene divider
* Optional: _Heading 3_  -->  New scene

#### Outline mode

* `{outline}` at the top of the document
* _Heading 1_  -->  New chapter title (beginning a new section)
* _Heading 2_  -->  New chapter title
* _Heading 3_  -->  New scene title
* All other text is considered to be chapter/scene description.
 

The  _pywlo_  extension is distributed under the [MIT License](http://www.opensource.org/licenses/mit-license.php).
