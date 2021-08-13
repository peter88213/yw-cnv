[Project home page](index) > Changelog

------------------------------------------------------------------------

## Changelog

### v1.4.7 Bugfix release

Fix a regression from v1.4.6 causing a crash if a scene has an 
hour, but no minute set.

Based on PyWriter v3.16.4

### v1.4.6 Optional update

- Major refactoring of the templase based export.

Based on PyWriter v3.12.5

### v1.4.5 Fix links in the help files

Based on PyWriter v3.8.2

### v1.4.4 Optional update

- Major refactoring of the yw7 file processing.

Based on PyWriter v3.8.0

### v1.4.3 Linux compatible

Make the Python script Linux compatible.

Based on PyWriter 3.6.6

### v1.4.2 Bugfix update

- Fix the "proofread tags" paragraph styles.
- Make the html helpfiles Linux compatible.

Based on PyWriter v3.6.6

### v1.4.1 Add a "Format" menu entry

- *Replace scene dividers with blank lines* is added to the **Format** menu.

Based on PyWriter v3.6.5

### v1.4.0 New document formatting, general program improvement

It is highly advised to update to this version.

- Imported *Scenes/chapters* and *manuscript* documents now have three-line 
"* * *" scene dividers instead of single blank lines.
- Scene titles as comments at the beginning of the scenes in the *manuscript* 
and *work in progress* documents are written back to yWriter.
- Chapter titles in the *manuscript*, *scene descriptions*, *part descriptions*,
and *chapter descriptions* documents are written back to yWriter.
- Generate temporary files in the user's temp folder.

Based on PyWriter v3.6.5

### v1.2.2 Optional update

- Modify plot list titles.

Based on PyWriter v3.4.2

### v1.2.1 Bugfix update

- Cope better with the chapter/scene type overdetermination since yWriter 7.0.7.2.

Based on PyWriter v3.4.1

### v1.2.0 Drop yWriter 6 support

Support only yWriter 7 projects for better maintainability and speed.

Based on PyWriter v3.4.0

### v1.0.2 Optional update

- Refactor for faster execution.

Based on Pywriter v3.2.1

### v1.0.1 Bugfix release

- Fix a link in the help script.

Based on Pywriter v3.2.0

### v1.0.0 Official release

- Update the underlying class library with changed API for better maintainability.
- Revise the GUI.

Based on PyWriter v3.0.0

### v0.40.3 Bugfix release

Fix a regression causing wrong file type detection. 

Based on PyWriter v2.14.3

### v0.40.2 Service release

- Generate compressed ODF documents.
- Add a project website link to the help files.

Based on PyWriter v2.14.2

### v0.40.1 Improve the processing of comma-separated lists

- Fix incorrectly defined tags during yWriter import.
- Protect the processing of comma-separated lists against wrongly set
  blanks.
- Update HTML help and documentation.

### v0.40.0 Export cross references

Generate an OpenDocument text file containing navigable cross
references:

- Scenes per character,
- scenes per location,
- scenes per item,
- scenes per tag,
- characters per tag,
- locations per tag,
- items per tag.

### v0.38.0 Add advanced features to the "Files" menu

The "advanced features" are meant to be used by experienced
LibreOffice users only. If you aren't familiar with Calc and the
concept of sections in Writer, please do not use the advanced features.
There is a risk of damaging the project when writing back if you don't
respect the section boundaries in the odt text documents, or if you mess
up the IDs in the ods tables.

Export files with invisible markers (to be written back after editing):

- Manuscript with invisible chapter and scene sections.
- Very brief synopsis (Part descriptions): Titles and descriptions of
  chapters "beginning a new section".
- Brief synopsis (Chapter descriptions): Chapter titles and chapter
  descriptions.
- Full synopsis (Scene descriptions): Chapter titles and scene
  descriptions.
- Character sheet: Character descriptions, bio and goals.
- Location sheet: Location descriptions.
- Item sheet: Item descriptions.
- Scene list: Spreadsheet with all scene metadata.
- Plot list: Spreadsheet with scene metadata following conventions
  described in the help text.

Based on PyWriter v2.12.3

### v0.37.4 Support ods spreadsheets

Change scene/plot list import (advanced features) from csv to ods file
format.

Based on PyWriter v2.12.3

### v0.37.0 Import to ods spreadsheets

Change character/location/item list import from csv to ods file format.

Based on PyWriter v2.11.0

### v0.36.0 Import/export csv lists

Import/export character/location/item lists to/from Calc spreadsheets.
Rows may be added, deleted and re-ordered.

Based on PyWriter v2.10.0

### v0.35.0 Underline and strikethrough no longer supported

That is because a real support would require considering nesting and
multi paragraph formatting. That would make everything too complicated,
considering that such formatting is not common in a fictional prose
text.

Based on PyWriter v2.9.0

### v0.34.0

Delete the temporary file unconditionally after execution. Process all
yWriter formatting tags.

- Convert underline and strikethrough. 
- Discard alignment. 
- Discard highlighting.

Based on PyWriter v2.8.0

### v0.33.1

- Refactor and update docstrings.
- Work around a yWriter 7.1.1.2 bug.

Based on PyWriter v2.7.2

### v0.33.0

- Update UI application context.
- Set a blank line as scene divider template.

Based on PyWriter v2.7.0

### v0.32.3

Work around a bug found in yWriter 7.1.1.2 assigning invalid viewpoint
characters to scenes created by splitting.

### v0.32.2

- Refactor the code for better maintainability.

Based on PyWriter v2.6.1

### v0.31.1

- Add strict project structure check.
- Improve screen output.
- Do not indent a chapter's first scene even if marked "append to
  previous".
- Can now write complete yw5 Files.
- Convert work in progress that contains empty chapter titles.
- Fix location descriptions export.
- Refactor the code for better maintainability.

Based on PyWriter v2.5.1

### v0.30.0

Adapt to modified yw7 file format (yWriter 7.0.7.2+): 

- "Info" chapters are replaced by "Notes" chapters (always unused). 
- New "Todo" chapter type (always unused). 
- Distinguish between "Notes scene", "Todo scene" and "Unused scene". 
- Chapter/scene tag colors in "proofread" export correspond to those of the 
  yWriter chapter list.

Bugfix: 

- Suppress chapter title if required.

Based on PyWriter v2.2.0

### v0.29.6

ODT export: Begin appended scenes with first-line-indent style. Based on
PyWriter v2.1.4

### v0.29.5

Adapt to yWriter 7.0.4.9 Beta: Don't replace dashes any longer by
"safe" double hyphens when writing yw7 (PyWriter library v2.1.3).

### v0.29.4

Rewrite large parts of the code (PyWriter library v2.1.0). Support
author's comments: Text `/* commented out */` in yWriter scenes is
exported as comment and vice versa.

### v0.28.2

Fix a bug making the output filename lowercase.
