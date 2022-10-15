[Project home page](index) > Changelog

------------------------------------------------------------------------

## Changelog

### v1.30.1

- Store the document language and country codes as project variables, thus making them accessible in yWriter.
- Add a chapter about document language handling to the help text.

Based on PyWriter v7.12.2

### v1.30.0

Introduce a notation for assigning text passages to another language/country. This is mainly for spell checking in Office Writer.

Based on PyWriter v7.11.4

### v1.29.5 Improvements

- When exporting a work in progress to yWriter, process lists and quotation paragraphs.
- When importing from yWriter, paragraphs that start with "> " are formatted as "Quotations".

Based on PyWriter v7.10.3

### v1.29.4 

- Change the German wording: Szene --> Abschnitt.
- Process the document's language and country (if any).
- Support "no document language" settings.

Based on PyWriter v7.10.2

### v1.28.0

- New document types for "ToDo" chapters import and export.
- Improved spreadsheet export.
- Improved ODT export.
- Add internationalization according to GNU coding standards.
- Provide a full German localization.
- Consider project names containing a reserved suffix.
- Fix a bug where scene titles are not imported from work in progress as specified.

Based on PyWriter v7.4.8

### v1.26.3 Minor update

- Modify scene split warning marker for descriptions.

Based on PyWriter v5.18.1

### v1.26.2 Optional update

Show a warning message if new scenes are created during export to yWriter.

Based on PyWriter v5.18.0

### v1.26.1 Bugfix release

- Import from yWriter: Fix quotations at scene start.
- Fix and refactor inline code removal.

Based on PyWriter v5.16.1

### v1.26.0 Consider quotations

When importing chapters and scenes from yWriter, set style of paragraphs 
that start with "> " to "Quotations".

Based on PyWriter v5.16.0

### v1.24.0 Consider inline raw code

When exporting chapters and scenes to odt,
- ignore scenes beginning with `<HTML>` or `<TEX>`,
- remove inline raw code.

Based on PyWriter v5.14.0

### v1.22.2 Improved word counting

- Fix word counting considering ellipses.

Based on PyWriter v5.12.4

### v1.22.1 Improved word counting

- Fix word counting considering comments, hyphens, and dashes.

Based on PyWriter v5.12.3

### v1.22.0 Optional update

When splitting scenes, process title and description of new
chapters/scenes, separated by "|".

Based on PyWriter v5.12.0

### v1.20.1 Optional update

When generating chapters from an outline or work in progress, add 
chapter type as in yWriter version 7.0.7.2+

Based on PyWriter v5.10.2

### v1.20.0 Support quotations and lists

- The *Indent paragraphs that start with '> '* macro now assigns the
  *Quotations* paragraph style.
- When exporting to yWriter, *Quotations* style paragraphs will
automatically marked with "> " at the beginning.
- When exporting to yWriter, Lists will
automatically marked with "- " list strokes.

Based on PyWriter v5.4.2

### v1.18.1 Bugfix release

Fix a bug in the formatting macros where characters are deleted if no list strokes are found.

Based on PyWriter v5.4.0

### v1.18.0

- Add a macro that replaces list strokes with bullets.

Based on PyWriter v5.4.0

### v1.16.1

- Add "bullets" when converting lists to yWriter format.

Based on PyWriter v5.4.0

### v1.16.0

- New document types for "Notes" chapters import and export.
- Drop support for the "Plot List" document types.
- Change the way how scenes are split: Scene divider is "###". 
  Optionally, the scene title may be appended.

Based on PyWriter v5.2.0

### v1.14.2 Optional update

- Clean up ODF templates to make the code and generated documents more compact.

Based on PyWriter v5.0.3

### v1.14.1 Optional update

- Refactor and reformat the code.

Based on PyWriter v5.0.2

### v1.14.0

- Fix a bug where "To do" chapters cause an exception.
- Rework the user interface. 
- Refactor the code.

Based on PyWriter v5.0.0

### v1.12.3 Bugfix release

Avoid exception on trying to re-import an empty document.

Based on PyWriter v4.0.4

### v1.12.2 Bugfix release

- Update bugfixed library

Based on PyWriter v4.0.4

### v1.12.1 Bugfix release

- Make sure that a brief synopsis cannot be exported to yWriter.
- Fix a regression where writing back to yWriter throws an exception in v1.12.0.

Based on PyWriter v4.0.3

### v1.12.0 New entry in the "Format" menu

- Add a macro that indents paragraphs that begin with "> ".
- Move the "Export to yWriter" entry to the top of the submenu.

Based on pyWriter v3.32.2

### v1.10.1 Update text formatting

When creating odt files, make paragraphs after blank lines "Text body" without indent.

Based on PyWriter v3.32.2

### v1.10.0 Expand the document structure by splitting scenes

- Split scenes in manuscript and proof reading files using part, chapter and scene separators.

Based on PyWriter v3.32.0

### v1.8.0 Brief synopsis

- Add a brief synopsis with chapter and scene titles.

Based on PyWriter v3.28.2

### v1.6.2 Optional update

PyWriter Library update:
- Refactor for more compact code. 
- (v3.28.1) Fix a bug where "Notes" and "To do" scenes might not be processed the right way.

Based on PyWriter v3.28.1

### v1.6.1 Optional update

- Change the csv separator (applies only to temporary files).

Based on PyWriter v3.24.0

### v1.6.0 Feature update

- Add notes to the character descriptions.

Based on PyWriter v3.22.0

### v1.4.8 Bugfix release

- Fix a bug causing an exception when a new item is added to a scene.

Based on PyWriter v3.18.1

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
