[Project home page](https://peter88213.github.io/yw-cnv/) > Main help page

------------------------------------------------------------------------

# yWriter import/export

## Command reference

### "Files" menu

-   [Export to yWriter](#export-to-ywriter) 
-   [Import from yWriter](#import-from-ywriter)
-   [Import from yWriter for proof reading](#import-from-ywriter-for-proof-reading)
-   [Brief synopsis](#brief-synopsis)
-   [Character list](#character-list)
-   [Location list](#location-list)
-   [Item list](#item-list)
-   [Cross reference](#cross-reference)
-   [Advanced features](help-adv)

### "Format" menu

-   [Replace scene dividers with blank lines](#replace-scene-dividers-with-blank-lines)
-   [Indent paragraphs that start with '> '](#indent-paragraphs-that-start-with)
-   [Replace list strokes with bullets](#replace-list-strokes-with-bullets)

------------------------------------------------------------------------

## General

### About formatting text

It is assumed that very few types of text markup are needed for a fictional text:

- *Emphasized* (usually shown as italics).
- *Strongly emphasized* (usually shown as capitalized).
- *Citation* (paragraph visually distinguished from body text).

When importing from yw7 format, the converter replaces these formattings as follows: 

- Text with *italics* in yWriter is formatted as *Emphasized*.
- *Bold* text in yWriter is formatted as *Strongly emphasized*. 
- Paragraphs starting with `"> "` in yWriter, are formatted as *Quote*.

When exporting to yw7 format, the reverse is the case. 

### About document language handling

ODF documents are generally assigned a language that determines spell checking and country-specific character substitutions. In addition, Office Writer lets you assign text passages to languages other than the document language to mark foreign language usage or to suspend spell checking. 

#### Document overall

- If a document language (Language code acc. to ISO 639-1 and country code acc. to ISO 3166-2) is detected in the source document during conversion to yw7 format, these codes are set as yWriter project variables. 

- If language code and country code exist as project variables during conversion from yw7 format, they are inserted into the generated ODF document. 

- If no language and country code exist as project variables when converting from yw7 format, language and country code from the operating system settings are entered into the generated ODF document. 

- The language and country codes are checked superficially. If they obviously do not comply with the ISO standards, they are replaced by the values for "No language". These are:
    - Language = zxx
    - Country = none

#### Text passages in scenes

If text markups for other languages are detected during conversion to the yw7 format, they are converted and transferred to the yWriter scene. 

This then looks like this, for example:

`xxx xxxx [lang=en-AU]yyy yyyy yyyy[/lang=en-AU] xxx xxx` 

To prevent these text markups from interfering with *yWriter*, they are automatically set as project variables in such a way that *yWriter* interprets them as HTML instructions during document export. 

For the example shown above, the project variable definition for the opening tag looks like this: 

- *Variable Name:* `lang=en-AU` 
- *Value/Text:* `<HTM <SPAN LANG="en-AU"> /HTM>`

The point of this is that such language assignments are preserved even after multiple conversions in both directions, so they are always effective for spell checking in the ODT document.

It is recommended not to modify such markups in yWriter to avoid unwanted nesting and broken enclosing. 

## HowTo

## How to set up a work in progress for export

A work in progress has no third level heading.

-   *Heading 1* → New chapter title (beginning a new section).
-   *Heading 2* → New chapter title.
-   `* * *` → Scene divider (not needed for the first scenes in a
    chapter).
-   Comments right at the scene beginning are considered scene titles.
-   All other text is considered scene content.

## How to set up an outline for export

An outline has at least one third level heading.

-   *Heading 1* → New chapter title (beginning a new section).
-   *Heading 2* → New chapter title.
-   *Heading 3* → New scene title.
-   All other text is considered to be chapter/scene description.

[Top of page](#top)

------------------------------------------------------------------------

## Export to yWriter

This writes back the document's content to the yWriter project file.

-   Make sure not to change a generated document's file name before
    writing back to yWriter format.
-   The yWriter 7 project to rewrite must exist in the same folder as
    the document.
-   If the document's file name has no suffix, the document is
    considered a [Work in
    progress](#how-to-set-up-a-work-in-progress-for-export) or an
    [Outline](#how-to-set-up-an-outline-for-export) to be exported into
    a newly created yWriter project. **Note:** Existing yWriter projects
    will not be overwritten.

[Top of page](#top)

------------------------------------------------------------------------

## Import from yWriter

This will write yWriter 7 chapters and scenes into a new OpenDocument
text document (odt).

-   The document is placed in the same folder as the yWriter project.
-   Document's **filename**: `<yW project name>.odt`.
-   Text markup: Bold and italics are supported. Other highlighting such
    as underline and strikethrough are lost.
-   Only "normal" chapters and scenes are imported. Chapters and
    scenes marked "unused", "todo" or "notes" are not imported.
-   Only scenes that are intended for RTF export in yWriter will be
    imported.
-   Scenes beginning with `<HTML>` or `<TEX>` are not imported.
-   Comments in the text bracketed with slashes and asterisks (like
    `/* this is a comment */`) are converted to author's comments.
-   Interspersed HTML, TEX, or RTF commands are removed.
-   Gobal variables and project variables are not resolved.
-   Chapter titles appear as first level heading if the chapter is
    marked as beginning of a new section in yWriter. Such headings are
    considered as "part" headings.
-   Chapter titles appear as second level heading if the chapter is not
    marked as beginning of a new section. Such headings are considered
    as "chapter" headings.
-   Scene titles appear as navigable comments pinned to the beginning of
    the scene.
-   Scenes are separated by `* * *`. The first line is not
    indented. You can replace the scene separators by blank lines with 
    the menu command **Format >  Replace scene dividers with blank lines**.
-   Starting from the second paragraph, paragraphs begin with
    indentation of the first line.
-   Scenes marked "attach to previous scene" in yWriter appear like
    continuous paragraphs.
-   Paragraphs starting with `> ` are formatted as quotations.


[Top of page](#top)

------------------------------------------------------------------------

## Import from yWriter for proof reading

This will write yWriter 7 chapters and scenes into a new OpenDocument
text document (odt) with chapter and scene markers. File name suffix is
`_proof`.

-   The proof read document is placed in the same folder as the yWriter
    project.
-   Document's filename: `<yW project name>_proof.odt`.
-   Text markup: Bold and italics are supported. Other highlighting such
    as underline and strikethrough are lost.
-   Scenes beginning with `<HTML>` or `<TEX>` are not imported.
-   All other chapters and scenes are imported, whether "used" or
    "unused".
-   Interspersed HTML, TEX, or RTF commands are taken over unchanged.
-   The document contains chapter `[ChID:x]` and scene `[ScID:y]`
    markers according to yWriter 5 standard. **Do not touch lines
    containing the markers** if you want to be able to reimport the
    document into yWriter.
-   Chapters and scenes can neither be rearranged nor deleted. 
-   You can split scenes by inserting headings or a scene divider:
    -  *Heading 1* → New chapter title (beginning a new section).
    -  *Heading 2* → New chapter title.
    -  `###` → Scene divider. Optionally, you can append the 
       scene title to the scene divider.
-   Paragraphs starting with `> ` are formatted as quotations.

You can write back the scene contents to the yWriter 7 project file
with the [Export to yWriter](#export-to-ywriter) command.

[Top of page](#top)

------------------------------------------------------------------------

## Brief synopsis

This will write a brief synopsis with chapter and scenes titles into a new 
OpenDocument text document.  File name suffix is `_brf_synopsis`.
 
-   Only scenes that are intended for RTF export in yWriter will be
    imported.
-   Titles of scenes beginning with `<HTML>` or `<TEX>` are not imported.
-   Chapter titles appear as first level heading if the chapter is
    marked as beginning of a new section in yWriter. Such headings are
    considered as "part" headings.
-   Chapter titles appear as second level heading if the chapter is not
    marked as beginning of a new section. Such headings are considered
    as "chapter" headings.
-   Scene titles appear as plain paragraphs.



[Top of page](#top)

------------------------------------------------------------------------

## Character list

This will generate a new OpenDocument spreadsheet (ods) containing a
character list that can be edited in Office Calc and written back to
yWriter format. File name suffix is `_charlist`.

You may change the sort order of the rows. You may also add or remove
rows. New entities must get a unique ID.

You can write back the edited table to the yWriter 7 project file with
the [Export to yWriter](#export-to-ywriter) command.

[Top of page](#top)

------------------------------------------------------------------------

## Location list

This will generate a new OpenDocument spreadsheet (ods) containing a
location list that can be edited in Office Calc and written back to
yWriter format. File name suffix is `_loclist`.

You may change the sort order of the rows. You may also add or remove
rows. New entities must get a unique ID.

You can write back the edited table to the yWriter 7 project file with
the [Export to yWriter](#export-to-ywriter) command.

[Top of page](#top)

------------------------------------------------------------------------

## Item list

This will generate a new OpenDocument spreadsheet (ods) containing an
item list that can be edited in Office Calc and written back to yWriter
format. File name suffix is `_itemlist`.

You may change the sort order of the rows. You may also add or remove
rows. New entities must get a unique ID.

You can write back the edited table to the yWriter 7 project file with
the [Export to yWriter](#export-to-ywriter) command.

[Top of page](#top)

------------------------------------------------------------------------

## Cross reference

This will generate a new OpenDocument text document (odt) containing
navigable cross references. File name suffix is `_xref`. The cross
references are:

-   Scenes per character,
-   scenes per location,
-   scenes per item,
-   scenes per tag,
-   characters per tag,
-   locations per tag,
-   items per tag.

[Top of page](#top)

------------------------------------------------------------------------

## Replace scene dividers with blank lines

This will replace the three-line "* * *" scene dividers
with single blank lines. The style of the scene-dividing
lines will be changed from  _Heading 4_  to  _Heading 5_.

[Top of page](#top)

------------------------------------------------------------------------

## Indent paragraphs that start with '> '

This will select all paragraphs that start with "> "
and change their paragraph style to _Quotations_.

Note: When exporting to yWriter, _Quotations_ style paragraphs will
automatically marked with "> " at the beginning.

[Top of page](#top)

------------------------------------------------------------------------

## Replace list strokes with bullets

This will select all paragraphs that start with "- "
and apply a list paragraph style.

Note: When exporting to yWriter, Lists will
automatically marked with "- " list strokes.

[Top of page](#top)

------------------------------------------------------------------------

