[Projekt-Homepage](https://peter88213.github.io/yw-cnv/) &gt;
Haupt-Hilfeseite

------------------------------------------------------------------------

# yWriter import/export

## Befehlsreferenz

### "Files" menu

-   [Zu yWriter exportieren](#export-to-ywriter)
-   [Von yWriter importieren](#import-from-ywriter)
-   [Von yWriter zum Korrekturlesen
    importieren](#import-from-ywriter-for-proof-reading)
-   [Kurze Zusammenfassung](#brief-synopsis)
-   [Figurenliste](#character-list)
-   [Schauplatzliste](#location-list)
-   [Gegenständeliste](#item-list)
-   [Querverweise](#cross-reference)
-   [Für Fortgeschrittene](help-adv-de.html)

### "Format" menu

-   [Szenentrenner durch Leerzeilen
    ersetzen](#replace-scene-dividers-with-blank-lines)
-   [Absätze einrücken, die mit '&gt;'
    beginnen'](#indent-paragraphs-that-start-with)
-   [Aufzählungsstriche ersetzen](#replace-list-strokes-with-bullets)

## HowTo

-   [How to set up a work in progress for
    export](#how-to-set-up-a-work-in-progress-for-export)
-   [How to set up an outline for
    export](#how-to-set-up-an-outline-for-export)

------------------------------------------------------------------------

## Zu yWriter exportieren

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

## Von yWriter importieren

This will load yWriter 7 chapters and scenes into a new OpenDocument
text document (odt).

-   The document is placed in the same folder as the yWriter project.
-   Document's **filename**: `<yW project name>.odt`.
-   Text markup: Bold and italics are supported. Other highlighting such
    as underline and strikethrough are lost.
-   Only "normal" chapters and scenes are imported. Chapters and scenes
    marked "unused", "todo" or "notes" are not imported.
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
-   Scenes are separated by `* * *`. The first line is not indented. You
    can replace the scene separators by blank lines with the menu
    command **Format &gt; Szenentrenner durch Leerzeilen ersetzen**.
-   Starting from the second paragraph, paragraphs begin with
    indentation of the first line.
-   Scenes marked "attach to previous scene" in yWriter appear like
    continuous paragraphs.
-   Paragraphs starting with `>` are formatted as quotations.

[Top of page](#top)

------------------------------------------------------------------------

## Von yWriter zum Korrekturlesen importieren

This will load yWriter 7 chapters and scenes into a new OpenDocument
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
    -   *Heading 1* --› New chapter title (beginning a new section).
    -   *Heading 2* --› New chapter title.
    -   `###` --› Scene divider. Optionally, you can append the scene
        title to the scene divider.

You can write back the scene contents to the yWriter 7 project file with
the [Zu yWriter exportieren](#export-to-ywriter) command.

[Top of page](#top)

------------------------------------------------------------------------

## Kurze Zusammenfassung

This will load a brief synopsis with chapter and scenes titles into a
new OpenDocument text document (odt).

-   The document is placed in the same folder as the yWriter project.
-   Document's **filename**: `<yW project name_brf_synopsis>.odt`.
-   Only "normal" chapters and scenes are imported. Chapters and scenes
    marked "unused", "todo" or "notes" are not imported.
-   Only scenes that are intended for RTF export in yWriter will be
    imported.
-   Titles of scenes beginning with `<HTML>` or `<TEX>` are not
    imported.
-   Chapter titles appear as first level heading if the chapter is
    marked as beginning of a new section in yWriter. Such headings are
    considered as "part" headings.
-   Chapter titles appear as second level heading if the chapter is not
    marked as beginning of a new section. Such headings are considered
    as "chapter" headings.
-   Scene titles appear as plain paragraphs.

[Top of page](#top)

------------------------------------------------------------------------

## Figurenliste

This will generate a new OpenDocument spreadsheet (ods) containing a
character list that can be edited in Office Calc and written back to
yWriter format. File name suffix is `_charlist`.

You may change the sort order of the rows. You may also add or remove
rows. New entities must get a unique ID.

You can write back the edited table to the yWriter 7 project file with
the [Zu yWriter exportieren](#export-to-ywriter) command.

[Top of page](#top)

------------------------------------------------------------------------

## Schauplatzliste

This will generate a new OpenDocument spreadsheet (ods) containing a
location list that can be edited in Office Calc and written back to
yWriter format. File name suffix is `_loclist`.

You may change the sort order of the rows. You may also add or remove
rows. New entities must get a unique ID.

You can write back the edited table to the yWriter 7 project file with
the [Zu yWriter exportieren](#export-to-ywriter) command.

[Top of page](#top)

------------------------------------------------------------------------

## Gegenständeliste

This will generate a new OpenDocument spreadsheet (ods) containing an
item list that can be edited in Office Calc and written back to yWriter
format. File name suffix is `_itemlist`.

You may change the sort order of the rows. You may also add or remove
rows. New entities must get a unique ID.

You can write back the edited table to the yWriter 7 project file with
the [Zu yWriter exportieren](#export-to-ywriter) command.

[Top of page](#top)

------------------------------------------------------------------------

## Querverweise

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

## How to set up a work in progress for export

Generate a new yWriter 7 project from a work in progress:

-   The new yWriter project is placed in the same folder as the
    document.
-   yWriter project's filename: `<document name>.yw7`.
-   Existing yWriter 7 projects will not be overwritten.

### How to format a work in progress:

A work in progress has no third level heading.

-   *Heading 1* --› New chapter title (beginning a new section).
-   *Heading 2* --› New chapter title.
-   `* * *` --› Scene divider (not needed for the first scenes in a
    chapter).
-   Comments right at the scene beginning are considered scene titles.
-   All other text is considered scene content.

[Top of page](#top)

------------------------------------------------------------------------

## How to set up an outline for export

Generate a new yWriter 7 project from an outline:

-   The new yWriter project is placed in the same folder as the
    document.
-   yWriter project's filename: `<document name>.yw7`.
-   Existing yWriter 7 projects will not be overwritten.

### How to format an outline:

An outline has at least one third level heading.

-   *Heading 1* --› New chapter title (beginning a new section).
-   *Heading 2* --› New chapter title.
-   *Heading 3* --› New scene title.
-   All other text is considered to be chapter/scene description.

[Top of page](#top)

------------------------------------------------------------------------

## Szenentrenner durch Leerzeilen ersetzen

This will replace the three-line "\* \* \*" scene dividers with single
blank lines. The style of the scene-dividing lines will be changed from
*Heading 4* to *Heading 5*.

[Top of page](#top)

------------------------------------------------------------------------

## Absätze einrücken, die mit '&gt;' beginnen'

This will select all paragraphs that start with "&gt; " and change their
paragraph style to *Quotations*.

Note: When exporting to yWriter, *Quotations* style paragraphs will
automatically marked with "&gt; " at the beginning.

[Top of page](#top)

------------------------------------------------------------------------

## Aufzählungsstriche ersetzen

This will select all paragraphs that start with "- " and apply a list
paragraph style.

Note: When exporting to yWriter, Lists will automatically marked with "-
" list strokes.

[Top of page](#top)

------------------------------------------------------------------------
