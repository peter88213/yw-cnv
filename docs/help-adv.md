[Project home page](https://peter88213.github.io/yw-cnv/) > [Main help page](help) > Advanced features

------------------------------------------------------------------------

# yWriter import/export: advanced features

## Warning!

The "advanced features" are meant to be used by experienced
OpenOffice/LibreOffice users only. If you aren't familiar with **Calc**
and the concept of **sections in Writer**, please do not use the
advanced features. There is a risk of damaging the project when writing
back if you don't respect the section boundaries in the ODT text
documents, or if you mess up the IDs in the ODS tables. Be sure to try
out the features with a test project first.

## Command reference

-   [Manuscript with chapter and scene
    sections](#manuscript-with-chapter-and-scene-sections)
-   [Scene descriptions](#scene-descriptions)
-   [Chapter descriptions](#chapter-descriptions)
-   [Part descriptions](#part-descriptions)
-   [Character descriptions](#character-descriptions)
-   [Location descriptions](#location-descriptions)
-   [Item descriptions](#item-descriptions)
-   [Scene list](#scene-list)
-   [Notes chapters](#notes-chapters)

------------------------------------------------------------------------

## Manuscript with chapter and scene sections

This will load yWriter 7 chapters and scenes into a new OpenDocument
text document (odt) with invisible chapter and scene sections (to be
seen in the Navigator). File name suffix is `_manuscript`.

-  Chapters and scenes can neither be rearranged nor deleted.
-  You can split scenes by inserting headings or a scene divider:
    -  *Heading 1* --› New chapter title (beginning a new section).
    -  *Heading 2* --› New chapter title.
    -  `###` --› Scene divider. Optionally, you can also append the 
       scene title to the scene divider.

You can write back the scene contents and the chapter/part headings to 
the yWriter 7 project file with the [Export to yWriter](help#export-to-ywriter) 
command. 

-   Comments within scenes are written back as scene titles
    if surrounded by `~`.

[Top of page](#top)

------------------------------------------------------------------------

## Scene descriptions

This will generate a new OpenDocument text document (odt) containing a
**full synopsis** with chapter titles and scene descriptions that can be
edited and written back to yWriter format. File name suffix is
`_scenes`.

You can write back the scene descriptions and the chapter/part headings 
to the yWriter 7 project file with the [Export to yWriter](help#export-to-ywriter) 
command. Comments right at the beginning of the scene descriptions are 
written back as scene titles.

[Top of page](#top)

------------------------------------------------------------------------

## Chapter descriptions

This will generate a new OpenDocument text document (odt) containing a
**brief synopsis** with chapter titles and chapter descriptions that can
be edited and written back to yWriter format. File name suffix is
`_chapters`.

**Note:** Doesn't apply to chapters marked
`This chapter begins a new section` in yWriter.

You can write back the headings and descriptions to the yWriter 7 project 
file with the [Export to yWriter](help#export-to-ywriter) command.

[Top of page](#top)

------------------------------------------------------------------------

## Part descriptions

This will generate a new OpenDocument text document (odt) containing a
**very brief synopsis** with part titles and part descriptions that can
be edited and written back to yWriter format. File name suffix is
`_parts`.

**Note:** Applies only to chapters marked
`This chapter  begins a new section` in yWriter.

You can write back the headings and descriptions to the yWriter 7 project
file with the [Export to yWriter](help#export-to-ywriter) command.

[Top of page](#top)

------------------------------------------------------------------------

## Character descriptions

This will generate a new OpenDocument text document (odt) containing
character descriptions, bio, goals, and notes that can be edited in Office
Writer and written back to yWriter format. File name suffix is
`_characters`.

You can write back the descriptions to the yWriter 7 project
file with the [Export to yWriter](help#export-to-ywriter) command.

[Top of page](#top)

------------------------------------------------------------------------

## Location descriptions

This will generate a new OpenDocument text document (odt) containing
location descriptions that can be edited in Office Writer and written
back to yWriter format. File name suffix is `_locations`.

You can write back the descriptions to the yWriter 7 project
file with the [Export to yWriter](help#export-to-ywriter) command.

[Top of page](#top)

------------------------------------------------------------------------

## Item descriptions

This will generate a new OpenDocument text document (odt) containing
item descriptions that can be edited in Office Writer and written back
to yWriter format. File name suffix is `_items`.

You can write back the descriptions to the yWriter 7 project
file with the [Export to yWriter](help#export-to-ywriter) command.

[Top of page](#top)

------------------------------------------------------------------------

## Scene list

This will generate a new OpenDocument spreadsheet (ods) listing the following:

- Hyperlink to the manuscript's scene section
- Scene title
- Scene description
- Tags
- Scene notes
- A/R
- Goal
- Conflict
- Outcome
- Sequential scene number
- Words total
- Rating 1
- Rating 2
- Rating 3
- Rating 4
- Word count
- Letter count
- Status
- Characters
- Locations
- Items

Only "normal" scenes that would be exported as RTF in yWriter get a 
row in the scene list. Scenes of the "Unused", "Notes", or "ToDo" 
type are omitted.

File name suffix is `_scenelist`.

You can write back the table contents to the yWriter 7 project file with
the [Export to yWriter](help#export-to-ywriter) command.

The following columns can be written back to the yWriter project:

- Title
- Description
- Tags (comma-separated)
- Scene notes
- A/R (action/reaction scene)
- Goal
- Conflict
- Outcome
- Rating 1
- Rating 2
- Rating 3
- Rating 4
- Status ('Outline', 'Draft', '1st Edit', '2nd Edit', 'Done')

[Top of page](#top)

------------------------------------------------------------------------

## Notes chapters

This will write yWriter 7 "Notes" chapters with child scenes into a new 
OpenDocument text document (odt) with invisible chapter and scene 
sections (to be seen in the Navigator). File name suffix is `_notes`.

-  Comments within scenes are written back as scene titles
   if surrounded by `~`.
-  Chapters and scenes can neither be rearranged nor deleted.
-  Scenes can be split by inserting headings or a scene divider:
    -  *Heading 1* --› New chapter title (beginning a new section).
    -  *Heading 2* --› New chapter title.
    -  `###` --› Scene divider. Optionally, you can append the 
       scene title to the scene divider.

[Top of page](#top)

------------------------------------------------------------------------
