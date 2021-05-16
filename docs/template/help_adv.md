# yWriter import/export: advanced features {#top}

## Warning!

The \"advanced features\" are meant to be used by experienced
OpenOffice/LibreOffice users only. If you aren\'t familiar with **Calc**
and the concept of **sections in Writer**, please do not use the
advanced features. There is a risk of damaging the project when writing
back if you don\'t respect the section boundaries in the ODT text
documents, or if you mess up the IDs in the ODS tables. Be sure to try
out the features with a test project first.

[Main help page](@help@)

## Command reference

-   [Manuscript with chapter and scene
    sections](#manuscript_with_chapter_and_scene_sections)
-   [Scene descriptions](#scene_descriptions)
-   [Chapter descriptions](#chapter_descriptions)
-   [Part descriptions](#part_descriptions)
-   [Character descriptions](#character_descriptions)
-   [Location descriptions](#location_descriptions)
-   [Item descriptions](#item_descriptions)
-   [Scene list](#scene_list)
-   [Plot list](#plot_list)

------------------------------------------------------------------------

## Manuscript with chapter and scene sections {#manuscript_with_chapter_and_scene_sections}

This will load yWriter 6/7 chapters and scenes into a new OpenDocument
text document (odt) with invisible chapter and scene sections (to be
seen in the Navigator). File name suffix is `_manuscript`.

You can write back the edited scene contents to the yWriter 6/7 project
file with the [Export to yWriter](help.html#export_to_ywriter) command.

[Top of page](#top)

------------------------------------------------------------------------

## Scene descriptions {#scene_descriptions}

This will generate a new OpenDocument text document (odt) containing a
**full synopsis** with chapter titles and scene descriptions that can be
edited and written back to yWriter format. File name suffix is
`_scenes`.

You can write back the edited descriptions to the yWriter 6/7 project
file with the [Export to yWriter](help.html#export_to_ywriter) command.

[Top of page](#top)

------------------------------------------------------------------------

## Chapter descriptions {#chapter_descriptions}

This will generate a new OpenDocument text document (odt) containing a
**brief synopsis** with chapter titles and chapter descriptions that can
be edited and written back to yWriter format. File name suffix is
`_chapters`.

**Note:** Doesn\'t apply to chapters marked
`This chapter begins a new section` in yWriter.

You can write back the edited descriptions to the yWriter 6/7 project
file with the [Export to yWriter](help.html#export_to_ywriter) command.

[Top of page](#top)

------------------------------------------------------------------------

## Part descriptions {#part_descriptions}

This will generate a new OpenDocument text document (odt) containing a
**very brief synopsis** with part titles and part descriptions that can
be edited and written back to yWriter format. File name suffix is
`_parts`.

**Note:** Applies only to chapters marked
`This chapter  begins a new section` in yWriter.

You can write back the edited descriptions to the yWriter 6/7 project
file with the [Export to yWriter](help.html#export_to_ywriter) command.

[Top of page](#top)

------------------------------------------------------------------------

## Character descriptions {#character_descriptions}

This will generate a new OpenDocument text document (odt) containing
character descriptions, bio and goals that can be edited in Office
Writer and written back to yWriter format. File name suffix is
`_characters`.

You can write back the edited descriptions to the yWriter 6/7 project
file with the [Export to yWriter](help.html#export_to_ywriter) command.

[Top of page](#top)

------------------------------------------------------------------------

## Location descriptions {#location_descriptions}

This will generate a new OpenDocument text document (odt) containing
location descriptions that can be edited in Office Writer and written
back to yWriter format. File name suffix is `_locations`.

You can write back the edited descriptions to the yWriter 6/7 project
file with the [Export to yWriter](help.html#export_to_ywriter) command.

[Top of page](#top)

------------------------------------------------------------------------

## Item descriptions {#item_descriptions}

This will generate a new OpenDocument text document (odt) containing
item descriptions that can be edited in Office Writer and written back
to yWriter format. File name suffix is `_items`.

You can write back the edited descriptions to the yWriter 6/7 project
file with the [Export to yWriter](help.html#export_to_ywriter) command.

[Top of page](#top)

------------------------------------------------------------------------

## Scene list {#scene_list}

This will generate a new OpenDocument spreadsheet (ods) listing scene
title, scene descriptions, and links to the manuscript\'s scene
sections. Further scene metadata (e.g. tags, goals, time), if any. File
name suffix is `_scenelist`.

You can write back the edited table to the yWriter 6/7 project file with
the [Export to yWriter](help.html#export_to_ywriter) command.

[Top of page](#top)

------------------------------------------------------------------------

## Plot list {#plot_list}

This will generate a new OpenDocument spreadsheet (ods) listing plot
related metadata that can be displayed and edited. File name suffix is
`_plotlist`.

### Plotting conventions

In yWriter, you can divide your novel into **Plot Sections** (e.g. acts
or steps) by inserting chapters marked \"other\". They will show up in
blue color and won\'t get exported.

**Plot-related events** (e.g. \"Mid Point\", \"Climax\") can be
identified by \"scene tags\" if you want to link them to a specific
scene.

You can use scene notes for **plot-specific explanations**.

If you want to **visualize character arcs**, you can use the project\'s
rating names by changing them to the names of up to four main
characters. Then you can quantify the state of these four characters and
put them into the scenes. It\'s easy then to let OpenOffice Calc show a
diagram for the scene ratings over scene count or word count.

-   Only rating field names corresponding to character names or
    containing the string \"story\", e.g. \"A-Story\", \"BStoryline\"
    (up to 10 case insensitive characters) appear in the plot list.
-   Only ratings greater than 1 appear in the plot list, i.e. 1 means
    \"a rating is not set for this chapter\".
-   Recommended ratings:
    -   1 = N/A
    -   2 = unhappy
    -   3 = dissatisfied
    -   4 = vague
    -   5 = satisfied
    -   6 = happy
-   Ratings deleted while editing the plotlist will be converted to 1 on
    writing back.

You can write back the edited table to the yWriter 6/7 project file with
the [Export to yWriter](help.html#export_to_ywriter) command.

[Top of page](#top)

------------------------------------------------------------------------
