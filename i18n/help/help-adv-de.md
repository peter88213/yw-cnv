[Projekt-Homepage](https://peter88213.github.io/yw-cnv/) &gt;
[Haupt-Hilfeseite](help-de.html) &gt; Für Fortgeschrittene

------------------------------------------------------------------------

# yWriter import/export: für Fortgeschrittene

## Warnung!

Diese Funktionen sind nur für erfahrene OpenOffice/LibreOffice-Benutzer gedacht. Wenn Sie mit **Calc** und dem Konzept der **Bereiche in Writer** nicht vertraut sind, verwenden Sie die erweiterten Funktionen bitte nicht. Es besteht die Gefahr, dass das Projekt beim Zurückschreiben beschädigt wird, wenn Sie die Bereichsgrenzen in den ODT-Textdokumenten nicht einhalten oder die IDs in den ODS-Tabellen durcheinander bringen. Probieren Sie die Funktionen unbedingt zuerst mit einem Testprojekt aus.

Übersetzt mit www.DeepL.com/Translator (kostenlose Version)
## Command reference

-   [Manuskript mit Kapitel- und Szenenbereichen](#manuskript-mit-kapitel--und-szenenbereichen)
-   [Szenebeschreibungen](#szenebeschreibungen)
-   [Kapitelbeschreibungen](#kapitelbeschreibungen)
-   [Teilebeschreibungen](#teilebeschreibungen)
-   [Figurenbeschreibungen](#figurenbeschreibungen)
-   [Schauplatzbeschreibungen](#schauplatzbeschreibungen)
-   [Gegenstandsbeschreibungen](#gegenstandsbeschreibungen)
-   [Szenenliste](#szenenliste)
-   [Notizen-Kapitel](#notizen-kapitel)
-   [Planungs-Kapitel](#planungs-kapitel)

------------------------------------------------------------------------

## Manuskript mit Kapitel- und Szenenbereichen

Dies lädt yWriter 7-Kapitel und -Szenen in ein neues OpenDocument-Textdokument (odt) mit unsichtbaren Kapitel- und Szenebereichen (die im Navigator zu sehen sind). Das Suffix des Dateinamens ist `_manuscript`.

- Es werden nur "normale" Kapitel und Szenen importiert. Kapitel und Szenen, die als "unbenutzt", "ToDo" oder "Notizen" gekennzeichnet sind, werden nicht importiert.
- Szenen, die mit `<HTML>` oder `<TEX>` beginnen, werden nicht importiert.
- Kommentare innerhalb von Szenen werden als Szenentitel zurückgeschrieben, wenn sie von `~` umgeben sind.
- Kommentare im Text, die mit Schrägstrichen und Sternchen eingeklammert sind (z.B. `/* dies ist ein Kommentar */`), werden in Autorenkommentare umgewandelt.
- Eingestreute HTML-, TEX- oder RTF-Befehle werden unverändert übernommen.
- Gobal- und Projektvariablen werden nicht aufgelöst.
- Kapitel und Szenen können weder neu geordnet noch gelöscht werden.
- Sie können Szenen durch Einfügen von Überschriften oder einem Szenentrenner aufteilen:
    - *Überschrift 1* --> Neue Kapitelüberschrift (Beginn eines neuen Abschnitts).
    - *Überschrift 2* --> Neuer Kapiteltitel.
    - `###` --> Szenentrenner. Optional können Sie den Szenentitel auch an den Szenentrenner anhängen.




Übersetzt mit www.DeepL.com/Translator (kostenlose Version)
You can write back the scene contents and the chapter/part headings to the yWriter 7 project file with the [Export to yWriter](help#export-to-ywriter) command.

-   Comments within scenes are written back as scene titles if surrounded by `~`.

[Zum Seitenbeginn](#top)

------------------------------------------------------------------------

## Szenebeschreibungen

This will generate a new OpenDocument text document (odt) containing a **full synopsis** with chapter titles and scene descriptions that can be edited and written back to yWriter format. File name suffix is `_scenes`.

You can write back the scene descriptions and the chapter/part headings to the yWriter 7 project file with the [Export to yWriter](help#export-to-ywriter) command. Comments right at the beginning of the scene descriptions are written back as scene titles.

[Zum Seitenbeginn](#top)

------------------------------------------------------------------------

## Kapitelbeschreibungen

This will generate a new OpenDocument text document (odt) containing a **brief synopsis** with chapter titles and chapter descriptions that can be edited and written back to yWriter format. File name suffix is `_chapters`.

**Note:** Doesn't apply to chapters marked `This chapter begins a new section` in yWriter.

You can write back the headings and descriptions to the yWriter 7 project file with the [Export to yWriter](help#export-to-ywriter) command.

[Zum Seitenbeginn](#top)

------------------------------------------------------------------------

## Teilebeschreibungen

This will generate a new OpenDocument text document (odt) containing a **very brief synopsis** with part titles and part descriptions that can be edited and written back to yWriter format. File name suffix is `_parts`.

**Note:** Applies only to chapters marked `This chapter  begins a new section` in yWriter.

You can write back the headings and descriptions to the yWriter 7 project file with the [Export to yWriter](help#export-to-ywriter) command.

[Zum Seitenbeginn](#top)

------------------------------------------------------------------------

## Figurenbeschreibungen

This will generate a new OpenDocument text document (odt) containing character descriptions, bio, goals, and notes that can be edited in Office Writer and written back to yWriter format. File name suffix is `_characters`.

You can write back the descriptions to the yWriter 7 project file with the [Export to yWriter](help#export-to-ywriter) command.

[Zum Seitenbeginn](#top)

------------------------------------------------------------------------

## Schauplatzbeschreibungen

This will generate a new OpenDocument text document (odt) containing location descriptions that can be edited in Office Writer and written back to yWriter format. File name suffix is `_locations`.

You can write back the descriptions to the yWriter 7 project file with the [Export to yWriter](help#export-to-ywriter) command.

[Zum Seitenbeginn](#top)

------------------------------------------------------------------------

## Gegenstandsbeschreibungen

This will generate a new OpenDocument text document (odt) containing item descriptions that can be edited in Office Writer and written back to yWriter format. File name suffix is `_items`.

You can write back the descriptions to the yWriter 7 project file with the [Export to yWriter](help#export-to-ywriter) command.

[Zum Seitenbeginn](#top)

------------------------------------------------------------------------

## Szenenliste

This will generate a new OpenDocument spreadsheet (ods) listing the following:

-   Hyperlink to the manuscript's scene section
-   Scene title
-   Scene description
-   Tags
-   Scene notes
-   A/R
-   Goal
-   Conflict
-   Outcome
-   Sequential scene number
-   Words total
-   Rating 1
-   Rating 2
-   Rating 3
-   Rating 4
-   Word count
-   Letter count
-   Status
-   Characters
-   Locations
-   Items

Only "normal" scenes that would be exported as RTF in yWriter get a row in the scene list. Scenes of the "Unused", "Notes", or "ToDo" type are omitted.

File name suffix is `_scenelist`.

You can write back the table contents to the yWriter 7 project file with the [Export to yWriter](help#export-to-ywriter) command.

The following columns can be written back to the yWriter project:

-   Title
-   Description
-   Tags (comma-separated)
-   Scene notes
-   A/R (action/reaction scene)
-   Goal
-   Conflict
-   Outcome
-   Rating 1
-   Rating 2
-   Rating 3
-   Rating 4
-   Status ('Outline', 'Draft', '1st Edit', '2nd Edit', 'Done')

[Zum Seitenbeginn](#top)

------------------------------------------------------------------------

## Notizen-Kapitel

This will write yWriter 7 "Notes" chapters with child scenes into a new OpenDocument text document (odt) with invisible chapter and scene sections (to be seen in the Navigator). File name suffix is `_notes`.

-   Comments within scenes are written back as scene titles if surrounded by `~`.
-   Chapters and scenes can neither be rearranged nor deleted.
-   Scenes can be split by inserting headings or a scene divider:
    -   *Heading 1* --› New chapter title (beginning a new section).
    -   *Heading 2* --› New chapter title.
    -   `###` --› Scene divider. Optionally, you can append the scene
        title to the scene divider.

[Zum Seitenbeginn](#top)

------------------------------------------------------------------------

## Planungs-Kapitel

This will write yWriter 7 "Todo" chapters with child scenes into a new 
OpenDocument text document (odt) with invisible chapter and scene 
sections (to be seen in the Navigator). File name suffix is `_todo`.

-  Comments within scenes are written back as scene titles
   if surrounded by `~`.
-  Chapters and scenes can neither be rearranged nor deleted.
-  Scenes can be split by inserting headings or a scene divider:
    -  *Heading 1* --› New chapter title (beginning a new section).
    -  *Heading 2* --› New chapter title.
    -  `###` --› Scene divider. Optionally, you can append the 
       scene title to the scene divider.

[Zum Seitenbeginn](#top)

------------------------------------------------------------------------
