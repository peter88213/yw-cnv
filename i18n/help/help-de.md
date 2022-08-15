[Projekt-Homepage](https://peter88213.github.io/yw-cnv/) &gt;
Haupt-Hilfeseite

------------------------------------------------------------------------

# yWriter import/export

## Befehlsreferenz

### "Files" menu

-   [Zu yWriter exportieren](#zu-ywriter-exportieren)
-   [Von yWriter importieren](#von-ywriter-importieren)
-   [Von yWriter zum Korrekturlesen
    importieren](#von-ywriter-zum-korrekturlesen-importieren)
-   [Kurze Zusammenfassung](#kurze-zusammenfassung)
-   [Figurenliste](#figurenliste)
-   [Schauplatzliste](#schauplatzliste)
-   [Gegenstandsliste](#gegenstandsliste)
-   [Querverweise](#querverweise)
-   [Für Fortgeschrittene](help-adv-de.html)

### "Format" menu

-   [Szenentrenner durch Leerzeilen
    ersetzen](#szenentrenner-durch-leerzeilen-ersetzen)
-   [Absätze einrücken, die mit '> ' beginnen'](#absätze-einrücken-die-mit-beginnen)
-   [Aufzählungsstriche ersetzen](#aufzählungsstriche-ersetzen)

## So geht's

-   [Ein Manuskript zum Export vorbereiten](#ein-manuskript-zum-export-vorbereiten)
-   [Eine Gliederung zum Export vorbereiten](#eine-gliederung-zum-export-vorbereiten)

------------------------------------------------------------------------

## Zu yWriter exportieren

Dadurch wird der Inhalt des Dokuments in die yWriter-Projektdatei zurückgeschrieben.

-   Stellen Sie sicher, dass Sie den Dateinamen eines generierten Dokuments vor dem Zurückschreiben in das yWriter-Formanicht ändern.
-   Das zu überschreibende yWriter 7 Projekt muss sich im selben Ordner befinden wie das Dokument.
-   Wenn der Dateiname des Dokuments kein Suffix hat, dient das Dokument als [Manuskript in Arbeit](#how-to-set-up-a-work-in-progress-for-export) oder eine [Gliederung](#how-to-set-up-an-outline-for-export) zum Exportieren in ein neu zu erstellendes yWriter-Projekt. **Hinweis:** Bestehende  yWriter-Projekte werden nicht überschrieben.

[Zum Seitenbeginn](#top)

------------------------------------------------------------------------

## Von yWriter importieren

Dadurch werden yWriter 7-Kapitel und -Szenen in ein neues OpenDocument Textdokument (odt) importiert.

- Das Dokument wird im selben Ordner wie das yWriter-Projekt abgelegt.
- Der **Dateiname** des Dokuments: `<yW-Projektname>.odt`.
- Textauszeichnung: Fett und kursiv werden unterstützt. Andere Hervorhebungen wie Unterstreichen und Durchstreichen gehen verloren.
- Es werden nur "normale" Kapitel und Szenen importiert. Kapitel und Szenen, die als "unbenutzt", "ToDo" oder "Notizen" gekennzeichnet sind, werden nicht importiert.
- Nur Szenen, die für den RTF-Export in yWriter vorgesehen sind, werden importiert.
- Szenen, die mit `<HTML>` oder `<TEX>` beginnen, werden nicht importiert.
- Kommentare im Text, die mit Schrägstrichen und Sternchen eingeklammert sind (z.B. `/* this is a comment */`), werden in Autorenkommentare umgewandelt.
- Eingestreute HTML-, TEX- oder RTF-Befehle werden entfernt.
- Grundlegende Variablen und Projektvariablen werden nicht aufgelöst.
- Kapitelüberschriften erscheinen als Überschriften der ersten Ebene, wenn das Kapitel in yWriter als Beginn eines neuen Abschnitts markiert ist. Solche Überschriften werden als "Teil"-Überschriften betrachtet.
- Kapitelüberschriften erscheinen als Überschriften der zweiten Ebene, wenn das Kapitel nicht als Beginn eines neuen Abschnitts markiert ist. Solche Überschriften werden als "Kapitel"-Überschriften betrachtet.
- Szenentitel erscheinen als navigierbare Kommentare, die an den Anfang der Szene gepinnt sind.
- Szenen werden durch `* * *` getrennt. Die erste Zeile wird nicht eingerückt. Mit dem Menübefehl **Format &gt; Szenentrenner durch Leerzeilen ersetzen** können Sie die Szenentrenner durch Leerzeilen ersetzen.
- Ab dem zweiten Absatz beginnen Absätze mit der Einrückung der ersten Zeile.
- Szenen, die in yWriter mit "an vorherige Szene anhängen" gekennzeichnet sind, erscheinen wie durchgehende Absätze.
- Absätze, die mit `>` beginnen, werden als Zitate formatiert.

[Zum Seitenbeginn](#top)

------------------------------------------------------------------------

## Von yWriter zum Korrekturlesen importieren

This will load yWriter 7 chapters and scenes into a new OpenDocument text document (odt) with chapter and scene markers. File name suffix is `_proof`.

-   The proof read document is placed in the same folder as the yWriter  project.
-   Document's filename: `<yW project name>_proof.odt`.
-   Text markup: Bold and italics are supported. Other highlighting such as underline and strikethrough are lost.
-   Scenes beginning with `<HTML>` or `<TEX>` are not imported.
-   All other chapters and scenes are imported, whether "used" or "unused".
-   Interspersed HTML, TEX, or RTF commands are taken over unchanged.
-   The document contains chapter `[ChID:x]` and scene `[ScID:y]` markers according to yWriter 5 standard. **Do not touch lines containing the markers** if you want to be able to reimport the document into yWriter.
-   Chapters and scenes can neither be rearranged nor deleted.
-   You can split scenes by inserting headings or a scene divider:
    -   *Heading 1* --› New chapter title (beginning a new section).
    -   *Heading 2* --› New chapter title.
    -   `###` --› Scene divider. Optionally, you can append the scene title to the scene divider.

You can write back the scene contents to the yWriter 7 project file with the [Zu yWriter exportieren](#export-to-ywriter) command.

[Zum Seitenbeginn](#top)

------------------------------------------------------------------------

## Kurze Zusammenfassung

This will load a brief synopsis with chapter and scenes titles into a
new OpenDocument text document (odt).

-   The document is placed in the same folder as the yWriter project.
-   Document's **filename**: `<yW project name_brf_synopsis>.odt`.
-   Only "normal" chapters and scenes are imported. Chapters and scenes marked "unused", "todo" or "notes" are not imported.
-   Only scenes that are intended for RTF export in yWriter will be imported.
-   Titles of scenes beginning with `<HTML>` or `<TEX>` are not imported.
-   Chapter titles appear as first level heading if the chapter is marked as beginning of a new section in yWriter. Such headings are considered as "part" headings.
-   Chapter titles appear as second level heading if the chapter is not  marked as beginning of a new section. Such headings are considered as "chapter" headings.
-   Scene titles appear as plain paragraphs.

[Zum Seitenbeginn](#top)

------------------------------------------------------------------------

## Figurenliste

This will generate a new OpenDocument spreadsheet (ods) containing a character list that can be edited in Office Calc and written back to yWriter format. File name suffix is `_charlist`.

You may change the sort order of the rows. You may also add or remove rows. New entities must get a unique ID.

You can write back the edited table to the yWriter 7 project file with the [Zu yWriter exportieren](#export-to-ywriter) command.

[Zum Seitenbeginn](#top)

------------------------------------------------------------------------

## Schauplatzliste

This will generate a new OpenDocument spreadsheet (ods) containing a location list that can be edited in Office Calc and written back to yWriter format. File name suffix is `_loclist`.

You may change the sort order of the rows. You may also add or remove rows. New entities must get a unique ID.

You can write back the edited table to the yWriter 7 project file with the [Zu yWriter exportieren](#export-to-ywriter) command.

[Zum Seitenbeginn](#top)

------------------------------------------------------------------------

## Gegenstandsliste

This will generate a new OpenDocument spreadsheet (ods) containing an item list that can be edited in Office Calc and written back to yWriter format. File name suffix is `_itemlist`.

You may change the sort order of the rows. You may also add or remove rows. New entities must get a unique ID.

You can write back the edited table to the yWriter 7 project file with the [Zu yWriter exportieren](#export-to-ywriter) command.

[Zum Seitenbeginn](#top)

------------------------------------------------------------------------

## Querverweise

This will generate a new OpenDocument text document (odt) containing navigable cross references. File name suffix is `_xref`. The cross references are:

-   Scenes per character,
-   scenes per location,
-   scenes per item,
-   scenes per tag,
-   characters per tag,
-   locations per tag,
-   items per tag.

[Zum Seitenbeginn](#top)

------------------------------------------------------------------------

## Ein Manuskript zum Export vorbereiten

Generate a new yWriter 7 project from a work in progress:

-   The new yWriter project is placed in the same folder as the document.
-   yWriter project's filename: `<document name>.yw7`.
-   Existing yWriter 7 projects will not be overwritten.

### Formatierung des Manuskripts:

A work in progress has no third level heading.

-   *Heading 1* --› New chapter title (beginning a new section).
-   *Heading 2* --› New chapter title.
-   `* * *` --› Scene divider (not needed for the first scenes in a chapter).
-   Comments right at the scene beginning are considered scene titles.
-   All other text is considered scene content.

[Zum Seitenbeginn](#top)

------------------------------------------------------------------------

## Eine Gliederung zum Export vorbereiten

Generate a new yWriter 7 project from an outline:

-   The new yWriter project is placed in the same folder as the document.
-   yWriter project's filename: `<document name>.yw7`.
-   Existing yWriter 7 projects will not be overwritten.

### Formatierung der Gliederung:

An outline has at least one third level heading.

-   *Heading 1* --› New chapter title (beginning a new section).
-   *Heading 2* --› New chapter title.
-   *Heading 3* --› New scene title.
-   All other text is considered to be chapter/scene description.

[Zum Seitenbeginn](#top)

------------------------------------------------------------------------

## Szenentrenner durch Leerzeilen ersetzen

This will replace the three-line "\* \* \*" scene dividers with single blank lines. The style of the scene-dividing lines will be changed from *Heading 4* to *Heading 5*.

[Zum Seitenbeginn](#top)

------------------------------------------------------------------------

## Absätze einrücken, die mit '&gt;' beginnen'

This will select all paragraphs that start with "&gt; " and change their paragraph style to *Quotations*.

Note: When exporting to yWriter, *Quotations* style paragraphs will automatically marked with "&gt; " at the beginning.

[Zum Seitenbeginn](#top)

------------------------------------------------------------------------

## Aufzählungsstriche ersetzen

This will select all paragraphs that start with "- " and apply a list paragraph style.

Note: When exporting to yWriter, Lists will automatically marked with "-" list strokes.

[Zum Seitenbeginn](#top)

------------------------------------------------------------------------
