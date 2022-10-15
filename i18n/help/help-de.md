[Projekt-Homepage](https://peter88213.github.io/yw-cnv/) &gt;
Haupt-Hilfeseite

------------------------------------------------------------------------

# yWriter Import/Export


## Befehlsreferenz

### "Datei"-Menü

-   [Zu yWriter exportieren](#zu-ywriter-exportieren)
-   [Von yWriter importieren](#von-ywriter-importieren)
-   [Von yWriter zum Korrekturlesen importieren](#von-ywriter-zum-korrekturlesen-importieren)
-   [Kurze Zusammenfassung](#kurze-zusammenfassung)
-   [Figurenliste](#figurenliste)
-   [Schauplatzliste](#schauplatzliste)
-   [Gegenstandsliste](#gegenstandsliste)
-   [Querverweise](#querverweise)
-   [Für Fortgeschrittene](help-adv-de.html)

### "Format"-Menü

-   [Szenentrenner durch Leerzeilen
    ersetzen](#szenentrenner-durch-leerzeilen-ersetzen)
-   [Absätze einrücken, die mit '> ' beginnen'](#absätze-einrücken-die-mit-beginnen)
-   [Aufzählungsstriche ersetzen](#aufzählungsstriche-ersetzen)

## So geht's

-   [Ein Manuskript zum Export vorbereiten](#ein-manuskript-zum-export-vorbereiten)
-   [Eine Gliederung zum Export vorbereiten](#eine-gliederung-zum-export-vorbereiten)

## Allgemeines

-   [Umgang mit der Sprache des Dokuments](#umgang-mit-der-sprache-des-dokuments)


------------------------------------------------------------------------

## Zu yWriter exportieren

Dadurch wird der Inhalt des Dokuments in die yWriter-Projektdatei zurückgeschrieben.

-   Stellen Sie sicher, dass Sie den Dateinamen eines generierten Dokuments vor dem Zurückschreiben in das yWriter-Formanicht ändern.
-   Das zu überschreibende yWriter 7 Projekt muss sich im selben Ordner befinden wie das Dokument.
-   Wenn der Dateiname des Dokuments kein Suffix hat, dient das Dokument als [Manuskript in Arbeit](#ein-manuskript-zum-export-vorbereiten) oder eine [Gliederung](#eine-gliederung-zum-export-vorbereiten) zum Exportieren in ein neu zu erstellendes yWriter-Projekt. **Hinweis:** Bestehende  yWriter-Projekte werden nicht überschrieben.

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
- Kommentare im Text, die mit Schrägstrichen und Sternchen eingeklammert sind (z.B. `/* das ist ein Kommentar */`), werden in Autorenkommentare umgewandelt.
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

Dies lädt yWriter 7-Kapitel und -Szenen in ein neues OpenDocument-Textdokument (odt) mit Kapitel- und Szenenmarkierungen. Das Suffix des Dateinamens ist `_proof`.

- Das Probedokument wird im selben Ordner wie das yWriter-Projekt abgelegt.
- Der Dateiname des Dokuments: `<yW Projektname>_proof.odt`.
- Textauszeichnung: Fett und kursiv werden unterstützt. Andere Hervorhebungen wie Unterstreichen und Durchstreichen gehen verloren.
- Szenen, die mit `<HTML>` oder `<TEX>` beginnen, werden nicht importiert.
- Alle anderen Kapitel und Szenen werden importiert, egal ob "benutzt" oder "unbenutzt".
- Eingestreute HTML-, TEX- oder RTF-Befehle werden unverändert übernommen.
- Das Dokument enthält Kapitel `[ChID:x]` und Szenen `[ScID:y]` Markierungen entsprechend dem yWriter 5 Standard. **Ändern Sie die Zeilen mit den Markierungen nicht**, wenn Sie das Dokument wieder in yWriter importieren wollen.
- Kapitel und Szenen können weder umgeordnet noch gelöscht werden.
- Sie können Szenen aufteilen, indem Sie Überschriften oder einen Szenentrenner einfügen:
    - *Überschrift 1* → Neue Kapitelüberschrift (Beginn eines neuen Abschnitts).
    - *Überschrift 2* → Neuer Kapiteltitel.
    - `###` → Szenentrenner. Optional können Sie den Szenentitel an den Szenentrenner anhängen.
- Absätze, die mit `>` beginnen, werden als Zitate formatiert.

Sie können den Inhalt der Szene mit dem Befehl [Zu yWriter exportieren](#zu-ywriter-exportieren) in die yWriter 7 Projektdatei zurückschreiben.

[Zum Seitenbeginn](#top)

------------------------------------------------------------------------

## Kurze Zusammenfassung

Dies lädt eine kurze Zusammenfassung mit Kapitel- und Szenenüberschriften in ein
neues OpenDocument-Textdokument (odt).

- Das Dokument wird im selben Ordner wie das yWriter-Projekt abgelegt.
- Der **Dateiname** des Dokuments: `<yW-Projektname_brf_synopsis>.odt`.
- Es werden nur "normale" Kapitel und Szenen importiert. Kapitel und Szenen, die als "unbenutzt", "ToDo" oder "Notizen" gekennzeichnet sind, werden nicht importiert.
- Nur Szenen, die für den RTF-Export in yWriter vorgesehen sind, werden importiert.
- Titel von Szenen, die mit `<HTML>` oder `<TEX>` beginnen, werden nicht importiert.
- Kapitelüberschriften erscheinen als Überschrift der ersten Ebene, wenn das Kapitel in yWriter als Beginn eines neuen Abschnitts markiert ist. Solche Überschriften werden als "Teil"-Überschriften betrachtet.
- Kapitelüberschriften erscheinen als Überschriften der zweiten Ebene, wenn das Kapitel nicht als Beginn eines neuen Abschnitts markiert ist. Solche Überschriften werden als "Kapitel"-Überschriften betrachtet.
- Szenentitel erscheinen als einfache Absätze.

[Zum Seitenbeginn](#top)

------------------------------------------------------------------------

## Figurenliste

Damit wird ein neues OpenDocument Tabellenblatt (ods) erzeugt, das eine Figurenliste enthält, die in Office Calc bearbeitet und in das yWriter-Format zurückgeschrieben werden kann. Das Suffix des Dateinamens ist `_charlist`.

Sie können die Sortierreihenfolge der Zeilen ändern. Sie können auch Zeilen hinzufügen oder entfernen. Neue Einträge müssen eine eindeutige ID erhalten.

Sie können die bearbeitete Tabelle mit dem Befehl [Zu yWriter exportieren](#zu-ywriter-exportieren) in die yWriter 7 Projektdatei zurückschreiben.

[Zum Seitenbeginn](#top)

------------------------------------------------------------------------

## Schauplatzliste

Dies erzeugt ein neues OpenDocument Tabellenblatt (ods), das eine Schauplatzliste enthält, die in Office Calc bearbeitet und in das yWriter Format zurückgeschrieben werden kann. Das Suffix des Dateinamens ist `_loclist`.

Sie können die Sortierreihenfolge der Zeilen ändern. Sie können auch Zeilen hinzufügen oder entfernen. Neue Einträge müssen eine eindeutige ID erhalten.

Sie können die bearbeitete Tabelle mit dem Befehl [Zu yWriter exportieren](#zu-ywriter-exportieren) in die yWriter 7 Projektdatei zurückschreiben.

[Zum Seitenbeginn](#top)

------------------------------------------------------------------------

## Gegenstandsliste

Dies erzeugt ein neues OpenDocument Spreadsheet (ods), das eine Gegenstandsliste enthält, die in Office Calc bearbeitet und in das yWriter-Format zurückgeschrieben werden kann. Das Suffix des Dateinamens ist `_itemlist`.

Sie können die Sortierreihenfolge der Zeilen ändern. Sie können auch Zeilen hinzufügen oder entfernen. Neue Einträge müssen eine eindeutige ID erhalten.

Sie können die bearbeitete Tabelle mit dem Befehl [Zu yWriter exportieren](#zu-ywriter-exportieren) in die yWriter 7 Projektdatei zurückschreiben.

[Zum Seitenbeginn](#top)

------------------------------------------------------------------------

## Querverweise

Dadurch wird ein neues OpenDocument-Textdokument (odt) erzeugt, das navigierbare Querverweise enthält. Das Suffix des Dateinamens ist `_xref`. Die Querverweise sind:

- Szenen pro Figur,
- Szenen pro Schauplatz,
- Szenen pro Gegenstand,
- Szenen pro Tag,
- Figuren pro Tag,
- Schauplätze pro Tag,
- Gegenstände pro Tag.

[Zum Seitenbeginn](#top)

------------------------------------------------------------------------

## Ein Manuskript zum Export vorbereiten

Erzeugt ein neues yWriter 7-Projekt aus einem in Arbeit befindlichen Dokument:

- Das neue yWriter-Projekt wird in demselben Ordner wie das Dokument abgelegt.
- Dateiname des yWriter-Projekts: `<Dokumentname>.yw7`.
- Bestehende yWriter 7-Projekte werden nicht überschrieben.

### Formatierung des Manuskripts:

Ein in Arbeit befindliches Dokument hat keine Überschrift auf der dritten Ebene.

- *Überschrift 1* → Neue Kapitelüberschrift (Beginn eines neuen Abschnitts).
- *Überschrift 2* → Neuer Kapiteltitel.
- `* * *` → Szenentrenner (nicht erforderlich für die ersten Szenen eines Kapitels).
- Kommentare direkt am Anfang einer Szene gelten als Szenentitel.
- Alle anderen Texte gelten als Szeneninhalt.

[Zum Seitenbeginn](#top)

------------------------------------------------------------------------

## Eine Gliederung zum Export vorbereiten

Erzeugt ein neues yWriter 7-Projekt aus einer Gliederung:

- Das neue yWriter-Projekt wird in demselben Ordner wie das Dokument abgelegt.
- Dateiname des yWriter-Projekts: `<Dokumentname>.yw7`.
- Bestehende yWriter 7-Projekte werden nicht überschrieben.

### Formatierung der Gliederung:

Eine Gliederung hat mindestens eine Überschrift auf der dritten Ebene.

- *Überschrift 1* → Neue Kapitelüberschrift (Beginn eines neuen Abschnitts).
- *Überschrift 2* → Neue Kapitelüberschrift.
- *Überschrift 3* → Neuer Szenentitel.
- Alle anderen Texte werden als Kapitel-/Szenenbeschreibung betrachtet.

[Zum Seitenbeginn](#top)

------------------------------------------------------------------------

## Szenentrenner durch Leerzeilen ersetzen

Dadurch werden die dreizeiligen "\* \* \*"-Szenentrennlinien durch einzelne Leerzeilen ersetzt. Der Stil der szenenunterteilenden Zeilen wird von *Überschrift 4* auf *Überschrift 5* geändert.

[Zum Seitenbeginn](#top)

------------------------------------------------------------------------

## Absätze einrücken, die mit '&gt;' beginnen'

Dadurch werden alle Absätze ausgewählt, die mit "&gt; " beginnen, und ihr Absatzstil wird in *Quotations* geändert.

Hinweis: Beim Export in yWriter werden Absätze im Stil *Quotations* automatisch mit "&gt; " am Anfang markiert.

[Zum Seitenbeginn](#top)

------------------------------------------------------------------------

## Aufzählungsstriche ersetzen

Dadurch werden alle Absätze, die mit "-" beginnen, ausgewählt und mit einem Listenabsatzstil versehen.

Hinweis: Beim Export in yWriter werden Listen automatisch mit "-"-Listenstrichen markiert.

[Zum Seitenbeginn](#top)

------------------------------------------------------------------------

## Allgemeines

### Umgang mit der Sprache des Dokuments

ODF-Dokumenten ist im allgemeinen eine Sprache zugeordnet, welche die Rechtschreibprüfung und länderspezifische Zeichenersetzungen bestimmt. Außerdem kann man in Office Writer Textpassagen von der Dokumentsprache abweichende Sprachen zuordnen, um Fremdsprachengebrauch zu markieren oder die Rechtschreibprüfung auszusetzen. 

#### Dokument insgesamt

- Wenn bei der Konvertierung in das yw7-Format eine Dokument-Sprache (Sprachencode gem. ISO 639-1 und Ländercode gem. ISO 3166-2) im Ausgangsdokument erkannt wird, werden diese Codes als yWriter-Projektvariablen eingetragen.

- Wenn bei der Konvertierung aus dem yw7-Format Sprach- und Ländercode als Projektvariablen existieren, werden sie ins erzeugte ODF-Dokument eingetragen. 

- Wenn bei der Konvertierung aus dem yw7-Format keine Sprach- und Ländercode als Projektvariablen existieren, werden Sprach- und Ländercode aus den Betriebssystemeinstellungen ins erzeugte ODF-Dokument eingetragen. 

- Die Sprach- und Ländercodes werden oberflächlich geprüft. Wenn sie offensichtlich nicht den ISO-Standards entsprechen, werden sie durch die Werte für "Keine Sprache" ersetzt. Das sind:
    - Language = zxx
    - Country = none
    
#### Textpassagen in Szenen  

Wenn bei der Konvertierung in das yw7-Format Textauszeichnungen für andere Sprachen erkannt werden, werden diese umgewandelt und in die yWriter-Szene übernommen. 

Das sieht dann beispielsweise so aus:

`xxx xxxx [lang=de-CH]yyy yyyy yyyy[/lang=de-CH] xxx xxx`

Damit diese Textauszeichnungen in *yWriter* nicht stören, werden sie automatisch als Projektvariablen eingetragen, und zwar so, dass *yWriter* sie beim Dokumentenexport als HTML-Anweisungen interpretiert. 

Für das oben gezeigte Beispiel sieht die Projektvariablen-Definition für das öffnende Tag so aus:

- *Variable Name:* `lang=de-CH` 
- *Value/Text:* `<HTM <SPAN LANG="de-CH"> /HTM>`

Der Sinn der Sache liegt darin, dass solche Sprachenzuweisungen auch bei mehrmaligem Konvertieren in beide Richtungen erhalten bleiben, also im ODT-Dokument immer für die Rechtschreibprüfung wirksam sind.

Es wird empfohlen, solche Auszeichnungen nicht in yWriter zu verändern, um ungewollte Verschachtelungen und unterbrochene Umschließungen zu vermeiden. 

[Zum Seitenbeginn](#top)

------------------------------------------------------------------------

