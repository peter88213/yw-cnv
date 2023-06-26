[Projekt-Homepage](https://peter88213.github.io/yw-cnv/) &gt;
Haupt-Hilfeseite

------------------------------------------------------------------------

# yw7 Import/Export


## Befehlsreferenz

### "Datei"-Menü

-   [Zu yw7 exportieren](#zu-yw7-exportieren)
-   [Von yw7 importieren](#von-yw7-importieren)
-   [Von yw7 zum Korrekturlesen importieren](#von-yw7-zum-korrekturlesen-importieren)
-   [Kurze Zusammenfassung](#kurze-zusammenfassung)
-   [Figurenliste](#figurenliste)
-   [Schauplatzliste](#schauplatzliste)
-   [Gegenstandsliste](#gegenstandsliste)
-   [Querverweise](#querverweise)
-   [Für Fortgeschrittene](help-adv-de.html)

### "Format"-Menü

-   [Abschnittstrenner durch Leerzeilen
    ersetzen](#abschnittstrenner-durch-leerzeilen-ersetzen)
-   [Absätze einrücken, die mit '> ' beginnen'](#absätze-einrücken-die-mit-beginnen)
-   [Aufzählungsstriche ersetzen](#aufzählungsstriche-ersetzen)

------------------------------------------------------------------------

## Allgemeines


### Zur Textformatierung

Es wird davon ausgegangen, dass für einen fiktionalen Text nur sehr wenige Arten von Textauszeichnungen erforderlich sind:

- *Betont* (üblicherweise in Kursivschrift dargestellt).
- *Stark betont* (üblicherweise in Großbuchstaben dargestellt).
- *Zitat* (Absatz optisch vom Fließtext unterschieden).

Beim Importieren aus dem yw7-Format ersetzt der Konverter diese Formatierungen wie folgt: 

- In yw7 *kursiv* gesetzter Text wird als *Betont* formatiert.
- In yw7 *fett* gesetzter Text wird als *Stark betont* formatiert. 
- Absätze, die in yw7 mit `"> "` beginnen, werden als *Zitat* formatiert. 

Beim Exportieren in das yw7-Format gilt das Umgekehrte.

### Zur Sprache des Dokuments

ODF-Dokumenten ist im allgemeinen eine Sprache zugeordnet, welche die Rechtschreibprüfung und länderspezifische Zeichenersetzungen bestimmt. Außerdem kann man in Office Writer Textpassagen von der Dokumentsprache abweichende Sprachen zuordnen, um Fremdsprachengebrauch zu markieren oder die Rechtschreibprüfung auszusetzen. 

#### Dokument insgesamt

- Wenn bei der Konvertierung in das yw7-Format eine Dokument-Sprache (Sprachencode gem. ISO 639-1 und Ländercode gem. ISO 3166-2) im Ausgangsdokument erkannt wird, werden diese Codes als yw7-Projektvariablen eingetragen.

- Wenn bei der Konvertierung aus dem yw7-Format Sprach- und Ländercode als Projektvariablen existieren, werden sie ins erzeugte ODF-Dokument eingetragen. 

- Wenn bei der Konvertierung aus dem yw7-Format keine Sprach- und Ländercode als Projektvariablen existieren, werden Sprach- und Ländercode aus den Betriebssystemeinstellungen ins erzeugte ODF-Dokument eingetragen. 

- Die Sprach- und Ländercodes werden oberflächlich geprüft. Wenn sie offensichtlich nicht den ISO-Standards entsprechen, werden sie durch die Werte für "Keine Sprache" ersetzt. Das sind:
    - Language = zxx
    - Country = none
    
#### Textpassagen in Abschnitten

Wenn bei der Konvertierung in das yw7-Format Textauszeichnungen für andere Sprachen erkannt werden, werden diese umgewandelt und in drn yw7-Abschnitt übernommen. 

Das sieht dann beispielsweise so aus:

`xxx xxxx [lang=de-CH]yyy yyyy yyyy[/lang=de-CH] xxx xxx`

Damit diese Textauszeichnungen in *yWriter* nicht stören, werden sie automatisch als Projektvariablen eingetragen, und zwar so, dass *yWriter* sie beim Dokumentenexport als HTML-Anweisungen interpretiert. 

Für das oben gezeigte Beispiel sieht die Projektvariablen-Definition für das öffnende Tag so aus:

- *Variable Name:* `lang=de-CH` 
- *Value/Text:* `<HTM <SPAN LANG="de-CH"> /HTM>`

Der Sinn der Sache liegt darin, dass solche Sprachenzuweisungen auch bei mehrmaligem Konvertieren in beide Richtungen erhalten bleiben, also im ODT-Dokument immer für die Rechtschreibprüfung wirksam sind.

Es wird empfohlen, solche Auszeichnungen nicht in *yWriter* zu verändern, um ungewollte Verschachtelungen und unterbrochene Umschließungen zu vermeiden. 


## So wird's gemacht

## Ein Manuskript zum Export vorbereiten

Ein in Arbeit befindliches Dokument hat keine Überschrift auf der dritten Ebene.

- *Überschrift 1* → Neue Kapitelüberschrift (Beginn eines neuen Abschnitts).
- *Überschrift 2* → Neuer Kapiteltitel.
- `* * *` → Abschnittstrenner (nicht erforderlich für die ersten Abschnitte eines Kapitels).
- Kommentare direkt am Anfang eines Abschnitts gelten als Abschnittstitel.
- Alle anderen Texte gelten als Abschnittsinhalt.

## Eine Gliederung zum Export vorbereiten

Eine Gliederung hat mindestens eine Überschrift auf der dritten Ebene.

- *Überschrift 1* → Neue Kapitelüberschrift (Beginn eines neuen Abschnitts).
- *Überschrift 2* → Neue Kapitelüberschrift.
- *Überschrift 3* → Neuer Abschnittstitel.
- Alle anderen Texte werden als Kapitel-/Abschnittsbeschreibung betrachtet.

[Zum Seitenbeginn](#top)

------------------------------------------------------------------------

## Zu yw7 exportieren

Dadurch wird der Inhalt des Dokuments in die yw7-Projektdatei zurückgeschrieben.

-   Stellen Sie sicher, dass Sie den Dateinamen eines generierten Dokuments vor dem Zurückschreiben 
in das yw7-Format nicht ändern.
-   Das zu überschreibende yw7 Projekt muss sich im selben Ordner befinden wie das Dokument.
-   Wenn der Dateiname des Dokuments kein Suffix hat, dient das Dokument als [Manuskript in Arbeit](#ein-manuskript-zum-export-vorbereiten) oder eine [Gliederung](#eine-gliederung-zum-export-vorbereiten) zum Exportieren in ein neu zu erstellendes yw7-Projekt. **Hinweis:** Bestehende  yw7-Projekte werden nicht überschrieben.


[Zum Seitenbeginn](#top)

------------------------------------------------------------------------

## Von yw7 importieren

Dadurch werden yw7-Kapitel und -Abschnitte in ein neues OpenDocument Textdokument (odt) importiert.

- Das Dokument wird im selben Ordner wie das yw7-Projekt abgelegt.
- Der **Dateiname** des Dokuments: `<yW-Projektname>.odt`.
- Textauszeichnung: Fett und kursiv werden unterstützt. Andere Hervorhebungen wie Unterstreichen und Durchstreichen gehen verloren.
- Es werden nur "normale" Kapitel und Abschnitte importiert. Kapitel und Abschnitte, die als "unbenutzt", "ToDo" oder "Notizen" gekennzeichnet sind, werden nicht importiert.
- Nur Abschnitte, die in *yWriter* für den RTF-Export vorgesehen sind, werden importiert.
- Abschnitte, die mit `<HTML>` oder `<TEX>` beginnen, werden nicht importiert.
- Kommentare im Text, die mit Schrägstrichen und Sternchen eingeklammert sind (z.B. `/* das ist ein Kommentar */`), werden in Autorenkommentare umgewandelt.
- Gekennzeichnete Kommentare (wie `/*@en Das ist eine Endnote. */`) 
  werden zu Fußnoten oder Endnoten konvertiert. Kennzeichnungen:
    - `@fn*` -- Einfache Fußnote mit Sternchen, 
    - `@fn` -- Nummerierte Fußnote.
    - `@en` -- Nummerierte Endnote.  
- Eingestreute HTML-, TEX- oder RTF-Befehle werden entfernt.
- Grundlegende Variablen und Projektvariablen werden nicht aufgelöst.
- Kapitelüberschriften erscheinen als Überschriften der ersten Ebene, wenn das Kapitel in yw7 als Beginn eines neuen Abschnitts markiert ist. Solche Überschriften werden als "Teil"-Überschriften betrachtet.
- Kapitelüberschriften erscheinen als Überschriften der zweiten Ebene, wenn das Kapitel nicht als Beginn eines neuen Abschnitts markiert ist. Solche Überschriften werden als "Kapitel"-Überschriften betrachtet.
- Abschnittstitel erscheinen als navigierbare Kommentare, die an den Anfang des Abschnitts gepinnt sind.
- Abschnitte werden durch `* * *` getrennt. Die erste Zeile wird nicht eingerückt. Mit dem Menübefehl **Format &gt; Abschnittstrenner durch Leerzeilen ersetzen** können Sie die Abschnittstrenner durch Leerzeilen ersetzen.
- Ab dem zweiten Absatz beginnen Absätze mit der Einrückung der ersten Zeile.
- Abschnitte, die in yw7 mit "an vorherigen Abschnitt anhängen" gekennzeichnet sind, erscheinen wie durchgehende Absätze.
- Absätze, die mit `>` beginnen, werden als Zitate formatiert.

[Zum Seitenbeginn](#top)

------------------------------------------------------------------------

## Von yw7 zum Korrekturlesen importieren

Dies lädt yw7-Kapitel und -Abschnitte in ein neues OpenDocument-Textdokument (odt) mit sichtbaren Abschnittsmarkierungen. Das Suffix des Dateinamens ist `_proof`.

- Es werden nur "normale" Kapitel und Abschnitte importiert. Kapitel und Abschnitte, die als "unbenutzt", "ToDo" oder "Notizen" gekennzeichnet sind, werden nicht importiert.
- Abschnitte, die mit `<HTML>` oder `<TEX>` beginnen, werden nicht importiert.
- Eingestreute HTML-, TEX- oder RTF-Befehle werden unverändert übernommen.
- Das Dokument enthält Abschnittsmarkierungen `[ScID:x]`. **Ändern Sie die Zeilen mit den Markierungen nicht**, wenn Sie das Dokument wieder in yw7 importieren wollen.
- Kapitel und Abschnitte können weder umgeordnet noch gelöscht werden.
- Sie können Abschnitte aufteilen, indem Sie Überschriften oder einen Abschnittstrenner einfügen:
    - *Überschrift 1* → Neue Kapitelüberschrift (Beginn eines neuen Abschnitts). Sie können eine Beschreibung anfügen, durch `|` getrennt.
    - *Überschrift 2* → Neuer Kapiteltitel. Sie können eine Beschreibung anfügen, durch `|` getrennt.
    - `###` → Abschnittstrenner. Optional können Sie den Abschnittstitel an den Abschnittstrenner anhängen. Sie können auch eine Beschreibung anfügen, durch `|` getrennt.
- Absätze, die mit `>` beginnen, werden als Zitate formatiert.
- Textauszeichnung: Fett und kursiv werden unterstützt. Andere Hervorhebungen wie Unterstreichen und Durchstreichen gehen verloren.

Sie können den Inhalt des Abschnitts mit dem Befehl [Zu yw7 exportieren](#zu-yw7-exportieren) in die yw7 Projektdatei zurückschreiben.

[Zum Seitenbeginn](#top)

------------------------------------------------------------------------

## Kurze Zusammenfassung

Dies lädt eine kurze Zusammenfassung mit Kapitel- und Abschnittsüberschriften in ein
neues OpenDocument-Textdokument (odt).

- Das Dokument wird im selben Ordner wie das yw7-Projekt abgelegt.
- Der **Dateiname** des Dokuments: `<yW-Projektname_brf_synopsis>.odt`.
- Es werden nur "normale" Kapitel und Abschnitte importiert. Kapitel und Abschnitte, die als "unbenutzt", "ToDo" oder "Notizen" gekennzeichnet sind, werden nicht importiert.
- Nur Abschnitte, die in *yWriter* für den RTF-Export vorgesehen sind, werden importiert.
- Titel von Abschnitten, die mit `<HTML>` oder `<TEX>` beginnen, werden nicht importiert.
- Kapitelüberschriften erscheinen als Überschrift der ersten Ebene, wenn das Kapitel in yw7 als Beginn eines neuen Abschnitts markiert ist. Solche Überschriften werden als "Teil"-Überschriften betrachtet.
- Kapitelüberschriften erscheinen als Überschriften der zweiten Ebene, wenn das Kapitel nicht als Beginn eines neuen Abschnitts markiert ist. Solche Überschriften werden als "Kapitel"-Überschriften betrachtet.
- Abschnittstitel erscheinen als einfache Absätze.

[Zum Seitenbeginn](#top)

------------------------------------------------------------------------

## Figurenliste

Damit wird ein neues OpenDocument Tabellenblatt (ods) erzeugt, das eine Figurenliste enthält, die in Office Calc bearbeitet und in das yw7-Format zurückgeschrieben werden kann. Das Suffix des Dateinamens ist `_charlist`.

Sie können die Sortierreihenfolge der Zeilen ändern. Sie können auch Zeilen hinzufügen oder entfernen. Neue Einträge müssen eine eindeutige ID erhalten.

Sie können die bearbeitete Tabelle mit dem Befehl [Zu yw7 exportieren](#zu-yw7-exportieren) in die yw7 Projektdatei zurückschreiben.

[Zum Seitenbeginn](#top)

------------------------------------------------------------------------

## Schauplatzliste

Dies erzeugt ein neues OpenDocument Tabellenblatt (ods), das eine Schauplatzliste enthält, die in Office Calc bearbeitet und in das yw7 Format zurückgeschrieben werden kann. Das Suffix des Dateinamens ist `_loclist`.

Sie können die Sortierreihenfolge der Zeilen ändern. Sie können auch Zeilen hinzufügen oder entfernen. Neue Einträge müssen eine eindeutige ID erhalten.

Sie können die bearbeitete Tabelle mit dem Befehl [Zu yw7 exportieren](#zu-yw7-exportieren) in die yw7 Projektdatei zurückschreiben.

[Zum Seitenbeginn](#top)

------------------------------------------------------------------------

## Gegenstandsliste

Dies erzeugt ein neues OpenDocument Spreadsheet (ods), das eine Gegenstandsliste enthält, die in Office Calc bearbeitet und in das yw7-Format zurückgeschrieben werden kann. Das Suffix des Dateinamens ist `_itemlist`.

Sie können die Sortierreihenfolge der Zeilen ändern. Sie können auch Zeilen hinzufügen oder entfernen. Neue Einträge müssen eine eindeutige ID erhalten.

Sie können die bearbeitete Tabelle mit dem Befehl [Zu yw7 exportieren](#zu-yw7-exportieren) in die yw7 Projektdatei zurückschreiben.

[Zum Seitenbeginn](#top)

------------------------------------------------------------------------

## Querverweise

Dadurch wird ein neues OpenDocument-Textdokument (odt) erzeugt, das navigierbare Querverweise enthält. Das Suffix des Dateinamens ist `_xref`. Die Querverweise sind:

- Abschnitte pro Figur,
- Abschnitte pro Schauplatz,
- Abschnitte pro Gegenstand,
- Abschnitte pro Tag,
- Figuren pro Tag,
- Schauplätze pro Tag,
- Gegenstände pro Tag.

[Zum Seitenbeginn](#top)

------------------------------------------------------------------------

## Abschnittstrenner durch Leerzeilen ersetzen

Dadurch werden die dreizeiligen "\* \* \*"-Abschnittstrennlinien durch einzelne Leerzeilen ersetzt. Der Stil der szenenunterteilenden Zeilen wird von *Überschrift 4* auf *Überschrift 5* geändert.

[Zum Seitenbeginn](#top)

------------------------------------------------------------------------

## Absätze einrücken, die mit '&gt;' beginnen'

Dadurch werden alle Absätze ausgewählt, die mit "&gt; " beginnen, und ihr Absatzstil wird in *Quotations* geändert.

Hinweis: Beim Export in yw7 werden Absätze im Stil *Quotations* automatisch mit "&gt; " am Anfang markiert.

[Zum Seitenbeginn](#top)

------------------------------------------------------------------------

## Aufzählungsstriche ersetzen

Dadurch werden alle Absätze, die mit "-" beginnen, ausgewählt und mit einem Listenabsatzstil versehen.

Hinweis: Beim Export in yw7 werden Listen automatisch mit "-"-Listenstrichen markiert.

[Zum Seitenbeginn](#top)

------------------------------------------------------------------------

