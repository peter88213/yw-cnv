[Projekt-Homepage](https://peter88213.github.io/yw-cnv/) &gt;
[Haupt-Hilfeseite](help-de.html) &gt; Für Fortgeschrittene

------------------------------------------------------------------------

# yw7 Import/Export: für Fortgeschrittene

## Warnung!

Diese Funktionen sind nur für erfahrene OpenOffice/LibreOffice-Benutzer gedacht. Wenn Sie mit **Calc** und dem Konzept der **Bereiche in Writer** nicht vertraut sind, verwenden Sie die erweiterten Funktionen bitte nicht. Es besteht die Gefahr, dass das Projekt beim Zurückschreiben beschädigt wird, wenn Sie die Bereichsgrenzen in den ODT-Textdokumenten nicht einhalten oder die IDs in den ODS-Tabellen durcheinander bringen. Probieren Sie die Funktionen unbedingt zuerst mit einem Testprojekt aus.

## Befehlsreferenz

-   [Manuskript mit Kapitel- und Abschnittsbereichen](#manuskript-mit-kapitel--und-szenenbereichen)
-   [Abschnittsbeschreibungen](#abschnittsbeschreibungen)
-   [Kapitelbeschreibungen](#kapitelbeschreibungen)
-   [Teilebeschreibungen](#teilebeschreibungen)
-   [Figurenbeschreibungen](#figurenbeschreibungen)
-   [Schauplatzbeschreibungen](#schauplatzbeschreibungen)
-   [Gegenstandsbeschreibungen](#gegenstandsbeschreibungen)
-   [Abschnittsliste](#abschnittsliste)
-   [Notizen-Kapitel](#notizen-kapitel)
-   [Planungs-Kapitel](#planungs-kapitel)

------------------------------------------------------------------------

## Manuskript mit Kapitel- und Abschnittsbereichen

Dies lädt yw7-Kapitel und -Abschnitte in ein neues OpenDocument-Textdokument (odt) mit unsichtbaren Kapitel- und Abschnittsbereichen (die im Navigator zu sehen sind). Das Suffix des Dateinamens ist `_manuscript`.

- Es werden nur "normale" Kapitel und Abschnitte importiert. Kapitel und Abschnitte, die als "unbenutzt", "ToDo" oder "Notizen" gekennzeichnet sind, werden nicht importiert.
- Abschnitte, die mit `<HTML>` oder `<TEX>` beginnen, werden nicht importiert.
- Kommentare innerhalb von Abschnitten werden als Abschnittstitel zurückgeschrieben, wenn sie von `~` umgeben sind.
- Kommentare im Text, die mit Schrägstrichen und Sternchen eingeklammert sind (z.B. `/* dies ist ein Kommentar */`), werden in Autorenkommentare umgewandelt.
- Eingestreute HTML-, TEX- oder RTF-Befehle werden unverändert übernommen.
- Gobal- und Projektvariablen werden nicht aufgelöst.
- Kapitel und Abschnitte können weder neu geordnet noch gelöscht werden.
- Sie können Abschnitte durch Einfügen von Überschriften oder einem Abschnittstrenner aufteilen:
    - *Überschrift 1* → Neue Kapitelüberschrift (Beginn eines neuen Abschnitts). Sie können eine Beschreibung anfügen, durch `|` getrennt.
    - *Überschrift 2* → Neue Kapitelüberschrift. Sie können eine Beschreibung anfügen, durch `|` getrennt.
    - `###` → Abschnittstrenner. Optional können Sie den Abschnittstitel an den Abschnittstrenner anhängen. Sie können auch eine Beschreibung anfügen, durch `|` getrennt.
    - **Hinweis:** Exportieren Sie Dokumente mit aufgeteilten Abschnitten nicht mehrmals zu yw7. 
- Absätze, die mit `>` beginnen, werden als Zitate formatiert.


Mit dem Befehl [Zu yw7 exportieren](help-de.html#zu-yw7-exportieren) können Sie den Abschnittsinhalt und die Kapitel-/Abschnittsüberschriften in die yw7 Projektdatei zurückschreiben.

-   Kommentare innerhalb von Abschnitten werden als Abschnittstitel zurückgeschrieben, wenn sie von `~` umgeben sind.

[Zum Seitenbeginn](#top)

------------------------------------------------------------------------

## Abschnittsbeschreibungen

Dies erzeugt ein neues OpenDocument-Textdokument (odt), das eine **vollständige Zusammenfassung** mit Kapitelüberschriften und Abschnittsbeschreibungen enthält, die bearbeitet und in das yw7-Format zurückgeschrieben werden kann. Das Suffix des Dateinamens ist `_scenes`.

Sie können die Abschnittsbeschreibungen und die Kapitel-/Teilüberschriften mit dem Befehl [Zu yw7 exportieren](help-de.html#zu-yw7-exportieren) in die yw7 Projektdatei zurückschreiben. Kommentare direkt am Anfang der Abschnittsbeschreibungen werden als Abschnittstitel zurückgeschrieben.

[Zum Seitenbeginn](#top)

------------------------------------------------------------------------

## Kapitelbeschreibungen

Dies erzeugt ein neues OpenDocument-Textdokument (odt), das eine **Kurzzusammenfassung** mit Kapitelüberschriften und -beschreibungen enthält, die bearbeitet und in das yw7-Format zurückgeschrieben werden können. Die Endung des Dateinamens ist `_chapters`.

**Anmerkung:** Gilt nicht für Kapitel, die in yw7 mit `Dieses Kapitel beginnt einen neuen Abschnitt` gekennzeichnet sind.

Sie können die Überschriften und Beschreibungen mit dem Befehl [Zu yw7 exportieren](help-de.html#zu-yw7-exportieren) in die yw7 Projektdatei zurückschreiben.


[Zum Seitenbeginn](#top)

------------------------------------------------------------------------

## Teilebeschreibungen

Dies erzeugt ein neues OpenDocument-Textdokument (odt), das eine **sehr kurze Zusammenfassung** mit Teiltiteln und Teilbeschreibungen enthält, die bearbeitet und in das yw7-Format zurückgeschrieben werden können. Die Endung des Dateinamens ist `_parts`.

**Hinweis:** Gilt nur für Kapitel, die in yw7 mit `Dieses Kapitel beginnt einen neuen Abschnitt` gekennzeichnet sind.

Sie können die Überschriften und Beschreibungen mit dem Befehl [Zu yw7 exportieren](help-de.html#zu-yw7-exportieren) in die yw7 Projektdatei zurückschreiben.

[Zum Seitenbeginn](#top)

------------------------------------------------------------------------

## Figurenbeschreibungen

Dies erzeugt ein neues OpenDocument-Textdokument (odt) mit Figurenbeschreibungen, Lebenslauf, Zielen und Notizen, das in Office Writer bearbeitet und in das yw7-Format zurückgeschrieben werden kann. Das Suffix des Dateinamens ist `_characters`.

Sie können die Beschreibungen mit dem Befehl [Zu yw7 exportieren](help-de.html#zu-yw7-exportieren) in die yw7 Projektdatei zurückschreiben.

[Zum Seitenbeginn](#top)

------------------------------------------------------------------------

## Schauplatzbeschreibungen

Dies erzeugt ein neues OpenDocument Textdokument (odt) mit Schauplatzbeschreibungen, die in Office Writer bearbeitet und in das yw7-Format zurückgeschrieben werden können. Das Suffix des Dateinamens ist `_locations`.

Sie können die Beschreibungen mit dem Befehl [Zu yw7 exportieren](help-de.html#zu-yw7-exportieren) in die yw7 Projektdatei zurückschreiben.

[Zum Seitenbeginn](#top)

------------------------------------------------------------------------

## Gegenstandsbeschreibungen

Dies erzeugt ein neues OpenDocument-Textdokument (odt) mit Gegenstandsbeschreibungen, die in Office Writer bearbeitet und in das yw7-Format zurückgeschrieben werden können. Das Suffix des Dateinamens ist `_items`.

Sie können die Beschreibungen mit dem Befehl [Zu yw7 exportieren](help-de.html#zu-yw7-exportieren) in die yw7 Projektdatei zurückschreiben.

[Zum Seitenbeginn](#top)

------------------------------------------------------------------------

## Abschnittsliste

Dies erzeugt ein neues OpenDocument Tabellenblatt (ods), das Folgendes enthält:

-   Hyperlink zum Abschnittsbereich des Manuskripts
-   Abschnittstitel
-   Abschnittsbeschreibung
-   Tags
-   Abschnitts-Notizen
-   A/R
-   Ziel
-   Konflikt
-   Ausgang
-   Aufeinanderfolgende Abschnittsnummern
-   Gesamtwortzahl
-   Bewertung 1
-   Bewertung 2
-   Bewertung 3
-   Bewertung 4
-   Wortzahl
-   Zeichenzahl
-   Status
-   Figuren
-   Schauplätze
-   Gegenstände

Nur "normale" Abschnitte, die in *yWriter* als RTF exportiert werden würden, erhalten eine Zeile in der Abschnittsliste. Abschnitte vom Typ "Unused", "Notes" oder "ToDo" werden ausgelassen.

Das Suffix des Dateinamens ist `_scenelist`.

Sie können den Tabelleninhalt mit dem Befehl [Zu yw7 exportieren](help-de.html#zu-yw7-exportieren) in die yw7 Projektdatei zurückschreiben.

Die folgenden Spalten können in das yw7-Projekt zurückgeschrieben werden:

-   Title
-   Beschreibung
-   Tags (durch Kommas getrennt)
-   Abschnitts-Notizen
-   A/R (action/reaction scene)
-   Ziel
-   Konflikt
-   Ausgang
-   Bewertung 1
-   Bewertung 2
-   Bewertung 3
-   Bewertung 4
-   Status ('Gliederung', 'Entwurf', '1. Überarbeitung', '2. Überarbeitung', 'Fertig')

*Hinweis: In der Tabelle erscheinen die Spaltentitel in englischer Sprache. 
Das ist für das Zurückschreiben ins yw7-Format notwendig.*

[Zum Seitenbeginn](#top)

------------------------------------------------------------------------

## Notizen-Kapitel

Dies schreibt yw7 "Notizen"-Kapitel mit untergeordneten Abschnitten in ein neues OpenDocument-Textdokument (odt) mit unsichtbaren Kapitel- und Abschnittsbereichen (zu sehen im Navigator). Das Suffix des Dateinamens ist `_notes`.

- Kommentare innerhalb von Abschnitten werden als Abschnittstitel zurückgeschrieben, 
  wenn sie von `~` umgeben sind.
- Kapitel und Abschnitte können weder umgeordnet noch gelöscht werden.
- Sie können Abschnitte durch Einfügen von Überschriften oder einem Abschnittstrenner aufteilen:
    - *Überschrift 1* → Neue Kapitelüberschrift (Beginn eines neuen Abschnitts). Sie können eine Beschreibung anfügen, durch `|` getrennt.
    - *Überschrift 2* → Neue Kapitelüberschrift. Sie können eine Beschreibung anfügen, durch `|` getrennt.
    - `###` → Abschnittstrenner. Optional können Sie den Abschnittstitel an den Abschnittstrenner anhängen. Sie können auch eine Beschreibung anfügen, durch `|` getrennt.
    - **Hinweis:** Exportieren Sie Dokumente mit aufgeteilten Abschnitten nicht mehrmals zu yw7. 
- Absätze, die mit `>` beginnen, werden als Zitate formatiert.

[Zum Seitenbeginn](#top)

------------------------------------------------------------------------

## Planungs-Kapitel

Dies schreibt yw7 "Todo"-Kapitel mit untergeordneten Abschnitten in ein neues 
OpenDocument-Textdokument (odt) mit unsichtbaren Kapitel- und Abschnittsbereichen (die im Navigator zu sehen sind). Die Dateinamenserweiterung ist `_todo`.

- Kommentare innerhalb von Abschnitten werden als Abschnittstitel zurückgeschrieben,
  wenn sie von `~` umgeben sind.
- Kapitel und Abschnitte können weder umgeordnet noch gelöscht werden.
- Sie können Abschnitte durch Einfügen von Überschriften oder einem Abschnittstrenner aufteilen:
    - *Überschrift 1* → Neue Kapitelüberschrift (Beginn eines neuen Abschnitts). Sie können eine Beschreibung anfügen, durch `|` getrennt.
    - *Überschrift 2* → Neue Kapitelüberschrift. Sie können eine Beschreibung anfügen, durch `|` getrennt.
    - `###` → Abschnittstrenner. Optional können Sie den Abschnittstitel an den Abschnittstrenner anhängen. Sie können auch eine Beschreibung anfügen, durch `|` getrennt.
    - **Hinweis:** Exportieren Sie Dokumente mit aufgeteilten Abschnitten nicht mehrmals zu yw7. 
- Absätze, die mit `>` beginnen, werden als Zitate formatiert.

[Zum Seitenbeginn](#top)

------------------------------------------------------------------------
