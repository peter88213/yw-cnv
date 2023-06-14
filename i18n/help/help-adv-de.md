[Projekt-Homepage](https://peter88213.github.io/yw-cnv/) &gt;
[Haupt-Hilfeseite](help-de.html) &gt; Für Fortgeschrittene

------------------------------------------------------------------------

# yw7 Import/Export: für Fortgeschrittene

## Warnung!

Diese Funktionen sind nur für erfahrene OpenOffice/LibreOffice-Benutzer gedacht. Wenn Sie mit **Calc** und dem Konzept der **Bereiche in Writer** nicht vertraut sind, verwenden Sie die erweiterten Funktionen bitte nicht. Es besteht die Gefahr, dass das Projekt beim Zurückschreiben beschädigt wird, wenn Sie die Bereichsgrenzen in den ODT-Textdokumenten nicht einhalten oder die IDs in den ODS-Tabellen durcheinander bringen. Probieren Sie die Funktionen unbedingt zuerst mit einem Testprojekt aus.

## Befehlsreferenz

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

Dies lädt yw7-Kapitel und -Szenen in ein neues OpenDocument-Textdokument (odt) mit unsichtbaren Kapitel- und Szenebereichen (die im Navigator zu sehen sind). Das Suffix des Dateinamens ist `_manuscript`.

- Es werden nur "normale" Kapitel und Szenen importiert. Kapitel und Szenen, die als "unbenutzt", "ToDo" oder "Notizen" gekennzeichnet sind, werden nicht importiert.
- Szenen, die mit `<HTML>` oder `<TEX>` beginnen, werden nicht importiert.
- Kommentare innerhalb von Szenen werden als Szenentitel zurückgeschrieben, wenn sie von `~` umgeben sind.
- Kommentare im Text, die mit Schrägstrichen und Sternchen eingeklammert sind (z.B. `/* dies ist ein Kommentar */`), werden in Autorenkommentare umgewandelt.
- Eingestreute HTML-, TEX- oder RTF-Befehle werden unverändert übernommen.
- Gobal- und Projektvariablen werden nicht aufgelöst.
- Kapitel und Szenen können weder neu geordnet noch gelöscht werden.
- Sie können Szenen durch Einfügen von Überschriften oder einem Szenentrenner aufteilen:
    - *Überschrift 1* → Neue Kapitelüberschrift (Beginn eines neuen Abschnitts). Sie können eine Beschreibung anfügen, durch `|` getrennt.
    - *Überschrift 2* → Neue Kapitelüberschrift. Sie können eine Beschreibung anfügen, durch `|` getrennt.
    - `###` → Szenentrenner. Optional können Sie den Szenentitel an den Szenentrenner anhängen. Sie können auch eine Beschreibung anfügen, durch `|` getrennt.
    - **Hinweis:** Exportieren Sie Dokumente mit aufgeteilten Szenen nicht mehrmals zu yw7. 
- Absätze, die mit `>` beginnen, werden als Zitate formatiert.


Mit dem Befehl [Zu yw7 exportieren](help-de.html#zu-yw7-exportieren) können Sie den Szeneninhalt und die Kapitel-/Abschnittsüberschriften in die yw7 Projektdatei zurückschreiben.

-   Kommentare innerhalb von Szenen werden als Szenentitel zurückgeschrieben, wenn sie von `~` umgeben sind.

[Zum Seitenbeginn](#top)

------------------------------------------------------------------------

## Szenebeschreibungen

Dies erzeugt ein neues OpenDocument-Textdokument (odt), das eine **vollständige Zusammenfassung** mit Kapitelüberschriften und Szenenbeschreibungen enthält, die bearbeitet und in das yw7-Format zurückgeschrieben werden kann. Das Suffix des Dateinamens ist `_scenes`.

Sie können die Szenenbeschreibungen und die Kapitel-/Teilüberschriften mit dem Befehl [Zu yw7 exportieren](help-de.html#zu-yw7-exportieren) in die yw7 Projektdatei zurückschreiben. Kommentare direkt am Anfang der Szenenbeschreibungen werden als Szenentitel zurückgeschrieben.

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

## Szenenliste

Dies erzeugt ein neues OpenDocument Tabellenblatt (ods), das Folgendes enthält:

-   Hyperlink zum Szenenbereich des Manuskripts
-   Szenentitel
-   Szenenbeschreibung
-   Tags
-   Szenen-Notizen
-   A/R
-   Ziel
-   Konflikt
-   Ausgang
-   Aufeinanderfolgende Szenennummern
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

Nur "normale" Szenen, die in *yWriter* als RTF exportiert werden würden, erhalten eine Zeile in der Szenenliste. Szenen vom Typ "Unused", "Notes" oder "ToDo" werden ausgelassen.

Das Suffix des Dateinamens ist `_scenelist`.

Sie können den Tabelleninhalt mit dem Befehl [Zu yw7 exportieren](help-de.html#zu-yw7-exportieren) in die yw7 Projektdatei zurückschreiben.

Die folgenden Spalten können in das yw7-Projekt zurückgeschrieben werden:

-   Title
-   Beschreibung
-   Tags (durch Kommas getrennt)
-   Szenen-Notizen
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

Dies schreibt yw7 "Notizen"-Kapitel mit untergeordneten Szenen in ein neues OpenDocument-Textdokument (odt) mit unsichtbaren Kapitel- und Szenenabschnitten (zu sehen im Navigator). Das Suffix des Dateinamens ist `_notes`.

- Kommentare innerhalb von Szenen werden als Szenentitel zurückgeschrieben, 
  wenn sie von `~` umgeben sind.
- Kapitel und Szenen können weder umgeordnet noch gelöscht werden.
- Sie können Szenen durch Einfügen von Überschriften oder einem Szenentrenner aufteilen:
    - *Überschrift 1* → Neue Kapitelüberschrift (Beginn eines neuen Abschnitts). Sie können eine Beschreibung anfügen, durch `|` getrennt.
    - *Überschrift 2* → Neue Kapitelüberschrift. Sie können eine Beschreibung anfügen, durch `|` getrennt.
    - `###` → Szenentrenner. Optional können Sie den Szenentitel an den Szenentrenner anhängen. Sie können auch eine Beschreibung anfügen, durch `|` getrennt.
    - **Hinweis:** Exportieren Sie Dokumente mit aufgeteilten Szenen nicht mehrmals zu yw7. 
- Absätze, die mit `>` beginnen, werden als Zitate formatiert.

[Zum Seitenbeginn](#top)

------------------------------------------------------------------------

## Planungs-Kapitel

Dies schreibt yw7 "Todo"-Kapitel mit untergeordneten Szenen in ein neues 
OpenDocument-Textdokument (odt) mit unsichtbaren Kapitel- und Szenen 
Abschnitten (die im Navigator zu sehen sind). Die Dateinamenserweiterung ist `_todo`.

- Kommentare innerhalb von Szenen werden als Szenentitel zurückgeschrieben,
  wenn sie von `~` umgeben sind.
- Kapitel und Szenen können weder umgeordnet noch gelöscht werden.
- Sie können Szenen durch Einfügen von Überschriften oder einem Szenentrenner aufteilen:
    - *Überschrift 1* → Neue Kapitelüberschrift (Beginn eines neuen Abschnitts). Sie können eine Beschreibung anfügen, durch `|` getrennt.
    - *Überschrift 2* → Neue Kapitelüberschrift. Sie können eine Beschreibung anfügen, durch `|` getrennt.
    - `###` → Szenentrenner. Optional können Sie den Szenentitel an den Szenentrenner anhängen. Sie können auch eine Beschreibung anfügen, durch `|` getrennt.
    - **Hinweis:** Exportieren Sie Dokumente mit aufgeteilten Szenen nicht mehrmals zu yw7. 
- Absätze, die mit `>` beginnen, werden als Zitate formatiert.

[Zum Seitenbeginn](#top)

------------------------------------------------------------------------
