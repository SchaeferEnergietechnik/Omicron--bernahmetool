# Daten und Integrationen

## Quelldaten

### Omicron-Prüfdatei

- Dateiendung: `*.occ`
- Bedeutung: Omicron-Prüfung
- Verwendung: Quelle für den Export der Messdaten
- Technische Exportmethode: sichtbare Bedienung der Omicron-Anwendung; die bestehende funktionierende Implementierung soll übernommen werden

Für die Integration werden benötigt:

- Quellcode oder ausführbarer Aufruf der bestehenden Verarbeitung,
- benötigte Omicron-Software und Version,
- notwendige lokale Installation und Lizenz,
- verwendete Technik für Fenster-, Fokus-, Maus- oder Tastaturautomatisierung,
- erzeugtes Zwischenformat des Messdatenexports,
- Verhalten bei beschädigten oder nicht unterstützten `*.occ`-Dateien.

## Interaktive Windows-Sitzung

Die Exportlogik benötigt eine sichtbare Omicron-Anwendung in einer interaktiven Windows-Sitzung. Die Automatisierung darf nicht als unsichtbarer Hintergrunddienst eingeplant werden, solange keine andere Omicron-Schnittstelle verfügbar ist.

Auswirkungen:

- Die Sitzung darf während des Exports nicht gesperrt oder getrennt werden, sofern die verwendete Automatisierung das nicht unterstützt.
- Andere Maus-, Tastatur- und Fokusaktionen können den Ablauf stören.
- Der Benutzer muss vor Beginn deutlich gewarnt werden.
- Für echtes paralleles Arbeiten ist eine getrennte Sitzung, VM oder ein zweiter Rechner erforderlich.
- Die Automatisierung und die GUI sollen in getrennten Prozessen laufen, damit Statusüberwachung und kontrollierter Abbruch möglich bleiben.

## Terminexcel

Die Terminexcel liegt später als `Y:\GES Energietechnik\Termine.xlsx` in einer als Laufwerk `Y:` eingebundenen Cloud-Ablage. Es gibt keinen weiteren Unterordner. Sie kann über den normalen Windows-Dateipfad gelesen werden; eine Cloud-API ist für diesen Zugriff zunächst nicht erforderlich.

Benötigte Quelldaten:

- Prüfername,
- Prüfdatum,
- Kunde.

Der Suchschlüssel besteht zunächst aus Prüfername und Prüfdatum. Da diese Kombination mehrfach vorkommen kann, muss das Tool mehrdeutige Treffer anzeigen statt automatisch einen Kunden auszuwählen.

Die hochgeladene Beispieldatei verwendet das Blatt `Termine`. Jeder Prüfer besitzt dort einen eigenen Spaltenblock aus Datum, Prüfer-/Termintext, Kunde und weiteren Projektdaten. Die Checkbox-Kurzbezeichnungen der Prüfdatenmappe müssen über eine feste Alias-Tabelle den ausgeschriebenen Namen der Terminexcel zugeordnet werden.

Beispiel:

| Checkbox | Terminexcel | Datumsspalte | Kundenspalte |
|---|---|---|---|
| `Helmchen` | `Niklas Helmchen` | `Q` | `T` |
| `Fäthke` | `Pascal Fäthke` | `B` | `E` |
| `Schmidt` | `Hagen Schmidt` | `AF` | `AI` |
| `Koehn` | `Kevin Koehn` | `CN` | `CQ` |
| `Wendt` | `Sebastian Wendt` | `BJ` | `BM` |
| `Mummhardt` | `Elias Mummhardt` | `BY` | `CB` |
| `Kolzer` | `Finn Kolzer` | `AU` | `AX` |

Die in V19m noch vorhandene Checkbox `Schäfer` gilt als veraltet und wird nicht für eine automatische Terminzuordnung verwendet. `Mundkowski` ist in V19m nicht mehr vorhanden.

## Verifizierte Zellen der Prüfdaten-Arbeitsmappe

| Information | Technische Quelle |
|---|---|
| Prüfer | Formular-Checkboxen über Zeile 5 im Blatt `Schutzprüf-Checkliste` |
| Prüfdatum | `Schutzprüf-Checkliste!B7` |
| Ziel Kunde | `Allgemeine Angaben!C2` |
| Kundenliste | `Kunden!A1:A35` |

Die Prüfer-Checkboxen sind nicht mit Hilfszellen verknüpft. Ihr Name und Auswahlstatus müssen über Excel-COM oder direkt aus den VML-Formularinformationen gelesen werden. Im analysierten Beispiel ist `H. Schmidt` ausgewählt.

`Allgemeine Angaben!C2` verwendet eine erweiterte Excel-Datenvalidierung mit `Kunden!A1:A35` als Quelle. Das Tool darf einen Wert nur nach erfolgreichem Abgleich mit dieser Liste automatisch einsetzen; unbekannte Kunden werden zur Prüfung angezeigt.

## Zieldaten

### Excel-Arbeitsmappe

Voraussichtliche Formate:

- `*.xlsx`
- `*.xlsm`, falls Arbeitsmappen mit Makros verwendet werden

Noch zu klären:

- Name oder Muster des Zielarbeitsblatts,
- Zielzellen beziehungsweise Tabellenstruktur,
- Verhalten bei bereits vorhandenen Messdaten,
- Erhalt von Formeln, Formatierungen, Diagrammen und Makros,
- Notwendigkeit einer Sicherungskopie,
- Verhalten bei geöffneter oder gesperrter Datei.

## Zuordnungsmodell

Eine Verarbeitungseinheit besteht mindestens aus:

| Feld | Bedeutung |
|---|---|
| `workingDirectory` | Fachlicher Unterordner der Prüfung |
| `occPath` | Vollständiger Pfad zur Omicron-Prüfdatei |
| `excelPath` | Vollständiger Pfad zur ausgewählten Excel-Datei |
| `mappingStatus` | Eindeutig, Auswahl erforderlich, fehlt oder ausgeschlossen |
| `processingStatus` | Ausstehend, läuft, erfolgreich, Fehler oder abgebrochen |
| `message` | Warnung, Fehler oder Ergebnisinformation |

### Bekannte Dateinamensregel

Dateien der zweiten Station tragen im Dateinamen den Marker `EZE`. Der Marker kann an beliebiger Position stehen und wird ohne Beachtung der Groß-/Kleinschreibung erkannt. Technisch entspricht die Erkennung einer Suche nach `eze` im normalisierten Dateinamen.

Wenn in einem fachlichen Ordner genau eine Excel-Datei liegt, ist sie das gemeinsame Ziel aller dort gefundenen `*.occ`-Dateien. Das gilt insbesondere für den häufigen Fall aus zwei Omicron-Prüfungen und einer Excel-Arbeitsmappe. Der EZE-Marker unterscheidet dann die zweite Station beziehungsweise den dafür vorgesehenen Datenbereich innerhalb dieser Arbeitsmappe.

Nur wenn mehrere geeignete Excel-Dateien vorhanden sind, wird der EZE-Marker auch zur Zielauswahl verwendet: Genau eine EZE-markierte `*.occ`- und eine EZE-markierte Excel-Datei bilden einen Zuordnungsvorschlag. Mehrere oder fehlende EZE-Kandidaten werden als mehrdeutig beziehungsweise unvollständig markiert.

## Lokale Einstellungen

Mindestens zu speichern:

- zuletzt verwendeter Startordner,
- gegebenenfalls Fenster- und Anzeigeeinstellungen,
- später optional bestätigte Zuordnungsregeln.

Die Einstellungen sollen benutzerbezogen lokal gespeichert werden. Projektdateien dürfen dadurch nicht verändert werden.

## Schreibsicherheit

Bevor eine Excel-Datei ersetzt wird, soll das Tool in eine temporäre Datei schreiben und das Ergebnis prüfen. Erst danach wird die Zieldatei atomar ersetzt oder eine kontrollierte Sicherung angelegt. Das genaue Verfahren hängt von der verwendeten Excel-Bibliothek und dem Betriebssystem ab.

## Datenschutz und Netzwerk

- Die Verarbeitung erfolgt lokal.
- Der Startordner befindet sich lokal auf dem Arbeitsplatz.
- Quelldateien und Messdaten werden nicht an externe Dienste übertragen.
- Netzwerkpfade sind für den rekursiv durchsuchten Prüfungsordner nicht vorgesehen.
- Die Terminexcel unter dem gemappten Cloud-Pfad `Y:\GES Energietechnik\Termine.xlsx` ist die ausdrücklich vorgesehene Ausnahme und wird nur lesend zur Kundenzuordnung verwendet.
