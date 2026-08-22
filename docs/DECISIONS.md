# Entscheidungen und offene Punkte

Dieses Dokument hält bestätigte Entscheidungen und Fragen fest, die vor der jeweiligen Implementierungsphase beantwortet werden müssen.

## Bestätigte Entscheidungen

### D-001: Lokale Verarbeitung

**Status:** entschieden

Nach einer vorgelagerten Kopie aus dem Cloud-Quellordner finden die Prüfung und Excel-Verarbeitung ausschließlich im lokalen Arbeitsordner statt. Die Quelldaten in der Cloud werden nur gelesen und niemals verändert, verschoben oder gelöscht.

### D-002: Frei wählbare und gespeicherte Ordner

**Status:** entschieden

Der Benutzer wählt den Cloud-Quellordner und den lokalen Arbeitsordner über die GUI. Beide zuletzt gültigen Ordner werden lokal gespeichert und beim nächsten Start wieder angeboten.

### D-003: Rekursive Suche

**Status:** entschieden

Zuerst werden der ausgewählte Cloud-Quellordner und seine Unterordner nach Ordnern mit Omicron-Prüfungen durchsucht. Nach dem Kopieren werden die lokalen Fundordner nach `*.occ`- und Excel-Dateien durchsucht.

### D-004: Vorschau vor Verarbeitung

**Status:** entschieden

Vor dem Schreiben werden gefundene Unterordner, `*.occ`-Dateien und vorgesehene Excel-Ziele angezeigt. Mehrdeutige Zuordnungen müssen erkennbar sein.

### D-005: Kontrollierter Abbruch

**Status:** entschieden

Die Stapelverarbeitung kann durch den Benutzer abgebrochen werden. Der Abbruch darf keine teilweise geschriebene Excel-Datei hinterlassen.

### D-006: Kennzeichnung der zweiten Station durch EZE

**Status:** entschieden

Dateien der zweiten Station enthalten im Dateinamen an beliebiger Stelle den Marker `EZE`. Die Erkennung erfolgt ohne Beachtung der Groß-/Kleinschreibung, sodass beispielsweise `EZE`, `eze` und `Eze` gleich behandelt werden.

Bei mehreren Stationen dient dieser Marker als erste feste Zuordnungsregel: Eine als EZE erkannte Omicron-Prüfung soll der ebenfalls als EZE erkannten Excel-Datei zugeordnet werden. Fehlt auf einer Seite eine eindeutige EZE-Kennzeichnung oder entstehen mehrere Kandidaten, bleibt eine manuelle Auswahl erforderlich.

### D-007: Export durch sichtbare Bedienung

**Status:** entschieden

Die bestehende Exportlogik bedient die Omicron-Anwendung sichtbar. Sie benötigt damit die aktive interaktive Windows-Sitzung und kann von Fokuswechseln sowie Maus- oder Tastatureingaben beeinflusst werden.

Während eines aktiven Exports ist zuverlässiges paralleles Arbeiten in derselben Windows-Sitzung nicht möglich. Für ungestörtes Weiterarbeiten muss die Verarbeitung in einer separaten Windows-Sitzung, virtuellen Maschine oder auf einem zweiten Rechner laufen.

### D-008: Mehrere OCC-Dateien und eine Excel-Datei

**Status:** entschieden

Häufig enthält ein fachlicher Ordner zwei `*.occ`-Dateien, aber nur eine Excel-Arbeitsmappe. In diesem Fall werden die Messdaten beider Omicron-Prüfungen in dieselbe Excel-Datei übernommen.

Der Marker `EZE` dient in diesem Fall der Unterscheidung der Stationen beziehungsweise Datenbereiche innerhalb der gemeinsamen Arbeitsmappe. Er ist erst dann für die Auswahl zwischen Excel-Zielen relevant, wenn im Ordner mehrere geeignete Excel-Dateien vorhanden sind.

### D-009: Kundenzuordnung über Prüfer und Prüfdatum

**Status:** Zielablauf bestätigt, technische Details der Terminexcel offen

Das Tool soll den angehakten Prüfer aus Zeile 5 des Blatts `Schutzprüf-Checkliste` und das Prüfdatum aus Zelle `B7` lesen. Mit Prüfer und Datum sucht es den passenden Termin in der Terminexcel. Aus der Spalte `Kunde` übernimmt es den Kunden in `Allgemeine Angaben!C2` der Prüfdaten-Arbeitsmappe.

Die Kundenzelle `C2` besitzt bereits eine Dropdown-Liste mit der Quelle `Kunden!A1:A35`. Der aus der Terminexcel gelesene Kunde soll deshalb gegen diese Liste abgeglichen und als vorhandener vollständiger Listeneintrag geschrieben werden. Bei keinem oder mehreren Treffern ist vor dem Schreiben eine Benutzerentscheidung erforderlich.

### D-010: Speicherbereich der Terminexcel

**Status:** entschieden

Die produktive Terminexcel liegt unter `Y:\GES Energietechnik\Termine.xlsx`. Es gibt keinen weiteren Unterordner und keine dynamische Auswahl anhand eines Datums im Dateinamen. Die Terminexcel wird nur gelesen. Davon unabhängig wird der konfigurierbare Cloud-Quellordner nach Prüfungsordnern durchsucht; deren Verarbeitung erfolgt erst nach der lokalen Kopie.

### D-011: Nicht destruktiver Cloud-Import

**Status:** entschieden

Nach manuellem Start durchsucht das Tool den in der GUI gespeicherten Cloud-Quellordner rekursiv nach Ordnern mit mindestens einer `*.occ`-Datei. Diese Fundordner werden vollständig und unter Erhalt ihrer relativen Struktur in den ebenfalls gespeicherten lokalen Arbeitsordner kopiert. Bestehende lokale Dateien werden nicht ungeprüft überschrieben. Der eigentliche OCC-Export und die Excel-Datenübernahme arbeiten nur mit den lokalen Kopien.

### D-012: Erfolgreich bearbeitete Ordner ablegen

**Status:** entschieden

Sobald alle Verarbeitungseinheiten eines lokalen Fundordners erfolgreich abgeschlossen sind, verschiebt das Tool den vollständigen Fundordner nach `<lokaler Arbeitsordner>/Protokollentwürfe`. Ordner mit fehlgeschlagenen, abgebrochenen, ausgeschlossenen oder nicht gestarteten Einträgen bleiben außerhalb dieses Unterordners, damit offene Arbeit sichtbar bleibt. `Protokollentwürfe` wird von zukünftigen Kopier- und Scanvorgängen ausgeschlossen. Besteht dort bereits ein gleichnamiger Ordner, ist eine Benutzerentscheidung erforderlich; ein automatisches Überschreiben findet nicht statt.

### D-013: Manuelle OCC->Excel-Zuordnung vor Start

**Status:** entschieden und in GUI umgesetzt

Wenn ein Fundordner mehrere Excel-Dateien enthält, wird pro OCC-Datei eine manuelle Zuordnung OCC->Excel in der GUI erfasst. Der Verarbeitungsstart ist blockiert, solange mindestens eine OCC-Datei im betroffenen Fundordner nicht zugeordnet ist.

### D-014: Nachtlauf mit Fehlerbericht

**Status:** entschieden und im Worker umgesetzt

Der unbeaufsichtigte Lauf arbeitet bei Einzelfehlern weiter und sammelt Fehler sowie übersprungene Einträge. Nach Laufende wird bei Bedarf ein maschinenlesbarer Fehlerbericht erzeugt.

### D-015: Ausschluss erledigter Ordner beim Scan

**Status:** entschieden und in Desktop-Scan umgesetzt

Ordner, deren Name den Bestandteil `erledigt` enthält, werden bei der rekursiven Suche vollständig ausgeschlossen. Der Ausschluss gilt unabhängig von Groß-/Kleinschreibung und umfasst den gesamten Teilbaum inklusive aller Unterordner.

Ziel ist eine kürzere Suchlaufzeit und die Vermeidung unnötiger Verarbeitung bereits abgeschlossener Bereiche.

## Offene Entscheidungen

### O-001: Zuordnung OCC zu Excel

**Geklärt:** Ist genau eine Excel-Datei vorhanden, ist sie das gemeinsame Ziel aller `*.occ`-Dateien des Fundordners. Gibt es mehrere Excel-Dateien, erfolgt eine manuelle OCC->Excel-Zuordnung in der GUI vor Start.

**Verbleibende Frage:** Soll künftig zusätzlich eine automatische Vorbelegung der manuellen Zuordnung aus Namensmustern erfolgen?

### O-002: Bestehende Exportlogik

**Teilweise geklärt:** Die bestehende Einzelordner-Verarbeitung bedient Omicron sichtbar über die grafische Oberfläche.

**Verbleibende Frage:** In welcher Sprache und mit welcher Automatisierungstechnik funktioniert diese Bedienung?

Die Antwort bestimmt, wie die Automatisierung in einen separaten Prozess ausgelagert und wie ein sicherer Abbruch umgesetzt werden kann.

### O-003: Excel-Struktur

**Frage:** In welche Arbeitsblätter, Tabellen oder Zellen werden Messdaten geschrieben und wie werden bestehende Werte behandelt?

### O-004: Unterstützte Excel-Formate

**Frage:** Müssen neben `*.xlsx` auch `*.xlsm` oder ältere `*.xls`-Dateien unterstützt werden?

### O-005: Sicherung und Überschreiben

**Frage:** Wird die originale Excel-Datei direkt ersetzt, vorher gesichert oder als neue Ergebnisdatei gespeichert?

### O-006: Zielplattform

**Frage:** Welche Windows-Versionen und Omicron-Installationen müssen unterstützt werden?

### O-007: Scan-Grenzen

**Teilweise geklärt:** `Protokollentwürfe` und Ordner mit Namensbestandteil `erledigt` sind ausgeschlossen.

**Verbleibende Frage:** Sollen darüber hinaus weitere Ordner, Dateimuster oder Archivbereiche von der rekursiven Suche ausgeschlossen werden?

### O-008: Aufbau und Zugriff der Terminexcel

**Fragen:**

- Wie werden neu hinzukommende Prüfer dauerhaft in der Alias-Tabelle gepflegt?
- Wie wird entschieden, wenn ein Prüfer am selben Tag mehrere Kundentermine hat?
- Enthält die Spalte `Kunde` exakt den Eintrag aus `Kunden!A1:A35` oder nur einen verkürzten Kundennamen?
- Wie sollen interne Termine, Urlaub, Elternzeit und freie Tage behandelt werden?

## Entscheidungsformat

Wenn ein offener Punkt geklärt ist, wird er in eine bestätigte Entscheidung überführt und erhält:

- Entscheidung,
- Begründung,
- betrachtete Alternativen,
- Auswirkungen auf Implementierung und Tests.