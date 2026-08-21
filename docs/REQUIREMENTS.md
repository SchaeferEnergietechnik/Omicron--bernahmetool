# Anforderungen

## Funktionale Anforderungen

| ID | Anforderung | Priorität |
|---|---|---|
| F-001 | Der Benutzer kann den Cloud-Quellordner und den lokalen Arbeitsordner über grafische Ordnerauswahlen festlegen. | Muss |
| F-002 | Die zuletzt gültigen Ordner werden lokal gespeichert und beim nächsten Start angeboten. | Muss |
| F-003 | Das Tool durchsucht nach dem Cloud-Import den lokalen Arbeitsordner und alle relevanten Unterordner nach `*.occ`-Dateien. | Muss |
| F-004 | Das Tool erfasst geeignete Excel-Dateien in den relevanten Ordnern. | Muss |
| F-005 | Vor dem Schreiben wird eine Vorschau aller gefundenen Verarbeitungseinheiten angezeigt. | Muss |
| F-006 | Jede `*.occ`-Datei wird mit ihrem vorgesehenen Excel-Ziel dargestellt. | Muss |
| F-007 | Mehrdeutige Zuordnungen werden nicht automatisch verarbeitet. | Muss |
| F-008 | Der Benutzer kann bei Mehrdeutigkeit eine Excel-Datei auswählen oder den Eintrag ausschließen. | Muss |
| F-009 | Die Messdaten werden mit der bestehenden Exportlogik aus der `*.occ`-Datei exportiert. | Muss |
| F-010 | Die exportierten Messdaten werden in die zugeordnete Excel-Datei übernommen. | Muss |
| F-011 | Mehrere Verarbeitungseinheiten werden nacheinander bearbeitet. | Muss |
| F-012 | Die laufende Stapelverarbeitung kann kontrolliert abgebrochen werden. | Muss |
| F-013 | Fortschritt und aktuell verarbeiteter Eintrag werden angezeigt. | Muss |
| F-014 | Nach Abschluss oder Abbruch wird ein Ergebnis pro Eintrag angezeigt. | Muss |
| F-015 | Nicht lesbare Dateien und Ordner werden verständlich protokolliert. | Muss |
| F-016 | Der Benutzer kann fehlgeschlagene oder nicht gestartete Einträge erneut ausführen. | Soll |
| F-017 | Das Ergebnisprotokoll kann als Datei gespeichert werden. | Soll |
| F-018 | Bestätigte Zuordnungsmuster können später als Regel wiederverwendet werden. | Kann |
| F-019 | Der Arbeitsordner für Zuordnung und Verarbeitung liegt lokal auf dem Arbeitsplatz. | Muss |
| F-020 | Der Marker `EZE` wird im gesamten Dateinamen ohne Beachtung der Groß-/Kleinschreibung erkannt. | Muss |
| F-021 | Bei genau einer EZE-markierten `*.occ`- und einer EZE-markierten Excel-Datei schlägt das Tool diese als Paar der zweiten Station vor. | Muss |
| F-022 | Vor Beginn der sichtbaren Omicron-Automatisierung warnt das Tool, dass die aktive Windows-Sitzung nicht parallel bedient werden darf. | Muss |
| F-023 | Ein robuster Abbruchmechanismus kann die sichtbare Automatisierung kontrolliert beenden. | Muss |
| F-024 | Sind in einem fachlichen Ordner mehrere `*.occ`-Dateien und genau eine Excel-Datei vorhanden, werden alle Prüfungen dieser gemeinsamen Excel-Datei zugeordnet. | Muss |
| F-025 | Bei einer gemeinsamen Excel-Datei kennzeichnet `EZE` die zweite Station beziehungsweise ihren Datenbereich und führt nicht zur Suche nach einer zweiten Excel-Datei. | Muss |
| F-026 | Das Tool liest den angehakten Prüfer aus den Formular-Checkboxen über Zeile 5 im Blatt `Schutzprüf-Checkliste`. | Muss |
| F-027 | Das Tool liest das Prüfdatum aus `Schutzprüf-Checkliste!B7`. | Muss |
| F-028 | Das Tool sucht in der Terminexcel nach der Kombination aus Prüfer und Prüfdatum. | Muss |
| F-029 | Bei genau einem Termin liest das Tool den Wert aus der Spalte `Kunde`. | Muss |
| F-030 | Der gelesene Kunde wird gegen die Liste `Kunden!A1:A35` der Prüfdaten-Arbeitsmappe abgeglichen. | Muss |
| F-031 | Bei eindeutigem Abgleich schreibt das Tool den vollständigen Kundeneintrag nach `Allgemeine Angaben!C2`. | Muss |
| F-032 | Kein oder mehrere Termin- beziehungsweise Kundentreffer erfordern eine sichtbare Benutzerentscheidung. | Muss |
| F-033 | Die Terminexcel wird unter `Y:\GES Energietechnik\Termine.xlsx` gelesen. | Muss |
| F-034 | Prüfer-Kurzbezeichnungen aus den Checkboxen werden über eine konfigurierbare Alias-Tabelle den Prüferblöcken der Terminexcel zugeordnet. | Muss |
| F-035 | Interne Einträge, Urlaub, Elternzeit und freie Tage werden nicht ungeprüft als Kundenauftrag übernommen. | Muss |
| F-036 | Der Produktivlauf prüft, ob das Laufwerk `Y:` und die Terminexcel erreichbar sind, bevor eine Prüfdaten-Datei verändert wird. | Muss |
| F-037 | Nach manuellem Start durchsucht das Tool zuerst einen in der GUI konfigurierbaren Cloud-Quellordner einschließlich seiner Unterordner nach Ordnern mit mindestens einer `*.occ`-Datei. | Muss |
| F-038 | Der zuletzt gültige Cloud-Quellordner wird lokal gespeichert und beim nächsten Start vorausgewählt. | Muss |
| F-039 | Der Benutzer kann in der GUI einen lokalen Zielordner für die Cloud-Kopien festlegen; der zuletzt gültige Zielordner wird lokal gespeichert und beim nächsten Start vorausgewählt. | Muss |
| F-040 | Das Tool kopiert jeden gefundenen Ordner mit `*.occ`-Datei einschließlich seines Inhalts in den lokalen Zielordner und erhält dabei seine relative Ordnerstruktur. | Muss |
| F-041 | Beim Cloud-Scan und Kopieren werden keine Dateien oder Ordner in der Cloud verändert, verschoben oder gelöscht. | Muss |
| F-042 | Die Zuordnung und Datenübernahme aus `*.occ` in Excel erfolgt ausschließlich mit den lokalen Kopien. | Muss |
| F-043 | Vorhandene lokale Zieldateien werden nicht ungeprüft überschrieben; Konflikte werden angezeigt und erfordern eine Benutzerentscheidung. | Muss |
| F-044 | Nach erfolgreicher Verarbeitung aller Verarbeitungseinheiten eines lokalen Fundordners verschiebt das Tool diesen Ordner in den Unterordner `Protokollentwürfe` des lokalen Arbeitsordners. | Muss |
| F-045 | Enthält ein Fundordner fehlgeschlagene, abgebrochene, ausgeschlossene oder nicht gestartete Verarbeitungseinheiten, bleibt er außerhalb von `Protokollentwürfe` im lokalen Arbeitsordner. | Muss |
| F-046 | Der Unterordner `Protokollentwürfe` wird bei Cloud-Kopie und lokaler Suche ausgeschlossen, damit erledigte Ordner nicht erneut verarbeitet werden. | Muss |
| F-047 | Besteht im Ziel bereits ein gleichnamiger Ordner unter `Protokollentwürfe`, wird der Fundordner nicht automatisch überschrieben; der Konflikt wird protokolliert und verlangt eine Benutzerentscheidung. | Muss |

## Nichtfunktionale Anforderungen

| ID | Anforderung |
|---|---|
| N-001 | Nach dem ausschließlich lesenden Import aus dem konfigurierten Cloud-Ordner verarbeitet die Anwendung die Daten lokal ohne weiteren externen Datentransfer. |
| N-002 | Ein Fehler an einer Datei beendet nicht automatisch den gesamten Stapel. |
| N-003 | Eine Excel-Datei darf bei Fehler oder Abbruch nicht beschädigt oder nur teilweise geschrieben zurückbleiben. |
| N-004 | Scan und Vorprüfung blockieren die Oberfläche nicht; während der sichtbaren Omicron-Automatisierung bleiben mindestens Status und Abbruchmechanismus verfügbar. |
| N-005 | Lange Pfade und Dateinamen werden vollständig einsehbar dargestellt. |
| N-006 | Warnungen und Fehler nennen betroffene Quelle und betroffenes Ziel. |
| N-007 | Die Anwendung muss auf dem vorgesehenen Windows-Arbeitsplatz lauffähig sein. |
| N-008 | Die sichtbare Automatisierung darf nicht den Eindruck erwecken, sie könne störungsfrei parallel zur normalen Bedienung derselben Windows-Sitzung laufen. |
| N-009 | Automatisierung und Bedienoberfläche sollen in getrennten Prozessen laufen, damit ein Fehler der Automatisierung die Steuerung nicht unmittelbar beendet. |

## Abnahmeszenarien

### A-001: Einzelner eindeutiger Ordner

Ein Ordner enthält eine `*.occ`- und eine passende Excel-Datei. Die Vorschau zeigt genau eine eindeutige Zuordnung. Nach Bestätigung werden die Messdaten übernommen und als erfolgreich protokolliert.

### A-002: Rekursive Verarbeitung

Der Startordner enthält mehrere Unterordner mit eindeutigen Paaren. Alle Unterordner werden gefunden und nach Bestätigung nacheinander verarbeitet.

### A-003: Zwei Trafostationen

Ein lokaler fachlicher Bereich enthält zwei `*.occ`- und zwei Excel-Dateien. Die Dateien der zweiten Station enthalten jeweils an beliebiger Stelle `EZE` oder eine andere Groß-/Kleinschreibung dieses Markers. Das Tool ordnet die beiden EZE-Dateien einander als zweite Station zu, zeigt beide Prüfungen in der Vorschau und verlangt bei verbleibender Mehrdeutigkeit eine Benutzerentscheidung, bevor eine Datei verändert wird.

### A-004: Fehlendes Excel-Ziel

Zu einer `*.occ`-Datei existiert kein geeignetes Excel-Ziel. Der Eintrag wird mit Warnung angezeigt und nicht verarbeitet.

### A-005: Benutzerabbruch

Der Benutzer bricht während eines Stapels ab. Die aktuelle Schreiboperation endet konsistent, weitere Einträge starten nicht und die Ergebnisliste unterscheidet erfolgreiche von abgebrochenen Einträgen.

### A-006: Gesperrte Excel-Datei

Eine zugeordnete Excel-Datei ist geöffnet oder nicht beschreibbar. Der Eintrag schlägt mit verständlicher Meldung fehl; andere Einträge können weiterlaufen.

### A-007: Sichtbare Automatisierung

Vor dem ersten Export weist das Tool auf die benötigte ungestörte Windows-Sitzung hin. Nach Bestätigung bedient es Omicron sichtbar und zeigt den aktuellen Eintrag sowie einen erreichbaren Abbruchmechanismus an.

### A-008: Zwei OCC-Dateien und eine Excel-Datei

Ein lokaler fachlicher Ordner enthält zwei `*.occ`-Dateien, davon eine mit dem Marker `EZE`, und genau eine Excel-Arbeitsmappe. Die Vorschau zeigt beide Prüfungen mit derselben Excel-Datei als Ziel. Beide Exporte werden durchgeführt und anschließend gemeinsam in diese Arbeitsmappe übernommen.

### A-009: Eindeutige Kundenzuordnung

In der Prüfdaten-Arbeitsmappe ist genau ein Prüfer angehakt und in `B7` steht ein gültiges Prüfdatum. Die Terminexcel enthält genau einen Termin für diese Kombination. Dessen Kunde entspricht genau einem Eintrag aus `Kunden!A1:A35`. Das Tool schreibt diesen Eintrag nach `Allgemeine Angaben!C2`, bevor die Protokollnummer erzeugt wird.

### A-010: Mehrdeutiger Termin

Die Terminexcel enthält mehrere Termine für denselben Prüfer am selben Tag. Das Tool verändert `Allgemeine Angaben!C2` nicht automatisch, zeigt alle möglichen Kunden an und verlangt eine Auswahl.

### A-011: Interner oder Abwesenheitstermin

Für Prüfer und Datum enthält die Terminexcel einen Eintrag wie `intern`, `Urlaub`, `Elternzeit` oder `frei`. Das Tool übernimmt diesen Wert nicht ungeprüft nach `Allgemeine Angaben!C2`, sondern meldet, dass kein eindeutiger Kundenauftrag gefunden wurde.

### A-012: Terminexcel nicht erreichbar

Das Laufwerk `Y:` oder die Terminexcel ist nicht erreichbar. Das Tool zeigt den Fehler vor der Excel-Nachverarbeitung an und verändert die Prüfdaten-Arbeitsmappe nicht.

### A-013: Cloud-Ordner lokal bereitstellen

Nach manuellem Start durchsucht das Tool den gespeicherten Cloud-Quellordner rekursiv. Einen Unterordner mit mindestens einer `*.occ`-Datei kopiert es vollständig unter Beibehaltung der relativen Struktur in den gespeicherten lokalen Zielordner. Die anschließende Vorschau und Verarbeitung verwenden nur diese lokale Kopie; die Dateien und Ordner in der Cloud bleiben unverändert.

### A-014: Konflikt im lokalen Zielordner

Für einen gefundenen Cloud-Ordner bestehen im lokalen Ziel bereits gleichnamige Dateien. Das Tool überschreibt sie nicht automatisch, sondern zeigt den Konflikt an und verlangt vor dem Kopieren eine Benutzerentscheidung. Die Cloud-Quelldaten bleiben in jedem Fall unverändert.

### A-015: Erfolgreichen Fundordner ablegen

Alle Verarbeitungseinheiten eines lokalen Fundordners wurden erfolgreich in Excel übernommen. Nach dem Speichern verschiebt das Tool den vollständigen Fundordner nach `Protokollentwürfe` unterhalb des lokalen Arbeitsordners. Bei einem Folge-Scan wird dieser Ablageordner nicht erneut durchsucht.

### A-016: Unvollständigen Fundordner liegen lassen

Mindestens eine Verarbeitungseinheit eines lokalen Fundordners ist fehlgeschlagen, abgebrochen, ausgeschlossen oder nicht gestartet. Der Fundordner wird nicht nach `Protokollentwürfe` verschoben und bleibt zur erkennbaren Nachbearbeitung im lokalen Arbeitsordner.
