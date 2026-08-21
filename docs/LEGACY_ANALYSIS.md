# Analyse des bestehenden Python-Programms

## Analysierter Bestand

- Python-Programm: `legacy/occ_export_sehr gut mit Excel übernahme-ohne pdf.py`
- Beispielmappe: `samples/V19g_Übergeordneter_Entkupplungsschutz.xlsm`
- Omicron-Beispiele:
  - `samples/V24_NAP_V2_Bensheim.occ`
  - `samples/V9_EZE_mit_Anregung_Bensheim.occ`

Die Analyse erfolgte statisch. Weder Omicron-Dateien noch VBA-Makros wurden in der Entwicklungsumgebung ausgeführt.

## Aktueller Ablauf

1. `Microsoft.Mashup.Container.Loader.exe` wird per `taskkill /F` beendet.
2. Alle `*.csv` unter `C:\Omicron_Datenexport` werden gelöscht.
3. Das aktuelle Arbeitsverzeichnis wird nicht rekursiv nach `*.occ` durchsucht.
4. Jede gefundene `*.occ` wird mit `os.startfile` in Omicron geöffnet.
5. `pywinauto` bedient sichtbar `Datei -> Daten exportieren...` und bestätigt den Exportdialog.
6. Omicron wird geschlossen; eine mögliche Speicherabfrage wird mit `Nein` beantwortet.
7. Nach allen OCC-Exporten wird die alphabetisch erste `*.xlsm` im aktuellen Ordner ausgewählt.
8. Excel wird sichtbar über COM geöffnet und `RefreshAll` ausgeführt.
9. Das Blatt `Prüfprotokoll` wird aktiviert.
10. Drei VBA-Makros werden ausdrücklich gestartet.
11. Die Arbeitsmappe wird gespeichert und Excel beendet.

## Datenfluss

```mermaid
flowchart LR
    A[Mehrere OCC-Dateien im aktuellen Ordner] --> B[Sichtbarer Omicron-Export]
    B --> C[CSV-Dateien unter C:\Omicron_Datenexport]
    C --> D[Power Query RefreshAll]
    D --> E[Eine gemeinsame XLSM-Arbeitsmappe]
    E --> F[VBA-Nachverarbeitung]
```

Das hochgeladene Beispiel bestätigt den häufigen Fall aus zwei `*.occ`-Dateien und einer gemeinsamen `*.xlsm`. Die Arbeitsmappe enthält unter anderem die Blätter `Daten`, `Daten EZE`, `Measurements`, `Seq_Measurement` und `PlsRmp_Measurement` sowie drei Power-Query-Verbindungen.

## Python-Abhängigkeiten

- Python 3.10 oder neuer wegen der Typangabe `str | None`
- `pywinauto`
- `pywin32` für `win32com.client`
- Windows mit Omicron Control Center und Microsoft Excel

## Excel-Verhalten

### Verifizierte Eingabedaten

- Der Prüfer wird über mehrere Formular-Checkboxen über Zeile 5 des Blatts `Schutzprüf-Checkliste` ausgewählt. Die Checkboxen sind nicht mit Zellen verknüpft; im Beispiel ist `H. Schmidt` angehakt.
- Das Prüfdatum steht in `Schutzprüf-Checkliste!B7`; im Beispiel ist der gespeicherte Wert der 19.06.2026.
- Der Auftraggeber steht in `Allgemeine Angaben!C2`.
- `C2` verwendet eine erweiterte Datenvalidierung mit `Kunden!A1:A35` als Dropdown-Quelle.

### Beim Öffnen

Das Modul `DieseArbeitsmappe` ist leer. In der analysierten Beispielmappe wurden keine klassischen Auto-Start-Prozeduren gefunden:

- kein `Workbook_Open`
- kein `Auto_Open`
- kein `Workbook_Activate`

Es existieren `Worksheet_SelectionChange`-Ereignisse auf einzelnen Tabellenblättern. Diese reagieren auf eine Auswahländerung, sind aber keine expliziten Öffnungs-Makros.

Die statische Analyse markiert außerdem ActiveX-Ereignisnamen vorsorglich als mögliche AutoExec-Ereignisse. Der Handler `txtVersion_Change` im PDF-Dialog ist leer. Ein weiterer Handler heißt `ommandButton1_Click` ohne führendes `C` und entspricht daher nicht dem üblichen Ereignisnamen `CommandButton1_Click`.

### Vom Python-Programm ausdrücklich gestartete Makros

#### `Tabelle1.Protokollnummer_generieren_unsichtbar`

- liest Jahr, Monat, Bearbeiter und Projektdaten aus der Arbeitsmappe,
- öffnet `C:\Firma\Übersicht Prüfprotokolle_auto.xlsx`,
- ermittelt die nächste laufende Nummer,
- schreibt zwei Protokollnummern in die aktuelle Arbeitsmappe,
- ergänzt und sortiert die zentrale Übersicht,
- speichert die Übersicht.

Das Makro akzeptiert derzeit ausschließlich die fest codierten Namen `Gunnar Schäfer` und `Kevin Koehn`. Bei anderen Namen erscheint eine Meldung und das Makro endet.

#### `Modul1.BereicheEinOderAusblenden_Start`

- wertet Angaben aus dem Blatt `Allgemeine Angaben` aus,
- blendet Bereiche des aktiven Prüfprotokolls abhängig von vorhandenen Schutzfunktionen ein oder aus.

#### `Tabelle7.ZeilenAusblendenWennLeer`

- prüft im Blatt `Prüfprotokoll` die Zeilen 161 bis 182,
- blendet Zeilen mit leerer Spalte A aus und andere wieder ein.

## Risiken des aktuellen Programms

### Globale CSV-Löschung

Vor jedem Lauf werden ausnahmslos alle `*.csv` in `C:\Omicron_Datenexport` gelöscht. Ein Abbruch oder paralleler Prozess kann dadurch fremde beziehungsweise noch benötigte Exporte verlieren.

### Globaler Prozessabbruch

Das Programm beendet jeden Prozess namens `Microsoft.Mashup.Container.Loader.exe` mit `/F`. Das kann auch andere geöffnete Excel- oder Power-Query-Vorgänge betreffen.

### Auswahl der Excel-Datei

Bei mehreren `*.xlsm` wird ohne Rückfrage die alphabetisch erste Datei verwendet. Das widerspricht der geplanten sicheren Vorschau und Zuordnung.

### Excel-Instanz

Es wird `win32com.client.Dispatch("Excel.Application")` verwendet und am Ende `excel.Quit()` aufgerufen. Für die neue Anwendung ist eine isolierte Excel-Instanz vorzuziehen, damit eine bereits laufende Benutzerinstanz und deren Arbeitsmappen nicht versehentlich beeinflusst werden.

### Fehlende Transaktionssicherheit

Die Zielarbeitsmappe wird direkt gespeichert. Es gibt noch keine Sicherungskopie, temporäre Ergebnisdatei oder atomare Ersetzung.

### Abbruch

`KeyboardInterrupt` wird nur am äußeren Programmeinstieg behandelt. Es gibt keinen kooperativen Abbruchstatus zwischen den einzelnen UI-Schritten und keine definierte Wiederaufnahme.

### Feste Pfade und Namen

Fest codiert sind unter anderem:

- `C:\Omicron_Datenexport`
- `C:\Firma\Übersicht Prüfprotokolle_auto.xlsx`
- Blatt `Prüfprotokoll`
- Makronamen und Bearbeiternamen

Diese Werte müssen vor der Integration als Konfiguration oder bewusst bestätigte Fachkonstanten eingeordnet werden.

## Wiederverwendbare Teile

- robuste Suche nach Omicron-Fenstern über reguläre Fenstertitel,
- mehrere Fallbacks für Menü und Exportdialog,
- Behandlung der Omicron-Speicherabfrage,
- sequenzielle Verarbeitung mehrerer OCC-Dateien,
- Warten auf Excel-Berechnung und asynchrone Abfragen,
- explizite Reihenfolge der drei VBA-Makros,
- grundlegende Protokollierung pro Datei.

## Empfohlener Integrationsschnitt

Die vorhandene Logik sollte zunächst funktional unverändert in einen separaten Windows-Worker überführt werden. GUI und Worker kommunizieren über klar definierte Aufträge und Statusmeldungen. Vor dem Worker-Aufruf übernimmt die GUI rekursive Suche, Vorschau und Zuordnung; der Worker verarbeitet danach jeweils einen bestätigten fachlichen Ordner mit seinen OCC-Dateien und genau einem Excel-Ziel.