# Projektstand bei Pausierung

**Stand:** 9. August 2026  
**Status:** Electron-Desktop-GUI, nativer Cloud-Import und Python-Worker-Anbindung als erste produktive Windows-Grundlage umgesetzt. Fachliche End-to-End-Erweiterungen sind noch offen.

## Ziel

Das spätere Windows-Tool soll einen lokalen Startordner rekursiv nach `*.occ` und zugehörigen Excel-Dateien durchsuchen, die Zuordnung vorab anzeigen und anschließend die Omicron-Daten sichtbar exportieren und in die Arbeitsmappen übernehmen. Der Vorgang muss kontrolliert abbrechbar sein.

## Festgelegter Ablauf

1. Lokalen Startordner auswählen; den zuletzt verwendeten Ordner speichern.
2. Startordner und Unterordner rekursiv nach OCC- und Excel-Dateien durchsuchen.
3. Gefundene Dateien und vorgeschlagene Zuordnungen anzeigen.
4. Mehrdeutige Zuordnungen vom Benutzer bestätigen lassen.
5. OCC-Dateien nacheinander über die sichtbare Omicron-Oberfläche exportieren.
6. Exportdaten über Power Query in die zugeordnete Excel-Datei laden.
7. Kunden anhand von Prüfer und Prüfdatum aus `Y:\GES Energietechnik\Termine.xlsx` ermitteln und gegen `Kunden!A1:A35` prüfen.
8. Kundenwert vor der Protokollnummernerzeugung nach `Allgemeine Angaben!C2` schreiben.
9. Die drei vorgesehenen VBA-Makros ausführen und die Arbeitsmappe speichern.

## Erledigt

- Native Electron-Desktop-Hülle mit sicherer Preload-/IPC-Grenze ergänzt.
- Native Windows-Ordnerdialoge für Cloud-Quelle und lokalen Arbeitsordner ergänzt.
- Rekursiver Cloud-Scan und nicht überschreibender lokaler Import in den Desktop-Prozess verlagert.
- GUI mit dem separaten Python-Worker, JSON-Lines-Fortschritt und kontrolliertem Abbruch verbunden.
- Electron-Installer-Konfiguration für einen Windows-NSIS-Build ergänzt.

- React-/TypeScript-/Vite-Oberfläche als Renderer der Desktop-Anwendung umgesetzt.
- Produktvision, Ablauf, Anforderungen, Integrationen, Entscheidungen und Roadmap dokumentiert.
- Bestehendes Python-Programm statisch analysiert und unverändert unter `legacy/` abgelegt.
- Anonymisierte OCC-, XLSM- und Terminexcel-Beispiele unter `samples/` statisch analysiert.
- Zuordnungsregeln für mehrere OCC- und Excel-Dateien sowie den EZE-Fall festgelegt.
- Terminexcel-Aufbau und Prüferspalten analysiert.
- Produktivpfad der Terminexcel auf `Y:\GES Energietechnik\Termine.xlsx` festgelegt; kein weiterer Unterordner und kein datierter Dateiname.
- V19g und V19m auf VBA-Quelltextebene verglichen.
- Fehlendes öffentliches V19m-Makro als importierbares Modul unter `legacy/vba/V19m_Modul1_Ergaenzung.bas` angelegt.
- VBA-Migrationsanleitung erstellt und statisch validiert.

## VBA-Ergebnis für V19m

Das Python-Programm erwartet drei Makros:

| Makro | V19m-Stand |
|---|---|
| `Tabelle1.Protokollnummer_generieren_unsichtbar` | vorhanden und mit V19g identisch |
| `Modul1.BereicheEinOderAusblenden_Start` | fehlt; Ergänzung liegt als BAS-Datei vor |
| `Tabelle7.ZeilenAusblendenWennLeer` | vorhanden und mit V19g identisch |

Das bestehende private Ereignis `Tabelle7.CommandButton1_Click` in V19m darf nicht ersetzt werden. Es enthält zusätzlich eine W14-Regel für die Zeilen 159 bis 164. Die öffentliche Prozedur wird nur ergänzend in das leere `Modul1` importiert.

Die Binärdatei V19m wurde im Repository nicht verändert. Import, Kompilierung und Funktionstest müssen später unter Windows mit Excel an einer Sicherungskopie erfolgen.

## Prüfer und Terminzuordnung

Aktuell für die automatische Terminzuordnung vorgesehen:

- Helmchen
- Fäthke
- Schmidt
- Koehn
- Wendt
- Mummhardt
- Kolzer

Die in V19m noch vorhandene Checkbox `Schäfer` gilt fachlich als veraltet und wird ignoriert. `Mundkowski` ist in V19m nicht mehr enthalten.

Interne Termine und Abwesenheiten, beispielsweise Urlaub oder Elternzeit, dürfen nicht automatisch als Kunde übernommen werden. Mehrere passende Termine oder nicht eindeutig normalisierbare Kundennamen erfordern eine Benutzerentscheidung.

## Offene fachliche Punkte

- Mehrdeutige Excel-Ziele werden derzeit sicher als Konflikt angezeigt, aber noch nicht in der GUI manuell aufgelöst.
- Erfolgreiche Fundordner werden noch nicht automatisch nach `Protokollentwürfe` verschoben.
- Die Terminexcel-Kundenauflösung ist noch nicht in den Worker integriert.
- Abschlussprotokoll, Wiederholung fehlgeschlagener Einträge und atomare Excel-Sicherung fehlen noch.

- Nummernkreise und vollständige Namen weiterer Protokollersteller festlegen. `Allgemeine Angaben!C7` und das Protokollnummernmakro unterstützen derzeit nur `Gunnar Schäfer` und `Kevin Koehn`.
- Verhalten bei mehreren Terminen desselben Prüfers am selben Tag abschließend festlegen.
- Regeln zur Normalisierung abweichender Kundennamen zwischen Terminexcel und Kundenliste bestätigen.
- Pflegeweg für künftig hinzukommende Prüfer festlegen.

## Technische Risiken

- Omicron wird sichtbar über `pywinauto` bedient und benötigt eine aktive Windows-Sitzung.
- Parallelbetrieb in derselben Windows-Sitzung ist nicht zuverlässig möglich.
- Das Legacy-Programm löscht global CSV-Dateien unter `C:\Omicron_Datenexport` und beendet den Power-Query-Loader zwangsweise; dies muss vor Produktiveinsatz eingegrenzt werden.
- Das Legacy-Programm wählt aktuell nur die alphabetisch erste XLSM-Datei und speichert ohne Transaktion oder Sicherung.
- Ein sauberer Abbruch zwischen Omicron-, Export- und Excel-Schritten fehlt noch.
- Die Terminexcel auf Laufwerk `Y:` muss erreichbar sein, bevor eine Arbeitsmappe verändert wird.

## Empfohlene Wiederaufnahme

1. V19m an einer Sicherungskopie gemäß `V19M_VBA_MIGRATION.md` ergänzen, kompilieren und alle drei Makros testen.
2. Offene Nummernkreise für Protokollersteller klären.
3. Mehrdeutige OCC-/Excel-Zuordnungen in der GUI manuell auflösbar machen.
4. Legacy-Omicron-Automation kapseln und sichere CSV-Arbeitsverzeichnisse einführen.
5. Terminexcel-Kundenauflösung und Excel-Verarbeitung integrieren.
6. Erfolgsablage, Abschlussprotokoll, Sicherung und Wiederanlauf ergänzen.
7. End-to-End-Test unter Windows mit Kopien der Beispieldateien durchführen.

## Validierungsstand

- Frontend: `npm run lint` und `npm run build` bestanden beim letzten Frontend-Stand.
- Legacy-Python: Syntaxprüfung bestanden.
- VBA-Ergänzung: Makroname, Modulname und verwendete Steuerzellen statisch gegen V19g geprüft.
- Dokumentation: Markdown-Links und `git diff --check` geprüft.
- Nicht möglich im Linux-Container: Omicron-Automation, Excel-COM, VBA-Kompilierung und End-to-End-Ausführung.

## Wichtige Einstiegsdokumente

- `WORKFLOW.md`: fachlicher Gesamtablauf
- `REQUIREMENTS.md`: Muss-Anforderungen und Akzeptanzfälle
- `DECISIONS.md`: bestätigte Entscheidungen und offene Punkte
- `LEGACY_ANALYSIS.md`: Verhalten und Risiken des bestehenden Python-Programms
- `TERMINEXCEL_ANALYSIS.md`: Aufbau und Zuordnungslogik der Terminexcel
- `V19M_VBA_MIGRATION.md`: konkrete VBA-Übernahme in V19m
- `ROADMAP.md`: geplante Umsetzungsschritte