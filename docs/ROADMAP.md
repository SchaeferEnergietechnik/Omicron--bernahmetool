# Roadmap

## Phase 0: Bestehende Logik sichern

- vorhandene Verarbeitung des aktuellen Ordners beschaffen
- Laufzeitumgebung und Abhängigkeiten dokumentieren
- verwendete Fenster-, Fokus-, Maus- und Tastaturautomatisierung dokumentieren
- repräsentative, anonymisierte Testdateien bereitstellen
- aktuellen Export und Excel-Schreibvorgang mit Tests charakterisieren

**Ergebnis:** Die bereits funktionierende Kernlogik ist reproduzierbar und kann ohne Funktionsverlust aufgerufen werden.

## Phase 1: Desktop-Grundlage und Ordnerwahl

- Zielplattform und Desktop-Technik festlegen
- grafische Ordnerauswahl umsetzen
- letzten gültigen Ordner lokal speichern
- aktuellen Einzelordner-Ablauf aus der GUI starten

**Ergebnis:** Der heutige Ablauf ist über eine GUI für einen frei wählbaren Ordner nutzbar.

## Phase 2: Rekursiver Scan und Vorschau

- Unterordner rekursiv durchsuchen
- `*.occ`- und Excel-Dateien inventarisieren
- eindeutige und unvollständige Fälle erkennen
- Vorschautabelle mit Pfaden und Status bauen

**Ergebnis:** Vor jeder Änderung ist vollständig sichtbar, welche Dateien verarbeitet würden.

## Phase 3: Sichere Zuordnung

- fachliche Zuordnungsregeln anhand echter Beispiele definieren
- automatische Vorschläge implementieren
- manuelle OCC->Excel-Auswahl bei mehreren Excel-Dateien ermöglichen
- unklare Einträge standardmäßig von der Verarbeitung ausschließen

**Ergebnis:** Mehrere Trafostationen im gleichen Bereich können sicher getrennt zugeordnet werden; der Start erfolgt nur mit vollständiger Zuordnung.

## Phase 4: Stapelverarbeitung und Abbruch

- Warteschlange der bestätigten Einträge aufbauen
- Fortschritt und aktuellen Eintrag anzeigen
- sichtbare Automatisierung in einen separaten Prozess auslagern
- Warnung vor der exklusiven Nutzung der aktiven Windows-Sitzung anzeigen
- kontrollierten Abbruch implementieren
- Excel-Dateien atomar beziehungsweise mit Sicherung schreiben
- Fehler je Eintrag isolieren
- Laufzeit und aktuellen Verarbeitungsschritt anzeigen
- Fehlerbericht für unbeaufsichtigte Läufe schreiben

**Ergebnis:** Große Ordnerbäume können nachvollziehbar und unterbrechbar verarbeitet werden.

## Phase 5: Protokoll und Freigabe

- Abschlussprotokoll und Wiederholungsfunktion ergänzen
- Abnahmeszenarien automatisieren
- Installation und Updateweg für Zielrechner festlegen
- Bedien- und Fehlerdokumentation erstellen

**Ergebnis:** Das Tool ist für den praktischen Einsatz auslieferbar.

## Nächster konkreter Schritt

Terminexcel-Kundenzuordnung in den Worker integrieren, danach die Erfolgsablage nach `Protokollentwürfe` einschließlich Konfliktbehandlung implementieren und mit einem echten Windows-End-to-End-Nachtlauf validieren.
