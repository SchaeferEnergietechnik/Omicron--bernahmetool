# Produktvision

## Ausgangslage

Omicron-Prüfungen liegen als `*.occ`-Dateien in Projektordnern. Aus diesen Dateien können Messdaten exportiert werden. Die exportierten Werte werden in eine Excel-Datei übernommen, die im gleichen fachlichen Arbeitsordner liegt.

Eine bestehende Verarbeitung kann diesen Ablauf bereits für den aktuellen Ordner durchführen. Sie durchsucht jedoch weder frei wählbare Ordner noch deren Unterordner und bietet keine geeignete grafische Kontrolle für mehrere Prüfungen und Excel-Dateien.

## Ziel

Das Omicron Übernahmetool soll diesen Ablauf als lokale, nachvollziehbare Stapelverarbeitung bereitstellen:

1. Der Benutzer wählt einen Cloud-Quellordner und einen lokalen Arbeitsordner aus.
2. Das Tool merkt sich beide Ordner für den nächsten Start.
3. Nach manuellem Start werden Cloud-Unterordner mit `*.occ`-Dateien gesucht und vollständig lokal kopiert, ohne die Cloud-Daten zu verändern.
4. Die lokalen Kopien werden nach `*.occ`-Dateien und passenden Excel-Dateien durchsucht.
5. Vor der Verarbeitung erscheint eine Vorschau der geplanten Zuordnungen.
6. Eindeutige Fälle können gemeinsam verarbeitet werden.
7. Mehrdeutige Fälle müssen sichtbar sein und dürfen nicht stillschweigend falsch zugeordnet werden.
8. Der laufende Vorgang kann kontrolliert abgebrochen werden.
9. Am Ende ist nachvollziehbar, was erfolgreich war, übersprungen wurde oder fehlgeschlagen ist.

Die Zuordnung und Verarbeitung erfolgen ausschließlich im lokalen Arbeitsordner. Dateien der zweiten Station sind durch den Marker `EZE` im Dateinamen gekennzeichnet; Groß-/Kleinschreibung und Position des Markers spielen keine Rolle.

## Nutzer

Primäre Nutzer sind Mitarbeitende, die Omicron-Prüfungen aus mehreren Trafostationen oder Projektordnern gesammelt auswerten und die Messdaten in bestehende Excel-Arbeitsunterlagen übernehmen.

## Erfolgskriterien

Das Tool ist fachlich brauchbar, wenn:

- ein frei wählbarer Cloud-Quellordner rekursiv nach OCC-Fundordnern durchsucht werden kann,
- Fundordner ohne Veränderung der Cloud-Daten lokal kopiert werden,
- Cloud-Quellordner und lokaler Arbeitsordner nach einem Neustart wieder angeboten werden,
- jede `*.occ`-Datei vor dem Schreiben einem sichtbaren Excel-Ziel zugeordnet ist,
- unklare Zuordnungen eine Benutzerentscheidung verlangen,
- keine Excel-Datei unbemerkt oder teilweise beschädigt wird,
- ein Abbruch keine aktuell geschriebene Excel-Datei in einem inkonsistenten Zustand hinterlässt,
- ein Abschlussprotokoll alle Einzelergebnisse enthält.

## Nicht im ersten Zielumfang

- direkte Verarbeitung oder Veränderung von Dateien in der Cloud
- Mehrbenutzerbetrieb
- Verarbeitung überwachter Netzwerkordner ohne Benutzeraktion
- allgemeiner Import beliebiger CSV- oder JSON-Dateien
