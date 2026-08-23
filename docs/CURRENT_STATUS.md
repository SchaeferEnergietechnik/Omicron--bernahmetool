# Projektstand

**Stand:** 22. August 2026  
**Status:** Electron-Desktop-GUI, nativer Cloud-Import, Python-Worker mit robuster Excel-Verarbeitung, manuelle OCC-zu-Excel-Zuordnung, sichtbare Schrittanzeige und optionales Herunterfahren nach Laufende sind lauffähig.  
**Letzter Commit:** `48d0812` – Always save Excel even when customer lookup or macros fail

## Sicherung und Dokumentation (22.08.2026)

- Technischer Sicherungsstand auf `main`: `21de035`
- Fachliche Änderung Kundenzuordnung (Teilwortabgleich mit Stopwort-Filter): `d41eb95`
- Konsistenzanpassungen Doku und UI-Texte: `21de035`
- Packaged-Worker-Pfad und `asarUnpack` für Python-Worker: `95b1512`
- Robuste Nachtlauf-Verarbeitung (Kundenzuordnung/Makrofehler als Warnung): `fedcf35`, `48d0812`
- Option "Nach Beendigung Rechner herunterfahren" in der GUI: `d66b148`
- Wiederherstellung auf diesen Stand ist jederzeit per Git-Checkout der Commit-ID oder über den zugehörigen Tag möglich.

## Ziel

Das Windows-Tool soll einen Cloud-Quellordner lokal bereitstellen, Fundordner rekursiv erfassen, OCC-Dateien vor dem Start eindeutig auf Excel-Ziele zuordnen und danach robust als Stapel abarbeiten. Für Nachtschichten soll der Lauf bei Einzelfehlern weiterlaufen und am Ende einen nachvollziehbaren Fehlerbericht erzeugen.

## Festgelegter Ablauf

1. Lokalen Startordner auswählen; den zuletzt verwendeten Ordner speichern.
2. Startordner und Unterordner rekursiv nach OCC- und Excel-Dateien durchsuchen.
3. Gefundene Dateien und Zuordnungsstatus anzeigen.
4. Bei mehreren Excel-Dateien pro Ordner OCC-Dateien manuell in der GUI zuordnen.
5. Verarbeitung erst starten, wenn alle Zuordnungen eindeutig sind.
6. Je Zuordnungsgruppe: Mashup beenden, OCC sichtbar exportieren, danach Excel aktualisieren und Makros ausführen.
7. Laufzeit und aktueller Schritt laufend anzeigen.
8. Bei Einzelfehlern nicht abbrechen, sondern weiterarbeiten, speichern was möglich ist und Fehler sammeln.
9. Nach Laufende Zusammenfassung und Fehlerbericht ausgeben.
10. Optional Rechner nach Abschluss automatisch herunterfahren.

## Erledigt

- Native Electron-Desktop-Hülle mit sicherer Preload-/IPC-Grenze ergänzt.
- Native Windows-Ordnerdialoge für Cloud-Quelle und lokalen Arbeitsordner ergänzt.
- Rekursiver Cloud-Scan und nicht überschreibender lokaler Import in den Desktop-Prozess verlagert.
- Cloud-Scan optimiert: Ordner mit `erledigt` im Namen werden inklusive Unterordner ausgeschlossen.
- Cloud-Scan weiter optimiert: Ordner mit Präfix `zz_` werden inklusive Unterordner ausgeschlossen.
- GUI mit dem separaten Python-Worker, JSON-Lines-Fortschritt und kontrolliertem Abbruch verbunden.
- Electron-Installer-Konfiguration für einen Windows-NSIS-Build ergänzt.
- Verarbeitungsschritte im UI sichtbar gemacht: Mashup, OCC-Export, Excel-Bearbeitung, Abschluss.
- Laufzeit-Timer im UI ergänzt.
- Scan-Fortschrittsanzeige für große Verzeichnisbäume robuster gemacht (animierter Fortschrittsbalken und Zusatzhinweis bei großen Archivbereichen).
- Fehlerbericht als JSON-Datei aus dem Worker ergänzt, wenn Fehler oder Skip-Fälle auftreten.
- Worker-Reihenfolge für Fundordner stabilisiert: OCC-Reihenfolge NAP/sonstige vor EZE.
- Nachtlauf-Verhalten umgesetzt: Einzelfehler beenden nicht den gesamten Lauf.
- Excel-Verarbeitung gehärtet: Datei wird auch bei Kundenzuordnungs- oder Makrofehlern gespeichert.
- Kundenzuordnung erweitert: wenn kein Treffer gefunden wird, bleibt der vorhandene Kunde aus `Allgemeine Angaben!C2` als Fallback erhalten.
- Excel-Dialoge zur Link-/Datenaktualisierung werden für unbeaufsichtigte Läufe unterdrückt.
- Option in der GUI ergänzt: Rechner nach Beendigung automatisch herunterfahren (standardmäßig deaktiviert).
- Manuelle Zuordnung OCC -> Excel in der GUI vor Verarbeitungsstart ergänzt.
- Terminexcel-Kundenauflösung im Worker integriert (Teilwortabgleich gegen Kundenliste mit Stopwort-Filter, Datumsvergleich ohne Uhrzeit, interne Begriffe gefiltert).

- **Worker-Integration (neu):** Python-Worker führt alle drei erforderlichen Excel-Makros aus:
  - `Tabelle1.Protokollnummer_generieren_unsichtbar` (Protokollnummer-Generierung)
  - `Modul1.BereicheEinOderAusblenden_Start` (mit Fallback auf `BereicheEinOderAusblenden_Start`)
  - `Tabelle7.ZeilenAusblendenWennLeer` (Leerzeilen ausblenden)
- **Makro-Fehlerbehandlung (neu):** Versucht Bereichsmakro mit zwei Kandidaten; gibt aussagekräftige Fehlermeldung, wenn beide fehlen.
- **Vorbedingung für V19m:** VBA-Ergänzung `legacy/vba/V19m_Modul1_Ergaenzung.bas` muss noch manuell unter Windows mit Excel in V19m importiert werden.

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

Das bestehende private Ereignis `Tabelle7.CommandButton1_Click` in V19m darf nicht ersetzt werden. Es enthält die W14-Regel für die Zeilen 159 bis 164. Die öffentliche Prozedur in `Modul1` ist ergänzend vorhanden und enthält diese W14-Regel ebenfalls, damit Python-Aufruf und Button konsistent arbeiten.

Die Binärdatei V19m wurde im Repository nicht verändert. Import, Kompilierung und Funktionstest müssen später unter Windows mit Excel an einer Sicherungskopie erfolgen.

## Aktuelle Worker-Implementierung

Der Python-Worker unter `worker/occ_worker.py` ist nun vollständig implementiert:

- **OCC-Export:** Öffnet OCC sichtbar, öffnet Exportdialog, wartet auf stabile CSV-Ausgabe, schließt Omicron
- **Excel-Refresh:** Öffnet Arbeitsmappe, führt `RefreshAll` durch, wartet auf Berechnung
- **Excel ohne Rückfragen:** Unterdrückt Excel-Link-/Aktualisierungsabfragen beim Öffnen für Nachtläufe
- **Makro-Ausführung:** Ruft nacheinander auf:
  1. `Tabelle1.Protokollnummer_generieren_unsichtbar` (Protokollnummer)
  2. `Modul1.BereicheEinOderAusblenden_Start` oder `BereicheEinOderAusblenden_Start` (mit Fallback-Logik)
  3. `Tabelle7.ZeilenAusblendenWennLeer` (Leerzeilen-Ausblendung)
- **Fehlerbehandlung:** Sammelt Fehler und Skip-Fälle für den Abschlussbericht, setzt Verarbeitung bei Einzelfehlern fort
- **Speicherstrategie:** Speichert Ergebnisdateien auch bei Makro- oder Kundenzuordnungsproblemen; bei SaveAs-Fehlern Fallback auf `Save()` der Originaldatei
- **Kundenfallback:** Bei nicht auflösbarem Terminexcel-Treffer wird der vorhandene Vorlagenkunde aus `Allgemeine Angaben!C2` beibehalten
- **Fortschritt:** Emittiert JSON-Lines-Events für Mashup, OCC, Excel und Gesamtfortschritt
- **Abbruch:** Respektiert Abbruchdatei an sicheren Grenzen

## Vorbedingungen für Produktiveinsatz

1. **V19m-VBA:** `legacy/vba/V19m_Modul1_Ergaenzung.bas` muss in V19m per `Alt+F11` -> Importieren importiert werden
2. **Excel-COM:** Windows-System mit Excel und VBA-Laufzeitumgebung erforderlich
3. **Omicron:** Omicron muss auf dem Windows-System installiert sein
4. **Cloud-Quelle:** Erreichbarer Cloud-Ordner oder lokaler Quellordner konfiguriert
5. **Terminexcel:** `Y:\GES Energietechnik\Termine.xlsx` muss für die Kundenauflösung erreichbar sein


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

Interne Termine und Abwesenheiten, beispielsweise Urlaub oder Elternzeit, dürfen nicht automatisch als Kunde übernommen werden. Mehrere passende Termine oder nicht eindeutig normalisierbare Kundennamen werden protokolliert; der Lauf wird fortgesetzt und die Vorlage bleibt mit bestehendem Kundenwert speicherbar.

## Offene fachliche Punkte

- Erfolgreiche Fundordner werden noch nicht automatisch nach `Protokollentwürfe` verschoben.
- Abschlussprotokoll, Wiederholung fehlgeschlagener Einträge und atomare Excel-Sicherung fehlen noch.

- Nummernkreise und vollständige Namen weiterer Protokollersteller festlegen. `Allgemeine Angaben!C7` und das Protokollnummernmakro unterstützen derzeit nur `Gunnar Schäfer` und `Kevin Koehn`.
- Pflegeweg für zusätzliche interne Ausschlussbegriffe festlegen.
- Pflegeweg für künftig hinzukommende Prüfer festlegen.

## Technische Risiken

- Omicron wird sichtbar über `pywinauto` bedient und benötigt eine aktive Windows-Sitzung.
- Parallelbetrieb in derselben Windows-Sitzung ist nicht zuverlässig möglich.
- Das Legacy-Programm löscht global CSV-Dateien unter `C:\Omicron_Datenexport` und beendet den Power-Query-Loader zwangsweise; dies muss vor Produktiveinsatz eingegrenzt werden.
- Das Legacy-Programm wählt aktuell nur die alphabetisch erste XLSM-Datei und speichert ohne Transaktion oder Sicherung.
- Für End-to-End-Validierung bleibt ein echter Windows-Testlauf mit Omicron und Excel erforderlich.
- Die Terminexcel auf Laufwerk `Y:` muss erreichbar sein, bevor eine Arbeitsmappe verändert wird.

## Nächste Entwicklungsschritte

1. **V19m-VBA manuell ergänzen** (Vorbedingung für Tests) – unter Windows an einer Sicherungskopie durchführen:
   - `legacy/vba/V19m_Modul1_Ergaenzung.bas` öffnen
   - V19m in Excel öffnen → `Alt+F11` → VBA-Editor
   - Leeres `Modul1` markieren und Inhalt einfügen oder BAS-Datei importieren
   - Makros testen: `Modul1.BereicheEinOderAusblenden_Start`, `Tabelle1.Protokollnummer_generieren_unsichtbar`, `Tabelle7.ZeilenAusblendenWennLeer`

2. **Terminexcel-Kundenauflösung in Worker integrieren** – Kundennamen aus Terminexcel zuordnen

3. **Erfolgsablage nach `Protokollentwürfe`** – erfolgreiche Fundordner automatisch verschieben

4. **End-to-End-Test unter Windows** – mit Kopien der Beispieldateien und echtem Omicron

5. **Wiederanlauf fehlgeschlagener Einträge** – neue Jobs aus Skip-Liste generieren

6. **Atomare Excel-Sicherung** – Backup vor Makro-Ausführung, Wiederherstellung bei Fehler

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