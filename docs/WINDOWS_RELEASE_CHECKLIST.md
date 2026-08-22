# Windows Release Checklist

## 1) Version und Metadaten

- package.json Version erhoehen (kein 0.0.0 fuer produktive Auslieferung)
- package.json Felder description und author pflegen
- Product Name und App ID im Build-Block pruefen
- Falls gewuenscht: eigenes App-Icon hinterlegen

## 2) Abhaengigkeiten und Qualitaet

- npm install
- npm run lint
- npm run build
- Optional: npm audit und bewusste Bewertung offener Findings

## 3) Build auf Windows (Installer)

- Build auf echtem Windows-Rechner ausfuehren
- Befehl: npm run desktop:dist
- Erwartete Artefakte in dist:
  - Omicron Uebernahmetool Setup <version>.exe (NSIS)
  - win-unpacked Verzeichnis

## 4) Build im Linux-Container (Fallback)

- Fuer schnelle Abnahme: npx electron-builder --win portable
- Erwartetes Artefakt:
  - Omicron Uebernahmetool <version>.exe (portable)
- Hinweis: NSIS-Setup kann in headless Umgebungen an Wine/X11 scheitern

## 5) Signatur und Reputation

- Falls vorhanden: Code Signing Zertifikat konfigurieren
- Signierte EXE erzeugen und Signatur pruefen
- SmartScreen-Dialog auf sauberem Testsystem pruefen

## 6) Laufzeitvoraussetzungen am Zielsystem

- Windows mit installiertem Microsoft Excel
- Omicron Control Center installiert
- Python 3.10+ vorhanden
- Python Worker-Abhaengigkeiten installiert:
  - python -m pip install -r worker\requirements.txt
- Zugriff auf Y:\GES Energietechnik\Termine.xlsx verifiziert

## 7) Smoke Test vor Freigabe

- App startet ohne Fehler
- Cloud-Ordnerwahl funktioniert
- Lokaler Import zeigt Fundordner korrekt
- Manuelle OCC-zu-Excel-Zuordnung funktioniert
- Kundenzuordnung pruefen:
  - Teilwort-Matching greift
  - Rechtsformwoerter (z. B. gmbh, ag, kg) fuehren nicht allein zum Treffer
  - Mehrdeutigkeit loest manuelle Auswahl aus
- Ein Durchlauf mit Testdaten erzeugt erwartete Excel-Ergebnisse
- Fehlerbericht wird bei Fehler/Skip geschrieben

## 8) Release-Dokumentation

- Release-Notiz aktualisieren (Aenderungen, Risiken, offene Punkte)
- Commit-ID und Tag dokumentieren
- Build-Artefaktname und Build-Datum dokumentieren

## 9) Rollback und Sicherung

- Vor Auslieferung Git-Tag setzen
- Vorherige stabile EXE archivieren
- Rueckfallpfad benennen (Tag/Commit + altes Artefakt)

## 10) Auslieferung

- Setup/portable EXE in freigegebenen Verteiler kopieren
- Dateihash (z. B. SHA256) mitliefern
- Kurze Installations- und Betriebsanweisung beilegen
