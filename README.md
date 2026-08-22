# Omicron Übernahmetool

Lokales Werkzeug, das Ordnerbäume nach Omicron-Prüfdateien (`*.occ`) durchsucht, die Messdaten exportiert und in die jeweils zugehörige Excel-Datei im selben Arbeitsbereich übernimmt.

## Zielbild

Die Anwendung soll einen frei wählbaren Startordner einschließlich aller Unterordner prüfen. Vor der Verarbeitung zeigt sie die gefundenen Prüfdateien und ihre vorgesehenen Excel-Ziele an. Der Benutzer kann mehrdeutige Zuordnungen kontrollieren und den laufenden Vorgang jederzeit kontrolliert abbrechen.

Die ausführliche Beschreibung liegt unter:

- [Aktueller Projektstand und Wiederaufnahme](docs/CURRENT_STATUS.md)
- [Produktvision](docs/VISION.md)
- [Fachlicher Ablauf](docs/WORKFLOW.md)
- [Daten und Integrationen](docs/DATA_AND_INTEGRATIONS.md)
- [Anforderungen](docs/REQUIREMENTS.md)
- [Roadmap](docs/ROADMAP.md)
- [Entscheidungen und offene Punkte](docs/DECISIONS.md)
- [Analyse des bestehenden Python-Programms](docs/LEGACY_ANALYSIS.md)
- [Analyse der Terminexcel](docs/TERMINEXCEL_ANALYSIS.md)
- [VBA-Anpassung für V19m](docs/V19M_VBA_MIGRATION.md)
- [Release 2026-08-22](docs/RELEASE_2026-08-22.md)
- [Windows Release Checklist](docs/WINDOWS_RELEASE_CHECKLIST.md)

## Aktueller Stand

Die produktive Desktop-Basis ist umgesetzt: Electron-Oberfläche, nativer Cloud-Import, Python-Worker, manuelle OCC-zu-Excel-Zuordnung, Schrittanzeige und robuster Nachtlauf mit Fehlerbericht sind vorhanden. Die Terminexcel-Kundenauflösung ist integriert (Teilwortabgleich mit Stopwort-Filter). Wenn keine eindeutige Kundenzuordnung gelingt oder Makros fehlschlagen, wird der Lauf fortgesetzt und die Excel dennoch gespeichert. Optional kann der Rechner nach Laufende automatisch heruntergefahren werden (standardmäßig deaktiviert). Offen sind vor allem die automatische Erfolgsablage nach `Protokollentwürfe`, der Wiederanlauf fehlgeschlagener Einträge und die atomare Excel-Sicherung. Der detaillierte Übergabestand steht unter [Aktueller Projektstand und Wiederaufnahme](docs/CURRENT_STATUS.md).

## Windows-Desktopversion

Die produktive Oberfläche läuft als Electron-Desktopprogramm. Sie verwendet native
Windows-Ordnerdialoge und startet den Python-Worker als separaten Prozess. Für die
sichtbare Omicron- und Excel-Automatisierung wird Windows mit installiertem Omicron
Control Center, Excel und Python 3.10 oder neuer benötigt.

```powershell
npm install
npm run desktop
```

### EXE erstellen

Windows-Setup (NSIS):

```powershell
npm install
npm run desktop:dist
```

Hinweis: In headless Linux-Containern kann das NSIS-Setup wegen fehlender GUI/Wine-Umgebung fehlschlagen. Als lauffaehige Alternative kann eine portable EXE gebaut werden:

```bash
npm install
npx electron-builder --win portable
```

Ausgabe typischerweise unter `dist/Omicron Uebernahmetool 0.0.0.exe`.

Für den Worker müssen die Python-Abhängigkeiten installiert sein:

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r worker\requirements.txt
```

Die Desktop-App kopiert Cloud-Fundordner zunächst ausschließlich lesend in den
lokalen Arbeitsordner. Vor der echten Verarbeitung müssen die angezeigten
Zuordnungen geprüft werden. Während Omicron sichtbar bedient wird, darf die aktive
Windows-Sitzung nicht anderweitig verwendet werden.

Ordner mit `Protokollentwürfe`, einem Namensbestandteil `erledigt` oder dem
Präfix `zz_` werden beim rekursiven Cloud-Scan einschließlich aller Unterordner
bewusst übersprungen.

## Entwicklung

Voraussetzung: Node.js 20 oder neuer.

```bash
npm install
npm run dev
```

Die Anwendung ist anschließend standardmäßig unter `http://localhost:5173` erreichbar.

## Qualitätssicherung

```bash
npm run lint
npm run build
```

## Technische Basis

- React 19
- TypeScript
- Vite
- ESLint
- Lucide React
