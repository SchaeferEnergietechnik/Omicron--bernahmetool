# VBA-Anpassung für V19m

## Ergebnis des Vergleichs

Die VBA-Projekte aus V19g und V19m wurden statisch verglichen.

| Vom Python-Tool aufgerufenes Makro | V19g | V19m | Maßnahme |
|---|---|---|---|
| `Tabelle1.Protokollnummer_generieren_unsichtbar` | vorhanden | vorhanden, identisch | nichts kopieren |
| `Modul1.BereicheEinOderAusblenden_Start` | vorhanden | fehlt; `Modul1` ist leer | aus V19g übernehmen |
| `Tabelle7.ZeilenAusblendenWennLeer` | vorhanden | vorhanden, identisch | nichts kopieren |

V19m enthält die weitgehend gleiche Bereichslogik bereits im privaten Ereignis `Tabelle7.CommandButton1_Click`. Python kann dieses private Button-Ereignis jedoch nicht als öffentliches Makro aufrufen. Deshalb wurde die Logik in V19g zusätzlich als `Public Sub BereicheEinOderAusblenden_Start` in `Modul1` bereitgestellt.

Die beiden Varianten unterscheiden sich an einer Stelle: Der private V19m-Button wertet zusätzlich `Schutzprüf-Checkliste!W14` aus und blendet damit die Zeilen 159 bis 164 ein oder aus.

**Aktueller Stand:** Das bestehende private Button-Ereignis in V19m bleibt unverändert. Die öffentliche Prozedur in `Modul1` enthält zusätzlich ebenfalls die W14-Regel für die Zeilen 159 bis 164, damit der Python-Aufruf und der Button konsistent ausblenden.

## Empfohlene Übernahme

Die Datei `legacy/vba/V19m_Modul1_Ergaenzung.bas` enthält eine bereinigte, direkt importierbare Variante des fehlenden Makros.

Änderungen gegenüber dem ursprünglichen V19g-Code:

- `Option Explicit` ergänzt,
- Zeilennummern als `Long` statt `Integer`,
- `Allgemeine Angaben` und `Prüfprotokoll` ausdrücklich über `ThisWorkbook` referenziert,
- alle `Rows(...)` ausdrücklich auf `Prüfprotokoll` bezogen,
- wiederholte Ein-/Ausblendlogik vereinfacht,
- W14-Regel für `Schutzprüf-Checkliste!W14` ergänzt (Zeilen 159 bis 164).

Die Fachlogik und verwendeten Zellen bleiben gleich.

## Manuelle Schritte in Excel

1. Eine Sicherungskopie von V19m erstellen.
2. V19m öffnen und mit `Alt+F11` den VBA-Editor starten.
3. Im Projekt von V19m das leere `Modul1` markieren.
4. Entweder den Inhalt aus `V19m_Modul1_Ergaenzung.bas` in `Modul1` einfügen oder das leere Modul entfernen und die BAS-Datei über `Datei -> Datei importieren...` importieren.
5. `Tabelle7.CommandButton1_Click` nicht ersetzen oder löschen, da es die zusätzliche W14-Regel enthält.
6. Im VBA-Editor `Debuggen -> VBAProject kompilieren` ausführen.
7. Die Arbeitsmappe weiterhin als `*.xlsm` speichern.
8. Das Makro `Modul1.BereicheEinOderAusblenden_Start` einmal an einer Kopie manuell testen.
9. Den bestehenden Excel-Button separat testen, damit seine W14-Regel erhalten bleibt.
10. Danach einen vollständigen Testlauf des Python-Programms an einer Kopie durchführen.

## Automatisierte Übertragung unter Windows

Alternativ kann die Migration für die beiden Beispiel-Dateien automatisiert per Excel-COM erfolgen:

```powershell
python samples/topics/excel-basis/legacy-python/migrate_v19g_vba_to_v19m.py
```

Das Skript:

- erstellt Sicherungskopien beider Sample-Dateien,
- exportiert `Modul1` aus V19g,
- ersetzt das Standardmodul `Modul1` in V19m,
- ergänzt bei Bedarf die W14-Ausblendlogik in `BereicheEinOderAusblenden_Start`,
- prüft, dass `Public Sub BereicheEinOderAusblenden_Start()` in V19m vorhanden ist,
- speichert V19m.

Voraussetzungen:

- Windows mit installiertem Excel,
- Python mit `pywin32`,
- in Excel ist der Zugriff auf das VBA-Projektmodell erlaubt:
	`Datei -> Optionen -> Trust Center -> Einstellungen fuer das Trust Center -> Makroeinstellungen -> Zugriff auf das VBA-Projektobjektmodell vertrauen`.

## Python-Aufrufe

Nach der Ergänzung kann das bestehende Python-Programm alle drei Aufrufe unverändert ausführen:

```python
Tabelle1.Protokollnummer_generieren_unsichtbar
Modul1.BereicheEinOderAusblenden_Start
Tabelle7.ZeilenAusblendenWennLeer
```

## Separater Prüfpunkt: Protokollersteller

Das Makro `Protokollnummer_generieren_unsichtbar` ist zwar zwischen V19g und V19m identisch, unterstützt in `Allgemeine Angaben!C7` aber weiterhin nur:

- `Gunnar Schäfer`
- `Kevin Koehn`

Die Prüfer-Checkboxen in `Schutzprüf-Checkliste` sind davon technisch getrennt. Falls künftig weitere Personen als Protokollersteller in `C7` auswählbar sein sollen, müssen deren Nummernkreise fachlich festgelegt und anschließend sowohl die Datenvalidierung in `C7` als auch beide `Select Case`-Blöcke im Makro erweitert werden.

## Aktuelle Prüfer-Checkboxen in V19m

Statisch gefunden wurden:

- `Schäfer` (laut Fachvorgabe veraltet und nicht mehr zu verwenden),
- `Helmchen`,
- `Fäthke`,
- `Schmidt`,
- `Koehn`,
- `Wendt`,
- `Mummhardt`,
- `Kolzer`.

`Mundkowski` ist in V19m nicht mehr enthalten.