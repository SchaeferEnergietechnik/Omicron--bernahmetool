# Analyse der Terminexcel

## Analysierte Datei

`samples/Termine 21.07.2026_Beispieldaten.xlsx`

Die Analyse erfolgte statisch. Die Datei wurde nicht mit Excel geöffnet und enthält laut OOXML-Struktur weder VBA-Projekt noch externe Datenverbindungen.

## Produktiver Speicherbereich

Die produktive Terminexcel liegt später unter `Y:\GES Energietechnik\Termine.xlsx`. Es gibt keinen weiteren Unterordner; der Dateiname ist fest und enthält kein Auswahldatum.

## Arbeitsblätter

Die Beispieldatei enthält unter anderem:

- einzelne Prüferblätter wie `Pascal`, `Niklas`, `Hagen`, `Kevin`, `Finn`, `Basti` und `Elias`,
- das zusammengeführte Blatt `Termine`,
- `Kundenadressen`,
- weitere organisatorische Blätter.

Für die automatische Suche ist das zusammengeführte Blatt `Termine` geeignet, weil es Datum, Prüferzuordnung und Kunde gemeinsam enthält. Die einzelnen Prüferblätter enthalten nicht durchgehend eine eigene Kundenspalte.

## Struktur des Blatts Termine

Jeder Prüfer besitzt einen eigenen horizontalen Spaltenblock. Beispiele:

| Prüferüberschrift | Datum | Kunde |
|---|---|---|
| `Pascal Fäthke` in `D1` | `B` | `E` |
| `Niklas Helmchen` in `S1` | `Q` | `T` |
| `Hagen Schmidt` in `AH1` | `AF` | `AI` |
| `Finn Kolzer` in `AW1` | `AU` | `AX` |
| `Sebastian Wendt` in `BL1` | `BJ` | `BM` |
| `Elias Mummhardt` in `CA1` | `BY` | `CB` |
| `Kevin Koehn` in `CP1` | `CN` | `CQ` |

Die Zuordnung sollte anhand der Überschriften erkannt und nicht ausschließlich über fest codierte Spaltenbuchstaben implementiert werden. Dadurch bleiben eingefügte oder verschobene Spalten beherrschbar.

## Prüfer-Aliase

Die Prüfdatenmappe verwendet verkürzte Checkbox-Beschriftungen, während die Terminexcel ausgeschriebene Namen nutzt. Statisch bestätigt ist:

| Prüfdatenmappe | Terminexcel |
|---|---|
| `Fäthke` | `Pascal Fäthke` |
| `Helmchen` | `Niklas Helmchen` |
| `Schmidt` | `Hagen Schmidt` |
| `Koehn` | `Kevin Koehn` |
| `Wendt` | `Sebastian Wendt` |
| `Mummhardt` | `Elias Mummhardt` |
| `Kolzer` | `Finn Kolzer` |

Die V19m-Datei enthält zusätzlich noch die Checkbox `Schäfer`. Laut Fachvorgabe prüft G. Schäfer nicht mehr; dieser Eintrag gilt daher als veraltet und wird nicht automatisch zugeordnet. `Mundkowski` ist in V19m nicht mehr enthalten.

## Ergebnis für die hochgeladenen Beispieldaten

Aus der Prüfdaten-Arbeitsmappe wurden gelesen:

- ausgewählter Prüfer: `H. Schmidt`,
- Prüfdatum: 19.06.2026.

Nach Aliasauflösung entspricht dies dem Block `Hagen Schmidt`. In der Terminexcel steht am 19.06.2026:

- Termintext: `Elternzeit`,
- Kunde: `intern G.E.S. Energietechnik`.

Dieser Treffer passt nicht zum bereits in `Allgemeine Angaben!C2` gespeicherten Kunden `Faber E-Tec GmbH`. Das Beispiel belegt daher, dass Prüfer und Datum allein zwar einen Datensatz finden können, aber interne oder Abwesenheitstermine nicht automatisch als Kundenauftrag übernommen werden dürfen.

## Vorgeschlagene Suchlogik

1. Checkbox-Kurzname über Alias-Tabelle auf den vollständigen Prüfernamen abbilden.
2. Im Kopf des Blatts `Termine` den Prüferblock anhand des vollständigen Namens finden.
3. Innerhalb dieses Blocks die Spalten `Datum` und `Kunde` relativ zur Prüferüberschrift bestimmen.
4. Datumswerte als echte Excel-Datumswerte vergleichen, nicht als formatierte Zeichenfolge.
5. Leere, interne oder Abwesenheitseinträge herausfiltern.
6. Den verbleibenden Kunden normalisiert gegen `Kunden!A1:A35` der Prüfdatenmappe abgleichen.
7. Nur bei genau einem fachlich gültigen Treffer automatisch nach `Allgemeine Angaben!C2` schreiben.

## Offene Punkte

- Pflegeweg für später neu hinzukommende Prüfer,
- Schlüsselwörter für interne und Abwesenheitstermine,
- Entscheidung bei mehreren Außenterminen am selben Tag,
- Abgleich verkürzter Kundennamen aus der Terminexcel mit vollständigen Adressblöcken der Kundenliste.