# Änderungs-Matrix V20a (Schutzprüf-Checkliste -> Auswirkungen)

- Datei: samples/topics/excel-basis/V20a_Übergeordneter_Entkupplungsschutz.xlsm
- Erstellt (UTC): 2026-08-23 09:56:34
- Zweck: Schnelle Risikoabschätzung vor Änderungen an Eingabezellen in Schutzprüf-Checkliste.

## Schnellfazit

- Änderungen in Schutzprüf-Checkliste wirken breit in Allgemeine Angaben und Prüfprotokoll hinein.
- Kritisch für Warnanzeigen: E63/E64 sowie neu E77/E78 (NICHT OK-Logik im Protokollbereich).
- Kritisch für Zeilensichtbarkeit: W14 (VBA blendet Prüfprotokoll 159:164 ein/aus).

## Top-Eingabezellen nach Anzahl direkter Formelabhängigkeiten

| Zelle (Schutzprüf-Checkliste) | Anzahl direkter Formel-Ziele |
|---|---:|
| C25 | 56 |
| M24 | 16 |
| D60 | 10 |
| D66 | 9 |
| E64 | 8 |
| E63 | 6 |
| H66 | 6 |
| D61 | 5 |
| B7 | 5 |
| H65 | 5 |
| C67 | 4 |
| C68 | 4 |
| Q24 | 4 |
| K24 | 4 |
| I24 | 4 |
| D29 | 3 |
| E74 | 3 |
| E91 | 3 |
| I72 | 3 |
| B4 | 3 |
| B6 | 3 |
| K47 | 3 |
| C83 | 3 |
| L14 | 2 |
| N14 | 2 |

## Fokus-Matrix für geplante Änderungen

| Eingabezelle | Direkte Formel-Ziele | CF-Auswirkung | VBA-Auswirkung | Hinweis |
|---|---:|---|---|---|
| E63 | 6 | Ja, indirekt: J167 wird rot bei NICHT OK (über Allgemeine Angaben C103). | Keine direkte Regel erkannt. | "!" in E63 setzt C103 auf NICHT OK; J167/J184 reagieren. |
| E64 | 8 | Ja, indirekt: J168:J175 wird rot bei NICHT OK (über C104). | Keine direkte Regel erkannt. | "!" in E64 setzt C104 auf NICHT OK; J168/J184 reagieren. |
| E77 | 2 | Ja, indirekt: J175 wird rot bei NICHT OK. | Keine direkte Regel erkannt. | "!" in E77 oder E78 setzt J175 auf NICHT OK. |
| E78 | 2 | Ja, indirekt: J175 wird rot bei NICHT OK. | Keine direkte Regel erkannt. | "!" in E77 oder E78 setzt J175 auf NICHT OK. |
| W14 | 0 | Keine direkte CF-Regel. | Ja: blendet Prüfprotokoll-Zeilen 159:164. | Steuert Leistungsschutz-Block (P>, P>> usw.). |
| E75 | 2 | Keine direkte CF-Regel, aber Anzeige-/Textfelder im Prüfprotokoll ändern sich. | Keine direkte Regel erkannt. | Beeinflusst Abschluss-/Abschaltungszeilen. |
| D60 | 10 | Indirekt möglich über abhängige Zielzellen. | Keine direkte Regel erkannt. |  |
| D66 | 9 | Indirekt möglich über abhängige Zielzellen. | Keine direkte Regel erkannt. |  |
| H65 | 5 | Indirekt möglich über abhängige Zielzellen. | Keine direkte Regel erkannt. |  |
| H66 | 6 | Indirekt möglich über abhängige Zielzellen. | Keine direkte Regel erkannt. |  |
| C67 | 4 | Indirekt möglich über abhängige Zielzellen. | Keine direkte Regel erkannt. |  |
| C68 | 4 | Indirekt möglich über abhängige Zielzellen. | Keine direkte Regel erkannt. |  |

## Detail E63

- Direkte Formelziele: 6
  - Allgemeine Angaben!C103: =IF('Schutzprüf-Checkliste'!E63="x","x",IF('Schutzprüf-Checkliste'!E63="!","NICHT OK",""))
  - Prüfprotokoll!C124: =IF('Schutzprüf-Checkliste'!E63="x",'Schutzprüf-Checkliste'!H65 &"; I-Wandler: "&'Allgemeine Angaben'!C78&"A//"&"1A", "Keine Auslösung MSA")
  - Prüfprotokoll!C69: =IF('Schutzprüf-Checkliste'!E63="x",'Schutzprüf-Checkliste'!H65 &"; I-Wandler: "&'Allgemeine Angaben'!C78&"A//"&"1A", "Keine Auslösung MSA")
  - Prüfprotokoll!C74: =IF((G70="-"),"-",IF('Schutzprüf-Checkliste'!E63="x",'Schutzprüf-Checkliste'!H65&"; I-Wandler: "&'Allgemeine Angaben'!C78&"A//"&"1A","Keine Auslösung MSA"))
  - Prüfprotokoll!C79: =IF((G75="-"),"-",IF('Schutzprüf-Checkliste'!E63="x",'Schutzprüf-Checkliste'!H65&"; I-Wandler: "&'Allgemeine Angaben'!C78&"A//"&"1A","Keine Auslösung MSA"))
  - Prüfprotokoll!C84: =IF((G80="-"),"-",IF('Schutzprüf-Checkliste'!E63="x",'Schutzprüf-Checkliste'!H65&"; I-Wandler: "&'Allgemeine Angaben'!C78&"A//"&"1A","Keine Auslösung MSA"))

## Detail E64

- Direkte Formelziele: 8
  - Allgemeine Angaben!C104: =IF('Schutzprüf-Checkliste'!E64="x","x",IF('Schutzprüf-Checkliste'!E64="!","NICHT OK",""))
  - Allgemeine Angaben!D103: =IF('Schutzprüf-Checkliste'!E64="x", "", "x")
  - Prüfprotokoll!C105: =IF('Schutzprüf-Checkliste'!E64="x",'Schutzprüf-Checkliste'!H66,(IF('Schutzprüf-Checkliste'!D66="MS","Keine Auslösung MSA",IF('Schutzprüf-Checkliste'!D66="NS","Keine Auslösung NSHV-LS"))))
  - Prüfprotokoll!C117: =IF('Schutzprüf-Checkliste'!E64="x",'Schutzprüf-Checkliste'!H66,(IF('Schutzprüf-Checkliste'!D66="MS","Keine Auslösung MSA",IF('Schutzprüf-Checkliste'!D66="NS","Keine Auslösung NSHV-LS"))))
  - Prüfprotokoll!C158: =IF('Schutzprüf-Checkliste'!E64="x",'Schutzprüf-Checkliste'!H66,(IF('Schutzprüf-Checkliste'!D66="MS","Keine Auslösung MSA",IF('Schutzprüf-Checkliste'!D66="NS","Keine Auslösung NSHV-LS"))))
  - Prüfprotokoll!C164: =IF('Schutzprüf-Checkliste'!E64="x",'Schutzprüf-Checkliste'!H66,(IF('Schutzprüf-Checkliste'!D66="MS","Keine Auslösung MSA",IF('Schutzprüf-Checkliste'!D66="NS","Keine Auslösung NSHV-LS"))))
  - Prüfprotokoll!C92: =IF('Schutzprüf-Checkliste'!E64="x",'Schutzprüf-Checkliste'!H66,(IF('Schutzprüf-Checkliste'!D66="MS","Keine Auslösung MSA",IF('Schutzprüf-Checkliste'!D66="NS","Keine Auslösung NSHV-LS"))))
  - Prüfprotokoll!C98: =IF('Schutzprüf-Checkliste'!E64="x",'Schutzprüf-Checkliste'!H66,(IF('Schutzprüf-Checkliste'!D66="MS","Keine Auslösung MSA",IF('Schutzprüf-Checkliste'!D66="NS","Keine Auslösung NSHV-LS"))))

## Detail E77

- Direkte Formelziele: 2
  - Prüfprotokoll!A175: =IF(OR('Schutzprüf-Checkliste'!E77="x",'Schutzprüf-Checkliste'!E78="x",'Schutzprüf-Checkliste'!E77="!",'Schutzprüf-Checkliste'!E78="!"),"Trafoschutzfunktion(en) löst Schalter aus","")
  - Prüfprotokoll!J175: =IF(OR('Schutzprüf-Checkliste'!E77="!",'Schutzprüf-Checkliste'!E78="!"),"NICHT OK",IF(OR('Schutzprüf-Checkliste'!E77="x",'Schutzprüf-Checkliste'!E78="x"),"x",""))

## Detail E78

- Direkte Formelziele: 2
  - Prüfprotokoll!A175: =IF(OR('Schutzprüf-Checkliste'!E77="x",'Schutzprüf-Checkliste'!E78="x",'Schutzprüf-Checkliste'!E77="!",'Schutzprüf-Checkliste'!E78="!"),"Trafoschutzfunktion(en) löst Schalter aus","")
  - Prüfprotokoll!J175: =IF(OR('Schutzprüf-Checkliste'!E77="!",'Schutzprüf-Checkliste'!E78="!"),"NICHT OK",IF(OR('Schutzprüf-Checkliste'!E77="x",'Schutzprüf-Checkliste'!E78="x"),"x",""))

## Detail W14

- Keine direkten Formelziele gefunden.
- VBA-Bezug: Tabelle7.cls
- Regel: Bei W14 = False werden im Prüfprotokoll Zeilen 159:164 ausgeblendet, sonst eingeblendet.

## Bedingte Formatierung mit Warncharakter in Prüfprotokoll

- Bereich J167: Typ=cellIs, Operator=equal, Formel="NICHT OK"
- Bereich J168:J175: Typ=cellIs, Operator=equal, Formel="NICHT OK"
- Bereich A184:J184: Typ=expression, Operator=None, Formel=OR(J167="NICHT OK",J168="NICHT OK")

## Nächste Schritte für Änderungen

1. Änderungen zuerst in einer Kopie testen (vor allem E63/E64/E77/E78/W14).
2. Nach jeder Regeländerung die Felder J167, J168, J175 und A184:J184 kontrollieren.
3. Bei Änderungen rund um Leistungsschutz zusätzlich W14-Schalter und Zeilen 159:164 prüfen.