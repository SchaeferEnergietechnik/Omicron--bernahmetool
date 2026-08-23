# Excel-Analyse V20a

- Datei: samples/topics/excel-basis/V20a_Übergeordneter_Entkupplungsschutz.xlsm
- Analysezeitpunkt (UTC): 2026-08-23 09:49:07
- Umgebung: statische Analyse unter Linux mit openpyxl und oletools (kein Excel-COM).

## Kurzfazit

- Die Arbeitsmappe ist makrofähig und enthält VBA-Logik für Bereichs-Ein/Ausblendung.
- Im Blatt Prüfprotokoll existieren bedingte Formatierungsregeln, darunter Warnlogik für J167/J168 (NICHT OK).
- Es bestehen zahlreiche Formelabhängigkeiten von Schutzprüf-Checkliste nach Allgemeine Angaben und weiter nach Prüfprotokoll.
- Die Ausblendung des Leistungsschutz-Blocks (Zeilen 159-164) ist über die W14-Regel in VBA gekoppelt.

## Blattübersicht

| Blatt | Sichtbarkeit | Zeilen | Spalten | Nicht leer | Formeln | CF-Regeln | Datenvalidierungen |
|---|---:|---:|---:|---:|---:|---:|---:|
| Schutzprüf-Checkliste | sichtbar | 1220 | 25 | 491 | 70 | 6 | 1 |
| Bürdemessung | sichtbar | 30 | 7 | 126 | 22 | 0 | 0 |
| Allgemeine Angaben | sichtbar | 141 | 24 | 484 | 169 | 0 | 5 |
| Prüfprotokoll | sichtbar | 189 | 12 | 547 | 280 | 6 | 0 |
| Wandlerprüfprotokoll | sichtbar | 82 | 9 | 255 | 134 | 0 | 0 |
| Protokolluntergeord. Schutz EZE | sichtbar | 59 | 16 | 219 | 150 | 2 | 0 |
| Daten | sichtbar | 260 | 19 | 837 | 614 | 0 | 0 |
| Daten EZE | sichtbar | 199 | 22 | 525 | 374 | 0 | 0 |
| Measurements | sichtbar | 19 | 15 | 15 | 0 | 0 | 0 |
| Seq_Measurement | sichtbar | 19 | 13 | 13 | 0 | 0 | 0 |
| PlsRmp_Measurement | sichtbar | 10 | 14 | 14 | 0 | 0 | 0 |
| Wandlerdaten | sichtbar | 15 | 18 | 121 | 0 | 0 | 0 |
| Kunden | sichtbar | 38 | 1 | 38 | 0 | 0 | 0 |
| ProtNr | sichtbar | 2 | 4 | 6 | 2 | 0 | 0 |

## Abhängigkeitsgraph (Formeln zwischen Blättern)

| Quelle | Zielblatt | Anzahl Referenzen |
|---|---|---:|
| Allgemeine Angaben | Schutzprüf-Checkliste | 77 |
| Allgemeine Angaben | Daten | 8 |
| Allgemeine Angaben | ProtNr | 2 |
| Allgemeine Angaben | Daten EZE | 1 |
| Daten | Measurements | 224 |
| Daten | Seq_Measurement | 200 |
| Daten EZE | Seq_Measurement | 96 |
| Daten EZE | PlsRmp_Measurement | 40 |
| Daten EZE | Measurements | 35 |
| Daten EZE | Schutzprüf-Checkliste | 17 |
| Daten EZE | Allgemeine Angaben | 16 |
| Daten EZE | REF | 1 |
| ProtNr | Schutzprüf-Checkliste | 2 |
| Protokolluntergeord. Schutz EZE | Daten EZE | 142 |
| Protokolluntergeord. Schutz EZE | Schutzprüf-Checkliste | 88 |
| Protokolluntergeord. Schutz EZE | Allgemeine Angaben | 72 |
| Protokolluntergeord. Schutz EZE | Die Schutzeinrichtung ist nicht funktionstüchtig | 1 |
| Protokolluntergeord. Schutz EZE | Die Schutzeinrichtung ist funktionstüchtig | 1 |
| Prüfprotokoll | Allgemeine Angaben | 160 |
| Prüfprotokoll | Daten | 159 |
| Prüfprotokoll | Schutzprüf-Checkliste | 68 |
| Prüfprotokoll | Die Schutzeinrichtung ist nicht funktionstüchtig | 1 |
| Prüfprotokoll | Die Schutzeinrichtung ist funktionstüchtig | 1 |
| Schutzprüf-Checkliste | Spannungsfaktoren EZE anpassen | 1 |
| Wandlerprüfprotokoll | Allgemeine Angaben | 128 |
| Wandlerprüfprotokoll | Schutzprüf-Checkliste | 91 |
| Wandlerprüfprotokoll | Bürdemessung | 60 |

## Bedingte Formatierung (alle Regeln)

| Blatt | Bereich | Typ | Operator | Formel(n) |
|---|---|---|---|---|
| Schutzprüf-Checkliste | E3:E7 E9:E14 E19:E26 E28:E33 E38:E40 E44:E45 E47 E50:E51 E53:E58 E60:E80 E82:E92 E94:E125 E127:E132 E134:E135 E148:E1048576 | containsText | containsText | NOT(ISERROR(SEARCH("!",E3))) |
| Schutzprüf-Checkliste | E3:E7 E9:E14 E19:E26 E28:E33 E38:E40 E44:E45 E47 E50:E51 E53:E58 E60:E80 E82:E92 E94:E125 E127:E132 E134:E135 E148:E1048576 | containsText | containsText | NOT(ISERROR(SEARCH("?",E3))) |
| Schutzprüf-Checkliste | B56:B57 E1 E3:E7 E9:E14 E19:E26 E28:E33 E38:E40 E44:E45 E47 E50:E51 E53:E58 E60:E80 E82:E125 E127:E132 E134:E135 E148:E1048576 | containsText | containsText | NOT(ISERROR(SEARCH("x",B1))) |
| Schutzprüf-Checkliste | C134 | containsText | containsText | NOT(ISERROR(SEARCH("!",C134))) |
| Schutzprüf-Checkliste | C134 | containsText | containsText | NOT(ISERROR(SEARCH("?",C134))) |
| Schutzprüf-Checkliste | C134 | containsText | containsText | NOT(ISERROR(SEARCH("x",C134))) |
| Prüfprotokoll | J167 | cellIs | equal | "NICHT OK" |
| Prüfprotokoll | J168:J175 | cellIs | equal | "NICHT OK" |
| Prüfprotokoll | A184:J184 | expression |  | OR(J167="NICHT OK",J168="NICHT OK") |
| Prüfprotokoll | C92:J92 | cellIs | equal | "Keine Auslösung NSHV-LS" |
| Prüfprotokoll | C92:J92 | cellIs | equal | "Keine Auslösung MSA" |
| Prüfprotokoll | C92:J92 | cellIs | equal | FALSE |
| Protokolluntergeord. Schutz EZE | F50 | cellIs | equal | "NICHT OK" |
| Protokolluntergeord. Schutz EZE | A48:I48 | expression |  | F50="NICHT OK" |

## Relevante Bedingungslogik (Warnung/NICHT OK)

- Prüfprotokoll J167: CF-Regel rot bei Zellwert NICHT OK.
- Prüfprotokoll J168:J175: CF-Regel rot bei Zellwert NICHT OK.
- Prüfprotokoll A184:J184: Ausdrucksregel OR(J167="NICHT OK",J168="NICHT OK").
- Formelweg:
  - Prüfprotokoll J167 = Allgemeine Angaben C103
  - Allgemeine Angaben C103 = IF(Schutzprüf-Checkliste E63 = "!", "NICHT OK", ... )
  - Prüfprotokoll J168 = Allgemeine Angaben C104
  - Allgemeine Angaben C104 = IF(Schutzprüf-Checkliste E64 = "!", "NICHT OK", ... )

## Datenvalidierungen

| Blatt | Bereich | Typ | Operator | Formel1 | Formel2 | Leer erlaubt |
|---|---|---|---|---|---|---|
| Schutzprüf-Checkliste | A14:B14 |  |  |  |  | True |
| Allgemeine Angaben | C34 C55 C96 D133:D134 D136 | list |  | $F$103:$F$105 |  | True |
| Allgemeine Angaben | C97:C101 C103:D115 C125:D125 |  |  |  |  | True |
| Allgemeine Angaben | C72 |  |  |  |  | True |
| Allgemeine Angaben | C52:C53 C63:C66 C69:C70 C84:C86 |  |  |  |  | True |
| Allgemeine Angaben | C7 | list |  | "Gunnar Schäfer, Kevin Koehn" |  | True |

## Definierte Namen

- Keine definierten Namen gefunden.

## Formel-Referenzen aus Prüfprotokoll

- Referenzen auf Schutzprüf-Checkliste: 36
  - C9: ='Schutzprüf-Checkliste'!B6
  - H15: =IF(ISBLANK('Allgemeine Angaben'!D133),'Schutzprüf-Checkliste'!B11," ")
  - H16: =IF(ISBLANK('Allgemeine Angaben'!D134),'Schutzprüf-Checkliste'!B10," ")
  - H17: ='Schutzprüf-Checkliste'!B9
  - A20: ="Uc = "&'Schutzprüf-Checkliste'!D60&" kV"&IF('Schutzprüf-Checkliste'!H27="kein Trafo",""," - Trafostellung: "&'Schutzprüf-Checkliste'!D57&" UNS = "&'Schutzprüf-Checkliste'!D60/'Schutzprüf-Checkliste'!H27&" kV")
  - C69: =IF('Schutzprüf-Checkliste'!E63="x",'Schutzprüf-Checkliste'!H65 &"; I-Wandler: "&'Allgemeine Angaben'!C78&"A//"&"1A", "Keine Auslösung MSA")
  - C74: =IF((G70="-"),"-",IF('Schutzprüf-Checkliste'!E63="x",'Schutzprüf-Checkliste'!H65&"; I-Wandler: "&'Allgemeine Angaben'!C78&"A//"&"1A","Keine Auslösung MSA"))
  - C79: =IF((G75="-"),"-",IF('Schutzprüf-Checkliste'!E63="x",'Schutzprüf-Checkliste'!H65&"; I-Wandler: "&'Allgemeine Angaben'!C78&"A//"&"1A","Keine Auslösung MSA"))
  - C84: =IF((G80="-"),"-",IF('Schutzprüf-Checkliste'!E63="x",'Schutzprüf-Checkliste'!H65&"; I-Wandler: "&'Allgemeine Angaben'!C78&"A//"&"1A","Keine Auslösung MSA"))
  - C92: =IF('Schutzprüf-Checkliste'!E64="x",'Schutzprüf-Checkliste'!H66,(IF('Schutzprüf-Checkliste'!D66="MS","Keine Auslösung MSA",IF('Schutzprüf-Checkliste'!D66="NS","Keine Auslösung NSHV-LS"))))
  - C93: ="Uc = "&'Schutzprüf-Checkliste'!D60&" kV"
  - C98: =IF('Schutzprüf-Checkliste'!E64="x",'Schutzprüf-Checkliste'!H66,(IF('Schutzprüf-Checkliste'!D66="MS","Keine Auslösung MSA",IF('Schutzprüf-Checkliste'!D66="NS","Keine Auslösung NSHV-LS"))))
  - C99: ="Uc = "&'Schutzprüf-Checkliste'!D60&" kV"
  - C104: ='Allgemeine Angaben'!C94&"        Uc = "&'Schutzprüf-Checkliste'!D60&" kV"
  - C105: =IF('Schutzprüf-Checkliste'!E64="x",'Schutzprüf-Checkliste'!H66,(IF('Schutzprüf-Checkliste'!D66="MS","Keine Auslösung MSA",IF('Schutzprüf-Checkliste'!D66="NS","Keine Auslösung NSHV-LS"))))
  - C117: =IF('Schutzprüf-Checkliste'!E64="x",'Schutzprüf-Checkliste'!H66,(IF('Schutzprüf-Checkliste'!D66="MS","Keine Auslösung MSA",IF('Schutzprüf-Checkliste'!D66="NS","Keine Auslösung NSHV-LS"))))
  - C124: =IF('Schutzprüf-Checkliste'!E63="x",'Schutzprüf-Checkliste'!H65 &"; I-Wandler: "&'Allgemeine Angaben'!C78&"A//"&"1A", "Keine Auslösung MSA")
  - C157: =IF(ISBLANK('Allgemeine Angaben'!C96),"-","Uc = "&'Schutzprüf-Checkliste'!D60&" kV")
  - C158: =IF('Schutzprüf-Checkliste'!E64="x",'Schutzprüf-Checkliste'!H66,(IF('Schutzprüf-Checkliste'!D66="MS","Keine Auslösung MSA",IF('Schutzprüf-Checkliste'!D66="NS","Keine Auslösung NSHV-LS"))))
  - C164: =IF('Schutzprüf-Checkliste'!E64="x",'Schutzprüf-Checkliste'!H66,(IF('Schutzprüf-Checkliste'!D66="MS","Keine Auslösung MSA",IF('Schutzprüf-Checkliste'!D66="NS","Keine Auslösung NSHV-LS"))))
  - A167: =IF((OR('Allgemeine Angaben'!C103="x",'Allgemeine Angaben'!C103="NICHT OK")),"UMZ: Abschaltung MS-LS"&'Schutzprüf-Checkliste'!C67 &" im Fehlerfall geprüft (gesamte Wirkungskette)","")
  - A168: =IF(OR('Allgemeine Angaben'!C104="x",'Allgemeine Angaben'!C104="NICHT OK"),"Entkupplungsschutz: Abschaltung "&'Schutzprüf-Checkliste'!D66&"-LS"&'Schutzprüf-Checkliste'!C68 &" im Fehlerfall geprüft (gesamte Wirkungskette)","")
  - A169: =IF('Schutzprüf-Checkliste'!I70=TRUE,"Abschaltung MS-LS"&'Schutzprüf-Checkliste'!C67 &" nach Ausfall der Hilfsspannung - AuxDC","")
  - A170: =IF('Schutzprüf-Checkliste'!K70=TRUE,"Abschaltung NS-LS"&'Schutzprüf-Checkliste'!C68 &" nach Ausfall der Hilfsspannung- AuxDC","")
  - A171: =IF('Schutzprüf-Checkliste'!I71=TRUE,"Abschaltung MS-LS"&'Schutzprüf-Checkliste'!C67 &" nach Ausfall Schutzrelais (Live Contact)","")
  - A172: =IF('Schutzprüf-Checkliste'!K71=TRUE,"Abschaltung NS-LS"&'Schutzprüf-Checkliste'!C68 &" nach Ausfall Schutzrelais (Live Contact)","")
  - A173: =IF('Schutzprüf-Checkliste'!I72=TRUE,'Allgemeine Angaben'!A114,"")
  - A174: =IF('Schutzprüf-Checkliste'!K72=TRUE,'Allgemeine Angaben'!A115,"")
  - A175: =IF('Schutzprüf-Checkliste'!E77="x","Trafoschutzfunktion(en) löst " &"Schalter aus","")
  - J175: =IF('Schutzprüf-Checkliste'!E77="x","x","")
  - A180: =IF('Schutzprüf-Checkliste'!E75="x","Prüfklemmleiste vorhanden","")
  - J180: =IF('Schutzprüf-Checkliste'!E75="x","x","")
  - A182: =IF('Schutzprüf-Checkliste'!E74="x","DC USV für übergeordneter Schutz/ggfs. UMZ-Schutz in Ordnung","")
  - A183: =IF('Schutzprüf-Checkliste'!E79="x", "Ruhestromüberwachung bei räumlicher Trennung Schutz/Schalter vorhanden","")
  - D187: =IF('Allgemeine Angaben'!C8="x", 'Schutzprüf-Checkliste'!B4,"")
  - G187: ='Schutzprüf-Checkliste'!B7

- Referenzen auf Allgemeine Angaben: 116
  - C7: =IF('Allgemeine Angaben'!C7= "Kevin Koehn", 'Allgemeine Angaben'!U7,'Allgemeine Angaben'!H7)
  - F7: =IF('Allgemeine Angaben'!C8="x", IF(('Allgemeine Angaben'!C137=0),"","Ort der Prüfung:"),"")
  - H7: =IF('Allgemeine Angaben'!C8="x",IF(('Allgemeine Angaben'!C137=0),"",'Allgemeine Angaben'!C137),"")
  - C8: =IF('Allgemeine Angaben'!C7= "Kevin Koehn", 'Allgemeine Angaben'!G7,'Allgemeine Angaben'!F7)
  - H8: ='Allgemeine Angaben'!C138
  - I9: ='Allgemeine Angaben'!C139
  - H10: ='Allgemeine Angaben'!C5
  - A13: ='Allgemeine Angaben'!C2
  - H13: ='Allgemeine Angaben'!C3
  - H14: ='Allgemeine Angaben'!C140
  - H15: =IF(ISBLANK('Allgemeine Angaben'!D133),'Schutzprüf-Checkliste'!B11," ")
  - H16: =IF(ISBLANK('Allgemeine Angaben'!D134),'Schutzprüf-Checkliste'!B10," ")
  - H18: =IF(ISBLANK('Allgemeine Angaben'!D136),'Allgemeine Angaben'!C136," ")
  - H19: ='Allgemeine Angaben'!C4
  - C66: =IF(ISBLANK('Allgemeine Angaben'!C78),"-",TEXT(Daten!F27*'Allgemeine Angaben'!C78,"#,0#\ A"))
  - C68: =IF((G65="-"),"-",'Allgemeine Angaben'!C90)
  - C69: =IF('Schutzprüf-Checkliste'!E63="x",'Schutzprüf-Checkliste'!H65 &"; I-Wandler: "&'Allgemeine Angaben'!C78&"A//"&"1A", "Keine Auslösung MSA")
  - C71: =IF((G70="-"),"-",TEXT(Daten!F35*'Allgemeine Angaben'!C78,"#,0#\ A"))
  - C73: =IF((G70="-"),"-",'Allgemeine Angaben'!C91)
  - C74: =IF((G70="-"),"-",IF('Schutzprüf-Checkliste'!E63="x",'Schutzprüf-Checkliste'!H65&"; I-Wandler: "&'Allgemeine Angaben'!C78&"A//"&"1A","Keine Auslösung MSA"))
  - C76: =IF((G75="-"),"-",TEXT(Daten!F43*'Allgemeine Angaben'!C78,"#,0#\ A"))
  - C78: =IF((G75="-"),"-",IF(ISBLANK('Allgemeine Angaben'!C92),"-",('Allgemeine Angaben'!C92)))
  - C79: =IF((G75="-"),"-",IF('Schutzprüf-Checkliste'!E63="x",'Schutzprüf-Checkliste'!H65&"; I-Wandler: "&'Allgemeine Angaben'!C78&"A//"&"1A","Keine Auslösung MSA"))
  - C81: =IF((G80="-"),"-",TEXT(Daten!F49*'Allgemeine Angaben'!C78,"#,0#\ A"))
  - C83: =IF((G75="-"),"-",IF(ISBLANK('Allgemeine Angaben'!C93),"-",('Allgemeine Angaben'!C93)))
  - C84: =IF((G80="-"),"-",IF('Schutzprüf-Checkliste'!E63="x",'Schutzprüf-Checkliste'!H65&"; I-Wandler: "&'Allgemeine Angaben'!C78&"A//"&"1A","Keine Auslösung MSA"))
  - C89: =TEXT(Daten!I3/100*'Allgemeine Angaben'!$C$40,"0,0#\ kV")
  - C95: =TEXT(Daten!I11/100*'Allgemeine Angaben'!$C$40,"0,0#\ kV")
  - C101: =TEXT(Daten!I19/100*'Allgemeine Angaben'!$C$40,"#,0#\ kV")
  - C104: ='Allgemeine Angaben'!C94&"        Uc = "&'Schutzprüf-Checkliste'!D60&" kV"
  - C124: =IF('Schutzprüf-Checkliste'!E63="x",'Schutzprüf-Checkliste'!H65 &"; I-Wandler: "&'Allgemeine Angaben'!C78&"A//"&"1A", "Keine Auslösung MSA")
  - D128: =IF(ISBLANK('Allgemeine Angaben'!C96),"-",TEXT('Allgemeine Angaben'!C97, "0,0\ \M\W"))
  - H128: =IF(ISBLANK('Allgemeine Angaben'!C96),"-",'Allgemeine Angaben'!C98)
  - C130: =IF(ISBLANK('Allgemeine Angaben'!C96),"-",IF(H128="Leistungsschwelle", TEXT('Allgemeine Angaben'!C97*0.05*1000, "0,00\ \k\var"), "3°"))
  - D130: =IF(ISBLANK('Allgemeine Angaben'!C96),"-",IF(H128="Leistungsschwelle", TEXT('Allgemeine Angaben'!C97*0.05*100000/'Allgemeine Angaben'!C78/'Allgemeine Angaben'!C40, "0,000\ \var"), "3°"))
  - H130: =IF(ISBLANK('Allgemeine Angaben'!C96),"-",TEXT('Allgemeine Angaben'!C100*'Allgemeine Angaben'!C78, "0,00\ A"))
  - J130: =IF(ISBLANK('Allgemeine Angaben'!C96),"-",TEXT('Allgemeine Angaben'!C100, "0,00\ A"))
  - C131: =IF(ISBLANK('Allgemeine Angaben'!C96),"-",'Allgemeine Angaben'!C40*'Allgemeine Angaben'!C99/100 & " kV")
  - D131: =IF(ISBLANK('Allgemeine Angaben'!C96),"-",TEXT('Allgemeine Angaben'!C99,"0,0\ \V"))
  - J131: =IF(ISBLANK('Allgemeine Angaben'!C96),"-",TEXT('Allgemeine Angaben'!C101, "0,0\ \s"))
  - C137: =IF(ISBLANK('Allgemeine Angaben'!C96),"-",TEXT(Daten!L55, "0,00\ \M\W"))
  - D137: =IF(ISBLANK('Allgemeine Angaben'!C96),"-",TEXT(Daten!L56, "0,00\ \M\W"))
  - E137: =IF(ISBLANK('Allgemeine Angaben'!C96),"-","0,1 V")
  - G137: =IF(ISBLANK('Allgemeine Angaben'!C96),"-",Daten!L57)
  - H137: =IF(ISBLANK('Allgemeine Angaben'!C96),"-",Daten!L58)
  - J137: =IF(ISBLANK('Allgemeine Angaben'!C96),"-","0,001 A")
  - E140: =IF(ISBLANK('Allgemeine Angaben'!C96),"-","0,05 °")
  - J140: =IF(ISBLANK('Allgemeine Angaben'!C96),"-","0,1 var")
  - E141: =IF(ISBLANK('Allgemeine Angaben'!C96),"-","0,05 °")
  - J141: =IF(ISBLANK('Allgemeine Angaben'!C96),"-","0,1 var")
  - D146: =IF(ISBLANK('Allgemeine Angaben'!$C$96),"-","Nein")
  - F146: =IF(ISBLANK('Allgemeine Angaben'!$C$96),"-","x")
  - G146: =IF(ISBLANK('Allgemeine Angaben'!$C$96),"-",'Allgemeine Angaben'!C99*0.9648)
  - H146: =IF(ISBLANK('Allgemeine Angaben'!$C$96),"-",'Allgemeine Angaben'!C99*1.07)
  - J146: =IF(ISBLANK('Allgemeine Angaben'!$C$96),"-",'Allgemeine Angaben'!C99*1.07)
  - D147: =IF(ISBLANK('Allgemeine Angaben'!$C$96),"-","Nein")
  - F147: =IF(ISBLANK('Allgemeine Angaben'!$C$96),"-","x")
  - G147: =IF(ISBLANK('Allgemeine Angaben'!$C$96),"-",'Allgemeine Angaben'!C99*1.07)
  - H147: =IF(ISBLANK('Allgemeine Angaben'!$C$96),"-",'Allgemeine Angaben'!C99*0.9648)
  - J147: =IF(ISBLANK('Allgemeine Angaben'!$C$96),"-",'Allgemeine Angaben'!C99*1.07)
  - D148: =IF(ISBLANK('Allgemeine Angaben'!$C$96),"-","Nein")
  - F148: =IF(ISBLANK('Allgemeine Angaben'!$C$96),"-","x")
  - G148: =IF(ISBLANK('Allgemeine Angaben'!$C$96),"-",'Allgemeine Angaben'!C99*1.07)
  - H148: =IF(ISBLANK('Allgemeine Angaben'!$C$96),"-",'Allgemeine Angaben'!C99*1.07)
  - J148: =IF(ISBLANK('Allgemeine Angaben'!$C$96),"-",'Allgemeine Angaben'!C99*0.9648)
  - D149: =IF(ISBLANK('Allgemeine Angaben'!$C$96),"-","Nein")
  - F149: =IF(ISBLANK('Allgemeine Angaben'!$C$96),"-","x")
  - G149: =IF(ISBLANK('Allgemeine Angaben'!$C$96),"-",'Allgemeine Angaben'!C99*0.9648)
  - H149: =IF(ISBLANK('Allgemeine Angaben'!$C$96),"-",'Allgemeine Angaben'!C99*1.176)
  - J149: =IF(ISBLANK('Allgemeine Angaben'!$C$96),"-",'Allgemeine Angaben'!C99*0.9648)
  - D150: =IF(ISBLANK('Allgemeine Angaben'!$C$96),"-","Ja")
  - F150: =IF(ISBLANK('Allgemeine Angaben'!$C$96),"-","x")
  - G150: =IF(ISBLANK('Allgemeine Angaben'!$C$96),"-",'Allgemeine Angaben'!C99*0.9648)
  - H150: =IF(ISBLANK('Allgemeine Angaben'!$C$96),"-",'Allgemeine Angaben'!C99*0.9648)
  - J150: =IF(ISBLANK('Allgemeine Angaben'!$C$96),"-",'Allgemeine Angaben'!C99*0.9648)
  - C152: =IF(ISBLANK('Allgemeine Angaben'!$C$96),"-",'Allgemeine Angaben'!C99*0.9648)
  - D153: =IF(ISBLANK('Allgemeine Angaben'!$C$96),"-","0,5 A")
  - E153: =IF(ISBLANK('Allgemeine Angaben'!$C$96),"-","135")
  - G153: =IF(ISBLANK('Allgemeine Angaben'!$C$96),"-","15")
  - H153: =IF(ISBLANK('Allgemeine Angaben'!$C$96),"-","255")

## Formel-Referenzen aus Allgemeine Angaben

- Referenzen auf Schutzprüf-Checkliste: 51
  - C3: ='Schutzprüf-Checkliste'!B3
  - C16: ='Schutzprüf-Checkliste'!L14
  - C17: ='Schutzprüf-Checkliste'!L14
  - C18: ='Schutzprüf-Checkliste'!N14
  - C19: ='Schutzprüf-Checkliste'!N14
  - C20: ='Schutzprüf-Checkliste'!L16
  - C21: ='Schutzprüf-Checkliste'!L17
  - C22: ='Schutzprüf-Checkliste'!J17
  - C23: ='Schutzprüf-Checkliste'!J15
  - C24: ='Schutzprüf-Checkliste'!L15
  - C25: ='Schutzprüf-Checkliste'!J16
  - C40: ='Schutzprüf-Checkliste'!D60
  - C52: ='Schutzprüf-Checkliste'!D60*10&"V L-N"
  - C53: =IF('Schutzprüf-Checkliste'!E112="x", "Zählerschrank; ","")& IF('Schutzprüf-Checkliste'!E115="x", " 2. Wicklung; ","")& IF('Schutzprüf-Checkliste'!E118="x", " Schutz-Wicklung; ","")& IF('Schutzprüf-Checkliste'!E121="x", " da-dn (Klemmen)","")
  - C68: =IF('Schutzprüf-Checkliste'!D29/2<61,'Schutzprüf-Checkliste'!D29/2 & " A", "50 A")
  - C69: =IF('Schutzprüf-Checkliste'!E96="x", "Zählerschrank; ","")& IF('Schutzprüf-Checkliste'!E99="x", " 2. Kern; ","")& IF('Schutzprüf-Checkliste'!E102="x", " Schutzkern","")
  - C72: ='Schutzprüf-Checkliste'!J38
  - C78: ='Schutzprüf-Checkliste'!D61
  - C87: =IF(C78/2<61, 'Schutzprüf-Checkliste'!D61/2 & " A", "50 A")
  - C97: =IF(Daten!F59="-","",Daten!F59*20*('Schutzprüf-Checkliste'!D61*'Schutzprüf-Checkliste'!D60)/100000)
  - C103: =IF('Schutzprüf-Checkliste'!E63="x","x",IF('Schutzprüf-Checkliste'!E63="!","NICHT OK",""))
  - D103: =IF('Schutzprüf-Checkliste'!E64="x", "", "x")
  - C104: =IF('Schutzprüf-Checkliste'!E64="x","x",IF('Schutzprüf-Checkliste'!E64="!","NICHT OK",""))
  - C105: ='Schutzprüf-Checkliste'!J70
  - D105: =IF('Schutzprüf-Checkliste'!I70=FALSE, "x", "")
  - C106: ='Schutzprüf-Checkliste'!L70
  - C107: ='Schutzprüf-Checkliste'!J71
  - D107: =IF('Schutzprüf-Checkliste'!I71=FALSE, "x", "")
  - C108: ='Schutzprüf-Checkliste'!L71
  - C109: =IF('Schutzprüf-Checkliste'!E74="x","x","")
  - D109: =IF('Schutzprüf-Checkliste'!E74="x", "", "x")
  - C110: =IF('Schutzprüf-Checkliste'!E91="x","x","")
  - A114: =IF('Schutzprüf-Checkliste'!I72=TRUE, "Abschaltung Leistungsschalter " &'Schutzprüf-Checkliste'!D65&'Schutzprüf-Checkliste'!C67&  " bei Ausfall Messspannung (Spannungswandlerschutzschalter)","")
  - C114: ='Schutzprüf-Checkliste'!J72
  - D114: =IF('Schutzprüf-Checkliste'!I72=FALSE, "x", "")
  - A115: =IF('Schutzprüf-Checkliste'!K72=TRUE, "Abschaltung Leistungsschalter " &'Schutzprüf-Checkliste'!D66&'Schutzprüf-Checkliste'!C68&  " bei Ausfall Messspannung (Spannungswandlerschutzschalter)","")
  - C115: ='Schutzprüf-Checkliste'!L72
  - C117: =IF('Schutzprüf-Checkliste'!$K$24=TRUE,'Allgemeine Angaben'!H117,IF('Schutzprüf-Checkliste'!$Q$24="WAHR",'Allgemeine Angaben'!G117, IF('Schutzprüf-Checkliste'!$I$24=TRUE,'Allgemeine Angaben'!F117,"Eingabe falsch")))
  - C119: =IF(C118="x",I119,IF('Schutzprüf-Checkliste'!$K$24=TRUE,'Allgemeine Angaben'!H119,IF('Schutzprüf-Checkliste'!$Q$24="WAHR",'Allgemeine Angaben'!G119,IF('Schutzprüf-Checkliste'!$I$24=TRUE,'Allgemeine Angaben'!F119,"Eingabe falsch"))))
  - C121: =IF('Schutzprüf-Checkliste'!$K$24=TRUE,'Allgemeine Angaben'!H121,IF('Schutzprüf-Checkliste'!$Q$24="WAHR",'Allgemeine Angaben'!G121, IF('Schutzprüf-Checkliste'!$I$24=TRUE,'Allgemeine Angaben'!F121,"Eingabe falsch")))
  - C122: =IF('Schutzprüf-Checkliste'!$K$24=TRUE,'Allgemeine Angaben'!H122,IF('Schutzprüf-Checkliste'!$Q$24="WAHR",'Allgemeine Angaben'!G122, IF('Schutzprüf-Checkliste'!$I$24=TRUE,'Allgemeine Angaben'!F122,"Eingabe falsch")))
  - C125: =IF('Schutzprüf-Checkliste'!E82="x","x",(IF('Schutzprüf-Checkliste'!E82="!","NICHT OK","")))
  - D125: =IF('Schutzprüf-Checkliste'!E82="x","","x")
  - C133: ='Schutzprüf-Checkliste'!C21
  - C134: ='Schutzprüf-Checkliste'!C20
  - C135: ='Schutzprüf-Checkliste'!B9:D9
  - C136: ='Schutzprüf-Checkliste'!B12
  - C137: ='Schutzprüf-Checkliste'!B4
  - C138: ='Schutzprüf-Checkliste'!B7
  - C140: ="Schutzschrank, "&IF(ISBLANK('Schutzprüf-Checkliste'!D65),'Schutzprüf-Checkliste'!D66,'Schutzprüf-Checkliste'!D65&IF(ISBLANK('Schutzprüf-Checkliste'!D66),"",", "&'Schutzprüf-Checkliste'!D66))
  - C141: ="NSHV; " &'Schutzprüf-Checkliste'!H26

## VBA-Überblick

- VBA-Module gesamt: 20
- Module: DieseArbeitsmappe.cls, Tabelle1.cls, Tabelle2.cls, Tabelle3.cls, Tabelle4.cls, Tabelle5.cls, Tabelle6.cls, Tabelle7.cls, UserForm1.frm, Tabelle8.cls, Tabelle9.cls, Datumsetzen.frm, Tabelle13.cls, Tabelle11.cls, Tabelle12.cls, Tabelle10.cls, Tabelle14.cls, frmPDFDruck.frm, Modul2.bas, Modul1.bas

| Modul | Sichtbarkeit | Typ | Prozedur |
|---|---|---|---|
| Datumsetzen.frm | Default | Sub | Datum |
| Modul1.bas | Private | Sub | CommandButton1_Click |
| Modul1.bas | Public | Sub | BereicheEinOderAusblenden_Start |
| Tabelle1.cls | Default | Sub | Protokollnummer_generieren_unsichtbar |
| Tabelle1.cls | Default | Sub | SchreibeNeunTexte |
| Tabelle1.cls | Default | Sub | SchreibeVierTexte |
| Tabelle1.cls | Private | Sub | Worksheet_SelectionChange |
| Tabelle1.cls | Private | Sub | ommandButton1_Click |
| Tabelle7.cls | Default | Sub | StartePDFDruck |
| Tabelle7.cls | Default | Sub | ZeilenAusblendenWennLeer |
| Tabelle7.cls | Private | Sub | CommandButton1_Click |
| Tabelle7.cls | Private | Sub | Worksheet_SelectionChange |
| Tabelle9.cls | Private | Sub | CheckBox1_Click |
| Tabelle9.cls | Private | Sub | CommandButton2_Click |
| Tabelle9.cls | Private | Sub | Worksheet_SelectionChange |
| frmPDFDruck.frm | Default | Function | DateiMitVersion |
| frmPDFDruck.frm | Private | Sub | cmdOK_Click |
| frmPDFDruck.frm | Private | Sub | txtVersion_Change |

### Makro-Kandidaten für den Ablauf

- Tabelle1.cls: Default Sub Protokollnummer_generieren_unsichtbar
- Tabelle7.cls: Private Sub CommandButton1_Click
- Tabelle7.cls: Default Sub ZeilenAusblendenWennLeer
- Modul1.bas: Public Sub BereicheEinOderAusblenden_Start
- Modul1.bas: Private Sub CommandButton1_Click

## Spezifischer Punkt Leistungsschutz/Pave

- Im Blatt Prüfprotokoll liegt der Leistungsschutz-Block in den Zeilen 159-164 (inkl. P> und P>>).
- Die Ein-/Ausblendung dieses Blocks erfolgt über VBA und hängt an Schutzprüf-Checkliste W14.
- Bei W14 = False wird 159:164 ausgeblendet, bei True eingeblendet.

## Grenzen der Analyse

- Keine Berechnung in Excel ausgeführt, daher sind gecachte Ergebniswerte teilweise leer oder veraltet.
- ActiveX- und Formularsteuerelemente wurden nur indirekt über Formeln/VBA beurteilt.
- Visuelle Styles (exakte Farbwerte der CF) sind nicht vollständig extrahiert.
