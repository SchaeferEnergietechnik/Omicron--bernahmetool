from __future__ import annotations

import argparse
import os
import shutil
import sys
import tempfile
import time
from pathlib import Path

from openpyxl import load_workbook

DEFAULT_TARGETS_REAL = [
    Path("samples/V20d_Übergeordneter_Entkupplungsschutz.xlsm"),
    Path("samples/topics/excel-basis/V20d_Übergeordneter_Entkupplungsschutz.xlsm"),
]

DEFAULT_TARGETS_FALLBACK = [
    Path("samples/V20d_Uebergeordneter_Entkupplungsschutz.xlsm"),
    Path("samples/topics/excel-basis/V20d_Uebergeordneter_Entkupplungsschutz.xlsm"),
    Path("V20d_Übergeordneter_Entkupplungsschutz.xlsm"),
    Path("V20d_Uebergeordneter_Entkupplungsschutz.xlsm"),
]

DEFAULT_SOURCE = Path("samples/topics/excel-basis/Muster_Termine 17.08.2026.xlsx")
CUSTOMER_SHEET = "Kunden"
SOURCE_CANDIDATE_SHEETS = ("Kundenadressen", "Kunden")
LOGO_REFERENCE_CANDIDATES = [
    Path("samples/V19m_Übergeordneter_Entkupplungsschutz.xlsm"),
    Path("samples/V19g_Übergeordneter_Entkupplungsschutz.xlsm"),
    Path("samples/topics/V19m_Übergeordneter_Entkupplungsschutz.xlsm"),
    Path("samples/topics/excel-basis/V19m_Übergeordneter_Entkupplungsschutz.xlsm"),
]

FORM_CODE = r'''Option Explicit
Private Const EMAIL_CHECKBOX_NAME As String = "chkDirektEmail"

Private Sub UserForm_Initialize()
    EnsureEmailCheckbox
End Sub

Private Sub EnsureEmailCheckbox()
    Dim ctrl As Object
    On Error Resume Next
    Set ctrl = Me.Controls(EMAIL_CHECKBOX_NAME)
    On Error GoTo 0

    If ctrl Is Nothing Then
        Set ctrl = Me.Controls.Add("Forms.CheckBox.1", EMAIL_CHECKBOX_NAME, True)
        With ctrl
            .Caption = "PDF direkt per E-Mail vorbereiten"
            .Left = 18
            .Top = 166
            .Width = 220
            .Height = 18
            .Value = False
        End With
    End If
End Sub

Private Sub cmdOK_Click()
    Dim wb As Workbook
    Set wb = ThisWorkbook

    Dim pfad As String
    Dim dateiname As String
    Dim vollDateiname As String
    Dim vollPfad As String
    Dim schutz As String
    Dim stationPrefix As String
    Dim wsProt As Worksheet
    Dim exportiertePdfs As Collection
    Set exportiertePdfs = New Collection

    pfad = wb.Path & "\\"

    If optEntkupplung.Value = True Then
        schutz = "Entkupplungsschutz"
    ElseIf optUMZ.Value = True Then
        schutz = "UMZ-Schutz"
    Else
        MsgBox "Bitte eine Schutzart auswaehlen!", vbExclamation
        Exit Sub
    End If

    Set wsProt = wb.Sheets("Prüfprotokoll")
    If IstStationGesperrt(wsProt) Then
        stationPrefix = "Station-gesperrt_"
    Else
        stationPrefix = ""
    End If

    If chkBlatt1.Value = True Then
        dateiname = stationPrefix & wb.Sheets("Allgemeine Angaben").Range("C5").Value & "_" & _
                    wb.Sheets("Prüfprotokoll").Range("H13").Value & "_" & _
                    IIf(schutz = "Entkupplungsschutz", "Übergeordneter-Schutz", "UMZ-Schutz")

        vollDateiname = DateiMitVersion(pfad, dateiname)
        vollPfad = pfad & vollDateiname

        wb.Sheets("Prüfprotokoll").ExportAsFixedFormat Type:=xlTypePDF, _
            Filename:=vollPfad, Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False

        exportiertePdfs.Add vollPfad
    End If

    If chkBlatt2.Value = True Then
        dateiname = stationPrefix & wb.Sheets("Allgemeine Angaben").Range("C5").Value & "_" & _
                    wb.Sheets("Prüfprotokoll").Range("H13").Value & "_Wandlerprüfprotokoll"

        vollDateiname = DateiMitVersion(pfad, dateiname)
        vollPfad = pfad & vollDateiname

        wb.Sheets("Wandlerprüfprotokoll").ExportAsFixedFormat Type:=xlTypePDF, _
            Filename:=vollPfad, Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False

        exportiertePdfs.Add vollPfad
    End If

    If chkBlatt3.Value = True Then
        dateiname = stationPrefix & wb.Sheets("Allgemeine Angaben").Range("C6").Value & "_" & _
                    wb.Sheets("Prüfprotokoll").Range("H13").Value & "_Untergeordneter-Schutz"

        vollDateiname = DateiMitVersion(pfad, dateiname)
        vollPfad = pfad & vollDateiname

        wb.Sheets("Protokolluntergeord. Schutz EZE").ExportAsFixedFormat Type:=xlTypePDF, _
            Filename:=vollPfad, Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False

        exportiertePdfs.Add vollPfad
    End If

    If exportiertePdfs.Count = 0 Then
        MsgBox "Es wurde kein Blatt fuer den PDF-Export ausgewaehlt.", vbExclamation
        Exit Sub
    End If

    MsgBox "PDF(s) erfolgreich erstellt im Ordner: " & pfad

    If IsEmailCheckboxSelected() Then
        SendePdfsPerOutlook exportiertePdfs
    End If

    Unload Me
End Sub

Private Function IsEmailCheckboxSelected() As Boolean
    Dim ctrl As Object

    On Error Resume Next
    Set ctrl = Me.Controls(EMAIL_CHECKBOX_NAME)
    On Error GoTo 0

    If ctrl Is Nothing Then
        IsEmailCheckboxSelected = False
    Else
        IsEmailCheckboxSelected = CBool(ctrl.Value)
    End If
End Function

Private Sub SendePdfsPerOutlook(ByVal exportiertePdfs As Collection)
    Dim wb As Workbook
    Dim wsAngaben As Worksheet
    Dim wsKunden As Worksheet
    Dim wsCheck As Worksheet
    Dim kundeText As String
    Dim emailTo As String
    Dim bemerkungen As String
    Dim betreff As String
    Dim projektname As String
    Dim protokollLabel As String
    Dim bodyText As String
    Dim htmlBody As String
    Dim logoPath As String

    Set wb = ThisWorkbook
    Set wsAngaben = wb.Sheets("Allgemeine Angaben")
    Set wsKunden = wb.Sheets("Kunden")
    Set wsCheck = wb.Sheets("Schutzprüf-Checkliste")

    kundeText = CStr(wsAngaben.Range("C2").Value)
    emailTo = HoleEmailAusKundenblatt(wsKunden, kundeText)

    If Trim$(emailTo) = "" Then
        MsgBox "Keine E-Mail-Adresse fuer den Kunden gefunden. Bitte im Blatt 'Kunden' in Spalte B pflegen.", vbExclamation
        Exit Sub
    End If

    bemerkungen = HoleBemerkungenAusCheckliste(wsCheck)

    projektname = Trim$(CStr(wsCheck.Range("B3").Value))
    If projektname = "" Then
        projektname = Trim$(CStr(wsAngaben.Range("C5").Value))
    End If
    protokollLabel = IIf(exportiertePdfs.Count = 1, "Schutzprüfprotokoll", "Schutzprüfprotokolle")
    betreff = protokollLabel & IIf(projektname <> "", " - " & projektname, "")
    bodyText = ErzeugeStandardMailtext(kundeText, bemerkungen, projektname, exportiertePdfs.Count)
    logoPath = FindeLogoPfad(wb.Path)
    htmlBody = ErzeugeHtmlMailtext(kundeText, bemerkungen, logoPath, projektname, exportiertePdfs.Count)

    Dim olApp As Object
    Dim olMail As Object
    Dim i As Long

    On Error GoTo outlook_fehler
    Set olApp = CreateObject("Outlook.Application")
    Set olMail = olApp.CreateItem(0)

    With olMail
        .To = emailTo
        .Subject = betreff
        .Body = bodyText

        For i = 1 To exportiertePdfs.Count
            If Len(Dir(CStr(exportiertePdfs(i)))) > 0 Then
                .Attachments.Add CStr(exportiertePdfs(i))
            End If
        Next i

        If logoPath <> "" Then
            Dim logoAttachment As Object
            Set logoAttachment = .Attachments.Add(logoPath)
            logoAttachment.PropertyAccessor.SetProperty _
                "http://schemas.microsoft.com/mapi/proptag/0x3712001F", "geslogo"
        End If

        .HTMLBody = htmlBody

        .Display
    End With

    MsgBox "Outlook-E-Mail wurde vorbereitet.", vbInformation
    Exit Sub

outlook_fehler:
    MsgBox "Outlook konnte nicht gestartet werden oder der Entwurf konnte nicht erstellt werden.", vbCritical
End Sub

Private Function ErzeugeStandardMailtext(ByVal kundeText As String, ByVal bemerkungen As String, ByVal projektname As String, ByVal anzahlPdfs As Long) As String
    Dim text As String
    Dim protokollLabel As String
    Dim projektTeil As String

    protokollLabel = IIf(anzahlPdfs = 1, "Schutzprüfprotokoll", "Schutzprüfprotokolle")
    If Trim$(projektname) <> "" Then
        projektTeil = " zum Projekt """ & projektname & """"
    Else
        projektTeil = ""
    End If

    text = "Sehr geehrte Damen und Herren," & vbCrLf & vbCrLf & _
        "anbei erhalten Sie " & IIf(anzahlPdfs = 1, "das aktuelle ", "die aktuellen ") & protokollLabel & projektTeil & " als PDF-Anhang." & vbCrLf & _
        "Bitte prüfen Sie die Unterlagen und melden Sie sich gerne bei Rückfragen." & vbCrLf & vbCrLf

    If Trim$(bemerkungen) <> "" Then
        text = text & "Bemerkungen aus der Checkliste:" & vbCrLf & bemerkungen & vbCrLf & vbCrLf
    End If

    text = text & "Herzlichen Dank." & vbCrLf & vbCrLf & _
        "Mit freundlichen Grüßen" & vbCrLf & vbCrLf & _
        "Ihr Team von G.E.S. Energietechnik GmbH" & vbCrLf & vbCrLf & _
        "Der Inhalt dieser Email ist ausschließlich für den bezeichneten Adressaten bestimmt." & vbCrLf & _
        "Falls Sie nicht der vorgesehene Adressat dieser Email oder dessen Vertreter sein sollten," & vbCrLf & _
        "so beachten Sie bitte, dass jede Form der Kenntnisnahme, Veröffentlichung," & vbCrLf & _
        "Vervielfältigung oder Weitergabe des Inhalts dieser Email unzulässig ist." & vbCrLf & _
        "Wir bitten Sie, sich in diesem Fall mit dem Absender der Email in Verbindung zu setzen." & vbCrLf & vbCrLf & _
        "G.E.S. Energietechnik GmbH: Sitz der Gesellschaft: Altmärkische Wische - Amtsgericht Stendal," & vbCrLf & _
        "HRB 30020 - Geschäftsführer: Gunnar Schäfer"

    ErzeugeStandardMailtext = text
End Function

Private Function ErzeugeHtmlMailtext(ByVal kundeText As String, ByVal bemerkungen As String, ByVal logoPath As String, ByVal projektname As String, ByVal anzahlPdfs As Long) As String
    Dim text As String
    Dim bemerkHtml As String
    Dim protokollLabel As String
    Dim projektTeil As String

    bemerkHtml = Replace(HTMLEncode(bemerkungen), vbCrLf, "<br>")
    protokollLabel = IIf(anzahlPdfs = 1, "Schutzprüfprotokoll", "Schutzprüfprotokolle")
    If Trim$(projektname) <> "" Then
        projektTeil = " zum Projekt &bdquo;" & HTMLEncode(projektname) & "&ldquo;"
    Else
        projektTeil = ""
    End If

    text = "<html><body style='font-family:Calibri,Arial,sans-serif;font-size:11pt;'>" & _
           "<p>Sehr geehrte Damen und Herren,</p>" & _
             "<p>anbei erhalten Sie " & IIf(anzahlPdfs = 1, "das aktuelle ", "die aktuellen ") & HTMLEncode(protokollLabel) & projektTeil & " als PDF-Anhang.<br>" & _
            "Bitte prüfen Sie die Unterlagen und melden Sie sich gerne bei Rückfragen.</p>"

    If Trim$(bemerkungen) <> "" Then
        text = text & "<p><b>Bemerkungen aus der Checkliste:</b><br>" & bemerkHtml & "</p>"
    End If

    text = text & "<p>Herzlichen Dank.</p>" & _
                  "<p>Mit freundlichen Grüßen</p>" & _
                  "<p>Ihr Team von G.E.S. Energietechnik GmbH</p>" & _
                  "<p style='font-size:9pt;color:#555;'>" & _
                  "Der Inhalt dieser Email ist ausschließlich für den bezeichneten Adressaten bestimmt.<br>" & _
                  "Falls Sie nicht der vorgesehene Adressat dieser Email oder dessen Vertreter sein sollten,<br>" & _
                  "so beachten Sie bitte, dass jede Form der Kenntnisnahme, Veröffentlichung,<br>" & _
                  "Vervielfältigung oder Weitergabe des Inhalts dieser Email unzulässig ist.<br>" & _
                  "Wir bitten Sie, sich in diesem Fall mit dem Absender der Email in Verbindung zu setzen.<br><br>" & _
                  "G.E.S. Energietechnik GmbH: Sitz der Gesellschaft: Altmärkische Wische - Amtsgericht Stendal,<br>" & _
                  "HRB 30020 - Geschäftsführer: Gunnar Schäfer" & _
                  "</p>"

    If logoPath <> "" Then
        text = text & "<p><img src='cid:geslogo' style='max-width:220px;height:auto;'></p>"
    End If

    text = text & "</body></html>"
    ErzeugeHtmlMailtext = text
End Function

Private Function HTMLEncode(ByVal text As String) As String
    Dim t As String
    t = text
    t = Replace(t, "&", "&amp;")
    t = Replace(t, "<", "&lt;")
    t = Replace(t, ">", "&gt;")
    t = Replace(t, Chr(34), "&quot;")
    HTMLEncode = t
End Function

Private Function IstStationGesperrt(ByVal wsProt As Worksheet) As Boolean
    Dim zeile As Long
    Dim statusWert As String

    For zeile = 167 To 175
        statusWert = Trim$(CStr(wsProt.Cells(zeile, "J").Value))
        If StrComp(statusWert, "NICHT OK", vbTextCompare) = 0 Then
            IstStationGesperrt = True
            Exit Function
        End If
    Next zeile

    IstStationGesperrt = False
End Function

Private Function FindeLogoPfad(ByVal basisPfad As String) As String
    Dim kandidaten(1 To 7) As String
    Dim i As Long

    kandidaten(1) = basisPfad & "\\logo.png"
    kandidaten(2) = basisPfad & "\\Logo.png"
    kandidaten(3) = basisPfad & "\\logo.jpg"
    kandidaten(4) = basisPfad & "\\Logo.jpg"
    kandidaten(5) = basisPfad & "\\logo.jpeg"
    kandidaten(6) = basisPfad & "\\ges-logo.png"
    kandidaten(7) = basisPfad & "\\GES_Logo.png"

    For i = 1 To 7
        If Len(Dir(kandidaten(i))) > 0 Then
            FindeLogoPfad = kandidaten(i)
            Exit Function
        End If
    Next i
End Function

Private Function HoleBemerkungenAusCheckliste(ByVal ws As Worksheet) As String
    Dim zeile As Long
    Dim eintrag As String
    Dim gesammelt As String

    ' Unteres Bemerkungsfeld in V19m/V20a: B138:B147
    For zeile = 138 To 147
        eintrag = Trim$(CStr(ws.Cells(zeile, "B").Value))
        If eintrag <> "" Then
            If gesammelt <> "" Then
                gesammelt = gesammelt & vbCrLf
            End If
            gesammelt = gesammelt & CStr(zeile - 137) & ". " & eintrag
        End If
    Next zeile

    HoleBemerkungenAusCheckliste = gesammelt
End Function

Private Function HoleEmailAusKundenblatt(ByVal ws As Worksheet, ByVal kundenwert As String) As String
    Dim suchKey As String
    Dim zeile As Long
    Dim kandidat As String
    Dim kandidatKey As String
    Dim email As String

    suchKey = NormalisiereKundenname(kundenwert)
    If suchKey = "" Then
        Exit Function
    End If

    For zeile = 1 To ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        kandidat = CStr(ws.Cells(zeile, "A").Value)
        kandidatKey = NormalisiereKundenname(kandidat)
        email = Trim$(CStr(ws.Cells(zeile, "B").Value))

        If email <> "" Then
            If kandidatKey = suchKey Then
                HoleEmailAusKundenblatt = email
                Exit Function
            End If
        End If
    Next zeile

    For zeile = 1 To ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        kandidat = CStr(ws.Cells(zeile, "A").Value)
        kandidatKey = NormalisiereKundenname(kandidat)
        email = Trim$(CStr(ws.Cells(zeile, "B").Value))

        If email <> "" Then
            If (InStr(1, kandidatKey, suchKey, vbTextCompare) > 0) Or _
               (InStr(1, suchKey, kandidatKey, vbTextCompare) > 0) Then
                HoleEmailAusKundenblatt = email
                Exit Function
            End If
        End If
    Next zeile
End Function

Private Function NormalisiereKundenname(ByVal rawText As String) As String
    Dim t As String
    Dim teile() As String

    t = Replace(rawText, vbCr, vbLf)
    teile = Split(t, vbLf)
    t = Trim$(teile(0))

    t = LCase$(t)
    t = Replace(t, "ae", "a")
    t = Replace(t, "oe", "o")
    t = Replace(t, "ue", "u")
    t = Replace(t, "ss", "s")

    t = Replace(t, "gmbh", "")
    t = Replace(t, "co. kg", "")
    t = Replace(t, "co kg", "")
    t = Replace(t, "kg", "")
    t = Replace(t, "ag", "")

    t = Replace(t, "-", " ")
    t = Replace(t, ",", " ")
    t = Replace(t, ".", " ")

    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop

    NormalisiereKundenname = Trim$(t)
End Function

Function DateiMitVersion(pfad As String, BasisName As String) As String
    Dim version As Long
    Dim vollName As String

    version = 1
    vollName = BasisName & "_v" & version & ".pdf"

    Do While Dir(pfad & vollName) <> ""
        version = version + 1
        vollName = BasisName & "_v" & version & ".pdf"
    Loop

    DateiMitVersion = vollName
End Function

Private Sub txtVersion_Change()

End Sub
'''


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Erweitert den PDF-Export in frmPDFDruck um Outlook-E-Mail-Versand "
            "und pflegt E-Mail-Adressen in Kunden!B anhand Muster_Termine."
        )
    )
    parser.add_argument(
        "--source",
        type=Path,
        default=DEFAULT_SOURCE,
        help="Quelle fuer Kunde->E-Mail (xlsx, Blatt Kundenadressen/Kunden)",
    )
    parser.add_argument(
        "--targets",
        type=Path,
        nargs="+",
        default=DEFAULT_TARGETS_REAL,
        help="Ziel-xlsm-Dateien, die angepasst werden sollen",
    )
    parser.add_argument(
        "--backup-suffix",
        default=".bak_before_pdf_mail",
        help="Suffix fuer Sicherung vor der Endung .xlsm",
    )
    parser.add_argument(
        "--visible",
        action="store_true",
        help="Excel sichtbar starten (Debug)",
    )
    return parser.parse_args()


def ensure_exists(path: Path) -> None:
    if not path.exists():
        raise FileNotFoundError(f"Datei nicht gefunden: {path}")


def _unique_paths(paths: list[Path]) -> list[Path]:
    seen: set[str] = set()
    result: list[Path] = []
    for path in paths:
        key = str(path).lower()
        if key in seen:
            continue
        seen.add(key)
        result.append(path)
    return result


def resolve_source_path(source_path: Path) -> Path:
    if source_path.exists():
        return source_path

    fallback_candidates: list[Path] = [
        Path("samples/topics/excel-basis/Muster_Termine 17.08.2026.xlsx"),
        Path("samples/Muster_Termine 17.08.2026.xlsx"),
        Path("Muster_Termine 17.08.2026.xlsx"),
    ]
    for candidate in fallback_candidates:
        if candidate.exists():
            return candidate

    recursive_hits = sorted(Path(".").glob("**/Muster_Termine*.xlsx"))
    if recursive_hits:
        return recursive_hits[0]

    termine_hits = sorted(Path(".").glob("**/Termine*.xlsx"))
    if termine_hits:
        return termine_hits[0]

    raise FileNotFoundError(
        f"Datei nicht gefunden: {source_path}. "
        "Bitte --source explizit angeben."
    )


def resolve_target_paths(targets: list[Path]) -> list[Path]:
    existing = [target for target in targets if target.exists()]
    if existing:
        return _unique_paths(existing)

    candidates: list[Path] = []
    for candidate in DEFAULT_TARGETS_REAL + DEFAULT_TARGETS_FALLBACK:
        if candidate.exists():
            candidates.append(candidate)

    if candidates:
        return _unique_paths(candidates)

    recursive_hits: list[Path] = []
    for pattern in [
        "**/*V20d*Übergeordneter_Entkupplungsschutz*.xlsm",
        "**/*V20D*Übergeordneter_Entkupplungsschutz*.xlsm",
        "**/*v20d*Übergeordneter_Entkupplungsschutz*.xlsm",
        "**/*V20d*Uebergeordneter_Entkupplungsschutz*.xlsm",
        "**/*V20D*Uebergeordneter_Entkupplungsschutz*.xlsm",
        "**/*v20d*Uebergeordneter_Entkupplungsschutz*.xlsm",
    ]:
        recursive_hits.extend(sorted(Path(".").glob(pattern)))

    filtered = [
        path
        for path in recursive_hits
        if "bak" not in path.name.lower() and not path.name.startswith("~$")
    ]
    if filtered:
        return _unique_paths(filtered)

    raise FileNotFoundError(
        "Keine passende V20d-Ziel-XLSM gefunden. Bitte mit --targets die V20d-Datei explizit angeben."
    )


def resolve_logo_reference_path(target_path: Path) -> Path | None:
    for candidate in LOGO_REFERENCE_CANDIDATES:
        if candidate.exists() and candidate.resolve() != target_path.resolve():
            return candidate

    recursive_hits = sorted(Path(".").glob("**/*V19*Übergeordneter_Entkupplungsschutz*.xlsm"))
    for hit in recursive_hits:
        if hit.resolve() != target_path.resolve():
            return hit

    recursive_hits_ascii = sorted(Path(".").glob("**/*V19*Uebergeordneter_Entkupplungsschutz*.xlsm"))
    for hit in recursive_hits_ascii:
        if hit.resolve() != target_path.resolve():
            return hit

    return None


def _shape_intersects_d7_e7(shape) -> bool:
    try:
        tl_row = int(shape.TopLeftCell.Row)
        tl_col = int(shape.TopLeftCell.Column)
        br_row = int(shape.BottomRightCell.Row)
        br_col = int(shape.BottomRightCell.Column)
    except Exception:
        return False

    return not (br_row < 7 or tl_row > 7 or br_col < 4 or tl_col > 5)


def restore_logos_from_reference(excel_app, target_workbook, reference_path: Path) -> int:
    ref_workbook = None
    restored = 0
    try:
        ref_workbook = excel_app.Workbooks.Open(str(reference_path.resolve()), ReadOnly=True, AddToMru=False)
        ws_src = ref_workbook.Worksheets("Allgemeine Angaben")
        ws_dst = target_workbook.Worksheets("Allgemeine Angaben")

        # Restore possible formula/text content in D7:E7.
        ws_dst.Range("D7:E7").Formula = ws_src.Range("D7:E7").Formula

        # Remove existing shapes in D7:E7 before re-pasting from V19 reference.
        for i in range(int(ws_dst.Shapes.Count), 0, -1):
            shp = ws_dst.Shapes(i)
            if _shape_intersects_d7_e7(shp):
                try:
                    shp.Delete()
                except Exception:
                    pass

        for i in range(1, int(ws_src.Shapes.Count) + 1):
            shp = ws_src.Shapes(i)
            if not _shape_intersects_d7_e7(shp):
                continue

            left = float(shp.Left)
            top = float(shp.Top)
            width = float(shp.Width)
            height = float(shp.Height)

            shp.Copy()
            ws_dst.Paste()
            pasted = ws_dst.Shapes(int(ws_dst.Shapes.Count))
            pasted.Left = left
            pasted.Top = top
            pasted.Width = width
            pasted.Height = height
            restored += 1

        return restored
    finally:
        if ref_workbook is not None:
            ref_workbook.Close(SaveChanges=False)


def backup_file(path: Path, suffix: str) -> Path:
    backup_path = path.with_name(f"{path.stem}{suffix}{path.suffix}")
    shutil.copy2(path, backup_path)
    return backup_path


def norm_name(value: str) -> str:
    text = value.replace("\r", "\n")
    first = text.split("\n", 1)[0].strip().lower()
    repl = {
        "ae": "a",
        "oe": "o",
        "ue": "u",
        "ss": "s",
        "gmbh": "",
        "co. kg": "",
        "co kg": "",
        "kg": "",
        "ag": "",
        "-": " ",
        ",": " ",
        ".": " ",
    }
    for old, new in repl.items():
        first = first.replace(old, new)
    while "  " in first:
        first = first.replace("  ", " ")
    return first.strip()


def is_placeholder_email(value: str) -> bool:
    lower = value.strip().lower()
    if lower in {"", "?", "-"}:
        return True
    if "genauerausw" in lower:
        return True
    if "ueber " in lower:
        return True
    return False


def load_source_emails(source_path: Path) -> list[tuple[str, str, str]]:
    wb = load_workbook(source_path, data_only=True)
    ws = None
    for name in SOURCE_CANDIDATE_SHEETS:
        if name in wb.sheetnames:
            ws = wb[name]
            break
    if ws is None:
        raise RuntimeError(
            "Kein Blatt 'Kundenadressen' oder 'Kunden' in der Quelle gefunden."
        )

    result: list[tuple[str, str, str]] = []
    for row in range(1, ws.max_row + 1):
        customer = ws.cell(row, 1).value
        email = ws.cell(row, 2).value
        if customer is None:
            continue
        customer_text = str(customer).strip()
        if not customer_text or customer_text.lower() == "kunde":
            continue
        email_text = "" if email is None else str(email).strip()
        if is_placeholder_email(email_text):
            continue
        result.append((norm_name(customer_text), customer_text, email_text))
    return result


def pick_email_for_customer(customer_value: str, source_rows: list[tuple[str, str, str]]) -> str:
    target_key = norm_name(customer_value)
    if not target_key:
        return ""

    exact = [email for key, _raw, email in source_rows if key == target_key]
    if exact:
        return "; ".join(dict.fromkeys(exact))

    fuzzy = [
        email
        for key, _raw, email in source_rows
        if key and (target_key in key or key in target_key)
    ]
    if fuzzy:
        return "; ".join(dict.fromkeys(fuzzy))

    return ""


def update_customer_sheet(path: Path, source_rows: list[tuple[str, str, str]]) -> tuple[int, int]:
    wb = load_workbook(path, keep_vba=True)
    if CUSTOMER_SHEET not in wb.sheetnames:
        raise RuntimeError(f"Blatt '{CUSTOMER_SHEET}' nicht gefunden in {path}.")
    ws = wb[CUSTOMER_SHEET]

    updated = 0
    missing = 0
    for row in range(1, ws.max_row + 1):
        customer = ws.cell(row, 1).value
        if customer is None:
            continue
        customer_text = str(customer).strip()
        if not customer_text:
            continue

        email = pick_email_for_customer(customer_text, source_rows)
        if email:
            if ws.cell(row, 2).value != email:
                ws.cell(row, 2).value = email
                updated += 1
        else:
            ws.cell(row, 2).value = ""
            missing += 1

    wb.save(path)
    wb.close()
    return updated, missing


def update_customer_sheet_excel_com(workbook, source_rows: list[tuple[str, str, str]]) -> tuple[int, int]:
    try:
        ws = workbook.Worksheets(CUSTOMER_SHEET)
    except Exception as error:
        raise RuntimeError(f"Blatt '{CUSTOMER_SHEET}' nicht gefunden in {workbook.Name}.") from error

    last_row = int(ws.Cells(ws.Rows.Count, 1).End(-4162).Row)  # xlUp
    updated = 0
    missing = 0

    for row in range(1, last_row + 1):
        customer_value = ws.Cells(row, 1).Value
        if customer_value is None:
            continue
        customer_text = str(customer_value).strip()
        if not customer_text:
            continue

        email = pick_email_for_customer(customer_text, source_rows)
        current_email = ws.Cells(row, 2).Value
        current_text = "" if current_email is None else str(current_email)

        if email:
            if current_text != email:
                ws.Cells(row, 2).Value = email
                updated += 1
        else:
            if current_text != "":
                ws.Cells(row, 2).Value = ""
            missing += 1

    return updated, missing


def _extract_com_hresult(error: Exception) -> int | None:
    if not getattr(error, "args", None):
        return None
    first_arg = error.args[0]
    if isinstance(first_arg, int):
        return first_arg
    return None


def _is_excel_call_rejected(error: Exception) -> bool:
    # RPC_E_CALL_REJECTED
    return _extract_com_hresult(error) == -2147418111


def _retry_excel_call(callable_fn, attempts: int = 20, wait_seconds: float = 0.35):
    for attempt in range(1, attempts + 1):
        try:
            return callable_fn()
        except Exception as error:
            if (not _is_excel_call_rejected(error)) or attempt == attempts:
                raise
            time.sleep(wait_seconds)


def apply_abschlussbemerkungen_nicht_ok_logic(workbook) -> None:
    ws_angaben = workbook.Worksheets("Allgemeine Angaben")
    ws_protokoll = workbook.Worksheets("Prüfprotokoll")

    # Map "!" from checklist-linked fields to "NICHT OK" so existing warning semantics apply.
    formula_updates = {
        "C105": "=IF('Schutzprüf-Checkliste'!J70=\"x\",\"x\",IF('Schutzprüf-Checkliste'!J70=\"!\",\"NICHT OK\",\"\"))",
        "C106": "=IF('Schutzprüf-Checkliste'!L70=\"x\",\"x\",IF('Schutzprüf-Checkliste'!L70=\"!\",\"NICHT OK\",\"\"))",
        "C107": "=IF('Schutzprüf-Checkliste'!J71=\"x\",\"x\",IF('Schutzprüf-Checkliste'!J71=\"!\",\"NICHT OK\",\"\"))",
        "C108": "=IF('Schutzprüf-Checkliste'!L71=\"x\",\"x\",IF('Schutzprüf-Checkliste'!L71=\"!\",\"NICHT OK\",\"\"))",
        "C114": "=IF('Schutzprüf-Checkliste'!J72=\"x\",\"x\",IF('Schutzprüf-Checkliste'!J72=\"!\",\"NICHT OK\",\"\"))",
        "C115": "=IF('Schutzprüf-Checkliste'!L72=\"x\",\"x\",IF('Schutzprüf-Checkliste'!L72=\"!\",\"NICHT OK\",\"\"))",
        "C109": "=IF('Schutzprüf-Checkliste'!E74=\"x\",\"x\",IF('Schutzprüf-Checkliste'!E74=\"!\",\"NICHT OK\",\"\"))",
    }
    for cell, formula in formula_updates.items():
        ws_angaben.Range(cell).Formula = formula

    # Ensure description rows in Abschlussbemerkungen are also shown for "!".
    protokoll_formula_updates = {
        "A169": "=IF(OR('Schutzprüf-Checkliste'!J70=TRUE,'Schutzprüf-Checkliste'!J70=\"x\",'Schutzprüf-Checkliste'!J70=\"!\"),\"Abschaltung MS-LS\"&'Schutzprüf-Checkliste'!C67 &\" nach Ausfall der Hilfsspannung - AuxDC\",\"\")",
        "A170": "=IF(OR('Schutzprüf-Checkliste'!L70=TRUE,'Schutzprüf-Checkliste'!L70=\"x\",'Schutzprüf-Checkliste'!L70=\"!\"),\"Abschaltung NS-LS\"&'Schutzprüf-Checkliste'!C68 &\" nach Ausfall der Hilfsspannung- AuxDC\",\"\")",
        "A171": "=IF(OR('Schutzprüf-Checkliste'!J71=TRUE,'Schutzprüf-Checkliste'!J71=\"x\",'Schutzprüf-Checkliste'!J71=\"!\"),\"Abschaltung MS-LS\"&'Schutzprüf-Checkliste'!C67 &\" nach Ausfall Schutzrelais (Live Contact)\",\"\")",
        "A172": "=IF(OR('Schutzprüf-Checkliste'!L71=TRUE,'Schutzprüf-Checkliste'!L71=\"x\",'Schutzprüf-Checkliste'!L71=\"!\"),\"Abschaltung NS-LS\"&'Schutzprüf-Checkliste'!C68 &\" nach Ausfall Schutzrelais (Live Contact)\",\"\")",
        "A173": "=IF(OR('Schutzprüf-Checkliste'!J72=TRUE,'Schutzprüf-Checkliste'!J72=\"x\",'Schutzprüf-Checkliste'!J72=\"!\"),'Allgemeine Angaben'!A114,\"\")",
        "A174": "=IF(OR('Schutzprüf-Checkliste'!L72=TRUE,'Schutzprüf-Checkliste'!L72=\"x\",'Schutzprüf-Checkliste'!L72=\"!\"),'Allgemeine Angaben'!A115,\"\")",
        "A182": "=IF(OR('Schutzprüf-Checkliste'!E74=\"x\",'Schutzprüf-Checkliste'!E74=\"!\"),\"DC USV für übergeordneter Schutz/ggfs. UMZ-Schutz in Ordnung\",\"\")",
    }
    for cell, formula in protokoll_formula_updates.items():
        ws_protokoll.Range(cell).Formula = formula

    # Existing red rule covers J168:J175. Add same semantic rule for J182 (linked to row 74).
    target_cell = ws_protokoll.Range("J182")
    has_nicht_ok_rule = False
    for i in range(1, int(target_cell.FormatConditions.Count) + 1):
        rule = target_cell.FormatConditions(i)
        try:
            if int(rule.Type) == 1 and int(rule.Operator) == 3:
                formula = str(rule.Formula1).replace("=", "").replace('"', "").strip().upper()
                if formula == "NICHT OK":
                    has_nicht_ok_rule = True
                    break
        except Exception:
            continue

    if not has_nicht_ok_rule:
        added_rule = target_cell.FormatConditions.Add(Type=1, Operator=3, Formula1='="NICHT OK"')
        try:
            ref_rule = ws_protokoll.Range("J168").FormatConditions(1)
            added_rule.Font.Color = ref_rule.Font.Color
            added_rule.Interior.Color = ref_rule.Interior.Color
            added_rule.StopIfTrue = ref_rule.StopIfTrue
        except Exception:
            # Style copy is best effort; semantic condition is the functional requirement.
            pass

    # Show station lock message in row 3 when any lower status is NICHT OK.
    lock_message = "Station gesperrt - Nicht alle Schutzfunktionen in Ordnung"
    lock_formula = (
        "=IF(OR(COUNTIF($J$167:$J$175,\"NICHT OK\")>0,$J$182=\"NICHT OK\"),"
        f"\"{lock_message}\""
        ",\"Prüfer\")"
    )
    ws_protokoll.Range("A3").Formula = lock_formula

    lock_cf_formula = '=OR(COUNTIF($J$167:$J$175,"NICHT OK")>0,$J$182="NICHT OK")'
    a3_cell = ws_protokoll.Range("A3")
    has_lock_cf_rule = False
    for i in range(1, int(a3_cell.FormatConditions.Count) + 1):
        rule = a3_cell.FormatConditions(i)
        try:
            if int(rule.Type) == 2:
                if str(rule.Formula1).replace(" ", "") == lock_cf_formula.replace(" ", ""):
                    has_lock_cf_rule = True
                    break
        except Exception:
            continue

    if not has_lock_cf_rule:
        lock_rule = a3_cell.FormatConditions.Add(Type=2, Formula1=lock_cf_formula)
        try:
            ref_rule = ws_protokoll.Range("J168").FormatConditions(1)
            lock_rule.Font.Color = ref_rule.Font.Color
            lock_rule.Interior.Color = ref_rule.Interior.Color
            lock_rule.StopIfTrue = ref_rule.StopIfTrue
        except Exception:
            # Style copy is best effort; enforce red background if no reference style is available.
            lock_rule.Interior.Color = 255
            lock_rule.Font.Color = 16777215


def patch_pdf_form_vba(path: Path, visible: bool) -> None:
    try:
        import pythoncom
        import win32com.client
    except ImportError as error:
        raise RuntimeError(
            "Excel-COM ist nicht verfuegbar. Bitte unter Windows mit pywin32 ausfuehren."
        ) from error

    pythoncom.CoInitialize()
    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = bool(visible)
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False

        workbook = excel.Workbooks.Open(str(path.resolve()))
        component = workbook.VBProject.VBComponents("frmPDFDruck")
        module = component.CodeModule

        if module.CountOfLines > 0:
            module.DeleteLines(1, module.CountOfLines)
        module.AddFromString(FORM_CODE)

        workbook.Save()
        workbook.Close(SaveChanges=True)
        workbook = None
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        if excel is not None:
            excel.Quit()
        pythoncom.CoUninitialize()


def apply_changes_with_excel_com(
    path: Path,
    source_rows: list[tuple[str, str, str]],
    visible: bool,
) -> tuple[int, int]:
    try:
        import pythoncom
        import win32com.client
        import win32api
    except ImportError as error:
        raise RuntimeError(
            "Excel-COM ist nicht verfuegbar. Bitte unter Windows mit pywin32 ausfuehren."
        ) from error

    pythoncom.CoInitialize()
    excel = None
    workbook = None
    temp_dir: str | None = None
    temp_workbook_path: Path | None = None
    original_path = path.resolve()

    def _open_workbook(excel_app, workbook_path_str: str):
        last_error: Exception | None = None
        # Try normal open first.
        open_variants = [
            {},
            {
                "UpdateLinks": 0,
                "ReadOnly": False,
                "IgnoreReadOnlyRecommended": True,
                "AddToMru": False,
                "Local": True,
            },
            {
                "UpdateLinks": 0,
                "ReadOnly": False,
                "IgnoreReadOnlyRecommended": True,
                "AddToMru": False,
                "Local": True,
                # 1 = xlRepairFile
                "CorruptLoad": 1,
            },
        ]

        for kwargs in open_variants:
            try:
                if kwargs:
                    return _retry_excel_call(lambda kwargs=kwargs: excel_app.Workbooks.Open(workbook_path_str, **kwargs))
                return _retry_excel_call(lambda: excel_app.Workbooks.Open(workbook_path_str))
            except Exception as error:
                last_error = error
                continue

        if last_error is not None:
            raise last_error
        raise RuntimeError("Unbekannter Fehler beim Oeffnen der Excel-Datei.")

    try:
        excel = _retry_excel_call(lambda: win32com.client.DispatchEx("Excel.Application"))
        excel.Visible = bool(visible)
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False

        workbook_path = str(original_path)
        try:
            short_path = win32api.GetShortPathName(workbook_path)
            if short_path:
                workbook_path = short_path
        except Exception:
            pass

        try:
            workbook = _open_workbook(excel, workbook_path)
        except Exception:
            # Fallback: copy to temp path with ASCII-only filename and retry there.
            temp_dir = tempfile.mkdtemp(prefix="ges_excel_open_")
            temp_workbook_path = Path(temp_dir) / "workbook.xlsm"
            shutil.copy2(original_path, temp_workbook_path)
            workbook = _open_workbook(excel, str(temp_workbook_path))

        updated, missing = _retry_excel_call(lambda: update_customer_sheet_excel_com(workbook, source_rows))
        _retry_excel_call(lambda: apply_abschlussbemerkungen_nicht_ok_logic(workbook))

        reference_logo_path = resolve_logo_reference_path(path)
        restored_logos = 0
        if reference_logo_path is not None:
            restored_logos = _retry_excel_call(
                lambda: restore_logos_from_reference(excel, workbook, reference_logo_path)
            )
            print(f"{path}: Logos aus V19 wiederhergestellt (Anzahl: {restored_logos})")
        else:
            print(f"{path}: Keine V19-Referenzdatei fuer Logo-Wiederherstellung gefunden")

        component = _retry_excel_call(lambda: workbook.VBProject.VBComponents("frmPDFDruck"))
        module = component.CodeModule
        if module.CountOfLines > 0:
            _retry_excel_call(lambda: module.DeleteLines(1, module.CountOfLines))
        _retry_excel_call(lambda: module.AddFromString(FORM_CODE))

        _retry_excel_call(lambda: workbook.Save())
        _retry_excel_call(lambda: workbook.Close(SaveChanges=True))
        workbook = None

        if temp_workbook_path is not None:
            shutil.copy2(temp_workbook_path, original_path)

        return updated, missing
    finally:
        if workbook is not None:
            _retry_excel_call(lambda: workbook.Close(SaveChanges=False))
        if excel is not None:
            _retry_excel_call(lambda: excel.Quit())
        pythoncom.CoUninitialize()
        if temp_dir is not None:
            shutil.rmtree(temp_dir, ignore_errors=True)


def main() -> int:
    args = parse_args()
    source_path = resolve_source_path(args.source)

    requested_targets = [Path(p) for p in args.targets]
    targets = resolve_target_paths(requested_targets)

    source_rows = load_source_emails(source_path)
    if not source_rows:
        raise RuntimeError("Keine nutzbaren Kunde->E-Mail-Daten in der Quelle gefunden.")

    print(f"Quelle geladen: {source_path}")
    print(f"Nutzbare E-Mail-Zuordnungen: {len(source_rows)}")
    print("Ziel-Dateien:")
    for target in targets:
        print(f"- {target}")

    failures: list[str] = []

    for target in targets:
        try:
            backup_path = backup_file(target, args.backup_suffix)
            print(f"Sicherung erstellt: {backup_path}")

            updated, missing = apply_changes_with_excel_com(target, source_rows, visible=args.visible)
            print(f"{target}: Kundenblatt aktualisiert (gesetzt: {updated}, ohne Treffer: {missing})")
            print(f"{target}: VBA frmPDFDruck aktualisiert")
        except Exception as error:
            failures.append(f"{target}: {error}")
            print(f"Warnung: Verarbeitung fehlgeschlagen fuer {target}: {error}")

    if failures:
        print("\nFolgende Dateien konnten nicht verarbeitet werden:", file=sys.stderr)
        for item in failures:
            print(f"- {item}", file=sys.stderr)
        return 1

    print("Fertig: PDF->E-Mail-Option und Kunden-E-Mail-Spalte sind eingepflegt.")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"Fehler: {exc}", file=sys.stderr)
        raise SystemExit(1)
