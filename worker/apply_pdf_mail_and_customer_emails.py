from __future__ import annotations

import argparse
import shutil
import sys
from pathlib import Path

from openpyxl import load_workbook

DEFAULT_TARGETS_REAL = [
    Path("samples/V19m_Übergeordneter_Entkupplungsschutz.xlsm"),
    Path("samples/topics/excel-basis/V20a_Übergeordneter_Entkupplungsschutz.xlsm"),
]

DEFAULT_SOURCE = Path("samples/topics/excel-basis/Muster_Termine 17.08.2026.xlsx")
CUSTOMER_SHEET = "Kunden"
SOURCE_CANDIDATE_SHEETS = ("Kundenadressen", "Kunden")

FORM_CODE = r'''Attribute VB_Name = "frmPDFDruck"
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

    If chkBlatt1.Value = True Then
        dateiname = wb.Sheets("Allgemeine Angaben").Range("C5").Value & "_" & _
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
        dateiname = wb.Sheets("Allgemeine Angaben").Range("C5").Value & "_" & _
                    wb.Sheets("Prüfprotokoll").Range("H13").Value & "_Wandlerprüfprotokoll"

        vollDateiname = DateiMitVersion(pfad, dateiname)
        vollPfad = pfad & vollDateiname

        wb.Sheets("Wandlerprüfprotokoll").ExportAsFixedFormat Type:=xlTypePDF, _
            Filename:=vollPfad, Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False

        exportiertePdfs.Add vollPfad
    End If

    If chkBlatt3.Value = True Then
        dateiname = wb.Sheets("Allgemeine Angaben").Range("C6").Value & "_" & _
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

    betreff = "Pruefprotokoll " & CStr(wsAngaben.Range("C5").Value)
    bodyText = ErzeugeStandardMailtext(kundeText, bemerkungen)
    logoPath = FindeLogoPfad(wb.Path)
    htmlBody = ErzeugeHtmlMailtext(kundeText, bemerkungen, logoPath)

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

Private Function ErzeugeStandardMailtext(ByVal kundeText As String, ByVal bemerkungen As String) As String
    Dim text As String

    text = "Sehr geehrte Damen und Herren," & vbCrLf & vbCrLf & _
        "bitte entnehmen Sie dem Anhang die Rechnungen." & vbCrLf & vbCrLf

    If Trim$(bemerkungen) <> "" Then
        text = text & "Bemerkungen aus der Checkliste:" & vbCrLf & bemerkungen & vbCrLf & vbCrLf
    End If

    text = text & "Herzlichen Dank." & vbCrLf & vbCrLf & _
        "Mit freundlichen Gruessen" & vbCrLf & vbCrLf & _
        "Gunnar Schaefer" & vbCrLf & vbCrLf & _
        "Dipl.-Ing. Elektrotechnik" & vbCrLf & vbCrLf & _
        "G.E.S. Energietechnik GmbH" & vbCrLf & vbCrLf & _
        "Ferchlipp 16" & vbCrLf & _
        "39615 Altmaerkische Wische" & vbCrLf & vbCrLf & _
        "E-Mail g.schaefer@ges-energietechnik.com" & vbCrLf & _
        "Internet www.ges-energietechnik.com" & vbCrLf & vbCrLf & _
        "Der Inhalt dieser Email ist ausschliesslich fuer den bezeichneten Adressaten bestimmt." & vbCrLf & _
        "Falls Sie nicht der vorgesehene Adressat dieser Email oder dessen Vertreter sein sollten," & vbCrLf & _
        "so beachten Sie bitte, dass jede Form der Kenntnisnahme, Veroeffentlichung," & vbCrLf & _
        "Vervielfaeltigung oder Weitergabe des Inhalts dieser Email unzulaessig ist." & vbCrLf & _
        "Wir bitten Sie, sich in diesem Fall mit dem Absender der Email in Verbindung zu setzen." & vbCrLf & vbCrLf & _
        "G.E.S. Energietechnik GmbH: Sitz der Gesellschaft: Altmaerkische Wische - Amtsgericht Stendal," & vbCrLf & _
        "HRB 30020 - Geschaeftsfuehrer: Gunnar Schaefer"

    ErzeugeStandardMailtext = text
End Function

Private Function ErzeugeHtmlMailtext(ByVal kundeText As String, ByVal bemerkungen As String, ByVal logoPath As String) As String
    Dim text As String
    Dim bemerkHtml As String

    bemerkHtml = Replace(HTMLEncode(bemerkungen), vbCrLf, "<br>")

    text = "<html><body style='font-family:Calibri,Arial,sans-serif;font-size:11pt;'>" & _
           "<p>Sehr geehrte Damen und Herren,</p>" & _
           "<p>bitte entnehmen Sie dem Anhang die Rechnungen.</p>"

    If Trim$(bemerkungen) <> "" Then
        text = text & "<p><b>Bemerkungen aus der Checkliste:</b><br>" & bemerkHtml & "</p>"
    End If

    text = text & "<p>Herzlichen Dank.</p>" & _
                  "<p>Mit freundlichen Gruessen</p>" & _
                  "<p>Gunnar Schaefer</p>" & _
                  "<p>Dipl.-Ing. Elektrotechnik</p>" & _
                  "<p>G.E.S. Energietechnik GmbH</p>" & _
                  "<p>Ferchlipp 16<br>39615 Altmaerkische Wische</p>" & _
                  "<p>E-Mail <a href='mailto:g.schaefer@ges-energietechnik.com'>g.schaefer@ges-energietechnik.com</a><br>" & _
                  "Internet <a href='http://www.ges-energietechnik.com'>www.ges-energietechnik.com</a></p>" & _
                  "<p style='font-size:9pt;color:#555;'>" & _
                  "Der Inhalt dieser Email ist ausschliesslich fuer den bezeichneten Adressaten bestimmt.<br>" & _
                  "Falls Sie nicht der vorgesehene Adressat dieser Email oder dessen Vertreter sein sollten,<br>" & _
                  "so beachten Sie bitte, dass jede Form der Kenntnisnahme, Veroeffentlichung,<br>" & _
                  "Vervielfaeltigung oder Weitergabe des Inhalts dieser Email unzulaessig ist.<br>" & _
                  "Wir bitten Sie, sich in diesem Fall mit dem Absender der Email in Verbindung zu setzen.<br><br>" & _
                  "G.E.S. Energietechnik GmbH: Sitz der Gesellschaft: Altmaerkische Wische - Amtsgericht Stendal,<br>" & _
                  "HRB 30020 - Geschaeftsfuehrer: Gunnar Schaefer" & _
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


def main() -> int:
    args = parse_args()
    ensure_exists(args.source)

    targets = [Path(p) for p in args.targets]
    for target in targets:
        ensure_exists(target)

    source_rows = load_source_emails(args.source)
    if not source_rows:
        raise RuntimeError("Keine nutzbaren Kunde->E-Mail-Daten in der Quelle gefunden.")

    print(f"Quelle geladen: {args.source}")
    print(f"Nutzbare E-Mail-Zuordnungen: {len(source_rows)}")

    for target in targets:
        backup_path = backup_file(target, args.backup_suffix)
        print(f"Sicherung erstellt: {backup_path}")

        updated, missing = update_customer_sheet(target, source_rows)
        print(f"{target}: Kundenblatt aktualisiert (gesetzt: {updated}, ohne Treffer: {missing})")

        patch_pdf_form_vba(target, visible=args.visible)
        print(f"{target}: VBA frmPDFDruck aktualisiert")

    print("Fertig: PDF->E-Mail-Option und Kunden-E-Mail-Spalte sind eingepflegt.")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"Fehler: {exc}", file=sys.stderr)
        raise SystemExit(1)
