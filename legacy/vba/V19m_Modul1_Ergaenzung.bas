Attribute VB_Name = "Modul1"
Option Explicit

Public Sub BereicheEinOderAusblenden_Start()
    Dim wsAngaben As Worksheet
    Dim wsProtokoll As Worksheet
    Dim startZeile As Long
    Dim H14 As Long
    Dim H15 As Long
    Dim H16 As Long
    Dim H17 As Long
    Dim H18 As Long
    Dim H19 As Long
    Dim H20 As Long
    Dim H21 As Long
    Dim H22 As Long
    Dim H23 As Long
    Dim H24 As Long
    Dim H25 As Long
    Dim zeile As Long

    Set wsAngaben = ThisWorkbook.Worksheets("Allgemeine Angaben")
    Set wsProtokoll = ThisWorkbook.Worksheets("Prüfprotokoll")

    startZeile = 1
    H14 = startZeile + wsAngaben.Range("H14").Value
    H15 = H14 + wsAngaben.Range("H15").Value
    H20 = H15 + wsAngaben.Range("H20").Value
    H21 = H20 + wsAngaben.Range("H21").Value
    H16 = H21 + wsAngaben.Range("H16").Value
    H17 = H16 + wsAngaben.Range("H17").Value
    H18 = H17 + wsAngaben.Range("H18").Value
    H19 = H18 + wsAngaben.Range("H19").Value
    H23 = H19 + wsAngaben.Range("H23").Value
    H24 = H23 + wsAngaben.Range("H24").Value
    H22 = H24 + wsAngaben.Range("H22").Value
    H25 = H22 + wsAngaben.Range("H25").Value

    For zeile = H14 To H15
        wsProtokoll.Rows(zeile).Hidden = (wsAngaben.Range("C15").Value = "")
    Next zeile

    For zeile = H15 To H20
        wsProtokoll.Rows(zeile).Hidden = True
    Next zeile

    For zeile = H20 To H21
        wsProtokoll.Rows(zeile).Hidden = (wsAngaben.Range("C21").Value = "")
    Next zeile

    For zeile = H21 To H16
        wsProtokoll.Rows(zeile).Hidden = (wsAngaben.Range("C16").Value = "")
    Next zeile

    For zeile = H16 To H17
        wsProtokoll.Rows(zeile).Hidden = (wsAngaben.Range("C17").Value = "")
    Next zeile

    For zeile = H17 To H18
        wsProtokoll.Rows(zeile).Hidden = (wsAngaben.Range("C18").Value = "")
    Next zeile

    For zeile = H18 To H19
        wsProtokoll.Rows(zeile).Hidden = (wsAngaben.Range("C19").Value = "")
    Next zeile

    For zeile = H19 To H23
        wsProtokoll.Rows(zeile).Hidden = (wsAngaben.Range("C23").Value = "")
    Next zeile

    For zeile = H23 To H24
        wsProtokoll.Rows(zeile).Hidden = (wsAngaben.Range("C24").Value = "")
    Next zeile

    For zeile = H24 To H22
        wsProtokoll.Rows(zeile).Hidden = (wsAngaben.Range("C22").Value = "")
    Next zeile

    For zeile = H22 To H25
        wsProtokoll.Rows(zeile).Hidden = (wsAngaben.Range("C25").Value = "")
    Next zeile

    If ThisWorkbook.Worksheets("Schutzprüf-Checkliste").Range("W14").Value = False Then
        wsProtokoll.Rows("159:164").EntireRow.Hidden = True
    Else
        wsProtokoll.Rows("159:164").EntireRow.Hidden = False
    End If
End Sub