Attribute VB_Name = "mod_ReceiptTag"
Option Explicit

' ============================================================
' ERP LSM [CITY]  Smart Receipt Tag Generator
' Compatible: Excel 2010 / Windows 7
' Author: LSM VBA Core  TAG1801 GSL  Public Sector 2026
' ============================================================
' ARCHITECTURE: Offline-First, Zero External Dependencies
' QR code replaced with local verification code (100% VBA).
' Public sector compliant - no internet connection required.
' ============================================================
' IMPORTANT: QR code generation requires internet connection.
' If offline, the macro will skip the QR image and still
' produce a clean printable PDF receipt tag.
' ============================================================

Public Sub GenerateReceiptTagPDF()

    Dim wsData     As Worksheet
    Dim wsTemplate As Worksheet
    Dim qrData     As String
    Dim pdfPath    As String
    Dim receiptID  As String
    Dim activeRow  As Long
    Dim pic        As Picture

    ' 1. Define sheets
    Set wsData = ThisWorkbook.Sheets(mod_Config.SHEET_MOUVEMENTS)
    Set wsTemplate = ThisWorkbook.Sheets("RECEIPT_TAG")

    ' 2. Get data from the last active movement row
    activeRow = ActiveCell.Row
    If activeRow < 3 Then
        MsgBox "S" & Chr(233) & "lectionnez une ligne de mouvement (ligne 3 ou plus) avant de lancer la macro.", vbExclamation
        Exit Sub
    End If

    ' 3. Read movement data
    Dim mvtDate    As String: mvtDate = Format(wsData.Cells(activeRow, 1).Value, "DD/MM/YYYY")
    Dim artCode    As String: artCode = wsData.Cells(activeRow, 2).Value
    Dim artDesc    As String: artDesc = wsData.Cells(activeRow, 3).Value      ' auto from VLOOKUP
    Dim mvtType    As String: mvtType = wsData.Cells(activeRow, 4).Value
    Dim qty        As String: qty = wsData.Cells(activeRow, 5).Value
    Dim prixUnit   As String: prixUnit = wsData.Cells(activeRow, 6).Value
    Dim valeur     As String: valeur = wsData.Cells(activeRow, 7).Value
    Dim refDoc     As String: refDoc = wsData.Cells(activeRow, 8).Value

    receiptID = artCode & "-" & Format(Now, "YYYYMMDD-HHMM")

    ' 4. Build tracking string for QR code (Excel 2010 Compatible)
    qrData = "ID:" & receiptID & _
             "|DATE:" & mvtDate & _
             "|ART:" & artCode & _
             "|QTY:" & qty & _
             "|TYPE:" & mvtType & _
             "|REF:" & refDoc

    ' Basic URL Encoding for Excel 2010 (Replace spaces and pipes)
    qrData = Replace(qrData, " ", "%20")
    qrData = Replace(qrData, "|", "%7C")

    ' 5. Populate the RECEIPT_TAG template sheet
    With wsTemplate
        .Range("TAG_ID").Value = receiptID
        .Range("TAG_DATE").Value = mvtDate
        .Range("TAG_CODE").Value = artCode
        .Range("TAG_DESC").Value = artDesc
        .Range("TAG_TYPE").Value = mvtType
        .Range("TAG_QTY").Value = qty
        .Range("TAG_PRIX").Value = prixUnit & " DA"
        .Range("TAG_VALEUR").Value = valeur & " DA"
        .Range("TAG_REF").Value = refDoc
    End With

    ' 6. Generate local verification code (Offline-First, no internet)
    Dim verifyCode As String
    verifyCode = mod_Utilities.GenerateVerifyCode(receiptID & mvtDate & artCode & qty & mvtType)

    ' Place verification code on the tag
    With wsTemplate
        .Range("E4:F9").Merge
        .Range("E4").Value = verifyCode
        .Range("E4").Font.Size = 9
        .Range("E4").Font.Bold = True
        .Range("E4").Font.name = "Courier New"
        .Range("E4").HorizontalAlignment = xlCenter
        .Range("E4").VerticalAlignment = xlCenter
        .Range("E4").Interior.Color = RGB(240, 248, 255)
        .Range("E4").BorderAround Color:=RGB(0, 0, 0), Weight:=xlThin
    End With

    ' 7. Export RECEIPT_TAG sheet as PDF
    pdfPath = ThisWorkbook.Path & "\receipt_tags\ReceiptTag_" & receiptID & ".pdf"

    ' Create receipt_tags folder if needed
    If Dir(ThisWorkbook.Path & "\receipt_tags", vbDirectory) = "" Then
        MkDir ThisWorkbook.Path & "\receipt_tags"
    End If

    wsTemplate.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        fileName:=pdfPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=False, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True

    ' 8. Confirmation message
    MsgBox "Bon de rception gn" & Chr(233) & "r" & Chr(233) & " avec succ" & Chr(232) & "s!" & vbCrLf & _
           "Fichier: " & pdfPath & vbCrLf & _
           "Code de v" & Chr(233) & "rification: " & verifyCode, vbInformation, "ERP Acad" & Chr(233) & "mie v13"

End Sub

' ============================================================
' SETUP: Run this once to create named ranges on RECEIPT_TAG sheet
' ============================================================
Public Sub SetupReceiptTagSheet()

    Dim ws As Worksheet
    Dim sheetExists As Boolean: sheetExists = False
    Dim s As Worksheet
    Dim fields As Variant
    Dim i As Integer
    Dim valueCell As Range

    ' Create RECEIPT_TAG sheet if it doesn't exist
    For Each s In ThisWorkbook.Sheets
        If s.name = "RECEIPT_TAG" Then sheetExists = True
    Next s

    If Not sheetExists Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        ws.name = "RECEIPT_TAG"
    Else
        Set ws = ThisWorkbook.Sheets("RECEIPT_TAG")
    End If

    ' Build the bilingual receipt tag layout
    With ws
        .Cells.Clear
        .PageSetup.PaperSize = xlPaperA5        ' A5  half page, print 2 per A4
        .PageSetup.Orientation = xlPortrait
        .PageSetup.LeftMargin = Application.InchesToPoints(0.4)
        .PageSetup.RightMargin = Application.InchesToPoints(0.4)
        .PageSetup.TopMargin = Application.InchesToPoints(0.4)
        .PageSetup.BottomMargin = Application.InchesToPoints(0.4)

        ' Header
        .Range("A1:F1").Merge
        .Range("A1").Value = "MAGASIN MDIDRIYA TARBIYA [CITY]  /  ???? ???????? ?????"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 11
        .Range("A1").HorizontalAlignment = xlCenter

        ' Subheader
        .Range("A2:F2").Merge
        .Range("A2").Value = "BON DE MOUVEMENT / ????? ???? ?????"
        .Range("A2").Font.Size = 9
        .Range("A2").HorizontalAlignment = xlCenter

        ' Field labels (FR left | AR right | Value center)
        fields = Array( _
            Array("Identifiant / ????? ???????", "TAG_ID", "A4", "D4"), _
            Array("Date / ???????", "TAG_DATE", "A5", "D5"), _
            Array("Code Article / ??? ??????", "TAG_CODE", "A6", "D6"), _
            Array("D" & Chr(233) & "signation / ???????", "TAG_DESC", "A7", "D7"), _
            Array("Mouvement / ??? ??????", "TAG_TYPE", "A8", "D8"), _
            Array("Quantit / ??????", "TAG_QTY", "A9", "D9"), _
            Array("Prix Unit. / ??? ??????", "TAG_PRIX", "A10", "D10"), _
            Array("Valeur / ?????? ?????????", "TAG_VALEUR", "A11", "D11"), _
            Array("Rf. Document / ??? ???????", "TAG_REF", "A12", "D12") _
        )

        For i = 0 To UBound(fields)
            .Range(fields(i)(2)).Value = fields(i)(0)
            .Range(fields(i)(2)).Font.Size = 9
            .Range(fields(i)(2)).Font.Bold = True

            ' Create named range for value cell
            Set valueCell = .Range(fields(i)(3))
            ThisWorkbook.Names.Add name:=fields(i)(1), RefersTo:=valueCell
            valueCell.Font.Size = 10
            valueCell.Interior.Color = RGB(255, 252, 196) ' light yellow input
        Next i

        ' QR code placeholder
        .Range("E4:F9").Merge
        .Range("E4").Value = "[QR CODE]"
        .Range("E4").HorizontalAlignment = xlCenter
        .Range("E4").VerticalAlignment = xlCenter
        .Range("E4").Font.Color = RGB(180, 180, 180)
        .Range("E4").Font.Size = 8
        ThisWorkbook.Names.Add name:="TAG_QR", RefersTo:=.Range("E4")

        ' Separator line
        .Range("A13:F13").Merge
        .Range("A13").Value = "????????? ??????? ???  Usage interne exclusivement"
        .Range("A13").Font.Size = 8
        .Range("A13").Font.Italic = True
        .Range("A13").HorizontalAlignment = xlCenter
        .Range("A13").Font.Color = RGB(100, 100, 100)
    End With

    MsgBox "Feuille RECEIPT_TAG cr" & Chr(233) & "e avec succ" & Chr(232) & "s! Lancez la macro apr" & Chr(232) & "s avoir s" & Chr(233) & "lectionn" & Chr(233) & " une ligne de mouvement.", _
           vbInformation, "Setup OK"
End Sub

' ============================================================
' END -- mod_ReceiptTag.bas
' ============================================================
Function GenerateLocalVerifyCode( _
    ByVal receiptID As String, _
    ByVal mvtDate As String, _
    ByVal artCode As String, _
    ByVal qty As String, _
    ByVal mvtType As String) As String

    Dim rawString As String
    Dim checksum As Long
    Dim i As Integer
    Dim chCode As Integer
    Dim hexPart1 As String, hexPart2 As String, hexPart3 As String

    ' Build raw data string
    rawString = receiptID & mvtDate & artCode & qty & mvtType

    ' Generate checksum from character codes
    checksum = 0
    For i = 1 To Len(rawString)
        chCode = Asc(Mid(rawString, i, 1))
        checksum = checksum + (chCode * (i + 7)) Mod 9973
    Next i

    ' Split into 3 hex segments (4 chars each)
    hexPart1 = Right("0000" & Hex(checksum And &HFFFF&), 4)
    hexPart2 = Right("0000" & Hex((checksum \ 17) And &HFFFF&), 4)
    hexPart3 = Right("0000" & Hex((checksum \ 257) And &HFFFF&), 4)

    ' Return formatted verification code
    GenerateLocalVerifyCode = "V-" & hexPart1 & "-" & hexPart2 & "-" & hexPart3

End Function

