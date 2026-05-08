Attribute VB_Name = "mod_Barcode"
Option Explicit

Private Const BARCODE_SHEET As String = "STAGING_BUFFER"
Private Const BARCODE_RANGE As String = "BARCODE_MAP"

Public Function LookupBarcode(ByVal barcode As String) As String
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim code As String

    barcode = Trim(UCase(barcode))
    If Len(barcode) = 0 Then
        LookupBarcode = ""
        Exit Function
    End If

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(BARCODE_SHEET)
    On Error GoTo 0

    If Not ws Is Nothing Then
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        For i = 1 To lastRow
            If Trim(UCase(CStr(ws.Cells(i, 1).Value))) = BARCODE_RANGE Then
                Dim j As Long
                j = i + 1
                Do While j <= lastRow
                    Dim mapBarcode As String
                    Dim mapArticle As String
                    mapBarcode = Trim(UCase(CStr(ws.Cells(j, 1).Value)))
                    mapArticle = Trim(UCase(CStr(ws.Cells(j, 2).Value)))
                    If Len(mapBarcode) = 0 Then Exit Do
                    If mapBarcode = barcode Then
                        LookupBarcode = mapArticle
                        Exit Function
                    End If
                    j = j + 1
                Loop
                Exit For
            End If
        Next i
    End If

    If barcode Like "ART-###" Then
        LookupBarcode = barcode
        Exit Function
    End If

    Dim defaultBarcode As String
    defaultBarcode = GetDefaultBarcodeMapping(barcode)
    If Len(defaultBarcode) > 0 Then
        LookupBarcode = defaultBarcode
        Exit Function
    End If

    LookupBarcode = ""
End Function

Public Sub ScanBarcode()
    Dim barcode As String
    Dim articleCode As String
    Dim articleDesig As String

    barcode = InputBox("Scannez le code-barres:" & vbCrLf & vbCrLf & _
                       "Placez le curseur dans la zone, puis scannez.", _
                       "Lecteur Code-Barres", "")

    If Len(Trim(barcode)) = 0 Then Exit Sub

    articleCode = LookupBarcode(Trim(barcode))

    If Len(articleCode) = 0 Then
        Dim choice As VbMsgBoxResult
        choice = MsgBox("Code-barres non reconnu: " & barcode & vbCrLf & vbCrLf & _
                       "Voulez-vous en[ARTICLE_DESC]r ce code-barres pour un article?", _
                       vbYesNo + vbQuestion, "Code-barres inconnu")
        If choice = vbYes Then
            Call RegisterBarcode(barcode)
        End If
        Exit Sub
    End If

    articleDesig = mod_Utilities.GetArticleField(articleCode, "DESIG")

    Dim result As VbMsgBoxResult
    result = MsgBox("Article trouv" & Chr(233) & ":" & vbCrLf & vbCrLf & _
                   "Code    : " & articleCode & vbCrLf & _
                   "D" & Chr(233) & "signation : " & articleDesig & vbCrLf & vbCrLf & _
                   "Ouvrir le formulaire de saisie?", _
                   vbYesNo + vbInformation, "Code-barres reconnu")

    If result = vbYes Then
        frmStockEntry.Show
    End If
End Sub

Public Sub RegisterBarcode(ByVal barcode As String)
    Dim ws As Worksheet
    Dim articleCode As String
    Dim lastRow As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(BARCODE_SHEET)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = BARCODE_SHEET
    End If
    ws.Unprotect Password:=mod_Config.MASTER_PWD
    On Error GoTo 0

    articleCode = InputBox("Code article pour le code-barres " & barcode & ":" & vbCrLf & _
                          "Exemple: ART-001", "Associer code-barres", "ART-")

    If Len(Trim(articleCode)) = 0 Then Exit Sub
    articleCode = Trim(UCase(articleCode))

    If Not (articleCode Like "ART-###") Then
        MsgBox "Format de code article invalide. Utilisez ART-001.", vbExclamation
        Exit Sub
    End If

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    Dim found As Boolean
    found = False
    Dim i As Long
    For i = 1 To lastRow
        If Trim(UCase(CStr(ws.Cells(i, 1).Value))) = BARCODE_RANGE Then
            Dim j As Long
            j = i + 1
            Do While j <= ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
                If Trim(UCase(CStr(ws.Cells(j, 1).Value))) = UCase(barcode) Then
                    ws.Cells(j, 2).Value = articleCode
                    found = True
                    Exit Do
                End If
                If Len(Trim(CStr(ws.Cells(j, 1).Value))) = 0 Then
                    ws.Cells(j, 1).Value = barcode
                    ws.Cells(j, 2).Value = articleCode
                    found = True
                    Exit Do
                End If
                j = j + 1
            Loop
            Exit For
        End If
    Next i

    If Not found Then
        ws.Cells(lastRow, 1).Value = BARCODE_RANGE
        ws.Cells(lastRow + 1, 1).Value = barcode
        ws.Cells(lastRow + 1, 2).Value = articleCode
    End If

    On Error Resume Next
    ws.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    On Error GoTo 0

    MsgBox "Code-barres " & barcode & " associ" & Chr(233) & " " & Chr(224) & " " & articleCode & ".", _
           vbInformation, "Code-barres enregistr" & Chr(233)
End Sub

Public Sub SetupDefaultBarcodes()
    Dim ws As Worksheet
    Dim i As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(BARCODE_SHEET)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = BARCODE_SHEET
    End If
    ws.Unprotect Password:=mod_Config.MASTER_PWD
    On Error GoTo 0

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 1 To lastRow
        If Trim(UCase(CStr(ws.Cells(i, 1).Value))) = BARCODE_RANGE Then
            MsgBox "Les codes-barres par d" & Chr(233) & "faut existent d" & Chr(233) & "j" & Chr(224) & ".", vbInformation
            GoTo CleanUpBarcode
        End If
    Next i

    Dim nextRow As Long
    nextRow = lastRow + 1
    ws.Cells(nextRow, 1).Value = BARCODE_RANGE
    ws.Cells(nextRow, 2).Value = "Default barcode mapping"

    Dim articles As Variant
    articles = Array("ART-001", "ART-002", "ART-003", "ART-004", "ART-005", _
                     "ART-006", "ART-007", "ART-008", "ART-009", "ART-010", _
                     "ART-011", "ART-012")

    For i = LBound(articles) To UBound(articles)
        Dim code As String
        code = articles(i)
        ws.Cells(nextRow + 1 + i, 1).Value = Format(i + 1, "000")
        ws.Cells(nextRow + 1 + i, 2).Value = code
    Next i

    MsgBox "Codes-barres par d" & Chr(233) & "faut install" & Chr(233) & "s (12 articles).", vbInformation, "Setup Barcode"

CleanUpBarcode:
    On Error Resume Next
    ws.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    On Error GoTo 0
End Sub

Private Function GetDefaultBarcodeMapping(ByVal barcode As String) As String
    Select Case Trim(UCase(barcode))
        Case "001": GetDefaultBarcodeMapping = "ART-001"
        Case "002": GetDefaultBarcodeMapping = "ART-002"
        Case "003": GetDefaultBarcodeMapping = "ART-003"
        Case "004": GetDefaultBarcodeMapping = "ART-004"
        Case "005": GetDefaultBarcodeMapping = "ART-005"
        Case "006": GetDefaultBarcodeMapping = "ART-006"
        Case "007": GetDefaultBarcodeMapping = "ART-007"
        Case "008": GetDefaultBarcodeMapping = "ART-008"
        Case "009": GetDefaultBarcodeMapping = "ART-009"
        Case "010": GetDefaultBarcodeMapping = "ART-010"
        Case "011": GetDefaultBarcodeMapping = "ART-011"
        Case "012": GetDefaultBarcodeMapping = "ART-012"
        Case Else:  GetDefaultBarcodeMapping = ""
    End Select
End Function
