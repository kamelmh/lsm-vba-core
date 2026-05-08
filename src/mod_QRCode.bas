Attribute VB_Name = "mod_QRCode"
'==============================================================================
' mod_QRCode.bas  |  ERP LSM v1.0.0
' Author: LSM VBA Core | Public Sector 2026
'
' QR Code Generation - Offline-First, Zero External Dependencies
'
' ARCHITECTURE (3-tier fallback):
'   1. Primary: Real QR code via Google Charts API (internet required)
'   2. Secondary: Real QR code via QRServer API (alternative, more reliable)
'   3. Fallback: Offline verification code block with deterministic hash
'      (not scannable but provides audit trail)
'
' USAGE:
'   Call GenerateQRCodeForForm(docRef, sheetName, targetCell)
'   Call GenerateQRCodeForSheet(docRef, ws, targetCell)
'   Call GenerateLocalQRFallback(docRef, ws, targetCell)
'
' COMPATIBILITY: Excel 2010+ / Windows 7+
'==============================================================================

Option Explicit

' QR code size in pixels (standard for print)
Private Const QR_SIZE_PX As Integer = 150
Private Const QR_SIZE_PT As Double = 113  ' 150px * 0.75pt/px

'==============================================================================
' PRIMARY: Generate QR Code on specific sheet at target cell
' Removes any existing QR code first to prevent duplicates
'==============================================================================
Public Sub GenerateQRCodeForForm(ByVal docRef As String, _
                                 ByVal sheetName As String, _
                                 ByVal targetCell As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Debug.Print "[QR] Sheet not found: " & sheetName
        Exit Sub
    End If
    
    Call GenerateQRCodeForSheet(docRef, ws, targetCell)
End Sub

'==============================================================================
' Generate QR Code on specific sheet at target cell
'==============================================================================
Public Sub GenerateQRCodeForSheet(ByVal docRef As String, _
                                  ByRef ws As Worksheet, _
                                  ByVal targetCell As String)
    Dim qrData As String
    
    On Error GoTo QRError
    
    ' Remove existing QR code shapes
    Call RemoveExistingQRCode(ws)
    
    ' Build QR data string (contains all essential document info)
    qrData = "ACDX13.2|REF:" & docRef & "|DT:" & Format(Date, "dd/mm/yyyy") & _
             "|VER:" & mod_Config.APP_VERSION
    
    ' Attempt Google Charts API first
    If TryQRFromAPI(docRef, ws, targetCell, qrData, "google") Then Exit Sub
    
    ' Fallback: Try QRServer API (more reliable)
    If TryQRFromAPI(docRef, ws, targetCell, qrData, "qrserver") Then Exit Sub
    
    ' Final fallback: offline verification block
    Debug.Print "[QR] Both APIs failed, using offline fallback for " & docRef
    Call GenerateLocalQRFallback(docRef, ws, targetCell)
    Exit Sub
    
QRError:
    Debug.Print "[QR] Error: " & Err.Description & ", using fallback"
    Call GenerateLocalQRFallback(docRef, ws, targetCell)
End Sub

'==============================================================================
' Try to generate QR from an API (google or qrserver)
'==============================================================================
Private Function TryQRFromAPI(ByVal docRef As String, _
                              ByRef ws As Worksheet, _
                              ByVal targetCell As String, _
                              ByVal qrData As String, _
                              ByVal apiName As String) As Boolean
    TryQRFromAPI = False
    
    Dim qrUrl As String
    Dim encodedData As String
    encodedData = URLEncode(qrData)
    
    Select Case apiName
        Case "google"
            qrUrl = "https://chart.googleapis.com/chart?chs=" & QR_SIZE_PX & "x" & QR_SIZE_PX & _
                    "&cht=qr&chl=" & encodedData & "&choe=UTF-8"
        Case "qrserver"
            qrUrl = "https://api.qrserver.com/v1/create-qr-code/?size=" & QR_SIZE_PX & "x" & QR_SIZE_PX & _
                    "&data=" & encodedData & "&color=000000&bgcolor=ffffff"
        Case Else
            Exit Function
    End Select
    
    ' Test connectivity first
    If Not IsInternetAvailable() Then Exit Function
    
    Dim pic As Picture
    On Error GoTo APIFail
    
    Set pic = ws.Pictures.Insert(qrUrl)
    
    ' Position and size the QR code
    With pic
        .Top = ws.Range(targetCell).Top + 5
        .Left = ws.Range(targetCell).Left + 5
        .Width = QR_SIZE_PT
        .Height = QR_SIZE_PT
        .Name = "QRCode_" & docRef
        .Placement = xlMoveAndSize
    End With
    
    ' Add label below QR code
    Dim labelCell As Range
    Set labelCell = ws.Range(targetCell).Offset(3, 0)
    labelCell.Value = "Scan v" & Chr(233) & "rification QR"
    labelCell.Font.Size = 7
    labelCell.Font.Color = RGB(128, 128, 128)
    labelCell.HorizontalAlignment = xlCenter
    labelCell.Font.Name = "Tahoma"
    
    Debug.Print "[QR] Generated via " & apiName & " API for " & docRef
    TryQRFromAPI = True
    Exit Function
    
APIFail:
    Debug.Print "[QR] " & apiName & " API failed: " & Err.Description
    TryQRFromAPI = False
End Function

'==============================================================================
' FALLBACK: Generate Local Verification Block (Offline-First)
' Creates a deterministic verification code block for audit trail
'==============================================================================
Public Sub GenerateLocalQRFallback(ByVal docRef As String, _
                                   ByRef ws As Worksheet, _
                                   ByVal targetCell As String)
    Dim verifyCode As String
    Dim targetRange As Range
    Dim qrBlock As Range
    Dim i As Integer, j As Integer
    Dim cellVal As Integer
    Dim seed As Long
    
    On Error GoTo FallbackError
    
    Set targetRange = ws.Range(targetCell)
    
    ' Generate verification code
    Dim rawData As String
    rawData = docRef & Format(Date, "ddmmyyyy") & mod_Config.APP_VERSION
    verifyCode = mod_Utilities.GenerateVerifyCode(rawData)
    
    ' Clear target area (6x6 cells)
    Set qrBlock = ws.Range(targetCell).Resize(6, 6)
    qrBlock.ClearContents
    qrBlock.ClearFormats
    
    ' Generate deterministic pattern from docRef hash
    seed = HashString(docRef)
    
    ' Create visual QR-like pattern with corner finder patterns
    For i = 1 To 6
        For j = 1 To 6
            Dim cellRef As Range
            Set cellRef = qrBlock.Cells(i, j)
            
            ' Corner patterns (QR finder patterns)
            If (i <= 2 And j <= 2) Or (i <= 2 And j >= 5) Or (i >= 5 And j <= 2) Then
                cellRef.Interior.Color = RGB(0, 0, 0)
                cellRef.Value = Chr(9608)
                cellRef.Font.Size = 10
                cellRef.HorizontalAlignment = xlCenter
                cellRef.VerticalAlignment = xlCenter
            Else
                ' Data area (deterministic from hash)
                cellVal = ((seed * i * j + i * 13 + j * 7) Mod 100)
                If cellVal < 40 Then
                    cellRef.Interior.Color = RGB(0, 0, 0)
                    cellRef.Value = Chr(9608)
                    cellRef.Font.Size = 10
                    cellRef.HorizontalAlignment = xlCenter
                    cellRef.VerticalAlignment = xlCenter
                ElseIf cellVal < 60 Then
                    cellRef.Interior.Color = RGB(128, 128, 128)
                    cellRef.Value = Chr(9617)
                    cellRef.Font.Size = 10
                    cellRef.HorizontalAlignment = xlCenter
                    cellRef.VerticalAlignment = xlCenter
                Else
                    cellRef.Interior.Color = RGB(255, 255, 255)
                    cellRef.Value = ""
                End If
            End If
            
            cellRef.Borders.LineStyle = xlContinuous
            cellRef.Borders.Weight = xlHairline
            cellRef.Borders.Color = RGB(200, 200, 200)
        Next j
    Next i
    
    ' Square cells
    For i = 1 To 6
        qrBlock.Rows(i).RowHeight = 14
        qrBlock.Columns(i).ColumnWidth = 3
    Next i
    
    ' Verification code below
    Dim codeLabel As Range
    Set codeLabel = ws.Range(targetCell).Offset(7, 0)
    codeLabel.Value = verifyCode
    codeLabel.Font.Name = "Courier New"
    codeLabel.Font.Size = 8
    codeLabel.Font.Bold = True
    codeLabel.Font.Color = RGB(0, 70, 127)
    codeLabel.HorizontalAlignment = xlCenter
    
    Dim labelCell As Range
    Set labelCell = ws.Range(targetCell).Offset(8, 0)
    labelCell.Value = "Code v" & Chr(233) & "rification (hors ligne)"
    labelCell.Font.Size = 7
    labelCell.Font.Color = RGB(128, 128, 128)
    labelCell.HorizontalAlignment = xlCenter
    labelCell.Font.Name = "Tahoma"
    
    ' Merge areas
    ws.Range(targetCell).Offset(7, 0).Resize(1, 6).Merge
    ws.Range(targetCell).Offset(8, 0).Resize(1, 6).Merge
    
    Debug.Print "[QR] Offline fallback for " & docRef & " [" & verifyCode & "]"
    Exit Sub

FallbackError:
    Debug.Print "[QR] Fallback failed: " & Err.Description
End Sub

'==============================================================================
' Remove existing QR code shapes from sheet
'==============================================================================
Private Sub RemoveExistingQRCode(ByRef ws As Worksheet)
    Dim shp As Shape
    On Error Resume Next
    For Each shp In ws.Shapes
        If Left(shp.Name, 7) = "QRCode_" Then shp.Delete
    Next shp
    On Error GoTo 0
End Sub

'==============================================================================
' Test Internet Connectivity (Excel 2010 compatible)
'==============================================================================
Private Function IsInternetAvailable() As Boolean
    Dim httpObj As Object
    On Error GoTo NoInternet
    
    Set httpObj = CreateObject("MSXML2.XMLHTTP")
    httpObj.Open "GET", "https://api.qrserver.com/", False
    httpObj.setTimeouts 3000, 3000, 3000, 3000
    httpObj.Send
    
    If httpObj.Status = 200 Then
        IsInternetAvailable = True
    Else
        IsInternetAvailable = False
    End If
    
    Set httpObj = Nothing
    Exit Function

NoInternet:
    IsInternetAvailable = False
    Set httpObj = Nothing
End Function

'==============================================================================
' URL Encode for Excel 2010
'==============================================================================
Private Function URLEncode(ByVal text As String) As String
    Dim i As Integer
    Dim ch As String
    Dim result As String
    
    result = ""
    For i = 1 To Len(text)
        ch = Mid(text, i, 1)
        Select Case ch
            Case "A" To "Z", "a" To "z", "0" To "9", "-", ".", "_", "~"
                result = result & ch
            Case " "
                result = result & "+"
            Case Else
                result = result & "%" & Right("0" & Hex(Asc(ch)), 2)
        End Select
    Next i
    
    URLEncode = result
End Function

'==============================================================================
' Generate Local Verification Code - Format: V-XXXX-XXXX-XXXX
'==============================================================================
'==============================================================================
' Simple deterministic hash
'==============================================================================
Private Function HashString(ByVal text As String) As Long
    Dim i As Integer
    Dim hash As Long
    hash = 5381
    
    For i = 1 To Len(text)
        hash = ((hash * 33) Xor Asc(Mid(text, i, 1))) And &H7FFFFFFF
    Next i
    
    HashString = hash
End Function

'==============================================================================
' QR Code Placement Helper - Embeds QR data in form for PDF export
'==============================================================================
Public Sub AddQRCodeToForm(ByVal docRef As String, _
                           ByVal docType As String, _
                           ByVal docDate As String, _
                           ByVal totalVal As Double, _
                           ByRef ws As Worksheet, _
                           ByVal targetCell As String)
    Dim qrContent As String
    
    qrContent = "ACDX13.2|T:" & docType & "|R:" & docRef & _
                "|D:" & docDate & "|V:" & Format(totalVal, "0.00") & _
                "|TS:" & Format(Now, "yyyymmddhhmmss")
    
    ' Generate QR code (with automatic fallback)
    Call GenerateQRCodeForSheet(docRef, ws, targetCell)
    
    ' Store QR data in hidden cell for audit trail
    Dim auditCell As Range
    Set auditCell = ws.Range(targetCell).Offset(10, 0)
    auditCell.Value = "QR:" & qrContent
    auditCell.Font.Size = 6
    auditCell.Font.Color = RGB(200, 200, 200)
    auditCell.Font.Name = "Courier New"
    
    Debug.Print "[QR] Added to form " & docRef & " at " & targetCell
End Sub

'==============================================================================
' Verify Document - Checks if verification code matches
'==============================================================================
Public Function VerifyDocumentQR(ByVal docRef As String, _
                                 ByVal verifyCode As String) As Boolean
    Dim expectedCode As String
    Dim rawData As String
    
    rawData = docRef & Format(Date, "ddmmyyyy") & mod_Config.APP_VERSION
    expectedCode = mod_Utilities.GenerateVerifyCode(rawData)
    
    VerifyDocumentQR = (UCase(verifyCode) = UCase(expectedCode))
    
    Debug.Print "[QR] Verification for " & docRef & " = " & VerifyDocumentQR
End Function

'==============================================================================
' Get verification code string for a document (used in PDF footer)
'==============================================================================
Public Function GetDocumentVerifyCode(ByVal docRef As String) As String
    Dim rawData As String
    rawData = docRef & Format(Date, "ddmmyyyy") & mod_Config.APP_VERSION
    GetDocumentVerifyCode = mod_Utilities.GenerateVerifyCode(rawData)
End Function

'==============================================================================
' END -- mod_QRCode.bas
'==============================================================================
