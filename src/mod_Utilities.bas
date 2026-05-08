Attribute VB_Name = "mod_Utilities"
'=======================================================================================
' MODULE: mod_Utilities.bas
' PROJECT: ERP Acad�mie v13
' DESCRIPTION: General utility functions used across the application.
'=======================================================================================
Option Explicit

'=======================================================================================
' SUB: RestoreMouvementsHeaders
' Forces the correct headers into Row 1 of the MOUVEMENTS sheet to ensure PDF export works.
'=======================================================================================
Public Sub RestoreMouvementsHeaders(Optional ByVal silent As Boolean = False)
    Dim wsMouv As Worksheet
    Set wsMouv = ThisWorkbook.Sheets(mod_Config.SHEET_MOUVEMENTS)
    
    wsMouv.Unprotect Password:=mod_Config.MASTER_PWD
    
    Dim headers As Variant
    headers = Array("DATE", "CODE_ARTICLE", "DESIGNATION", "TYPE_MVT", "QTE", "VALEUR", "REF_DOCUMENT", "PRIX_UNITAIRE", "THIRD_PARTY", "NOTES")
    
    wsMouv.Range("A1:J1").Value = headers
    wsMouv.Range("A1:J1").Font.Bold = True
    wsMouv.Range("A1:J1").Interior.Color = RGB(200, 200, 200)
    
    wsMouv.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    If Not silent Then
        MsgBox "Headers restored!", vbInformation, mod_Config.SYS_TITLE
    End If
End Sub

'=======================================================================================
' FUNCTION: SafeVal
' Safe numeric cast � returns 0 for empty / non-numeric values
'=======================================================================================
Public Function SafeVal(ByVal v As Variant) As Double
    If IsNumeric(v) And Not IsEmpty(v) Then
        SafeVal = CDbl(v)
    Else
        SafeVal = 0
    End If
End Function

'-------------------------------------------------------------------------------------' FUNCTION: GetArticleField
' Retrieves a specific field (CODE, DESIG, QTE, PU, CAT) for a given SKU from the ARTICLES sheet
'--------------------------------------------------------------------------------------
Public Function GetArticleField(ByVal sku As String, ByVal fieldType As String) As String
    Dim wsArt    As Worksheet
    Dim foundRow As Variant
    Dim colIdx   As Integer
    
    On Error Resume Next
    Set wsArt = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    On Error GoTo 0
    
    If wsArt Is Nothing Then
        GetArticleField = ""
        Exit Function
    End If
    
    ' Determine column based on fieldType
    Select Case UCase(fieldType)
        Case "CODE":  colIdx = 1  ' Column A: Article Code
        Case "DESIG": colIdx = 2  ' Column B: Designation
        Case "QTE":   GetArticleField = "0": Exit Function  ' No stock column in ARTICLES
        Case "PU":    colIdx = 8  ' Column H: Prix Unitaire
        Case "CAT":   colIdx = 5  ' Column E: Categorie
        Case Else:    colIdx = 2  ' Default to Designation
    End Select
    
    foundRow = Application.Match(sku, wsArt.Range("A:A"), 0)
    
    If IsError(foundRow) Then
        GetArticleField = ""
    Else
        GetArticleField = Trim(CStr(wsArt.Cells(foundRow, colIdx).Value))
    End If
End Function

'--------------------------------------------------------------------------------------
' FUNCTION: IsValidDate
' Validates if a string is a valid date in strict DD/MM/YYYY format
'--------------------------------------------------------------------------------------
Public Function IsValidDate(ByVal dateStr As String) As Boolean
    If Len(Trim(dateStr)) <> 10 Then
        IsValidDate = False
        Exit Function
    End If
    
    ' Check slash positions (DD/MM/YYYY)
    If Mid(dateStr, 3, 1) <> "/" Or Mid(dateStr, 6, 1) <> "/" Then
        IsValidDate = False
        Exit Function
    End If
    
    ' Validate actual date components
    On Error Resume Next
    Dim dayPart As Integer, monthPart As Integer, yearPart As Integer
    dayPart = CInt(Mid(dateStr, 1, 2))
    monthPart = CInt(Mid(dateStr, 4, 2))
    yearPart = CInt(Right(dateStr, 4))
    
    Dim testDate As Date
    testDate = DateSerial(yearPart, monthPart, dayPart)
    IsValidDate = (Err.Number = 0)
    On Error GoTo 0
End Function

'=======================================================================================
' SUB: SetupLocationDropdown
' Creates a dynamic dropdown list for EMPLACEMENT column (Col H)
'=======================================================================================
Public Sub SetupLocationDropdown()
    Dim wsMaster As Worksheet, wsLists As Worksheet
    Dim lastRow As Long
    Dim listRange As Range
    
    Set wsMaster = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    
    ' IMPORTANT: Unprotect sheet to allow modification of Data Validation
    wsMaster.Unprotect Password:=mod_Config.MASTER_PWD
    
    On Error Resume Next
    Set wsLists = ThisWorkbook.Sheets("System_Lists")
    If wsLists Is Nothing Then
        Set wsLists = ThisWorkbook.Worksheets.Add(After:=wsMaster)
        wsLists.name = "System_Lists"
        wsLists.Visible = xlSheetVeryHidden
    End If
    On Error GoTo ErrorHandler
    wsLists.Cells.Clear
    
    lastRow = wsMaster.Cells(wsMaster.Rows.count, "H").End(xlUp).Row
    
    If lastRow > 1 Then
        wsMaster.Range("H2:H" & lastRow).Copy Destination:=wsLists.Range("A1")
        wsLists.Range("A:A").RemoveDuplicates Columns:=1, Header:=xlNo
        
        Dim listCount As Long
        listCount = wsLists.Cells(wsLists.Rows.count, "A").End(xlUp).Row
        
        If listCount < 1 Then
            MsgBox "No location data found in Column H of ARTICLES sheet.", vbExclamation, mod_Config.SYS_TITLE
            Exit Sub
        End If
        
        Set listRange = wsLists.Range("A1:A" & listCount)
        
        With wsMaster.Range("H2:H1000").Validation
            .Delete
            ' Use absolute reference and ensure the range is valid
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                 xlBetween, Formula1:="='System_Lists'!$A$1:$A$" & listCount
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = "Select Location"
            .InputMessage = "Please pick a valid warehouse zone from the list."
            .ErrorMessage = "This location doesn't exist. Please add it to the system first."
            .ShowInput = True
            .ShowError = True
        End With
    Else
        MsgBox "No location data found to synchronize.", vbInformation, mod_Config.SYS_TITLE
    End If
    
    ' Reprotect the sheet with UserInterfaceOnly:=True to allow future VBA modifications
    wsMaster.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    
    Application.ScreenUpdating = True
    MsgBox "Dropdown list for EMPLACEMENT has been synchronized!", vbInformation, mod_Config.SYS_TITLE
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Failed to setup dropdown: " & Err.Description, vbCritical, "VBA Error"
End Sub

'=======================================================================================
' SUB: ApplyInventoryHeatmap
' Applies ABC-aware conditional formatting to QTE_STOCK column
'=======================================================================================
Public Sub ApplyInventoryHeatmap()
    Dim ws As Worksheet
    Dim stockRange As Range
    Dim lastRow As Long
    
    Set ws = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    
    Set stockRange = ws.Range("C2:C" & lastRow)
    
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    
    ' IMPORTANT: Unprotect sheet to allow modification of Conditional Formatting
    ws.Unprotect Password:=mod_Config.MASTER_PWD
    
    stockRange.FormatConditions.Delete
    
    With stockRange.FormatConditions.Add(Type:=xlExpression, Formula1:="=$C2<=$F2")
        .Interior.Color = RGB(255, 199, 206)
        .Font.Color = RGB(156, 0, 6)
        .Font.Bold = True
    End With
    
    With stockRange.FormatConditions.Add(Type:=xlExpression, Formula1:="=$C2<=($F2*1.2)")
        .Interior.Color = RGB(255, 235, 156)
        .Font.Color = RGB(156, 101, 0)
    End With
    
    Application.ScreenUpdating = True
    
    ' Reprotect the sheet to ensure integrity
    ws.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    
    MsgBox "Stock heatmap applied successfully!", vbInformation, mod_Config.SYS_TITLE
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Formatting Error: " & Err.Description, vbCritical, mod_Config.SYS_TITLE
End Sub

'=======================================================================================
' SUB: ExportLowStockPDF
' Generates a Low Stock Report PDF for management
'=======================================================================================
Public Sub ExportLowStockPDF()
    Dim wsSource As Worksheet, wsReport As Worksheet
    Dim lastRow As Long, reportRow As Long, i As Long
    Dim pdfPath As String
    Dim reportName As String
    
    Set wsSource = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    reportName = "Low_Stock_Report_" & Format(Date, "yyyy-mm-dd")
    pdfPath = ThisWorkbook.Path & "\" & reportName & ".pdf"
    
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    
    Set wsReport = Worksheets.Add
    wsReport.name = "TEMP_REPORT"
    
    With wsReport
        .Range("A1").Value = "LOW STOCK ALERT - " & mod_Config.SYS_TITLE
        .Range("A2").Value = "Directorate of Education - [CITY]"
        .Range("A3").Value = "Date: " & Now
        
        .Range("A5").Value = "CODE"
        .Range("B5").Value = "DESIGNATION (AR)"
        .Range("C5").Value = "STOCK"
        .Range("D5").Value = "MIN"
        .Range("E5").Value = "LOCATION"
    End With
    
    lastRow = wsSource.Cells(wsSource.Rows.count, "A").End(xlUp).Row
    reportRow = 6
    
    For i = 3 To lastRow
        If wsSource.Cells(i, 3).Value <= wsSource.Cells(i, 6).Value Then
            wsReport.Cells(reportRow, 1).Value = wsSource.Cells(i, 1).Value
            wsReport.Cells(reportRow, 2).Value = wsSource.Cells(i, 2).Value
            wsReport.Cells(reportRow, 3).Value = wsSource.Cells(i, 3).Value
            wsReport.Cells(reportRow, 4).Value = wsSource.Cells(i, 6).Value
            wsReport.Cells(reportRow, 5).Value = wsSource.Cells(i, 8).Value
            reportRow = reportRow + 1
        End If
    Next i
    
    With wsReport.Range("A5:E" & reportRow - 1)
        .Borders.LineStyle = xlContinuous
        .Columns.AutoFit
    End With
    wsReport.Range("A5:E5").Interior.Color = RGB(0, 51, 102)
    wsReport.Range("A5:E5").Font.Color = vbWhite
    wsReport.Range("A5:E5").Font.Bold = True
    
    wsReport.ExportAsFixedFormat Type:=xlTypePDF, fileName:=pdfPath, Quality:=xlQualityStandard
    
    Application.DisplayAlerts = False
    wsReport.Delete
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    MsgBox "Report generated successfully: " & pdfPath, vbInformation, mod_Config.SYS_TITLE
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Failed to generate PDF: " & Err.Description, vbCritical, mod_Config.SYS_TITLE
    On Error Resume Next
    If Not wsReport Is Nothing Then
        Application.DisplayAlerts = False
        wsReport.Delete
        Application.DisplayAlerts = True
    End If
End Sub

'================================================================================
' FUNCTION: GenerateVerifyCode
' Unified verification code generator - replaces 3 duplicate implementations
' in mod_ExportEngine, mod_QRCode, and mod_ReceiptTag.
' Returns format: V-XXXX-XXXX-XXXX
'================================================================================
Public Function GenerateVerifyCode(ByVal rawData As String) As String
    Dim checksum As Long
    Dim i As Integer
    Dim chCode As Integer
    Dim hexPart1 As String, hexPart2 As String, hexPart3 As String

    checksum = 0
    For i = 1 To Len(rawData)
        chCode = Asc(Mid(rawData, i, 1))
        checksum = checksum + (chCode * (i + 7)) Mod 9973
    Next i

    hexPart1 = Right("0000" & Hex(checksum And &HFFFF&), 4)
    hexPart2 = Right("0000" & Hex((checksum \ 17) And &HFFFF&), 4)
    hexPart3 = Right("0000" & Hex((checksum \ 257) And &HFFFF&), 4)

    GenerateVerifyCode = "V-" & hexPart1 & "-" & hexPart2 & "-" & hexPart3
End Function
