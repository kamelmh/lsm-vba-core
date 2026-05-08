Attribute VB_Name = "mod_Reports"
'=======================================================================================
' MODULE: mod_Reports.bas
' PROJECT: ERP Acad" & Chr(233) & "mie v13
' DESCRIPTION: Generates summary reports for management.
'=======================================================================================
Option Explicit

'=======================================================================================
' SUB: GenerateMonthlyReport
' Creates monthly summary report (inputs, outputs, stock status)
'=======================================================================================
Public Sub GenerateMonthlyReport(Optional ByVal rptMonth As Integer = 0)
    Dim wsMouv As Worksheet, wsArt As Worksheet, wsReport As Worksheet
    Dim desktopPath As String, fileName As String, fullPath As String
    
    If rptMonth = 0 Then rptMonth = Month(Date)
    
    On Error GoTo ReportError
    
    Set wsMouv = ThisWorkbook.Sheets(mod_Config.SHEET_MOUVEMENTS)
    Set wsArt = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    
    desktopPath = mod_SharedEnvironment.GetSharedExportPath()
    fileName = "Rapport_Mensuel_" & Format(Date, "yyyy-mm") & ".xlsx"
    fullPath = desktopPath & fileName
    
    Application.ScreenUpdating = False
    Set wsReport = Worksheets.Add
    wsReport.name = "RAPPORT_MENSUEL"
    
    With wsReport
        .Range("A1").Value = "RAPPORT MENSUEL - " & mod_Config.SYS_TITLE
        .Range("A2").Value = "Direction de l'" & Chr(201) & "ducation - [CITY]"
        .Range("A3").Value = "Mois: " & Format(Date, "mmmm yyyy")
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        
        .Range("A5").Value = "CODE"
        .Range("B5").Value = "D" & Chr(201) & "SIGNATION"
        .Range("C5").Value = "ENTR" & Chr(201) & "ES"
        .Range("D5").Value = "SORTIES"
        .Range("E5").Value = "STOCK FINAL"
        .Range("F5").Value = "VALEUR (DZD)"
        .Range("G5").Value = "CLASSE ABC"
        .Range("H5").Value = "CONTRIB. %"
        .Range("A5:H5").Font.Bold = True
        .Range("A5:H5").Interior.Color = RGB(0, 112, 192)
        .Range("A5:H5").Font.Color = vbWhite
    End With
    
    Dim lastArtRow As Long: lastArtRow = wsArt.Cells(wsArt.Rows.count, 1).End(xlUp).Row
    Dim reportRow As Integer: reportRow = 6
    Dim totalIn As Double, totalOut As Double, totalValue As Double
    
    Dim i As Long
    For i = 3 To lastArtRow
        Dim sku As String: sku = Trim(wsArt.Cells(i, 1).Value)
        Dim name As String: name = Trim(wsArt.Cells(i, 2).Value)
        
        If sku <> "" Then
            Dim monthIn As Double, monthOut As Double
            On Error Resume Next
            wsMouv.Unprotect Password:=mod_Config.MASTER_PWD
            On Error GoTo ReportError
            monthIn = WorksheetFunction.SumIfs(wsMouv.Range("E:E"), wsMouv.Range("B:B"), sku, wsMouv.Range("D:D"), "IN", _
                                             wsMouv.Range("A:A"), ">=" & DateSerial(Year(Date), rptMonth, 1), _
                                             wsMouv.Range("A:A"), "<" & DateSerial(Year(Date), rptMonth + 1, 1))
            monthOut = WorksheetFunction.SumIfs(wsMouv.Range("E:E"), wsMouv.Range("B:B"), sku, wsMouv.Range("D:D"), "OUT", _
                                               wsMouv.Range("A:A"), ">=" & DateSerial(Year(Date), rptMonth, 1), _
                                               wsMouv.Range("A:A"), "<" & DateSerial(Year(Date), rptMonth + 1, 1))
            wsMouv.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
            
            Dim totalStock As Double: totalStock = wsArt.Cells(i, 7).Value
            Dim pu As Double: pu = wsArt.Cells(i, 6).Value
            Dim rowValue As Double: rowValue = totalStock * pu
            
            With wsReport
                .Cells(reportRow, 1).Value = sku
                .Cells(reportRow, 2).Value = name
                .Cells(reportRow, 3).Value = monthIn
                .Cells(reportRow, 4).Value = monthOut
                .Cells(reportRow, 5).Value = totalStock
                .Cells(reportRow, 6).Value = rowValue
                .Cells(reportRow, 6).NumberFormat = "#,##0.00"
                
                ' --- Enhanced Metrics ---
                ' Pull ABC Class from ARTICLES sheet (Col E = 5)
                .Cells(reportRow, 7).Value = wsArt.Cells(i, 5).Value
                
                ' Contribution % (calculated at the end or updated later)
                ' We'll leave Col H for now and fill it in a second pass
                .Cells(reportRow, 8).Value = 0
            End With
            
            totalIn = totalIn + monthIn
            totalOut = totalOut + monthOut
            totalValue = totalValue + rowValue
            reportRow = reportRow + 1
        End If
    Next i
    
    ' --- Final Totals Row ---
    With wsReport
        .Range("A" & reportRow + 1).Value = "TOTAUX"
        .Range("A" & reportRow + 1).Font.Bold = True
        .Range("C" & reportRow + 1).Value = totalIn
        .Range("D" & reportRow + 1).Value = totalOut
        .Range("F" & reportRow + 1).Value = totalValue
        .Range("F" & reportRow + 1).NumberFormat = "#,##0.00"
        .Range("A" & reportRow + 1 & ":H" & reportRow + 1).Interior.Color = RGB(217, 217, 217)
        
        ' Calculate contribution percentages for each row
        Dim r As Long
        For r = 6 To reportRow
            Dim valLigne As Double: valLigne = .Cells(r, 6).Value
            If totalValue > 0 Then
                .Cells(r, 8).Value = valLigne / totalValue
                .Cells(r, 8).NumberFormat = "0.00%"
            End If
        Next r
        
        .Columns("A:H").AutoFit
    End With
    
    wsReport.SaveAs fullPath
    wsReport.Delete
    Application.ScreenUpdating = True
    MsgBox "Rapport g�n�r�: " & fullPath, vbInformation, mod_Config.SYS_TITLE
    Exit Sub
ReportError:
    Application.ScreenUpdating = True
    MsgBox "Erreur rapport: " & Err.Description, vbCritical
End Sub

'=======================================================================================
' SUB: GenerateStockCard
' Creates stock card (fiche de stock) for a specific article
'=======================================================================================
Public Sub GenerateStockCard(Optional ByVal sku As String = "")
    Dim wsArt As Worksheet, wsMouv As Worksheet, wsCard As Worksheet
    Dim desktopPath As String, fileName As String, fullPath As String
    
    On Error GoTo ReportError
    Set wsArt = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    Set wsMouv = ThisWorkbook.Sheets(mod_Config.SHEET_MOUVEMENTS)
    
    If sku = "" Then
        sku = InputBox("Entrez le code article:", "Fiche de Stock")
        If sku = "" Then Exit Sub
    End If
    
    Dim artRow As Variant: artRow = Application.Match(sku, wsArt.Range("A:A"), 0)
    If IsError(artRow) Then
        MsgBox "Article non trouv�: " & sku, vbExclamation
        Exit Sub
    End If
    
    desktopPath = mod_SharedEnvironment.GetSharedExportPath()
    fileName = "Fiche_Stock_" & sku & ".xlsx"
    fullPath = desktopPath & fileName
    
    Application.ScreenUpdating = False
    Set wsCard = Worksheets.Add
    wsCard.name = "FICHE_STOCK"
    
    With wsCard
        .Range("A1").Value = "FICHE DE STOCK - " & sku
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        .Range("A3").Value = "CODE:": .Range("B3").Value = sku
        .Range("A4").Value = "D" & Chr(233) & "SIGNATION:": .Range("B4").Value = wsArt.Cells(artRow, 2).Value
        .Range("A5").Value = "STOCK ACTUEL:": .Range("B5").Value = wsArt.Cells(artRow, 7).Value
        .Range("A6").Value = "PRIX UNITAIRE:": .Range("B6").Value = wsArt.Cells(artRow, 6).Value
        .Range("A8").Value = "DATE": .Range("B8").Value = "TYPE": .Range("C8").Value = "QT" & Chr(233): .Range("D8").Value = "VALEUR": .Range("E8").Value = "R" & Chr(233) & "F" & Chr(233) & "RENCE"
        .Range("A8:E8").Font.Bold = True
        .Range("A8:E8").Interior.Color = RGB(0, 112, 192)
        .Range("A8:E8").Font.Color = vbWhite
    End With
    
    Dim lastMvtRow As Long: lastMvtRow = wsMouv.Cells(wsMouv.Rows.count, 1).End(xlUp).Row
    Dim cardRow As Integer: cardRow = 9
    Dim j As Long
    On Error Resume Next
    wsMouv.Unprotect Password:=mod_Config.MASTER_PWD
    On Error GoTo ReportError
    For j = 2 To lastMvtRow
        If Trim(wsMouv.Cells(j, 2).Value) = sku Then
            wsCard.Cells(cardRow, 1).Value = wsMouv.Cells(j, 1).Value
            wsCard.Cells(cardRow, 2).Value = wsMouv.Cells(j, 3).Value
            wsCard.Cells(cardRow, 3).Value = wsMouv.Cells(j, 5).Value
            wsCard.Cells(cardRow, 4).Value = wsMouv.Cells(j, 7).Value
            wsCard.Cells(cardRow, 5).Value = wsMouv.Cells(j, 8).Value
            cardRow = cardRow + 1
        End If
    Next j
    wsMouv.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    
    wsCard.Columns("A:E").AutoFit
    wsCard.SaveAs fullPath
    wsCard.Delete
    Application.ScreenUpdating = True
    MsgBox "Fiche de stock g�n�r�e: " & fullPath, vbInformation, mod_Config.SYS_TITLE
    Exit Sub
ReportError:
    Application.ScreenUpdating = True
    MsgBox "Erreur: " & Err.Description, vbCritical
End Sub
