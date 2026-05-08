Attribute VB_Name = "mod_Dashboard"
Option Explicit

'=======================================================================================
' MODULE: mod_Dashboard.bas
' PROJECT: ERP LSM v1.0.0
' PURPOSE: Dynamic Dashboard Controller
'
' This module transforms the static DASHBOARD sheet into a live monitoring console.
' It reads the computed metrics from the ARTICLES and MOUVEMENTS sheets (via mod_SyncBridge)
' and visualizes the critical stock state, ABC-XYZ distribution, and global KPIs.
'=======================================================================================

'--------------------------------------------------------------------------------------
' MAIN ENTRY POINT: RefreshDashboard
' Updates all KPIs and tables on the DASHBOARD sheet.
'--------------------------------------------------------------------------------------
Public Sub RefreshDashboard()
    Dim ws As Worksheet
    
    ' 1. Get or Create Dashboard Sheet
    Set ws = GetOrCreateDashboardSheet()
    
    ' 2. Update Global KPIs
    Call UpdateKPIs(ws)
    
    ' 3. Update Critical Items Table
    Call UpdateCriticalTable(ws)
    
    ' 4. Update ABC-XYZ Summary
    Call UpdateABCXYZSummary(ws)
    
    ' 5. Final Touch: Update timestamp
    ws.Range("B1").Value = "Derniere actualisation : " & Format(Now, "DD/MM/YYYY HH:MM:SS")
    ws.Range("B1").Font.Size = 8
    ws.Range("B1").Font.Italic = True
    
    MsgBox "Tableau de bord actualise avec succes !", vbInformation, "Dashboard Sync"
End Sub

'--------------------------------------------------------------------------------------
' HELPER: Update Global KPIs
'--------------------------------------------------------------------------------------
Private Sub UpdateKPIs(ws As Worksheet)
    Dim wsArt As Worksheet: Set wsArt = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    Dim totalSKUs As Long
    Dim countRupture As Long
    Dim countAlert As Long
    Dim totalValue As Double
    
    totalSKUs = wsArt.Cells(wsArt.Rows.count, 1).End(xlUp).Row - 1
    If totalSKUs < 0 Then totalSKUs = 0
    
    ' Iterate through ARTICLES to find status and total value
    Dim i As Long
    For i = 2 To totalSKUs + 1
        Dim stock As Long: stock = Val(wsArt.Cells(i, 3).Value) ' Col C: Stock
        Dim pu As Double: pu = Val(wsArt.Cells(i, 8).Value)     ' Col H: PU
        
        totalValue = totalValue + (stock * pu)
        
        ' Simple status check based on ROP/SS from mod_StockEngine
        Dim sku As String: sku = Trim(wsArt.Cells(i, 1).Value)
        Dim ss As Double: ss = mod_StockEngine.GetSafetyStock(sku)
        Dim AnnualDemand As Double: AnnualDemand = mod_StockEngine.GetAnnualDemandFromHistory(sku)
        Dim rop As Double: rop = mod_StockEngine.ComputeROP(AnnualDemand / mod_Config.WORKING_DAYS_PER_YEAR, sku)
        
        If stock <= 0 Then
            countRupture = countRupture + 1
        ElseIf stock <= rop Then
            countAlert = countAlert + 1
        End If
    Next i
    
    ' Layout KPIs
    ws.Range("B2").Value = "Total Articles"
    ws.Range("C2").Value = totalSKUs
    
    ws.Range("B3").Value = "Articles en RUPTURE"
    ws.Range("C3").Value = countRupture
    ws.Range("C3").Font.Color = RGB(200, 0, 0)
    
    ws.Range("B4").Value = "Articles en ALERTE"
    ws.Range("C4").Value = countAlert
    ws.Range("C4").Font.Color = RGB(200, 100, 0)
    
    ws.Range("B5").Value = "Valeur Total Stock"
    ws.Range("C5").Value = totalValue
    ws.Range("C5").NumberFormat = "#,##0.00 ""DZD"""
    
    ws.Range("B6").Value = "Rotation Moyenne (ITR)"
    ' Simplified ITR for Dashboard
    ws.Range("C6").Value = "Calcul Local"
    
    ' Formatting
    With ws.Range("B2:B6")
        .Font.Bold = True
        .HorizontalAlignment = xlRight
    End With
    With ws.Range("C2:C5")
        .Font.Size = 12
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
    End With
End Sub

'--------------------------------------------------------------------------------------
' HELPER: Update Critical Items Table (Top 5 closest to rupture)
'--------------------------------------------------------------------------------------
Private Sub UpdateCriticalTable(ws As Worksheet)
    ' Header
    ws.Range("D2:G2").Value = Array("SKU", "Designation", "Stock", "Etat")
    ws.Range("D2:G2").Interior.Color = RGB(0, 70, 127)
    ws.Range("D2:G2").Font.Color = RGB(255, 255, 255)
    ws.Range("D2:G2").Font.Bold = True
    ws.Range("D2:G2").HorizontalAlignment = xlCenter
    
    Dim wsArt As Worksheet: Set wsArt = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    Dim lastRow As Long: lastRow = wsArt.Cells(wsArt.Rows.count, 1).End(xlUp).Row
    
    ' Store critical items in a temporary array
    Dim criticalList(1 To 1000, 1 To 4) As Variant
    Dim countCrit As Integer: countCrit = 0
    
    Dim i As Long
    For i = 2 To lastRow
        Dim sku As String: sku = Trim(wsArt.Cells(i, 1).Value)
        Dim stock As Long: stock = Val(wsArt.Cells(i, 3).Value)
        
        Dim AnnualDemand As Double: AnnualDemand = mod_StockEngine.GetAnnualDemandFromHistory(sku)
        Dim rop As Double: rop = mod_StockEngine.ComputeROP(AnnualDemand / mod_Config.WORKING_DAYS_PER_YEAR, sku)
        
        If stock <= rop Then
            countCrit = countCrit + 1
            If countCrit > 1000 Then Exit For
            
            criticalList(countCrit, 1) = sku
            criticalList(countCrit, 2) = wsArt.Cells(i, 2).Value
            criticalList(countCrit, 3) = stock
            criticalList(countCrit, 4) = IIf(stock <= 0, "RUPTURE", "ALERTE")
        End If
    Next i
    
    ' Write Top 5 to sheet
    Dim rowNum As Integer: rowNum = 3
    For i = 1 To countCrit
        If i > 5 Then Exit For
        ws.Cells(rowNum, 4).Value = criticalList(i, 1)
        ws.Cells(rowNum, 5).Value = criticalList(i, 2)
        ws.Cells(rowNum, 6).Value = criticalList(i, 3)
        ws.Cells(rowNum, 7).Value = criticalList(i, 4)
        
        If criticalList(i, 3) <= 0 Then
            ws.Range("D" & rowNum & ":G" & rowNum).Interior.Color = RGB(255, 200, 200)
        End If
        rowNum = rowNum + 1
    Next i
    
    ws.Range("D3:G" & rowNum - 1).Borders.LineStyle = xlContinuous
    ws.Columns("D:G").AutoFit
End Sub

'--------------------------------------------------------------------------------------
' HELPER: Update ABC-XYZ Summary
'--------------------------------------------------------------------------------------
Private Sub UpdateABCXYZSummary(ws As Worksheet)
    ws.Range("I2").Value = "Classe"
    ws.Range("J2").Value = "Nombre d'articles"
    ws.Range("I2:J2").Interior.Color = RGB(0, 70, 127)
    ws.Range("I2:J2").Font.Color = RGB(255, 255, 255)
    ws.Range("I2:J2").Font.Bold = True
    
    Dim wsArt As Worksheet: Set wsArt = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    Dim lastRow As Long: lastRow = wsArt.Cells(wsArt.Rows.count, 1).End(xlUp).Row
    
    Dim classes As Variant: classes = Array("A", "B", "C")
    Dim rowNum As Integer: rowNum = 3
    
    Dim c As Integer
    For c = 0 To UBound(classes)
        Dim cls As String: cls = classes(c)
        Dim countCls As Long: countCls = 0
        
        Dim i As Long
        For i = 2 To lastRow
            If wsArt.Cells(i, 5).Value = cls Then ' Col E: ABC Class
                countCls = countCls + 1
            End If
        Next i
        
        ws.Cells(rowNum, 9).Value = cls
        ws.Cells(rowNum, 10).Value = countCls
        rowNum = rowNum + 1
    Next c
    
    ws.Range("I3:J" & rowNum - 1).Borders.LineStyle = xlContinuous
    ws.Columns("I:J").AutoFit
End Sub

'--------------------------------------------------------------------------------------
' HELPER: Get or Create Dashboard Sheet
'--------------------------------------------------------------------------------------
Private Function GetOrCreateDashboardSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("DASHBOARD")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
        ws.name = "DASHBOARD"
    End If
    
    ' Basic Clean
    ws.Cells.Clear
    ws.Cells.Interior.Color = RGB(245, 245, 245)
    
    Set GetOrCreateDashboardSheet = ws
End Function
