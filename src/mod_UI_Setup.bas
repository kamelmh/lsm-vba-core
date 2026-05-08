Attribute VB_Name = "mod_UI_Setup"
'=======================================================================================
' MODULE: mod_UI_Setup.bas
' PROJECT: ERP LSM v1.0.0
' PURPOSE: Programmatically creates UI elements on the Dashboard sheet
'=======================================================================================
Option Explicit

'--------------------------------------------------------------------------------------
' MAIN ENTRY POINT: SetupDashboardButtons
' Creates professional buttons on the DASHBOARD sheet.
'--------------------------------------------------------------------------------------
Public Sub SetupDashboardButtons()
    Dim ws As Worksheet
    Dim btn As Button
    Dim rowIdx As Long
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("DASHBOARD")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
        ws.Name = "DASHBOARD"
    End If
    
    ws.Buttons.Delete
    ws.Cells.Interior.Color = RGB(245, 245, 245)
    
    ws.Range("B2").Value = "MANAGEMENT CONSOLE"
    ws.Range("B2").Font.Size = 16
    ws.Range("B2").Font.Bold = True
    ws.Range("B2").Font.Color = RGB(0, 70, 127)
    ws.Range("B4").Value = "Quick Actions:"
    ws.Range("B4").Font.Italic = True
    
    rowIdx = 50

    '-- Section 1: Data Entry & Sync
    AddButton ws, rowIdx, "[ENTRY] Stock Entry Form", "mod_Navigation.OpenStockForm"
    rowIdx = rowIdx + 35
    AddButton ws, rowIdx, "[SYNC] Sync Metrics", "mod_SyncBridge.SyncMetricsFromLedger"
    rowIdx = rowIdx + 35

    '-- Section 2: CSV & Barcode
    AddButton ws, rowIdx, "[CSV] Export MOUVEMENTS", "mod_CSVImportExport.ExportMouvementsToCSV"
    rowIdx = rowIdx + 35
    AddButton ws, rowIdx, "[CSV] Import MOUVEMENTS", "mod_CSVImportExport.ImportMouvementsFromCSV"
    rowIdx = rowIdx + 35
    AddButton ws, rowIdx, "[BARCODE] Scan Article", "mod_Barcode.ScanBarcode"
    rowIdx = rowIdx + 45

    '-- Section 3: Analysis Modules
    AddButton ws, rowIdx, "[RECONCILE] Physical Inventory vs System", "mod_InventoryReconciliation.RunInventoryReconciliation"
    rowIdx = rowIdx + 35
    AddButton ws, rowIdx, "[FORECAST] Stock-Out Prediction", "mod_StockOutPredictor.RunStockOutPrediction"
    rowIdx = rowIdx + 35
    AddButton ws, rowIdx, "[SUPPLIER] Supplier Scorecard", "mod_SupplierScorecard.RunSupplierScorecard"
    rowIdx = rowIdx + 35
    AddButton ws, rowIdx, "[AGING] Stock Aging Report", "mod_StockAging.RunStockAgingReport"
    rowIdx = rowIdx + 35
    AddButton ws, rowIdx, "[VALIDATE] Data Integrity Check", "mod_DataValidator.RunDataValidation"
    rowIdx = rowIdx + 45

    '-- Section 4: Reports & Utilities
    AddButton ws, rowIdx, "[REPORT] Generate Monthly Report", "mod_Reports.GenerateMonthlyReport"
    rowIdx = rowIdx + 35
    AddButton ws, rowIdx, "[REPORT] Stock Card", "mod_Reports.GenerateStockCard"
    rowIdx = rowIdx + 35
    AddButton ws, rowIdx, "[ORDER] Procurement Report", "mod_Procurement.GenerateOrderReport"
    rowIdx = rowIdx + 35
    AddButton ws, rowIdx, "[ABC] Update ABC Classification", "mod_StockEngine.UpdateAllABCClassifications"
    rowIdx = rowIdx + 35
    AddButton ws, rowIdx, "[CMUP] Refresh CMUP", "mod_StockEngine.RefreshAllCMUP"
    rowIdx = rowIdx + 35
    AddButton ws, rowIdx, "[DASHBOARD] Refresh KPIs", "mod_Dashboard.RefreshDashboard"

    ws.Range("B4").Value = "Quick Actions (" & Format(Now, "DD/MM/YYYY HH:MM") & "):"
End Sub

Private Sub AddButton(ByVal ws As Worksheet, ByVal topPos As Long, _
                       ByVal caption As String, ByVal action As String)
    Dim btn As Button
    Set btn = ws.Buttons.Add(50, topPos, 320, 28)
    With btn
        .Caption = caption
        .OnAction = action
        .Font.Bold = True
        .Font.Size = 9
        .Font.Name = "Calibri"
    End With
End Sub
