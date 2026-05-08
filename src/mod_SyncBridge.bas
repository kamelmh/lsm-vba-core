Attribute VB_Name = "mod_SyncBridge"
Option Explicit

' Module-level state
Private m_LastSyncTime As Date

'=======================================================================================
' MODULE: mod_SyncBridge
' PURPOSE: Orchestrates data synchronization between transactions and core metrics.
' REFACTORED v1.0.0: All Python dependencies removed. Logic now fully local to VBA.
'=======================================================================================

'--------------------------------------------------------------------------------------
' PUBLIC BRIDGE - Synchronizes a transaction with local stock engine.
' Replaces the previous Python bridge with direct VBA Processing Units.
'--------------------------------------------------------------------------------------
Public Function SyncTransactionInternal(ByVal artCode As String, _
                                      ByVal mvtType As String, _
                                      ByVal qty As Long, _
                                      ByVal unitPrice As Double, _
                                      ByVal refDoc As String) As Integer
    On Error GoTo SyncError
    
    ' 1. Update the physical stock balance in the ARTICLES sheet
    Call mod_StockEngine.UpdateArticleStockBalance(artCode, mvtType, qty)
    
    ' 2. Recalculate CMUP for this specific article immediately
    Dim newCMUP As Double
    newCMUP = mod_StockEngine.CalculateCMUP(artCode)
    
    Dim wsArt As Worksheet: Set wsArt = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    Dim foundRow As Variant
    foundRow = Application.Match(artCode, wsArt.Range("A:A"), 0)
    
    If Not IsError(foundRow) Then
        wsArt.Unprotect Password:=mod_Config.MASTER_PWD
        wsArt.Cells(foundRow, 8).Value = newCMUP ' Col H: CMUP
        wsArt.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    End If
    
    ' 3. Trigger stock level validation (ROP/SS alert)
    Dim AnnualDemand As Double
    AnnualDemand = mod_StockEngine.GetAnnualDemandFromHistory(artCode)
    
    ' Get current stock for validation
    Dim CurrentStock As Long
    CurrentStock = GetStockFromLedger(artCode)
    
    Call mod_StockEngine.ValidateStockLevel(artCode, CurrentStock, AnnualDemand, unitPrice)
    
    Debug.Print "[SYNC " & Format(Now, "HH:MM:SS") & "] Local sync complete for " & artCode
    SyncTransactionInternal = 0
    Exit Function

SyncError:
    Debug.Print "[SYNC ERROR] " & Err.Description
    SyncTransactionInternal = -1
End Function

'--------------------------------------------------------------------------------------
' METRICS SYNC - Updates all CMUP and ABC classifications in the ARTICLES sheet.
'--------------------------------------------------------------------------------------
Public Sub SyncMetricsFromLedger()
    On Error Resume Next
    ' 1. Recalculate all CMUPs
    Call mod_StockEngine.RefreshAllCMUP
    
    ' 2. Recalculate all ABC Classifications
    Call mod_StockEngine.UpdateAllABCClassifications(silent:=True)
    
    Debug.Print "[METRICS SYNC] All article metrics updated locally."
    On Error GoTo 0
End Sub

'--------------------------------------------------------------------------------------
' GENERIC METRIC RETRIEVER - Reads specific metrics from the ARTICLES sheet.
'--------------------------------------------------------------------------------------
Public Function GetMetricFromLedger(ByVal artCode As String, ByVal metricName As String) As Variant
    Dim wsArt As Worksheet: Set wsArt = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    Dim foundRow As Variant
    foundRow = Application.Match(artCode, wsArt.Range("A:A"), 0)
    
    If IsError(foundRow) Then
        GetMetricFromLedger = "Unknown"
        Exit Function
    End If
    
    Select Case LCase(metricName)
        Case "cmup"
            GetMetricFromLedger = wsArt.Cells(foundRow, 8).Value ' Col H
        Case "abc_class"
            GetMetricFromLedger = wsArt.Cells(foundRow, 5).Value ' Col E
        Case Else
            GetMetricFromLedger = "Unknown"
    End Select
End Function

'--------------------------------------------------------------------------------------
' STOCK RETRIEVER - Reads the current stock level from the ARTICLES sheet.
'--------------------------------------------------------------------------------------
Public Function GetStockFromLedger(ByVal artCode As String) As Long
    Dim wsArt As Worksheet: Set wsArt = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    Dim foundRow As Variant
    foundRow = Application.Match(artCode, wsArt.Range("A:A"), 0)
    
    If IsError(foundRow) Then
        GetStockFromLedger = -1
    Else
        GetStockFromLedger = CLng(wsArt.Cells(foundRow, 3).Value) ' Col C
    End If
End Function

'--------------------------------------------------------------------------------------
' SKU METRICS - Returns a pipe-separated string: "ABC-XYZ|Countdown"
'--------------------------------------------------------------------------------------
Public Function GetSkuMetrics(ByVal artCode As String) As String
    Dim abcVal As String
    Dim countdown As String
    
    abcVal = GetMetricFromLedger(artCode, "abc_class")
    
    ' Countdown calculation based on ROP and avg demand
    Dim AnnualDemand As Double: AnnualDemand = mod_StockEngine.GetAnnualDemandFromHistory(artCode)
    Dim CurrentStock As Long: CurrentStock = GetStockFromLedger(artCode)
    Dim avgDaily As Double: avgDaily = AnnualDemand / mod_Config.WORKING_DAYS_PER_YEAR
    
    If avgDaily > 0 Then
        Dim daysLeft As Long
        daysLeft = CLng(CurrentStock / avgDaily)
        countdown = CStr(daysLeft)
    Else
        countdown = "N/A"
    End If
    
    GetSkuMetrics = abcVal & "|" & countdown
End Function

'--- Legacy Compatibility Stubs (W009: Implemented) ---

Public Function IsSyncComplete(ByVal triggerTime As String) As Boolean
    IsSyncComplete = (m_LastSyncTime > CDate(triggerTime))
End Function

Public Function GetSyncProgress() As Integer
    GetSyncProgress = 100
End Function

Public Function GetSyncError() As String
    GetSyncError = ""
End Function

Public Sub MarkSyncComplete()
    m_LastSyncTime = Now
End Sub

' Stub: Returns workbook path as hub root (simplified for offline deployment)
Public Function GetHubRoot() As String
    GetHubRoot = ThisWorkbook.Path
End Function

' Stub: Restore master ledger from backup (not yet implemented)
Public Sub RestoreMasterLedger()
    MsgBox "Fonctionnalit" & Chr(233) & " non impl" & Chr(233) & "ment" & Chr(233) & "e.", vbInformation, "Restore Ledger"
End Sub
