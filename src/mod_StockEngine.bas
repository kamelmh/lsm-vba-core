Attribute VB_Name = "mod_StockEngine"
'=======================================================================================
' MODULE: mod_StockEngine.bas
' PROJECT: ERP Acad" & Chr(233) & "mie v13
' DESCRIPTION: Handles ROP, EOQ and CMUP calculations for inventory management.
'=======================================================================================
Option Explicit

' ================================================================================
' CONSTANTS � Synchronized with Unité de traitement VBA GROUND_TRUTH
' ================================================================================
Private Const ORDER_COST_S  As Double = 500    ' DZD � full order cycle cost
Private Const HOLDING_RATE  As Double = 0.2    ' 20% of unit price per year
Private Const LEAD_TIME_DEFAULT As Integer = 2 ' Default delivery days

' Article-specific safety stocks � mirrors Unité de traitement VBA GROUND_TRUTH
Public Function GetSafetyStock(ByVal sku As String) As Double
    Select Case UCase(Trim(sku))
        Case "ART-001": GetSafetyStock = 200
        Case "ART-005": GetSafetyStock = 10
        Case "ART-002": GetSafetyStock = 50
        Case "ART-003": GetSafetyStock = 20
        Case Else:      GetSafetyStock = 50
    End Select
End Function

' ================================================================================
' FUNCTION: ComputeEOQ
' Formula: Q* = SQRT(2 x D x S / (P x t))
' ================================================================================
Public Function ComputeEOQ(ByVal AnnualDemand As Double, _
                            ByVal unitPrice As Double) As Double
    If unitPrice <= 0 Or AnnualDemand <= 0 Then
        ComputeEOQ = 0
        Exit Function
    End If

    Dim holdingCostH As Double
    holdingCostH = unitPrice * HOLDING_RATE

    ComputeEOQ = Sqr((2 * AnnualDemand * ORDER_COST_S) / holdingCostH)
End Function

' ================================================================================
' FUNCTION: ComputeROP
' Formula: ROP = (avg_daily_demand x lead_time) + safety_stock
' ================================================================================
Public Function ComputeROP(ByVal AvgDailyDemand As Double, _
                            ByVal sku As String, _
                            Optional ByVal LeadTimeDays As Integer = LEAD_TIME_DEFAULT) As Double
    ComputeROP = (AvgDailyDemand * LeadTimeDays) + GetSafetyStock(sku)
End Function

' ================================================================================
' SUB: ValidateStockLevel
' Fires a UI alert if current stock breaches ROP.
' ================================================================================
Public Sub ValidateStockLevel(ByVal sku As String, _
                               ByVal CurrentStock As Double, _
                               ByVal AnnualDemand As Double, _
                               ByVal unitPrice As Double)
    If AnnualDemand <= 0 Then Exit Sub

    Dim avgDaily As Double: avgDaily = AnnualDemand / mod_Config.WORKING_DAYS_PER_YEAR
    Dim rop As Double: rop = ComputeROP(avgDaily, sku)
    Dim ss As Double: ss = GetSafetyStock(sku)

    If CurrentStock <= rop Then
        Dim eoq As Double: eoq = ComputeEOQ(AnnualDemand, unitPrice)
        Dim alertLevel As String: alertLevel = IIf(CurrentStock <= ss, "RUPTURE IMMINENTE", "SEUIL D'ALERTE ATTEINT")

        MsgBox alertLevel & vbCrLf & vbCrLf & _
               "Article  : " & sku & vbCrLf & _
               "Stock    : " & CurrentStock & " unites" & vbCrLf & _
               "ROP      : " & Round(rop, 1) & " unites" & vbCrLf & _
               "SS       : " & ss & " unites" & vbCrLf & _
               "EOQ (Q*) : " & Round(eoq, 0) & " unites a commander", _
               vbExclamation, mod_Config.SYS_TITLE
    End If
End Sub

' ================================================================================
' FUNCTION: GetArticleStock
' Returns current stock quantity for an article (reads from ARTICLES column C)
' ================================================================================
Public Function GetArticleStock(ByVal sku As String) As Double
    Dim wsArt As Worksheet
    Dim foundRow As Variant
    
    On Error Resume Next
    Set wsArt = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    On Error GoTo 0
    
    If wsArt Is Nothing Then
        GetArticleStock = 0
        Exit Function
    End If
    
    foundRow = Application.Match(sku, wsArt.Range("A:A"), 0)
    
    If IsError(foundRow) Then
        GetArticleStock = 0
        Exit Function
    End If
    
    GetArticleStock = mod_Utilities.SafeVal(wsArt.Cells(foundRow, 3).Value)
End Function

' ================================================================================
' SUB: UpdateArticleStockBalance
' Directly updates the stock quantity in the ARTICLES sheet based on movements.
' ================================================================================
Public Sub UpdateArticleStockBalance(ByVal artCode As String, ByVal mvtSign As String, ByVal qty As Long)
    Dim wsArt As Worksheet
    Dim foundRow As Variant
    
    On Error Resume Next
    Set wsArt = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    On Error GoTo 0
    
    If wsArt Is Nothing Then Exit Sub
    
    foundRow = Application.Match(artCode, wsArt.Range("A:A"), 0)
    
    If Not IsError(foundRow) Then
        wsArt.Unprotect Password:=mod_Config.MASTER_PWD
        
        Dim currentQty As Double: currentQty = Val(wsArt.Cells(foundRow, 3).Value) ' Column C: Stock
        
        If mvtSign = "IN" Then
            wsArt.Cells(foundRow, 3).Value = currentQty + qty
        Else
            wsArt.Cells(foundRow, 3).Value = currentQty - qty
        End If
        
        wsArt.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    End If
End Sub

' ================================================================================
' FUNCTION: GetAnnualDemandFromHistory
' Aggregates annual demand from MOUVEMENTS sheet for a given SKU.
' ================================================================================
Public Function GetAnnualDemandFromHistory(ByVal sku As String) As Double
    On Error Resume Next
    Dim wsMouv As Worksheet: Set wsMouv = ThisWorkbook.Sheets(mod_Config.SHEET_MOUVEMENTS)
    If wsMouv Is Nothing Then GetAnnualDemandFromHistory = 0: Exit Function
    Dim currentYear As Integer: currentYear = Year(Date)
    
    wsMouv.Unprotect Password:=mod_Config.MASTER_PWD
    GetAnnualDemandFromHistory = WorksheetFunction.SumIfs( _
        wsMouv.Range("E:E"), _
        wsMouv.Range("B:B"), sku, _
        wsMouv.Range("D:D"), "OUT", _
        wsMouv.Range("A:A"), ">=" & DateSerial(currentYear, 1, 1))
    wsMouv.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    On Error GoTo 0
End Function

' ================================================================================
' FUNCTION: CalculateCMUP
' Formula: CMUP = Total IN Value / Total IN Quantity
' ================================================================================
Public Function CalculateCMUP(ByVal sku As String) As Double
    On Error Resume Next
    Dim wsMouv As Worksheet: Set wsMouv = ThisWorkbook.Sheets(mod_Config.SHEET_MOUVEMENTS)
    Dim wsArt As Worksheet: Set wsArt = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    If wsMouv Is Nothing Or wsArt Is Nothing Then CalculateCMUP = 0: Exit Function

    Dim totalInQty As Double, TotalINValue As Double
    wsMouv.Unprotect Password:=mod_Config.MASTER_PWD
    totalInQty = WorksheetFunction.SumIfs(wsMouv.Range("E:E"), wsMouv.Range("B:B"), sku, wsMouv.Range("D:D"), "IN")
    TotalINValue = WorksheetFunction.SumIfs(wsMouv.Range("G:G"), wsMouv.Range("B:B"), sku, wsMouv.Range("D:D"), "IN")

    ' CMUP = Total IN Value / Total IN Quantity (standard weighted average cost)
    If totalInQty > 0 Then
        CalculateCMUP = TotalINValue / totalInQty
    Else
        CalculateCMUP = 0
    End If
    On Error GoTo 0
End Function

' ================================================================================
' SUB: RefreshAllCMUP
' Recalculates CMUP for all articles in ARTICLES sheet
' ================================================================================
Public Sub RefreshAllCMUP()
    Dim wsArt As Worksheet: Set wsArt = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    Dim lastRow As Long: lastRow = wsArt.Cells(wsArt.Rows.count, 1).End(xlUp).Row
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim i As Long, cmup As Double
    For i = 2 To lastRow
        Dim sku As String: sku = Trim(wsArt.Cells(i, 1).Value)
        If sku <> "" Then
            cmup = CalculateCMUP(sku)
            If cmup > 0 Then wsArt.Cells(i, 12).Value = cmup
        End If
    Next i
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "CMUP (Prix Moyen) mis " & Chr(233) & " jour.", vbInformation, mod_Config.SYS_TITLE
End Sub

' ================================================================================
' SUB: UpdateAllABCClassifications
' Calculates ABC classification based on annual consumption value.
' A: Top 80%, B: 15%, C: 5%
' ================================================================================
Public Sub UpdateAllABCClassifications(Optional ByVal silent As Boolean = False)
    Dim wsArt As Worksheet: Set wsArt = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    Dim lastRow As Long: lastRow = wsArt.Cells(wsArt.Rows.count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    Dim i As Long
    Dim totalValue As Double: totalValue = 0
    Dim articleValues() As Double: ReDim articleValues(2 To lastRow)
    Dim articleCodes() As String: ReDim articleCodes(2 To lastRow)

    wsArt.Unprotect Password:=mod_Config.MASTER_PWD

    ' 1. Calculate total value for each article
    For i = 2 To lastRow
        Dim sku As String: sku = Trim(wsArt.Cells(i, 1).Value)
        If sku <> "" Then
            Dim AnnualDemand As Double: AnnualDemand = GetAnnualDemandFromHistory(sku)
            Dim pu As Double: pu = Val(wsArt.Cells(i, 8).Value) ' Column H: PU
            articleValues(i) = AnnualDemand * pu
            articleCodes(i) = sku
            totalValue = totalValue + articleValues(i)
        End If
    Next i

    If totalValue = 0 Then
        wsArt.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
        Exit Sub
    End If

    ' 2. Sort articles by value (Simple Bubble Sort for small lists)
    Dim j As Long, tempVal As Double, tempCode As String
    For i = 2 To lastRow - 1
        For j = i + 1 To lastRow
            If articleValues(i) < articleValues(j) Then
                tempVal = articleValues(i): articleValues(i) = articleValues(j): articleValues(j) = tempVal
                tempCode = articleCodes(i): articleCodes(i) = articleCodes(j): articleCodes(j) = tempCode
            End If
        Next j
    Next i

    ' 3. Assign classes
    Dim cumulativeValue As Double: cumulativeValue = 0
    For i = 2 To lastRow
        cumulativeValue = cumulativeValue + articleValues(i)
        Dim ratio As Double: ratio = cumulativeValue / totalValue
        Dim abcClass As String
        
        If ratio <= 0.8 Then
            abcClass = "A"
        ElseIf ratio <= 0.95 Then
            abcClass = "B"
        Else
            abcClass = "C"
        End If

        ' Update ARTICLES sheet (Column F = 6)
        Dim foundRow As Variant
        foundRow = Application.Match(articleCodes(i), wsArt.Range("A:A"), 0)
        If Not IsError(foundRow) Then
            wsArt.Cells(foundRow, 6).Value = abcClass
        End If
    Next i

    wsArt.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    If Not silent Then
        MsgBox "Classifications ABC mises a jour.", vbInformation, mod_Config.SYS_TITLE
    End If
End Sub

