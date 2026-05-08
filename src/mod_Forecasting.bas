Attribute VB_Name = "mod_Forecasting"
'==============================================================================
' mod_Forecasting.bas  -  ERP LSM v1.0.0
' Purpose: Demand forecasting engine using Moving Average methods
' Author : LSM VBA Core | Public Sector 2026
'
' Methods:
'   - 3-Day Moving Average (short-term sensitivity)
'   - 7-Day Moving Average (weekly cycle)
'   - 14-Day Moving Average (medium-term trend)
'   - Projected Demand (next 30 days)
'   - Stock Coverage Days (how long until stockout)
'
' Data Source: MOUVEMENTS sheet (IN/OUT transactions)
' Constraints: Excel 2010 compatible, zero external dependencies
'==============================================================================

Option Explicit

'================================================================================
' STRUCTURES
'================================================================================

Public Type ForecastResult
    ArticleCode       As String
    ArticleDesc       As String
    CurrentStock      As Double
    AvgDaily3         As Double       ' 3-day moving average
    AvgDaily7         As Double       ' 7-day moving average
    AvgDaily14        As Double       ' 14-day moving average
    ProjectedDemand30 As Double       ' Projected demand for next 30 days
    StockCoverDays    As Double       ' Days of stock remaining
    DaysUntilROP      As Double       ' Days until ROP is reached
    IsCritical        As Boolean      ' True if stock < projected demand
    ForecastAccuracy  As Double       ' MAD (Mean Absolute Deviation)
End Type

Public Type PeriodData
    PeriodStart       As Date
    PeriodEnd         As Date
    TotalIn           As Double
    TotalOut          As Double
    NetChange         As Double
    TransactionCount  As Long
End Type

'================================================================================
' PUBLIC API - Entry points
'================================================================================

'-- Calculate forecast for all articles
Public Sub CalculateAllForecasts(ByRef results() As ForecastResult)
    Dim articles As Variant
    articles = GetAllArticles()
    
    Dim n As Long
    n = UBound(articles)
    
    ReDim results(1 To n)
    
    Dim i As Long
    For i = 1 To n
        results(i) = CalculateArticleForecast(articles(i))
    Next i
End Sub

'-- Calculate forecast for a single article
Public Function CalculateArticleForecast(ByVal articleCode As String) As ForecastResult
    Dim result As ForecastResult
    result.ArticleCode = articleCode
    
    ' Get article info
    result.ArticleDesc = GetForecastArticleField(articleCode, "Description")
    result.CurrentStock = GetForecastArticleField(articleCode, "Stock")
    
    ' Get movement data for this article
    Dim movements As Variant
    movements = GetArticleMovements(articleCode)
    
    If Not IsArray(movements) Then
        ' No movements - forecast based on annual demand
        result.AvgDaily3 = 0
        result.AvgDaily7 = 0
        result.AvgDaily14 = 0
        result.ProjectedDemand30 = 0
        result.StockCoverDays = 999
        result.DaysUntilROP = 999
        CalculateArticleForecast = result
        Exit Function
    End If
    
    ' Calculate moving averages
    result.AvgDaily3 = CalcMovingAverage(movements, 3)
    result.AvgDaily7 = CalcMovingAverage(movements, 7)
    result.AvgDaily14 = CalcMovingAverage(movements, 14)
    
    ' Use 7-day MA as primary (best balance for weekly cycles)
    Dim avgDaily As Double
    avgDaily = result.AvgDaily7
    If avgDaily = 0 Then
        avgDaily = result.AvgDaily3
        If avgDaily = 0 Then avgDaily = result.AvgDaily14
    End If
    
    ' Project 30-day demand
    result.ProjectedDemand30 = avgDaily * 30
    
    ' Stock coverage days
    If avgDaily > 0 Then
        result.StockCoverDays = result.CurrentStock / avgDaily
    Else
        result.StockCoverDays = 999
    End If
    
    ' Days until ROP
    Dim rop As Double
    rop = GetForecastArticleField(articleCode, "ROP")
    If avgDaily > 0 And result.CurrentStock > rop Then
        result.DaysUntilROP = (result.CurrentStock - rop) / avgDaily
    ElseIf result.CurrentStock <= rop Then
        result.DaysUntilROP = 0
    Else
        result.DaysUntilROP = 999
    End If
    
    ' Critical flag
    result.IsCritical = (result.CurrentStock < result.ProjectedDemand30)
    
    ' Forecast accuracy (MAD)
    result.ForecastAccuracy = CalcMAD(movements, 7)
    
    CalculateArticleForecast = result
End Function

'================================================================================
' MOVING AVERAGE CALCULATION
'================================================================================

'-- Calculate N-day moving average from movement data
'   movements array: 2D array (Date, QtyIn, QtyOut, NetChange)
Public Function CalcMovingAverage(ByVal movements As Variant, ByVal periodDays As Long) As Double
    If Not IsArray(movements) Then
        CalcMovingAverage = 0
        Exit Function
    End If
    
    Dim rowCount As Long
    rowCount = UBound(movements, 1)
    
    If rowCount = 0 Then
        CalcMovingAverage = 0
        Exit Function
    End If
    
    ' Sum net consumption over the period
    Dim totalConsumption As Double
    Dim daysWithData As Long
    Dim i As Long
    
    For i = 1 To rowCount
        If i <= periodDays Then
            Dim qtyOut As Double
            qtyOut = Abs(CDbl(movements(i, 3)))  ' QtyOut column
            If qtyOut > 0 Then
                totalConsumption = totalConsumption + qtyOut
                daysWithData = daysWithData + 1
            End If
        End If
    Next i
    
    ' Calculate daily average
    If periodDays > 0 Then
        CalcMovingAverage = totalConsumption / periodDays
    Else
        CalcMovingAverage = 0
    End If
End Function

'================================================================================
' MEAN ABSOLUTE DEVIATION (Forecast Accuracy)
'================================================================================

Public Function CalcMAD(ByVal movements As Variant, ByVal periodDays As Long) As Double
    If Not IsArray(movements) Then
        CalcMAD = 0
        Exit Function
    End If
    
    Dim rowCount As Long
    rowCount = UBound(movements, 1)
    
    If rowCount < periodDays Then
        CalcMAD = 0
        Exit Function
    End If
    
    Dim sumErrors As Double
    Dim errorCount As Long
    Dim i As Long
    
    For i = periodDays + 1 To rowCount
        ' Actual consumption on this day
        Dim actual As Double
        actual = Abs(CDbl(movements(i, 3)))
        
        ' Forecast = average of previous periodDays days
        Dim forecast As Double
        Dim forecastSum As Double
        Dim j As Long
        For j = i - periodDays To i - 1
            forecastSum = forecastSum + Abs(CDbl(movements(j, 3)))
        Next j
        forecast = forecastSum / periodDays
        
        ' Absolute error
        Dim absError As Double
        absError = Abs(actual - forecast)
        
        sumErrors = sumErrors + absError
        errorCount = errorCount + 1
    Next i
    
    If errorCount > 0 Then
        CalcMAD = sumErrors / errorCount
    Else
        CalcMAD = 0
    End If
End Function

'================================================================================
' DATA RETRIEVAL
'================================================================================

'-- Get all article codes from ARTICLES sheet
Public Function GetAllArticles() As Variant
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim articles() As String
    Dim count As Long
    count = 0
    
    Dim i As Long
    For i = 2 To lastRow
        If Len(Trim(ws.Cells(i, "A").Value)) > 0 Then
            count = count + 1
            ReDim Preserve articles(1 To count)
            articles(count) = ws.Cells(i, "A").Value
        End If
    Next i
    
    If count > 0 Then
        GetAllArticles = articles
    Else
        GetAllArticles = Array()
    End If
End Function

'-- Get movement data for a specific article
'   Returns 2D array: (Row, 1=Date, 2=QtyIn, 3=QtyOut, 4=NetChange)
Public Function GetArticleMovements(ByVal articleCode As String) As Variant
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(mod_Config.SHEET_MOUVEMENTS)
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim movements() As Variant
    Dim count As Long
    count = 0
    
    Dim i As Long
    For i = 2 To lastRow
        Dim movArticle As String
        movArticle = Trim(ws.Cells(i, "C").Value)  ' ArticleCode column
        
        If movArticle = articleCode Then
            count = count + 1
            ReDim Preserve movements(1 To 4, 1 To count)
            
            Dim moveDate As Date
            On Error Resume Next
            moveDate = CDate(ws.Cells(i, "A").Value)  ' Date column
            If Err.Number <> 0 Then
                moveDate = DateValue("2026-01-01")
                Err.Clear
            End If
            On Error GoTo 0
            
            Dim qtyIn As Double, qtyOut As Double
            qtyIn = mod_Utilities.SafeVal(ws.Cells(i, "D").Value)   ' QtyIn column
            qtyOut = mod_Utilities.SafeVal(ws.Cells(i, "E").Value)  ' QtyOut column
            
            movements(1, count) = moveDate
            movements(2, count) = qtyIn
            movements(3, count) = qtyOut
            movements(4, count) = qtyIn - qtyOut
        End If
    Next i
    
    If count > 0 Then
        ' Transpose to (row, column) format
        Dim result() As Variant
        ReDim result(1 To count, 1 To 4)
        Dim r As Long, c As Long
        For r = 1 To count
            For c = 1 To 4
                result(r, c) = movements(c, r)
            Next c
        Next r
        
        ' Sort by date (most recent first)
        Call SortMovementsByDate(result, count)
        
        GetArticleMovements = result
    Else
        GetArticleMovements = Array()
    End If
End Function

'-- Sort movements array by date (descending - most recent first)
Private Sub SortMovementsByDate(ByRef arr As Variant, ByVal count As Long)
    Dim i As Long, j As Long
    Dim tempDate As Date, tempIn As Double, tempOut As Double, tempNet As Double
    
    For i = 1 To count - 1
        For j = i + 1 To count
            If arr(i, 1) < arr(j, 1) Then
                ' Swap
                tempDate = arr(i, 1)
                tempIn = arr(i, 2)
                tempOut = arr(i, 3)
                tempNet = arr(i, 4)
                
                arr(i, 1) = arr(j, 1)
                arr(i, 2) = arr(j, 2)
                arr(i, 3) = arr(j, 3)
                arr(i, 4) = arr(j, 4)
                
                arr(j, 1) = tempDate
                arr(j, 2) = tempIn
                arr(j, 3) = tempOut
                arr(j, 4) = tempNet
            End If
        Next j
    Next i
End Sub

'-- Get article field value (wrapper around mod_Utilities)
Public Function GetForecastArticleField(ByVal articleCode As String, ByVal fieldName As String) As Variant
    On Error Resume Next
    Select Case fieldName
        Case "Description": GetForecastArticleField = mod_Utilities.GetArticleField(articleCode, "Description")
        Case "Stock":       GetForecastArticleField = mod_Utilities.SafeVal(mod_Utilities.GetArticleField(articleCode, "Stock"))
        Case "ROP":         GetForecastArticleField = mod_Utilities.SafeVal(mod_Utilities.GetArticleField(articleCode, "ROP"))
        Case "SS":          GetForecastArticleField = mod_Utilities.SafeVal(mod_Utilities.GetArticleField(articleCode, "SS"))
        Case "EOQ":         GetForecastArticleField = mod_Utilities.SafeVal(mod_Utilities.GetArticleField(articleCode, "EOQ"))
        Case "CMUP":        GetForecastArticleField = mod_Utilities.SafeVal(mod_Utilities.GetArticleField(articleCode, "CMUP"))
        Case "ABC":         GetForecastArticleField = mod_Utilities.GetArticleField(articleCode, "ABC")
        Case Else:          GetForecastArticleField = ""
    End Select
    
    On Error GoTo 0
End Function

'================================================================================
' FORECAST SHEET - Generate/Refresh FORECAST sheet
'================================================================================

Public Sub RefreshForecastSheet()
    Debug.Print "[Forecast] Starting forecast refresh..."
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("FORECAST")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "FORECAST"
    End If
    
    ws.Cells.Clear
    
    ' Header row
    Dim headers As Variant
    headers = Array("Code Article", "Description", "Stock Actuel", _
                    "MA 3 Jours", "MA 7 Jours", "MA 14 Jours", _
                    "Demande Projetée (30j)", "Couverture (Jours)", _
                    "Jours avant ROP", "Statut", "MAD")
    
    Dim h As Integer
    For h = 0 To UBound(headers)
        ws.Cells(1, h + 1).Value = headers(h)
        With ws.Cells(1, h + 1).Font
            .Bold = True
            .Size = 10
            .Name = "Calibri"
            .Color = RGB(70, 70, 70)
        End With
        ws.Cells(1, h + 1).Interior.Color = RGB(220, 230, 240)
    Next h
    
    ' Calculate forecasts
    Dim articles As Variant
    articles = GetAllArticles()
    
    If IsArray(articles) Then
        Dim i As Long
        For i = 1 To UBound(articles)
            Dim fc As ForecastResult
            fc = CalculateArticleForecast(articles(i))
            
            Dim row As Long
            row = i + 1
            
            ws.Cells(row, 1).Value = fc.ArticleCode
            ws.Cells(row, 2).Value = fc.ArticleDesc
            ws.Cells(row, 3).Value = Round(fc.CurrentStock, 1)
            ws.Cells(row, 4).Value = Round(fc.AvgDaily3, 2)
            ws.Cells(row, 5).Value = Round(fc.AvgDaily7, 2)
            ws.Cells(row, 6).Value = Round(fc.AvgDaily14, 2)
            ws.Cells(row, 7).Value = Round(fc.ProjectedDemand30, 1)
            ws.Cells(row, 8).Value = Round(fc.StockCoverDays, 1)
            ws.Cells(row, 9).Value = Round(fc.DaysUntilROP, 1)
            ws.Cells(row, 10).Value = GetStatusText(fc)
            ws.Cells(row, 11).Value = Round(fc.ForecastAccuracy, 2)
            
            ' Color code status
            If fc.IsCritical Then
                ws.Cells(row, 10).Interior.Color = RGB(255, 220, 220)
                ws.Cells(row, 10).Font.Color = RGB(204, 0, 0)
            ElseIf fc.DaysUntilROP <= 7 Then
                ws.Cells(row, 10).Interior.Color = RGB(255, 243, 224)
                ws.Cells(row, 10).Font.Color = RGB(255, 140, 0)
            Else
                ws.Cells(row, 10).Interior.Color = RGB(211, 240, 224)
                ws.Cells(row, 10).Font.Color = RGB(40, 100, 40)
            End If
            
            ' Number formatting
            ws.Cells(row, 3).NumberFormat = "#,##0.0"
            ws.Cells(row, 4).NumberFormat = "0.00"
            ws.Cells(row, 5).NumberFormat = "0.00"
            ws.Cells(row, 6).NumberFormat = "0.00"
            ws.Cells(row, 7).NumberFormat = "#,##0.0"
            ws.Cells(row, 8).NumberFormat = "0.0"
            ws.Cells(row, 9).NumberFormat = "0.0"
            ws.Cells(row, 11).NumberFormat = "0.00"
        Next i
    End If
    
    ' Auto-fit columns
    ws.Columns("A:K").AutoFit
    
    ' Set column widths minimum
    ws.Columns("A").ColumnWidth = 12
    ws.Columns("B").ColumnWidth = 28
    ws.Columns("C").ColumnWidth = 14
    ws.Columns("D").ColumnWidth = 12
    ws.Columns("E").ColumnWidth = 12
    ws.Columns("F").ColumnWidth = 12
    ws.Columns("G").ColumnWidth = 18
    ws.Columns("H").ColumnWidth = 16
    ws.Columns("I").ColumnWidth = 16
    ws.Columns("J").ColumnWidth = 14
    ws.Columns("K").ColumnWidth = 10
    
    ' Protect sheet
    ws.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    Debug.Print "[Forecast] Forecast sheet refreshed"
End Sub

'-- Get status text for forecast result
Private Function GetStatusText(ByRef fc As ForecastResult) As String
    If fc.CurrentStock <= 0 Then
        GetStatusText = "RUPTURE"
    ElseIf fc.IsCritical Then
        GetStatusText = "CRITIQUE"
    ElseIf fc.DaysUntilROP <= 7 Then
        GetStatusText = "ALERTE"
    ElseIf fc.StockCoverDays < 30 Then
        GetStatusText = "ATTENTION"
    Else
        GetStatusText = "NORMAL"
    End If
End Function

'================================================================================
' DASHBOARD INTEGRATION - KPI values for mod_Dashboard
'================================================================================

'-- Get total critical articles (demand > stock within 30 days)
Public Function GetCriticalForecastCount() As Long
    Dim articles As Variant
    articles = GetAllArticles()
    
    Dim count As Long
    count = 0
    
    If IsArray(articles) Then
        Dim i As Long
        For i = 1 To UBound(articles)
            Dim fc As ForecastResult
            fc = CalculateArticleForecast(articles(i))
            If fc.IsCritical Then count = count + 1
        Next i
    End If
    
    GetCriticalForecastCount = count
End Function

'-- Get average forecast accuracy (MAD) across all articles
Public Function GetAverageForecastAccuracy() As Double
    Dim articles As Variant
    articles = GetAllArticles()
    
    Dim sumMAD As Double
    Dim articleCount As Long
    
    If IsArray(articles) Then
        Dim i As Long
        For i = 1 To UBound(articles)
            Dim fc As ForecastResult
            fc = CalculateArticleForecast(articles(i))
            If fc.ForecastAccuracy > 0 Then
                sumMAD = sumMAD + fc.ForecastAccuracy
                articleCount = articleCount + 1
            End If
        Next i
    End If
    
    If articleCount > 0 Then
        GetAverageForecastAccuracy = sumMAD / articleCount
    Else
        GetAverageForecastAccuracy = 0
    End If
End Function

'-- Get article with lowest stock coverage (most urgent)
Public Function GetMostUrgentArticle() As String
    Dim articles As Variant
    articles = GetAllArticles()
    
    Dim minCoverage As Double
    minCoverage = 999
    
    Dim urgentCode As String
    urgentCode = ""
    
    If IsArray(articles) Then
        Dim i As Long
        For i = 1 To UBound(articles)
            Dim fc As ForecastResult
            fc = CalculateArticleForecast(articles(i))
            If fc.StockCoverDays < minCoverage And fc.StockCoverDays > 0 Then
                minCoverage = fc.StockCoverDays
                urgentCode = fc.ArticleCode & " (" & fc.ArticleDesc & ")"
            End If
        Next i
    End If
    
    GetMostUrgentArticle = urgentCode
End Function

'================================================================================
' EXPORT - Forecast report as PDF
'================================================================================

Public Sub ExportForecastToPDF()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("FORECAST")
    
    If ws Is Nothing Then
        Call RefreshForecastSheet
        Set ws = ThisWorkbook.Sheets("FORECAST")
    End If
    
    ' Unprotect temporarily for export
    ws.Unprotect Password:=mod_Config.MASTER_PWD
    
    Dim exportPath As String
    exportPath = ThisWorkbook.Path & "\Forecast_Report_" & Format(Date, "YYYY-MM-DD") & ".pdf"
    
    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=exportPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True
    
    ' Re-protect
    ws.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    
    Debug.Print "[Forecast] PDF exported to: " & exportPath
End Sub

'================================================================================
' AUDIT - Log forecast calculation
'================================================================================

Public Sub LogForecastCalculation()
    If Not mod_AuditTrail.AuditLogInitialized Then Exit Sub
    
    Dim articleCount As Long
    articleCount = UBound(GetAllArticles())
    
    mod_AuditTrail.LogAction "FORECAST", _
        "Forecast calculated for " & articleCount & " articles", _
        "mod_Forecasting", _
        "ForecastEngine"
End Sub

'==============================================================================
' END -- mod_Forecasting.bas
'==============================================================================
