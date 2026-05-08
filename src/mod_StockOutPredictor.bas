Attribute VB_Name = "mod_StockOutPredictor"
Option Explicit

Private Const SHEET_STOCKOUT As String = "PREVISION_RUPTURE"

Public Sub RunStockOutPrediction()
    Dim wsArt As Worksheet, wsMouv As Worksheet, wsOut As Worksheet
    Dim lastArtRow As Long, lastMouvRow As Long
    Dim i As Long, j As Long, outRow As Long
    Dim artCode As String, designation As String
    Dim currentStock As Double, pu As Double
    Dim totalOut As Double, totalDays As Double, avgDaily As Double
    Dim stockCoverDays As Double, depletionDate As Date
    Dim status As String

    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Set wsArt = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    Set wsMouv = ThisWorkbook.Sheets(mod_Config.SHEET_MOUVEMENTS)
    lastArtRow = wsArt.Cells(wsArt.Rows.Count, 1).End(xlUp).Row
    lastMouvRow = wsMouv.Cells(wsMouv.Rows.Count, 1).End(xlUp).Row

    Set wsOut = GetOrCreateSheet(SHEET_STOCKOUT)
    wsOut.Unprotect Password:=mod_Config.MASTER_PWD
    wsOut.Cells.Clear

    With wsOut
        .Range("A1:H1").Merge
        .Cells(1, 1).Value = "PREVISION DES RUPTURES DE STOCK"
        .Cells(1, 1).Font.Bold = True
        .Cells(1, 1).Font.Size = 14
        .Cells(2, 1).Value = "Date: " & Format(Now, "DD/MM/YYYY HH:MM")
        .Cells(2, 1).Font.Italic = True

        .Cells(4, 1).Value = "CODE_ARTICLE"
        .Cells(4, 2).Value = "DESIGNATION"
        .Cells(4, 3).Value = "STOCK_ACTUEL"
        .Cells(4, 4).Value = "CONSOMMATION_MOYENNE/J"
        .Cells(4, 5).Value = "JOURS_COUVERTURE"
        .Cells(4, 6).Value = "DATE_PREVISION_RUPTURE"
        .Cells(4, 7).Value = "STATUT"
        .Cells(4, 8).Value = "VALEUR_STOCK (DZD)"
        .Range("A4:H4").Font.Bold = True
        .Range("A4:H4").Interior.Color = RGB(0, 70, 127)
        .Range("A4:H4").Font.Color = RGB(255, 255, 255)
    End With

    outRow = 5

    For i = 3 To lastArtRow
        artCode = Trim(wsArt.Cells(i, 1).Value)
        If artCode = "" Then GoTo NextArt

        designation = Trim(wsArt.Cells(i, 2).Value)
        currentStock = mod_Utilities.SafeVal(wsArt.Cells(i, 7).Value)
        pu = mod_Utilities.SafeVal(wsArt.Cells(i, 8).Value)

        totalOut = 0: totalDays = 0
        For j = 3 To lastMouvRow
            If Trim(wsMouv.Cells(j, 2).Value) = artCode And Trim(wsMouv.Cells(j, 4).Value) = "OUT" Then
                totalOut = totalOut + mod_Utilities.SafeVal(wsMouv.Cells(j, 5).Value)
                totalDays = totalDays + 1
            End If
        Next j

        avgDaily = 0
        If totalDays > 0 Then avgDaily = totalOut / totalDays

        stockCoverDays = 999
        If avgDaily > 0 Then stockCoverDays = currentStock / avgDaily

        depletionDate = Date + stockCoverDays

        If currentStock <= 0 Then
            status = "EN_RUPTURE"
        ElseIf stockCoverDays <= 7 Then
            status = "CRITIQUE"
        ElseIf stockCoverDays <= 30 Then
            status = "ALERTE"
        ElseIf stockCoverDays <= 60 Then
            status = "ATTENTION"
        Else
            status = "NORMAL"
        End If

        With wsOut
            .Cells(outRow, 1).Value = artCode
            .Cells(outRow, 2).Value = designation
            .Cells(outRow, 3).Value = currentStock
            .Cells(outRow, 4).Value = Round(avgDaily, 2)
            .Cells(outRow, 5).Value = Round(stockCoverDays, 1)
            .Cells(outRow, 6).Value = depletionDate
            .Cells(outRow, 6).NumberFormat = "DD/MM/YYYY"
            .Cells(outRow, 7).Value = status
            .Cells(outRow, 8).Value = currentStock * pu
            .Cells(outRow, 8).NumberFormat = "#,##0.00"

            Select Case status
                Case "EN_RUPTURE"
                    .Range(.Cells(outRow, 3), .Cells(outRow, 8)).Interior.Color = RGB(255, 180, 180)
                Case "CRITIQUE"
                    .Range(.Cells(outRow, 3), .Cells(outRow, 8)).Interior.Color = RGB(255, 220, 220)
                Case "ALERTE"
                    .Range(.Cells(outRow, 3), .Cells(outRow, 8)).Interior.Color = RGB(255, 243, 224)
                Case "ATTENTION"
                    .Range(.Cells(outRow, 3), .Cells(outRow, 8)).Interior.Color = RGB(255, 255, 210)
            End Select
        End With

        outRow = outRow + 1
NextArt:
    Next i

    wsOut.Columns("A:H").AutoFit
    wsOut.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Prevision des ruptures terminee. " & (outRow - 5) & " articles analyses.", vbInformation, "LSM v1.0.0"
    Exit Sub

ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Erreur prevision: " & Err.Description, vbCritical
End Sub

Public Sub ExportStockOutPDF()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_STOCKOUT)
    On Error GoTo 0
    If ws Is Nothing Then MsgBox "Execute d'abord RunStockOutPrediction.", vbExclamation: Exit Sub

    Dim pdfPath As String
    pdfPath = ThisWorkbook.Path & "\Prevision_Rupture_" & Format(Date, "YYYY-MM-DD") & ".pdf"

    ws.Unprotect Password:=mod_Config.MASTER_PWD
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, Quality:=xlQualityStandard
    ws.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    MsgBox "Rapport exporte: " & pdfPath, vbInformation
End Sub

Private Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = sheetName
    End If
    Set GetOrCreateSheet = ws
End Function
