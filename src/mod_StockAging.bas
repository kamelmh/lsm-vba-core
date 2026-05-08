Attribute VB_Name = "mod_StockAging"
Option Explicit

Private Const SHEET_AGING As String = "ANALYSE_VETUSTE"

Public Sub RunStockAgingReport()
    Dim wsArt As Worksheet, wsMouv As Worksheet, wsOut As Worksheet
    Dim lastArtRow As Long, lastMouvRow As Long
    Dim i As Long, j As Long, outRow As Long
    Dim artCode As String, designation As String
    Dim currentStock As Double, pu As Double
    Dim lastDate As Date, daysSinceLast As Double
    Dim category As String, totalValue As Double
    Dim countActive As Long, countSlow As Long, countDead As Long

    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Set wsArt = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    Set wsMouv = ThisWorkbook.Sheets(mod_Config.SHEET_MOUVEMENTS)
    lastArtRow = wsArt.Cells(wsArt.Rows.Count, 1).End(xlUp).Row
    lastMouvRow = wsMouv.Cells(wsMouv.Rows.Count, 1).End(xlUp).Row

    Set wsOut = GetOrCreateSheet(SHEET_AGING)
    wsOut.Unprotect Password:=mod_Config.MASTER_PWD
    wsOut.Cells.Clear

    With wsOut
        .Range("A1:G1").Merge
        .Cells(1, 1).Value = "ANALYSE DE VETUSTE DES STOCKS"
        .Cells(1, 1).Font.Bold = True
        .Cells(1, 1).Font.Size = 14
        .Cells(2, 1).Value = "Date: " & Format(Now, "DD/MM/YYYY HH:MM")
        .Cells(2, 1).Font.Italic = True

        .Cells(4, 1).Value = "CODE_ARTICLE"
        .Cells(4, 2).Value = "DESIGNATION"
        .Cells(4, 3).Value = "STOCK_ACTUEL"
        .Cells(4, 4).Value = "DERNIER_MOUVEMENT"
        .Cells(4, 5).Value = "JOURS_DERNIER_MVT"
        .Cells(4, 6).Value = "CATEGORIE"
        .Cells(4, 7).Value = "VALEUR (DZD)"
        .Range("A4:G4").Font.Bold = True
        .Range("A4:G4").Interior.Color = RGB(0, 70, 127)
        .Range("A4:G4").Font.Color = RGB(255, 255, 255)
    End With

    countActive = 0: countSlow = 0: countDead = 0
    outRow = 5

    For i = 3 To lastArtRow
        artCode = Trim(wsArt.Cells(i, 1).Value)
        If artCode = "" Then GoTo NextArt

        designation = Trim(wsArt.Cells(i, 2).Value)
        currentStock = mod_Utilities.SafeVal(wsArt.Cells(i, 7).Value)
        pu = mod_Utilities.SafeVal(wsArt.Cells(i, 8).Value)

        lastDate = DateSerial(2026, 1, 1)
        For j = 3 To lastMouvRow
            If Trim(wsMouv.Cells(j, 2).Value) = artCode Then
                Dim mvtDate As Date
                On Error Resume Next
                mvtDate = CDate(wsMouv.Cells(j, 1).Value)
                If Err.Number = 0 Then
                    If mvtDate > lastDate Then lastDate = mvtDate
                End If
                On Error GoTo ErrorHandler
            End If
        Next j

        daysSinceLast = DateDiff("d", lastDate, Date)
        If daysSinceLast < 0 Then daysSinceLast = 0

        If currentStock <= 0 Then
            category = "SANS_STOCK"
            countDead = countDead + 1
        ElseIf daysSinceLast >= 90 Then
            category = "MORT (90+ jours)"
            countDead = countDead + 1
        ElseIf daysSinceLast >= 60 Then
            category = "TRES LENT (60-89j)"
            countSlow = countSlow + 1
        ElseIf daysSinceLast >= 30 Then
            category = "LENT (30-59j)"
            countSlow = countSlow + 1
        Else
            category = "ACTIF (< 30j)"
            countActive = countActive + 1
        End If

        totalValue = currentStock * pu

        With wsOut
            .Cells(outRow, 1).Value = artCode
            .Cells(outRow, 2).Value = designation
            .Cells(outRow, 3).Value = currentStock
            .Cells(outRow, 4).Value = lastDate
            .Cells(outRow, 4).NumberFormat = "DD/MM/YYYY"
            .Cells(outRow, 5).Value = daysSinceLast
            .Cells(outRow, 6).Value = category
            .Cells(outRow, 7).Value = totalValue
            .Cells(outRow, 7).NumberFormat = "#,##0.00"

            Select Case True
                Case category Like "MORT*"
                    .Range(.Cells(outRow, 3), .Cells(outRow, 7)).Interior.Color = RGB(255, 200, 200)
                Case category Like "*LENT*"
                    .Range(.Cells(outRow, 3), .Cells(outRow, 7)).Interior.Color = RGB(255, 243, 224)
                Case category = "SANS_STOCK"
                    .Range(.Cells(outRow, 3), .Cells(outRow, 7)).Interior.Color = RGB(220, 220, 220)
            End Select
        End With

        outRow = outRow + 1
NextArt:
    Next i

    Dim sRow As Long
    sRow = outRow + 1
    With wsOut
        .Cells(sRow, 1).Value = "RESUME"
        .Cells(sRow, 1).Font.Bold = True
        .Cells(sRow + 1, 1).Value = "Articles actifs:"
        .Cells(sRow + 1, 2).Value = countActive
        .Cells(sRow + 2, 1).Value = "Articles a rotation lente:"
        .Cells(sRow + 2, 2).Value = countSlow
        .Cells(sRow + 3, 1).Value = "Articles morts/sans stock:"
        .Cells(sRow + 3, 2).Value = countDead
        .Columns("A:G").AutoFit
    End With

    wsOut.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Analyse de vetuste terminee." & vbCrLf & _
           countActive & " actifs, " & countSlow & " lents, " & countDead & " morts/sans stock.", _
           vbInformation, "LSM v1.0.0"
    Exit Sub

ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Erreur analyse vetuste: " & Err.Description, vbCritical
End Sub

Public Sub ExportAgingPDF()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_AGING)
    On Error GoTo 0
    If ws Is Nothing Then MsgBox "Execute d'abord RunStockAgingReport.", vbExclamation: Exit Sub

    Dim pdfPath As String
    pdfPath = ThisWorkbook.Path & "\Analyse_Vetuste_" & Format(Date, "YYYY-MM-DD") & ".pdf"

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
