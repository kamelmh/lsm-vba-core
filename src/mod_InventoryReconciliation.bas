Attribute VB_Name = "mod_InventoryReconciliation"
Option Explicit

Private Const SHEET_INVENTAIRE As String = "INVENTAIRE"
Private Const SHEET_OUTPUT As String = "RECONCILIATION"

Public Sub RunInventoryReconciliation()
    Dim wsInv As Worksheet, wsArt As Worksheet, wsOut As Worksheet
    Dim lastInvRow As Long, lastArtRow As Long
    Dim i As Long, outRow As Long
    Dim artCode As String, physCount As Double, sysStock As Double
    Dim variance As Double, variancePct As Double
    Dim found As Variant
    Dim totalItems As Long, totalMatch As Long, totalDiff As Long

    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Set wsArt = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    lastArtRow = wsArt.Cells(wsArt.Rows.Count, 1).End(xlUp).Row

    Set wsInv = GetOrCreateSheet(SHEET_INVENTAIRE)
    wsInv.Unprotect Password:=mod_Config.MASTER_PWD

    If wsInv.Cells(1, 1).Value <> "CODE_ARTICLE" Then
        InitializeInventorySheet wsInv
    End If

    lastInvRow = wsInv.Cells(wsInv.Rows.Count, 1).End(xlUp).Row

    Set wsOut = GetOrCreateSheet(SHEET_OUTPUT)
    wsOut.Unprotect Password:=mod_Config.MASTER_PWD
    wsOut.Cells.Clear

    With wsOut
        .Range("A1:G1").Merge
        .Cells(1, 1).Value = "RAPPORT DE RECONCILIATION INVENTAIRE"
        .Cells(1, 1).Font.Bold = True
        .Cells(1, 1).Font.Size = 14
        .Cells(2, 1).Value = "Date: " & Format(Now, "DD/MM/YYYY HH:MM")
        .Cells(2, 1).Font.Italic = True

        .Cells(4, 1).Value = "CODE_ARTICLE"
        .Cells(4, 2).Value = "DESIGNATION"
        .Cells(4, 3).Value = "INVENTAIRE_PHYSIQUE"
        .Cells(4, 4).Value = "STOCK_SYSTEME"
        .Cells(4, 5).Value = "ECART"
        .Cells(4, 6).Value = "ECART_%"
        .Cells(4, 7).Value = "STATUT"
        .Range("A4:G4").Font.Bold = True
        .Range("A4:G4").Interior.Color = RGB(0, 70, 127)
        .Range("A4:G4").Font.Color = RGB(255, 255, 255)
    End With

    outRow = 5
    totalItems = 0: totalMatch = 0: totalDiff = 0

    For i = 3 To lastInvRow
        artCode = Trim(wsInv.Cells(i, 1).Value)
        If artCode = "" Then GoTo NextInvRow

        physCount = mod_Utilities.SafeVal(wsInv.Cells(i, 3).Value)
        totalItems = totalItems + 1

        found = Application.Match(artCode, wsArt.Range("A:A"), 0)
        If IsError(found) Then
            sysStock = 0
        Else
            sysStock = mod_Utilities.SafeVal(wsArt.Cells(found, 7).Value)
        End If

        variance = physCount - sysStock
        variancePct = 0
        If sysStock > 0 Then variancePct = variance / sysStock

        Dim status As String
        If variance = 0 Then
            status = "OK"
            totalMatch = totalMatch + 1
        ElseIf Abs(variancePct) <= 0.05 Then
            status = "ECART_MINEUR"
            totalDiff = totalDiff + 1
        Else
            status = "ECART_MAJEUR"
            totalDiff = totalDiff + 1
        End If

        With wsOut
            .Cells(outRow, 1).Value = artCode
            .Cells(outRow, 2).Value = wsInv.Cells(i, 2).Value
            .Cells(outRow, 3).Value = physCount
            .Cells(outRow, 4).Value = sysStock
            .Cells(outRow, 5).Value = variance
            .Cells(outRow, 6).Value = variancePct
            .Cells(outRow, 6).NumberFormat = "0.00%"
            .Cells(outRow, 7).Value = status

            If variance <> 0 Then
                .Range(.Cells(outRow, 5), .Cells(outRow, 7)).Interior.Color = RGB(255, 235, 156)
            End If
            If Abs(variancePct) > 0.05 Then
                .Range(.Cells(outRow, 5), .Cells(outRow, 7)).Interior.Color = RGB(255, 200, 200)
            End If
        End With
        outRow = outRow + 1

NextInvRow:
    Next i

    Dim summaryRow As Long
    summaryRow = outRow + 1
    With wsOut
        .Cells(summaryRow, 1).Value = "RESUME"
        .Cells(summaryRow, 1).Font.Bold = True
        .Cells(summaryRow + 1, 1).Value = "Total articles inventories:"
        .Cells(summaryRow + 1, 2).Value = totalItems
        .Cells(summaryRow + 2, 1).Value = "Correspondance parfaite:"
        .Cells(summaryRow + 2, 2).Value = totalMatch
        .Cells(summaryRow + 3, 1).Value = "Ecart constate:"
        .Cells(summaryRow + 3, 2).Value = totalDiff
        .Columns("A:G").AutoFit
    End With

    wsInv.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    wsOut.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Reconciliation inventaire terminee." & vbCrLf & _
           totalItems & " articles, " & totalMatch & " OK, " & totalDiff & " ecarts.", _
           vbInformation, "LSM v1.0.0"
    Exit Sub

ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Erreur reconciliation: " & Err.Description, vbCritical
End Sub

Public Sub ExportReconciliationReport()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_OUTPUT)
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "Aucun rapport de reconciliation trouve. Executez d'abord RunInventoryReconciliation.", vbExclamation
        Exit Sub
    End If

    Dim pdfPath As String
    pdfPath = ThisWorkbook.Path & "\Reconciliation_Inventaire_" & Format(Date, "YYYY-MM-DD") & ".pdf"

    ws.Unprotect Password:=mod_Config.MASTER_PWD
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, Quality:=xlQualityStandard
    ws.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True

    MsgBox "Rapport exporte: " & pdfPath, vbInformation
End Sub

Private Sub InitializeInventorySheet(ByVal ws As Worksheet)
    With ws
        .Cells(1, 1).Value = "CODE_ARTICLE"
        .Cells(1, 2).Value = "DESIGNATION"
        .Cells(1, 3).Value = "QTE_PHYSIQUE"
        .Cells(1, 4).Value = "EMPLACEMENT"
        .Cells(1, 5).Value = "OBSERVATIONS"
        .Range("A1:E1").Font.Bold = True
        .Range("A1:E1").Interior.Color = RGB(44, 62, 80)
        .Range("A1:E1").Font.Color = RGB(255, 255, 255)
        .Columns("A:E").AutoFit
    End With
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
