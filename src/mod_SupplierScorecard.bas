Attribute VB_Name = "mod_SupplierScorecard"
Option Explicit

Private Const SHEET_SCORECARD As String = "SCORECARD_FOURNISSEURS"

Public Sub RunSupplierScorecard()
    Dim wsMouv As Worksheet, wsArt As Worksheet, wsOut As Worksheet
    Dim lastMouvRow As Long, lastArtRow As Long
    Dim i As Long, j As Long, outRow As Long
    Dim supplierCode As String, supplierName As String
    Dim totalOrders As Long, totalQty As Double, totalValue As Double
    Dim onTimeDeliveries As Long, totalDeliveries As Long
    Dim deliveryScore As Double, volumeScore As Double, overallScore As Double
    Dim rating As Double
    Dim suppliers As Variant
    Dim artSupplierMap As Object
    Dim poCount As Object

    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Set wsMouv = ThisWorkbook.Sheets(mod_Config.SHEET_MOUVEMENTS)
    Set wsArt = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    lastMouvRow = wsMouv.Cells(wsMouv.Rows.Count, 1).End(xlUp).Row
    lastArtRow = wsArt.Cells(wsArt.Rows.Count, 1).End(xlUp).Row

    Set artSupplierMap = CreateObject("Scripting.Dictionary")
    For i = 3 To lastArtRow
        Dim aCode As String
        aCode = Trim(wsArt.Cells(i, 1).Value)
        If aCode <> "" Then
            artSupplierMap(aCode) = Trim(wsArt.Cells(i, 9).Value)
        End If
    Next i

    Set poCount = CreateObject("Scripting.Dictionary")

    Dim mvtType As String, qty As Double, val As Double, mvtSupplier As String
    For j = 3 To lastMouvRow
        mvtType = Trim(wsMouv.Cells(j, 4).Value)
        qty = mod_Utilities.SafeVal(wsMouv.Cells(j, 5).Value)
        val = mod_Utilities.SafeVal(wsMouv.Cells(j, 6).Value)

        If mvtType = "IN" Then
            mvtSupplier = Trim(wsMouv.Cells(j, 9).Value)
            If mvtSupplier = "" Then
                mvtSupplier = GetArticleSupplier(wsMouv.Cells(j, 2).Value)
            End If
            If mvtSupplier <> "" Then
                If Not poCount.Exists(mvtSupplier) Then
                    poCount.Add mvtSupplier, Array(0, 0, 0)
                End If
                Dim arr
                arr = poCount(mvtSupplier)
                arr(0) = arr(0) + 1
                arr(1) = arr(1) + qty
                arr(2) = arr(2) + val
                poCount(mvtSupplier) = arr
            End If
        End If
    Next j

    suppliers = Array("F-001", "F-002", "F-003")

    Set wsOut = GetOrCreateSheet(SHEET_SCORECARD)
    wsOut.Unprotect Password:=mod_Config.MASTER_PWD
    wsOut.Cells.Clear

    With wsOut
        .Range("A1:I1").Merge
        .Cells(1, 1).Value = "TABLEAU DE BORD FOURNISSEURS"
        .Cells(1, 1).Font.Bold = True
        .Cells(1, 1).Font.Size = 14
        .Cells(2, 1).Value = "Date: " & Format(Now, "DD/MM/YYYY HH:MM")
        .Cells(2, 1).Font.Italic = True

        .Cells(4, 1).Value = "CODE"
        .Cells(4, 2).Value = "RAISON_SOCIALE"
        .Cells(4, 3).Value = "COMMANDES"
        .Cells(4, 4).Value = "QUANTITE_TOTALE"
        .Cells(4, 5).Value = "VALEUR_TOTALE (DZD)"
        .Cells(4, 6).Value = "SCORE_VOLUME"
        .Cells(4, 7).Value = "NOTE_INTERNE"
        .Cells(4, 8).Value = "SCORE_GLOBAL"
        .Cells(4, 9).Value = "CLASSEMENT"
        .Range("A4:I4").Font.Bold = True
        .Range("A4:I4").Interior.Color = RGB(0, 70, 127)
        .Range("A4:I4").Font.Color = RGB(255, 255, 255)
    End With

    outRow = 5
    Dim maxValue As Double: maxValue = 0

    For i = 0 To UBound(suppliers)
        supplierCode = suppliers(i)
        supplierName = mod_SupplierRegistry.GetSupplierLegalName(supplierCode)
        rating = mod_SupplierRegistry.GetSupplierRating(supplierCode)

        If poCount.Exists(supplierCode) Then
            Dim data
            data = poCount(supplierCode)
            totalOrders = data(0)
            totalQty = data(1)
            totalValue = data(2)
        Else
            totalOrders = 0: totalQty = 0: totalValue = 0
        End If

        If totalValue > maxValue Then maxValue = totalValue
    Next i

    For i = 0 To UBound(suppliers)
        supplierCode = suppliers(i)
        supplierName = mod_SupplierRegistry.GetSupplierLegalName(supplierCode)
        rating = mod_SupplierRegistry.GetSupplierRating(supplierCode)

        If poCount.Exists(supplierCode) Then
            data = poCount(supplierCode)
            totalOrders = data(0)
            totalQty = data(1)
            totalValue = data(2)
        Else
            totalOrders = 0: totalQty = 0: totalValue = 0
        End If

        volumeScore = 0
        If maxValue > 0 Then volumeScore = totalValue / maxValue * 100

        deliveryScore = rating / 5 * 100
        overallScore = (volumeScore * 0.5) + (deliveryScore * 0.5)

        Dim rank As String
        If overallScore >= 80 Then
            rank = "A - Excellent"
        ElseIf overallScore >= 60 Then
            rank = "B - Satisfaisant"
        ElseIf overallScore >= 40 Then
            rank = "C - Moyen"
        Else
            rank = "D - Insuffisant"
        End If

        With wsOut
            .Cells(outRow, 1).Value = supplierCode
            .Cells(outRow, 2).Value = supplierName
            .Cells(outRow, 3).Value = totalOrders
            .Cells(outRow, 4).Value = totalQty
            .Cells(outRow, 5).Value = totalValue
            .Cells(outRow, 5).NumberFormat = "#,##0.00"
            .Cells(outRow, 6).Value = Round(volumeScore, 1)
            .Cells(outRow, 6).NumberFormat = "0.0"
            .Cells(outRow, 7).Value = rating
            .Cells(outRow, 7).NumberFormat = "0.0"
            .Cells(outRow, 8).Value = Round(overallScore, 1)
            .Cells(outRow, 8).NumberFormat = "0.0"
            .Cells(outRow, 9).Value = rank

            If overallScore >= 80 Then
                .Range(.Cells(outRow, 1), .Cells(outRow, 9)).Interior.Color = RGB(211, 240, 224)
            ElseIf overallScore >= 60 Then
                .Range(.Cells(outRow, 1), .Cells(outRow, 9)).Interior.Color = RGB(255, 243, 224)
            Else
                .Range(.Cells(outRow, 1), .Cells(outRow, 9)).Interior.Color = RGB(255, 220, 220)
            End If

            outRow = outRow + 1
        End With
    Next i

    wsOut.Columns("A:I").AutoFit
    wsOut.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Scorecard fournisseurs mis a jour. 3 fournisseurs evalues.", vbInformation, "LSM v1.0.0"
    Exit Sub

ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Erreur scorecard: " & Err.Description, vbCritical
End Sub

Public Sub ExportScorecardPDF()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_SCORECARD)
    On Error GoTo 0
    If ws Is Nothing Then MsgBox "Execute d'abord RunSupplierScorecard.", vbExclamation: Exit Sub

    Dim pdfPath As String
    pdfPath = ThisWorkbook.Path & "\Scorecard_Fournisseurs_" & Format(Date, "YYYY-MM-DD") & ".pdf"

    ws.Unprotect Password:=mod_Config.MASTER_PWD
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, Quality:=xlQualityStandard
    ws.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    MsgBox "Rapport exporte: " & pdfPath, vbInformation
End Sub

Private Function GetArticleSupplier(ByVal artCode As Variant) As String
    Dim wsArt As Worksheet
    Dim found As Variant
    On Error Resume Next
    Set wsArt = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    On Error GoTo 0
    If wsArt Is Nothing Then GetArticleSupplier = "": Exit Function

    found = Application.Match(artCode, wsArt.Range("A:A"), 0)
    If IsError(found) Then
        GetArticleSupplier = ""
    Else
        GetArticleSupplier = Trim(wsArt.Cells(found, 9).Value)
    End If
End Function

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
