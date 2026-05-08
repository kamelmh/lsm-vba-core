Attribute VB_Name = "mod_DataValidator"
Option Explicit

Private Const SHEET_VALIDATION As String = "VALIDATION_INTEGRITE"

Public Sub RunDataValidation()
    Dim wsOut As Worksheet
    Dim issues As Collection
    Dim issueCount As Long

    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False

    Set issues = New Collection

    Call ValidateMouvements(issues)
    Call ValidateArticles(issues)
    Call ValidateSuppliers(issues)

    Set wsOut = GetOrCreateSheet(SHEET_VALIDATION)
    wsOut.Unprotect Password:=mod_Config.MASTER_PWD
    wsOut.Cells.Clear

    With wsOut
        .Range("A1:E1").Merge
        .Cells(1, 1).Value = "RAPPORT DE VALIDATION D'INTEGRITE DES DONNEES"
        .Cells(1, 1).Font.Bold = True
        .Cells(1, 1).Font.Size = 14
        .Cells(2, 1).Value = "Date: " & Format(Now, "DD/MM/YYYY HH:MM")
        .Cells(2, 1).Font.Italic = True

        .Cells(4, 1).Value = "SOURCE"
        .Cells(4, 2).Value = "TYPE"
        .Cells(4, 3).Value = "DESCRIPTION"
        .Cells(4, 4).Value = "VALEUR"
        .Cells(4, 5).Value = "GRAVITE"
        .Range("A4:E4").Font.Bold = True
        .Range("A4:E4").Interior.Color = RGB(0, 70, 127)
        .Range("A4:E4").Font.Color = RGB(255, 255, 255)
    End With

    issueCount = issues.Count
    Dim rowNum As Long
    rowNum = 5

    If issueCount = 0 Then
        wsOut.Cells(rowNum, 1).Value = "AUCUN PROBLEME DETECTE"
        wsOut.Range("A" & rowNum & ":E" & rowNum).Merge
        wsOut.Cells(rowNum, 1).Font.Color = RGB(0, 128, 0)
        wsOut.Cells(rowNum, 1).Font.Bold = True
    Else
        Dim idx As Long
        For idx = 1 To issueCount
            Dim item
            item = issues.Item(idx)
            With wsOut
                .Cells(rowNum, 1).Value = item(0)
                .Cells(rowNum, 2).Value = item(1)
                .Cells(rowNum, 3).Value = item(2)
                .Cells(rowNum, 4).Value = item(3)
                .Cells(rowNum, 5).Value = item(4)

                If item(4) = "CRITIQUE" Then
                    .Range(.Cells(rowNum, 1), .Cells(rowNum, 5)).Interior.Color = RGB(255, 200, 200)
                ElseIf item(4) = "AVERTISSEMENT" Then
                    .Range(.Cells(rowNum, 1), .Cells(rowNum, 5)).Interior.Color = RGB(255, 243, 224)
                End If
            End With
            rowNum = rowNum + 1
        Next idx
    End If

    Dim sRow As Long
    sRow = rowNum + 1
    With wsOut
        .Cells(sRow, 1).Value = "RESUME"
        .Cells(sRow, 1).Font.Bold = True
        .Cells(sRow + 1, 1).Value = "Total problemes:"
        .Cells(sRow + 1, 2).Value = issueCount
        .Columns("A:E").AutoFit
    End With

    wsOut.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True

    Application.ScreenUpdating = True
    If issueCount > 0 Then
        MsgBox "Validation terminee. " & issueCount & " probleme(s) detecte(s). Consultez l'onglet " & SHEET_VALIDATION & ".", _
               vbExclamation, "LSM v1.0.0"
    Else
        MsgBox "Validation terminee. Aucun probleme detecte.", vbInformation, "LSM v1.0.0"
    End If
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Erreur validation: " & Err.Description, vbCritical
End Sub

Private Sub ValidateMouvements(ByRef issues As Collection)
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim artCode As String, mvtType As String, refDoc As String
    Dim qty As Double, mvtDate As Variant
    Dim wsArt As Worksheet, found As Variant

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(mod_Config.SHEET_MOUVEMENTS)
    Set wsArt = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    On Error GoTo 0

    If ws Is Nothing Then
        issues.Add Array("MOUVEMENTS", "ABSENT", "Feuille MOUVEMENTS introuvable", "", "CRITIQUE")
        Exit Sub
    End If

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
        artCode = Trim(ws.Cells(i, 2).Value)
        mvtType = Trim(ws.Cells(i, 4).Value)
        refDoc = Trim(ws.Cells(i, 7).Value)
        qty = mod_Utilities.SafeVal(ws.Cells(i, 5).Value)
        mvtDate = ws.Cells(i, 1).Value

        If artCode = "" Then
            issues.Add Array("MOUVEMENTS", "CODE_VIDE", "Ligne " & i & " - Code article manquant", "", "CRITIQUE")
            GoTo NextMvt
        End If

        If Len(artCode) < 5 Then
            issues.Add Array("MOUVEMENTS", "CODE_COURT", "Ligne " & i & " - Code suspect: " & artCode, artCode, "AVERTISSEMENT")
        End If

        If mvtType <> "IN" And mvtType <> "OUT" Then
            issues.Add Array("MOUVEMENTS", "TYPE_INVALIDE", "Ligne " & i & " - Type mouvement invalide: " & mvtType, mvtType, "CRITIQUE")
        End If

        If qty <= 0 Then
            issues.Add Array("MOUVEMENTS", "QUANTITE_INVALIDE", "Ligne " & i & " - Quantite <= 0", qty, "AVERTISSEMENT")
        End If

        If IsEmpty(mvtDate) Then
            issues.Add Array("MOUVEMENTS", "DATE_VIDE", "Ligne " & i & " - Date manquante", "", "CRITIQUE")
        ElseIf Not IsDate(mvtDate) Then
            issues.Add Array("MOUVEMENTS", "DATE_INVALIDE", "Ligne " & i & " - Date invalide: " & CStr(mvtDate), CStr(mvtDate), "CRITIQUE")
        End If

        If Not wsArt Is Nothing Then
            found = Application.Match(artCode, wsArt.Range("A:A"), 0)
            If IsError(found) Then
                issues.Add Array("MOUVEMENTS", "ARTICLE_ORPHELIN", "Ligne " & i & " - " & artCode & " n'existe pas dans ARTICLES", artCode, "CRITIQUE")
            End If
        End If
NextMvt:
    Next i

    ' Check duplicate ref
    Dim refs As Object
    Set refs = CreateObject("Scripting.Dictionary")
    For i = 2 To lastRow
        refDoc = Trim(ws.Cells(i, 7).Value)
        If refDoc <> "" Then
            If refs.Exists(refDoc) Then
                issues.Add Array("MOUVEMENTS", "REF_DUPLICATE", "Reference dupliquee: " & refDoc, refDoc, "AVERTISSEMENT")
            Else
                refs.Add refDoc, True
            End If
        End If
    Next i
End Sub

Private Sub ValidateArticles(ByRef issues As Collection)
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim artCode As String, stock As Double, pu As Double
    Dim codes As Object

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    On Error GoTo 0

    If ws Is Nothing Then
        issues.Add Array("ARTICLES", "ABSENT", "Feuille ARTICLES introuvable", "", "CRITIQUE")
        Exit Sub
    End If

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Set codes = CreateObject("Scripting.Dictionary")

    For i = 2 To lastRow
        artCode = Trim(ws.Cells(i, 1).Value)

        If artCode = "" Then
            issues.Add Array("ARTICLES", "CODE_VIDE", "Ligne " & i & " - Code article manquant", "", "CRITIQUE")
            GoTo NextArt
        End If

        If codes.Exists(artCode) Then
            issues.Add Array("ARTICLES", "DUPLICATE", "Code article duplique: " & artCode, artCode, "CRITIQUE")
        Else
            codes.Add artCode, True
        End If

        stock = mod_Utilities.SafeVal(ws.Cells(i, 3).Value)
        If stock < 0 Then
            issues.Add Array("ARTICLES", "STOCK_NEGATIF", artCode & " - Stock negatif: " & stock, stock, "CRITIQUE")
        End If

        pu = mod_Utilities.SafeVal(ws.Cells(i, 8).Value)
        If pu <= 0 Then
            issues.Add Array("ARTICLES", "PRIX_INVALIDE", artCode & " - Prix unitaire <= 0", pu, "AVERTISSEMENT")
        End If

        Dim designation As String
        designation = Trim(ws.Cells(i, 2).Value)
        If designation = "" Then
            issues.Add Array("ARTICLES", "DESIGNATION_VIDE", artCode & " - Designation manquante", artCode, "AVERTISSEMENT")
        End If
NextArt:
    Next i
End Sub

Private Sub ValidateSuppliers(ByRef issues As Collection)
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim fouCode As String, nif As String, nis As String

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(mod_Config.SHEET_FOURNISSEURS)
    On Error GoTo 0

    If ws Is Nothing Then
        issues.Add Array("FOURNISSEURS", "ABSENT", "Feuille FOURNISSEURS introuvable", "", "AVERTISSEMENT")
        Exit Sub
    End If

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = 3 To lastRow
        fouCode = Trim(ws.Cells(i, 1).Value)
        If fouCode = "" Then
            issues.Add Array("FOURNISSEURS", "CODE_VIDE", "Ligne " & i & " - Code fournisseur manquant", "", "AVERTISSEMENT")
            GoTo NextFou
        End If

        nif = Trim(ws.Cells(i, 5).Value)
        nis = Trim(ws.Cells(i, 6).Value)
        If Len(nif) < 10 Then
            issues.Add Array("FOURNISSEURS", "NIF_COURT", fouCode & " - NIF semble invalide: " & nif, nif, "AVERTISSEMENT")
        End If
        If Len(nis) < 10 Then
            issues.Add Array("FOURNISSEURS", "NIS_COURT", fouCode & " - NIS semble invalide: " & nis, nis, "AVERTISSEMENT")
        End If
NextFou:
    Next i
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
