Attribute VB_Name = "mod_CSVImportExport"
Option Explicit

Private Const CSV_EXT As String = "CSV Files (*.csv),*.csv"
Private Const ALL_EXT As String = "All Files (*.*),*.*"

Public Sub ExportMouvementsToCSV()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim csvContent As String
    Dim savePath As Variant
    Dim i As Long

    Set ws = ThisWorkbook.Sheets(mod_Config.SHEET_MOUVEMENTS)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 1 Then
        MsgBox "Aucune donn" & Chr(233) & "e trouv" & Chr(233) & "e dans MOUVEMENTS.", vbExclamation
        Exit Sub
    End If

    savePath = Application.GetSaveAsFilename( _
        FileFilter:=CSV_EXT, _
        Title:="Exporter MOUVEMENTS en CSV", _
        InitialFileName:="MOUVEMENTS_Export_" & Format(Date, "yyyy-mm-dd") & ".csv")

    If savePath = False Then Exit Sub

    csvContent = BuildCSVLine("DATE", "CODE_ARTICLE", "DESIGNATION", "TYPE_MVT", "QTE", "VALEUR", "REF_DOCUMENT", "PRIX_UNITAIRE", "THIRD_PARTY", "NOTES")

    For i = 2 To lastRow
        Dim dateVal As Variant, codeVal As Variant, desigVal As Variant
        Dim typeVal As Variant, qteVal As Variant, valVal As Variant
        Dim refVal As Variant, puVal As Variant, tpVal As Variant, notesVal As Variant

        dateVal = ws.Cells(i, 1).Value
        codeVal = ws.Cells(i, 2).Value
        desigVal = ws.Cells(i, 3).Value
        typeVal = ws.Cells(i, 4).Value
        qteVal = ws.Cells(i, 5).Value
        valVal = ws.Cells(i, 6).Value
        refVal = ws.Cells(i, 7).Value
        puVal = ws.Cells(i, 8).Value
        tpVal = ws.Cells(i, 9).Value
        notesVal = ws.Cells(i, 12).Value

        csvContent = csvContent & vbCrLf & BuildCSVLine( _
            FormatDateTime(dateVal, vbShortDate), _
            CStr(codeVal), _
            CStr(desigVal), _
            CStr(typeVal), _
            CStr(qteVal), _
            CStr(valVal), _
            CStr(refVal), _
            CStr(puVal), _
            CStr(tpVal), _
            CStr(notesVal))
    Next i

    WriteCSVFile savePath, csvContent
    MsgBox "Export" & Chr(233) & " termin" & Chr(233) & ": " & savePath, vbInformation, "CSV Export"
End Sub

Public Sub ExportArticlesToCSV()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim csvContent As String
    Dim savePath As Variant
    Dim i As Long

    Set ws = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 1 Then
        MsgBox "Aucune donn" & Chr(233) & "e trouv" & Chr(233) & "e dans ARTICLES.", vbExclamation
        Exit Sub
    End If

    savePath = Application.GetSaveAsFilename( _
        FileFilter:=CSV_EXT, _
        Title:="Exporter ARTICLES en CSV", _
        InitialFileName:="ARTICLES_Export_" & Format(Date, "yyyy-mm-dd") & ".csv")

    If savePath = False Then Exit Sub

    csvContent = BuildCSVLine("CODE_ARTICLE", "DESIGNATION", "QTE_STOCK", "PRIX_UNITAIRE", "CATEGORIE", "MIN_STOCK")

    For i = 2 To lastRow
        csvContent = csvContent & vbCrLf & BuildCSVLine( _
            CStr(ws.Cells(i, 1).Value), _
            CStr(ws.Cells(i, 2).Value), _
            CStr(ws.Cells(i, 3).Value), _
            CStr(ws.Cells(i, 4).Value), _
            CStr(ws.Cells(i, 5).Value), _
            CStr(ws.Cells(i, 6).Value))
    Next i

    WriteCSVFile savePath, csvContent
    MsgBox "Export" & Chr(233) & " termin" & Chr(233) & ": " & savePath, vbInformation, "CSV Export"
End Sub

Public Sub ExportFournisseursToCSV()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim csvContent As String
    Dim savePath As Variant
    Dim i As Long

    Set ws = ThisWorkbook.Sheets(mod_Config.SHEET_FOURNISSEURS)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 3 Then
        MsgBox "Aucune donn" & Chr(233) & "e trouv" & Chr(233) & "e dans FOURNISSEURS.", vbExclamation
        Exit Sub
    End If

    savePath = Application.GetSaveAsFilename( _
        FileFilter:=CSV_EXT, _
        Title:="Exporter FOURNISSEURS en CSV", _
        InitialFileName:="FOURNISSEURS_Export_" & Format(Date, "yyyy-mm-dd") & ".csv")

    If savePath = False Then Exit Sub

    csvContent = BuildCSVLine("Code", "Raison Sociale", "Adresse", "T" & Chr(233) & "l" & Chr(233) & "phone", "NIF", "NIS", "RC", "Art. Imposition", "Cat" & Chr(233) & "gorie")

    For i = 3 To lastRow
        csvContent = csvContent & vbCrLf & BuildCSVLine( _
            CStr(ws.Cells(i, 1).Value), _
            CStr(ws.Cells(i, 2).Value), _
            CStr(ws.Cells(i, 3).Value), _
            CStr(ws.Cells(i, 4).Value), _
            CStr(ws.Cells(i, 5).Value), _
            CStr(ws.Cells(i, 6).Value), _
            CStr(ws.Cells(i, 7).Value), _
            CStr(ws.Cells(i, 8).Value), _
            CStr(ws.Cells(i, 9).Value))
    Next i

    WriteCSVFile savePath, csvContent
    MsgBox "Export" & Chr(233) & " termin" & Chr(233) & ": " & savePath, vbInformation, "CSV Export"
End Sub

Public Sub ImportMouvementsFromCSV()
    Dim filePath As Variant
    Dim lines() As String
    Dim totalLines As Long
    Dim imported As Long
    Dim errors As Long
    Dim i As Long
    Dim errorMsg As String

    filePath = Application.GetOpenFilename( _
        FileFilter:=CSV_EXT, _
        Title:="Importer MOUVEMENTS depuis CSV", _
        MultiSelect:=False)

    If filePath = False Then Exit Sub

    lines = ReadCSVLines(filePath)
    totalLines = UBound(lines) - LBound(lines) + 1

    If totalLines < 2 Then
        MsgBox "Fichier CSV vide ou invalide.", vbExclamation
        Exit Sub
    End If

    imported = 0
    errors = 0
    errorMsg = ""

    Application.ScreenUpdating = False

    For i = 1 To totalLines - 1
        Dim line As String
        Dim fields() As String
        Dim dateStr As String, codeStr As String, desigStr As String
        Dim typeStr As String, qtyStr As String, valStr As String
        Dim refStr As String, puStr As String, tpStr As String

        line = Trim(lines(i))
        If Len(line) = 0 Then GoTo NextLine

        fields = ParseCSVLine(line)
        If UBound(fields) < 5 Then
            errors = errors + 1
            errorMsg = errorMsg & "Ligne " & (i + 1) & ": champs insuffisants" & vbCrLf
            GoTo NextLine
        End If

        dateStr = Trim(fields(0))
        codeStr = Trim(fields(1))
        desigStr = Trim(fields(2))
        typeStr = Trim(fields(3))
        qtyStr = Trim(fields(4))
        valStr = IIf(UBound(fields) >= 5, Trim(fields(5)), "0")
        refStr = IIf(UBound(fields) >= 6, Trim(fields(6)), "")
        puStr = IIf(UBound(fields) >= 7, Trim(fields(7)), "0")
        tpStr = IIf(UBound(fields) >= 8, Trim(fields(8)), "")

        If Len(dateStr) = 0 Or Len(codeStr) = 0 Or Len(typeStr) = 0 Then
            errors = errors + 1
            errorMsg = errorMsg & "Ligne " & (i + 1) & ": champs obligatoires manquants (date, code, type)" & vbCrLf
            GoTo NextLine
        End If

        If Not IsNumeric(qtyStr) Or Not IsNumeric(puStr) Then
            errors = errors + 1
            errorMsg = errorMsg & "Ligne " & (i + 1) & ": quantit" & Chr(233) & " ou prix non num" & Chr(233) & "rique" & vbCrLf
            GoTo NextLine
        End If

        On Error Resume Next
        Call mod_Database.SecureWriteTransaction( _
            docDate:=CDate(dateStr), _
            typeSign:=typeStr, _
            refDoc:=refStr, _
            codeArticle:=codeStr, _
            designation:=desigStr, _
            quantity:=CLng(qtyStr), _
            unitPrice:=CDbl(puStr), _
            lineValue:=CDbl(valStr), _
            thirdParty:=tpStr)

        If Err.Number <> 0 Then
            errors = errors + 1
            errorMsg = errorMsg & "Ligne " & (i + 1) & ": " & Err.Description & vbCrLf
            Err.Clear
        Else
            imported = imported + 1
        End If
        On Error GoTo 0

NextLine:
    Next i

    Application.ScreenUpdating = True

    Dim result As String
    result = "Import termin" & Chr(233) & "." & vbCrLf & vbCrLf & _
             "Import" & Chr(233) & "s: " & imported & vbCrLf & _
             "Erreurs: " & errors & vbCrLf & _
             "Total: " & (imported + errors) & " / " & (totalLines - 1)

    If Len(errorMsg) > 0 Then
        result = result & vbCrLf & vbCrLf & "D" & Chr(233) & "tails des erreurs:" & vbCrLf & errorMsg
    End If

    MsgBox result, vbInformation, "CSV Import"
End Sub

Public Sub ImportArticlesFromCSV()
    Dim filePath As Variant
    Dim lines() As String
    Dim ws As Worksheet
    Dim nextRow As Long
    Dim imported As Long
    Dim errors As Long
    Dim i As Long
    Dim errorMsg As String

    filePath = Application.GetOpenFilename( _
        FileFilter:=CSV_EXT, _
        Title:="Importer ARTICLES depuis CSV", _
        MultiSelect:=False)

    If filePath = False Then Exit Sub

    lines = ReadCSVLines(filePath)

    If UBound(lines) - LBound(lines) + 1 < 2 Then
        MsgBox "Fichier CSV vide ou invalide.", vbExclamation
        Exit Sub
    End If

    Set ws = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    ws.Unprotect Password:=mod_Config.MASTER_PWD

    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    imported = 0
    errors = 0
    errorMsg = ""

    Application.ScreenUpdating = False

    For i = 1 To UBound(lines) - LBound(lines)
        Dim line As String
        Dim fields() As String

        line = Trim(lines(i))
        If Len(line) = 0 Then GoTo NextArt

        fields = ParseCSVLine(line)
        If UBound(fields) < 1 Then
            errors = errors + 1
            GoTo NextArt
        End If

        On Error Resume Next
        ws.Cells(nextRow, 1).Value = Trim(fields(0))
        ws.Cells(nextRow, 2).Value = IIf(UBound(fields) >= 1, Trim(fields(1)), "")
        ws.Cells(nextRow, 3).Value = IIf(UBound(fields) >= 2, mod_Utilities.SafeVal(fields(2)), 0)
        ws.Cells(nextRow, 4).Value = IIf(UBound(fields) >= 3, mod_Utilities.SafeVal(fields(3)), 0)
        ws.Cells(nextRow, 5).Value = IIf(UBound(fields) >= 4, Trim(fields(4)), "")
        ws.Cells(nextRow, 6).Value = IIf(UBound(fields) >= 5, mod_Utilities.SafeVal(fields(5)), 0)

        If Err.Number <> 0 Then
            errors = errors + 1
            errorMsg = errorMsg & "Ligne " & (i + 1) & ": " & Err.Description & vbCrLf
            Err.Clear
        Else
            imported = imported + 1
            nextRow = nextRow + 1
        End If
        On Error GoTo 0

NextArt:
    Next i

    ws.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    Application.ScreenUpdating = True

    MsgBox "Import termin" & Chr(233) & "." & vbCrLf & vbCrLf & _
           "Import" & Chr(233) & "s: " & imported & vbCrLf & _
           "Erreurs: " & errors, vbInformation, "CSV Import"
End Sub

Public Sub ImportFournisseursFromCSV()
    Dim filePath As Variant
    Dim lines() As String
    Dim ws As Worksheet
    Dim nextRow As Long
    Dim imported As Long
    Dim errors As Long
    Dim i As Long

    filePath = Application.GetOpenFilename( _
        FileFilter:=CSV_EXT, _
        Title:="Importer FOURNISSEURS depuis CSV", _
        MultiSelect:=False)

    If filePath = False Then Exit Sub

    lines = ReadCSVLines(filePath)

    If UBound(lines) - LBound(lines) + 1 < 2 Then
        MsgBox "Fichier CSV vide ou invalide.", vbExclamation
        Exit Sub
    End If

    Set ws = ThisWorkbook.Sheets(mod_Config.SHEET_FOURNISSEURS)
    ws.Unprotect Password:=mod_Config.MASTER_PWD

    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    imported = 0
    errors = 0

    Application.ScreenUpdating = False

    For i = 1 To UBound(lines) - LBound(lines)
        Dim line As String
        Dim fields() As String

        line = Trim(lines(i))
        If Len(line) = 0 Then GoTo NextFou

        fields = ParseCSVLine(line)
        If UBound(fields) < 1 Then
            errors = errors + 1
            GoTo NextFou
        End If

        On Error Resume Next
        ws.Cells(nextRow, 1).Value = Trim(fields(0))
        ws.Cells(nextRow, 2).Value = IIf(UBound(fields) >= 1, Trim(fields(1)), "")
        ws.Cells(nextRow, 3).Value = IIf(UBound(fields) >= 2, Trim(fields(2)), "")
        ws.Cells(nextRow, 4).Value = IIf(UBound(fields) >= 3, Trim(fields(3)), "")
        ws.Cells(nextRow, 5).Value = IIf(UBound(fields) >= 4, Trim(fields(4)), "")
        ws.Cells(nextRow, 6).Value = IIf(UBound(fields) >= 5, Trim(fields(5)), "")
        ws.Cells(nextRow, 7).Value = IIf(UBound(fields) >= 6, Trim(fields(6)), "")
        ws.Cells(nextRow, 8).Value = IIf(UBound(fields) >= 7, Trim(fields(7)), "")
        ws.Cells(nextRow, 9).Value = IIf(UBound(fields) >= 8, Trim(fields(8)), "")

        If Err.Number <> 0 Then
            errors = errors + 1
            Err.Clear
        Else
            imported = imported + 1
            nextRow = nextRow + 1
        End If
        On Error GoTo 0

NextFou:
    Next i

    ws.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    Application.ScreenUpdating = True

    MsgBox "Import termin" & Chr(233) & "." & vbCrLf & vbCrLf & _
           "Import" & Chr(233) & "s: " & imported & vbCrLf & _
           "Erreurs: " & errors, vbInformation, "CSV Import"
End Sub

Private Function BuildCSVLine(ParamArray values() As Variant) As String
    Dim i As Integer
    Dim parts() As String
    Dim count As Integer
    Dim val As String

    count = UBound(values) - LBound(values) + 1
    ReDim parts(count - 1)

    For i = LBound(values) To UBound(values)
        val = CStr(values(i))
        If InStr(val, ",") > 0 Or InStr(val, """") > 0 Or InStr(val, vbCrLf) > 0 Then
            val = Replace(val, """", """""")
            val = """" & val & """"
        End If
        parts(i - LBound(values)) = val
    Next i

    BuildCSVLine = Join(parts, ",")
End Function

Private Function ParseCSVLine(ByVal line As String) As String()
    Dim fields() As String
    Dim inQuote As Boolean
    Dim i As Long
    Dim current As String
    Dim fieldList As Collection
    Dim delim As String
    Dim result() As String
    Dim j As Long

    delim = ","
    If InStr(line, ";") > 0 And InStr(line, ",") = 0 Then
        delim = ";"
    End If

    Set fieldList = New Collection
    inQuote = False
    current = ""

    For i = 1 To Len(line)
        Dim ch As String
        ch = Mid(line, i, 1)

        If ch = """" Then
            inQuote = Not inQuote
        ElseIf ch = delim And Not inQuote Then
            fieldList.Add Trim(current)
            current = ""
        Else
            current = current & ch
        End If
    Next i

    fieldList.Add Trim(current)

    ReDim result(fieldList.Count - 1)
    For j = 1 To fieldList.Count
        result(j - 1) = fieldList(j)
    Next j

    ParseCSVLine = result
End Function

Private Function ReadCSVLines(ByVal filePath As String) As String()
    Dim fileContent As String
    Dim lines() As String
    Dim f As Integer

    f = FreeFile
    Open filePath For Input As #f
    fileContent = Input$(LOF(f), f)
    Close #f

    lines = Split(fileContent, vbCrLf)
    If UBound(lines) = 0 Then
        lines = Split(fileContent, vbLf)
    End If

    ReadCSVLines = lines
End Function

Private Sub WriteCSVFile(ByVal filePath As String, ByVal content As String)
    Dim f As Integer

    f = FreeFile
    Open filePath For Output As #f
    Print #f, content
    Close #f
End Sub
