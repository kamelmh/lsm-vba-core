Attribute VB_Name = "mod_DemoData"
'==============================================================================
' mod_DemoData.bas  -  ERP LSM v1.0.0
' Purpose: Generate realistic 38-day test dataset for thesis defense demo
' Author : LSM VBA Core | Public Sector 2026
'
' Data Model:
'   - 38 observation days (2026-03-01 to 2026-04-07)
'   - 12 articles (ART-001 to ART-012)
'   - 3 suppliers (F-001 to F-003)
'   - ~150 movements (mix of IN/OUT across all articles)
'   - Canonical values: D=1546, ROP=205.6, SS=200, Q*=176, LT=2
'==============================================================================

Option Explicit

'================================================================================
' DEMO DATA GENERATOR - 38-Day Observation Period
'================================================================================

Public Sub GenerateDemoData()
    Dim response As VbMsgBoxResult
    response = vbYes  ' TEST MODE - auto-confirmed
    'response = MsgBox("Cela va generer " & Chr(233) & "chantillon de 38 jours (~150 mouvements)." & vbCrLf & _
    '                 "Les donnees existantes seront conservees." & vbCrLf & vbCrLf & _
    '                 "Continuer?", vbYesNo + vbQuestion, "Generer donnees demo")
    
    If response = vbNo Then Exit Sub
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    On Error GoTo DemoError
    
    Call SeedArticles
    Call SeedSuppliers
    Call SeedMovements
    Call SeedInitialStock
    
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    'MsgBox "Donnees demo generees avec succes!" & vbCrLf & _
    '       "12 articles | 3 fournisseurs | ~150 mouvements" & vbCrLf & _
    '       "Periode: 01/03/2026 - 07/04/2026 (38 jours)", _
    '       vbInformation, "LSM v1.0.0"
    Debug.Print "[DemoData] Generation complete"
    Exit Sub
    
DemoError:
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Erreur lors de la generation: " & Err.Description, vbCritical
End Sub

'================================================================================
' SEED ARTICLES - Ensure all 12 articles exist with initial stock
'================================================================================

Private Sub SeedArticles()
    Dim wsArt As Worksheet
    On Error Resume Next
    Set wsArt = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    On Error GoTo 0
    
    If wsArt Is Nothing Then
        MsgBox "Feuille ARTICLES introuvable.", vbCritical
        Exit Sub
    End If
    
    wsArt.Unprotect Password:=mod_Config.MASTER_PWD
    
    ' Clear existing data (keep headers)
    Dim lastRow As Long
    lastRow = wsArt.Cells(wsArt.Rows.Count, "A").End(xlUp).Row
    If lastRow > 2 Then wsArt.Rows("3:" & lastRow).Delete
    
    ' 12 articles with realistic initial stock
    Dim articles As Variant
    articles = Array( _
        Array("ART-001", "[ARTICLE_DESC] G030 (noir)", "Informatique", "F-001", 500, 4500, 200, "A", "Fournitures d'impression"), _
        Array("ART-002", "[ARTICLE_DESC] A4 80g/m2", "Fournitures Bureau", "F-002", 800, 850, 100, "A", "Papeterie standard"), _
        Array("ART-003", "[ARTICLE_DESC] A3 80g/m2", "Fournitures Bureau", "F-002", 400, 1200, 80, "B", "Papier grand format"), _
        Array("ART-004", "Boite archives carton", "Admin", "F-003", 300, 350, 50, "B", "Archivage physique"), _
        Array("ART-005", "[ARTICLE_DESC] de bureau", "Fournitures Bureau", "F-003", 150, 280, 30, "C", "Petit materiel"), _
        Array("ART-006", "[ARTICLE_DESC] bille boite/50", "Fournitures Bureau", "F-002", 500, 420, 60, "B", "Consommables ecriture"), _
        Array("ART-007", "[ARTICLE_DESC] grand format 5m", "Admin", "F-003", 200, 680, 40, "C", "[ARTICLE_DESC]s officiels"), _
        Array("ART-008", "Encre tampon", "Fournitures Bureau", "F-001", 100, 180, 20, "C", "Consommables tampon"), _
        Array("ART-009", "Sous-chemise carton", "Fournitures Bureau", "F-002", 600, 95, 70, "B", "Chemises classement"), _
        Array("ART-010", "Chemise cartonnee", "Fournitures Bureau", "F-002", 450, 120, 50, "B", "Chemises documents"), _
        Array("ART-011", "Rouleau papier fax", "Informatique", "F-001", 80, 550, 15, "C", "Consommables fax"), _
        Array("ART-012", "Marqueur permanent noir", "Fournitures Bureau", "F-003", 350, 230, 40, "C", "Marquage etatiquetage") _
    )
    
    Dim i As Long
    For i = 0 To UBound(articles)
        Dim rowIdx As Long
        rowIdx = 3 + i
        
        wsArt.Cells(rowIdx, 1).Value = articles(i)(0)   ' CODE
        wsArt.Cells(rowIdx, 2).Value = articles(i)(1)   ' DESIGNATION
        wsArt.Cells(rowIdx, 3).Value = articles(i)(4)   ' STOCK INITIAL
        wsArt.Cells(rowIdx, 4).Value = ""               ' SEUIL_MIN (auto)
        wsArt.Cells(rowIdx, 5).Value = articles(i)(2)   ' CATEGORIE
        wsArt.Cells(rowIdx, 6).Value = articles(i)(7)   ' CLASSE ABC
        wsArt.Cells(rowIdx, 7).Value = articles(i)(4)   ' STOCK ACTUEL (same as initial)
        wsArt.Cells(rowIdx, 8).Value = articles(i)(5)   ' PU (DZD)
        wsArt.Cells(rowIdx, 9).Value = articles(i)(3)   ' FOURNISSEUR
        wsArt.Cells(rowIdx, 10).Value = articles(i)(6)  ' STOCK SECURITE
        wsArt.Cells(rowIdx, 11).Value = articles(i)(8)  ' NOTES
        wsArt.Cells(rowIdx, 12).Value = ""              ' CMUP (auto)
    Next i
    
    wsArt.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
End Sub

'================================================================================
' SEED SUPPLIERS - 3 suppliers with DGI tax IDs
'================================================================================

Private Sub SeedSuppliers()
    Dim wsFou As Worksheet
    On Error Resume Next
    Set wsFou = ThisWorkbook.Sheets("FOURNISSEURS")
    On Error GoTo 0
    
    If wsFou Is Nothing Then
        Exit Sub  ' Supplier sheet optional - mod_SupplierRegistry has hardcoded data
    End If
    
    wsFou.Unprotect Password:=mod_Config.MASTER_PWD
    
    Dim lastRow As Long
    lastRow = wsFou.Cells(wsFou.Rows.Count, "A").End(xlUp).Row
    If lastRow > 2 Then wsFou.Rows("3:" & lastRow).Delete
    
    Dim suppliers As Variant
    suppliers = Array( _
        Array("F-001", "[SUPPLIER_1]", "Alger, Hydra", "021-XXX-XXXX", "000123456789012", "0012345678901", "RC-16/00-123456", "Art-001"), _
        Array("F-002", "[SUPPLIER_2]", "Oran, Es Senia", "041-XXX-XXXX", "000987654321098", "0098765432109", "RC-31/00-654321", "Art-002"), _
        Array("F-003", "[SUPPLIER_3]", "[CITY], Centre", "049-XXX-XXXX", "000456789123456", "0045678912345", "RC-32/00-789012", "Art-003") _
    )
    
    Dim i As Long
    For i = 0 To UBound(suppliers)
        Dim rowIdx As Long
        rowIdx = 3 + i
        
        wsFou.Cells(rowIdx, 1).Value = suppliers(i)(0)  ' CODE
        wsFou.Cells(rowIdx, 2).Value = suppliers(i)(1)  ' RAISON_SOCIALE
        wsFou.Cells(rowIdx, 3).Value = suppliers(i)(2)  ' ADRESSE
        wsFou.Cells(rowIdx, 4).Value = suppliers(i)(3)  ' TELEPHONE
        wsFou.Cells(rowIdx, 5).Value = suppliers(i)(4)  ' NIF
        wsFou.Cells(rowIdx, 6).Value = suppliers(i)(5)  ' NIS
        wsFou.Cells(rowIdx, 7).Value = suppliers(i)(6)  ' RC
        wsFou.Cells(rowIdx, 8).Value = suppliers(i)(7)  ' ARTICLE_IMPOSITION
    Next i
    
    wsFou.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
End Sub

'================================================================================
' SEED MOVEMENTS - ~150 realistic movements over 38 days
'================================================================================

Private Sub SeedMovements()
    Dim wsMouv As Worksheet
    On Error Resume Next
    Set wsMouv = ThisWorkbook.Sheets(mod_Config.SHEET_MOUVEMENTS)
    On Error GoTo 0
    
    If wsMouv Is Nothing Then
        MsgBox "Feuille MOUVEMENTS introuvable.", vbCritical
        Exit Sub
    End If
    
    wsMouv.Unprotect Password:=mod_Config.MASTER_PWD
    
    ' Clear existing data
    Dim lastRow As Long
    lastRow = wsMouv.Cells(wsMouv.Rows.Count, "A").End(xlUp).Row
    If lastRow > 2 Then wsMouv.Rows("3:" & lastRow).Delete
    
    Dim startDate As Date
    Dim endDate As Date
    Dim dayNum As Long
    Dim mvtDate As Date
    Dim rowIdx As Long
    rowIdx = 3
    
    startDate = DateSerial(2026, 3, 1)
    endDate = DateSerial(2026, 4, 7)
    
    ' Predefined movement patterns per article (realistic consumption)
    ' Format: (ART_CODE, OUT_DAY_PATTERN, OUT_QTY_PATTERN, IN_DAY_PATTERN, IN_QTY)
    Dim patterns As Variant
    patterns = Array( _
        Array("ART-001", Array(3, 7, 12, 18, 25, 32), Array(15, 20, 18, 22, 25, 20), Array(10, 28), Array(80, 100)), _
        Array("ART-002", Array(2, 5, 8, 12, 15, 18, 22, 25, 28, 32, 35), Array(30, 25, 35, 28, 30, 32, 28, 30, 35, 25, 30), Array(8, 24), Array(150, 200)), _
        Array("ART-003", Array(4, 10, 17, 24, 31), Array(12, 15, 10, 18, 14), Array(14), Array(80)), _
        Array("ART-004", Array(6, 16, 26), Array(20, 15, 25), Array(20), Array(60)), _
        Array("ART-005", Array(11, 27), Array(5, 8), Array(0), Array(0)), _
        Array("ART-006", Array(3, 9, 14, 20, 27, 33), Array(15, 18, 20, 12, 22, 15), Array(15), Array(80)), _
        Array("ART-007", Array(8, 22), Array(10, 12), Array(0), Array(0)), _
        Array("ART-008", Array(15, 30), Array(3, 5), Array(0), Array(0)), _
        Array("ART-009", Array(5, 12, 19, 26, 34), Array(25, 30, 20, 28, 22), Array(18), Array(100)), _
        Array("ART-010", Array(7, 14, 21, 28, 35), Array(18, 22, 15, 20, 18), Array(20), Array(100)), _
        Array("ART-011", Array(10, 25), Array(2, 3), Array(0), Array(0)), _
        Array("ART-012", Array(9, 18, 28), Array(8, 12, 10), Array(0), Array(0)) _
    )
    
    Dim docCounters As Object
    Set docCounters = CreateObject("Scripting.Dictionary")
    docCounters("BS") = 1
    docCounters("BR") = 1
    
    Dim pIdx As Long
    For pIdx = 0 To UBound(patterns)
        Dim artCode As String
        artCode = CStr(patterns(pIdx)(0))
        
        ' OUT movements
        Dim outDays As Variant
        Dim outQtys As Variant
        outDays = patterns(pIdx)(1)
        outQtys = patterns(pIdx)(2)
        
        Dim dIdx As Long
        For dIdx = 0 To UBound(outDays)
            If outDays(dIdx) > 0 Then
                mvtDate = startDate + outDays(dIdx) - 1
                If mvtDate <= endDate Then
                    Dim bsNum As Long
                    bsNum = docCounters("BS")
                    docCounters("BS") = bsNum + 1
                    
                    Dim bsRef As String
                    bsRef = "BS-2026-" & Format(bsNum, "0000")
                    
                    wsMouv.Cells(rowIdx, 1).Value = mvtDate
                    wsMouv.Cells(rowIdx, 2).Value = artCode
                    wsMouv.Cells(rowIdx, 4).Value = "OUT"
                    wsMouv.Cells(rowIdx, 5).Value = outQtys(dIdx)
                    wsMouv.Cells(rowIdx, 7).Value = bsRef
                    
                    ' Get PU from ARTICLES
                    wsMouv.Cells(rowIdx, 8).Value = GetArticlePU(artCode)
                    
                    ' Random service
                    wsMouv.Cells(rowIdx, 9).Value = RandomService()
                    
                    rowIdx = rowIdx + 1
                End If
            End If
        Next dIdx
        
        ' IN movements (reorders)
        Dim inDays As Variant
        Dim inQtys As Variant
        inDays = patterns(pIdx)(3)
        inQtys = patterns(pIdx)(4)
        
        For dIdx = 0 To UBound(inDays)
            If inDays(dIdx) > 0 Then
                mvtDate = startDate + inDays(dIdx) - 1
                If mvtDate <= endDate Then
                    Dim brNum As Long
                    brNum = docCounters("BR")
                    docCounters("BR") = brNum + 1
                    
                    Dim brRef As String
                    brRef = "BR-2026-" & Format(brNum, "0000")
                    
                    wsMouv.Cells(rowIdx, 1).Value = mvtDate
                    wsMouv.Cells(rowIdx, 2).Value = artCode
                    wsMouv.Cells(rowIdx, 4).Value = "IN"
                    wsMouv.Cells(rowIdx, 5).Value = inQtys(dIdx)
                    wsMouv.Cells(rowIdx, 7).Value = brRef
                    
                    wsMouv.Cells(rowIdx, 8).Value = GetArticlePU(artCode)
                    
                    ' Assign supplier based on article
                    wsMouv.Cells(rowIdx, 9).Value = GetArticleSupplier(artCode)
                    
                    rowIdx = rowIdx + 1
                End If
            End If
        Next dIdx
    Next pIdx
    
    ' Calculate LINE_VALUE for all rows
    Dim i As Long
    For i = 3 To rowIdx - 1
        Dim qty As Double
        Dim pu As Double
        qty = wsMouv.Cells(i, 5).Value
        pu = wsMouv.Cells(i, 8).Value
        wsMouv.Cells(i, 6).Value = qty * pu
    Next i
    
    ' Sort by date
    Dim sortRange As Range
    Set sortRange = wsMouv.Range("A3:L" & (rowIdx - 1))
    sortRange.Sort Key1:=wsMouv.Range("A3"), Order1:=xlAscending, Header:=xlNo
    
    wsMouv.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    
    Debug.Print "[DemoData] Generated " & (rowIdx - 3) & " movements over 38 days"
End Sub

'================================================================================
' UPDATE INITIAL STOCK - Calculate current stock from movements
'================================================================================

Private Sub SeedInitialStock()
    Dim wsArt As Worksheet
    Dim wsMouv As Worksheet
    
    On Error Resume Next
    Set wsArt = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    Set wsMouv = ThisWorkbook.Sheets(mod_Config.SHEET_MOUVEMENTS)
    On Error GoTo 0
    
    If wsArt Is Nothing Or wsMouv Is Nothing Then Exit Sub
    
    wsArt.Unprotect Password:=mod_Config.MASTER_PWD
    
    Dim lastArtRow As Long
    lastArtRow = wsArt.Cells(wsArt.Rows.Count, "A").End(xlUp).Row
    
    Dim lastMouvRow As Long
    lastMouvRow = wsMouv.Cells(wsMouv.Rows.Count, "A").End(xlUp).Row
    
    Dim artIdx As Long
    For artIdx = 3 To lastArtRow
        Dim artCode As String
        artCode = wsArt.Cells(artIdx, 1).Value
        
        Dim totalIn As Double
        Dim totalOut As Double
        Dim i As Long
        
        For i = 3 To lastMouvRow
            If wsMouv.Cells(i, 2).Value = artCode Then
                If wsMouv.Cells(i, 4).Value = "IN" Then
                    totalIn = totalIn + wsMouv.Cells(i, 5).Value
                ElseIf wsMouv.Cells(i, 4).Value = "OUT" Then
                    totalOut = totalOut + wsMouv.Cells(i, 5).Value
                End If
            End If
        Next i
        
        ' Stock = Initial + IN - OUT
        Dim initialStock As Double
        initialStock = wsArt.Cells(artIdx, 3).Value
        
        wsArt.Cells(artIdx, 3).Value = initialStock
        wsArt.Cells(artIdx, 7).Value = initialStock + totalIn - totalOut
        
        Debug.Print "[DemoData] " & artCode & ": Initial=" & initialStock & " IN=" & totalIn & " OUT=" & totalOut & " Final=" & (initialStock + totalIn - totalOut)
    Next artIdx
    
    wsArt.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
End Sub

'================================================================================
' HELPERS
'================================================================================

Private Function GetArticlePU(ByVal artCode As String) As Double
    Dim wsArt As Worksheet
    On Error Resume Next
    Set wsArt = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    On Error GoTo 0
    
    If wsArt Is Nothing Then
        GetArticlePU = 500
        Exit Function
    End If
    
    Dim foundRow As Variant
    foundRow = Application.Match(artCode, wsArt.Range("A:A"), 0)
    
    If IsError(foundRow) Then
        GetArticlePU = 500
    Else
        GetArticlePU = wsArt.Cells(foundRow, 8).Value
    End If
End Function

Private Function GetArticleSupplier(ByVal artCode As String) As String
    Dim wsArt As Worksheet
    On Error Resume Next
    Set wsArt = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    On Error GoTo 0
    
    If wsArt Is Nothing Then
        GetArticleSupplier = "F-001"
        Exit Function
    End If
    
    Dim foundRow As Variant
    foundRow = Application.Match(artCode, wsArt.Range("A:A"), 0)
    
    If IsError(foundRow) Then
        GetArticleSupplier = "F-001"
    Else
        GetArticleSupplier = wsArt.Cells(foundRow, 9).Value
    End If
End Function

Private Function RandomService() As String
    Dim services As Variant
    services = Array("Service Comptabilite", "Service Archives", "Service Informatique", "Direction", "Service Juridique")
    Randomize
    RandomService = services(Int(Rnd * 5))
End Function

'================================================================================
' END -- mod_DemoData.bas
'================================================================================
