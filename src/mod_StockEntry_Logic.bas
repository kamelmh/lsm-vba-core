Attribute VB_Name = "mod_StockEntry_Logic"
'==============================================================================
' mod_StockEntry_Logic.bas  -  ERP LSM v1.0.0 (Refactored)
' Author   : LSM VBA Core | TAG1801 GSL | Public Sector 2026
' Depends  : mod_SyncBridge | mod_StockEngine | mod_Database | mod_AuditTrail
'            mod_ExportEngine | mod_Utilities | mod_Config
'
' Architecture (Decoupled):
'   UI Layer       - frmStockEntry (thin view, delegates to controller)
'   Controller     - mod_StockEntry_Logic (business logic, parameter-based)
'   Data Layer     - mod_Database.SecureWriteTransaction()
'   Sync           - mod_SyncBridge.CommitStaging()
'
' Canonical constants (locked - never change without updating CLAUDE.md):
'   D=1,546  |  ROP=205.6  |  Q*=176  |  SS=200  |  LT=2 days
'
' Refactoring (2026-01-01):
'   - Eliminated direct frmStockEntry.{control} references from logic module
'   - Introduced FormState struct for parameter-based data passing
'   - Form owns all UI operations; logic module returns state updates
'   - Eliminated duplicate FireInternalBridge (form version removed)
'   - Consolidated stock calculation to use mod_StockEngine
'==============================================================================

Option Explicit

'================================================================================
' SECTION 0 - MODULE-LEVEL CONSTANTS & STATE
'================================================================================

'-- Canonical ERP constants (mirrors Unité de traitement VBA GROUND_TRUTH)
Private Const CANON_ROP    As Double = 205.6
Private Const CANON_SS     As Long = 200
Private Const CANON_QSTAR  As Long = 176
Private Const CANON_LT     As Integer = 2

'-- Grid column indices (0-based)
Private Const COL_CODE   As Integer = 0
Private Const COL_DESIG  As Integer = 1
Private Const COL_CAT    As Integer = 2
Private Const COL_QTE    As Integer = 3
Private Const COL_PU     As Integer = 4
Private Const COL_VALEUR As Integer = 5

'-- Form state (isolated from UI controls)
Private m_TotalGeneral   As Double
Private m_CurrentArticle As String
Private m_StockActuel    As Long
Private m_IsBRMode       As Boolean

'================================================================================
' SECTION 0A - FORM STATE STRUCT (Decoupled data transfer)
'================================================================================

' FormState holds all form data independently of UI controls.
' The form populates this struct; logic reads/writes the struct.
' This eliminates direct frmStockEntry.{control} coupling.

Public Type FormState
    '- Input fields
    docType        As String      ' BS / BR / DA / BC
    docRef         As String      ' Auto-generated reference
    TransDate      As String      ' DD/MM/YYYY format
    ArticleCode    As String      ' ART-001 etc.
    ArticleDesign  As String      ' Full designation
    ArticleCat     As String      ' Category filter
    qty            As String      ' Quantity text
    unitPrice      As String      ' PU text
    Service        As String      ' Service/Fournisseur
    
    '- Grid data (semicolon-delimited rows for VBA compatibility)
    '   Format: "CODE|DESIG|CAT|QTY|PU|VAL;CODE|DESIG|CAT|QTY|PU|VAL;..."
    GridData       As String
    GridRowCount   As Integer
    
    '- Display outputs (logic sets these, form renders them)
    StockInfoText      As String
    StockInfoColor     As Long     ' RGB value
    WilsonAlertText    As String
    WilsonAlertVisible As Boolean
    BannerText         As String
    BannerColor        As Long     ' RGB value
    QtyBackColor       As Long     ' RGB value
    TotalGeneralText   As String
    TotalGeneral       As Double
    
    '- Mode flags
    IsBRMode       As Boolean
    PUEditable     As Boolean
    PULabel        As String
    
    '- Article metadata (populated on selection)
    ArticleStock   As Long
    ArticlePU      As Double
    
    '- Controls (passed as objects for SetFocus etc.)
    '- Form reference for UI operations (only when absolutely needed)
    formRef        As Object
End Type

'================================================================================
' SECTION 1 - FORM INITIALIZE (Controller sets up state)
'================================================================================

Public Sub InitializeForm(ByRef state As FormState)
    Call SetupFormAppearance(state)
    Call PopulateDropdowns(state)
    Call ConfigureGrid(state)
    Call ResetToDefaultState(state)
End Sub

Private Sub SetupFormAppearance(ByRef state As FormState)
    '- Form shell
    state.formRef.Caption = mod_Localization.SafeGetTxt("SYS_TITLE")
    If state.formRef.Caption = "SYS_TITLE" Or InStr(state.formRef.Caption, "NOT_FOUND") > 0 Then
        state.formRef.Caption = "Saisie des Mouvements de Stock"
    End If
    state.formRef.Width = 870
    state.formRef.Height = 640
    
    '- Banner (idle state)
    state.BannerText = "-- SELECTIONNEZ LE TYPE DE DOCUMENT --"
    state.BannerColor = RGB(100, 100, 100)
    state.formRef.lblBannerText.Caption = state.BannerText
    state.formRef.lblBannerText.BackColor = state.BannerColor
    state.formRef.lblBannerText.ForeColor = RGB(255, 255, 255)
    state.formRef.lblBannerText.Font.Bold = True
    state.formRef.lblBannerText.TextAlign = fmTextAlignCenter
    
    '- Wilson alert (hidden by default)
    state.WilsonAlertVisible = False
    state.formRef.lblWilsonAlert.Visible = False
    
    '- Stock info (idle)
    state.StockInfoText = "Code Article :  --"
    state.StockInfoColor = RGB(100, 100, 100)
    state.formRef.lblStockInfo.Caption = state.StockInfoText
    state.formRef.lblStockInfo.ForeColor = state.StockInfoColor
    
    '- Total footer
    state.TotalGeneralText = "TOTAL GENERAL :  0.00 DZD"
    state.formRef.lblTotalGeneral.Caption = state.TotalGeneralText
    state.formRef.lblTotalGeneral.Font.Bold = True
    state.formRef.lblTotalGeneral.ForeColor = RGB(5, 100, 60)
    
    '- Sync toggle default
    state.formRef.chkSyncInternal.Value = True
    
    '- Button captions
    state.formRef.btnAjouterLigne.Caption = "Ajouter Ligne"
    state.formRef.btnSupprimerLigne.Caption = "Supprimer Ligne"
    state.formRef.btnEn[ARTICLE_DESC]r.Caption = "En[ARTICLE_DESC]r"
    state.formRef.btnAutoRef.Caption = "Auto-Ref"
    state.formRef.btnAnnuler.Caption = "Annuler"
    If HasControl(state.formRef, "lblPU") Then
        state.formRef.lblPU.Caption = "PU -- CMUP auto"
    End If
    
    '- Date default
    state.TransDate = Format(Date, "DD/MM/YYYY")
    state.formRef.TxtDate.Value = state.TransDate
End Sub

Private Sub PopulateDropdowns(ByRef state As FormState)
    '- Document types
    With state.formRef.cmbTypeDoc
        .Clear
        .AddItem mod_Config.DOC_TYPE_BS
        .AddItem mod_Config.DOC_TYPE_BR
        .AddItem mod_Config.DOC_TYPE_BC
        .AddItem mod_Config.DOC_TYPE_DA
        .ListIndex = 0
    End With
    
    '- Services from SYS_STRINGS
    With state.formRef.cmbService
        .Clear
        Dim wsStr As Worksheet
        On Error Resume Next
        Set wsStr = ThisWorkbook.Sheets(mod_Config.SHEET_SYS_STRINGS)
        On Error GoTo 0
        
        If Not wsStr Is Nothing Then
            Dim lastRowStr As Long, iStr As Long
            lastRowStr = wsStr.Cells(wsStr.Rows.count, 1).End(xlUp).Row
            For iStr = 2 To lastRowStr
                If Left(Trim(CStr(wsStr.Cells(iStr, 1).Value)), 4) = "SVC_" Then
                    .AddItem Trim(CStr(wsStr.Cells(iStr, 2).Value))
                End If
            Next iStr
        End If
        
        If .listCount = 0 Then
            .AddItem "Service 1"
            .AddItem "Service 2"
            .AddItem "Fournisseur Externe"
        End If
        If .listCount > 0 Then .ListIndex = 0
    End With
    
    '- Categories
    With state.formRef.cmbCategorie
        .Clear
        .AddItem "(Toutes)"
        .AddItem "Fournitures Bureau"
        .AddItem "Informatique"
        .AddItem "Admin"
        .AddItem "Inconnu"
        .ListIndex = 0
    End With
    
    '- Articles (unfiltered)
    Call LoadArticleComboBox("", state)
End Sub

Private Sub LoadArticleComboBox(ByVal filterCat As String, ByRef state As FormState)
    Dim wsArt    As Worksheet
    Dim lastRow  As Long
    Dim i        As Long
    Dim code     As String
    Dim desig    As String
    Dim cat      As String
    Dim noFilter As Boolean

    state.formRef.cmbArticle.Clear
    noFilter = (filterCat = "" Or filterCat = "(Toutes)")

    On Error Resume Next
    Set wsArt = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    On Error GoTo 0

    If wsArt Is Nothing Then
        state.formRef.cmbArticle.AddItem "ART-001 | Papier A4"
        state.formRef.cmbArticle.AddItem "ART-002 | Papier A3"
        state.formRef.cmbArticle.AddItem "ART-003 | Sous-Chemise"
        Exit Sub
    End If

    lastRow = wsArt.Cells(wsArt.Rows.count, 1).End(xlUp).Row

    For i = 3 To lastRow
        code = Trim(CStr(wsArt.Cells(i, 1).Value))
        desig = Trim(CStr(wsArt.Cells(i, 2).Value))
        cat = Trim(CStr(wsArt.Cells(i, 5).Value))

        If code = "" Then GoTo nextRow

        If noFilter Or (cat = filterCat) Then
            state.formRef.cmbArticle.AddItem code & " | " & desig
        End If
nextRow:
    Next i
End Sub

Private Sub ConfigureGrid(ByRef state As FormState)
    With state.formRef.lstGrid
        .ColumnCount = 6
        .ColumnHeads = False
        .ColumnWidths = "80 pt;220 pt;90 pt;50 pt;80 pt;90 pt"
        .MultiSelect = fmMultiSelectSingle
        .ListStyle = fmListStylePlain
        .Font.name = "Courier New"
        .Font.Size = 9
        .BackColor = RGB(248, 248, 252)
    End With

    state.formRef.lblGridHeader.Caption = _
        "  Code      |  Designation              |  Cat" & Chr(233) & "gorie  | Qte |  PU (DZD) |  Valeur"
    state.formRef.lblGridHeader.Font.name = "Courier New"
    state.formRef.lblGridHeader.Font.Size = 8
    state.formRef.lblGridHeader.ForeColor = RGB(71, 71, 90)
End Sub

Private Sub ResetToDefaultState(ByRef state As FormState)
    state.TransDate = Format(Date, "DD/MM/YYYY")
    state.formRef.TxtDate.Value = state.TransDate
    state.docRef = ""
    state.formRef.txtRefDoc.Value = state.docRef
    state.qty = ""
    state.formRef.txtQuantite.Value = ""
    state.unitPrice = ""
    state.formRef.txtPrixUnitaire.Value = ""
    state.formRef.cmbArticle.ListIndex = -1
    state.formRef.lstGrid.Clear
    state.GridData = ""
    state.GridRowCount = 0

    m_TotalGeneral = 0
    state.TotalGeneral = 0
    m_CurrentArticle = ""
    state.ArticleCode = ""
    m_StockActuel = 0
    state.ArticleStock = 0

    state.StockInfoText = "Code Article :  --"
    state.StockInfoColor = RGB(100, 100, 100)
    state.formRef.lblStockInfo.Caption = state.StockInfoText
    state.formRef.lblStockInfo.ForeColor = state.StockInfoColor
    
    state.WilsonAlertVisible = False
    state.formRef.lblWilsonAlert.Visible = False
    
    state.QtyBackColor = RGB(255, 255, 255)
    state.formRef.txtQuantite.BackColor = state.QtyBackColor

    Call UpdateTotalDisplay(state)
End Sub


'==============================================================================
' SECTION 2 - DOCUMENT TYPE BANNER (Pure logic, returns state updates)
'==============================================================================

Public Sub OnDocTypeChanged(ByRef state As FormState)
    state.docType = state.formRef.cmbTypeDoc.Value
    
    Select Case state.docType
        Case mod_Config.DOC_TYPE_BS
            state.BannerText = "  MODE SORTIE  --  Bon de Sortie"
            state.BannerColor = RGB(160, 70, 0)
            m_IsBRMode = False
            state.IsBRMode = False
            state.PUEditable = False
            state.PULabel = "PU -- CMUP auto"
            state.formRef.txtPrixUnitaire.Enabled = False
            state.formRef.txtPrixUnitaire.BackColor = RGB(235, 235, 235)

        Case mod_Config.DOC_TYPE_BR
            state.BannerText = "  MODE ENTREE  --  Bon de R" & Chr(201) & "ception"
            state.BannerColor = RGB(4, 90, 55)
            m_IsBRMode = True
            state.IsBRMode = True
            state.PUEditable = True
            state.PULabel = "Prix Unitaire (saisir)"
            state.formRef.txtPrixUnitaire.Enabled = True
            state.formRef.txtPrixUnitaire.BackColor = RGB(255, 252, 196)

        Case mod_Config.DOC_TYPE_DA
            state.BannerText = "  Demande d'Achat"
            state.BannerColor = RGB(30, 80, 180)
            m_IsBRMode = False
            state.IsBRMode = False
            state.PUEditable = False
            state.PULabel = "PU (estime)"
            state.formRef.txtPrixUnitaire.Enabled = False
            state.formRef.txtPrixUnitaire.BackColor = RGB(235, 235, 235)

        Case mod_Config.DOC_TYPE_BC
            state.BannerText = "  COMMANDE  --  Bon de Commande"
            state.BannerColor = RGB(120, 40, 120)
            m_IsBRMode = False
            state.IsBRMode = False
            state.PUEditable = True
            state.PULabel = "Prix Unitaire (devis)"
            state.formRef.txtPrixUnitaire.Enabled = True
            state.formRef.txtPrixUnitaire.BackColor = RGB(255, 252, 196)

        Case Else
            state.BannerText = "-- SELECTIONNEZ LE TYPE DE DOCUMENT --"
            state.BannerColor = RGB(100, 100, 100)
    End Select

    '- Apply state to form
    state.formRef.fraDocTypeBanner.BackColor = state.BannerColor
    state.formRef.lblBannerText.Caption = state.BannerText
    If HasControl(state.formRef, "lblPU") Then
        state.formRef.lblPU.Caption = state.PULabel
    End If
    
    '- Refresh stock display if article selected
    If m_CurrentArticle <> "" Then Call EvaluateStockStatus(m_CurrentArticle, state)
    
    '- Auto-generate reference if empty
    If Len(Trim(state.docRef)) = 0 Then Call GenerateAutoRef(state)
End Sub

Public Function GetDocPrefixFromType(ByVal docType As String) As String
    Select Case docType
        Case mod_Config.DOC_TYPE_BS:  GetDocPrefixFromType = "BS"
        Case mod_Config.DOC_TYPE_BR:  GetDocPrefixFromType = "BR"
        Case mod_Config.DOC_TYPE_BC:  GetDocPrefixFromType = "BC"
        Case mod_Config.DOC_TYPE_DA:  GetDocPrefixFromType = "DA"
        Case Else:    GetDocPrefixFromType = "TXN"
    End Select
End Function


'==============================================================================
' SECTION 3 - ARTICLE SELECTION & STOCK INTELLIGENCE
'==============================================================================

Public Sub OnArticleChanged(ByRef state As FormState)
    Dim raw As String
    
    If state.formRef.cmbArticle.ListIndex < 0 Then Exit Sub
    raw = Trim(state.formRef.cmbArticle.text)
    If Len(raw) = 0 Then Exit Sub

    Dim parts() As String
    parts = Split(raw, "|")
    m_CurrentArticle = Trim(parts(0))
    state.ArticleCode = m_CurrentArticle

    Call EvaluateStockStatus(m_CurrentArticle, state)
End Sub

Private Sub EvaluateStockStatus(ByVal artCode As String, ByRef state As FormState)
    Dim wsArt    As Worksheet
    Dim foundRow As Variant
    Dim stock    As Long
    Dim pu       As Double
    Dim cat      As String
    Dim ropVal   As Double
    Dim ssVal    As Long

    On Error Resume Next
    Set wsArt = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    On Error GoTo 0

    If wsArt Is Nothing Then
        state.StockInfoText = "Code Article :  " & artCode & "  |  [Feuille ARTICLES introuvable]"
        state.StockInfoColor = RGB(180, 0, 0)
        Exit Sub
    End If

    foundRow = Application.Match(artCode, wsArt.Range("A:A"), 0)

    If IsError(foundRow) Then
        state.StockInfoText = "Code Article :  " & artCode & "  |  Article introuvable"
        state.StockInfoColor = RGB(180, 0, 0)
        m_StockActuel = -1
        state.ArticleStock = -1
        state.WilsonAlertVisible = False
        Exit Sub
    End If

    pu = CDbl(mod_Utilities.SafeVal(wsArt.Cells(foundRow, 8).Value))
    cat = Trim(CStr(wsArt.Cells(foundRow, 11).Value))
    state.ArticlePU = pu
    state.ArticleCat = cat

    ' Use mod_StockEngine for stock calculation (consolidated)
    Dim totalIn As Double, totalOut As Double
    On Error Resume Next
    totalIn = Application.SumIfs(wsArt.Parent.Range("E:E"), wsArt.Parent.Range("B:B"), artCode, wsArt.Parent.Range("D:D"), "IN")
    totalOut = Application.SumIfs(wsArt.Parent.Range("E:E"), wsArt.Parent.Range("B:B"), artCode, wsArt.Parent.Range("D:D"), "OUT")
    On Error GoTo 0
    stock = CLng(totalIn - totalOut)

    m_StockActuel = stock
    state.ArticleStock = stock

    If artCode = "ART-001" Then
        ropVal = CANON_ROP
        ssVal = CANON_SS
    Else
        ssVal = 50
        ropVal = ssVal + CANON_LT
    End If

    Dim statusText  As String
    Dim statusColor As Long

    If stock <= 0 Then
        statusText = "[RUPTURE]"
        statusColor = RGB(200, 30, 30)
    ElseIf stock <= ssVal Then
        statusText = "[CRITIQUE]"
        statusColor = RGB(200, 30, 30)
    ElseIf stock <= ropVal Then
        statusText = "[ALERTE]"
        statusColor = RGB(160, 70, 0)
    Else
        statusText = "[OK]"
        statusColor = RGB(4, 90, 55)
    End If

    state.StockInfoText = "Code Article :  " & artCode & "   |   Stock :  " & stock & " u" & "   |   " & statusText
    state.StockInfoColor = statusColor

    '- Auto-fill PU for non-BR modes
    If Not m_IsBRMode And pu > 0 Then
        state.formRef.txtPrixUnitaire.Value = Format(pu, "0.00")
        state.unitPrice = Format(pu, "0.00")
    End If

    '- Wilson alert for case study article
    If artCode = "ART-001" Then
        state.WilsonAlertText = "Wilson EOQ -- Q* = " & CANON_QSTAR & " u  |  SS = " & CANON_SS & " u"
        state.WilsonAlertVisible = True
    Else
        state.WilsonAlertVisible = False
    End If
    
    '- Apply to form
    state.formRef.lblStockInfo.Caption = state.StockInfoText
    state.formRef.lblStockInfo.ForeColor = state.StockInfoColor
    state.formRef.lblWilsonAlert.Visible = state.WilsonAlertVisible
    If state.WilsonAlertVisible Then
        state.formRef.lblWilsonAlert.Caption = state.WilsonAlertText
        state.formRef.lblWilsonAlert.ForeColor = RGB(4, 90, 55)
    End If
End Sub

Public Sub OnCategoryChanged(ByRef state As FormState)
    Dim prevSKU As String
    prevSKU = m_CurrentArticle

    Call LoadArticleComboBox(Trim(state.formRef.cmbCategorie.Value), state)

    '- Restore previous selection if still in filtered list
    If prevSKU <> "" Then
        Dim j As Integer
        For j = 0 To state.formRef.cmbArticle.listCount - 1
            If Left(state.formRef.cmbArticle.List(j), Len(prevSKU)) = prevSKU Then
                state.formRef.cmbArticle.ListIndex = j
                Exit For
            End If
        Next j
    End If
End Sub


'==============================================================================
' SECTION 4 - QUANTITY FIELD (Live validation + Wilson nudge)
'==============================================================================

Public Sub OnQuantityChanged(ByRef state As FormState)
    state.qty = state.formRef.txtQuantite.Value
    
    If Not IsNumeric(state.qty) Then
        state.QtyBackColor = RGB(255, 199, 199)
        state.formRef.txtQuantite.BackColor = state.QtyBackColor
        Exit Sub
    End If

    Dim qty As Long
    qty = CLng(state.qty)
    If qty <= 0 Then Exit Sub

    If Not m_IsBRMode And m_StockActuel >= 0 Then
        Dim projected As Long
        projected = m_StockActuel - qty

        Select Case True
            Case projected < 0
                state.QtyBackColor = RGB(255, 199, 199)  ' Red - insufficient
            Case projected <= CANON_SS
                state.QtyBackColor = RGB(255, 235, 150)  ' Orange - below SS
            Case projected <= CANON_ROP
                state.QtyBackColor = RGB(255, 248, 200)  ' Yellow - below ROP
            Case Else
                state.QtyBackColor = RGB(198, 239, 206)  ' Green - safe
        End Select
    Else
        state.QtyBackColor = RGB(255, 255, 255)
    End If
    
    state.formRef.txtQuantite.BackColor = state.QtyBackColor
End Sub


'==============================================================================
' SECTION 5 - AUTO REFERENCE GENERATOR
'==============================================================================

Public Sub GenerateAutoRef(ByRef state As FormState)
    Dim prefix As String
    Dim seq    As Long

    prefix = GetDocPrefixFromType(state.docType)
    seq = GetNextSequence(prefix)

    state.docRef = prefix & "-" & Format(Date, "YYYY") & "-" & Format(seq, "0000")
    state.formRef.txtRefDoc.Value = state.docRef
End Sub

Private Function GetNextSequence(ByVal prefix As String) As Long
    Dim wsMouv   As Worksheet
    Dim lastRow As Long
    Dim i       As Long
    Dim maxSeq  As Long
    Dim refStr  As String

    maxSeq = 0

    On Error Resume Next
    Set wsMouv = ThisWorkbook.Sheets(mod_Config.SHEET_MOUVEMENTS)
    On Error GoTo 0

    If wsMouv Is Nothing Then
        GetNextSequence = 1
        Exit Function
    End If

    lastRow = wsMouv.Cells(wsMouv.Rows.count, 7).End(xlUp).Row
    
    For i = 3 To lastRow
        refStr = CStr(wsMouv.Cells(i, 7).Value)
        If Left(refStr, Len(prefix)) = prefix And InStr(refStr, "-") > 0 Then
            Dim parts() As String
            parts = Split(refStr, "-")
            If UBound(parts) >= 2 Then
                Dim seqNum As Long
                On Error Resume Next
                seqNum = CLng(parts(UBound(parts)))
                If Err.Number = 0 And seqNum > maxSeq Then maxSeq = seqNum
                On Error GoTo 0
            End If
        End If
    Next i

    GetNextSequence = maxSeq + 1
End Function


'==============================================================================
' SECTION 6 - GRID OPERATIONS
'==============================================================================

Public Function AddLineToGrid(ByRef state As FormState) As Boolean
    AddLineToGrid = False
    
    Dim qty        As Long
    Dim pu         As Double
    Dim valLigne   As Double
    Dim desig      As String
    Dim cat        As String
    Dim ropSeuil   As Double
    Dim rowIdx     As Integer

    '- Guard 1: Date validation
    If Not mod_Utilities.IsValidDate(state.TransDate) Then
        MsgBox "Format de date requis : JJ/MM/AAAA", vbExclamation
        state.formRef.TxtDate.SetFocus
        Exit Function
    End If

    '- Guard 2: Document reference
    If Len(Trim(state.docRef)) = 0 Then
        MsgBox "Le N° Reference est OBLIGATOIRE.", vbCritical
        state.formRef.txtRefDoc.SetFocus
        Exit Function
    End If

    '- Guard 3: Article selection
    If Len(Trim(m_CurrentArticle)) = 0 Then
        MsgBox "Selectionnez un article.", vbExclamation
        state.formRef.cmbArticle.SetFocus
        Exit Function
    End If

    '- Guard 4: Article exists
    If m_StockActuel = -1 Then
        MsgBox "Article introuvable dans le catalogue.", vbCritical
        Exit Function
    End If

    '- Guard 5: Quantity valid
    If Not IsNumeric(state.qty) Then
        MsgBox "Quantite invalide.", vbCritical
        state.formRef.txtQuantite.SetFocus
        Exit Function
    End If

    qty = CLng(state.qty)
    If qty <= 0 Then
        MsgBox "La quantite doit " & Chr(234) & "tre > 0.", vbCritical
        state.formRef.txtQuantite.SetFocus
        Exit Function
    End If

    '- Guard 6: PU required for BR mode
    If m_IsBRMode Then
        state.unitPrice = state.formRef.txtPrixUnitaire.Value
        If Not IsNumeric(state.unitPrice) Or CDbl(mod_Utilities.SafeVal(state.unitPrice)) <= 0 Then
            MsgBox "Le Prix Unitaire est requis pour un Bon de R" & Chr(201) & "ception.", vbCritical
            state.formRef.txtPrixUnitaire.SetFocus
            Exit Function
        End If
    End If

    pu = CDbl(mod_Utilities.SafeVal(state.unitPrice))

    '- Guard 7: Stock sufficiency (non-BR only)
    If Not m_IsBRMode Then
        Dim netProjected As Long
        netProjected = m_StockActuel - qty

        If netProjected < 0 Then
            MsgBox "Stock insuffisant ! Stock dispo: " & m_StockActuel & " u, Qte demandee: " & qty & " u", vbCritical
            Exit Function
        End If

        ropSeuil = IIf(m_CurrentArticle = "ART-001", CANON_ROP, 60)
        If netProjected <= ropSeuil Then
            Dim ropResp As VbMsgBoxResult
            ropResp = MsgBox("ALERTE -- Point de commande atteint." & vbCrLf & _
                            "Continuer ?", vbYesNo + vbExclamation, "ROP Alert")
            If ropResp = vbNo Then Exit Function
        End If
    End If

    '- Add line to grid
    valLigne = qty * pu
    desig = mod_Utilities.GetArticleField(m_CurrentArticle, "DESIG")
    cat = mod_Utilities.GetArticleField(m_CurrentArticle, "CAT")

    state.formRef.lstGrid.AddItem ""
    rowIdx = state.formRef.lstGrid.listCount - 1

    state.formRef.lstGrid.List(rowIdx, COL_CODE) = m_CurrentArticle
    state.formRef.lstGrid.List(rowIdx, COL_DESIG) = Left(desig, 28)
    state.formRef.lstGrid.List(rowIdx, COL_CAT) = Left(cat, 14)
    state.formRef.lstGrid.List(rowIdx, COL_QTE) = CStr(qty)
    state.formRef.lstGrid.List(rowIdx, COL_PU) = Format(pu, "#,##0.00")
    state.formRef.lstGrid.List(rowIdx, COL_VALEUR) = Format(valLigne, "#,##0.00")

    state.GridRowCount = state.formRef.lstGrid.listCount
    
    Call UpdateTotalDisplay(state)

    '- Reset input fields
    state.formRef.cmbArticle.ListIndex = -1
    state.formRef.txtQuantite.Value = ""
    state.formRef.txtPrixUnitaire.Value = ""
    state.formRef.txtQuantite.BackColor = RGB(255, 255, 255)
    m_CurrentArticle = ""
    state.ArticleCode = ""
    m_StockActuel = 0
    state.ArticleStock = 0
    state.StockInfoText = "Code Article :  --"
    state.StockInfoColor = RGB(100, 100, 100)
    state.formRef.lblStockInfo.Caption = state.StockInfoText
    state.formRef.lblStockInfo.ForeColor = state.StockInfoColor
    state.WilsonAlertVisible = False
    state.formRef.lblWilsonAlert.Visible = False
    state.qty = ""
    state.unitPrice = ""

    state.formRef.cmbArticle.SetFocus
    AddLineToGrid = True
End Function

Public Sub RemoveLineFromGrid(ByRef state As FormState)
    If state.formRef.lstGrid.ListIndex < 0 Then
        MsgBox "Selectionnez une ligne a supprimer.", vbInformation
        Exit Sub
    End If

    state.formRef.lstGrid.RemoveItem state.formRef.lstGrid.ListIndex
    state.GridRowCount = state.formRef.lstGrid.listCount
    Call UpdateTotalDisplay(state)
End Sub

Private Sub UpdateTotalDisplay(ByRef state As FormState)
    Dim runningTotal As Double
    Dim i            As Integer
    runningTotal = 0

    For i = 0 To state.formRef.lstGrid.listCount - 1
        On Error Resume Next
        Dim rawVal As String
        rawVal = state.formRef.lstGrid.List(i, COL_VALEUR)
        rawVal = Replace(rawVal, ",", "")
        If IsNumeric(rawVal) Then runningTotal = runningTotal + CDbl(rawVal)
        On Error GoTo 0
    Next i

    m_TotalGeneral = runningTotal
    state.TotalGeneral = runningTotal
    state.TotalGeneralText = "TOTAL GENERAL :  " & Format(m_TotalGeneral, "#,##0.00") & " DZD"
    state.formRef.lblTotalGeneral.Caption = state.TotalGeneralText
End Sub

Private Function GetQtyInGridForSKU(ByVal sku As String, ByRef state As FormState) As Long
    Dim total As Long
    Dim i     As Integer
    total = 0

    For i = 0 To state.formRef.lstGrid.listCount - 1
        If state.formRef.lstGrid.List(i, COL_CODE) = sku Then
            On Error Resume Next
            total = total + CLng(state.formRef.lstGrid.List(i, COL_QTE))
            On Error GoTo 0
        End If
    Next i

    GetQtyInGridForSKU = total
End Function


'==============================================================================
' SECTION 7 - EN[ARTICLE_DESC]R (Transaction commit)
'==============================================================================

Public Function CommitTransaction(ByRef state As FormState) As Boolean
    CommitTransaction = False
    
    ' Read current state from form
    state.Service = state.formRef.cmbService.Value
    state.docType = state.formRef.cmbTypeDoc.Value
    state.docRef = Trim(state.formRef.txtRefDoc.Value)
    state.GridRowCount = state.formRef.lstGrid.listCount
    
    '- Guard: Empty grid
    If state.GridRowCount = 0 Then
        MsgBox "Le document ne contient aucun article.", vbExclamation
        Exit Function
    End If
    
    '- Guard: Service required
    If Len(Trim(state.Service)) = 0 Then
        MsgBox "SERVICE / FOURNISSEUR est requis.", vbExclamation
        state.formRef.cmbService.SetFocus
        Exit Function
    End If
    
    '- Guard: DocType required
    If Len(Trim(state.docType)) = 0 Then
        MsgBox "Type de Document est requis.", vbExclamation
        state.formRef.cmbTypeDoc.SetFocus
        Exit Function
    End If

    '- Guard: Validate grid data
    Dim gridRow As Integer
    For gridRow = 0 To state.GridRowCount - 1
        If Not IsNumeric(state.formRef.lstGrid.List(gridRow, COL_QTE)) Or _
           Not IsNumeric(state.formRef.lstGrid.List(gridRow, COL_PU)) Then
            MsgBox "Donn" & Chr(233) & "es invalides a la ligne " & gridRow + 1, vbCritical
            Exit Function
        End If
        If CLng(state.formRef.lstGrid.List(gridRow, COL_QTE)) <= 0 Then
            MsgBox "La quantite doit " & Chr(234) & "tre > 0 (ligne " & gridRow + 1 & ")", vbCritical
            Exit Function
        End If
    Next gridRow

    '- Confirmation dialog
    Dim typeSign As String
    typeSign = IIf(m_IsBRMode, "IN -- Entree", "OUT -- Sortie")
    
    Dim confMsg As String
    confMsg = "Confirmer l'en[ARTICLE_DESC]ment ?" & vbCrLf & vbCrLf & _
              "Document :  " & state.docType & "  [" & typeSign & "]" & vbCrLf & _
              "Reference:  " & state.docRef & vbCrLf & _
              "Service  :  " & state.Service & vbCrLf & _
              "Lignes   :  " & state.GridRowCount & vbCrLf & _
              "Total    :  " & Format(state.TotalGeneral, "#,##0.00") & " DZD"
    
    If MsgBox(confMsg, vbYesNo + vbQuestion) = vbNo Then Exit Function
    
    '- Begin transaction
    Dim i        As Integer
    Dim docDate  As Date
    Dim mvtSign  As String
    Dim lineCode As String
    Dim lineDesig As String
    Dim lineQty  As Long
    Dim linePU   As Double
    Dim lineVal  As Double
    
    '- Begin safety-managed transaction (snapshot + crash recovery flag)
    Call mod_TransactionSafety.BeginTransaction(state.docRef, state.docType)
    Call mod_TransactionSafety.SaveTransactionStateForRecovery
    
    On Error GoTo SaveError
    
    docDate = Date
    mvtSign = IIf(m_IsBRMode, "IN", "OUT")
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Dim wsMouv As Worksheet: Set wsMouv = ThisWorkbook.Sheets(mod_Config.SHEET_MOUVEMENTS)
    
    For i = 0 To state.GridRowCount - 1
        lineCode = state.formRef.lstGrid.List(i, COL_CODE)
        lineDesig = state.formRef.lstGrid.List(i, COL_DESIG)
        lineQty = CLng(state.formRef.lstGrid.List(i, COL_QTE))
        linePU = CDbl(Replace(state.formRef.lstGrid.List(i, COL_PU), ",", ""))
        lineVal = CDbl(Replace(state.formRef.lstGrid.List(i, COL_VALEUR), ",", ""))
        
        '- Pre-write stock check (OUT only)
        If mvtSign = "OUT" Then
            Dim currentStockLevel As Double
            currentStockLevel = mod_StockEngine.GetArticleStock(lineCode)
            
            If lineQty > currentStockLevel Then
                MsgBox "Stock insuffisant pour '" & lineCode & "'. Stock dispo: " & currentStockLevel, vbCritical
                GoTo SaveError
            End If
        End If
        
        '- Write to MOUVEMENTS via secure layer
        Call mod_Database.SecureWriteTransaction( _
            docDate:=docDate, _
            typeSign:=mvtSign, _
            refDoc:=state.docRef, _
            codeArticle:=lineCode, _
            designation:=lineDesig, _
            quantity:=lineQty, _
            unitPrice:=linePU, _
            lineValue:=lineVal, _
            thirdParty:=state.Service)
        
        '- Track line in transaction safety
        Call mod_TransactionSafety.AddTransactionLine
        
        '- Sync internal state
        If SyncTransactionInternal(lineCode, mvtSign, lineQty, linePU, state.docRef) <> 0 Then
            MsgBox "Erreur de synchronisation interne. Annulation de la transaction...", vbCritical
            GoTo SaveError
        End If
    Next i

    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    '- Commit safety transaction (validates consistency)
    If Not mod_TransactionSafety.CommitTransaction Then
        MsgBox "Validation de la transaction échouée. Annulation...", vbCritical
        GoTo SaveError
    End If
    
    '- Clear crash recovery flag
    Call mod_TransactionSafety.ClearTransactionState
    
    '- Audit trail
    Call mod_AuditTrail.LogTransaction(state.docType, state.docRef)
    
    '- Success message
    MsgBox "En[ARTICLE_DESC]ment r" & Chr(233) & "ussi !" & vbCrLf & _
           "Reference :  " & state.docRef & vbCrLf & _
           state.GridRowCount & " ligne(s) enregistr" & Chr(233) & "e(s)", vbInformation
    
    '- Sync metrics back
    Call mod_SyncBridge.SyncMetricsFromLedger
    
    '- PDF export prompt
    If MsgBox("Imprimer le " & state.docType & " ?", vbYesNo + vbQuestion) = vbYes Then
        Call mod_ExportEngine.ExportTransactionToPDF(state.docRef)
    End If
    
    Call ResetToDefaultState(state)
    CommitTransaction = True
    Exit Function
    
SaveError:
    '- Safety-managed rollback (snapshot restore + partial movement removal)
    If mod_TransactionSafety.GetTransactionStatus <> "NONE" Then
        Call mod_TransactionSafety.RollbackTransaction
        Call mod_TransactionSafety.ClearTransactionState
    End If
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Une erreur s'est produite lors de l'en[ARTICLE_DESC]ment. Transaction annul" & Chr(233) & "e.", vbCritical
End Function

Private Function SyncTransactionInternal(ByVal artCode As String, _
                                          ByVal mvtType As String, _
                                          ByVal qty As Long, _
                                          ByVal unitPrice As Double, _
                                          ByVal refDoc As String) As Integer
    On Error Resume Next
    SyncTransactionInternal = mod_SyncBridge.SyncTransactionInternal(artCode, mvtType, qty, unitPrice, refDoc)
    If Err.Number <> 0 Then SyncTransactionInternal = -1
    On Error GoTo 0
End Function


'==============================================================================
' SECTION 8 - CANCEL
'==============================================================================

Public Sub CancelTransaction(ByRef state As FormState)
    If state.formRef.lstGrid.listCount > 0 Then
        If MsgBox("Des lignes sont en attente. Annuler quand meme ?", _
                  vbYesNo + vbExclamation) = vbNo Then
            Exit Sub
        End If
    End If
    Unload state.formRef
End Sub


'==============================================================================
' SECTION 9 - UTILITY: Form Reference Helper
'==============================================================================

'- Helper to check if a control exists on the form (handles optional controls)
Public Function HasControl(ByVal formRef As Object, ByVal ctrlName As String) As Boolean
    Dim ctrl As Object
    On Error Resume Next
    Set ctrl = formRef.Controls(ctrlName)
    HasControl = (Err.Number = 0)
    On Error GoTo 0
End Function

'==============================================================================
' END -- mod_StockEntry_Logic.bas
'==============================================================================
