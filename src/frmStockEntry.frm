VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmStockEntry 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmStockEntry.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmStockEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'- Module-level state
Private m_State As FormState
Private m_Initialized As Boolean

'==============================================================================
' FORM LIFECYCLE
'==============================================================================

Private Sub UserForm_Initialize()
    '- Build all controls programmatically
    Call BuildUI
    
    '- Initialize FormState struct
    Set m_State.formRef = Me
    m_State.IsBRMode = False
    
    '- Delegate to controller
    Call mod_StockEntry_Logic.InitializeForm(m_State)
    
    m_Initialized = True
End Sub

'==============================================================================
' PROGRAMMATIC UI BUILDER
' Creates all 25 controls at runtime - no designer needed
'==============================================================================

Private Sub BuildUI()
    Dim ctrl As Object
    Dim i As Integer
    
    '- Set form dimensions
    Me.Width = 870
    Me.Height = 640
    Me.Caption = "ERP LSM - Saisie des Mouvements"
    Me.StartUpPosition = 1  ' CenterOwner
    
    '--------------------------------------------------------------------------
    ' 1. DOCUMENT TYPE BANNER (Frame + Label)
    '--------------------------------------------------------------------------
    Set ctrl = Me.Controls.Add("Forms.Frame.1", "fraDocTypeBanner", True)
    With ctrl
        .Left = 10
        .Top = 10
        .Width = 840
        .Height = 50
        .BackColor = RGB(100, 100, 100)
        .BorderStyle = 0
    End With
    
    Set ctrl = Me.Controls.Add("Forms.Label.1", "lblBannerText", True)
    With ctrl
        .Left = 20
        .Top = 18
        .Width = 820
        .Height = 35
        .Caption = "-- SELECTIONNEZ LE TYPE DE DOCUMENT --"
        .ForeColor = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 14
        .Font.name = "Tahoma"
        .TextAlign = fmTextAlignCenter
        .BackStyle = fmBackStyleTransparent
    End With
    
    '--------------------------------------------------------------------------
    ' 2. DOCUMENT TYPE DROPDOWN
    '--------------------------------------------------------------------------
    Set ctrl = Me.Controls.Add("Forms.ComboBox.1", "cmbTypeDoc", True)
    With ctrl
        .Left = 10
        .Top = 70
        .Width = 200
        .Height = 24
        .Font.Size = 9
        .Font.name = "Tahoma"
        .Style = fmStyleDropDownList
    End With
    
    '--------------------------------------------------------------------------
    ' 3. SERVICE DROPDOWN
    '--------------------------------------------------------------------------
    Set ctrl = Me.Controls.Add("Forms.ComboBox.1", "cmbService", True)
    With ctrl
        .Left = 220
        .Top = 70
        .Width = 200
        .Height = 24
        .Font.Size = 9
        .Font.name = "Tahoma"
        .Style = fmStyleDropDownList
    End With
    
    '--------------------------------------------------------------------------
    ' 4. CATEGORY FILTER
    '--------------------------------------------------------------------------
    Set ctrl = Me.Controls.Add("Forms.ComboBox.1", "cmbCategorie", True)
    With ctrl
        .Left = 430
        .Top = 70
        .Width = 150
        .Height = 24
        .Font.Size = 9
        .Font.name = "Tahoma"
        .Style = fmStyleDropDownList
    End With
    
    '--------------------------------------------------------------------------
    ' 5. DATE FIELD
    '--------------------------------------------------------------------------
    Set ctrl = Me.Controls.Add("Forms.TextBox.1", "TxtDate", True)
    With ctrl
        .Left = 590
        .Top = 70
        .Width = 100
        .Height = 24
        .Font.Size = 9
        .Font.name = "Tahoma"
        .TextAlign = fmTextAlignCenter
        .BackColor = RGB(240, 240, 240)
        .Locked = True
    End With
    
    '--------------------------------------------------------------------------
    ' 6. AUTO-REF BUTTON
    '--------------------------------------------------------------------------
    Set ctrl = Me.Controls.Add("Forms.CommandButton.1", "btnAutoRef", True)
    With ctrl
        .Left = 700
        .Top = 70
        .Width = 150
        .Height = 24
        .Caption = "Auto-Ref"
        .Font.Size = 9
        .Font.name = "Tahoma"
        .Font.Bold = True
    End With
    
    '--------------------------------------------------------------------------
    ' 7. ARTICLE SELECTION
    '--------------------------------------------------------------------------
    Set ctrl = Me.Controls.Add("Forms.ComboBox.1", "cmbArticle", True)
    With ctrl
        .Left = 10
        .Top = 110
        .Width = 500
        .Height = 24
        .Font.Size = 9
        .Font.name = "Tahoma"
        .Style = fmStyleDropDownList
    End With
    
    '--------------------------------------------------------------------------
    ' 8. STOCK INFO LABEL
    '--------------------------------------------------------------------------
    Set ctrl = Me.Controls.Add("Forms.Label.1", "lblStockInfo", True)
    With ctrl
        .Left = 520
        .Top = 110
        .Width = 330
        .Height = 24
        .Caption = "Code Article :  --"
        .Font.Size = 9
        .Font.name = "Tahoma"
        .ForeColor = RGB(100, 100, 100)
    End With
    
    '--------------------------------------------------------------------------
    ' 9. WILSON ALERT LABEL
    '--------------------------------------------------------------------------
    Set ctrl = Me.Controls.Add("Forms.Label.1", "lblWilsonAlert", True)
    With ctrl
        .Left = 10
        .Top = 140
        .Width = 840
        .Height = 20
        .Caption = ""
        .Font.Size = 8
        .Font.name = "Tahoma"
        .Font.Bold = True
        .ForeColor = RGB(4, 90, 55)
        .Visible = False
    End With
    
    '--------------------------------------------------------------------------
    ' 10. GRID HEADER LABEL
    '--------------------------------------------------------------------------
    Set ctrl = Me.Controls.Add("Forms.Label.1", "lblGridHeader", True)
    With ctrl
        .Left = 10
        .Top = 170
        .Width = 610
        .Height = 20
        .Caption = "  Code  |  Designation  |  Categorie  | Qte |  PU (DZD) |  Valeur"
        .Font.name = "Courier New"
        .Font.Size = 8
        .ForeColor = RGB(71, 71, 90)
        .BackStyle = fmBackStyleTransparent
    End With
    
    '--------------------------------------------------------------------------
    ' 11. TRANSACTION GRID (ListBox)
    '--------------------------------------------------------------------------
    Set ctrl = Me.Controls.Add("Forms.ListBox.1", "lstGrid", True)
    With ctrl
        .Left = 10
        .Top = 195
        .Width = 610
        .Height = 250
        .Font.name = "Courier New"
        .Font.Size = 9
        .BackColor = RGB(248, 248, 252)
        .ColumnCount = 6
        .ColumnHeads = False
        .ColumnWidths = "80;220;90;50;80;90"
        .MultiSelect = fmMultiSelectSingle
        .ListStyle = fmListStylePlain
    End With
    
    '--------------------------------------------------------------------------
    ' 12. QUANTITY FIELD
    '--------------------------------------------------------------------------
    Set ctrl = Me.Controls.Add("Forms.TextBox.1", "txtQuantite", True)
    With ctrl
        .Left = 630
        .Top = 195
        .Width = 100
        .Height = 24
        .Font.Size = 11
        .Font.name = "Tahoma"
        .TextAlign = fmTextAlignRight
        .BackColor = RGB(255, 255, 255)
    End With
    
    '--------------------------------------------------------------------------
    ' 13. UNIT PRICE FIELD
    '--------------------------------------------------------------------------
    Set ctrl = Me.Controls.Add("Forms.TextBox.1", "txtPrixUnitaire", True)
    With ctrl
        .Left = 630
        .Top = 230
        .Width = 100
        .Height = 24
        .Font.Size = 11
        .Font.name = "Tahoma"
        .TextAlign = fmTextAlignRight
        .Enabled = False
        .BackColor = RGB(235, 235, 235)
    End With
    
    '--------------------------------------------------------------------------
    ' 14. PU LABEL
    '--------------------------------------------------------------------------
    Set ctrl = Me.Controls.Add("Forms.Label.1", "lblPU", True)
    With ctrl
        .Left = 740
        .Top = 232
        .Width = 110
        .Height = 20
        .Caption = "PU -- CMUP auto"
        .Font.Size = 8
        .Font.name = "Tahoma"
        .ForeColor = RGB(128, 128, 128)
    End With
    
    '--------------------------------------------------------------------------
    ' 15. DOCUMENT REFERENCE
    '--------------------------------------------------------------------------
    Set ctrl = Me.Controls.Add("Forms.TextBox.1", "txtRefDoc", True)
    With ctrl
        .Left = 630
        .Top = 265
        .Width = 220
        .Height = 24
        .Font.Size = 9
        .Font.name = "Courier New"
        .TextAlign = fmTextAlignCenter
        .BackColor = RGB(255, 252, 196)
        .Font.Bold = True
    End With
    
    '--------------------------------------------------------------------------
    ' 16. AJOUTER LIGNE BUTTON
    '--------------------------------------------------------------------------
    Set ctrl = Me.Controls.Add("Forms.CommandButton.1", "btnAjouterLigne", True)
    With ctrl
        .Left = 630
        .Top = 300
        .Width = 100
        .Height = 28
        .Caption = "+ Ajouter"
        .Font.Size = 9
        .Font.name = "Tahoma"
        .Font.Bold = True
        .BackColor = RGB(198, 239, 206)
    End With
    
    '--------------------------------------------------------------------------
    ' 17. SUPPRIMER LIGNE BUTTON
    '--------------------------------------------------------------------------
    Set ctrl = Me.Controls.Add("Forms.CommandButton.1", "btnSupprimerLigne", True)
    With ctrl
        .Left = 740
        .Top = 300
        .Width = 110
        .Height = 28
        .Caption = "- Supprimer"
        .Font.Size = 9
        .Font.name = "Tahoma"
    End With
    
    '--------------------------------------------------------------------------
    ' 18. TOTAL GENERAL LABEL
    '--------------------------------------------------------------------------
    Set ctrl = Me.Controls.Add("Forms.Label.1", "lblTotalGeneral", True)
    With ctrl
        .Left = 10
        .Top = 450
        .Width = 610
        .Height = 28
        .Caption = "TOTAL GENERAL :  0.00 DZD"
        .Font.Size = 12
        .Font.name = "Tahoma"
        .Font.Bold = True
        .ForeColor = RGB(5, 100, 60)
    End With
    
    '--------------------------------------------------------------------------
    ' 19. EN[ARTICLE_DESC]R BUTTON (primary action)
    '--------------------------------------------------------------------------
    Set ctrl = Me.Controls.Add("Forms.CommandButton.1", "btnEn[ARTICLE_DESC]r", True)
    With ctrl
        .Left = 630
        .Top = 450
        .Width = 110
        .Height = 32
        .Caption = "En[ARTICLE_DESC]r"
        .Font.Size = 10
        .Font.name = "Tahoma"
        .Font.Bold = True
        .BackColor = RGB(0, 102, 204)
        .ForeColor = RGB(255, 255, 255)
    End With
    
    '--------------------------------------------------------------------------
    ' 20. ANNULER BUTTON
    '--------------------------------------------------------------------------
    Set ctrl = Me.Controls.Add("Forms.CommandButton.1", "btnAnnuler", True)
    With ctrl
        .Left = 750
        .Top = 450
        .Width = 100
        .Height = 32
        .Caption = "Annuler"
        .Font.Size = 10
        .Font.name = "Tahoma"
    End With
    
    '--------------------------------------------------------------------------
    ' 21. SYNC MASTER DATA BUTTON
    '--------------------------------------------------------------------------
    Set ctrl = Me.Controls.Add("Forms.CommandButton.1", "btnSyncMasterData", True)
    With ctrl
        .Left = 10
        .Top = 490
        .Width = 180
        .Height = 24
        .Caption = "Sync Master Data"
        .Font.Size = 8
        .Font.name = "Tahoma"
    End With
    
    '--------------------------------------------------------------------------
    ' 22. GENERATE REPORT BUTTON
    '--------------------------------------------------------------------------
    Set ctrl = Me.Controls.Add("Forms.CommandButton.1", "btnGenerateReport", True)
    With ctrl
        .Left = 200
        .Top = 490
        .Width = 150
        .Height = 24
        .Caption = "Generate Report"
        .Font.Size = 8
        .Font.name = "Tahoma"
    End With
    
    '--------------------------------------------------------------------------
    ' 23. IMPRIMER BON BUTTON
    '--------------------------------------------------------------------------
    Set ctrl = Me.Controls.Add("Forms.CommandButton.1", "btnImprimerBon", True)
    With ctrl
        .Left = 360
        .Top = 490
        .Width = 150
        .Height = 24
        .Caption = "Imprimer Bon"
        .Font.Size = 8
        .Font.name = "Tahoma"
    End With
    
    '--------------------------------------------------------------------------
    ' 24. SYNC INTERNAL CHECKBOX
    '--------------------------------------------------------------------------
    Set ctrl = Me.Controls.Add("Forms.CheckBox.1", "chkSyncInternal", True)
    With ctrl
        .Left = 520
        .Top = 492
        .Width = 180
        .Height = 20
        .Caption = "Sync interne automatique"
        .Font.Size = 8
        .Font.name = "Tahoma"
        .Value = True
    End With
    
    '--------------------------------------------------------------------------
    ' 25. SEPARATOR LINE (using Label)
    '--------------------------------------------------------------------------
    Set ctrl = Me.Controls.Add("Forms.Label.1", "lblSeparator", True)
    With ctrl
        .Left = 10
        .Top = 485
        .Width = 840
        .Height = 2
        .BackStyle = fmBackStyleOpaque
        .BackColor = RGB(200, 200, 200)
    End With
    
    Debug.Print "BuildUI complete - " & Me.Controls.count & " controls created"
    
    ' Apply professional theme
    Call ApplyProfessionalTheme
End Sub

'==============================================================================
' PROFESSIONAL THEME APPLICATION
'==============================================================================

Private Sub ApplyProfessionalTheme()
    On Error Resume Next
    
    ' Form-level styling
    Me.BackColor = RGB(245, 245, 250)
    
    ' Banner
    With Me.Controls("fraDocTypeBanner")
        .BackColor = RGB(0, 70, 127)
        .BorderStyle = 0
    End With
    
    With Me.Controls("lblBannerText")
        .Font.Name = "Calibri"
        .Font.Size = 14
        .Font.Bold = True
        .ForeColor = RGB(255, 255, 255)
    End With
    
    ' Text inputs
    ApplyTextboxTheme "TxtDate"
    ApplyTextboxTheme "txtRefDoc"
    ApplyTextboxTheme "txtQuantite"
    ApplyTextboxTheme "txtPrixUnitaire"
    
    ' Comboboxes
    ApplyComboboxTheme "cmbTypeDoc"
    ApplyComboboxTheme "cmbArticle"
    ApplyComboboxTheme "cmbService"
    ApplyComboboxTheme "cmbCategorie"
    
    ' Listbox
    With Me.Controls("lstGrid")
        .BorderStyle = 1
        .BorderColor = RGB(192, 192, 192)
        .BackColor = RGB(255, 255, 255)
        .Font.Name = "Consolas"
        .Font.Size = 9
    End With
    
    ' Labels
    ApplyLabelTheme "lblStockInfo", RGB(70, 70, 70), 9, False
    ApplyLabelTheme "lblWilsonAlert", RGB(40, 100, 40), 8, True
    ApplyLabelTheme "lblGridHeader", RGB(100, 100, 100), 8, False
    ApplyLabelTheme "lblTotalGeneral", RGB(5, 100, 60), 12, True
    ApplyLabelTheme "lblPU", RGB(128, 128, 128), 8, False
    
    ' Buttons
    With Me.Controls("btnEn[ARTICLE_DESC]r")
        .BackColor = RGB(0, 102, 204)
        .ForeColor = RGB(255, 255, 255)
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Bold = True
        .BorderStyle = 0
    End With
    
    With Me.Controls("btnAjouterLigne")
        .BackColor = RGB(232, 245, 233)
        .ForeColor = RGB(40, 100, 40)
        .Font.Name = "Calibri"
        .Font.Size = 10
        .Font.Bold = True
        .BorderStyle = 0
    End With
    
    With Me.Controls("btnSupprimerLigne")
        .BackColor = RGB(252, 228, 236)
        .ForeColor = RGB(204, 0, 0)
        .Font.Name = "Calibri"
        .Font.Size = 10
        .Font.Bold = True
        .BorderStyle = 0
    End With
    
    With Me.Controls("btnAnnuler")
        .BackColor = RGB(245, 245, 250)
        .ForeColor = RGB(70, 70, 70)
        .Font.Name = "Calibri"
        .Font.Size = 10
        .BorderStyle = 0
    End With
    
    With Me.Controls("btnAutoRef")
        .BackColor = RGB(245, 245, 250)
        .ForeColor = RGB(0, 102, 204)
        .Font.Name = "Calibri"
        .Font.Size = 9
        .Font.Bold = True
        .BorderStyle = 0
    End With
    
    With Me.Controls("btnImprimerBon")
        .BackColor = RGB(245, 245, 250)
        .ForeColor = RGB(0, 102, 204)
        .Font.Name = "Calibri"
        .Font.Size = 9
        .BorderStyle = 0
    End With
    
    With Me.Controls("btnSyncMasterData")
        .BackColor = RGB(245, 245, 250)
        .ForeColor = RGB(70, 70, 70)
        .Font.Name = "Calibri"
        .Font.Size = 8
        .BorderStyle = 0
    End With
    
    With Me.Controls("btnGenerateReport")
        .BackColor = RGB(245, 245, 250)
        .ForeColor = RGB(70, 70, 70)
        .Font.Name = "Calibri"
        .Font.Size = 8
        .BorderStyle = 0
    End With
    
    ' Checkbox
    With Me.Controls("chkSyncInternal")
        .Font.Name = "Calibri"
        .Font.Size = 9
        .ForeColor = RGB(70, 70, 70)
    End With
    
    ' Separator
    With Me.Controls("lblSeparator")
        .BackColor = RGB(224, 224, 224)
    End With
    
    Debug.Print "Professional theme applied"
End Sub

Private Sub ApplyTextboxTheme(ByVal ctrlName As String)
    On Error Resume Next
    With Me.Controls(ctrlName)
        .BorderStyle = 1
        .BorderColor = RGB(192, 192, 192)
        .BackColor = RGB(255, 255, 255)
        .ForeColor = RGB(70, 70, 70)
        .Font.Name = "Calibri"
        .Font.Size = 10
    End With
End Sub

Private Sub ApplyComboboxTheme(ByVal ctrlName As String)
    On Error Resume Next
    With Me.Controls(ctrlName)
        .BorderStyle = 1
        .BorderColor = RGB(192, 192, 192)
        .BackColor = RGB(255, 255, 255)
        .ForeColor = RGB(70, 70, 70)
        .Font.Name = "Calibri"
        .Font.Size = 10
    End With
End Sub

Private Sub ApplyLabelTheme(ByVal ctrlName As String, ByVal foreColor As Long, ByVal fontSize As Integer, ByVal isBold As Boolean)
    On Error Resume Next
    With Me.Controls(ctrlName)
        .ForeColor = foreColor
        .Font.Name = "Calibri"
        .Font.Size = fontSize
        .Font.Bold = isBold
    End With
End Sub


'==============================================================================
' EVENT HANDLERS (All delegate to controller)
'==============================================================================

'-- Document type changed
Private Sub cmbTypeDoc_Change()
    If Not m_Initialized Then Exit Sub
    Call mod_StockEntry_Logic.OnDocTypeChanged(m_State)
End Sub

'-- Article selection changed
Private Sub cmbArticle_Change()
    If Not m_Initialized Then Exit Sub
    Call mod_StockEntry_Logic.OnArticleChanged(m_State)
End Sub

'-- Category filter changed
Private Sub cmbCategorie_Change()
    If Not m_Initialized Then Exit Sub
    Call mod_StockEntry_Logic.OnCategoryChanged(m_State)
End Sub

'-- Quantity field changed (live validation)
Private Sub txtQuantite_Change()
    If Not m_Initialized Then Exit Sub
    Call mod_StockEntry_Logic.OnQuantityChanged(m_State)
End Sub

'-- Auto-generate reference
Private Sub btnAutoRef_Click()
    Call mod_StockEntry_Logic.GenerateAutoRef(m_State)
End Sub

'-- Add line to grid
Private Sub btnAjouterLigne_Click()
    Call mod_StockEntry_Logic.AddLineToGrid(m_State)
End Sub

'-- Remove line from grid
Private Sub btnSupprimerLigne_Click()
    Call mod_StockEntry_Logic.RemoveLineFromGrid(m_State)
End Sub

'-- En[ARTICLE_DESC]r (commit transaction)
Private Sub btnEn[ARTICLE_DESC]r_Click()
    Call mod_StockEntry_Logic.CommitTransaction(m_State)
End Sub

'-- Cancel / Annuler
Private Sub btnAnnuler_Click()
    Call mod_StockEntry_Logic.CancelTransaction(m_State)
End Sub

'-- Sync master data
Private Sub btnSyncMasterData_Click()
    On Error Resume Next
    Call mod_SyncBridge.SyncMetricsFromLedger
    If Err.Number = 0 Then
        MsgBox "Synchronisation réussie.", vbInformation, "Sync Master"
    Else
        MsgBox "Erreur: " & Err.Description, vbCritical, "Sync Error"
    End If
    On Error GoTo 0
End Sub

'-- Generate report
Private Sub btnGenerateReport_Click()
    Call mod_Procurement.GenerateOrderReport
End Sub

'-- Print bon
Private Sub btnImprimerBon_Click()
    Dim docRef As String
    docRef = Trim(Me.Controls("txtRefDoc").Value)
    If Len(docRef) > 0 Then
        Call mod_ExportEngine.ExportTransactionToPDF(docRef)
    Else
        MsgBox "Veuillez générer une référence d'abord.", vbExclamation
    End If
End Sub


'==============================================================================
' KEYBOARD SHORTCUTS
'==============================================================================

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, _
                              ByVal Shift As Integer)
    Select Case KeyCode
        Case 27: Call btnAnnuler_Click
        Case 13: Call btnAjouterLigne_Click
    End Select
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Call btnAnnuler_Click
    End If
End Sub

'==============================================================================
' END -- frmStockEntry.frm (Programmatic Controls - 450 lines)
' All business logic in mod_StockEntry_Logic.bas (996 lines)
'==============================================================================


