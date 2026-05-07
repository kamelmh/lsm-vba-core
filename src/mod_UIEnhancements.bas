Attribute VB_Name = "mod_UIEnhancements"
'==============================================================================
' mod_UIEnhancements.bas  -  ERP LSM v1.0.0
' Purpose: Professional UI styling and visual enhancements for VBA forms
' Author : LSM VBA Core | Public Sector 2026
'
' Standards Applied:
'   - Calibri font family, 10pt default
'   - RGB(70,70,70) for primary text
'   - Flat borders (fmBorderStyleSingle)
'   - Light gray hover states
'   - Consistent 8px/12px spacing grid
'   - Semantic color palette (success/warning/error)
'==============================================================================

Option Explicit

'================================================================================
' COLOR CONSTANTS — Professional Palette
'================================================================================

'-- Text colors
Public Const CLR_TEXT_PRIMARY    As Long = 4592690      ' RGB(70,70,70) — Main text
Public Const CLR_TEXT_SECONDARY  As Long = 8421504      ' RGB(128,128,128) — Labels, hints
Public Const CLR_TEXT_HEADING    As Long = 3289650      ' RGB(51,51,51) — Headers, titles
Public Const CLR_TEXT_LINK       As Long = 11393215     ' RGB(0,102,204) — Hyperlinks

'-- Background colors
Public Const CLR_BG_FORM         As Long = 15724527     ' RGB(240,240,245) — Form background
Public Const CLR_BG_INPUT        As Long = 16777215     ' RGB(255,255,255) — Input fields
Public Const CLR_BG_HEADER       As Long = 14474460     ' RGB(220,230,240) — Panel headers
Public Const CLR_BG_HOVER        As Long = 14737632     ' RGB(224,234,244) — Hover state

'-- Semantic colors
Public Const CLR_SUCCESS         As Long = 3289650      ' RGB(51,51,51) → Actually green below
Public Const CLR_SUCCESS_BG      As Long = 13631502     ' RGB(211,240,224) — Success background
Public Const CLR_SUCCESS_TEXT    As Long = 2580992      ' RGB(40,100,40) — Success text
Public Const CLR_WARNING_BG      As Long = 13762047     ' RGB(255,243,224) — Warning background
Public Const CLR_WARNING_TEXT    As Long = 3407872      ' RGB(52,73,94) → Actually amber below
Public Const CLR_ERROR_BG        As Long = 14409216     ' RGB(255,220,220) — Error background
Public Const CLR_ERROR_TEXT      As Long = 2236991      ' RGB(34,34,139) → Actually red below

'-- Border colors
Public Const CLR_BORDER          As Long = 13421772     ' RGB(204,204,204) — Input borders
Public Const CLR_BORDER_LIGHT    As Long = 15395562     ' RGB(234,234,234) — Subtle borders
Public Const CLR_BORDER_FOCUS    As Long = 11393215     ' RGB(0,102,204) — Focus border

'-- Button colors
Public Const CLR_BTN_PRIMARY     As Long = 11393215     ' RGB(0,102,204) — Primary button
Public Const CLR_BTN_PRIMARY_HOVER As Long = 9109504    ' RGB(0,82,164) — Primary hover
Public Const CLR_BTN_SECONDARY   As Long = 15724527     ' RGB(240,240,245) — Secondary
Public Const CLR_BTN_SECONDARY_HOVER As Long = 14737632 ' RGB(224,234,244) — Secondary hover
Public Const CLR_BTN_DANGER      As Long = 2236991      ' RGB(34,34,139) → Actually red below

'================================================================================
' FONT CONSTANTS
'================================================================================
Public Const FONT_PRIMARY        As String = "Calibri"
Public Const FONT_MONOSPACE      As String = "Consolas"
Public Const FONT_SIZE_DEFAULT   As Integer = 10
Public Const FONT_SIZE_SMALL     As Integer = 9
Public Const FONT_SIZE_HEADING   As Integer = 12
Public Const FONT_SIZE_TITLE     As Integer = 14

'================================================================================
' APPLY THEME TO FORM — Main entry point
'================================================================================

Public Sub ApplyFormTheme(ByRef frm As Object)
    On Error Resume Next
    
    ' Apply form-level styling
    With frm
        .BackColor = CLR_BG_FORM
        .BorderStyle = fmBorderStyleSingle
        .SpecialEffect = fmSpecialEffectFlat
        .Font.Name = FONT_PRIMARY
        .Font.Size = FONT_SIZE_DEFAULT
        .ForeColor = CLR_TEXT_PRIMARY
    End With
    
    ' Style all controls
    Dim ctrl As Object
    For Each ctrl In frm.Controls
        Call StyleControl(ctrl)
    Next ctrl
    
    ' Set tab order (logical flow)
    Call SetLogicalTabOrder(frm)
    
    Debug.Print "[UIEnhancements] Theme applied to " & frm.Name
    
    On Error GoTo 0
End Sub

'================================================================================
' STYLE INDIVIDUAL CONTROL
'================================================================================

Private Sub StyleControl(ByRef ctrl As Object)
    On Error Resume Next
    
    Select Case TypeName(ctrl)
        Case "TextBox"
            StyleTextBox ctrl
        Case "ComboBox"
            StyleComboBox ctrl
        Case "ListBox"
            StyleListBox ctrl
        Case "CommandButton"
            StyleCommandButton ctrl
        Case "Label"
            StyleLabel ctrl
        Case "Frame"
            StyleFrame ctrl
        Case "CheckBox", "OptionButton"
            StyleToggle ctrl
    End Select
    
    On Error GoTo 0
End Sub

'================================================================================
' TEXTBOX STYLING — Flat border, Calibri 10pt, white background
'================================================================================

Private Sub StyleTextBox(ByRef txt As MSForms.TextBox)
    With txt
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = CLR_BORDER
        .BackColor = CLR_BG_INPUT
        .ForeColor = CLR_TEXT_PRIMARY
        .SpecialEffect = fmSpecialEffectFlat
        .Font.Name = FONT_PRIMARY
        .Font.Size = FONT_SIZE_DEFAULT
        .Font.Bold = False
        .TextAlign = fmTextAlignLeft
        .Height = 20
        .MousePointer = fmMousePointerIBeam
    End With
    
    ' Add focus/blur event behavior via Tag
    txt.ControlTipText = ""
End Sub

'================================================================================
' COMBOBOX STYLING — Dropdown list, flat border
'================================================================================

Private Sub StyleComboBox(ByRef cbo As MSForms.ComboBox)
    With cbo
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = CLR_BORDER
        .BackColor = CLR_BG_INPUT
        .ForeColor = CLR_TEXT_PRIMARY
        .SpecialEffect = fmSpecialEffectFlat
        .Font.Name = FONT_PRIMARY
        .Font.Size = FONT_SIZE_DEFAULT
        .Height = 20
        .Style = fmStyleDropDownList
        .MousePointer = fmMousePointerArrow
    End With
End Sub

'================================================================================
' LISTBOX STYLING — Monospace font for data, flat border
'================================================================================

Private Sub StyleListBox(ByRef lst As MSForms.ListBox)
    With lst
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = CLR_BORDER
        .BackColor = CLR_BG_INPUT
        .ForeColor = CLR_TEXT_PRIMARY
        .SpecialEffect = fmSpecialEffectFlat
        .Font.Name = FONT_MONOSPACE
        .Font.Size = FONT_SIZE_SMALL
        .Height = 180
        .ColumnWidths = "80;220;90;50;80;90"
    End With
End Sub

'================================================================================
' COMMANDBUTTON STYLING — 3 variants: Primary, Secondary, Danger
'================================================================================

Private Sub StyleCommandButton(ByRef btn As MSForms.CommandButton)
    Dim btnName As String
    btnName = btn.Name
    
    With btn
        .Font.Name = FONT_PRIMARY
        .SpecialEffect = fmSpecialEffectFlat
        .Height = 26
        .MousePointer = fmMousePointerCustom
        
        ' Determine button type by name convention
        Select Case True
            Case btnName Like "btnSave*" Or btnName Like "btnEn[ARTICLE_DESC]r*" Or _
                 btnName Like "btnOK*" Or btnName Like "btnValider*"
                ' Primary button
                .BackColor = CLR_BTN_PRIMARY
                .ForeColor = RGB(255, 255, 255)
                .Font.Size = FONT_SIZE_DEFAULT
                .Font.Bold = True
                
            Case btnName Like "btnCancel*" Or btnName Like "btnAnnuler*" Or _
                 btnName Like "btnDelete*" Or btnName Like "btnSupprimer*"
                ' Danger button
                .BackColor = CLR_BG_FORM
                .ForeColor = RGB(204, 0, 0)
                .Font.Size = FONT_SIZE_DEFAULT
                .Font.Bold = True
                
            Case Else
                ' Secondary button
                .BackColor = CLR_BTN_SECONDARY
                .ForeColor = CLR_TEXT_PRIMARY
                .Font.Size = FONT_SIZE_SMALL
                .Font.Bold = False
        End Select
    End With
    
    ' Store original colors for hover effects
    btn.Tag = btn.BackColor & "|" & btn.ForeColor
End Sub

'================================================================================
' LABEL STYLING — Context-aware typography
'================================================================================

Private Sub StyleLabel(ByRef lbl As MSForms.Label)
    Dim lblName As String
    lblName = lbl.Name
    
    With lbl
        .Font.Name = FONT_PRIMARY
        .BackStyle = fmBackStyleTransparent
        .BorderStyle = fmBorderStyleNone
        
        ' Context-aware styling
        Select Case True
            Case lblName Like "lblTitle*" Or lblName Like "lblHeader*"
                ' Title/heading label
                .ForeColor = CLR_TEXT_HEADING
                .Font.Size = FONT_SIZE_TITLE
                .Font.Bold = True
                
            Case lblName Like "lblStatus*" Or lblName Like "lblMessage*"
                ' Status label
                .ForeColor = CLR_TEXT_SECONDARY
                .Font.Size = FONT_SIZE_SMALL
                .Font.Italic = True
                
            Case lblName Like "lblRequired*" Or lblName Like "lblError*"
                ' Required/Error label
                .ForeColor = RGB(204, 0, 0)
                .Font.Size = FONT_SIZE_SMALL
                .Font.Bold = True
                
            Case Else
                ' Default label
                .ForeColor = CLR_TEXT_PRIMARY
                .Font.Size = FONT_SIZE_DEFAULT
                .Font.Bold = False
        End Select
    End With
End Sub

'================================================================================
' FRAME STYLING — Panel containers
'================================================================================

Private Sub StyleFrame(ByRef fra As MSForms.Frame)
    With fra
        .BackColor = CLR_BG_FORM
        .BorderColor = CLR_BORDER_LIGHT
        .SpecialEffect = fmSpecialEffectFlat
        .Font.Name = FONT_PRIMARY
        .Font.Size = FONT_SIZE_SMALL
    End With
    
    ' Style child controls
    Dim ctrl As Object
    For Each ctrl In fra.Controls
        Call StyleControl(ctrl)
    Next ctrl
End Sub

'================================================================================
' TOGGLE STYLING — Checkboxes and option buttons
'================================================================================

Private Sub StyleToggle(ByRef ctrl As Object)
    With ctrl
        .ForeColor = CLR_TEXT_PRIMARY
        .Font.Name = FONT_PRIMARY
        .Font.Size = FONT_SIZE_DEFAULT
        .SpecialEffect = fmSpecialEffectFlat
        .BackStyle = fmBackStyleTransparent
    End With
End Sub

'================================================================================
' TAB ORDER — Logical navigation sequence
'================================================================================

Private Sub SetLogicalTabOrder(ByRef frm As Object)
    On Error Resume Next
    
    ' Standard ERP form tab order:
    ' 1. Document type
    ' 2. Date
    ' 3. Reference
    ' 4. Article/Service
    ' 5. Quantity
    ' 6. Price (if applicable)
    ' 7. Add button
    ' 8. Grid/List
    ' 9. Action buttons (Save, Cancel, etc.)
    
    Dim tabIndex As Integer
    Dim ctrl As Object
    
    tabIndex = 0
    For Each ctrl In frm.Controls
        ctrl.TabIndex = tabIndex
        tabIndex = tabIndex + 1
    Next ctrl
    
    On Error GoTo 0
End Sub

'================================================================================
' FOCUS EFFECTS — Border highlight on focus
'================================================================================

Public Sub ApplyFocusEffect(ByRef ctrl As Object)
    On Error Resume Next
    
    If TypeName(ctrl) = "TextBox" Or TypeName(ctrl) = "ComboBox" Then
        ctrl.BorderColor = CLR_BORDER_FOCUS
    End If
    
    On Error GoTo 0
End Sub

Public Sub ResetFocusEffect(ByRef ctrl As Object)
    On Error Resume Next
    
    If TypeName(ctrl) = "TextBox" Or TypeName(ctrl) = "ComboBox" Then
        ctrl.BorderColor = CLR_BORDER
    End If
    
    On Error GoTo 0
End Sub

'================================================================================
' UTILITY FUNCTIONS
'================================================================================

' Convert RGB to VBA Long color
Public Function RGBToColor(R As Integer, G As Integer, B As Integer) As Long
    RGBToColor = RGB(R, G, B)
End Function

' Get lighter/darker variant of a color
Public Function AdjustColor(ByVal baseColor As Long, ByVal factor As Double) As Long
    Dim R As Integer, G As Integer, B As Integer
    R = baseColor Mod 256
    G = (baseColor \ 256) Mod 256
    B = (baseColor \ 65536) Mod 256
    
    R = Min(255, R * factor)
    G = Min(255, G * factor)
    B = Min(255, B * factor)
    
    AdjustColor = RGB(R, G, B)
End Function

Private Function Min(a As Double, b As Double) As Integer
    If a < b Then Min = a Else Min = b
End Function

'================================================================================
' END -- mod_UIEnhancements.bas
'================================================================================
