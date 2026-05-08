Attribute VB_Name = "mod_SheetSetup"
Option Explicit

' Sheet protection password delegated to mod_Config.MASTER_PWD
' W003 FIX: Eliminated local duplicate to prevent password drift

Public Sub CreateAllSheets()
    Dim confirmation As VbMsgBoxResult
    
    confirmation = MsgBox("This will setup/verify all ERP sheets." & vbCrLf & _
                          "Existing data will be preserved." & vbCrLf & vbCrLf & _
                          "Continue?", vbYesNo + vbQuestion, "ERP v13 Sheet Setup")
    
    If confirmation <> vbYes Then Exit Sub
    
    CreateArticlesSheet
    CreateMouvementsSheet
    CreateFormInputSheet
    CreateStagingBufferSheet
    CreateSysStringsSheet
    CreateReceiptTagSheet
    CreateTemplateBonSheet
    
    MsgBox "ERP v13 sheets ready!", vbInformation, "Setup Complete"
End Sub

Private Sub CreateArticlesSheet()
    Dim ws As Worksheet
    Dim sheetExists As Boolean
    
    sheetExists = False
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("ARTICLES")
    If Err.Number = 0 Then sheetExists = True
    On Error GoTo 0
    
    If Not sheetExists Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        ws.name = "ARTICLES"
        
        With ws
            .Range("A1").Value = "CODE_ARTICLE"
            .Range("B1").Value = "DESIGNATION_AR"
            .Range("C1").Value = "QTE_STOCK"
            .Range("D1").Value = "PRIX_UNITAIRE"
            .Range("E1").Value = "CATEGORIE"
            .Range("F1").Value = "MIN_STOCK"
            .Range("A1:F1").Font.Bold = True
            .Range("A1:F1").Interior.Color = RGB(44, 62, 80)
            .Range("A1:F1").Font.Color = RGB(255, 255, 255)
            .Range("A1:F1").HorizontalAlignment = xlCenter
        End With
    End If
    
    On Error Resume Next
    ws.Unprotect Password:=mod_Config.MASTER_PWD
    On Error GoTo 0
    
    With ws
        .Tab.Color = RGB(39, 174, 96)
        .Columns("A").ColumnWidth = 12
        .Columns("B").ColumnWidth = 30
        .Columns("C").ColumnWidth = 12
        .Columns("D").ColumnWidth = 15
        .Columns("E").ColumnWidth = 18
        .Columns("F").ColumnWidth = 12
        .Protect mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    End With
End Sub

Private Sub CreateMouvementsSheet()
    Dim ws As Worksheet
    Dim sheetExists As Boolean
    
    sheetExists = False
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("MOUVEMENTS")
    If Err.Number = 0 Then sheetExists = True
    On Error GoTo 0
    
    If Not sheetExists Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        ws.name = "MOUVEMENTS"
        
        With ws
            .Range("A1").Value = "DATE"
            .Range("B1").Value = "CODE_ARTICLE"
            .Range("C1").Value = "DESIGNATION"
            .Range("D1").Value = "TYPE_MVT"
            .Range("E1").Value = "QTE"
            .Range("F1").Value = "REF_DOCUMENT"
            .Range("G1").Value = "TXN_ID"
            .Range("A1:G1").Font.Bold = True
            .Range("A1:G1").Interior.Color = RGB(44, 62, 80)
            .Range("A1:G1").Font.Color = RGB(255, 255, 255)
        End With
    End If
    
    On Error Resume Next
    ws.Unprotect Password:=mod_Config.MASTER_PWD
    On Error GoTo 0
    
    With ws
        .Tab.Color = RGB(52, 152, 219)
        .Protect mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    End With
End Sub

Private Sub CreateFormInputSheet()
    Dim ws As Worksheet
    Dim sheetExists As Boolean
    
    sheetExists = False
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("FORM_INPUT")
    If Err.Number = 0 Then sheetExists = True
    On Error GoTo 0
    
    If Not sheetExists Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        ws.name = "FORM_INPUT"
        
        With ws
            .Range("A1:D1").Merge
            .Range("A1").Value = "????? ???? ??????? - Entry Form"
            .Range("A1").Font.Bold = True
            .Range("A1").Font.Size = 14
            .Range("A3").Value = "??? ?????? (SKU):"
            .Range("B3").Value = ""
            .Range("A4").Value = "??? ?????? (IN/OUT):"
            .Range("B4").Value = "IN"
            .Range("A5").Value = "?????? (Quantity):"
            .Range("B5").Value = 0
            .Range("A6").Value = "???? ???????:"
            .Range("B6").Value = ""
        End With
    End If
    
    On Error Resume Next
    ws.Unprotect Password:=mod_Config.MASTER_PWD
    On Error GoTo 0
    
    With ws
        .Tab.Color = RGB(155, 89, 182)
        .Protect mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    End With
End Sub

Private Sub CreateStagingBufferSheet()
    Dim ws As Worksheet
    Dim sheetExists As Boolean
    
    sheetExists = False
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("STAGING_BUFFER")
    If Err.Number = 0 Then sheetExists = True
    On Error GoTo 0
    
    Application.DisplayAlerts = False
    If sheetExists Then ws.Delete
    Application.DisplayAlerts = True
    
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    ws.name = "STAGING_BUFFER"
    
    With ws
        .Range("A1").Value = "TXN_ID"
        .Range("B1").Value = "TIMESTAMP"
        .Range("C1").Value = "SKU"
        .Range("D1").Value = "TYPE"
        .Range("E1").Value = "QTY"
        .Range("F1").Value = "STATUS"
        .Range("A1:F1").Font.Bold = True
        .Tab.Color = RGB(230, 126, 34)
        .Protect mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    End With
End Sub

Private Sub CreateSysStringsSheet()
    Dim ws As Worksheet
    Dim sheetExists As Boolean
    
    sheetExists = False
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("SYS_STRINGS")
    If Err.Number = 0 Then
        sheetExists = True
        ws.Unprotect Password:=mod_Config.MASTER_PWD
        ws.Visible = xlSheetVisible
        ws.Delete
    End If
    On Error GoTo 0
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    ws.name = "SYS_STRINGS"
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    With ws
        .Range("A1").Value = "STRING_ID"
        .Range("B1").Value = "ARABIC_TEXT"
        .Range("A1:B1").Font.Bold = True
        
        .Range("A2").Value = "SYS_TITLE"
        .Range("B2").Value = "???? ================================================================================ - ????? ???????"
        
        .Range("A3").Value = "ERR_SYNC_ABSENT"
        .Range("B3").Value = "??? Python ??? ?????"
        
        .Range("A4").Value = "ERR_ART_NOT_FOUND"
        .Range("B4").Value = "?????? ??? ??????"
        
        .Range("A5").Value = "MSG_SUCCESS"
        .Range("B5").Value = "?? ????? ?????"
        
        .Range("A6").Value = "BTN_SAVE"
        .Range("B6").Value = "???"
        
        .Range("A7").Value = "BTN_CANCEL"
        .Range("B7").Value = "?????"
        
        .Range("A8").Value = "BTN_REFRESH"
        .Range("B8").Value = "?????"
        
        .Range("A9").Value = "LBL_STOCK"
        .Range("B9").Value = "???????"
        
        .Range("A10").Value = "LBL_EOF"
        .Range("B10").Value = "????? ???????"
        
        .Range("A11").Value = "LBL_ARTICLE"
        .Range("B11").Value = "??????"
        
        .Range("A12").Value = "LBL_QUANTITY"
        .Range("B12").Value = "??????"
        
        .Range("A13").Value = "LBL_UNIT_PRICE"
        .Range("B13").Value = "????? ??????"
        
        .Range("A14").Value = "LBL_DATE"
        .Range("B14").Value = "???????"
        
        .Range("A15").Value = "LBL_TYPE"
        .Range("B15").Value = "?????"
        
        .Range("A16").Value = "MSG_CONFIRM_SAVE"
        .Range("B16").Value = "?? ???? ??????"
        
        .Range("A17").Value = "MSG_LOADING"
        .Range("B17").Value = "???? ???????..."
        
        .Range("A18").Value = "ERR_INVALID_DATA"
        .Range("B18").Value = "?????? ??? ?????"
        
        ' === INVENTORY MANAGEMENT ===
        .Range("A19").Value = "MENU_STOCK"
        .Range("B19").Value = "????? ???????"
        
        .Range("A20").Value = "MENU_MOUVEMENTS"
        .Range("B20").Value = "???????"
        
        .Range("A21").Value = "MENU_ARTICLES"
        .Range("B21").Value = "??????"
        
        .Range("A22").Value = "MENU_DASHBOARD"
        .Range("B22").Value = "???? ??????"
        
        .Range("A23").Value = "MENU_REPORTS"
        .Range("B23").Value = "????????"
        
        .Range("A24").Value = "MENU_SETTINGS"
        .Range("B24").Value = "?????????"
        
        ' === STOCK TYPES ===
        .Range("A25").Value = "TYPE_ENTRY"
        .Range("B25").Value = "?????"
        
        .Range("A26").Value = "TYPE_EXIT"
        .Range("B26").Value = "?????"
        
        .Range("A27").Value = "TYPE_TRANSFER"
        .Range("B27").Value = "?????"
        
        .Range("A28").Value = "TYPE_ADJUST"
        .Range("B28").Value = "?????"
        
        ' === DASHBOARD ===
        .Range("A29").Value = "DASH_TOTAL_ITEMS"
        .Range("B29").Value = "?????? ??????"
        
        .Range("A30").Value = "DASH_TOTAL_VALUE"
        .Range("B30").Value = "?????? ?????????"
        
        .Range("A31").Value = "DASH_LOW_STOCK"
        .Range("B31").Value = "????? ?????"
        
        .Range("A32").Value = "DASH_OUT_OF_STOCK"
        .Range("B32").Value = "???? ???????"
        
        .Range("A33").Value = "DASH_NORMAL"
        .Range("B33").Value = "????? ?????"
        
        ' === ALERTS ===
        .Range("A34").Value = "ALERT_LOW_STOCK"
        .Range("B34").Value = "?????: ??????? ?????"
        
        .Range("A35").Value = "ALERT_OUT_OF_STOCK"
        .Range("B35").Value = "?????: ???? ???????"
        
        .Range("A36").Value = "ALERT_REORDER"
        .Range("B36").Value = "?????: reached ???? ?????"
        
        ' === REPORTS ===
        .Range("A37").Value = "REPORT_MONTHLY"
        .Range("B37").Value = "??????? ??????"
        
        .Range("A38").Value = "REPORT_STOCK_CARD"
        .Range("B38").Value = "????? ???????"
        
        .Range("A39").Value = "REPORT_LOW_STOCK"
        .Range("B39").Value = "????? ??????? ???????"
        
        .Range("A40").Value = "REPORT_MOVEMENTS"
        .Range("B40").Value = "????? ???????"
        
        ' === ACTIONS ===
        .Range("A41").Value = "ACTION_ADD"
        .Range("B41").Value = "?????"
        
        .Range("A42").Value = "ACTION_EDIT"
        .Range("B42").Value = "?????"
        
        .Range("A43").Value = "ACTION_DELETE"
        .Range("B43").Value = "???"
        
        .Range("A44").Value = "ACTION_SEARCH"
        .Range("B44").Value = "???"
        
        .Range("A45").Value = "ACTION_PRINT"
        .Range("B45").Value = "?????"
        
        .Range("A46").Value = "ACTION_EXPORT"
        .Range("B46").Value = "?????"
        
        .Range("A47").Value = "ACTION_IMPORT"
        .Range("B47").Value = "???????"
        
        .Range("A48").Value = "ACTION_CALCULATE"
        .Range("B48").Value = "????"
        
        ' === FIELDS ===
        .Range("A49").Value = "FIELD_CODE"
        .Range("B49").Value = "?????"
        
        .Range("A50").Value = "FIELD_NAME"
        .Range("B50").Value = "?????"
        
        .Range("A51").Value = "FIELD_DESIGNATION"
        .Range("B51").Value = "???????"
        
        .Range("A52").Value = "FIELD_CATEGORY"
        .Range("B52").Value = "?????"
        
        .Range("A53").Value = "FIELD_UNIT"
        .Range("B53").Value = "??????"
        
        .Range("A54").Value = "FIELD_MIN_STOCK"
        .Range("B54").Value = "???? ??????"
        
        .Range("A55").Value = "FIELD_MAX_STOCK"
        .Range("B55").Value = "???? ??????"
        
        .Range("A56").Value = "FIELD_SUPPLIER"
        .Range("B56").Value = "??????"
        
        .Range("A57").Value = "FIELD_REFERENCE"
        .Range("B57").Value = "??????"
        
        ' === MESSAGES ===
        .Range("A58").Value = "MSG_CONFIRM_DELETE"
        .Range("B58").Value = "?? ??? ????? ?? ??????"
        
        .Range("A59").Value = "MSG_SAVE_SUCCESS"
        .Range("B59").Value = "?? ????? ?????"
        
        .Range("A60").Value = "MSG_DELETE_SUCCESS"
        .Range("B60").Value = "?? ????? ?????"
        
        .Range("A61").Value = "MSG_UPDATE_SUCCESS"
        .Range("B61").Value = "?? ??????? ?????"
        
        .Range("A62").Value = "MSG_NO_DATA"
        .Range("B62").Value = "?? ???? ??????"
        
        .Range("A63").Value = "MSG_SELECT_ITEM"
        .Range("B63").Value = "???? ?????? ????"
        
        .Range("A64").Value = "MSG_REQUIRED_FIELD"
        .Range("B64").Value = "??? ?????"
        
        ' === BUTTONS ===
        .Range("A65").Value = "BTN_NEW"
        .Range("B65").Value = "????"
        
        .Range("A66").Value = "BTN_EDIT"
        .Range("B66").Value = "?????"
        
        .Range("A67").Value = "BTN_DELETE"
        .Range("B67").Value = "???"
        
        .Range("A68").Value = "BTN_CLOSE"
        .Range("B68").Value = "?????"
        
        .Range("A69").Value = "BTN_PRINT"
        .Range("B69").Value = "?????"
        
        .Range("A70").Value = "BTN_EXPORT"
        .Range("B70").Value = "?????"
        
        .Range("A71").Value = "BTN_CALCULATE_EOQ"
        .Range("B71").Value = "???? EOQ"
        
        .Range("A72").Value = "BTN_CALCULATE_CMUP"
        .Range("B72").Value = "???? CMUP"
        
        ' === ERRORS ===
        .Range("A73").Value = "ERR_REQUIRED"
        .Range("B73").Value = "??? ????? ?????"
        
        .Range("A74").Value = "ERR_NUMBER_ONLY"
        .Range("B74").Value = "??? ????? ???"
        
        .Range("A75").Value = "ERR_DUPLICATE"
        .Range("B75").Value = "??? ????? ????? ?????"
        
        .Range("A76").Value = "ERR_NEGATIVE_STOCK"
        .Range("B76").Value = "?? ???? ?? ???? ??????? ?????"
        
        .Range("A77").Value = "ERR_INSUFFICIENT_STOCK"
        .Range("B77").Value = "????? ??? ???"
        
        ' === UNITS ===
        .Range("A78").Value = "UNIT_PIECE"
        .Range("B78").Value = "????"
        
        .Range("A79").Value = "UNIT_BOX"
        .Range("B79").Value = "?????"
        
        .Range("A80").Value = "UNIT_PACK"
        .Range("B80").Value = "????"
        
        .Range("A81").Value = "UNIT_REAM"
        .Range("B81").Value = "????"
        
        .Range("A82").Value = "UNIT_LITER"
        .Range("B82").Value = "???"
        
        ' === EOQ/ROP ===
        .Range("A83").Value = "LABEL_EOQ"
        .Range("B83").Value = "?????? ?????? ?????"
        
        .Range("A84").Value = "LABEL_ROP"
        .Range("B84").Value = "???? ????? ?????"
        
        .Range("A85").Value = "LABEL_SAFETY_STOCK"
        .Range("B85").Value = "????? ??????"
        
        .Range("A86").Value = "LABEL_LEAD_TIME"
        .Range("B86").Value = "??? ???????"
        
        .Range("A87").Value = "LABEL_DAILY_DEMAND"
        .Range("B87").Value = "????? ??????"
        
        .Range("A88").Value = "LABEL_ANNUAL_DEMAND"
        .Range("B88").Value = "????? ??????"
        
        .Columns("A").ColumnWidth = 20
        .Columns("B").ColumnWidth = 35
        .Tab.Color = RGB(149, 165, 166)
        .Protect mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    End With
End Sub

Private Sub CreateReceiptTagSheet()
    Dim ws As Worksheet
    Dim sheetExists As Boolean
    
    sheetExists = False
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("RECEIPT_TAG")
    If Err.Number = 0 Then sheetExists = True
    On Error GoTo 0
    
    Application.DisplayAlerts = False
    If sheetExists Then ws.Delete
    Application.DisplayAlerts = True
    
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    ws.name = "RECEIPT_TAG"
    
    With ws
        On Error Resume Next
        .PageSetup.PaperSize = xlPaperA5
        .PageSetup.Orientation = xlPortrait
        On Error GoTo 0
        
        .Range("A1").Value = "RECEIPT TAG"
        .Range("A1").Font.Bold = True
        .Protect mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    End With
End Sub

Private Sub CreateTemplateBonSheet()
    Dim ws As Worksheet
    Dim sheetExists As Boolean
    
    sheetExists = False
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("TEMPLATE_BON")
    If Err.Number = 0 Then sheetExists = True
    On Error GoTo 0
    
    Application.DisplayAlerts = False
    If sheetExists Then ws.Delete
    Application.DisplayAlerts = True
    
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    ws.name = "TEMPLATE_BON"
    
    With ws
        .Range("A1").Value = "DOCUMENT OFFICIEL"
        .Range("A1").Font.Bold = True
        .Protect mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    End With
End Sub
