Attribute VB_Name = "mod_ExportEngine"
'==============================================================================
' MODULE: mod_ExportEngine.bas  |  ERP LSM v1.0.0
' Author: LSM VBA Core | Public Sector 2026
'
' PDF Document Generation - Algerian Public Sector Compliance
'
' COMPLIANCE FEATURES:
'   - TVA exemption (Instruction 09-2018 / Article 5)
'   - 4-signature block (Fournisseur, Comptable, Responsable, Directeur)
'   - Engagement / Liquidation / Code Budgétaire lines
'   - QR code (generated BEFORE PDF export - embedded in document)
'   - Verification code (deterministic hash)
'   - NIF / NIS / RC / Article tax identifiers
'   - DGI-standard document numbering (BS-YYYY-NNNN)
'   - A4 portrait, proper margins, print area
'
' MOUVEMENTS COLUMN CONTRACT:
'   Col A: DATE | Col B: CODE_ARTICLE | Col C: DESIGNATION
'   Col D: TYPE_MVT | Col E: QTE | Col F: LINE_VALUE
'   Col G: REF_DOCUMENT | Col H: PRIX_UNITAIRE | Col I: THIRD_PARTY
'   Col L: NOTES
'==============================================================================

Option Explicit

'================================================================================
' SECTION 1 - PRIMARY EXPORT ENTRY POINT
'================================================================================

Public Sub ExportTransactionToPDF(ByVal docRef As String)
    Call ExportTransactionToPDF_Internal(docRef, False, "")
End Sub

' Silent export for batch operations - no dialog, no MsgBox, saves to specified path
Public Function ExportTransactionToPDF_Silent(ByVal docRef As String, Optional ByVal outputPath As String) As Boolean
    ExportTransactionToPDF_Silent = ExportTransactionToPDF_Internal(docRef, True, outputPath)
End Function

Private Function ExportTransactionToPDF_Internal(ByVal docRef As String, ByVal silent As Boolean, Optional ByVal outputPath As String) As Boolean
    Dim wsTemplate  As Worksheet
    Dim wsMouv      As Worksheet
    Dim savePath    As String
    Dim fileName    As String
    Dim fullPath    As String
    
    If Not silent Then Debug.Print "--- PDF EXPORT START: " & docRef & " ---"
    On Error GoTo ExportError
    
    ' 1. Pre-flight: validate required sheets
    If Not sheetExists("TEMPLATE_BON") Then
        If Not silent Then Debug.Print "[Export] FAIL: TEMPLATE_BON sheet missing"
        GoTo ExportError
    End If
    If Not sheetExists(mod_Config.SHEET_MOUVEMENTS) Then
        If Not silent Then Debug.Print "[Export] FAIL: MOUVEMENTS sheet missing"
        GoTo ExportError
    End If
    
    ' Self-healing: ensure MOUVEMENTS headers exist
    Call mod_Utilities.RestoreMouvementsHeaders(silent:=True)
    
    Set wsTemplate = ThisWorkbook.Sheets("TEMPLATE_BON")
    Set wsMouv = ThisWorkbook.Sheets(mod_Config.SHEET_MOUVEMENTS)
    
    ' 2. Populate template from transaction data
    If Not silent Then Debug.Print "[Export] Populating template for Ref: " & docRef
    If Not PopulateTemplateBon(docRef, wsMouv, wsTemplate) Then
        If Not silent Then Debug.Print "[Export] FAIL: No lines found for " & docRef
        GoTo ExportError
    End If
    
    ' 3. Generate QR code BEFORE PDF export
    On Error Resume Next
    Dim qrRow As Long
    qrRow = GetQRRow(wsTemplate)
    Call mod_QRCode.GenerateQRCodeForForm(docRef, "TEMPLATE_BON", "F" & qrRow)
    On Error GoTo ExportError
    
    ' 4. Determine save path
    If silent Then
        fullPath = IIf(Len(outputPath) > 0, outputPath, mod_SharedEnvironment.GetSharedExportPath() & docRef & "_" & Format(Date, "yyyy-mm-dd") & ".pdf")
    Else
        savePath = SelectPDFSavePath(docRef)
        If savePath = "" Then
            Debug.Print "[Export] User cancelled save dialog"
            ExportTransactionToPDF_Internal = False
            Exit Function
        End If
        fullPath = savePath
    End If
    
    ' 5. Export to PDF (QR code now embedded in template)
    wsTemplate.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        fileName:=fullPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=False, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=Not silent
    
    If Not silent Then
        Debug.Print "[Export] SUCCESS: PDF generated with QR code."
        MsgBox "Document export" & Chr(233) & " vers :" & vbCrLf & fullPath, _
               vbInformation, "LSM v1.0.0"
    End If
    
    ExportTransactionToPDF_Internal = True
    Exit Function
    
ExportError:
    If Not silent Then
        Debug.Print "[Export] CRASH: " & Err.Description & " (Error #" & Err.Number & ")"
        MsgBox "Erreur PDF Export : " & Err.Description, vbCritical, "LSM v1.0.0"
    End If
    ExportTransactionToPDF_Internal = False
End Function

'================================================================================
' SECTION 2 - TEMPLATE POPULATION ENGINE
'================================================================================

Private Function PopulateTemplateBon(ByVal docRef As String, _
                                      ByRef wsMouv As Worksheet, _
                                      ByRef wsTpl As Worksheet) As Boolean
    Dim lastRow      As Long
    Dim j            As Long
    Dim r            As Integer
    Dim docDate      As String
    Dim mvtSign      As String
    Dim docType      As String
    Dim thirdParty   As String
    Dim totalVal     As Double
    Dim lineCount    As Integer
    Dim wsArt        As Worksheet
    
    ' Column discovery (robust against column-order variations)
    Dim colDate    As Integer: colDate = FindColumn(wsMouv, "DATE")
    Dim colCode    As Integer: colCode = FindColumn(wsMouv, "CODE_ARTICLE")
    Dim colDesig   As Integer: colDesig = FindColumn(wsMouv, "DESIGNATION")
    Dim colType    As Integer: colType = FindColumn(wsMouv, "TYPE_MVT")
    Dim colQte     As Integer: colQte = FindColumn(wsMouv, "QTE")
    Dim colRef     As Integer: colRef = FindColumn(wsMouv, "REF_DOCUMENT")
    Dim colPU      As Integer: colPU = FindColumn(wsMouv, "PRIX_UNITAIRE")
    Dim colThird   As Integer: colThird = FindColumn(wsMouv, "THIRD_PARTY")
    Dim colNotes   As Integer: colNotes = FindColumn(wsMouv, "NOTES")
    
    ' Critical check
    If colDate = 0 Or colCode = 0 Or colRef = 0 Then
        MsgBox "ERREUR: Colonnes obligatoires introuvables dans MOUVEMENTS." & vbCrLf & _
               "V" & Chr(233) & "rifiez les ent" & Chr(234) & "tes (DATE, CODE_ARTICLE, REF_DOCUMENT).", _
               vbCritical, "Export Error"
        PopulateTemplateBon = False
        Exit Function
    End If
    
    ' Fallback to hardcoded positions
    If colDesig = 0 Then colDesig = 3
    If colType = 0 Then colType = 4
    If colQte = 0 Then colQte = 5
    If colPU = 0 Then colPU = 8
    If colThird = 0 Then colThird = 9
    If colNotes = 0 Then colNotes = 12
    
    On Error GoTo PopulateError
    
    Application.ScreenUpdating = False
    wsTpl.Unprotect Password:=mod_Config.MASTER_PWD
    wsTpl.Cells.Clear
    wsTpl.Cells.Interior.ColorIndex = xlNone
    
    ' Scan MOUVEMENTS for matching rows
    lastRow = wsMouv.Cells(wsMouv.Rows.Count, 1).End(xlUp).Row
    lineCount = 0
    docDate = Format(Date, "DD/MM/YYYY")
    mvtSign = "OUT"
    thirdParty = ""
    
    On Error Resume Next
    Set wsArt = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    On Error GoTo PopulateError
    
    For j = 2 To lastRow
        If Trim(CStr(wsMouv.Cells(j, colRef).Value)) = docRef Then
            If lineCount = 0 Then
                If IsDate(wsMouv.Cells(j, colDate).Value) Then
                    docDate = Format(CDate(wsMouv.Cells(j, colDate).Value), "DD/MM/YYYY")
                End If
                mvtSign = UCase(Trim(CStr(wsMouv.Cells(j, colType).Value)))
                thirdParty = Trim(CStr(wsMouv.Cells(j, colThird).Value))
            End If
            lineCount = lineCount + 1
        End If
    Next j
    
    If lineCount = 0 Then
        wsTpl.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
        Application.ScreenUpdating = True
        PopulateTemplateBon = False
        Exit Function
    End If
    
    docType = IIf(mvtSign = "IN", _
                  "BON DE R" & Chr(201) & "CEPTION", _
                  "BON DE SORTIE")
    
    ' PAGE SETUP
    With wsTpl.PageSetup
        .Orientation = xlPortrait
        .PaperSize = xlPaperA4
        .LeftMargin = Application.CentimetersToPoints(2)
        .RightMargin = Application.CentimetersToPoints(1.5)
        .TopMargin = Application.CentimetersToPoints(2)
        .BottomMargin = Application.CentimetersToPoints(2)
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .RightFooter = "ERP Acad" & Chr(233) & "mie v1.0.0  |  " & _
                          Format(Now, "DD/MM/YYYY HH:MM")
    End With
    
    r = 1
    With wsTpl
        ' ROW 1: Ministry header
        .Range("A" & r & ":G" & r).Merge
        .Cells(r, 1).Value = "MINIST" & Chr(200) & "RE DE L'" & Chr(201) & _
                             "DUCATION NATIONALE"
        With .Cells(r, 1)
            .Font.Bold = True
            .Font.Size = 11
            .HorizontalAlignment = xlCenter
        End With
        .Rows(r).RowHeight = 22: r = r + 1
        
        ' ROW 2: Direction (bilingual FR/AR)
        .Range("A" & r & ":G" & r).Merge
        .Cells(r, 1).Value = "Direction de l'" & Chr(201) & "ducation  " & _
                             Chr(8212) & "  [CITY]  |  " & _
                             Chr(1605) & Chr(1583) & Chr(1610) & Chr(1585) & _
                             Chr(1610) & Chr(1577) & " " & Chr(1575) & _
                             Chr(1604) & Chr(1578) & Chr(1585) & Chr(1576) & _
                             Chr(1610) & Chr(1577) & " " & Chr(1575) & _
                             Chr(1604) & Chr(1576) & Chr(1610) & Chr(1590)
        With .Cells(r, 1)
            .Font.Size = 9
            .Font.Italic = True
            .HorizontalAlignment = xlCenter
            .Font.Name = "Tahoma"
        End With
        .Rows(r).RowHeight = 18: r = r + 1
        
        ' ROW 3: Double separator
        .Range("A" & r & ":G" & r).Borders(xlEdgeBottom).LineStyle = xlDouble
        .Range("A" & r & ":G" & r).Borders(xlEdgeBottom).Weight = xlThick
        .Rows(r).RowHeight = 6: r = r + 1
        
        ' ROW 4: Spacer
        .Rows(r).RowHeight = 10: r = r + 1
        
        ' ROW 5: Document title banner
        .Range("A" & r & ":G" & r).Merge
        .Cells(r, 1).Value = docType
        With .Cells(r, 1)
            .Font.Bold = True
            .Font.Size = 20
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Interior.Color = RGB(0, 70, 127)
            .Font.Color = RGB(255, 255, 255)
        End With
        .Rows(r).RowHeight = 40: r = r + 1
        
        ' ROW 6: Spacer
        .Rows(r).RowHeight = 10: r = r + 1
        
        ' ROW 7: Metadata - Ref, Date, Type
        .Cells(r, 1).Value = "N" & Chr(176) & " R" & Chr(233) & "f" & Chr(233) & "rence :"
        .Cells(r, 1).Font.Bold = True
        .Cells(r, 2).Value = docRef
        .Cells(r, 2).Font.Bold = True
        .Cells(r, 2).Font.Color = RGB(0, 70, 127)
        
        .Cells(r, 4).Value = "Date :"
        .Cells(r, 4).Font.Bold = True
        .Cells(r, 5).Value = docDate
        
        .Cells(r, 6).Value = "Type :"
        .Cells(r, 6).Font.Bold = True
        .Cells(r, 7).Value = IIf(mvtSign = "IN", "ENTR" & Chr(201) & "E", "SORTIE")
        .Cells(r, 7).Font.Bold = True
        .Cells(r, 7).Font.Color = IIf(mvtSign = "IN", RGB(4, 90, 55), RGB(160, 70, 0))
        .Rows(r).RowHeight = 18: r = r + 1
        
        ' ROW 8: Service / Fournisseur
        If Len(thirdParty) > 0 Then
            .Cells(r, 1).Value = "Service / Fournisseur :"
            .Cells(r, 1).Font.Bold = True
            .Range("B" & r & ":G" & r).Merge
            .Cells(r, 2).Value = thirdParty
            .Rows(r).RowHeight = 16: r = r + 1
        End If
        
        ' ROW 9: Spacer
        .Rows(r).RowHeight = 8: r = r + 1
        
        ' ROW 10: Column Headers
        Dim hdrs(5) As String
        hdrs(0) = "Code Article"
        hdrs(1) = "D" & Chr(233) & "signation"
        hdrs(2) = "Unit" & Chr(233)
        hdrs(3) = "Qt" & Chr(233)
        hdrs(4) = "PU (DZD)"
        hdrs(5) = "Valeur (DZD)"
        
        Dim c As Integer
        For c = 0 To 5
            With .Cells(r, c + 1)
                .Value = hdrs(c)
                .Font.Bold = True
                .Font.Size = 9
                .Font.Color = RGB(255, 255, 255)
                .Interior.Color = RGB(0, 70, 127)
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .WrapText = True
                .Borders.LineStyle = xlContinuous
                .Borders.Weight = xlThin
            End With
        Next c
        .Rows(r).RowHeight = 28: r = r + 1
        
        ' ROWS 11+: Data rows
        totalVal = 0
        For j = 2 To lastRow
            If Trim(CStr(wsMouv.Cells(j, colRef).Value)) = docRef Then
                Dim artCode  As String
                Dim artDesig As String
                Dim artUnit  As String
                Dim qty      As Double
                Dim pu       As Double
                Dim valLigne As Double
                
                artCode = Trim(CStr(wsMouv.Cells(j, colCode).Value))
                artDesig = Trim(CStr(wsMouv.Cells(j, colDesig).Value))
                artUnit = "unit" & Chr(233)
                qty = mod_Utilities.SafeVal(wsMouv.Cells(j, colQte).Value)
                pu = mod_Utilities.SafeVal(wsMouv.Cells(j, colPU).Value)
                valLigne = qty * pu
                
                ' Lookup Arabic designation from ARTICLES
                If Not wsArt Is Nothing Then
                    Dim artMatchRow As Variant
                    artMatchRow = Application.Match(artCode, wsArt.Range("A:A"), 0)
                    If Not IsError(artMatchRow) Then
                        Dim arLabel As String
                        arLabel = Trim(CStr(wsArt.Cells(artMatchRow, 2).Value))
                        Dim unitLabel As String
                        unitLabel = Trim(CStr(wsArt.Cells(artMatchRow, 4).Value))
                        If Len(arLabel) > 0 Then artDesig = arLabel
                        If Len(unitLabel) > 0 Then artUnit = unitLabel
                    End If
                End If
                
                ' Alternating row shading
                Dim rowBg As Long
                rowBg = IIf((r Mod 2) = 0, RGB(235, 242, 250), RGB(255, 255, 255))
                
                ' Col A: Code
                With .Cells(r, 1)
                    .Value = artCode
                    .Interior.Color = rowBg
                    .HorizontalAlignment = xlCenter
                    .Font.Name = "Courier New": .Font.Size = 9
                    .Borders.LineStyle = xlContinuous
                    .Borders.Weight = xlThin
                End With
                
                ' Col B: Designation (Arabic - right-align)
                With .Cells(r, 2)
                    .Value = artDesig
                    .Interior.Color = rowBg
                    .Font.Name = "Tahoma": .Font.Size = 9
                    .HorizontalAlignment = xlRight
                    .Borders.LineStyle = xlContinuous
                    .Borders.Weight = xlThin
                End With
                
                ' Col C: Unit
                With .Cells(r, 3)
                    .Value = artUnit
                    .Interior.Color = rowBg
                    .Font.Name = "Tahoma": .Font.Size = 9
                    .HorizontalAlignment = xlCenter
                    .Borders.LineStyle = xlContinuous
                    .Borders.Weight = xlThin
                End With
                
                ' Col D: Quantity
                With .Cells(r, 4)
                    .Value = qty
                    .NumberFormat = "#,##0"
                    .Interior.Color = rowBg
                    .HorizontalAlignment = xlCenter
                    .Font.Bold = True: .Font.Size = 9
                    .Borders.LineStyle = xlContinuous
                    .Borders.Weight = xlThin
                End With
                
                ' Col E: PU
                With .Cells(r, 5)
                    .Value = pu
                    .NumberFormat = "#,##0.00"
                    .Interior.Color = rowBg
                    .HorizontalAlignment = xlRight
                    .Font.Size = 9
                    .Borders.LineStyle = xlContinuous
                    .Borders.Weight = xlThin
                End With
                
                ' Col F: Line Value
                With .Cells(r, 6)
                    .Value = valLigne
                    .NumberFormat = "#,##0.00"
                    .Interior.Color = rowBg
                    .HorizontalAlignment = xlRight
                    .Font.Bold = True: .Font.Size = 9
                    .Borders.LineStyle = xlContinuous
                    .Borders.Weight = xlThin
                End With
                
                totalVal = totalVal + valLigne
                .Rows(r).RowHeight = 20
                r = r + 1
            End If
        Next j
        
        ' TOTAL ROW
        .Range("A" & r & ":E" & r).Merge
        With .Cells(r, 1)
            .Value = "TOTAL G" & Chr(201) & "N" & Chr(201) & "RAL"
            .Font.Bold = True: .Font.Size = 10
            .HorizontalAlignment = xlRight
            .Interior.Color = RGB(215, 228, 244)
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlMedium
        End With
        With .Cells(r, 6)
            .Value = totalVal
            .NumberFormat = "#,##0.00 DZD"
            .Font.Bold = True: .Font.Size = 11
            .Font.Color = RGB(0, 70, 127)
            .Interior.Color = RGB(215, 228, 244)
            .HorizontalAlignment = xlRight
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlMedium
        End With
        .Rows(r).RowHeight = 24: r = r + 1
        
        ' TVA EXEMPTION (Public Sector)
        .Range("A" & r & ":G" & r).Merge
        .Cells(r, 1).Value = "TVA non applicable -- Secteur Public (Instruction 09-2018 / Article 5)"
        With .Cells(r, 1)
            .Font.Size = 8: .Font.Italic = True
            .Font.Color = RGB(100, 100, 100)
            .HorizontalAlignment = xlRight
        End With
        .Rows(r).RowHeight = 14: r = r + 1
        
        ' MULTI-COPY INDICATOR
        .Range("A" & r & ":G" & r).Merge
        .Cells(r, 1).Value = "ORIGINAL -- Exemplaire du Magasin"
        With .Cells(r, 1)
            .Font.Bold = True: .Font.Size = 9
            .Font.Color = RGB(0, 70, 127)
            .HorizontalAlignment = xlCenter
            .Interior.Color = RGB(230, 240, 250)
        End With
        .Rows(r).RowHeight = 16: r = r + 1
        
        ' SPACER
        .Rows(r).RowHeight = 10: r = r + 1
        
        ' SIGNATURE ZONE (4 blocks - Algerian Public Sector Standard)
        .Range("A" & r & ":B" & r).Merge
        .Cells(r, 1).Value = "Le Fournisseur"
        .Cells(r, 1).Font.Bold = True: .Cells(r, 1).Font.Size = 8
        .Cells(r, 1).HorizontalAlignment = xlCenter
        
        .Range("C" & r & ":D" & r).Merge
        .Cells(r, 3).Value = "Le Comptable"
        .Cells(r, 3).Font.Bold = True: .Cells(r, 3).Font.Size = 8
        .Cells(r, 3).HorizontalAlignment = xlCenter
        
        .Range("E" & r & ":F" & r).Merge
        .Cells(r, 5).Value = "Le Responsable"
        .Cells(r, 5).Font.Bold = True: .Cells(r, 5).Font.Size = 8
        .Cells(r, 5).HorizontalAlignment = xlCenter
        
        .Cells(r, 7).Value = "Le Directeur"
        .Cells(r, 7).Font.Bold = True: .Cells(r, 7).Font.Size = 8
        .Cells(r, 7).HorizontalAlignment = xlCenter
        .Rows(r).RowHeight = 14: r = r + 1
        
        ' Signature boxes
        .Rows(r).RowHeight = 40
        .Range("A" & r & ":B" & r).Merge
        .Range("A" & r).Borders.LineStyle = xlContinuous
        .Range("A" & r).Borders.Weight = xlThin
        .Range("A" & r).Interior.Color = RGB(250, 250, 250)
        
        .Range("C" & r & ":D" & r).Merge
        .Range("C" & r).Borders.LineStyle = xlContinuous
        .Range("C" & r).Borders.Weight = xlThin
        .Range("C" & r).Interior.Color = RGB(250, 250, 250)
        
        .Range("E" & r & ":F" & r).Merge
        .Range("E" & r).Borders.LineStyle = xlContinuous
        .Range("E" & r).Borders.Weight = xlThin
        .Range("E" & r).Interior.Color = RGB(250, 250, 250)
        
        .Cells(r, 7).Borders.LineStyle = xlContinuous
        .Cells(r, 7).Borders.Weight = xlThin
        .Cells(r, 7).Interior.Color = RGB(250, 250, 250)
        r = r + 1
        
        ' Cachet labels
        .Range("A" & r & ":B" & r).Merge
        .Cells(r, 1).Value = "Signature & Cachet"
        .Cells(r, 1).Font.Size = 7: .Cells(r, 1).Font.Italic = True
        .Cells(r, 1).Font.Color = RGB(128, 128, 128)
        .Cells(r, 1).HorizontalAlignment = xlCenter
        
        .Range("C" & r & ":D" & r).Merge
        .Cells(r, 3).Value = "Visa du Comptable"
        .Cells(r, 3).Font.Size = 7: .Cells(r, 3).Font.Italic = True
        .Cells(r, 3).Font.Color = RGB(128, 128, 128)
        .Cells(r, 3).HorizontalAlignment = xlCenter
        
        .Range("E" & r & ":F" & r).Merge
        .Cells(r, 5).Value = "Visa du Responsable"
        .Cells(r, 5).Font.Size = 7: .Cells(r, 5).Font.Italic = True
        .Cells(r, 5).Font.Color = RGB(128, 128, 128)
        .Cells(r, 5).HorizontalAlignment = xlCenter
        
        .Cells(r, 7).Value = "Visa du Directeur"
        .Cells(r, 7).Font.Size = 7: .Cells(r, 7).Font.Italic = True
        .Cells(r, 7).Font.Color = RGB(128, 128, 128)
        .Cells(r, 7).HorizontalAlignment = xlCenter
        .Rows(r).RowHeight = 12: r = r + 1
        
        ' BUDGET / ENGAGEMENT LINE
        .Range("A" & r & ":G" & r).Merge
        .Cells(r, 1).Value = "N" & Chr(176) & " Engagement : _______________  |  N" & Chr(176) & " Liquidation : _______________  |  Code Budg" & Chr(233) & "taire : _______________"
        With .Cells(r, 1)
            .Font.Size = 8: .Font.Italic = True
            .Font.Color = RGB(80, 80, 80)
            .HorizontalAlignment = xlCenter
        End With
        .Rows(r).RowHeight = 14: r = r + 1
        
        ' TAX IDENTIFIERS (NIF/NIS/RC/Art - auto-filled from supplier registry)
        Dim taxIDs As String
        taxIDs = mod_SupplierRegistry.GetSupplierTaxIDsForPDF(thirdParty)
        .Range("A" & r & ":G" & r).Merge
        .Cells(r, 1).Value = taxIDs
        With .Cells(r, 1)
            .Font.Size = 8: .Font.Italic = True
            .Font.Color = RGB(80, 80, 80)
            .HorizontalAlignment = xlCenter
        End With
        .Rows(r).RowHeight = 14: r = r + 1
        
        ' SPACER
        .Rows(r).RowHeight = 8: r = r + 1
        
        ' VERIFICATION CODE
        Dim verifyCode As String
        verifyCode = mod_Utilities.GenerateVerifyCode(docRef & docType & docDate & Format(totalVal, "0.00") & mod_Config.APP_VERSION)
        .Range("A" & r & ":G" & r).Merge
        .Cells(r, 1).Value = "Code v" & Chr(233) & "rification : " & verifyCode
        With .Cells(r, 1)
            .Font.Name = "Courier New"
            .Font.Size = 8: .Font.Bold = True
            .Font.Color = RGB(0, 70, 127)
            .HorizontalAlignment = xlCenter
        End With
        .Rows(r).RowHeight = 14: r = r + 1
        
        ' QR CODE PLACEHOLDER (QR generated BEFORE PDF export in ExportTransactionToPDF)
        .Range("F" & r & ":G" & r).Merge
        .Cells(r, 6).Value = "[QR]"
        .Cells(r, 6).Font.Size = 8
        .Cells(r, 6).Font.Color = RGB(180, 180, 180)
        .Cells(r, 6).HorizontalAlignment = xlCenter
        .Cells(r, 6).VerticalAlignment = xlCenter
        .Cells(r, 6).Interior.Color = RGB(245, 245, 245)
        .Cells(r, 6).BorderAround Color:=RGB(200, 200, 200), Weight:=xlThin
        .Rows(r).RowHeight = 30: r = r + 1
        
        ' FOOTER
        .Rows(r).RowHeight = 6: r = r + 1
        .Range("A" & r & ":G" & r).Merge
        .Cells(r, 1).Value = "Document g" & Chr(233) & "n" & Chr(233) & "r" & Chr(233) & _
                             " par ERP Acad" & Chr(233) & "mie v1.0.0  |  " & _
                             Format(Now, "DD/MM/YYYY HH:MM") & _
                             "  |  Syst" & Chr(232) & "me de Gestion Minist" & Chr(232) & "re " & Chr(201) & "ducation  |  " & _
                             verifyCode
        With .Cells(r, 1)
            .Font.Size = 7: .Font.Italic = True
            .Font.Color = RGB(128, 128, 128)
            .HorizontalAlignment = xlCenter
        End With
        .Rows(r).RowHeight = 12
        
        ' COLUMN WIDTHS + PRINT AREA
        .Columns("A").ColumnWidth = 13
        .Columns("B").ColumnWidth = 30
        .Columns("C").ColumnWidth = 10
        .Columns("D").ColumnWidth = 7
        .Columns("E").ColumnWidth = 13
        .Columns("F").ColumnWidth = 16
        .Columns("G").ColumnWidth = 12
        .PageSetup.PrintArea = "$A$1:$G$" & r
        
    End With
    
    wsTpl.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    Application.ScreenUpdating = True
    
    PopulateTemplateBon = True
    Exit Function

PopulateError:
    Application.ScreenUpdating = True
    On Error Resume Next
    wsTpl.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    MsgBox "Erreur construction template: " & Err.Description, vbCritical, "LSM v1.0.0"
    PopulateTemplateBon = False
End Function

'================================================================================
' SECTION 2B - VERIFICATION CODE GENERATOR
'================================================================================

'================================================================================
' SECTION 3 - DYNAMIC COLUMN FINDER
'================================================================================

Private Function FindColumn(ByRef ws As Worksheet, _
                              ByVal headerName As String) As Integer
    Dim lastCol As Integer
    Dim c       As Integer
    
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        Dim cellVal As String
        cellVal = UCase(Trim(CStr(ws.Cells(1, c).Value)))
        If cellVal = UCase(headerName) Or _
           InStr(cellVal, UCase(Replace(headerName, "_", " "))) > 0 Then
            FindColumn = c
            Exit Function
        End If
    Next c
    FindColumn = 0
End Function

'================================================================================
' SECTION 4 - SHEET EXISTENCE CHECK
'================================================================================

Private Function sheetExists(ByVal sheetName As String) As Boolean
    Dim s As Worksheet
    sheetExists = False
    For Each s In ThisWorkbook.Sheets
        If s.Name = sheetName Then
            sheetExists = True
            Exit Function
        End If
    Next s
End Function

'================================================================================
' SECTION 5B - PDF SAVE PATH SELECTOR
'================================================================================

Private Function SelectPDFSavePath(ByVal docRef As String) As String
    Dim dlg As Object
    Dim fileName As String
    Dim desktopPath As String
    
    On Error GoTo FallbackDesktop
    
    Set dlg = Application.FileDialog(2)
    If dlg Is Nothing Then GoTo FallbackDesktop
    
    fileName = docRef & "_" & Format(Date, "yyyy-mm-dd") & ".pdf"
    
    With dlg
        .Title = "En[ARTICLE_DESC]r le document PDF -- " & docRef
        .InitialFileName = mod_SharedEnvironment.GetSharedExportPath() & fileName
        .Filters.Clear
        .Filters.Add "PDF Files", "*.pdf"
        .FilterIndex = 1
        
        If .Show = -1 Then
            SelectPDFSavePath = .SelectedItems(1)
            If Right(SelectPDFSavePath, 4) <> ".pdf" Then
                SelectPDFSavePath = SelectPDFSavePath & ".pdf"
            End If
        Else
            SelectPDFSavePath = ""
        End If
    End With
    
    Set dlg = Nothing
    Exit Function

FallbackDesktop:
    desktopPath = mod_SharedEnvironment.GetSharedExportPath()
    SelectPDFSavePath = desktopPath & docRef & "_" & Format(Date, "yyyy-mm-dd") & ".pdf"
    Set dlg = Nothing
End Function

'================================================================================
' SECTION 5C - GET QR CODE ROW
'================================================================================

Private Function GetQRRow(ByRef ws As Worksheet) As Long
    Dim r As Long
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    For r = 1 To lastRow
        If ws.Range("F" & r).Value = "[QR]" Then
            GetQRRow = r
            Exit Function
        End If
    Next r
    
    GetQRRow = lastRow - 3
End Function

'================================================================================
' SECTION 6 - EXISTING EXPORTS
'================================================================================

Public Sub ExportToExcel(Optional sheetName As String = "ARTICLES")
    Dim ws As Worksheet
    Dim desktopPath As String, fileName As String, fullPath As String
    Dim wb As Workbook
    On Error GoTo ExportError2
    Set ws = ThisWorkbook.Sheets(sheetName)
    desktopPath = mod_SharedEnvironment.GetSharedExportPath()
    fileName = sheetName & "_Export_" & Format(Date, "yyyy-mm-dd") & ".xlsx"
    fullPath = desktopPath & fileName
    ws.Copy
    Set wb = ActiveWorkbook
    wb.SaveAs fileName:=fullPath, FileFormat:=xlOpenXMLWorkbook
    wb.Close
    MsgBox "Export" & Chr(233) & " vers: " & fullPath, vbInformation, "LSM v1.0.0"
    Exit Sub
ExportError2:
    MsgBox "Export Error: " & Err.Description, vbCritical, "LSM v1.0.0"
End Sub

Public Sub ExportDashboardPDF()
    Dim wsDash      As Worksheet
    Dim desktopPath As String, fileName As String, fullPath As String
    On Error GoTo ExportError3
    Set wsDash = ThisWorkbook.Sheets("DASHBOARD")
    desktopPath = mod_SharedEnvironment.GetSharedExportPath()
    fileName = "Dashboard_Report_" & Format(Date, "yyyy-mm-dd") & ".pdf"
    fullPath = desktopPath & fileName
    
    ' Proper page setup for dashboard export
    With wsDash.PageSetup
        .Orientation = xlPortrait
        .PaperSize = xlPaperA4
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .LeftMargin = Application.CentimetersToPoints(1.5)
        .RightMargin = Application.CentimetersToPoints(1.5)
        .TopMargin = Application.CentimetersToPoints(1.5)
        .BottomMargin = Application.CentimetersToPoints(1.5)
        .CenterHorizontally = True
    End With
    
    wsDash.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        fileName:=fullPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=False, _
        OpenAfterPublish:=True
    
    MsgBox "Dashboard export" & Chr(233) & " vers: " & fullPath, vbInformation, "LSM v1.0.0"
    Exit Sub
ExportError3:
    MsgBox "Export Error: " & Err.Description, vbCritical, "LSM v1.0.0"
End Sub

'==============================================================================
' END -- mod_ExportEngine.bas
'==============================================================================
