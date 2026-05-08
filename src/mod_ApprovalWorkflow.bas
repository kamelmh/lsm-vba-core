Attribute VB_Name = "mod_ApprovalWorkflow"
' Public Sector PUBLIC SECTOR FIX: 3-Step Approval Workflow
' Algerian Directorate of Education approval chain

Public Enum ApprovalLevel
    magasinier = 1
    comptable = 2
    directeur = 3
End Enum

' Validates transaction by Comptable (Financial validation)
' Sets "Validé par Comptable" + date in Column G
Public Sub ValidateByComptable(transactionID As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim found As Boolean
    
    Set ws = ThisWorkbook.Sheets(mod_Config.SHEET_MOUVEMENTS)
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    found = False
    
    For i = 2 To lastRow
        If CStr(ws.Cells(i, 1).Value) = transactionID Then
            ws.Cells(i, 7).Value = "Validé par Comptable"
            ws.Cells(i, 7).Interior.Color = RGB(255, 255, 200)
            ws.Cells(i, 8).Value = Date
            found = True
            Exit For
        End If
    Next i
    
    If Not found Then
        MsgBox "Transaction ID '" & transactionID & "' not found.", vbExclamation, "Validation Error"
    End If
End Sub

' Approves transaction by Directeur (Final authorization)
' Sets "Approuvé par Directeur" + visa stamp in Column H
Public Sub ApproveByDirecteur(transactionID As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim found As Boolean
    
    Set ws = ThisWorkbook.Sheets(mod_Config.SHEET_MOUVEMENTS)
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    found = False
    
    For i = 2 To lastRow
        If CStr(ws.Cells(i, 1).Value) = transactionID Then
            If ws.Cells(i, 7).Value <> "Validé par Comptable" Then
                MsgBox "Transaction must be validated by Comptable before Directeur approval.", vbExclamation, "Approval Error"
                Exit Sub
            End If
            ws.Cells(i, 8).Value = "Approuvé par Directeur"
            ws.Cells(i, 8).Interior.Color = RGB(200, 255, 200)
            ws.Cells(i, 9).Value = Date
            ws.Cells(i, 9).Font.Bold = True
            found = True
            Exit For
        End If
    Next i
    
    If Not found Then
        MsgBox "Transaction ID '" & transactionID & "' not found.", vbExclamation, "Approval Error"
    End If
End Sub

' Checks current approval status of a transaction
' Returns: magasinier (1), comptable (2), or directeur (3)
Public Function CheckApprovalStatus(transactionID As String) As ApprovalLevel
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    CheckApprovalStatus = magasinier
    
    Set ws = ThisWorkbook.Sheets(mod_Config.SHEET_MOUVEMENTS)
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    
    For i = 2 To lastRow
        If CStr(ws.Cells(i, 1).Value) = transactionID Then
            If ws.Cells(i, 8).Value = "Approuvé par Directeur" Then
                CheckApprovalStatus = directeur
            ElseIf ws.Cells(i, 7).Value = "Validé par Comptable" Then
                CheckApprovalStatus = comptable
            Else
                CheckApprovalStatus = magasinier
            End If
            Exit For
        End If
    Next i
End Function

' Initializes MOUVEMENTS sheet with required approval columns
Public Sub InitializeMouvementsColumns()
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(mod_Config.SHEET_MOUVEMENTS)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "MOUVEMENTS sheet not found.", vbExclamation, "Setup Error"
        Exit Sub
    End If
    
    ' Column G: Validation Comptable (date)
    ws.Cells(1, 7).Value = "Validation Comptable"
    ws.Cells(1, 7).Interior.Color = RGB(255, 255, 200)
    
    ' Column H: Approbation Directeur (date)
    ws.Cells(1, 8).Value = "Approbation Directeur"
    ws.Cells(1, 8).Interior.Color = RGB(200, 255, 200)
    
    ' Column I: Statut (En attente/Validé/Approuvé)
    ws.Cells(1, 9).Value = "Statut"
    ws.Cells(1, 9).Interior.Color = RGB(220, 220, 255)
    
    ' Set header row style
    Dim lastCol As Long
    lastCol = 9
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
End Sub
