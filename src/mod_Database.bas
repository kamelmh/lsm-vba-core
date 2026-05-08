Attribute VB_Name = "mod_Database"
'=======================================================================================
' MODULE: mod_Database.bas
' PROJECT: ERP Acad" & Chr(233) & "mie v13
' DESCRIPTION: Secure database access layer for transaction writes.
'=======================================================================================
Option Explicit

Public Sub SecureWriteTransaction(docDate As Date, _
                                  typeSign As String, _
                                  refDoc As String, _
                                  codeArticle As String, _
                                  designation As String, _
                                  quantity As Long, _
                                  unitPrice As Double, _
                                  lineValue As Double, _
                                  thirdParty As String, _
                                  Optional notes As String)
    Dim ws As Worksheet
    Dim nextRow As Long
    
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Set ws = ThisWorkbook.Sheets(mod_Config.SHEET_MOUVEMENTS)
    ws.Unprotect Password:=mod_Config.MASTER_PWD
    
    nextRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row + 1
    
    ws.Cells(nextRow, 1).Value = docDate
    ws.Cells(nextRow, 2).Value = codeArticle
    ws.Cells(nextRow, 3).Value = designation
    ws.Cells(nextRow, 4).Value = typeSign
    ws.Cells(nextRow, 5).Value = quantity
    ws.Cells(nextRow, 6).Value = lineValue
    ws.Cells(nextRow, 7).Value = refDoc
    ws.Cells(nextRow, 8).Value = unitPrice
    ws.Cells(nextRow, 9).Value = thirdParty
    ws.Cells(nextRow, 12).Value = notes
    
    ws.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    
CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Exit Sub
    
ErrorHandler:
    On Error Resume Next
    If Not ws Is Nothing Then ws.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    Resume CleanUp
End Sub
