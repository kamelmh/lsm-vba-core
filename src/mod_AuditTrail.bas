Attribute VB_Name = "mod_AuditTrail"
'=======================================================================================
' MODULE: mod_AuditTrail.bas
' PROJECT: ERP Acad" & Chr(233) & "mie v13
' DESCRIPTION: Handles secure, immutable logging of all system transactions.
'=======================================================================================
Option Explicit

' Records a system event to the audit trail.
Public Sub LogTransaction(ByVal ActionType As String, ByVal RefNum As String)
    Dim wsAudit As Worksheet
    Dim nextRow As Long
    Dim userName As String
    
    On Error GoTo ErrorHandler

    On Error Resume Next
    Set wsAudit = ThisWorkbook.Sheets("AUDIT_LOG")
    On Error GoTo ErrorHandler
    
    If wsAudit Is Nothing Then
        Err.Raise vbObjectError + 513, "mod_AuditTrail", "Critical Error: AUDIT_LOG sheet not found."
    End If

    userName = mod_SharedEnvironment.GetCurrentUserName
    If Len(userName) = 0 Then userName = Environ("USERNAME")  ' Fallback if session not initialized
    wsAudit.Unprotect Password:=mod_Config.MASTER_PWD

    nextRow = wsAudit.Cells(wsAudit.Rows.count, 1).End(xlUp).Row + 1

    With wsAudit
        .Cells(nextRow, 1).Value = Date
        .Cells(nextRow, 2).Value = Format(Now, "HH:mm:ss")
        .Cells(nextRow, 3).Value = userName
        .Cells(nextRow, 4).Value = ActionType
        .Cells(nextRow, 5).Value = RefNum
        .Cells(nextRow, 1).NumberFormat = "yyyy-mm-dd"
        .Cells(nextRow, 2).NumberFormat = "HH:mm:ss"
    End With

    wsAudit.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True

CleanExit:
    Exit Sub

ErrorHandler:
    If Not wsAudit Is Nothing Then wsAudit.Protect Password:=mod_Config.MASTER_PWD
    MsgBox "Audit Logging Failed: " & Err.Description, vbCritical, mod_Config.SYS_TITLE
    Resume CleanExit
End Sub

' Utility to clear logs (Administrative use only).
Public Sub ClearAuditLogs()
    If MsgBox("WARNING: This will permanently delete the audit trail. Proceed?", vbYesNo + vbCritical, "Admin Access") = vbYes Then
        Dim wsAudit As Worksheet: Set wsAudit = ThisWorkbook.Sheets("AUDIT_LOG")
        wsAudit.Unprotect Password:=mod_Config.MASTER_PWD
        wsAudit.Rows("2:" & wsAudit.Rows.count).ClearContents
        wsAudit.Protect Password:=mod_Config.MASTER_PWD
        MsgBox "Audit logs cleared successfully.", vbInformation
    End If
End Sub

' Returns True if AUDIT_LOG sheet exists and is accessible
Public Function AuditLogInitialized() As Boolean
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("AUDIT_LOG")
    AuditLogInitialized = Not (ws Is Nothing)
    On Error GoTo 0
End Function

' Standard audit log entry with full context
Public Sub LogAction(ByVal category As String, ByVal details As String, Optional ByVal moduleName As String = "", Optional ByVal procName As String = "")
    On Error Resume Next
    Dim wsAudit As Worksheet
    Set wsAudit = ThisWorkbook.Sheets("AUDIT_LOG")
    If wsAudit Is Nothing Then Exit Sub

    Dim nextRow As Long
    wsAudit.Unprotect Password:=mod_Config.MASTER_PWD
    nextRow = wsAudit.Cells(wsAudit.Rows.count, 1).End(xlUp).Row + 1

    Dim userName As String
    userName = mod_SharedEnvironment.GetCurrentUserName
    If Len(userName) = 0 Then userName = Environ("USERNAME")

    With wsAudit
        .Cells(nextRow, 1).Value = Date
        .Cells(nextRow, 2).Value = Format(Now, "HH:mm:ss")
        .Cells(nextRow, 3).Value = userName
        .Cells(nextRow, 4).Value = IIf(moduleName <> "", moduleName & "." & procName, category)
        .Cells(nextRow, 5).Value = details
        .Cells(nextRow, 1).NumberFormat = "yyyy-mm-dd"
        .Cells(nextRow, 2).NumberFormat = "HH:mm:ss"
    End With

    wsAudit.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    On Error GoTo 0
End Sub
