Attribute VB_Name = "mod_Navigation"
'==============================================================================
' modNavigation.bas - ERP LSM v1.0.0
' Sheet navigation and form launching
'==============================================================================
Option Explicit

Public Sub OpenStockForm()
    On Error GoTo ErrorHandler
    Dim frmName As String
    frmName = "frmStockEntry"

    If FormExists(frmName) Then
        VBA.UserForms.Add(frmName).Show
    Else
        MsgBox "Erreur: Le formulaire de saisie est manquant.", _
               vbCritical, "ERP LSM v1.0.0"
    End If
    Exit Sub

ErrorHandler:
    MsgBox "Erreur lors de l'ouverture: " & Err.Description, _
           vbCritical, "ERP LSM v1.0.0"
End Sub

Public Sub GoToSheet(ByVal sheetName As String)
    On Error GoTo ErrorHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)
    If Not ws Is Nothing Then
        ws.Activate
    End If
    Exit Sub

ErrorHandler:
    MsgBox "Feuille introuvable: " & sheetName, vbExclamation, "ERP LSM v1.0.0"
End Sub

Public Sub ShowWelcome()
    On Error GoTo ErrorHandler
    Dim versionStr As String
    On Error Resume Next
    versionStr = mod_Config.APP_VERSION
    On Error GoTo ErrorHandler
    If versionStr = "" Then versionStr = "v1.0.0"

    MsgBox "ERP LSM " & versionStr & vbCrLf & _
           "[ORGANIZATION] - [CITY]", _
           vbInformation, "Bienvenue"
    Exit Sub

ErrorHandler:
    MsgBox "Erreur: " & Err.Description, vbCritical, "ERP LSM v1.0.0"
End Sub

Private Function FormExists(ByVal formName As String) As Boolean
    Dim vbComp As Object
    On Error Resume Next
    Set vbComp = ThisWorkbook.VBProject.VBComponents(formName)
    FormExists = Not (vbComp Is Nothing)
    On Error GoTo 0
End Function
