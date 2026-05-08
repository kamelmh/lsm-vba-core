Attribute VB_Name = "mod_SharedEnvironment"
'==============================================================================
' mod_SharedEnvironment.bas  -  ERP LSM v1.0.0

' Depends: mod_Config, mod_Database, mod_AuditTrail
' Note: mod_ExportEngine dependency removed to break circular reference
'
' Features:
'   - Configurable shared export directory (network or local)
'   - User session tracking (login/logout with role)
'   - File lock detection (warn if opened by another user)
'   - Batch PDF export (all pending documents)
'   - Bulk movement import from Excel/CSV
'   - Auto-backup scheduling
'==============================================================================

Option Explicit

'================================================================================
' USER SESSION TYPE - Must be declared before any procedures
'================================================================================

Public Type UserSession
    UserName        As String       ' Windows username
    DisplayName     As String       ' Full name
    Role            As String       ' Magasinier, Comptable, Directeur
    LoginTime       As Date         ' When session started
    LastActivity    As Date         ' Last action timestamp
    IsActive        As Boolean      ' Whether session is active
End Type

Private m_CurrentUser As UserSession

'================================================================================
' SHARED PATHS - Configurable directories
'================================================================================

' Returns the shared export directory (network or local)
Public Function GetSharedExportPath() As String
    Dim sharedPath As String
    
    ' Try to read from STAGING_BUFFER first (configurable by admin)
    On Error Resume Next
    Dim wsStaging As Worksheet
    Set wsStaging = ThisWorkbook.Sheets("STAGING_BUFFER")
    sharedPath = wsStaging.Range("Y1").Value
    On Error GoTo 0
    
    If Len(sharedPath) > 0 And Dir(sharedPath, vbDirectory) <> "" Then
        GetSharedExportPath = sharedPath
        Exit Function
    End If
    
    ' Fallback: User-specific Documents folder (more reliable than Desktop)
    GetSharedExportPath = Environ("USERPROFILE") & "\Documents\LSM_Export\"
    
    ' Create folder if it doesn't exist
    If Dir(GetSharedExportPath, vbDirectory) = "" Then
        MkDir GetSharedExportPath
    End If
End Function

' Returns the shared backup directory
Public Function GetSharedBackupPath() As String
    GetSharedBackupPath = Environ("USERPROFILE") & "\Documents\LSM_Backup\"
    
    If Dir(GetSharedBackupPath, vbDirectory) = "" Then
        MkDir GetSharedBackupPath
    End If
End Function

' Set shared export path (for admin use)
Public Sub SetSharedExportPath(ByVal newPath As String)
    On Error Resume Next
    Dim wsStaging As Worksheet
    Set wsStaging = ThisWorkbook.Sheets("STAGING_BUFFER")
    
    If Dir(newPath, vbDirectory) <> "" Then
        wsStaging.Unprotect Password:=mod_Config.MASTER_PWD
        wsStaging.Range("Y1").Value = newPath
        wsStaging.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
        Debug.Print "[SharedEnv] Export path set to: " & newPath
    Else
        MsgBox "Le dossier n'existe pas: " & newPath, vbExclamation
    End If
    On Error GoTo 0
End Sub

' Initialize user session (called on Workbook_Open)
Public Sub InitUserSession()
    m_CurrentUser.UserName = Environ("USERNAME")
    m_CurrentUser.DisplayName = Application.UserName
    m_CurrentUser.LoginTime = Now
    m_CurrentUser.LastActivity = Now
    m_CurrentUser.IsActive = True
    
    ' Determine role based on username (simple mapping - can be extended)
    m_CurrentUser.Role = DetectUserRole(m_CurrentUser.UserName)
    
    ' Log session start
    LogSessionEvent "LOGIN", "User " & m_CurrentUser.DisplayName & " (" & m_CurrentUser.Role & ")"
    
    Debug.Print "[Session] User: " & m_CurrentUser.DisplayName & " Role: " & m_CurrentUser.Role
End Sub

' Get current user display name
Public Function GetCurrentUserName() As String
    GetCurrentUserName = m_CurrentUser.DisplayName
End Function

' Get current user role
Public Function GetCurrentUserRole() As String
    GetCurrentUserRole = m_CurrentUser.Role
End Function

' Get current user login time
Public Function GetCurrentSessionTime() As Date
    GetCurrentSessionTime = m_CurrentUser.LoginTime
End Function

' Get current user session (internal use only)
Private Function GetCurrentSession() As UserSession
    GetCurrentSession = m_CurrentUser
End Function

' Update last activity timestamp
Public Sub UpdateUserActivity()
    m_CurrentUser.LastActivity = Now
    
    ' Save to STAGING_BUFFER for audit
    On Error Resume Next
    Dim wsStaging As Worksheet
    Set wsStaging = ThisWorkbook.Sheets("STAGING_BUFFER")
    wsStaging.Unprotect Password:=mod_Config.MASTER_PWD
    wsStaging.Range("Y2").Value = Now
    wsStaging.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    On Error GoTo 0
End Sub

' Detect user role based on username
Private Function DetectUserRole(ByVal userName As String) As String
    ' Simple role mapping - can be extended with a ROLES sheet
    Select Case LCase(userName)
        Case "magasinier", "store", "stock"
            DetectUserRole = "Magasinier"
        Case "comptable", "accounting", "finance"
            DetectUserRole = "Comptable"
        Case "directeur", "director", "admin", "administrator"
            DetectUserRole = "Directeur"
        Case Else
            DetectUserRole = "Magasinier"  ' Default role
    End Select
End Function

' Check if current user has permission for an action
Public Function HasPermission(ByVal requiredRole As String) As Boolean
    Select Case m_CurrentUser.Role
        Case "Directeur"
            HasPermission = True  ' Directeur has all permissions
        Case "Comptable"
            HasPermission = (requiredRole = "Comptable" Or requiredRole = "Magasinier")
        Case "Magasinier"
            HasPermission = (requiredRole = "Magasinier")
        Case Else
            HasPermission = False
    End Select
End Function

'================================================================================
' FILE LOCK DETECTION - Warn if workbook opened by another user
'================================================================================

' Check if workbook is opened by another user
Public Function IsWorkbookLocked() As Boolean
    Dim lockFile As String
    lockFile = ThisWorkbook.Path & "\~$" & ThisWorkbook.Name
    
    IsWorkbookLocked = (Dir(lockFile) <> "")
End Function

' Get lock owner information (simplified - lock files are binary, best effort only)
Public Function GetLockOwner() As String
    On Error Resume Next
    Dim lockFile As String
    lockFile = ThisWorkbook.Path & "\~$" & ThisWorkbook.Name
    
    If Dir(lockFile) <> "" Then
        ' Lock file exists but content is binary - return filename as indicator
        GetLockOwner = "Utilisateur inconnu (fichier: ~$" & ThisWorkbook.Name & ")"
    Else
        GetLockOwner = ""
    End If
    On Error GoTo 0
End Function

' Warn user if workbook is already open
Public Sub CheckWorkbookAccess()
    If IsWorkbookLocked() Then
        Dim owner As String
        owner = GetLockOwner()
        
        Dim response As VbMsgBoxResult
        response = MsgBox("ATTENTION: Ce fichier est déjà ouvert par: " & owner & vbCrLf & _
                         vbCrLf & _
                         "L'ouverture en mode lecture seule est recommandée pour éviter" & vbCrLf & _
                         "la corruption des données. Continuer en mode lecture/écriture?", _
                         vbYesNo + vbExclamation, "Accès concurrent détecté")
        
        If response = vbNo Then
            ThisWorkbook.ChangeFileAccess Mode:=xlReadOnly
            MsgBox "Fichier ouvert en mode lecture seule.", vbInformation
        End If
    End If
End Sub

'================================================================================
' BATCH OPERATIONS - Mass tasking engine
'================================================================================

' Batch PDF Export - Export all documents for a given period
Public Sub BatchExportPDFs(ByVal startDate As Date, ByVal endDate As Date)
    Dim wsMouv As Worksheet
    On Error Resume Next
    Set wsMouv = ThisWorkbook.Sheets(mod_Config.SHEET_MOUVEMENTS)
    On Error GoTo 0
    
    If wsMouv Is Nothing Then
        MsgBox "Feuille MOUVEMENTS introuvable.", vbCritical
        Exit Sub
    End If
    
    Dim lastRow As Long
    lastRow = wsMouv.Cells(wsMouv.Rows.Count, "A").End(xlUp).Row
    
    ' Collect unique document refs within date range
    Dim docRefs As Object
    Set docRefs = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    For i = 2 To lastRow
        Dim mvtDate As Date
        On Error Resume Next
        mvtDate = CDate(wsMouv.Cells(i, "A").Value)
        On Error GoTo 0
        
        If mvtDate >= startDate And mvtDate <= endDate Then
            Dim refDoc As String
            refDoc = Trim(CStr(wsMouv.Cells(i, "G").Value))
            
            If Len(refDoc) > 0 Then
                docRefs(refDoc) = True
            End If
        End If
    Next i
    
    If docRefs.Count = 0 Then
        MsgBox "Aucun document trouvé pour la période sélectionnée.", vbInformation
        Exit Sub
    End If
    
    ' Export each document
    Dim exportPath As String
    exportPath = GetSharedExportPath()
    
    Dim successCount As Long
    Dim failCount As Long
    Dim key As Variant
    
    For Each key In docRefs.Keys
        On Error Resume Next
        Dim pdfPath As String
        pdfPath = exportPath & CStr(key) & "_" & Format(Date, "yyyy-mm-dd") & ".pdf"
        
        ' Silent export via Application.Run to break circular dependency with mod_ExportEngine
        On Error Resume Next
        Dim exportResult As Variant
        exportResult = Application.Run("mod_ExportEngine.ExportTransactionToPDF_Silent", CStr(key), pdfPath)
        If exportResult = True Then
            successCount = successCount + 1
        Else
            failCount = failCount + 1
            Debug.Print "[Batch] Failed: " & key
        End If
        On Error GoTo 0
    Next key
    
    MsgBox "Export par lot terminé:" & vbCrLf & _
           "Succès: " & successCount & vbCrLf & _
           "Échecs: " & failCount, _
           vbInformation, "Export par lot"
    
    Debug.Print "[Batch] Exported " & successCount & " documents (" & failCount & " failed)"
End Sub

' Bulk Movement Import - Import movements from Excel sheet
Public Sub BulkImportMovements(ByVal importSheetName As String)
    Dim wsImport As Worksheet
    On Error Resume Next
    Set wsImport = ThisWorkbook.Sheets(importSheetName)
    On Error GoTo 0
    
    If wsImport Is Nothing Then
        MsgBox "Feuille d'import introuvable: " & importSheetName, vbCritical
        Exit Sub
    End If
    
    Dim lastRow As Long
    lastRow = wsImport.Cells(wsImport.Rows.Count, "A").End(xlUp).Row
    
    If lastRow < 2 Then
        MsgBox "Aucune donnée à importer.", vbInformation
        Exit Sub
    End If
    
    Dim successCount As Long
    Dim failCount As Long
    Dim i As Long
    
    ' Expected columns: Date, Type(IN/OUT), RefDoc, ArticleCode, Qty, PU, ThirdParty
    For i = 2 To lastRow
        Dim mvtDate As Date
        Dim mvtType As String
        Dim refDoc As String
        Dim artCode As String
        Dim qty As Long
        Dim qtyRaw As Double
        Dim pu As Double
        Dim thirdParty As String
        
        On Error Resume Next
        mvtDate = CDate(wsImport.Cells(i, "A").Value)
        mvtType = UCase(Trim(CStr(wsImport.Cells(i, "B").Value)))
        refDoc = Trim(CStr(wsImport.Cells(i, "C").Value))
        artCode = Trim(CStr(wsImport.Cells(i, "D").Value))
        qtyRaw = CDbl(wsImport.Cells(i, "E").Value)
        pu = CDbl(wsImport.Cells(i, "F").Value)
        thirdParty = Trim(CStr(wsImport.Cells(i, "G").Value))
        On Error GoTo 0
        
        qty = CLng(qtyRaw)
        
        ' Validate required fields
        If mvtDate = 0 Or Len(refDoc) = 0 Or Len(artCode) = 0 Or qty = 0 Then
            failCount = failCount + 1
            Debug.Print "[BulkImport] Skip row " & i & " - missing required fields"
            GoTo NextRow
        End If
        
        ' Write transaction
        On Error Resume Next
        mod_Database.SecureWriteTransaction mvtDate, mvtType, refDoc, artCode, "", qty, pu, qty * pu, thirdParty, "BulkImport"
        
        If Err.Number = 0 Then
            successCount = successCount + 1
        Else
            failCount = failCount + 1
            Debug.Print "[BulkImport] Failed row " & i & " - " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
        
NextRow:
    Next i
    
    MsgBox "Import par lot terminé:" & vbCrLf & _
           "Succès: " & successCount & vbCrLf & _
           "Échecs: " & failCount, _
           vbInformation, "Import par lot"
    
    Debug.Print "[BulkImport] Imported " & successCount & " movements (" & failCount & " failed)"
End Sub

'================================================================================
' AUTO-BACKUP - Scheduled backup management
'================================================================================

' Perform automatic backup
Public Sub AutoBackup()
    Dim backupPath As String
    backupPath = GetSharedBackupPath()
    
    Dim fileName As String
    fileName = "LSM_Backup_" & Format(Now, "yyyymmdd_hhmmss") & ".xlsm"
    Dim fullPath As String
    fullPath = backupPath & fileName
    
    On Error Resume Next
    ThisWorkbook.SaveCopyAs fullPath
    
    If Err.Number = 0 Then
        Debug.Print "[AutoBackup] Saved to: " & fullPath
        LogSessionEvent "BACKUP", "Auto-backup: " & fileName
    Else
        Debug.Print "[AutoBackup] Failed: " & Err.Description
    End If
    On Error GoTo 0
End Sub

'================================================================================
' SESSION LOGGING
'================================================================================

Private Sub LogSessionEvent(ByVal eventType As String, ByVal message As String)
    On Error Resume Next
    Dim wsAudit As Worksheet
    Set wsAudit = ThisWorkbook.Sheets("AUDIT_LOG")
    
    If wsAudit Is Nothing Then Exit Sub
    
    Dim lastRow As Long
    lastRow = wsAudit.Cells(wsAudit.Rows.Count, "A").End(xlUp).Row + 1
    
    wsAudit.Unprotect Password:=mod_Config.MASTER_PWD
    
    wsAudit.Cells(lastRow, 1).Value = Now
    wsAudit.Cells(lastRow, 2).Value = "SESSION_" & eventType
    wsAudit.Cells(lastRow, 3).Value = m_CurrentUser.UserName
    wsAudit.Cells(lastRow, 4).Value = message
    wsAudit.Cells(lastRow, 5).Value = m_CurrentUser.Role
    
    wsAudit.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    On Error GoTo 0
End Sub

'================================================================================
' END -- mod_SharedEnvironment.bas
'================================================================================
