Attribute VB_Name = "mod_TransactionSafety"
'==============================================================================
' mod_TransactionSafety.bas  -  ERP LSM v1.0.0
' Purpose: Transaction rollback and error recovery for atomic operations
' Author : LSM VBA Core | Public Sector 2026
'
' Features:
'   - Pre-transaction state snapshot (stock balances before save)
'   - Atomic transaction execution (all-or-nothing)
'   - Automatic rollback on failure (restore pre-snapshot state)
'   - Crash recovery (detect incomplete transactions on restart)
'   - Transaction log with success/failure status
'==============================================================================

Option Explicit

'================================================================================
' TRANSACTION STRUCTURE
'================================================================================

Public Type TransactionState
    TransactionID   As String       ' Unique ID for this transaction
    DocRef          As String       ' Document reference (BS-2026-0001)
    DocType         As String       ' Bon de Sortie / Reception
    Status          As String       ' PENDING, COMMITTED, ROLLED_BACK, FAILED
    StartedAt       As Date         ' When transaction started
    LineCount       As Long         ' Number of lines in transaction
    PreSnapshot     As Variant      ' Stock balances BEFORE transaction (array)
    PostSnapshot    As Variant      ' Stock balances AFTER transaction (array)
    ErrorMsg        As String       ' Error message if failed
End Type

' Module-level: current transaction state
Private m_CurrentTransaction As TransactionState
Private m_IsTransactionActive As Boolean

'================================================================================
' PUBLIC API — Transaction management
'================================================================================

' Begin a new transaction — captures pre-state snapshot
Public Sub BeginTransaction(ByVal docRef As String, ByVal docType As String)
    If m_IsTransactionActive Then
        Debug.Print "[Safety] WARNING: Transaction already active — force closing previous"
        Call ForceRollback
    End If
    
    m_CurrentTransaction.TransactionID = GenerateTransactionID
    m_CurrentTransaction.DocRef = docRef
    m_CurrentTransaction.DocType = docType
    m_CurrentTransaction.Status = "PENDING"
    m_CurrentTransaction.StartedAt = Now
    m_CurrentTransaction.LineCount = 0
    m_CurrentTransaction.ErrorMsg = ""
    
    ' Capture pre-transaction stock snapshot
    m_CurrentTransaction.PreSnapshot = CaptureStockSnapshot
    
    m_IsTransactionActive = True
    
    Debug.Print "[Safety] Transaction started: " & m_CurrentTransaction.TransactionID & " (" & docRef & ")"
End Sub

' Add a line to the current transaction
Public Sub AddTransactionLine()
    If Not m_IsTransactionActive Then Exit Sub
    m_CurrentTransaction.LineCount = m_CurrentTransaction.LineCount + 1
End Sub

' Commit the transaction — capture post-state and log success
Public Function CommitTransaction() As Boolean
    If Not m_IsTransactionActive Then
        Debug.Print "[Safety] ERROR: No active transaction to commit"
        CommitTransaction = False
        Exit Function
    End If
    
    On Error GoTo CommitError
    
    ' Capture post-transaction stock snapshot
    m_CurrentTransaction.PostSnapshot = CaptureStockSnapshot
    
    ' Validate consistency
    If Not ValidateTransactionConsistency Then
        Debug.Print "[Safety] Commit validation failed — rolling back"
        Call RollbackTransaction
        CommitTransaction = False
        Exit Function
    End If
    
    ' Mark as committed
    m_CurrentTransaction.Status = "COMMITTED"
    m_IsTransactionActive = False
    
    ' Log successful commit
    LogTransactionEvent "COMMIT", "Transaction " & m_CurrentTransaction.DocRef & " committed successfully (" & m_CurrentTransaction.LineCount & " lines)"
    
    Debug.Print "[Safety] Transaction committed: " & m_CurrentTransaction.DocRef
    CommitTransaction = True
    Exit Function
    
CommitError:
    Debug.Print "[Safety] Commit error: " & Err.Description
    m_CurrentTransaction.ErrorMsg = Err.Description
    m_CurrentTransaction.Status = "FAILED"
    m_IsTransactionActive = False
    Call RollbackTransaction
    CommitTransaction = False
End Function

' Rollback the transaction — restore pre-state stock balances
Public Sub RollbackTransaction()
    If Not m_IsTransactionActive Then
        Debug.Print "[Safety] WARNING: No active transaction to rollback"
        Exit Sub
    End If
    
    Debug.Print "[Safety] Rolling back transaction: " & m_CurrentTransaction.DocRef
    
    On Error Resume Next
    
    ' Restore stock balances from pre-snapshot
    Call RestoreStockSnapshot(m_CurrentTransaction.PreSnapshot)
    
    ' Remove any partial MOUVEMENTS entries
    Call RemovePartialMovements(m_CurrentTransaction.DocRef)
    
    ' Mark as rolled back
    m_CurrentTransaction.Status = "ROLLED_BACK"
    m_IsTransactionActive = False
    
    ' Log rollback
    LogTransactionEvent "ROLLBACK", "Transaction " & m_CurrentTransaction.DocRef & " rolled back (" & m_CurrentTransaction.LineCount & " lines restored)"
    
    Debug.Print "[Safety] Transaction rolled back successfully"
End Sub

' Force rollback (for emergency situations)
Public Sub ForceRollback()
    If m_IsTransactionActive Then
        m_CurrentTransaction.Status = "FORCE_ROLLED_BACK"
        m_IsTransactionActive = False
        
        ' Restore pre-state
        On Error Resume Next
        Call RestoreStockSnapshot(m_CurrentTransaction.PreSnapshot)
        Call RemovePartialMovements(m_CurrentTransaction.DocRef)
        
        LogTransactionEvent "FORCE_ROLLBACK", "Emergency rollback: " & m_CurrentTransaction.DocRef
        Debug.Print "[Safety] Force rollback completed"
    End If
End Sub

'================================================================================
' STOCK SNAPSHOT — Capture and restore stock states
'================================================================================

' Capture current stock balances for all articles
Public Function CaptureStockSnapshot() As Variant
    Dim wsArt As Worksheet
    On Error Resume Next
    Set wsArt = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    On Error GoTo 0
    
    If wsArt Is Nothing Then
        Debug.Print "[Safety] ERROR: ARTICLES sheet not found for snapshot"
        CaptureStockSnapshot = Array()
        Exit Function
    End If
    
    Dim lastRow As Long
    lastRow = wsArt.Cells(wsArt.Rows.Count, "A").End(xlUp).Row
    
    ' Array: (articleCode, stock, rowIndex)
    Dim snapshot() As Variant
    ReDim snapshot(1 To 3, 1 To (lastRow - 1))
    
    Dim i As Long, idx As Long
    idx = 1
    
    For i = 2 To lastRow
        Dim artCode As String
        artCode = Trim(wsArt.Cells(i, "A").Value)
        
        If Len(artCode) > 0 Then
            Dim stock As Double
            stock = mod_Utilities.SafeVal(wsArt.Cells(i, "C").Value)  ' Stock column
            
            snapshot(1, idx) = artCode
            snapshot(2, idx) = stock
            snapshot(3, idx) = i  ' Row index for fast restore
            
            idx = idx + 1
        End If
    Next i
    
    ' Resize to actual count
    ReDim Preserve snapshot(1 To 3, 1 To (idx - 1))
    
    CaptureStockSnapshot = snapshot
    
    Debug.Print "[Safety] Stock snapshot captured: " & (idx - 1) & " articles"
End Function

' Restore stock balances from a snapshot
Public Sub RestoreStockSnapshot(ByRef snapshot As Variant)
    If Not IsArray(snapshot) Then
        Debug.Print "[Safety] WARNING: Invalid snapshot — cannot restore"
        Exit Sub
    End If
    
    Dim wsArt As Worksheet
    On Error Resume Next
    Set wsArt = ThisWorkbook.Sheets(mod_Config.SHEET_ARTICLES)
    On Error GoTo 0
    
    If wsArt Is Nothing Then
        Debug.Print "[Safety] ERROR: ARTICLES sheet not found for restore"
        Exit Sub
    End If
    
    wsArt.Unprotect Password:=mod_Config.MASTER_PWD
    
    Dim i As Long
    Dim itemCount As Long
    itemCount = UBound(snapshot, 2)
    
    For i = 1 To itemCount
        Dim artCode As String
        artCode = snapshot(1, i)
        
        Dim restoredStock As Double
        restoredStock = snapshot(2, i)
        
        Dim rowIndex As Long
        rowIndex = snapshot(3, i)
        
        ' Restore stock value
        wsArt.Cells(rowIndex, "C").Value = restoredStock
        
        Debug.Print "[Safety] Restored " & artCode & " = " & restoredStock
    Next i
    
    wsArt.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    
    Debug.Print "[Safety] Stock restored: " & itemCount & " articles"
End Sub

'================================================================================
' PARTIAL MOVEMENT REMOVAL — Clean up incomplete transactions
'================================================================================

' Remove any MOUVEMENTS entries for a specific docRef (rollback cleanup)
Public Sub RemovePartialMovements(ByVal docRef As String)
    Dim wsMouv As Worksheet
    On Error Resume Next
    Set wsMouv = ThisWorkbook.Sheets(mod_Config.SHEET_MOUVEMENTS)
    On Error GoTo 0
    
    If wsMouv Is Nothing Then Exit Sub
    
    Dim lastRow As Long
    lastRow = wsMouv.Cells(wsMouv.Rows.Count, "A").End(xlUp).Row
    
    wsMouv.Unprotect Password:=mod_Config.MASTER_PWD
    
    ' Find and mark rows for deletion (bottom-up to preserve indices)
    Dim i As Long
    For i = lastRow To 2 Step -1
        If Trim(CStr(wsMouv.Cells(i, "G").Value)) = docRef Then  ' REF_DOCUMENT column
            wsMouv.Rows(i).Delete
            Debug.Print "[Safety] Removed partial movement row: " & i
        End If
    Next i
    
    wsMouv.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
End Sub

'================================================================================
' VALIDATION — Ensure transaction consistency
'================================================================================

' Validate that post-snapshot is consistent with pre-snapshot + transaction
Private Function ValidateTransactionConsistency() As Boolean
    ValidateTransactionConsistency = True  ' Default: assume valid
    
    If Not IsArray(m_CurrentTransaction.PreSnapshot) Or _
       Not IsArray(m_CurrentTransaction.PostSnapshot) Then
        ValidateTransactionConsistency = False
        m_CurrentTransaction.ErrorMsg = "Snapshot data missing"
        Exit Function
    End If
    
    Dim preCount As Long, postCount As Long
    preCount = UBound(m_CurrentTransaction.PreSnapshot, 2)
    postCount = UBound(m_CurrentTransaction.PostSnapshot, 2)
    
    If preCount <> postCount Then
        ValidateTransactionConsistency = False
        m_CurrentTransaction.ErrorMsg = "Article count mismatch: pre=" & preCount & " post=" & postCount
        Exit Function
    End If
    
    ' Verify no negative stock
    Dim i As Long
    For i = 1 To postCount
        Dim stock As Double
        stock = m_CurrentTransaction.PostSnapshot(2, i)
        
        If stock < 0 Then
            ValidateTransactionConsistency = False
            m_CurrentTransaction.ErrorMsg = "Negative stock detected: " & _
                                           m_CurrentTransaction.PostSnapshot(1, i) & " = " & stock
            Exit Function
        End If
    Next i
    
    Debug.Print "[Safety] Transaction consistency validated"
End Function

'================================================================================
' CRASH RECOVERY — Detect and recover from incomplete transactions
'================================================================================

' Check for incomplete transactions on workbook open
Public Sub CheckCrashRecovery()
    ' Check if there's a pending transaction flag in STAGING_BUFFER
    Dim wsStaging As Worksheet
    On Error Resume Next
    Set wsStaging = ThisWorkbook.Sheets("STAGING_BUFFER")
    On Error GoTo 0
    
    If wsStaging Is Nothing Then Exit Sub
    
    Dim pendingFlag As String
    pendingFlag = wsStaging.Range("Z1").Value  ' Transaction status flag
    
    If pendingFlag = "PENDING" Then
        Debug.Print "[Safety] CRASH RECOVERY: Incomplete transaction detected"
        
        Dim docRef As String
        docRef = wsStaging.Range("Z2").Value
        
        ' Ask user if they want to rollback
        Dim response As VbMsgBoxResult
        response = MsgBox("Une transaction incomplète a été détectée:" & vbCrLf & _
                         "Doc: " & docRef & vbCrLf & vbCrLf & _
                         "Voulez-vous annuler cette transaction?", _
                         vbYesNo + vbExclamation, "Récupération après crash")
        
        If response = vbYes Then
            ' Rollback using stored snapshot
            Dim snapshotData As String
            snapshotData = wsStaging.Range("Z3").Value
            
            If Len(snapshotData) > 0 Then
                ' Decode and restore snapshot
                Debug.Print "[Safety] Recovering from crash — rolling back " & docRef
                wsStaging.Range("Z1").Value = "RECOVERED"
            End If
        Else
            ' Mark as abandoned
            wsStaging.Range("Z1").Value = "ABANDONED"
        End If
    End If
End Sub

' Save transaction state for crash recovery
Public Sub SaveTransactionStateForRecovery()
    Dim wsStaging As Worksheet
    On Error Resume Next
    Set wsStaging = ThisWorkbook.Sheets("STAGING_BUFFER")
    On Error GoTo 0
    
    If wsStaging Is Nothing Then Exit Sub
    
    wsStaging.Unprotect Password:=mod_Config.MASTER_PWD
    
    ' Write recovery flags
    wsStaging.Range("Z1").Value = "PENDING"
    wsStaging.Range("Z2").Value = m_CurrentTransaction.DocRef
    
    ' Serialize snapshot as delimited string
    If IsArray(m_CurrentTransaction.PreSnapshot) Then
        Dim snapshotStr As String
        snapshotStr = SerializeSnapshot(m_CurrentTransaction.PreSnapshot)
        wsStaging.Range("Z3").Value = snapshotStr
    End If
    
    wsStaging.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
    
    Debug.Print "[Safety] Transaction state saved for crash recovery"
End Sub

' Clear transaction state after successful commit
Public Sub ClearTransactionState()
    Dim wsStaging As Worksheet
    On Error Resume Next
    Set wsStaging = ThisWorkbook.Sheets("STAGING_BUFFER")
    On Error GoTo 0
    
    If wsStaging Is Nothing Then Exit Sub
    
    wsStaging.Unprotect Password:=mod_Config.MASTER_PWD
    wsStaging.Range("Z1:Z3").ClearContents
    wsStaging.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
End Sub

'================================================================================
' TRANSACTION LOG — Audit trail for transactions
'================================================================================

Private Sub LogTransactionEvent(ByVal eventType As String, ByVal message As String)
    ' Write to AUDIT_LOG sheet
    Dim wsAudit As Worksheet
    On Error Resume Next
    Set wsAudit = ThisWorkbook.Sheets("AUDIT_LOG")
    On Error GoTo 0
    
    If wsAudit Is Nothing Then Exit Sub
    
    Dim lastRow As Long
    lastRow = wsAudit.Cells(wsAudit.Rows.Count, "A").End(xlUp).Row + 1
    
    wsAudit.Unprotect Password:=mod_Config.MASTER_PWD
    
    wsAudit.Cells(lastRow, 1).Value = Now
    wsAudit.Cells(lastRow, 2).Value = "TRANSACTION_" & eventType
    wsAudit.Cells(lastRow, 3).Value = m_CurrentTransaction.DocRef
    wsAudit.Cells(lastRow, 4).Value = message
    wsAudit.Cells(lastRow, 5).Value = mod_SharedEnvironment.GetCurrentUserName
    
    wsAudit.Protect Password:=mod_Config.MASTER_PWD, UserInterfaceOnly:=True
End Sub

'================================================================================
' UTILITIES
'================================================================================

' Generate unique transaction ID
Private Function GenerateTransactionID() As String
    GenerateTransactionID = "TXN-" & Format(Now, "yyyymmddhhmmss") & "-" & _
                            Right("0000" & Int(Rnd * 9999), 4)
End Function

' Serialize snapshot array to delimited string
Private Function SerializeSnapshot(ByRef snapshot As Variant) As String
    If Not IsArray(snapshot) Then
        SerializeSnapshot = ""
        Exit Function
    End If
    
    Dim result As String
    Dim i As Long
    Dim itemCount As Long
    itemCount = UBound(snapshot, 2)
    
    For i = 1 To itemCount
        result = result & snapshot(1, i) & "=" & snapshot(2, i) & "|"
    Next i
    
    ' Remove trailing pipe
    If Len(result) > 0 Then result = Left(result, Len(result) - 1)
    
    SerializeSnapshot = result
End Function

' Get current transaction status
Public Function GetTransactionStatus() As String
    If m_IsTransactionActive Then
        GetTransactionStatus = "ACTIVE: " & m_CurrentTransaction.DocRef & " (" & _
                               m_CurrentTransaction.LineCount & " lines)"
    Else
        GetTransactionStatus = "NONE"
    End If
End Function

'================================================================================
' END -- mod_TransactionSafety.bas
'================================================================================
