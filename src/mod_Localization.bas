Attribute VB_Name = "mod_Localization"
'=======================================================================================
' MODULE: mod_Localization.bas
' PROJECT: ERP Acad�mie v13
' Localized strings for Arabic RTL support
'=======================================================================================
Option Explicit

' ================================================================================
' API DECLARATIONS (for Unicode/Arabic support)
' ================================================================================
#If VBA7 Then
    Private Declare PtrSafe Function MessageBoxW Lib "user32" ( _
        ByVal hwnd As LongPtr, _
        ByVal lpText As LongPtr, _
        ByVal lpCaption As LongPtr, _
        ByVal uType As Long) As Long
#Else
    Private Declare Function MessageBoxW Lib "user32" ( _
        ByVal hwnd As Long, _
        ByVal lpText As Long, _
        ByVal lpCaption As Long, _
        ByVal uType As Long) As Long
#End If

' ================================================================================
' FUNCTION: GetLocalizedString
' Returns Arabic text from SYS_STRINGS sheet by STRING_ID
' ================================================================================
Public Function GetLocalizedString(ByVal stringID As String) As String
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(mod_Config.SHEET_SYS_STRINGS)
    
    If ws Is Nothing Then
        GetLocalizedString = stringID
        Exit Function
    End If
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    
    Dim i As Long
    For i = 2 To lastRow
        If Trim(ws.Cells(i, 1).Value) = stringID Then
            GetLocalizedString = Trim(ws.Cells(i, 2).Value)
            Exit Function
        End If
    Next i
    
    GetLocalizedString = stringID
    On Error GoTo 0
End Function

' Compatibility wrapper for SafeGetTxt used in forms
Public Function SafeGetTxt(ByVal strID As String) As String
    SafeGetTxt = GetLocalizedString(strID)
End Function

' ================================================================================
' SUB: ShowLocalizedMessage
' Shows a localized message box in Arabic (Unicode supported)
' ================================================================================
Public Sub ShowLocalizedMessage(ByVal stringID As String, _
                         Optional ByVal msgType As VbMsgBoxStyle = vbInformation, _
                         Optional ByVal title As String = "")
    Dim msg As String
    msg = GetLocalizedString(stringID)
    
    If title = "" Then
        title = GetLocalizedString("SYS_TITLE")
        If title = "SYS_TITLE" Then title = mod_Config.SYS_TITLE
    End If
    
    ' Use Unicode MessageBox to prevent Arabic garbling
    UnicodeMsgBox msg, msgType, title
End Sub

' ================================================================================
' SUB: UnicodeMsgBox
' Wrapper for Windows API MessageBoxW to support Unicode/Arabic
' ================================================================================
Public Sub UnicodeMsgBox(ByVal msg As String, _
                         Optional ByVal msgType As VbMsgBoxStyle = vbInformation, _
                         Optional ByVal title As String = "")
    ' We use StrPtr to pass the pointer to the Unicode string (BSTR)
    MessageBoxW 0, StrPtr(msg), StrPtr(title), CLng(msgType)
End Sub
