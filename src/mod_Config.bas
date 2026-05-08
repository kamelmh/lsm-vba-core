Attribute VB_Name = "mod_Config"
Option Explicit

' Sheet names
Public Const SHEET_MOUVEMENTS As String = "MOUVEMENTS"
Public Const SHEET_ARTICLES As String = "ARTICLES"
Public Const SHEET_SYS_STRINGS As String = "SYS_STRINGS"
Public Const SHEET_AUDIT_LOG As String = "AUDIT_LOG"
Public Const SHEET_STAGING As String = "STAGING_BUFFER"
Public Const SHEET_FOURNISSEURS As String = "FOURNISSEURS"
Public Const SHEET_ACCUEIL As String = "ACCUEIL"

' Document types
Public Const DOC_TYPE_BS As String = "Bon de Sortie"
Public Const DOC_TYPE_DA As String = "Demande d'Achat"

' System constants
Public Const WORKING_DAYS_PER_YEAR As Integer = 250
Public Const OBSERVATION_DAYS As Integer = 38

' Properties (must come after all Const declarations)
Public Property Get SYS_TITLE() As String
    SYS_TITLE = "ERP Acad" & Chr(233) & "mie"
End Property

Public Property Get DOC_TYPE_BR() As String
    DOC_TYPE_BR = "Bon de R" & Chr(201) & "ception"
End Property

Public Property Get DOC_TYPE_BC() As String
    DOC_TYPE_BC = "Bon de Commande"
End Property

Public Property Get MASTER_PWD() As String
    MASTER_PWD = "[YOUR_MASTER_PASSWORD]"
End Property

Public Property Get APP_VERSION() As String
    APP_VERSION = "v1.0.0"
End Property
