VERSION 5.00
Begin VB.Form frmIcon 
   Caption         =   "Настройки на UcsHPHub"
   ClientHeight    =   5868
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8148
   Icon            =   "frmIcon.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5868
   ScaleWidth      =   8148
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuSysTray 
      Caption         =   "UcsFPHub"
      Visible         =   0   'False
      Begin VB.Menu mnuPopup 
         Caption         =   "Настройки"
         Index           =   0
      End
      Begin VB.Menu mnuPopup 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuPopup 
         Caption         =   "Рестарт"
         Index           =   2
      End
      Begin VB.Menu mnuPopup 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuPopup 
         Caption         =   "Изход"
         Index           =   4
      End
   End
End
Attribute VB_Name = "frmIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================
'
' UcsFPHub (c) 2019 by Unicontsoft
'
' Unicontsoft Fiscal Printers Hub
'
' This project is licensed under the terms of the MIT license
' See the LICENSE file in the project root for more information
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "frmIcon"

'=========================================================================
' API
'=========================================================================

'--- for ShellExecuteEx
Private Const SEE_MASK_NOASYNC              As Long = &H100
Private Const SEE_MASK_FLAG_NO_UI           As Long = &H400

Private Declare Function ShellExecuteEx Lib "shell32" Alias "ShellExecuteExA" (lpExecInfo As SHELLEXECUTEINFO) As Long

Private Type SHELLEXECUTEINFO
    cbSize              As Long
    fMask               As Long
    hWnd                As Long
    lpVerb              As String
    lpFile              As String
    lpParameters        As String
    lpDirectory         As Long
    nShow               As Long
    hInstApp            As Long
    '  optional fields
    lpIDList            As Long
    lpClass             As Long
    hkeyClass           As Long
    dwHotKey            As Long
    hIcon               As Long
    hProcess            As Long
End Type
'=========================================================================
' Constants and member variables
'=========================================================================

Private WithEvents m_oSysTray       As cSysTray
Attribute m_oSysTray.VB_VarHelpID = -1

Private Enum UcsMenuItems
    ucsMnuSettings
    ucsMnuSep1
    ucsMnuRestart
    ucsMnuSep2
    ucsMnuExit
End Enum

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunction As String)
    Debug.Print "Critical error: " & Err.Description & " [" & MODULE_NAME & "." & sFunction & "]"
    DebugLog Err.Description & " [" & MODULE_NAME & "." & sFunction & "]", vbLogEventTypeError
End Sub

'=========================================================================
' Methods
'=========================================================================

Public Function Init(sProductName As String) As Boolean
    Set m_oSysTray = New cSysTray
    m_oSysTray.Init Me, sProductName
    '--- success
    Init = True
End Function

Public Sub Restart(Optional AddParam As String)
    Dim uShell          As SHELLEXECUTEINFO
    
    TerminateEndpoints
    FlushDebugLog
    With uShell
        .cbSize = Len(uShell)
        .fMask = SEE_MASK_NOASYNC Or SEE_MASK_FLAG_NO_UI
        .lpFile = GetProcessName()
        .lpParameters = Trim$(Command$ & " " & ArgvQuote(AddParam))
    End With
    Call ShellExecuteEx(uShell)
    Unload Me
End Sub

'=========================================================================
' Methods
'=========================================================================

Private Sub mnuPopup_Click(Index As Integer)
    Const FUNC_NAME     As String = "mnuPopup_Click"
    
    On Error GoTo EH
    Select Case Index
    Case ucsMnuSettings
        Visible = True
        SetFocus
    Case ucsMnuRestart
        Restart
    Case ucsMnuExit
        Unload Me
    End Select
QH:
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume QH
End Sub

Private Sub m_oSysTray_Click()
    Const FUNC_NAME     As String = "m_oSysTray_Click"
    
    On Error GoTo EH
    Visible = True
QH:
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume QH
End Sub

Private Sub m_oSysTray_ContextMenu()
    Const FUNC_NAME     As String = "m_oSysTray_ContextMenu"
    
    On Error GoTo EH
    PopupMenu mnuSysTray
QH:
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume QH
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Const FUNC_NAME     As String = "Form_QueryUnload"
    
    On Error GoTo EH
    If UnloadMode = vbFormControlMenu Then
        Visible = False
        Cancel = 1
    End If
QH:
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume QH
End Sub
