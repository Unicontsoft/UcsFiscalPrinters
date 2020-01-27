VERSION 5.00
Begin VB.Form frmIcon 
   Caption         =   "Настройки"
   ClientHeight    =   1776
   ClientLeft      =   192
   ClientTop       =   240
   ClientWidth     =   4944
   Icon            =   "frmIcon.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1776
   ScaleWidth      =   4944
   StartUpPosition =   1  'CenterOwner
   Begin VB.Menu mnuMainPopup 
      Caption         =   "UcsFPHub"
      Visible         =   0   'False
      Begin VB.Menu mnuPopup 
         Caption         =   "Настройки..."
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
' UcsFPHub (c) 2019-2020 by Unicontsoft
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
' Constants and member variables
'=========================================================================

Private m_oPrinters                 As Object
Private m_sConfFile                 As String
Private WithEvents m_oSysTray       As cSysTray
Attribute m_oSysTray.VB_VarHelpID = -1

Private Enum UcsMenuItems
    ucsMnuPopupConfig = 0
    ucsMnuPopupSep1
    ucsMnuPopupRestart
    ucsMnuPopupSep2
    ucsMnuPopupExit
End Enum

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunction As String)
    #If USE_DEBUG_LOG <> 0 Then
        DebugLog MODULE_NAME, sFunction & "(" & Erl & ")", Err.Description & " &H" & Hex$(Err.Number), vbLogEventTypeError
    #Else
        Debug.Print "Critical error: " & Err.Description & " [" & MODULE_NAME & "." & sFunction & "]"
    #End If
End Sub

'=========================================================================
' Properties
'=========================================================================

Property Get Printers() As Object
    Set Printers = m_oPrinters
End Property

Property Get ConfFile() As String
     ConfFile = m_sConfFile
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function Init(oPrinters As Object, sConfFile As String, sProductName As String) As Boolean
    Const FUNC_NAME     As String = "Init"
    
    On Error GoTo EH
    Set m_oPrinters = oPrinters
    m_sConfFile = sConfFile
    Caption = sProductName
    '--- setup systray
    Set m_oSysTray = New cSysTray
    m_oSysTray.Init Me, sProductName
    '--- success
    Init = True
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Public Sub ShowConfig(Optional OwnerForm As Object)
    Const FUNC_NAME     As String = "ShowConfig"
    
    On Error GoTo EH
    frmSettings.Init OwnerForm
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Public Sub ShutDown()
    Const FUNC_NAME     As String = "Shutdown"
    Dim oFrm            As Object
    
    On Error GoTo EH
    TerminateEndpoints
    FlushDebugLog
    For Each oFrm In Forms
        Unload oFrm
    Next
    Set frmSettings = Nothing
    Set frmIcon = Nothing
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Public Sub Restart(Optional AddParam As Variant)
    Const FUNC_NAME     As String = "Restart"

    On Error GoTo EH
    ShutDown
    If IsMissing(AddParam) Or InIde Then
        Main
    Else
        ShellExec GetProcessName(), Trim$(Command$ & IIf(LenB(AddParam) <> 0, " " & ArgvQuote(AddParam & vbNullString), vbNullString))
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

'=========================================================================
' Events
'=========================================================================

Private Sub mnuPopup_Click(Index As Integer)
    Const FUNC_NAME     As String = "mnuPopup_Click"
    
    On Error GoTo EH
    Select Case Index
    Case ucsMnuPopupConfig
        ShowConfig
    Case ucsMnuPopupRestart
        Restart vbNullString
    Case ucsMnuPopupExit
        ShutDown
    End Select
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub m_oSysTray_Click()
    Const FUNC_NAME     As String = "m_oSysTray_Click"
    
    On Error GoTo EH
    mnuPopup_Click ucsMnuPopupConfig
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub m_oSysTray_ContextMenu()
    Const FUNC_NAME     As String = "m_oSysTray_ContextMenu"
    
    On Error GoTo EH
    PopupMenu mnuMainPopup, DefaultMenu:=mnuPopup(ucsMnuPopupConfig)
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Const FUNC_NAME     As String = "Form_Unload"
    
    On Error GoTo EH
    If Not m_oSysTray Is Nothing Then
        m_oSysTray.Terminate
        Set m_oSysTray = Nothing
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

