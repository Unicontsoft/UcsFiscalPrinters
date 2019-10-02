VERSION 5.00
Begin VB.Form frmIcon 
   Caption         =   "Настройки на UcsHPHub"
   ClientHeight    =   8400
   ClientLeft      =   192
   ClientTop       =   540
   ClientWidth     =   12720
   Icon            =   "frmIcon.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   700
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1060
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picMain 
      BorderStyle     =   0  'None
      HasDC           =   0   'False
      Height          =   5052
      Left            =   252
      ScaleHeight     =   5052
      ScaleWidth      =   7236
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   84
      Width           =   7236
      Begin VB.TextBox txtConfig 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   10.2
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3960
         Left            =   168
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   1
         Top             =   84
         Width           =   4296
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Файл"
      Index           =   0
      Begin VB.Menu mnuFile 
         Caption         =   "Запис"
         Index           =   0
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Рестарт"
         Index           =   2
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Изход"
         Index           =   4
         Shortcut        =   ^W
      End
   End
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

Private Const MSG_SAVE_SUCCESS          As String = "Успешен запис на '%1'!" & vbCrLf & vbCrLf & "Желаете ли да рестартирате %2 за да активирате промените?"

Private m_oOpt                      As Object
Private m_sConfFile                 As String
Private WithEvents m_oSysTray       As cSysTray
Attribute m_oSysTray.VB_VarHelpID = -1

Private Enum UcsMenuItems
    ucsMnuFileSave = 0
    ucsMnuFileSep1
    ucsMnuFileRestart
    ucsMnuFileSep2
    ucsMnuFileExit
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
    Debug.Print "Critical error: " & Err.Description & " [" & MODULE_NAME & "." & sFunction & "]"
    DebugLog Err.Description & " [" & MODULE_NAME & "." & sFunction & "]", vbLogEventTypeError
End Sub

'=========================================================================
' Methods
'=========================================================================

Public Function Init(oOpt As Object, sConfFile As String, sProductName As String) As Boolean
    Const FUNC_NAME     As String = "Init"
    Dim oConfig         As Object
    
    On Error GoTo EH
    Set m_oOpt = oOpt
    m_sConfFile = Zn(sConfFile, PathCombine(GetSpecialFolder(ucsOdtLocalAppData) & "\Unicontsoft\UcsFPHub", App.EXEName & ".conf"))
    '--- load config
    If LenB(sConfFile) <> 0 Then
        txtConfig.Text = ReadTextFile(sConfFile)
    Else
        mnuFile(ucsMnuFileRestart).Enabled = False
        JsonItem(oConfig, "Printers/Autodetect") = True
        JsonItem(oConfig, "Endpoints/0/Binding") = "RestHttp"
        JsonItem(oConfig, "Endpoints/0/Address") = "127.0.0.1:8192"
        txtConfig.Text = JsonDump(oConfig)
    End If
    '--- setup systray
    Set m_oSysTray = New cSysTray
    m_oSysTray.Init Me, sProductName
    '--- success
    Init = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Public Sub ShowConfig()
    Const FUNC_NAME     As String = "ShowConfig"
    
    On Error GoTo EH
    Show
    txtConfig.SetFocus
QH:
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume QH
End Sub

Public Sub Shutdown()
    Const FUNC_NAME     As String = "Shutdown"
    
    On Error GoTo EH
    TerminateEndpoints
    FlushDebugLog
    Unload Me
QH:
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume QH
End Sub

Public Sub Restart(Optional AddParam As Variant)
    Const FUNC_NAME     As String = "Restart"
    Dim uShell          As SHELLEXECUTEINFO
    
    On Error GoTo EH
    Shutdown
    If IsMissing(AddParam) Or InIde Then
        Main
    Else
        With uShell
            .cbSize = Len(uShell)
            .fMask = SEE_MASK_NOASYNC Or SEE_MASK_FLAG_NO_UI
            .lpFile = GetProcessName()
            .lpParameters = Trim$(Command$ & IIf(LenB(AddParam) <> 0, " " & ArgvQuote(AddParam & vbNullString), vbNullString))
        End With
        Call ShellExecuteEx(uShell)
    End If
QH:
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume QH
End Sub

'=========================================================================
' Methods
'=========================================================================

Private Sub mnuFile_Click(Index As Integer)
    Const FUNC_NAME     As String = "mnuFile_Click"
    Dim sError          As String
    
    On Error GoTo EH
    Select Case Index
    Case ucsMnuFileSave
        If LenB(m_sConfFile) = 0 Then
            GoTo QH
        End If
        If Not JsonParse(txtConfig.Text, Empty, Error:=sError) Then
            MsgBox sError, vbExclamation
            GoTo QH
        End If
        WriteTextFile m_sConfFile, txtConfig.Text, ucsFltUtf8
        If MsgBox(Printf(MSG_SAVE_SUCCESS, m_sConfFile, App.ProductName), vbQuestion Or vbYesNo) = vbYes Then
            Restart
        End If
    Case ucsMnuFileRestart
        Restart vbNullString
    Case ucsMnuFileExit
        Unload Me
    End Select
QH:
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume QH
End Sub

Private Sub mnuPopup_Click(Index As Integer)
    Const FUNC_NAME     As String = "mnuPopup_Click"
    
    On Error GoTo EH
    Select Case Index
    Case ucsMnuPopupConfig
        ShowConfig
    Case ucsMnuPopupRestart
        Restart vbNullString
    Case ucsMnuPopupExit
        Shutdown
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
    mnuPopup_Click ucsMnuPopupConfig
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

Private Sub Form_Resize()
    Const FUNC_NAME     As String = "Form_Resize"
    
    On Error GoTo EH
    If WindowState <> vbMinimized Then
        picMain.Move 0, 0, ScaleWidth, ScaleHeight
        txtConfig.Move 0, 0, picMain.ScaleWidth, picMain.ScaleHeight
    End If
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
