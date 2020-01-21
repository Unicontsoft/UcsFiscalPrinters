VERSION 5.00
Begin VB.Form frmSettings 
   Caption         =   "Настройки"
   ClientHeight    =   8124
   ClientLeft      =   192
   ClientTop       =   840
   ClientWidth     =   10764
   Icon            =   "frmSettings.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8124
   ScaleWidth      =   10764
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picMain 
      BorderStyle     =   0  'None
      HasDC           =   0   'False
      Height          =   7488
      Index           =   0
      Left            =   252
      ScaleHeight     =   7488
      ScaleWidth      =   9924
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   504
      Visible         =   0   'False
      Width           =   9924
      Begin VB.TextBox txtInfo 
         BackColor       =   &H8000000F&
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
         Height          =   2868
         Left            =   4368
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   84
         Width           =   5556
      End
      Begin VB.ListBox lstPrinters 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   10.2
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2928
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   3708
      End
   End
   Begin VB.PictureBox picMain 
      BorderStyle     =   0  'None
      HasDC           =   0   'False
      Height          =   7740
      Index           =   2
      Left            =   252
      ScaleHeight     =   7740
      ScaleWidth      =   10008
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   504
      Visible         =   0   'False
      Width           =   10008
      Begin VB.TextBox txtLog 
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
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   6
         Top             =   0
         Width           =   4296
      End
   End
   Begin VB.PictureBox picMain 
      BorderStyle     =   0  'None
      HasDC           =   0   'False
      Height          =   7404
      Index           =   1
      Left            =   252
      ScaleHeight     =   7404
      ScaleWidth      =   9924
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   504
      Visible         =   0   'False
      Width           =   9924
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
   Begin UcsFPHub.AlphaBlendTabStrip tabMain 
      Height          =   348
      Left            =   84
      Top             =   84
      Width           =   10512
      _ExtentX        =   18542
      _ExtentY        =   614
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Layout          =   "Устройства|Конфигурация|Журнал"
   End
   Begin VB.Menu mnuMainFile 
      Caption         =   "Файл"
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
End
Attribute VB_Name = "frmSettings"
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
' Constants and member variables
'=========================================================================

Private Const STR_CAPTION               As String = "Настройки на %1 v%2"
Private Const MSG_SAVE_CHANGES          As String = "Желаете ли да запазите модификациите на %1?"
Private Const MSG_SAVE_SUCCESS          As String = "Успешен запис на %1!" & vbCrLf & vbCrLf & "Желаете ли да рестартирате %2 за да активирате промените?"
Private Const GRID_SIZE                 As Long = 60
Private Const STR_CAPTION_CONFIG        As String = "Конфигурация"
Private Const STR_HEADER_PRINTERS       As String = "Сериен No.|Порт|Хост|Модел|Версия"

Private m_sConfFile                 As String
Private m_bInSet                    As Boolean
Private m_bChanged                  As Boolean

Private Enum UcsMenuItems
    ucsMnuFileSave = 0
    ucsMnuFileSep1
    ucsMnuFileRestart
    ucsMnuFileSep2
    ucsMnuFileExit
End Enum

Private Enum UcsTabsEnums
    ucsTabDevices
    ucsTabConfig
    ucsTabLog
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

Private Property Get pvChanged() As Boolean
    pvChanged = m_bChanged
End Property

Private Property Let pvChanged(ByVal bValue As Boolean)
    m_bChanged = bValue
    tabMain.TabCaption(ucsTabConfig) = STR_CAPTION_CONFIG & IIf(bValue, "*", vbNullString)
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function Init() As Boolean
    Const FUNC_NAME     As String = "Init"
    Dim oForm           As frmIcon
    Dim oConfig         As Object
    Dim vKey            As Variant
    Dim vSplit          As Variant
    
    On Error GoTo EH
    If LenB(m_sConfFile) = 0 Then
        Set oForm = MainForm
        m_sConfFile = Zn(oForm.ConfFile, PathCombine(GetSpecialFolder(ucsOdtLocalAppData) & "\Unicontsoft\UcsFPHub", App.EXEName & ".conf"))
        Caption = IIf(LenB(oForm.ConfFile), oForm.ConfFile & " - ", vbNullString) & Printf(STR_CAPTION, STR_SERVICE_NAME, STR_VERSION)
        '--- load printers
        vSplit = Split(STR_HEADER_PRINTERS, "|")
        lstPrinters.AddItem Pad(At(vSplit, 0), 15) & vbTab & Pad(At(vSplit, 1), 15) & vbTab & Pad(At(vSplit, 2), 15) & vbTab & _
            Pad(At(vSplit, 3), 23) & vbTab & Pad(At(vSplit, 4), 38)
        For Each vKey In JsonItem(oForm.Printers, "*/DeviceSerialNo")
            If LenB(vKey) <> 0 Then
                lstPrinters.AddItem Pad(vKey, 15) & vbTab & _
                    Pad(JsonItem(oForm.Printers, vKey & "/DevicePort"), 15) & vbTab & _
                    Pad(JsonItem(oForm.Printers, vKey & "/DeviceHost"), 15) & vbTab & _
                    Pad(JsonItem(oForm.Printers, vKey & "/DeviceModel"), 23) & vbTab & _
                    Pad(JsonItem(oForm.Printers, vKey & "/FirmwareVersion"), 38)
            End If
        Next
        '--- load config
        If FileExists(m_sConfFile) Then
            m_bInSet = True
            txtConfig.Text = ReadTextFile(m_sConfFile)
            m_bInSet = False
        Else
            mnuFile(ucsMnuFileRestart).Enabled = False
            JsonItem(oConfig, "Printers/Autodetect") = True
            JsonItem(oConfig, "Endpoints/0/Binding") = "RestHttp"
            JsonItem(oConfig, "Endpoints/0/Address") = "127.0.0.1:" & DEF_LISTEN_PORT
            m_bInSet = True
            txtConfig.Text = JsonDump(oConfig)
            m_bInSet = False
        End If
        '--- load log
        txtLog.Text = vbNullString
        Set tabMain.Font = SystemIconFont
        tabMain_Click
        pvChanged = False
    End If
    Show
    If WindowState = vbMinimized Then
        WindowState = vbNormal
    End If
    '--- success
    Init = True
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

'=========================================================================
' Events
'=========================================================================

Private Sub mnuFile_Click(Index As Integer)
    Const FUNC_NAME     As String = "mnuFile_Click"
    Dim sError          As String
    Dim lPos            As Long
    
    On Error GoTo EH
    Select Case Index
    Case ucsMnuFileSave
        If LenB(m_sConfFile) = 0 Then
            GoTo QH
        End If
        If Not JsonParse(txtConfig.Text, Empty, Error:=sError, LastPos:=lPos) Then
            MsgBox sError, vbExclamation
            txtConfig.SelStart = lPos - 1
            txtConfig.SelLength = 1
            GoTo QH
        End If
        WriteTextFile m_sConfFile, txtConfig.Text, ucsFltUtf8
        pvChanged = False
        If MsgBox(Printf(MSG_SAVE_SUCCESS, m_sConfFile, App.ProductName), vbQuestion Or vbYesNo) = vbYes Then
            MainForm.Restart
        End If
    Case ucsMnuFileRestart
        MainForm.Restart vbNullString
    Case ucsMnuFileExit
        MainForm.ShutDown
    End Select
QH:
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume QH
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Const FUNC_NAME     As String = "Form_KeyDown"
    
    On Error GoTo EH
    Select Case KeyCode Or Shift * &H1000
    Case vbKeyF5
        If tabMain.CurrentTab = ucsTabLog Then
            txtLog.Text = ConcatCollection(Logger.MemoryLog) & vbCrLf
            txtLog.SelStart = &H7FFF&
        End If
    End Select
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub Form_Resize()
    Const FUNC_NAME     As String = "Form_Resize"
    Dim dblTop          As Double
    Dim dblLeft         As Double
    
    On Error GoTo EH
    If WindowState <> vbMinimized Then
        tabMain.Move GRID_SIZE, GRID_SIZE / 2, ScaleWidth - 2 * GRID_SIZE
        With picMain(tabMain.CurrentTab)
            dblTop = tabMain.Top + tabMain.Height + GRID_SIZE / 2
            .Move 0, dblTop, ScaleWidth, ScaleHeight - dblTop
            Select Case tabMain.CurrentTab
            Case ucsTabDevices
                dblLeft = (.ScaleWidth - GRID_SIZE) / 2
                lstPrinters.Move GRID_SIZE, 0, dblLeft - GRID_SIZE, .ScaleHeight
                dblLeft = dblLeft + GRID_SIZE
                txtInfo.Move dblLeft, 0, .ScaleWidth - dblLeft - GRID_SIZE, .ScaleHeight
            Case ucsTabConfig
                txtConfig.Move 0, 0, .ScaleWidth, .ScaleHeight
            Case ucsTabLog
                txtLog.Move 0, 0, .ScaleWidth, .ScaleHeight
            End Select
        End With
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Const FUNC_NAME     As String = "Form_QueryUnload"
    
    On Error GoTo EH
    If pvChanged Then
        Select Case MsgBox(Printf(MSG_SAVE_CHANGES, m_sConfFile), vbQuestion Or vbYesNoCancel)
        Case vbYes
            mnuFile_Click ucsMnuFileSave
        Case vbCancel
            Cancel = 1
        End Select
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub Form_Unload(Cancel As Integer)
    m_sConfFile = vbNullString
End Sub

Private Sub tabMain_Click()
    Const FUNC_NAME     As String = "tabMain_Click"
    Dim lIdx            As Long
    
    On Error GoTo EH
    Form_Resize
    For lIdx = 0 To tabMain.TabCount - 1
        picMain(lIdx).Visible = (lIdx = tabMain.CurrentTab)
    Next
    If tabMain.CurrentTab = ucsTabLog And LenB(txtLog.Text) = 0 Then
        txtLog.Text = ConcatCollection(Logger.MemoryLog) & vbCrLf
        txtLog.SelStart = &H7FFF&
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub lstPrinters_Click()
    Const FUNC_NAME     As String = "lstPrinters_Click"
    Dim sKey            As String
    On Error GoTo EH
    sKey = Trim$(At(Split(lstPrinters.List(lstPrinters.ListIndex), vbTab), 0))
    txtInfo.Text = JsonDump(JsonItem(MainForm.Printers, sKey), Level:=-1)
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub txtConfig_Change()
    Const FUNC_NAME     As String = "txtConfig_Change"
    
    On Error GoTo EH
    If Not m_bInSet Then
        pvChanged = True
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub
