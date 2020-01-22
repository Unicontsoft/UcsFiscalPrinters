VERSION 5.00
Begin VB.Form frmSettings 
   Caption         =   "Настройки"
   ClientHeight    =   8124
   ClientLeft      =   192
   ClientTop       =   840
   ClientWidth     =   10764
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   7.8
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSettings.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8124
   ScaleWidth      =   10764
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picTab 
      BorderStyle     =   0  'None
      HasDC           =   0   'False
      Height          =   6984
      Index           =   0
      Left            =   252
      ScaleHeight     =   6984
      ScaleWidth      =   9924
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   504
      Visible         =   0   'False
      Width           =   9924
      Begin VB.Frame fraQuickSetup 
         Caption         =   "Бързи настройки"
         Height          =   5052
         Left            =   0
         TabIndex        =   12
         Tag             =   "FONT"
         Top             =   0
         Width           =   4464
         Begin VB.CheckBox chkAutoDetect 
            Caption         =   "Автоматично откриване на устройства"
            Height          =   348
            Left            =   252
            TabIndex        =   0
            Tag             =   "FONT"
            Top             =   420
            Width           =   3708
         End
         Begin VB.CommandButton cmdApply 
            Caption         =   "Прилагане"
            Height          =   432
            Left            =   2772
            TabIndex        =   17
            Top             =   3192
            Width           =   1440
         End
         Begin VB.TextBox txtDefPass 
            Height          =   288
            Left            =   2772
            TabIndex        =   4
            Tag             =   "FONT"
            Top             =   2520
            Width           =   1440
         End
         Begin VB.ComboBox cobProtocol 
            Height          =   288
            Left            =   2772
            TabIndex        =   1
            Tag             =   "FONT"
            Top             =   1008
            Width           =   1440
         End
         Begin VB.ComboBox cobSpeed 
            Height          =   288
            Left            =   2772
            TabIndex        =   3
            Tag             =   "FONT"
            Top             =   2016
            Width           =   1440
         End
         Begin VB.ComboBox cobPort 
            Height          =   288
            Left            =   2772
            TabIndex        =   2
            Tag             =   "FONT"
            Top             =   1512
            Width           =   1440
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Парола по подразбиране:"
            Height          =   192
            Left            =   252
            TabIndex        =   16
            Tag             =   "FONT"
            Top             =   2520
            UseMnemonic     =   0   'False
            Width           =   2640
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Протокол за комуникация:"
            Height          =   192
            Left            =   252
            TabIndex        =   15
            Tag             =   "FONT"
            Top             =   1008
            UseMnemonic     =   0   'False
            Width           =   2640
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Скорост на комуникация:"
            Height          =   192
            Left            =   252
            TabIndex        =   14
            Tag             =   "FONT"
            Top             =   2016
            UseMnemonic     =   0   'False
            Width           =   2640
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Комуникационен порт:"
            Height          =   192
            Left            =   252
            TabIndex        =   13
            Tag             =   "FONT"
            Top             =   1512
            UseMnemonic     =   0   'False
            Width           =   2640
            WordWrap        =   -1  'True
         End
      End
      Begin VB.TextBox txtInfo 
         BackColor       =   &H8000000F&
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
         Left            =   4620
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   3192
         Width           =   5052
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
         IntegralHeight  =   0   'False
         Left            =   4620
         TabIndex        =   5
         Top             =   84
         Width           =   5052
      End
   End
   Begin VB.PictureBox picTab 
      BorderStyle     =   0  'None
      HasDC           =   0   'False
      Height          =   6984
      Index           =   1
      Left            =   252
      ScaleHeight     =   6984
      ScaleWidth      =   9924
      TabIndex        =   6
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
         Height          =   5640
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   7
         Top             =   0
         Width           =   6396
      End
   End
   Begin VB.PictureBox picTab 
      BorderStyle     =   0  'None
      HasDC           =   0   'False
      Height          =   6984
      Index           =   2
      Left            =   252
      ScaleHeight     =   6984
      ScaleWidth      =   9924
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   504
      Visible         =   0   'False
      Width           =   9924
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
         Height          =   5808
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   11
         Top             =   0
         Width           =   9252
      End
   End
   Begin UcsFPHub.AlphaBlendTabStrip tabMain 
      Height          =   348
      Left            =   84
      Tag             =   "FONT"
      Top             =   84
      Width           =   10512
      _ExtentX        =   18542
      _ExtentY        =   614
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   7.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Layout          =   "Устройства|Конфигурация|Журнал"
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
   Begin VB.Menu mnuMain 
      Caption         =   "Редакция"
      Index           =   1
      Begin VB.Menu mnuEdit 
         Caption         =   "Върни"
         Index           =   0
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Изрежи"
         Index           =   2
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Копирай"
         Index           =   3
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Постави"
         Index           =   4
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Изтрий"
         Index           =   5
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Избери всичко"
         Index           =   7
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Обнови"
         Index           =   9
         Shortcut        =   {F5}
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
Private Const MODULE_NAME As String = "frmSettings"

'=========================================================================
' API
'=========================================================================

Private Const WM_GETTEXTLENGTH          As Long = &HE
Private Const EM_SETSEL                 As Long = &HB1
Private Const EM_CANUNDO                As Long = &HC6
Private Const EM_UNDO                   As Long = &HC7
Private Const WM_INITMENU               As Long = &H116
Private Const WM_CUT                    As Long = &H300
Private Const WM_COPY                   As Long = &H301
Private Const WM_PASTE                  As Long = &H302
Private Const WM_CLEAR                  As Long = &H303
Private Const CF_UNICODETEXT            As Long = 13

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_CAPTION               As String = "Настройки на %1 v%2"
Private Const STR_CAPTION_CONFIG        As String = "Конфигурация"
Private Const STR_HEADER_PRINTERS       As String = "Сериен No.|Порт|Хост|Модел|Версия"
Private Const STR_PROTOCOLS             As String = "DATECS/X|DATECS|DAISY|INCOTEX|TREMOL|ESC/POS|PROXY"
Private Const STR_SPEEDS                As String = "9600|19200|38400|57600|115200"
'--- messages
Private Const MSG_SAVE_CHANGES          As String = "Желаете ли да запазите модификациите на %1?"
Private Const MSG_SAVE_SUCCESS          As String = "Успешен запис на %1!" & vbCrLf & vbCrLf & "Желаете ли да рестартирате %2 за да активирате промените?"
'--- numeric
Private Const GRID_SIZE                 As Long = 60

Private m_sConfFile                 As String
Private m_sPrinterID                As String
Private m_bInSet                    As Boolean
Private m_bChanged                  As Boolean
Private m_pSubclass                 As IUnknown

Private Enum UcsMenuItems
    ucsMnuFileSave = 0
    ucsMnuFileSep1
    ucsMnuFileRestart
    ucsMnuFileSep2
    ucsMnuFileExit
    ucsMnuEditUndo = 0
    ucsMnuEditSep1
    ucsMnuEditCut
    ucsMnuEditCopy
    ucsMnuEditPaste
    ucsMnuEditDelete
    ucsMnuEditSep2
    ucsMnuEditSelectAll
    ucsMnuEditSep3
    ucsMnuEditRefresh
End Enum

Private Enum UcsTabsEnums
    ucsTabPrinters
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

Private Property Get pvAddressOfSubclassProc() As frmSettings
    Set pvAddressOfSubclassProc = InitAddressOfMethod(Me, 5)
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function Init() As Boolean
    Const FUNC_NAME     As String = "Init"
    Dim sConfFile       As String
    Dim oCtl            As Object
    
    On Error GoTo EH
    If LenB(m_sConfFile) = 0 Then
        Set m_pSubclass = InitSubclassingThunk(hWnd, Me, pvAddressOfSubclassProc.SubclassProc(0, 0, 0, 0, 0))
        '--- fix font size
        Set Me.Font = SystemIconFont
        For Each oCtl In Controls
            If InStr(oCtl.Tag, "FONT") Then
                Set oCtl.Font = Me.Font
            End If
        Next
        '--- setup caption
        sConfFile = MainForm.ConfFile
        m_sConfFile = Zn(sConfFile, PathCombine(GetSpecialFolder(ucsOdtLocalAppData) & "\Unicontsoft\UcsFPHub", App.EXEName & ".conf"))
        Caption = IIf(LenB(sConfFile), sConfFile & " - ", vbNullString) & Printf(STR_CAPTION, STR_SERVICE_NAME, STR_VERSION)
        '--- load combos
        pvLoadItemData cobProtocol, Split(STR_PROTOCOLS, "|")
        With New cFiscalPrinter
            pvLoadItemData cobPort, .SerialPorts
        End With
        pvLoadItemData cobSpeed, Split(STR_SPEEDS, "|")
        txtDefPass.Height = cobSpeed.Height
        '--- delay-load UI
        m_sPrinterID = vbNullString
        lstPrinters.Clear
        txtConfig.Text = vbNullString
        txtLog.Text = vbNullString
        tabMain_Click
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

Public Function SubclassProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, Handled As Boolean) As Long
Attribute SubclassProc.VB_MemberFlags = "40"
    #If hWnd And wParam And lParam And Handled Then '--- touch args
    #End If
    Select Case wMsg
    Case WM_INITMENU
        pvMenuNegotiate ActiveControl
    End Select
End Function

'= private ===============================================================

Private Function pvLoadPrinters() As Boolean
    Const FUNC_NAME     As String = "pvLoadPrinters"
    Dim oConfig         As Object
    Dim vKey            As Variant
    Dim oDevice         As Object
    Dim oForm           As frmIcon
    Dim vSplit          As Variant
    
    On Error GoTo EH
    '--- quick setup  frame
    If LenB(txtConfig.Text) = 0 Then
        pvLoadConfig m_sConfFile
    End If
    Set oConfig = JsonParseObject(txtConfig.Text)
    chkAutoDetect.Value = IIf(C_Bool(JsonItem(oConfig, "Printers/Autodetect")), vbChecked, vbUnchecked)
    m_sPrinterID = "DefaultPrinter"
    Set oDevice = ParseDeviceString(C_Str(JsonItem(oConfig, "Printers/" & m_sPrinterID & "/DeviceString")))
    If oDevice Is Nothing Then
        For Each vKey In JsonKeys(oConfig, "Printers")
            Set oDevice = ParseDeviceString(C_Str(JsonItem(oConfig, "Printers/" & vKey & "/DeviceString")))
            If Not oDevice Is Nothing Then
                m_sPrinterID = vKey
                Exit For
            End If
        Next
    End If
    cobProtocol.Text = JsonItem(oDevice, "Protocol")
    cobPort.Text = JsonItem(oDevice, "Port")
    cobSpeed.Text = JsonItem(oDevice, "Speed")
    txtDefPass.Text = JsonItem(oDevice, "DefaultPassword")
    '--- printers list
    Set oForm = MainForm
    vSplit = Split(STR_HEADER_PRINTERS, "|")
    lstPrinters.AddItem Pad(At(vSplit, 0), 15) & vbTab & Pad(At(vSplit, 1), 15) & vbTab & Pad(At(vSplit, 2), 15) & vbTab & _
        Pad(At(vSplit, 3), 23) & vbTab & At(vSplit, 4)
    For Each vKey In JsonItem(oForm.Printers, "*/DeviceSerialNo")
        If LenB(vKey) <> 0 Then
            lstPrinters.AddItem Pad(vKey, 15) & vbTab & _
                Pad(JsonItem(oForm.Printers, vKey & "/DevicePort"), 15) & vbTab & _
                Pad(JsonItem(oForm.Printers, vKey & "/DeviceHost"), 15) & vbTab & _
                Pad(JsonItem(oForm.Printers, vKey & "/DeviceModel"), 23) & vbTab & _
                JsonItem(oForm.Printers, vKey & "/FirmwareVersion")
        End If
    Next
    If lstPrinters.ListCount > 1 Then
        lstPrinters.ListIndex = 1
    End If
    '--- success
    pvLoadPrinters = True
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Private Function pvLoadConfig(sConfFile As String) As Boolean
    Const FUNC_NAME     As String = "pvLoadConfig"
    Dim oConfig         As Object
    
    On Error GoTo EH
    If FileExists(sConfFile) Then
        m_bInSet = True
        txtConfig.Text = ReadTextFile(sConfFile)
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
    pvChanged = False
    '--- success
    pvLoadConfig = True
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Private Function pvSaveConfig(sConfFile As String) As Boolean
    Const FUNC_NAME     As String = "pvSaveConfig"
    Dim sError          As String
    Dim lPos            As Long
    
    On Error GoTo EH
    If LenB(sConfFile) <> 0 Then
        If Not JsonParse(txtConfig.Text, Empty, Error:=sError, LastPos:=lPos) Then
            MsgBox sError, vbExclamation
            txtConfig.SelStart = lPos - 1
            txtConfig.SelLength = 1
            GoTo QH
        End If
        WriteTextFile sConfFile, txtConfig.Text, ucsFltUtf8
    End If
    pvChanged = False
    '--- success
    pvSaveConfig = True
QH:
    Exit Function
EH:
    sError = Err.Description
    PrintError FUNC_NAME
    MsgBox sError, vbExclamation
End Function

Private Function pvQuerySaveConfig(sConfFile As String) As Boolean
    Const FUNC_NAME     As String = "pvQuerySaveConfig"
    
    On Error GoTo EH
    If pvChanged Then
        Select Case MsgBox(Printf(MSG_SAVE_CHANGES, sConfFile), vbQuestion Or vbYesNoCancel)
        Case vbYes
            If Not pvSaveConfig(sConfFile) Then
                GoTo QH
            End If
        Case vbCancel
            GoTo QH
        End Select
    End If
    '--- success
    pvQuerySaveConfig = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Private Sub pvMenuNegotiate(oCtl As Object)
    Const FUNC_NAME     As String = "pvMenuNegotiate"
    
    On Error GoTo EH
    mnuEdit(ucsMnuEditUndo).Enabled = pvCanExecute(ucsMnuEditUndo, oCtl)
    mnuEdit(ucsMnuEditCut).Enabled = pvCanExecute(ucsMnuEditCut, oCtl)
    mnuEdit(ucsMnuEditCopy).Enabled = pvCanExecute(ucsMnuEditCopy, oCtl)
    mnuEdit(ucsMnuEditPaste).Enabled = pvCanExecute(ucsMnuEditPaste, oCtl)
    mnuEdit(ucsMnuEditDelete).Enabled = pvCanExecute(ucsMnuEditDelete, oCtl)
    mnuEdit(ucsMnuEditSelectAll).Enabled = pvCanExecute(ucsMnuEditSelectAll, oCtl)
    mnuEdit(ucsMnuEditRefresh).Enabled = oCtl Is txtLog Or oCtl Is txtConfig
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Function pvCanExecute(ByVal eMenu As UcsMenuItems, oCtl As Object) As Boolean
    Const FUNC_NAME     As String = "pvCanExecute"
    Dim lTextLen        As Long
    
    On Error GoTo EH
    If TypeName(ActiveControl) = "TextBox" Or TypeName(ActiveControl) = "ComboBox" Then
        Select Case eMenu
        Case ucsMnuEditUndo
            If oCtl.Enabled And Not oCtl.Locked Then
                pvCanExecute = (SendMessage(oCtl.hWnd, EM_CANUNDO, 0, ByVal 0&) <> 0)
            End If
        Case ucsMnuEditCopy
            If oCtl.Enabled Then
                pvCanExecute = (oCtl.SelLength > 0)
            End If
        Case ucsMnuEditCut, ucsMnuEditDelete
            If oCtl.Enabled And Not oCtl.Locked Then
                pvCanExecute = (oCtl.SelLength > 0)
            End If
        Case ucsMnuEditPaste
            If oCtl.Enabled And Not oCtl.Locked Then
                pvCanExecute = Clipboard.GetFormat(CF_UNICODETEXT)
            End If
        Case ucsMnuEditSelectAll
            If oCtl.Enabled Then
                lTextLen = SendMessage(oCtl.hWnd, WM_GETTEXTLENGTH, 0, ByVal 0&)
                pvCanExecute = (lTextLen > 0 And oCtl.SelLength < lTextLen)
            End If
        End Select
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Private Function pvLoadItemData(oCombo As ComboBox, vItemData As Variant) As Boolean
    Const FUNC_NAME     As String = "pvLoadItemData"
    Dim vElem           As Variant
    
    On Error GoTo EH
    oCombo.Clear
    If IsArray(vItemData) Then
        For Each vElem In vItemData
            oCombo.AddItem vElem
        Next
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Sub pvRestart()
    Const FUNC_NAME     As String = "pvRestart"
    Dim oFrm            As Object
    
    On Error GoTo EH
    TerminateEndpoints
    FlushDebugLog
    For Each oFrm In Forms
        If Not oFrm Is Me Then
            Unload oFrm
        End If
    Next
    Set frmIcon = Nothing
    Main
    '--- delay-load UI
    m_sPrinterID = vbNullString
    lstPrinters.Clear
    txtConfig.Text = vbNullString
    txtLog.Text = vbNullString
    tabMain_Click
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

'=========================================================================
' Events
'=========================================================================

Private Sub cmdApply_Click()
    Const FUNC_NAME     As String = "cmdApply_Click"
    Dim oConfig         As Object
    Dim oDevice         As Object
    Dim sDeviceString   As String
    Dim vKey            As Variant
    Dim sValue          As String
    
    On Error GoTo EH
    If LenB(txtConfig.Text) = 0 Then
        pvLoadConfig m_sConfFile
    End If
    Set oConfig = JsonParseObject(txtConfig.Text)
    JsonItem(oConfig, "Printers/Autodetect") = (chkAutoDetect.Value = vbChecked)
    Set oDevice = ParseDeviceString(C_Str(JsonItem(oConfig, "Printers/" & m_sPrinterID & "/DeviceString")))
    JsonItem(oDevice, "Protocol") = Zn(cobProtocol.Text, Empty)
    JsonItem(oDevice, "Port") = Zn(cobPort.Text, Empty)
    JsonItem(oDevice, "Speed") = Zn(cobSpeed.Text, Empty)
    JsonItem(oDevice, "DefaultPassword") = Zn(txtDefPass.Text, Empty)
    For Each vKey In JsonKeys(oDevice)
        '--- try to escape value
        sValue = C_Str(JsonItem(oDevice, vKey))
        If InStr(sValue, ";") > 0 Then
            If InStr(sValue, """") = 0 Then
                sValue = """" & sValue & """"
            Else
                sValue = "'" & sValue & "'"
            End If
        End If
        sDeviceString = IIf(LenB(sDeviceString) <> 0, sDeviceString & ";", vbNullString) & vKey & "=" & sValue
    Next
    JsonItem(oConfig, "Printers/" & m_sPrinterID & "/DeviceString") = sDeviceString
    txtConfig.Text = JsonDump(oConfig)
    pvChanged = True
    mnuFile_Click ucsMnuFileSave
QH:
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub mnuFile_Click(Index As Integer)
    Const FUNC_NAME     As String = "mnuFile_Click"
    
    On Error GoTo EH
    Screen.MousePointer = vbHourglass
    Select Case Index
    Case ucsMnuFileSave
        If pvChanged Then
            If Not pvSaveConfig(m_sConfFile) Then
                GoTo QH
            End If
            If MsgBox(Printf(MSG_SAVE_SUCCESS, m_sConfFile, App.ProductName), vbQuestion Or vbYesNo) = vbYes Then
                pvRestart
            End If
        End If
    Case ucsMnuFileRestart
        If Not pvQuerySaveConfig(m_sConfFile) Then
            GoTo QH
        End If
        pvRestart
    Case ucsMnuFileExit
        MainForm.ShutDown
    End Select
QH:
    Screen.MousePointer = vbDefault
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume QH
End Sub

Private Sub mnuEdit_Click(Index As Integer)
    Const FUNC_NAME     As String = "mnuEdit_Click"
    Dim oCtl            As Object
    
    On Error GoTo EH
    Screen.MousePointer = vbHourglass
    Set oCtl = ActiveControl
    Select Case Index
    Case ucsMnuEditUndo
        Call SendMessage(oCtl.hWnd, EM_UNDO, 0, ByVal 0&)
    Case ucsMnuEditCut, ucsMnuEditCopy
        Call SendMessage(oCtl.hWnd, IIf(Index = ucsMnuEditCut, WM_CUT, WM_COPY), 0, ByVal 0&)
    Case ucsMnuEditPaste
        Call SendMessage(oCtl.hWnd, WM_PASTE, 0, ByVal 0&)
    Case ucsMnuEditDelete
        Call SendMessage(oCtl.hWnd, WM_CLEAR, 0, ByVal 0&)
    Case ucsMnuEditSelectAll
        Call SendMessage(oCtl.hWnd, EM_SETSEL, 0, ByVal -1)
    Case ucsMnuEditRefresh
        If tabMain.CurrentTab = ucsTabConfig Then
            If Not pvQuerySaveConfig(m_sConfFile) Then
                GoTo QH
            End If
            pvLoadConfig m_sConfFile
        ElseIf tabMain.CurrentTab = ucsTabLog Then
            txtLog.Text = ConcatCollection(Logger.MemoryLog) & vbCrLf
            txtLog.SelStart = &H7FFF&
        End If
    End Select
QH:
    Screen.MousePointer = vbDefault
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume QH
End Sub

Private Sub Form_Resize()
    Const FUNC_NAME     As String = "Form_Resize"
    Dim dblTop          As Double
    Dim dblLeft         As Double
    Dim dblHeight       As Double
    
    On Error GoTo EH
    If WindowState <> vbMinimized Then
        MoveCtl tabMain, GRID_SIZE, GRID_SIZE / 2, ScaleWidth - 2 * GRID_SIZE
        With picTab(tabMain.CurrentTab)
            dblTop = tabMain.Top + tabMain.Height + GRID_SIZE / 2
            MoveCtl picTab(tabMain.CurrentTab), 0, dblTop, ScaleWidth, ScaleHeight - dblTop
            Select Case tabMain.CurrentTab
            Case ucsTabPrinters
                dblLeft = GRID_SIZE
                dblTop = GRID_SIZE
                MoveCtl fraQuickSetup, dblLeft, dblTop
                dblLeft = fraQuickSetup.Left + fraQuickSetup.Width + GRID_SIZE
                dblHeight = (.ScaleHeight - GRID_SIZE) / 2
                MoveCtl lstPrinters, dblLeft, 0, .ScaleWidth - dblLeft - GRID_SIZE, dblHeight
                dblTop = dblHeight + GRID_SIZE
                MoveCtl txtInfo, dblLeft, dblTop, .ScaleWidth - dblLeft - GRID_SIZE, .ScaleHeight - dblTop - GRID_SIZE
            Case ucsTabConfig
                dblLeft = GRID_SIZE
                MoveCtl txtConfig, dblLeft, 0, .ScaleWidth - dblLeft - GRID_SIZE, .ScaleHeight
            Case ucsTabLog
                dblLeft = GRID_SIZE
                MoveCtl txtLog, dblLeft, 0, .ScaleWidth - dblLeft - GRID_SIZE, .ScaleHeight
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
    If Cancel = 0 Then
        If Not pvQuerySaveConfig(m_sConfFile) Then
            Cancel = 1
        End If
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Const FUNC_NAME     As String = "Form_Unload"
    
    On Error GoTo EH
    Set m_pSubclass = Nothing
    m_sConfFile = vbNullString
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub tabMain_Click()
    Const FUNC_NAME     As String = "tabMain_Click"
    Dim lIdx            As Long
    
    On Error GoTo EH
    Screen.MousePointer = vbHourglass
    Form_Resize
    For lIdx = 0 To tabMain.TabCount - 1
        picTab(lIdx).Visible = (lIdx = tabMain.CurrentTab)
    Next
    pvMenuNegotiate ActiveControl
    If tabMain.CurrentTab = ucsTabPrinters And lstPrinters.ListCount = 0 Then
        pvLoadPrinters
    ElseIf tabMain.CurrentTab = ucsTabConfig And LenB(txtConfig.Text) = 0 Then
        pvLoadConfig m_sConfFile
    ElseIf tabMain.CurrentTab = ucsTabLog And LenB(txtLog.Text) = 0 Then
        mnuEdit_Click ucsMnuEditRefresh
    End If
QH:
    Screen.MousePointer = vbDefault
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume QH
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

'Private Sub chkAutoDetect_Click()
'    Const FUNC_NAME     As String = "chkAutoDetect_Click"
'    Dim bLocked         As Boolean
'    Dim vElem           As Variant
'
'    On Error GoTo EH
'    bLocked = (chkAutoDetect.Value = vbChecked)
'    For Each vElem In Array(cobProtocol, cobPort, cobSpeed, txtDefPass)
'        vElem.Locked = bLocked
'        vElem.BackColor = IIf(bLocked, vbButtonFace, vbWindowBackground)
'    Next
'    Exit Sub
'EH:
'    PrintError FUNC_NAME
'    Resume Next
'End Sub
