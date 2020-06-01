VERSION 5.00
Begin VB.Form frmSettings 
   Caption         =   "Настройки"
   ClientHeight    =   8124
   ClientLeft      =   192
   ClientTop       =   840
   ClientWidth     =   14772
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
   ScaleWidth      =   14772
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picTab 
      BorderStyle     =   0  'None
      HasDC           =   0   'False
      Height          =   6984
      Index           =   0
      Left            =   252
      ScaleHeight     =   6984
      ScaleWidth      =   9924
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   504
      Visible         =   0   'False
      Width           =   9924
      Begin VB.CommandButton cmdTest 
         Caption         =   "Тест"
         Height          =   432
         Left            =   8484
         TabIndex        =   8
         Tag             =   "FONT"
         Top             =   2436
         Width           =   1104
      End
      Begin VB.Frame fraQuickSetup 
         Caption         =   "Бързи настройки"
         Height          =   4380
         Left            =   0
         TabIndex        =   15
         Tag             =   "FONT"
         Top             =   0
         Width           =   4464
         Begin VB.TextBox txtSerialNo 
            Height          =   288
            Left            =   2772
            TabIndex        =   5
            Tag             =   "FONT"
            Top             =   3024
            Width           =   1440
         End
         Begin VB.CheckBox chkAutoDetect 
            Caption         =   "Автоматично откриване на устройства"
            Height          =   348
            Left            =   252
            TabIndex        =   0
            Tag             =   "FONT"
            Top             =   420
            Width           =   4128
         End
         Begin VB.CommandButton cmdApply 
            Caption         =   "Прилагане"
            Height          =   432
            Left            =   2772
            TabIndex        =   6
            Tag             =   "FONT"
            Top             =   3612
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
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Сериен номер на ФУ:"
            Height          =   192
            Left            =   252
            TabIndex        =   20
            Tag             =   "FONT"
            Top             =   3024
            UseMnemonic     =   0   'False
            Width           =   2640
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Парола по подразбиране:"
            Height          =   192
            Left            =   252
            TabIndex        =   19
            Tag             =   "FONT"
            Top             =   2520
            UseMnemonic     =   0   'False
            Width           =   2640
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Протокол за комуникация:"
            Height          =   192
            Left            =   252
            TabIndex        =   18
            Tag             =   "FONT"
            Top             =   1008
            UseMnemonic     =   0   'False
            Width           =   2640
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Скорост на комуникация:"
            Height          =   192
            Left            =   252
            TabIndex        =   17
            Tag             =   "FONT"
            Top             =   2016
            UseMnemonic     =   0   'False
            Width           =   2640
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Комуникационен порт:"
            Height          =   192
            Left            =   252
            TabIndex        =   16
            Tag             =   "FONT"
            Top             =   1512
            UseMnemonic     =   0   'False
            Width           =   2640
            WordWrap        =   -1  'True
         End
      End
      Begin VB.TextBox txtInfo 
         BackColor       =   &H8000000F&
         Height          =   2868
         Left            =   4620
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   9
         Tag             =   "MONO"
         Top             =   3192
         Width           =   5052
      End
      Begin VB.ListBox lstPrinters 
         Height          =   2928
         IntegralHeight  =   0   'False
         Left            =   4620
         TabIndex        =   7
         Tag             =   "MONO"
         Top             =   84
         Width           =   5052
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
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   504
      Visible         =   0   'False
      Width           =   9924
      Begin VB.TextBox txtLog 
         BorderStyle     =   0  'None
         Height          =   5808
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   11
         Tag             =   "MONO"
         Top             =   0
         Width           =   9252
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
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   504
      Visible         =   0   'False
      Width           =   9924
      Begin VB.TextBox txtConfig 
         BorderStyle     =   0  'None
         Height          =   5640
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   10
         Tag             =   "MONO"
         Top             =   0
         Width           =   6396
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
   Begin VB.Menu mnuMain 
      Caption         =   "Помощ"
      Index           =   2
      Begin VB.Menu mnuHelp 
         Caption         =   "Проверка нова версия"
         Index           =   0
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "Относно"
         Index           =   2
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

'--- Windows Messages
Private Const WM_SETREDRAW              As Long = &HB
Private Const WM_GETTEXTLENGTH          As Long = &HE
Private Const WM_GETMINMAXINFO          As Long = &H24
Private Const EM_SETSEL                 As Long = &HB1
Private Const EM_REPLACESEL             As Long = &HC2
Private Const EM_CANUNDO                As Long = &HC6
Private Const EM_UNDO                   As Long = &HC7
Private Const WM_VSCROLL                As Long = &H115
Private Const WM_INITMENU               As Long = &H116
Private Const WM_CUT                    As Long = &H300
Private Const WM_COPY                   As Long = &H301
Private Const WM_PASTE                  As Long = &H302
Private Const WM_CLEAR                  As Long = &H303
'--- clipboard format
Private Const CF_UNICODETEXT            As Long = 13
'--- for WM_VSCROLL
Private Const SB_BOTTOM                 As Long = 7

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Type POINTAPI
    X                   As Long
    Y                   As Long
End Type

Private Type MINMAXINFO
    ptReserved          As POINTAPI
    ptMaxSize           As POINTAPI
    ptMaxPosition       As POINTAPI
    ptMinTrackSize      As POINTAPI
    ptMaxTrackSize      As POINTAPI
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_CAPTION               As String = "Настройки на %1"
Private Const STR_CAPTION_PRINTERS      As String = "Устройства"
Private Const STR_CAPTION_CONFIG        As String = "Конфигурация"
Private Const STR_CAPTION_LOG           As String = "Журнал"
Private Const STR_HEADER_PRINTERS       As String = "Сериен No.|Порт|Хост|Модел|Версия"
Private Const STR_PROTOCOLS             As String = "TREMOL|DATECS|DATECS/X|DAISY|INCOTEX|ELTRADE|ESC/POS|PROXY"
Private Const STR_SPEEDS                As String = "9600|19200|38400|57600|115200"
Private Const STR_CAPTION_APPLY         As String = "Прилагане"
Private Const STR_CAPTION_DISCOVERY     As String = "Търсене"
Private Const STR_MONIKER               As String = "UcsFPHub.LocalEndpoint"
'--- messages
Private Const MSG_SAVE_CHANGES          As String = "Желаете ли да запазите модификациите на %1?"
Private Const MSG_SAVE_SUCCESS          As String = "Успешен запис на %1!" & vbCrLf & vbCrLf & "Желаете ли да рестартирате %2 за да активирате промените?"
Private Const MSG_PRINTER_NOT_FOUND     As String = "Не е открито фискалното устройство с тези настройки." & vbCrLf & vbCrLf & "Желаете ли повторно прилагане?"
Private Const MSG_SUCCESS_FOUND         As String = "Успешно конфигуриране на фискално устройство %1!"
Private Const MSG_UPDATE_FOUND          As String = "Желаете ли да обновите %1 до последна версия след рестартиране?"
Private Const MSG_NO_UPDATE             As String = "Не е намерена по-нова версия на %1"
'--- numeric
Private Const GRID_SIZE                 As Long = 60
Private Const DEF_MIN_WIDTH             As Single = 10000
Private Const DEF_MIN_HEIGHT            As Single = 6000

Private m_sConfFile                 As String
Private m_sPrinterID                As String
Private m_bInSet                    As Boolean
Private m_bQuickSettingsChanged     As Boolean
Private m_bConfigChanged            As Boolean
Private m_bLogChanged               As Boolean
Private m_pSubclass                 As IUnknown
Private m_lConfigPosition           As Long
Private m_pTimerLog                 As IUnknown
Private m_lLogMemoryCount           As Long

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
    ucsMnuHelpAutoUpdate = 0
    ucsMnuHelpSep1
    ucsMnuHelpAbout
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

Private Property Get pvQuickSettingsChanged() As Boolean
    pvQuickSettingsChanged = m_bQuickSettingsChanged
End Property

Private Property Let pvQuickSettingsChanged(ByVal bValue As Boolean)
    m_bQuickSettingsChanged = bValue
    tabMain.TabCaption(ucsTabPrinters) = STR_CAPTION_PRINTERS & IIf(bValue, "*", vbNullString)
End Property

Private Property Get pvConfigChanged() As Boolean
    pvConfigChanged = m_bConfigChanged
End Property

Private Property Let pvConfigChanged(ByVal bValue As Boolean)
    m_bConfigChanged = bValue
    tabMain.TabCaption(ucsTabConfig) = STR_CAPTION_CONFIG & IIf(bValue, "*", vbNullString)
End Property

Private Property Get pvLogChanged() As Boolean
    pvLogChanged = m_bLogChanged
End Property

Private Property Let pvLogChanged(ByVal bValue As Boolean)
    m_bLogChanged = bValue
    tabMain.TabCaption(ucsTabLog) = STR_CAPTION_LOG & IIf(bValue, "*", vbNullString)
End Property

Private Property Get pvAddressOfSubclassProc() As frmSettings
    Set pvAddressOfSubclassProc = InitAddressOfMethod(Me, 5)
End Property

Private Property Get pvConfigText() As String
    pvConfigText = txtConfig.Text
End Property

Private Property Let pvConfigText(sValue As String)
    m_bInSet = True
    txtConfig.Text = sValue
    m_bInSet = False
End Property

Private Property Get pvAddressOfTimerProc() As frmSettings
    Set pvAddressOfTimerProc = InitAddressOfMethod(Me, 0)
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function Init(Optional OwnerForm As Object) As Boolean
    Const FUNC_NAME     As String = "Init"
    Dim sConfFile       As String
    Dim oCtl            As Object
    
    On Error GoTo EH
    If LenB(m_sConfFile) = 0 Then
        WindowState = C_Lng(GetSetting(App.Title, MODULE_NAME, "WindowState", 0))
        Set m_pSubclass = InitSubclassingThunk(hWnd, Me, pvAddressOfSubclassProc.SubclassProc(0, 0, 0, 0, 0))
        '--- fix font size
        Set Me.Font = SystemIconFont
        For Each oCtl In Controls
            If InStr(oCtl.Tag, "FONT") Then
                Set oCtl.Font = Me.Font
            ElseIf InStr(oCtl.Tag, "MONO") Then
                oCtl.Font.Name = "Consolas"
                If oCtl.Font.Name <> "Consolas" Then
                    oCtl.Font.Name = "Courier New"
                End If
                oCtl.Font.Size = Me.Font.Size
            End If
        Next
        '--- setup caption
        sConfFile = MainForm.ConfFile
        m_sConfFile = Zn(sConfFile, PathCombine(GetSpecialFolder(ucsOdtLocalAppData) & "\Unicontsoft\UcsFPHub", App.EXEName & ".conf"))
        Caption = IIf(LenB(sConfFile), sConfFile & " - ", vbNullString) & Printf(STR_CAPTION, App.ProductName & " v" & STR_VERSION)
        '--- load combos
        pvLoadItemData cobProtocol, Split(STR_PROTOCOLS, "|")
        pvLoadItemData cobPort, EnumSerialPorts
        pvLoadItemData cobSpeed, Split(STR_SPEEDS, "|")
        txtSerialNo.Height = cobSpeed.Height
        txtDefPass.Height = cobSpeed.Height
        '--- delay-load UI
        m_sPrinterID = vbNullString
        lstPrinters.Clear
        pvQuickSettingsChanged = False
        pvConfigText = vbNullString
        pvConfigChanged = False
        pvLoadLog
        Set m_pTimerLog = InitFireOnceTimerThunk(Me, pvAddressOfTimerProc.TimerProc, Delay:=1000)
        tabMain_Click
    End If
    mnuHelp(ucsMnuHelpAutoUpdate).Enabled = (LenB(MainForm.ExeAutoUpdate) <> 0)
    If Not OwnerForm Is Nothing Then
        If Not pvShowModal(OwnerForm) Then
            If Not pvShowModal() Then
                Show
            End If
        End If
    Else
        Show
    End If
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
    Const FUNC_NAME     As String = "SubclassProc"
    Dim uInfo           As MINMAXINFO
    
    On Error GoTo EH
    #If hWnd And wParam And lParam And Handled Then '--- touch args
    #End If
    Select Case wMsg
    Case WM_INITMENU
        pvMenuNegotiate ActiveControl
    Case WM_GETMINMAXINFO
        Call CopyMemory(uInfo, ByVal lParam, LenB(uInfo))
        uInfo.ptMinTrackSize.X = DEF_MIN_WIDTH / ScreenTwipsPerPixelX
        uInfo.ptMinTrackSize.Y = DEF_MIN_HEIGHT / ScreenTwipsPerPixelY
        Call CopyMemory(ByVal lParam, uInfo, LenB(uInfo))
        Handled = True
    End Select
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Public Function TimerProc() As Long
Attribute TimerProc.VB_MemberFlags = "40"
    Const FUNC_NAME     As String = "TimerProc"
    
    On Error GoTo EH
    If tabMain.CurrentTab = ucsTabLog Then
        pvLoadLog
    ElseIf pvGetLogTextLength <> 0 Then
        pvLogChanged = (Logger.MemoryCount <> m_lLogMemoryCount)
    End If
    Set m_pTimerLog = InitFireOnceTimerThunk(Me, pvAddressOfTimerProc.TimerProc, Delay:=1000)
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

'= private ===============================================================

Private Function pvLoadPrinters() As Boolean
    Const FUNC_NAME     As String = "pvLoadPrinters"
    Dim oConfig         As Object
    Dim vKey            As Variant
    Dim oDevice         As Object
    Dim oForm           As frmMain
    Dim vSplit          As Variant
    
    On Error GoTo EH
    '--- quick setup  frame
    If LenB(pvConfigText) = 0 Then
        pvLoadConfig m_sConfFile
    End If
    Set oConfig = JsonParseObject(pvConfigText)
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
    m_bInSet = True
    cobProtocol.Text = JsonItem(oDevice, "Protocol")
    cobPort.Text = JsonItem(oDevice, "Port")
    cobSpeed.Text = JsonItem(oDevice, "Speed")
    txtSerialNo.Text = JsonItem(oDevice, "DeviceSerialNo")
    txtDefPass.Text = JsonItem(oDevice, "DefaultPassword")
    m_bInSet = False
    '--- printers list
    Set oForm = MainForm
    vSplit = Split(STR_HEADER_PRINTERS, "|")
    lstPrinters.Clear
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
    pvQuickSettingsChanged = False
    '--- success
    pvLoadPrinters = True
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Private Function pvLoadLog() As Boolean
    Const FUNC_NAME     As String = "pvLoadLog"
    Dim oMemoryLog      As Object
    Dim cNewChunk       As Collection
    Dim lIdx            As Long
    Dim vItem           As Variant
    
    On Error GoTo EH
    If m_lLogMemoryCount <> Logger.MemoryCount Then
        Set oMemoryLog = Logger.MemoryLog
        Set cNewChunk = New Collection
        For lIdx = m_lLogMemoryCount + 1 To Logger.MemoryCount
            If SearchCollection(oMemoryLog, "#" & lIdx, RetVal:=vItem) Then
                cNewChunk.Add vItem
            End If
        Next
        m_lLogMemoryCount = Logger.MemoryCount
        pvAppendLogText ConcatCollection(cNewChunk) & vbCrLf
    End If
    pvLogChanged = False
    '--- success
    pvLoadLog = True
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Private Function pvLoadConfig(sConfFile As String) As Boolean
    Const FUNC_NAME     As String = "pvLoadConfig"
    Dim oConfig         As Object
    
    On Error GoTo EH
    If FileExists(sConfFile) Then
        If FileLen(sConfFile) = 0 Then
            GoTo LoadDefaultConfig
        End If
        pvConfigText = ReadTextFile(sConfFile)
    Else
LoadDefaultConfig:
        mnuFile(ucsMnuFileRestart).Enabled = False
        JsonItem(oConfig, "Printers/Autodetect") = True
        JsonItem(oConfig, "Endpoints/0/Binding") = "RestHttp"
        JsonItem(oConfig, "Endpoints/0/Address") = "127.0.0.1:" & DEF_LISTEN_PORT
        pvConfigText = JsonDump(oConfig)
    End If
    txtConfig.SelStart = m_lConfigPosition
    pvConfigChanged = False
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
        If Not JsonParse(pvConfigText, Empty, Error:=sError, LastPos:=lPos) Then
            MsgBox sError, vbExclamation
            txtConfig.SelStart = lPos - 1
            txtConfig.SelLength = 1
            m_lConfigPosition = txtConfig.SelStart
            GoTo QH
        End If
        WriteTextFile sConfFile, pvConfigText, ucsFltUtf8
    End If
    pvConfigChanged = False
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
    If pvConfigChanged Then
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
    mnuFile(ucsMnuFileSave).Enabled = pvConfigChanged
    mnuEdit(ucsMnuEditUndo).Enabled = pvCanExecute(ucsMnuEditUndo, oCtl)
    mnuEdit(ucsMnuEditCut).Enabled = pvCanExecute(ucsMnuEditCut, oCtl)
    mnuEdit(ucsMnuEditCopy).Enabled = pvCanExecute(ucsMnuEditCopy, oCtl)
    mnuEdit(ucsMnuEditPaste).Enabled = pvCanExecute(ucsMnuEditPaste, oCtl)
    mnuEdit(ucsMnuEditDelete).Enabled = pvCanExecute(ucsMnuEditDelete, oCtl)
    mnuEdit(ucsMnuEditSelectAll).Enabled = pvCanExecute(ucsMnuEditSelectAll, oCtl)
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
        '--- success
        pvLoadItemData = True
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Friend Sub frRestart()
    Const FUNC_NAME     As String = "frRestart"
    Dim oFrm            As Object
    
    On Error GoTo EH
    If LenB(pvConfigText) <> 0 Then
        m_lConfigPosition = txtConfig.SelStart
    End If
    TerminateEndpoints
    FlushDebugLog
    For Each oFrm In Forms
        If Not oFrm Is Me Then
            Unload oFrm
        End If
    Next
    Set frmMain = Nothing
    Process SplitArgs(Command$ & " --systray"), True
    '--- delay-load UI
    m_sPrinterID = vbNullString
    lstPrinters.Clear
    pvQuickSettingsChanged = False
    pvConfigText = vbNullString
    pvConfigChanged = False
    pvLoadLog
    Set m_pTimerLog = InitFireOnceTimerThunk(Me, pvAddressOfTimerProc.TimerProc, Delay:=1000)
    tabMain_Click
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Function pvShowModal(Optional OwnerForm As Variant) As Boolean
    On Error GoTo QH
    Show vbModal, OwnerForm
    '--- success
    pvShowModal = True
QH:
End Function

Private Function pvGetLogTextLength() As Long
    pvGetLogTextLength = SendMessage(txtLog.hWnd, WM_GETTEXTLENGTH, 0, ByVal 0&)
End Function

Private Sub pvAppendLogText(sValue As String)
    Call SendMessage(txtLog.hWnd, WM_SETREDRAW, 0, ByVal 0)
    Call SendMessage(txtLog.hWnd, EM_SETSEL, 0, ByVal -1)
    Call SendMessage(txtLog.hWnd, EM_SETSEL, -1, ByVal -1)
    Call SendMessage(txtLog.hWnd, EM_REPLACESEL, 1, ByVal sValue)
    Call SendMessage(txtLog.hWnd, EM_SETSEL, 0, ByVal -1)
    Call SendMessage(txtLog.hWnd, EM_SETSEL, -1, ByVal -1)
    Call SendMessage(txtLog.hWnd, WM_SETREDRAW, 1, ByVal 0)
    Call SendMessage(txtLog.hWnd, WM_VSCROLL, SB_BOTTOM, ByVal 0)
End Sub

'=========================================================================
' Events
'=========================================================================

Private Sub cmdApply_Click()
    Const FUNC_NAME     As String = "cmdApply_Click"
    Dim oConfig         As Object
    Dim oDevice         As Object
    Dim sDeviceSerialNo As String
    
    On Error GoTo EH
RetryRestart:
    cmdApply.Enabled = False
    cmdApply.Caption = STR_CAPTION_DISCOVERY
    If LenB(pvConfigText) = 0 Then
        pvLoadConfig m_sConfFile
    End If
    Set oConfig = JsonParseObject(pvConfigText)
    JsonItem(oConfig, "Printers/Autodetect") = (chkAutoDetect.Value = vbChecked)
    Set oDevice = ParseDeviceString(C_Str(JsonItem(oConfig, "Printers/" & m_sPrinterID & "/DeviceString")))
    JsonItem(oDevice, "Protocol") = Zn(Trim$(cobProtocol.Text), Empty)
    JsonItem(oDevice, "Port") = Zn(Trim$(cobPort.Text), Empty)
    JsonItem(oDevice, "Speed") = Zn(Trim$(cobSpeed.Text), Empty)
    JsonItem(oDevice, "DeviceSerialNo") = Zn(Trim$(txtSerialNo.Text), Empty)
    JsonItem(oDevice, "DefaultPassword") = Zn(Trim$(txtDefPass.Text), Empty)
    JsonItem(oConfig, "Printers/" & m_sPrinterID & "/DeviceString") = Zn(ToDeviceString(oDevice), Empty)
    If UBound(JsonKeys(oConfig, "Printers/" & m_sPrinterID)) < 0 Then
        JsonItem(oConfig, "Printers/" & m_sPrinterID) = Empty
    End If
    pvConfigText = JsonDump(oConfig)
    If pvSaveConfig(m_sConfFile) Then
        frRestart
        pvLoadPrinters
        sDeviceSerialNo = C_Str(JsonItem(MainForm.Printers, "Aliases/" & m_sPrinterID & "/DeviceSerialNo"))
        If LenB(sDeviceSerialNo) <> 0 Then
            sDeviceSerialNo = C_Str(JsonItem(MainForm.Printers, sDeviceSerialNo & "/DeviceSerialNo"))
        End If
    End If
    cmdApply.Caption = STR_CAPTION_APPLY
    cmdApply.Enabled = True
    cmdApply.SetFocus
    If LenB(JsonItem(oDevice, "Protocol")) <> 0 Then
        If LenB(sDeviceSerialNo) <> 0 Then
            MsgBox Printf(MSG_SUCCESS_FOUND, sDeviceSerialNo), vbExclamation
        ElseIf MsgBox(MSG_PRINTER_NOT_FOUND, vbQuestion Or vbYesNo) = vbYes Then
            GoTo RetryRestart
        End If
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub cmdTest_Click()
    Const FUNC_NAME     As String = "cmdTest_Click"
    Const URL_INFO      As String = "/printers/%1?format=json"
    Dim sPrinterID      As String
    Dim sResponse       As String
    
    On Error GoTo EH
    Screen.MousePointer = vbHourglass
    sPrinterID = Trim$(At(Split(lstPrinters.Text, vbTab), 0))
    If Not GetObject(STR_MONIKER).ServiceRequest(Printf(URL_INFO, sPrinterID), vbNullString, sResponse) Then
        GoTo QH
    End If
    txtInfo.Text = sResponse
QH:
    Screen.MousePointer = vbDefault
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume QH
End Sub

Private Sub mnuFile_Click(Index As Integer)
    Const FUNC_NAME     As String = "mnuFile_Click"
    
    On Error GoTo EH
    Screen.MousePointer = vbHourglass
    Select Case Index
    Case ucsMnuFileSave
        If pvConfigChanged Then
            If Not pvSaveConfig(m_sConfFile) Then
                GoTo QH
            End If
            If MsgBox(Printf(MSG_SAVE_SUCCESS, m_sConfFile, App.ProductName & " v" & STR_VERSION), vbQuestion Or vbYesNo) = vbYes Then
                frRestart
            End If
        End If
    Case ucsMnuFileRestart
        If Not pvQuerySaveConfig(m_sConfFile) Then
            GoTo QH
        End If
        frRestart
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
        Select Case tabMain.CurrentTab
        Case ucsTabPrinters
            pvLoadPrinters
        Case ucsTabConfig
            If Not pvQuerySaveConfig(m_sConfFile) Then
                GoTo QH
            End If
            pvLoadConfig m_sConfFile
        Case ucsTabLog
            TimerProc
        End Select
    End Select
QH:
    Screen.MousePointer = vbDefault
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume QH
End Sub

Private Sub mnuHelp_Click(Index As Integer)
    Const FUNC_NAME     As String = "mnuHelp_Click"
    Dim bResult         As Boolean
    
    On Error GoTo EH
    Select Case Index
    Case ucsMnuHelpAutoUpdate
        Screen.MousePointer = vbHourglass
        bResult = MainForm.StartAutoUpdate(vbTrue)
        Screen.MousePointer = vbDefault
        If bResult Then
            If MsgBox(Printf(MSG_UPDATE_FOUND, App.ProductName & " v" & STR_VERSION), vbQuestion Or vbYesNo) = vbYes Then
                Screen.MousePointer = vbHourglass
                MainForm.StartAutoUpdate vbFalse
            End If
        Else
            MsgBox Printf(MSG_NO_UPDATE, App.ProductName & " v" & STR_VERSION), vbInformation
        End If
    Case ucsMnuHelpAbout
        MsgBox App.ProductName & " v" & STR_VERSION & vbCrLf & App.LegalCopyright, vbInformation
    End Select
QH:
    Screen.MousePointer = vbDefault
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume QH
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Const FUNC_NAME     As String = "Form_KeyDown"
    
    On Error GoTo EH
    Select Case KeyCode + Shift * &H1000&
    Case vbKeyS + vbCtrlMask * &H1000&
        mnuFile_Click ucsMnuFileSave
    Case vbKeyZ + vbCtrlMask * &H1000&
        If Not ActiveControl Is txtConfig Then
            GoTo QH
        End If
        mnuEdit_Click ucsMnuEditUndo
    Case vbKeyTab + vbCtrlMask * &H1000&, vbKeyTab + (vbCtrlMask Or vbShiftMask) * &H1000&
        tabMain.CurrentTab = (tabMain.CurrentTab + IIf((Shift And vbShiftMask) <> 0, tabMain.TabCount - 1, 1)) Mod tabMain.TabCount
        tabMain_Click
    Case Else
        GoTo QH
    End Select
    KeyCode = 0: Shift = 0
QH:
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
        SaveSetting App.Title, MODULE_NAME, "WindowState", WindowState
        MoveCtl tabMain, GRID_SIZE, GRID_SIZE / 2, ScaleWidth - 2 * GRID_SIZE
        With picTab(tabMain.CurrentTab)
            dblTop = tabMain.Top + tabMain.Height + GRID_SIZE / 2
            MoveCtl picTab(tabMain.CurrentTab), 0, dblTop, ScaleWidth, ScaleHeight - dblTop
            Select Case tabMain.CurrentTab
            Case ucsTabPrinters
                dblLeft = GRID_SIZE
                dblTop = GRID_SIZE
                MoveCtl fraQuickSetup, dblLeft, dblTop, fraQuickSetup.Width, .ScaleHeight - dblTop - GRID_SIZE
                dblLeft = fraQuickSetup.Left + fraQuickSetup.Width + GRID_SIZE
                dblHeight = (.ScaleHeight - GRID_SIZE) / 2 - cmdTest.Height
                MoveCtl lstPrinters, dblLeft, 0, .ScaleWidth - dblLeft - GRID_SIZE, dblHeight - GRID_SIZE
                MoveCtl cmdTest, lstPrinters.Left + lstPrinters.Width - cmdTest.Width - GRID_SIZE, dblHeight + 2 * GRID_SIZE
                dblTop = dblHeight + GRID_SIZE
                MoveCtl txtInfo, dblLeft, dblTop, .ScaleWidth - dblLeft - GRID_SIZE, .ScaleHeight - dblTop - GRID_SIZE
            Case ucsTabConfig
                dblLeft = GRID_SIZE
                MoveCtl txtConfig, dblLeft, 0, .ScaleWidth - dblLeft, .ScaleHeight
            Case ucsTabLog
                dblLeft = GRID_SIZE
                MoveCtl txtLog, dblLeft, 0, .ScaleWidth - dblLeft, .ScaleHeight
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
    Set m_pTimerLog = Nothing
    m_sConfFile = vbNullString
    m_lLogMemoryCount = 0
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
    '--- delay-load tabs
    Select Case tabMain.CurrentTab
    Case ucsTabPrinters
        If lstPrinters.ListCount = 0 Then
            pvLoadPrinters
        End If
    Case ucsTabConfig
        If LenB(pvConfigText) = 0 Then
            pvLoadConfig m_sConfFile
        End If
    Case ucsTabLog
        TimerProc
    End Select
QH:
    Screen.MousePointer = vbDefault
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume QH
End Sub

Private Sub lstPrinters_Click()
    Const FUNC_NAME     As String = "lstPrinters_Click"
    Dim sPrinterID      As String
    
    On Error GoTo EH
    sPrinterID = Trim$(At(Split(lstPrinters.Text, vbTab), 0))
    txtInfo.Text = JsonDump(JsonItem(MainForm.Printers, sPrinterID))
    cmdTest.Enabled = (lstPrinters.ListIndex > 0)
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub txtConfig_Change()
    Const FUNC_NAME     As String = "txtConfig_Change"
    
    On Error GoTo EH
    If Not m_bInSet Then
        pvConfigChanged = True
        m_lConfigPosition = txtConfig.SelStart
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub txtConfig_KeyPress(KeyAscii As Integer)
    Const FUNC_NAME     As String = "txtConfig_KeyPress"
    
    On Error GoTo EH
    '--- prevent beep (8 = backspace, 9 = tab, 13 = enter)
    If KeyAscii < 32 And KeyAscii <> 8 And KeyAscii <> 9 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub cobProtocol_Change()
    Const FUNC_NAME     As String = "cobProtocol_Change"
    Dim oConfig         As Object
    
    On Error GoTo EH
    If Not m_bInSet Then
        If LenB(cobProtocol.Text) <> 0 Then
            Set oConfig = JsonParseObject(pvConfigText)
            If Not IsObject(JsonItem(oConfig, "Printers/" & m_sPrinterID)) Then
                chkAutoDetect.Value = vbUnchecked
            End If
        End If
        pvQuickSettingsChanged = True
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub cobProtocol_Click()
    cobProtocol_Change
End Sub

Private Sub cobPort_Change()
    cobProtocol_Change
End Sub

Private Sub cobPort_Click()
    cobProtocol_Change
End Sub

Private Sub cobSpeed_Change()
    cobProtocol_Change
End Sub

Private Sub cobSpeed_Click()
    cobProtocol_Change
End Sub

Private Sub txtSerialNo_Change()
    cobProtocol_Change
End Sub

Private Sub txtDefPass_Change()
    cobProtocol_Change
End Sub
