Attribute VB_Name = "mdStartup"
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
Private Const MODULE_NAME As String = "mdStartup"

'=========================================================================
' API
'=========================================================================

Private Const HKEY_CLASSES_ROOT         As Long = &H80000000
Private Const SAM_WRITE                 As Long = &H20007
Private Const REG_SZ                    As Long = 1

Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Private Declare Function SetEnvironmentVariable Lib "kernel32" Alias "SetEnvironmentVariableA" (ByVal lpName As String, ByVal lpValue As String) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As Long, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function SHDeleteKey Lib "shlwapi" Alias "SHDeleteKeyA" (ByVal hKey As Long, ByVal szSubKey As String) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function InitCommonControls Lib "comctl32" () As Long

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_LATEST_COMMIT         As String = ""
Public Const STR_VERSION                As String = "0.1.47" & STR_LATEST_COMMIT
Public Const STR_SERVICE_NAME           As String = "UcsFPHub"
Public Const DEF_LISTEN_PORT            As Long = 8192
Private Const STR_APPID_GUID            As String = "{6E78E71A-35B2-4D23-A88C-4C2858430329}"
Private Const STR_SVC_INSTALL           As String = "Инсталира NT услуга %1..."
Private Const STR_SVC_UNINSTALL         As String = "Деинсталира NT услуга %1..."
Private Const STR_SUCCESS               As String = "Успех"
Private Const STR_FAILURE               As String = "Грешка: "
Private Const STR_WARN                  As String = "Предупреждение: "
Private Const STR_AUTODETECTING_PRINTERS As String = "Автоматично търсене на принтери"
Private Const STR_ENVIRON_VARS_FOUND    As String = "Конфигурирани %1 променливи на средата"
Private Const STR_ONE_PRINTER_FOUND     As String = "Намерен 1 принтер"
Private Const STR_PRINTERS_FOUND        As String = "Намерени %1 принтера"
Private Const STR_PRESS_CTRLC           As String = "Натиснете Ctrl+C за изход"
Private Const STR_LOADING_CONFIG        As String = "Зарежда конфигурация от %1"
Private Const STR_MONIKER               As String = "UcsFPHub.LocalEndpoint"
Private Const STR_REGISTER_APPID_FAILED As String = "Неуспешна регистрация на AppID. %1"
Private Const MSG_ALREADY_RUNNING       As String = "COM сървър с моникер %1 вече е стартиран" & vbCrLf & vbCrLf & "Желаете ли да отворите предишната инстанция?"
Private Const STR_PREFIX_ERROR          As String = "[Грешка] "
Private Const STR_PREFIX_WARNING        As String = "[Внимание] "
'--- errors
Private Const ERR_CONFIG_NOT_FOUND      As String = "Конфигурационен файл %1 не е намерен"
Private Const ERR_PARSING_CONFIG        As String = "Невалиден %1: %2"
Private Const ERR_ENUM_PORTS            As String = "Енумериране на серийни портове: %1"
Private Const ERR_WARN_ACCESS           As String = "Принтер %1 е недостъпен: %2"
'--- formats
Public Const FORMAT_TIME_ONLY           As String = "hh:nn:ss"
Public Const FORMAT_DATETIME_LOG        As String = "yyyy.MM.dd hh:nn:ss"
Public Const FORMAT_BASE_2              As String = "0.00"
Public Const FORMAT_BASE_3              As String = "0.000"
'--- log level
Public Const vbLogEventTypeDebug        As Long = vbLogEventTypeInformation + 1
Public Const vbLogEventTypeDataDump     As Long = vbLogEventTypeInformation + 2

Private m_oOpt                      As Object
Private m_oPrinters                 As Object
Private m_oConfig                   As Object
Private m_cEndpoints                As Collection
Private m_bIsService                As Boolean
Private m_bIsHidden                 As Boolean
Private m_bStarted                  As Boolean

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

Property Get IsRunningAsService() As Boolean
    IsRunningAsService = m_bIsService
End Property

Property Get IsRunningHidden() As Boolean
    IsRunningHidden = m_bIsHidden
End Property

Property Get MainForm() As frmMain
    Dim oForm       As Object
    
    For Each oForm In Forms
        If TypeOf oForm Is frmMain Then
            Set MainForm = oForm
            Exit Property
        End If
    Next
End Property

Property Get LocalEndpointForm() As frmLocalEndpoint
    Dim oForm       As Object
    
    For Each oForm In m_cEndpoints
        If TypeOf oForm Is frmLocalEndpoint Then
            Set LocalEndpointForm = oForm
            Exit Property
        End If
    Next
End Property

'=========================================================================
' Functions
'=========================================================================

Public Sub Main()
    Const FUNC_NAME     As String = "Main"
    Dim lExitCode       As Long
    
    On Error GoTo EH
    If Not m_bStarted Then
        If Not InIde Then
            '--- prepare for visual styles
            Call LoadLibrary("shell32.dll")
            Call InitCommonControls
        End If
        ApplyTheme
        SetCurrentDateTimer VBA.Now, TimerEx
        Logger.Log 0, MODULE_NAME, FUNC_NAME, App.ProductName & " v" & STR_VERSION
    End If
    lExitCode = Process(SplitArgs(Command$), m_bStarted)
    m_bStarted = True
    If Not InIde And lExitCode <> -1 Then
        Call ExitProcess(lExitCode)
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Public Function Process(vArgs As Variant, ByVal bStarted As Boolean) As Long
    Const FUNC_NAME     As String = "Process"
    Dim sConfFile       As String
    Dim sError          As String
    Dim vKey            As Variant
    Dim lIdx            As Long
    Dim sLogFile        As String
    Dim sLogDataDump    As String
    
    On Error GoTo EH
    Screen.MousePointer = vbHourglass
    Set m_oOpt = GetOpt(vArgs, "config:-config:c")
    '--- normalize options: convert -o and -option to proper long form (--option)
    For Each vKey In Split("nologo config:c install:i uninstall:u systray:s hidden console help:h:?")
        vKey = Split(vKey, ":")
        For lIdx = 0 To UBound(vKey)
            If IsEmpty(m_oOpt.Item("--" & At(vKey, 0))) And Not IsEmpty(m_oOpt.Item("-" & At(vKey, lIdx))) Then
                m_oOpt.Item("--" & At(vKey, 0)) = m_oOpt.Item("-" & At(vKey, lIdx))
            End If
        Next
    Next
    If Not C_Bool(m_oOpt.Item("--nologo")) And Not bStarted Then
        ConsolePrint App.ProductName & " v" & STR_VERSION & vbCrLf & Replace(App.LegalCopyright, "©", "(c)") & vbCrLf & vbCrLf
    End If
    If C_Bool(m_oOpt.Item("--help")) Then
        ConsolePrint "Usage: " & App.EXEName & ".exe [options...]" & vbCrLf & vbCrLf & _
                    "Options:" & vbCrLf & _
                    "  -c, --config FILE   read configuration from FILE" & vbCrLf & _
                    "  -i, --install       install NT service (with config file from -c option)" & vbCrLf & _
                    "  -u, --uninstall     remove NT service" & vbCrLf & _
                    "  -s, --systray       on startup minimize to systray" & vbCrLf & _
                    "  --console           output to console" & vbCrLf
        GoTo QH
    End If
    '--- setup config filename
    sConfFile = C_Str(m_oOpt.Item("--config"))
    If LenB(sConfFile) = 0 Then
        sConfFile = LocateFile(PathCombine(App.Path, App.EXEName & ".conf"))
        If LenB(sConfFile) = 0 Then
            sConfFile = PathCombine(GetSpecialFolder(ucsOdtLocalAppData) & "\Unicontsoft\UcsFPHub", App.EXEName & ".conf")
            If Not FileExists(sConfFile) Then
                sConfFile = vbNullString
            End If
        End If
    End If
    '--- setup service
    If NtServiceInit(STR_SERVICE_NAME) Then
        m_bIsService = True
        '--- cannot handle these as NT service
        m_oOpt.Item("--systray") = Empty
        m_oOpt.Item("--install") = Empty
        m_oOpt.Item("--uninstall") = Empty
        m_oOpt.Item("--console") = True
        m_oOpt.Item("--hidden") = True
    End If
    m_bIsHidden = C_Bool(m_oOpt.Item("--hidden"))
    If C_Bool(m_oOpt.Item("--install")) Then
        ConsolePrint Printf(STR_SVC_INSTALL, STR_SERVICE_NAME) & vbCrLf
        If LenB(sConfFile) <> 0 Then
            sConfFile = " --config " & ArgvQuote(sConfFile)
        End If
        If Not pvRegisterServiceAppID(STR_SERVICE_NAME, App.ProductName & " (" & STR_VERSION & ")", App.EXEName & ".exe", STR_APPID_GUID, Error:=sError) Then
            ConsoleError STR_WARN & sError & vbCrLf
        End If
        If Not NtServiceInstall(STR_SERVICE_NAME, App.ProductName & " (" & STR_VERSION & ")", GetProcessName() & sConfFile, Error:=sError) Then
            ConsoleError STR_FAILURE
            ConsoleColorError FOREGROUND_RED, FOREGROUND_MASK, sError & vbCrLf
        Else
            ConsolePrint STR_SUCCESS & vbCrLf
        End If
        GoTo QH
    ElseIf C_Bool(m_oOpt.Item("--uninstall")) Then
        ConsolePrint Printf(STR_SVC_UNINSTALL, STR_SERVICE_NAME) & vbCrLf
        If Not pvUnregisterServiceAppID(App.EXEName & ".exe", STR_APPID_GUID, Error:=sError) Then
            ConsoleError STR_WARN & sError & vbCrLf
        End If
        If Not NtServiceUninstall(STR_SERVICE_NAME, Error:=sError) Then
            ConsoleError STR_FAILURE
            ConsoleColorError FOREGROUND_RED, FOREGROUND_MASK, sError & vbCrLf
        Else
            ConsolePrint STR_SUCCESS & vbCrLf
        End If
        GoTo QH
    End If
    '--- check for previous instance
    If Not bStarted And Not C_Bool(m_oOpt.Item("--hidden")) And Not C_Bool(m_oOpt.Item("--console")) Then
        If IsObjectRunning(STR_MONIKER) Then
            Select Case MsgBox(Printf(MSG_ALREADY_RUNNING, STR_MONIKER), vbQuestion Or vbYesNoCancel)
            Case vbYes
                GetObject(STR_MONIKER).ShowConfig
                GoTo QH
            Case vbCancel
                GoTo QH
            End Select
        End If
    End If
    '--- respawn hidden in systray
    If Not C_Bool(m_oOpt.Item("--console")) Then
        If Not C_Bool(m_oOpt.Item("--hidden")) And Not InIde Then
            frmMain.Restart AddParam:="--hidden"
            GoTo QH
        End If
        Process = -1
    End If
    '--- read config file
    If LenB(sConfFile) <> 0 Then
        If Not FileExists(sConfFile) Then
            DebugLog MODULE_NAME, FUNC_NAME, Printf(ERR_CONFIG_NOT_FOUND, sConfFile), vbLogEventTypeError
            Process = 1
            GoTo QH
        End If
        If Not JsonParse(ReadTextFile(sConfFile), m_oConfig, Error:=sError) Then
            DebugLog MODULE_NAME, FUNC_NAME, Printf(ERR_PARSING_CONFIG, sConfFile, sError), vbLogEventTypeError
            Process = 1
            GoTo QH
        End If
        DebugLog MODULE_NAME, FUNC_NAME, Printf(STR_LOADING_CONFIG, sConfFile)
    Else
        JsonItem(m_oConfig, "Printers/Autodetect") = True
        JsonItem(m_oConfig, "Endpoints/0/Binding") = "RestHttp"
        JsonItem(m_oConfig, "Endpoints/0/Address") = "127.0.0.1:" & DEF_LISTEN_PORT
    End If
    '--- setup environment and procotol configuration
    lIdx = JsonItem(m_oConfig, -1)
    If lIdx > 0 Then
        DebugLog MODULE_NAME, FUNC_NAME, Printf(STR_ENVIRON_VARS_FOUND, lIdx)
        sLogFile = GetEnvironmentVar("_UCS_FISCAL_PRINTER_LOG")
        sLogDataDump = GetEnvironmentVar("_UCS_FISCAL_PRINTER_DATA_DUMP")
        For Each vKey In JsonKeys(m_oConfig, "Environment")
            Call SetEnvironmentVariable(vKey, C_Str(JsonItem(m_oConfig, "Environment/" & vKey)))
        Next
        If sLogFile <> GetEnvironmentVar("_UCS_FISCAL_PRINTER_LOG") _
                Or sLogDataDump <> GetEnvironmentVar("_UCS_FISCAL_PRINTER_DATA_DUMP") Then
            Set Logger = Nothing
            Logger.Log 0, MODULE_NAME, FUNC_NAME, App.ProductName & " v" & STR_VERSION
        End If
        JsonExpandEnviron m_oConfig
    End If
    Set ProtocolConfig = C_Obj(JsonItem(m_oConfig, "ProtocolConfig"))
    '-- clear printers collection
    JsonItem(m_oPrinters, vbNullString) = Empty
    For Each vKey In JsonKeys(m_oPrinters)
        JsonItem(m_oPrinters, vKey) = Empty
    Next
    '--- first register local endpoints
    If Not pvCreateEndpoints(m_oPrinters, "local", m_cEndpoints) Then
        GoTo QH
    End If
    '--- leave longer to complete auto-detection for last step
    If Not pvCollectPrinters(m_oPrinters) Then
        GoTo QH
    End If
    lIdx = C_Lng(JsonItem(m_oPrinters, "Count"))
    DebugLog MODULE_NAME, FUNC_NAME, Printf(IIf(lIdx = 1, STR_ONE_PRINTER_FOUND, STR_PRINTERS_FOUND), lIdx)
    '--- then register http/mssql endpoints
    If Not pvCreateEndpoints(m_oPrinters, "resthttp mssqlservicebroker mysqlmessagequeue", m_cEndpoints) Then
        GoTo QH
    End If
    If m_bIsService Then
        Do While Not NtServiceQueryStop()
            '--- do nothing
        Loop
        TerminateEndpoints
        NtServiceTerminate
        FlushDebugLog
    ElseIf C_Bool(m_oOpt.Item("--console")) Then
        Screen.MousePointer = vbDefault
        ConsolePrint STR_PRESS_CTRLC & vbCrLf
        Do
            ConsoleRead
            DoEvents
            FlushDebugLog
        Loop
    Else
        If Not frmMain.Init(m_oPrinters, sConfFile, App.ProductName & " v" & STR_VERSION, GetEnvironmentVar("_UCS_FP_HUB_AUTO_UPDATE")) Then
            Process = 1
            GoTo QH
        End If
        If Not C_Bool(m_oOpt.Item("--systray")) Then
            MainForm.ShowConfig
        End If
    End If
QH:
    Screen.MousePointer = vbDefault
    Exit Function
EH:
    PrintError FUNC_NAME
    Process = 100
End Function

Private Function pvCollectPrinters(oRetVal As Object) As Boolean
    Const FUNC_NAME     As String = "pvCollectPrinters"
    Dim oFP             As cFiscalPrinter
    Dim sResponse       As String
    Dim oJson           As Object
    Dim vKey            As Variant
    Dim oRequest        As Object
    Dim sDeviceString   As String
    Dim sKey            As String
    Dim oAliases        As Object
    Dim oInfo           As Object
    
    On Error GoTo EH
    Set oFP = New cFiscalPrinter
    JsonItem(oRetVal, "Ok") = True
    JsonItem(oRetVal, "Count") = 0
    If JsonItem(m_oConfig, "Printers/Autodetect") Then
        DebugLog MODULE_NAME, FUNC_NAME, STR_AUTODETECTING_PRINTERS
        If oFP.EnumPorts(sResponse) And JsonParse(sResponse, oJson) Then
            If Not JsonItem(oJson, "Ok") Then
                DebugLog MODULE_NAME, FUNC_NAME, Printf(ERR_ENUM_PORTS, vKey, JsonItem(oJson, "ErrorText")), vbLogEventTypeError
            Else
                For Each vKey In JsonKeys(oJson, "SerialPorts")
                    If LenB(C_Str(JsonItem(oJson, "SerialPorts/" & vKey & "/Protocol"))) <> 0 Then
                        Set oInfo = JsonParseObject(JsonDump(JsonItem(oJson, "SerialPorts/" & vKey), Minimize:=True))
                        JsonItem(oInfo, "Model") = Empty
                        JsonItem(oInfo, "Firmware") = Empty
                        sDeviceString = ToDeviceString(oInfo)
                        Set oRequest = Nothing
                        JsonItem(oRequest, "DeviceString") = sDeviceString
                        JsonItem(oRequest, "IncludeTaxNo") = True
                        If oFP.GetDeviceInfo(JsonDump(oRequest, Minimize:=True), sResponse) And JsonParse(sResponse, oInfo) Then
                            sDeviceString = Zn(JsonItem(oInfo, "DeviceString"), sDeviceString)
                            sKey = JsonItem(oInfo, "DeviceSerialNo")
                            If LenB(sKey) <> 0 Then
                                JsonItem(oInfo, "Ok") = Empty
                                JsonItem(oInfo, "DeviceString") = sDeviceString
                                JsonItem(oInfo, "DeviceHost") = GetErrorComputerName()
                                JsonItem(oInfo, "DevicePort") = pvGetDevicePort(sDeviceString)
                                JsonItem(oInfo, "Autodetected") = True
                                JsonItem(oRetVal, sKey) = oInfo
                                JsonItem(oRetVal, "Count") = JsonItem(oRetVal, "Count") + 1
                            End If
                        End If
                    End If
                Next
            End If
        End If
    End If
    For Each vKey In JsonKeys(m_oConfig, "Printers")
        sDeviceString = C_Str(JsonItem(m_oConfig, "Printers/" & vKey & "/DeviceString"))
        If LenB(sDeviceString) <> 0 Then
            Set oRequest = Nothing
            JsonItem(oRequest, "DeviceString") = sDeviceString
            JsonItem(oRequest, "IncludeTaxNo") = True
            If oFP.GetDeviceInfo(JsonDump(oRequest, Minimize:=True), sResponse) And JsonParse(sResponse, oInfo) Then
                If Not JsonItem(oInfo, "Ok") Then
                    DebugLog MODULE_NAME, FUNC_NAME, Printf(ERR_WARN_ACCESS, vKey, JsonItem(oInfo, "ErrorText")), vbLogEventTypeWarning
                Else
                    sDeviceString = Zn(JsonItem(oInfo, "DeviceString"), sDeviceString)
                    sKey = Zn(JsonItem(oInfo, "DeviceSerialNo"), vKey)
                    If LenB(sKey) <> 0 Then
                        JsonItem(oInfo, "Ok") = Empty
                        JsonItem(oInfo, "DeviceString") = sDeviceString
                        JsonItem(oInfo, "DeviceHost") = GetErrorComputerName()
                        JsonItem(oInfo, "DevicePort") = pvGetDevicePort(sDeviceString)
                        JsonItem(oInfo, "Description") = JsonItem(m_oConfig, "Printers/" & vKey & "/Description")
                        If IsEmpty(JsonItem(oRetVal, sKey)) Then
                            JsonItem(oRetVal, "Count") = JsonItem(oRetVal, "Count") + 1
                        End If
                        JsonItem(oRetVal, sKey) = oInfo
                        If IsEmpty(JsonItem(oAliases, vKey)) Then
                            JsonItem(oAliases, "Count") = JsonItem(oAliases, "Count") + 1
                        End If
                        JsonItem(oAliases, vKey & "/DeviceSerialNo") = sKey
                    End If
                End If
            End If
        End If
    Next
    If Not oAliases Is Nothing Then
        JsonItem(oRetVal, "Aliases") = oAliases
    End If
    '--- success
    pvCollectPrinters = True
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function pvCreateEndpoints(oPrinters As Object, sBindings As String, cRetVal As Collection) As Boolean
    Const FUNC_NAME     As String = "pvCreateEndpoints"
    Dim vKey            As Variant
    Dim oRestEndpoint   As cRestEndpoint
    Dim oQueueEndpoint  As cQueueEndpoint
    Dim oLocalEndpoint  As frmLocalEndpoint
    
    On Error GoTo EH
    If cRetVal Is Nothing Then
        Set cRetVal = New Collection
    End If
    '--- first local endpoint (faster registration)
    For Each vKey In JsonKeys(m_oConfig, "Endpoints")
        If InStr(sBindings, LCase$(JsonItem(m_oConfig, "Endpoints/" & vKey & "/Binding"))) > 0 Then
            Select Case LCase$(JsonItem(m_oConfig, "Endpoints/" & vKey & "/Binding"))
            Case "local"
                Set oLocalEndpoint = New frmLocalEndpoint
                If oLocalEndpoint.frInit(JsonItem(m_oConfig, "Endpoints/" & vKey), oPrinters) Then
                    cRetVal.Add oLocalEndpoint
                End If
            Case "resthttp"
                Set oRestEndpoint = New cRestEndpoint
                If oRestEndpoint.Init(JsonItem(m_oConfig, "Endpoints/" & vKey), oPrinters) Then
                    cRetVal.Add oRestEndpoint
                End If
            Case "mssqlservicebroker", "mysqlmessagequeue"
                Set oQueueEndpoint = New cQueueEndpoint
                If oQueueEndpoint.Init(JsonItem(m_oConfig, "Endpoints/" & vKey), oPrinters) Then
                    cRetVal.Add oQueueEndpoint
                End If
            End Select
        End If
    Next
    '--- always init local endpoint
    If oLocalEndpoint Is Nothing And InStr(sBindings, "local") > 0 Then
        Set oLocalEndpoint = New frmLocalEndpoint
        If oLocalEndpoint.frInit(Nothing, oPrinters) Then
            cRetVal.Add oLocalEndpoint
        End If
    End If
    '--- success
    pvCreateEndpoints = True
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Public Sub DebugLog(sModule As String, sFunction As String, sText As String, Optional ByVal eType As LogEventTypeConstants = vbLogEventTypeInformation)
    Dim sPrefix         As String
    
    Logger.Log eType, sModule, sFunction, sText
    If Logger.LogFile = -1 And m_bIsService Then
        App.LogEvent sText, Clamp(eType, 0, vbLogEventTypeInformation)
    ElseIf Not m_bIsHidden Then
        sPrefix = Format$(GetCurrentNow, FORMAT_TIME_ONLY) & Right$(Format$(GetCurrentTimer, FORMAT_BASE_3), 4) & ": "
        Select Case eType
        Case vbLogEventTypeError
            sPrefix = sPrefix & STR_PREFIX_ERROR
        Case vbLogEventTypeWarning
            sPrefix = sPrefix & STR_PREFIX_WARNING
        End Select
        sPrefix = sPrefix & IIf(Len(sText) > 200, Left$(sText, 200) & "...", sText) & vbCrLf
        If eType = vbLogEventTypeError Then
            ConsoleColorError FOREGROUND_RED, FOREGROUND_MASK, sPrefix
        Else
            ConsolePrint sPrefix
        End If
    End If
End Sub

Public Property Get IsLogDebugEnabled() As Boolean
    IsLogDebugEnabled = Logger.LogLevel >= vbLogEventTypeDebug
End Property

Public Property Get IsLogDataDumpEnabled() As Boolean
    IsLogDataDumpEnabled = Logger.LogLevel >= vbLogEventTypeDataDump
End Property

Public Sub TerminateEndpoints()
    Dim oElem           As IEndpoint
    
    If Not m_cEndpoints Is Nothing Then
        For Each oElem In m_cEndpoints
            oElem.Terminate
        Next
        Set m_cEndpoints = Nothing
    End If
End Sub

Private Function pvRegisterServiceAppID(sServiceName As String, sDisplayName As String, sExeFile As String, sGuid As String, Optional Error As String) As Boolean
    If Not pvRegSetStringValue(HKEY_CLASSES_ROOT, "AppID\" & sExeFile, "AppID", sGuid) Then
        GoTo QH
    End If
    If Not pvRegSetStringValue(HKEY_CLASSES_ROOT, "AppID\" & sGuid, vbNullString, sDisplayName) Then
        GoTo QH
    End If
    If Not pvRegSetStringValue(HKEY_CLASSES_ROOT, "AppID\" & sGuid, "LocalService", sServiceName) Then
        GoTo QH
    End If
    '--- success
    pvRegisterServiceAppID = True
QH:
    If Not pvRegisterServiceAppID Then
        Error = Printf(STR_REGISTER_APPID_FAILED, GetErrorDescription(Err.LastDllError))
    End If
End Function

Private Function pvRegSetStringValue(ByVal hRoot As Long, sSubKey As String, sName As String, sValue As String) As Boolean
    Dim hKey            As Long
    Dim dwDummy         As Long
    
    If RegCreateKeyEx(hRoot, sSubKey, 0, 0, 0, SAM_WRITE, 0, hKey, dwDummy) = 0 Then
        Call RegCloseKey(hKey)
    End If
    If RegOpenKeyEx(hRoot, sSubKey, 0, SAM_WRITE, hKey) <> 0 Then
        GoTo QH
    End If
    If RegSetValueEx(hKey, sName, 0, REG_SZ, ByVal sValue, Len(sValue)) <> 0 Then
        GoTo QH
    End If
    '--- success
    pvRegSetStringValue = True
QH:
    If hKey <> 0 Then
        Call RegCloseKey(hKey)
    End If
End Function

Private Function pvUnregisterServiceAppID(sExeFile As String, sGuid As String, Optional Error As String) As Boolean
    SHDeleteKey HKEY_CLASSES_ROOT, "AppID\" & sExeFile
    SHDeleteKey HKEY_CLASSES_ROOT, "AppID\" & sGuid
    Error = vbNullString
    '--- success
    pvUnregisterServiceAppID = True
End Function

Private Function pvGetDevicePort(sDeviceString As String) As String
    Dim oJson           As Object
    Dim sRetVal         As String
    Dim lPos            As Long
    
    Set oJson = ParseDeviceString(sDeviceString)
    If Not IsEmpty(JsonItem(oJson, "Url")) Then
        sRetVal = JsonItem(oJson, "Url")
        lPos = InStr(sRetVal, "://") + 3
        If lPos > 3 And InStr(lPos, sRetVal, "/") > 0 Then
            sRetVal = Left$(sRetVal, InStr(lPos, sRetVal, "/") - 1)
        End If
    ElseIf Not IsEmpty(JsonItem(oJson, "IP")) Then
        sRetVal = JsonItem(oJson, "Port")
        sRetVal = JsonItem(oJson, "IP") & IIf(LenB(sRetVal) <> 0, ":" & sRetVal, vbNullString)
    Else
        sRetVal = JsonItem(oJson, "Speed")
        If sRetVal = "115200" Then
            sRetVal = vbNullString
        End If
        sRetVal = JsonItem(oJson, "Port") & IIf(LenB(sRetVal) <> 0, "," & sRetVal, vbNullString)
    End If
    pvGetDevicePort = sRetVal
End Function
