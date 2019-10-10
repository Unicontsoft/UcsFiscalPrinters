Attribute VB_Name = "mdStartup"
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
Private Const MODULE_NAME As String = "mdStartup"

'=========================================================================
' API
'=========================================================================

Private Const HKEY_CLASSES_ROOT         As Long = &H80000000
Private Const SAM_WRITE                 As Long = &H20007
Private Const REG_SZ                    As Long = 1

Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Private Declare Function SetEnvironmentVariable Lib "kernel32" Alias "SetEnvironmentVariableA" (ByVal lpName As String, ByVal lpValue As String) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As Long, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function SHDeleteKey Lib "shlwapi" Alias "SHDeleteKeyA" (ByVal hKey As Long, ByVal szSubKey As String) As Long

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_VERSION               As String = "0.1.24"
Private Const STR_SERVICE_NAME          As String = "UcsFPHub"
Private Const STR_DISPLAY_NAME          As String = "Unicontsoft Fiscal Printers Hub (" & STR_VERSION & ")"
Private Const STR_APPID_GUID            As String = "{6E78E71A-35B2-4D23-A88C-4C2858430329}"
Private Const STR_SVC_INSTALL           As String = "Инсталира NT услуга %1..."
Private Const STR_SVC_UNINSTALL         As String = "Деинсталира NT услуга %1..."
Private Const STR_SUCCESS               As String = "Успех"
Private Const STR_FAILURE               As String = "Грешка: "
Private Const STR_WARN                  As String = "Предупреждение: "
Private Const STR_AUTODETECTING_PRINTERS As String = "Автоматично търсене на принтери"
Private Const STR_ENVIRON_VARS_FOUND    As String = "Конфигурирани %1 променливи на средата"
Private Const STR_PRINTERS_FOUND        As String = "Намерени %1 принтера"
Private Const STR_PRESS_CTRLC           As String = "Натиснете Ctrl+C за изход"
Private Const STR_LOADING_CONFIG        As String = "Зарежда конфигурация от %1"
'--- errors
Private Const ERR_CONFIG_NOT_FOUND      As String = "Грешка: Конфигурационен файл %1 не е намерен"
Private Const ERR_PARSING_CONFIG        As String = "Грешка: Невалиден %1: %2"
Private Const ERR_ENUM_PORTS            As String = "Грешка: Енумериране на серийни портове: %1"
Private Const ERR_WARN_ACCESS           As String = "Предупреждение: Принтер %1: %2"
Private Const ERR_REGISTER_APPID_FAILED As String = "Неуспешна регистрация на AppID. %1"
'--- formats
Private Const FORMAT_DATETIME_LOG       As String = "yyyy.MM.dd hh:nn:ss"
Private Const FORMAT_BASE_3             As String = "0.000"

Private m_oOpt                      As Object
Private m_oPrinters                 As Object
Private m_oConfig                   As Object
Private m_cEndpoints                As Collection
Private m_bIsService                As Boolean
Private m_nDebugLogFile             As Integer

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunction As String)
    Debug.Print "Critical error: " & Err.Description & " [" & MODULE_NAME & "." & sFunction & "]"
    DebugLog Err.Description & " [" & MODULE_NAME & "." & sFunction & "(" & Erl & ")]", vbLogEventTypeError
End Sub

'=========================================================================
' Properties
'=========================================================================

Property Get IsRunningAsService() As Boolean
    IsRunningAsService = m_bIsService
End Property

'=========================================================================
' Functions
'=========================================================================

Public Sub Main()
    Dim lExitCode       As Long
    
    lExitCode = Process(SplitArgs(Command$))
    If Not InIde And lExitCode <> -1 Then
        Call ExitProcess(lExitCode)
    End If
End Sub

Private Function Process(vArgs As Variant) As Long
    Const FUNC_NAME     As String = "Process"
    Dim sConfFile       As String
    Dim sError          As String
    Dim vKey            As Variant
    Dim lIdx            As Long
    
    On Error GoTo EH
    Set m_oOpt = GetOpt(vArgs, "config:-config:c")
    '--- normalize options: convert -o and -option to proper long form (--option)
    For Each vKey In Split("nologo config:c install:i uninstall:u systray:s hidden help:h:?")
        vKey = Split(vKey, ":")
        For lIdx = 0 To UBound(vKey)
            If IsEmpty(m_oOpt.Item("--" & At(vKey, 0))) And Not IsEmpty(m_oOpt.Item("-" & At(vKey, lIdx))) Then
                m_oOpt.Item("--" & At(vKey, 0)) = m_oOpt.Item("-" & At(vKey, lIdx))
            End If
        Next
    Next
    If Not m_oOpt.Item("--nologo") Then
        ConsolePrint App.ProductName & " v" & STR_VERSION & " (c) 2019 by Unicontsoft" & vbCrLf & vbCrLf
    End If
    If m_oOpt.Item("--help") Then
        ConsolePrint "Usage: " & App.EXEName & ".exe [options...]" & vbCrLf & vbCrLf & _
                    "Options:" & vbCrLf & _
                    "  -c, --config FILE   read configuration from FILE" & vbCrLf & _
                    "  -i, --install       install NT service (with config file from -c option)" & vbCrLf & _
                    "  -u, --uninstall     remove NT service" & vbCrLf & _
                    "  -s, --systray       show icon in systray" & vbCrLf
        GoTo QH
    End If
    If NtServiceInit(STR_SERVICE_NAME) Then
        m_bIsService = True
        '--- cannot handle these as NT service
        m_oOpt.Item("--systray") = Empty
        m_oOpt.Item("--install") = Empty
        m_oOpt.Item("--uninstall") = Empty
    End If
    '--- read config file
    sConfFile = m_oOpt.Item("--config")
    If LenB(sConfFile) = 0 Then
        sConfFile = PathCombine(App.Path, App.EXEName & ".conf")
        If Not FileExists(sConfFile) Then
            sConfFile = PathCombine(GetSpecialFolder(ucsOdtLocalAppData) & "\Unicontsoft\UcsFPHub", App.EXEName & ".conf")
            If Not FileExists(sConfFile) Then
                sConfFile = vbNullString
            End If
        End If
    End If
    If LenB(sConfFile) <> 0 Then
        DebugLog Printf(STR_LOADING_CONFIG, sConfFile) & " [" & MODULE_NAME & "." & FUNC_NAME & "]"
        If Not FileExists(sConfFile) Then
            DebugLog Printf(ERR_CONFIG_NOT_FOUND, sConfFile) & " [" & MODULE_NAME & "." & FUNC_NAME & "]", vbLogEventTypeError
            Process = 1
            GoTo QH
        End If
        If Not JsonParse(ReadTextFile(sConfFile), m_oConfig, Error:=sError) Then
            DebugLog Printf(ERR_PARSING_CONFIG, sConfFile, sError) & " [" & MODULE_NAME & "." & FUNC_NAME & "]", vbLogEventTypeError
            Process = 1
            GoTo QH
        End If
        JsonExpandEnviron m_oConfig
    Else
        JsonItem(m_oConfig, "Printers/Autodetect") = True
        JsonItem(m_oConfig, "Endpoints/0/Binding") = "RestHttp"
        JsonItem(m_oConfig, "Endpoints/0/Address") = "127.0.0.1:8192"
    End If
    If m_oOpt.Item("--systray") Then
        If Not m_oOpt.Item("--hidden") And Not InIde Then
            frmIcon.Restart AddParam:="--hidden"
            GoTo QH
        ElseIf Not frmIcon.Init(m_oOpt, sConfFile, App.ProductName & " v" & STR_VERSION) Then
            Process = 1
            GoTo QH
        End If
        Process = -1
    End If
    If m_oOpt.Item("--install") Then
        ConsolePrint Printf(STR_SVC_INSTALL, STR_SERVICE_NAME) & vbCrLf
        If LenB(sConfFile) <> 0 Then
            sConfFile = " --config " & ArgvQuote(sConfFile)
        End If
        If Not pvRegisterServiceAppID(STR_SERVICE_NAME, STR_DISPLAY_NAME, App.EXEName & ".exe", STR_APPID_GUID, Error:=sError) Then
            ConsoleError STR_WARN & sError & vbCrLf
        End If
        If Not NtServiceInstall(STR_SERVICE_NAME, STR_DISPLAY_NAME, GetProcessName() & sConfFile, Error:=sError) Then
            ConsoleError STR_FAILURE
            ConsoleColorError FOREGROUND_RED, FOREGROUND_MASK, sError & vbCrLf
        Else
            ConsolePrint STR_SUCCESS & vbCrLf
        End If
        GoTo QH
    ElseIf m_oOpt.Item("--uninstall") Then
        ConsolePrint Printf(STR_SVC_UNINSTALL, STR_SERVICE_NAME) & vbCrLf
        If Not pvUnregisterServiceAppID(App.EXEName & ".exe", STR_APPID_GUID, Error:=sError) Then
            ConsoleError STR_WARN & sError & vbCrLf
        End If
        If Not NtServiceUninstall(STR_SERVICE_NAME, Error:=sError) Then
            ConsoleError STR_FAILURE
            ConsoleColorError FOREGROUND_RED, FOREGROUND_MASK, sError
        Else
            ConsolePrint STR_SUCCESS & vbCrLf
        End If
        GoTo QH
    End If
    If UBound(JsonKeys(m_oConfig, "Environment")) >= 0 Then
        DebugLog Printf(STR_ENVIRON_VARS_FOUND, UBound(JsonKeys(m_oConfig, "Environment")) + 1) & " [" & MODULE_NAME & "." & FUNC_NAME & "]"
        For Each vKey In JsonKeys(m_oConfig, "Environment")
            Call SetEnvironmentVariable(vKey, C_Str(JsonItem(m_oConfig, "Environment/" & vKey)))
        Next
        FlushDebugLog
        m_nDebugLogFile = 0
    End If
    Set m_oPrinters = pvCollectPrinters()
    DebugLog Printf(STR_PRINTERS_FOUND, JsonItem(m_oPrinters, "Count")) & " [" & MODULE_NAME & "." & FUNC_NAME & "]"
    Set m_cEndpoints = pvCreateEndpoints(m_oPrinters)
    If m_bIsService Then
        Do While Not NtServiceQueryStop()
            '--- do nothing
        Loop
        TerminateEndpoints
        NtServiceTerminate
        FlushDebugLog
    ElseIf Not m_oOpt.Item("--systray") Then
        ConsolePrint STR_PRESS_CTRLC & vbCrLf
        Do
            ConsoleRead
            DoEvents
            FlushDebugLog
        Loop
    End If
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
    Process = 100
End Function

Private Function pvCollectPrinters() As Object
    Const FUNC_NAME     As String = "pvCollectPrinters"
    Dim oFP             As cFiscalPrinter
    Dim sResponse       As String
    Dim oJson           As Object
    Dim vKey            As Variant
    Dim oRequest        As Object
    Dim oRetVal         As Object
    Dim sDeviceString   As String
    Dim sKey            As String
    Dim oAliases        As Object
    
    On Error GoTo EH
    Set oFP = New cFiscalPrinter
    JsonItem(oRetVal, "Ok") = True
    JsonItem(oRetVal, "Count") = 0
    If JsonItem(m_oConfig, "Printers/Autodetect") Then
        DebugLog STR_AUTODETECTING_PRINTERS & " [" & MODULE_NAME & "." & FUNC_NAME & "]"
        If oFP.EnumPorts(sResponse) And JsonParse(sResponse, oJson) Then
            If Not JsonItem(oJson, "Ok") Then
                DebugLog Printf(ERR_ENUM_PORTS, vKey, JsonItem(oJson, "ErrorText")) & " [" & MODULE_NAME & "." & FUNC_NAME & "]", vbLogEventTypeError
            Else
                For Each vKey In JsonKeys(oJson, "SerialPorts")
                    If LenB(JsonItem(oJson, "SerialPorts/" & vKey & "/Protocol")) <> 0 Then
                        sDeviceString = "Protocol=" & JsonItem(oJson, "SerialPorts/" & vKey & "/Protocol") & _
                            ";Port=" & JsonItem(oJson, "SerialPorts/" & vKey & "/Port") & _
                            ";Speed=" & JsonItem(oJson, "SerialPorts/" & vKey & "/Speed")
                        Set oRequest = Nothing
                        JsonItem(oRequest, "DeviceString") = sDeviceString
                        JsonItem(oRequest, "IncludeTaxNo") = True
                        If oFP.GetDeviceInfo(JsonDump(oRequest, Minimize:=True), sResponse) And JsonParse(sResponse, oJson) Then
                            sKey = JsonItem(oJson, "DeviceSerialNo")
                            If LenB(sKey) <> 0 Then
                                JsonItem(oJson, "Ok") = Empty
                                JsonItem(oJson, "DeviceString") = sDeviceString
                                JsonItem(oJson, "Host") = GetErrorComputerName()
                                JsonItem(oJson, "Device") = pvToSimpleDevice(sDeviceString)
                                JsonItem(oJson, "Autodetected") = True
                                JsonItem(oRetVal, sKey) = oJson
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
            If oFP.GetDeviceInfo(JsonDump(oRequest, Minimize:=True), sResponse) And JsonParse(sResponse, oJson) Then
                If Not JsonItem(oJson, "Ok") Then
                    DebugLog Printf(ERR_WARN_ACCESS, vKey, JsonItem(oJson, "ErrorText")) & " [" & MODULE_NAME & "." & FUNC_NAME & "]", vbLogEventTypeWarning
                Else
                    sKey = JsonItem(oJson, "DeviceSerialNo")
                    If LenB(sKey) <> 0 Then
                        JsonItem(oJson, "Ok") = Empty
                        JsonItem(oJson, "DeviceString") = sDeviceString
                        JsonItem(oJson, "Host") = GetErrorComputerName()
                        JsonItem(oJson, "Device") = pvToSimpleDevice(sDeviceString)
                        JsonItem(oJson, "Description") = JsonItem(m_oConfig, "Printers/" & vKey & "/Description")
                        If IsEmpty(JsonItem(oRetVal, sKey)) Then
                            JsonItem(oRetVal, "Count") = JsonItem(oRetVal, "Count") + 1
                        End If
                        JsonItem(oRetVal, sKey) = oJson
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
    Set pvCollectPrinters = oRetVal
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function pvCreateEndpoints(oPrinters As Object) As Collection
    Const FUNC_NAME     As String = "pvCreateEndpoints"
    Dim cRetVal         As Collection
    Dim vKey            As Variant
    Dim oRestEndpoint   As cRestEndpoint
    Dim oMssqlEndpoint  As cMssqlEndpoint
    Dim oLocalEndpoint  As frmLocalEndpoint
    
    On Error GoTo EH
    Set cRetVal = New Collection
    For Each vKey In JsonKeys(m_oConfig, "Endpoints")
        Select Case LCase$(JsonItem(m_oConfig, "Endpoints/" & vKey & "/Binding"))
        Case "resthttp"
            Set oRestEndpoint = New cRestEndpoint
            If oRestEndpoint.Init(JsonItem(m_oConfig, "Endpoints/" & vKey), oPrinters) Then
                cRetVal.Add oRestEndpoint
            End If
        Case "mssqlservicebroker"
            Set oMssqlEndpoint = New cMssqlEndpoint
            If oMssqlEndpoint.Init(JsonItem(m_oConfig, "Endpoints/" & vKey), oPrinters) Then
                cRetVal.Add oMssqlEndpoint
            End If
        Case "local"
            Set oLocalEndpoint = New frmLocalEndpoint
            If oLocalEndpoint.frInit(JsonItem(m_oConfig, "Endpoints/" & vKey), oPrinters) Then
                cRetVal.Add oLocalEndpoint
            End If
        End Select
    Next
    '--- always init local endpoint
    If oLocalEndpoint Is Nothing Then
        Set oLocalEndpoint = New frmLocalEndpoint
        If oLocalEndpoint.frInit(Nothing, oPrinters) Then
            cRetVal.Add oLocalEndpoint
        End If
    End If
    Set pvCreateEndpoints = cRetVal
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Public Sub DebugDataDump(sText As String, Optional ByVal eType As LogEventTypeConstants = vbLogEventTypeInformation)
    Static lLogging     As Long
    
    If lLogging = 0 Then
        lLogging = IIf(CBool(Val(GetEnvironmentVar("_UCS_FISCAL_PRINTER_DATA_DUMP"))), 1, -1)
    End If
    If lLogging < 0 Then
        Exit Sub
    End If
    DebugLog sText, eType
End Sub

Public Sub DebugLog(sText As String, Optional ByVal eType As LogEventTypeConstants = vbLogEventTypeInformation)
    Dim sFile           As String
    Dim sPrefix         As String
    
    sPrefix = GetCurrentProcessId() & ": " & GetCurrentThreadId() & ": " & "(" & Format$(Now, FORMAT_DATETIME_LOG) & Right$(Format$(TimerEx, FORMAT_BASE_3), 4) & "): "
    If m_nDebugLogFile <> -1 Then
        If m_nDebugLogFile = 0 Then
            sFile = GetEnvironmentVar("_UCS_FP_HUB_LOG")
            If LenB(sFile) = 0 Then
                sFile = GetErrorTempPath() & "\UcsFPHub.log"
                If Not FileExists(sFile) Then
                    m_nDebugLogFile = -1
                    GoTo NoLogFile
                End If
            End If
            m_nDebugLogFile = FreeFile
            Open sFile For Append Access Write Shared As #m_nDebugLogFile
        End If
        Print #m_nDebugLogFile, sPrefix & sText
        Debug.Print sPrefix & sText
    Else
NoLogFile:
        If m_bIsService Then
            App.LogEvent sText, eType
            GoTo QH
        End If
        If eType = vbLogEventTypeError Then
            ConsoleColorError FOREGROUND_RED, FOREGROUND_MASK, sPrefix & sText & vbCrLf
        Else
            ConsolePrint sPrefix & sText & vbCrLf
        End If
    End If
QH:
    FlushDebugLog
End Sub

Public Sub FlushDebugLog()
    If m_nDebugLogFile <> 0 And m_nDebugLogFile <> -1 Then
        Close #m_nDebugLogFile
        m_nDebugLogFile = 0
    End If
End Sub

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
        Error = Printf(ERR_REGISTER_APPID_FAILED, GetErrorDescription(Err.LastDllError))
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

Private Function pvToSimpleDevice(sDeviceString As String) As String
    Dim oJson           As Object
    Dim sDevice         As String
    
    Set oJson = ParseDeviceString(sDeviceString)
    If IsEmpty(JsonItem(oJson, "IP")) Then
        sDevice = JsonItem(oJson, "Speed")
        If sDevice = "115200" Then
            sDevice = vbNullString
        End If
        sDevice = JsonItem(oJson, "Port") & IIf(LenB(sDevice) <> 0, "," & sDevice, vbNullString)
    Else
        sDevice = JsonItem(oJson, "Port")
        sDevice = JsonItem(oJson, "IP") & IIf(LenB(sDevice) <> 0, ":" & sDevice, vbNullString)
    End If
    pvToSimpleDevice = sDevice
End Function
