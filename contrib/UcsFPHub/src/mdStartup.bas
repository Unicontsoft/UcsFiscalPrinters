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


Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Private Declare Function SetEnvironmentVariable Lib "kernel32" Alias "SetEnvironmentVariableA" (ByVal lpName As String, ByVal lpValue As String) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_VERSION               As String = "0.1.16"
Private Const STR_SERVICE_NAME          As String = "UcsFPHub"
Private Const STR_DISPLAY_NAME          As String = "Unicontsoft Fiscal Printers Hub (" & STR_VERSION & ")"
Private Const STR_SVC_INSTALL           As String = "��������� NT ������ %1..."
Private Const STR_SVC_UNINSTALL         As String = "����������� NT ������ %1..."
Private Const STR_SUCCESS               As String = "�����"
Private Const STR_FAILURE               As String = "������: "
Private Const STR_AUTODETECTING_PRINTERS As String = "����������� ������� �� ��������"
Private Const STR_ENVIRON_VARS_FOUND    As String = "������������� %1 ���������� �� �������"
Private Const STR_PRINTERS_FOUND        As String = "�������� %1 ��������"
Private Const STR_PRESS_CTRLC           As String = "��������� Ctrl+C �� �����"
Private Const STR_LOADING_CONFIG        As String = "������� ������������ �� %1"
'--- errors
Private Const ERR_CONFIG_NOT_FOUND      As String = "������: ��������������� ���� %1 �� � �������"
Private Const ERR_PARSING_CONFIG        As String = "������: ��������� %1: %2"
Private Const ERR_ENUM_PORTS            As String = "������: ����������� �� ������� �������: %1"
Private Const ERR_WARN_ACCESS           As String = "��������������: ������� %1: %2"
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
    DebugLog Err.Description & " [" & MODULE_NAME & "." & sFunction & "]", vbLogEventTypeError
End Sub

'=========================================================================
' Functions
'=========================================================================

Private Sub Main()
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
    
    On Error GoTo EH
    Set m_oOpt = GetOpt(vArgs, "conf:c")
    '--- normalize options: convert -o and -option to proper long form (--option)
    For Each vKey In Split("nologo config:c install:i uninstall:u systray:s hidden")
        vKey = Split(vKey, ":")
        If IsEmpty(m_oOpt.Item("--" & At(vKey, 0))) And Not IsEmpty(m_oOpt.Item("-" & At(vKey, 0))) Then
            m_oOpt.Item("--" & At(vKey, 0)) = m_oOpt.Item("-" & At(vKey, 0))
        End If
        If LenB(At(vKey, 1)) <> 0 Then
            If IsEmpty(m_oOpt.Item("--" & At(vKey, 0))) And Not IsEmpty(m_oOpt.Item("-" & At(vKey, 1))) Then
                m_oOpt.Item("--" & At(vKey, 0)) = m_oOpt.Item("-" & At(vKey, 1))
            End If
        End If
    Next
    If Not m_oOpt.Item("--nologo") Then
        ConsolePrint App.ProductName & " v" & STR_VERSION & " (c) 2019 by Unicontsoft" & vbCrLf & vbCrLf
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
        DebugLog Printf(STR_LOADING_CONFIG, sConfFile)
        If Not FileExists(sConfFile) Then
            DebugLog Printf(ERR_CONFIG_NOT_FOUND, sConfFile), vbLogEventTypeError
            Process = 1
            GoTo QH
        End If
        If Not JsonParse(ReadTextFile(sConfFile), m_oConfig, Error:=sError) Then
            DebugLog Printf(ERR_PARSING_CONFIG, sConfFile, sError), vbLogEventTypeError
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
            frmIcon.Restart "--hidden"
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
        If Not NtServiceInstall(STR_SERVICE_NAME, STR_DISPLAY_NAME, GetProcessName() & sConfFile, Error:=sError) Then
            ConsoleError STR_FAILURE
            ConsoleColorError FOREGROUND_RED, FOREGROUND_MASK, sError & vbCrLf
        Else
            ConsolePrint STR_SUCCESS & vbCrLf
        End If
        GoTo QH
    ElseIf m_oOpt.Item("--uninstall") Then
        ConsolePrint Printf(STR_SVC_UNINSTALL, STR_SERVICE_NAME) & vbCrLf
        If Not NtServiceUninstall(STR_SERVICE_NAME, Error:=sError) Then
            ConsoleError STR_FAILURE
            ConsoleColorError FOREGROUND_RED, FOREGROUND_MASK, sError
        Else
            ConsolePrint STR_SUCCESS & vbCrLf
        End If
        GoTo QH
    End If
    If UBound(JsonKeys(m_oConfig, "Environment")) >= 0 Then
        DebugLog Printf(STR_ENVIRON_VARS_FOUND, UBound(JsonKeys(m_oConfig, "Environment")) + 1)
        For Each vKey In JsonKeys(m_oConfig, "Environment")
            Call SetEnvironmentVariable(vKey, C_Str(JsonItem(m_oConfig, "Environment/" & vKey)))
        Next
        FlushDebugLog
        m_nDebugLogFile = 0
    End If
    Set m_oPrinters = pvCollectPrinters()
    DebugLog Printf(STR_PRINTERS_FOUND, JsonItem(m_oPrinters, "Count"))
    Set m_cEndpoints = pvCreateEndpoints(m_oPrinters)
    If m_bIsService Then
        Do While Not NtServiceQueryStop()
            '--- do nothing
        Loop
        NtServiceTerminate
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
    
    On Error GoTo EH
    Set oFP = New cFiscalPrinter
    JsonItem(oRetVal, "Ok") = True
    JsonItem(oRetVal, "Count") = 0
    If JsonItem(m_oConfig, "Printers/Autodetect") Then
        DebugLog STR_AUTODETECTING_PRINTERS
        If oFP.EnumPorts(sResponse) And JsonParse(sResponse, oJson) Then
            If Not JsonItem(oJson, "Ok") Then
                DebugLog Printf(ERR_ENUM_PORTS, vKey, JsonItem(oJson, "ErrorText")), vbLogEventTypeError
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
                    DebugLog Printf(ERR_WARN_ACCESS, vKey, JsonItem(oJson, "ErrorText")), vbLogEventTypeWarning
                Else
                    sKey = JsonItem(oJson, "DeviceSerialNo")
                    If LenB(sKey) <> 0 Then
                        JsonItem(oJson, "Ok") = Empty
                        JsonItem(oJson, "DeviceString") = sDeviceString
                        JsonItem(oJson, "Host") = GetErrorComputerName()
                        JsonItem(oJson, "Description") = JsonItem(m_oConfig, "Printers/" & vKey & "/Description")
                        JsonItem(oRetVal, "Count") = JsonItem(oRetVal, "Count") + 1
                        JsonItem(oRetVal, "Aliases/Count") = JsonItem(oRetVal, "Aliases/Count") + 1
                        JsonItem(oRetVal, "Aliases/" & vKey & "/DeviceSerialNo") = sKey
                        JsonItem(oRetVal, sKey) = oJson
                    End If
                End If
            End If
        End If
    Next
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
            If oLocalEndpoint.Init(JsonItem(m_oConfig, "Endpoints/" & vKey), oPrinters) Then
                cRetVal.Add oLocalEndpoint
            End If
        End Select
    Next
    '--- always init local endpoint
    If oLocalEndpoint Is Nothing Then
        Set oLocalEndpoint = New frmLocalEndpoint
        If oLocalEndpoint.Init(Nothing, oPrinters) Then
            cRetVal.Add oLocalEndpoint
        End If
    End If
    Set pvCreateEndpoints = cRetVal
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

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
    Dim vElem           As Variant
    
    If Not m_cEndpoints Is Nothing Then
        For Each vElem In m_cEndpoints
            vElem.Terminate
        Next
        Set m_cEndpoints = Nothing
    End If
End Sub

