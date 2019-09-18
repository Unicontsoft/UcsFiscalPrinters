VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMssqlEndpoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
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
Private Const MODULE_NAME As String = "cMssqlEndpoint"
Implements IEndpoint

#Const ImplNoIdeProtection = (MST_NO_IDE_PROTECTION <> 0)
#Const ImplSelfContained = True

'=========================================================================
' API
'=========================================================================

'--- for winsock
Private Const AF_INET                   As Long = 2
Private Const SOCK_STREAM               As Long = 1
Private Const IPPROTO_TCP               As Long = 6
Private Const INVALID_SOCKET            As Long = -1
'--- for SetInformationJobObject
Private Const JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE As Long = &H2000
'--- for thunks
Private Const MEM_COMMIT                    As Long = &H1000
Private Const PAGE_EXECUTE_READWRITE        As Long = &H40
Private Const CRYPT_STRING_BASE64           As Long = 1

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function WSAStartup Lib "ws2_32" (ByVal wVersionRequired As Long, lpWSADATA As Any) As Long
Private Declare Function WSACleanup Lib "ws2_32" () As Long
Private Declare Function SocketInit Lib "ws2_32" Alias "socket" (ByVal lAf As Long, ByVal lSocketType As Long, ByVal lProtocol As Long) As Long
Private Declare Function SocketBind Lib "ws2_32" Alias "bind" (ByVal lSocket As Long, ByRef uName As SOCKADDR_IN, ByRef lNameLen As Long) As Long
Private Declare Function SocketListen Lib "ws2_32" Alias "listen" (ByVal lSocket As Long, ByVal lBackLog As Long) As Long
Private Declare Function SocketClose Lib "ws2_32" Alias "closesocket" (ByVal lSocket As Long) As Long
Private Declare Function SocketGetIP Lib "ws2_32" Alias "inet_addr" (ByVal sConnPoint As String) As Long
Private Declare Function SocketnTons Lib "ws2_32" Alias "htons" (ByVal lNetShort As Long) As Integer
'--- for thunks
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function CryptStringToBinary Lib "crypt32" Alias "CryptStringToBinaryA" (ByVal pszString As String, ByVal cchString As Long, ByVal dwFlags As Long, ByVal pbBinary As Long, pcbBinary As Long, Optional ByVal pdwSkip As Long, Optional ByVal pdwFlags As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
#If Not ImplNoIdeProtection Then
    Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, ByVal lpszClass As String, ByVal lpszWindow As String) As Long
    Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
#End If
#If ImplSelfContained Then
    Private Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long
    Private Declare Function SetEnvironmentVariable Lib "kernel32" Alias "SetEnvironmentVariableA" (ByVal lpName As String, ByVal lpValue As String) As Long
#End If

Private Type SOCKADDR_IN
    Family              As Integer
    Port                As Integer
    Addr                As Long
    Zero(7)             As Byte
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_CONN_OPEN             As String = "Свързан с %1"
Private Const STR_CONN_REOPEN           As String = "Възстановена връзка с %1"
Private Const STR_QUEUE_SETUP           As String = "Слуша на опашка %2 за %1 услуги"
Private Const STR_QUEUE_ELAPSED         As String = "%1 на принтер %2 изпълни заявка %3 за %4 сек."
Private Const STR_QUEUE_ERROR           As String = "Грешка в %1 на принтер %2: %3"
Private Const STR_CROSS_PRINTER_REQUEST As String = "Заявка %1 не е към устройството на услуга %2"
Private Const STR_DATETIME_SETUP        As String = "Установен час %2 на принтер %1"
Private Const STR_DATETIME_ERROR        As String = "Грешка от принтер %1: %2"
Private Const STR_QUEUE_PREFIX          As String = "UcsFpTargetQueue"
Private Const STR_SVC_PREFIX            As String = "UcsFpTargetService"
'--- errors
Private Const ERR_INVALID_SERVICE       As String = "Невалидна Service Broker услуга: %1"
Private Const ERR_INVALID_URL           As String = "Невалиден адрес: %1"
Private Const ERR_INVALID_XML           As String = "Невалиден XML"
Private Const ERR_UNKNOWN               As String = "Неизвестна грешка %1"
Private Const ERR_DB_E_NOTABLE          As Long = &H80040E37
'--- queue
Private Const DEF_QUEUE_TIMEOUT         As Long = 5000
Private Const STR_QUEUE_LOOP            As String = "DECLARE     @QueueName SYSNAME = '%1', @Result INT, @Handle UNIQUEIDENTIFIER, @Request NVARCHAR(MAX), @MsgType SYSNAME, @SvcName SYSNAME, @ErrorText NVARCHAR(255)" & vbCrLf & _
                                                    "EXEC        @Result = dbo.usp_sys_ServiceBrokerWaitRequest @QueueName, %2, @Handle OUTPUT, @Request OUTPUT, @MsgType OUTPUT, @SvcName OUTPUT, @ErrorText OUTPUT" & vbCrLf & _
                                                    "SELECT      @Result AS Result, @Handle AS Handle, @Request AS Request, @MsgType AS MsgType, @SvcName AS SvcName, @ErrorText AS ErrorText" & vbCrLf & _
                                                    "RAISERROR   ('Result=%d', 10, 0, @Result) WITH NOWAIT" & vbCrLf
Private Const LNG_QUEUE_FIELDS          As Long = 6
Private Const STR_SEND_ON_CONV          As String = "SEND ON CONVERSATION '%1' (N'%2')"
Private Const STR_END_CONV_WITH_ERROR   As String = "END CONVERSATION '%1' WITH ERROR = %2 DESCRIPTION = '%3'"
'--- params
Private Const PRM_SQLSERVER             As String = "SqlServer"
Private Const PRM_DBUSER                As String = "DbUser"
Private Const PRM_DBPASS                As String = "DbPass"
Private Const PRM_CONNSTR               As String = "ConnStr"
Private Const PRM_SSHSETTINGS           As String = "SshSettings"
'--- SSH consts
Private Const STR_SSH_TUNNEL_ERROR              As String = "Грешка при отваряне на SSH тунел" & vbCrLf & vbCrLf & "%1"
Private Const STR_SSH_TUNNEL_NO_PORT_AVAILABLE  As String = "Невъзможно отваряне на локален порт"
Private Const STR_SSH_TUNNEL_PORT_UNACCESSIBLE  As String = "Тунелът е отворен, но локалния порт е недостъпен"
Private Const STR_SSH_TUNNEL_MISSING_FILE       As String = "Липсва файл %1"
Private Const STR_SSH_TUNNEL_TIMEOUT            As String = "Отваряне на SSH тунел отне прекалено много време"
Private Const STR_SSH_TUNNEL_UNKNOWN_SERVER_KEY As String = "SSH сървърът използва неоторизиран ключ"

Private m_sLastError                As String
Private m_oConfig                   As Object
Private m_oPrinters                 As Object
Private m_pTimer                    As IUnknown
Private WithEvents m_oConn          As ADODB.Connection
Attribute m_oConn.VB_VarHelpID = -1
Private m_sQueueName                As String
Private m_lQueueTimeout             As Long
Private m_oController               As cServiceController
Private m_oSshTunnel                As cExec
Private m_lSshLocalPort             As Long
Private m_dSshLastAccess            As Date

Private Enum UcsSshSettingsIndexEnum
    ucsSixTunnel
    ucsSixHost
    ucsSixPort
    ucsSixUser
    ucsSixPass
    ucsSixKeyFile
    ucsSixLocalPort
    ucsSixExtraParams
End Enum

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunction As String)
    m_sLastError = Err.Description
    Debug.Print "Critical error: " & Err.Description & " [" & MODULE_NAME & "." & sFunction & "]"
    DebugLog Err.Description & " [" & MODULE_NAME & "." & sFunction & "]", vbLogEventTypeError
End Sub

Private Sub RaiseError(sFunction As String)
    PrintError sFunction
    Err.Raise Err.Number, MODULE_NAME & "." & sFunction & vbCrLf & Err.Source, Err.Description
End Sub

'=========================================================================
' Properties
'=========================================================================

Property Get LastError() As String
    LastError = m_sLastError
End Property

Private Property Get pvAddressOfTimerProc() As cMssqlEndpoint
    Set pvAddressOfTimerProc = InitAddressOfMethod(Me, 0)
End Property

Private Property Get STR_RC4_SALT() As String
    STR_RC4_SALT = GetEnvironmentVar("_UCS_RC4_SALT")
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function Init(oConfig As Object, oPrinters As Object) As Boolean
    Const FUNC_NAME     As String = "Init"
    Dim oRequestsCache  As Object
    Dim dClock          As Date

    On Error GoTo EH
    Set m_oConfig = oConfig
    Set m_oPrinters = oPrinters
    JsonItem(oRequestsCache, vbNullString) = Empty
    '--- init service controller
    Set m_oController = New cServiceController
    If Not m_oController.Init(oPrinters, oRequestsCache) Then
        m_sLastError = m_oController.LastError
        Set m_oController = Nothing
        GoTo QH
    End If
    '--- init DB connection
    If Not pvOpenConn(oConfig) Then
        Set m_oController = Nothing
        GoTo QH
    End If
    '--- sync devices date/time if needed
    If Not IsEmpty(JsonItem(oConfig, "SyncDateTimeAdjustTolerance")) Then
        dClock = m_oConn.Execute("SELECT GETDATE()").Fields(0).Value
        pvSyncDateTime dClock, C_Lng(JsonItem(oConfig, "SyncDateTimeAdjustTolerance"))
    End If
    '--- init Service Broker queue
    m_sQueueName = STR_QUEUE_PREFIX & "/" & Zn(C_Str(JsonItem(oConfig, "QueueName")), GetErrorComputerName() & "/" & Right$(Replace(GetMacAddress(), ":", vbNullString), 6))
    m_lQueueTimeout = Znl(C_Lng(JsonItem(oConfig, "QueueTimeout")), DEF_QUEUE_TIMEOUT)
    pvQueueSetup
    pvQueueLoopAsync
    '--- success
    Init = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Public Sub Terminate()
    Set m_pTimer = Nothing
    Set m_oConn = Nothing
    Set m_oSshTunnel = Nothing
    Set m_oController = Nothing
    Set m_oConfig = Nothing
    Set m_oPrinters = Nothing
End Sub

Public Function TimerProc() As Long
Attribute TimerProc.VB_MemberFlags = "40"
    Const FUNC_NAME     As String = "TimerProc"
    
    On Error GoTo EH
    '--- if heart-beat timer is reached -> bad things must have happened to connection so re-open it
    If pvOpenConn(m_oConfig) Then
        pvQueueLoopAsync
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

'= private ===============================================================

Private Function pvHandleRequest(sSvcName As String, sRequest As String, sResponse As String, lErrorCode As Long) As Boolean
    Const FUNC_NAME     As String = "pvHandleRequest"
    Dim vJson           As Variant
    Dim sPrinterId      As String
    Dim sError          As String
    Dim sRawUrl         As String
    Dim dblStartTimerEx As Double
    Dim vSplit          As Variant
    
    On Error GoTo EH
    dblStartTimerEx = TimerEx
    sPrinterId = At(Split(sSvcName, "/"), 1)
    If LenB(sPrinterId) = 0 Then
        lErrorCode = 400
        sResponse = Printf(ERR_INVALID_SERVICE, sSvcName)
        GoTo QH
    End If
    If Left$(sRequest, 1) = "<" Then
        AssignVariant vJson, JsonFromXmlDocument(sRequest)
        If IsEmpty(vJson) Then
            lErrorCode = 400
            sResponse = ERR_INVALID_XML
            GoTo QH
        End If
    ElseIf Not JsonParse(sRequest, vJson, Error:=sError) Then
        lErrorCode = 400
        sResponse = sError
        GoTo QH
    End If
    sRawUrl = JsonItem(C_Obj(vJson), "Url")
    If Left$(sRawUrl, 1) <> "/" Then
        lErrorCode = 400
        sResponse = Printf(ERR_INVALID_URL, sRawUrl)
        GoTo QH
    End If
    If LCase$("/printers/" & sPrinterId) <> LCase$(Left$(sRawUrl, Len("/printers/" & sPrinterId))) Then
        DebugLog Printf(STR_CROSS_PRINTER_REQUEST, sRawUrl, sSvcName)
    End If
    vSplit = Split2(sRawUrl, "?")
    If Not m_oController.ServiceRequest(At(vSplit, 0), At(vSplit, 1), sRequest, sResponse) Then
        lErrorCode = m_oController.LastCode
        sResponse = m_oController.LastError
        GoTo QH
    End If
    DebugLog Printf(STR_QUEUE_ELAPSED & " [" & MODULE_NAME & "." & FUNC_NAME & "]", m_sQueueName, sPrinterId, sRawUrl, Format(TimerEx - dblStartTimerEx, "0.000"))
    '--- success
    pvHandleRequest = True
QH:
    If Not pvHandleRequest Then
        DebugLog Printf(STR_QUEUE_ERROR & " [" & MODULE_NAME & "." & FUNC_NAME & "]", m_sQueueName, sPrinterId, sResponse)
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    sResponse = m_sLastError
    GoTo QH
End Function

Private Function pvSyncDateTime(ByVal dClock As Date, ByVal lTolerance As Long) As Boolean
    Const FUNC_NAME     As String = "pvSyncDateTime"
    Const URL_DATETIME  As String = "/printers/%1/datetime"
    Dim oRequest        As Object
    Dim vElem           As Variant
    Dim sResponse       As String
    Dim vJson           As Variant
    
    On Error GoTo EH
    JsonItem(oRequest, "DeviceDateTime") = dClock
    If lTolerance > 0 Then
        JsonItem(oRequest, "AdjustTolerance") = lTolerance
    End If
    For Each vElem In JsonItem(m_oPrinters, "*/DeviceSerialNo")
        If LenB(C_Str(vElem)) <> 0 Then
            If m_oController.ServiceRequest(Printf(URL_DATETIME, vElem), vbNullString, JsonDump(oRequest, Minimize:=True), sResponse) Then
                If JsonParse(sResponse, vJson) Then
                    If JsonItem(C_Obj(vJson), "Ok") Then
                        DebugLog Printf(STR_DATETIME_SETUP, vElem, JsonItem(C_Obj(vJson), "DeviceDateTime")) & " [" & MODULE_NAME & "." & FUNC_NAME & "]"
                    Else
                        DebugLog Printf(STR_DATETIME_ERROR, vElem, JsonItem(C_Obj(vJson), "ErrorText")) & " [" & MODULE_NAME & "." & FUNC_NAME & "]", vbLogEventTypeError
                    End If
                End If
            End If
        End If
    Next
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function pvQueueSetup() As Boolean
    Const FUNC_NAME     As String = "pvQueueSetup"
    Dim lIdx            As Long
    Dim vElem           As Variant
    
    On Error GoTo EH
    m_oConn.usp_sys_ServiceBrokerSetupService m_sQueueName, STR_SVC_PREFIX & Mid$(m_sQueueName, Len(STR_QUEUE_PREFIX) + 1), "DROP_EXISTING"
    lIdx = lIdx + 1
    For Each vElem In JsonItem(m_oPrinters, "*/DeviceSerialNo")
        If LenB(C_Str(vElem)) <> 0 Then
            m_oConn.usp_sys_ServiceBrokerSetupService m_sQueueName, STR_SVC_PREFIX & "/" & C_Str(vElem), "DROP_SERVICE"
            lIdx = lIdx + 1
        End If
    Next
    DebugLog Printf(STR_QUEUE_SETUP & " [" & MODULE_NAME & "." & FUNC_NAME & "]", lIdx, m_sQueueName)
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function pvQueueLoopAsync() As Boolean
    Const FUNC_NAME     As String = "pvQueueLoopAsync"
    
    On Error GoTo EH
    m_oConn.Execute Printf(STR_QUEUE_LOOP, m_sQueueName, m_lQueueTimeout), , adAsyncExecute Or adAsyncFetch
    '--- start heart-beat timer
    Set m_pTimer = InitFireOnceTimerThunk(Me, pvAddressOfTimerProc.TimerProc, Delay:=m_lQueueTimeout + 3000)
    '--- success
    pvQueueLoopAsync = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Private Function pvOpenConn(oConfig As Object) As Boolean
    Const FUNC_NAME     As String = "pvOpenConn"
    Const DEF_SECTION   As String = "Profile"
    Dim STR_DATA_SOURCE As String
    Dim STR_FAILOVER_PARTNER As String
    Dim sIniFile        As String
    Dim sSqlServer      As String
    Dim sDbUser         As String
    Dim sDbPass         As String
    Dim sLocalConnStr   As String
    Dim oSshSettings    As Object
    Dim sSshError       As String
    Dim oTunnel         As cExec
    Dim lLocalPort      As Long
    Dim vSplit          As Variant
    Dim vSqlServer      As Variant
    Dim vMirrorServer   As Variant
    Dim bReopen         As Boolean
    
    On Error GoTo EH
    bReopen = Not m_oConn Is Nothing
    sLocalConnStr = JsonItem(oConfig, "ConnectString")
    Set oSshSettings = ParseConnectString(JsonItem(oConfig, "SshSettings"))
    sIniFile = C_Str(JsonItem(oConfig, "IniFile"))
    If LenB(sIniFile) <> 0 Then
        If LenB(sLocalConnStr) <> 0 Then
            sLocalConnStr = pvGetProfileString(sIniFile, DEF_SECTION, PRM_CONNSTR, vbNullString)
            sDbUser = pvGetProfileString(sIniFile, DEF_SECTION, PRM_DBUSER, vbNullString)
            If LenB(sDbUser) <> 0 Then
                sSqlServer = pvGetProfileString(sIniFile, DEF_SECTION, PRM_SQLSERVER, "(local)")
                sDbPass = CryptRC4(pvFromHexDump(pvGetProfileString(sIniFile, DEF_SECTION, PRM_DBPASS, vbNullString)), LCase$(sSqlServer) & STR_RC4_SALT & LCase$(sDbUser))
                sLocalConnStr = sLocalConnStr & ";Password=" & sDbPass
            End If
        End If
        If oSshSettings Is Nothing Then
            vSplit = Split(pvGetProfileString(sIniFile, DEF_SECTION, PRM_SSHSETTINGS, vbNullString), ",")
            If C_Lng(At(vSplit, ucsSixTunnel)) <> 0 Then
                JsonItem(oSshSettings, "Host") = At(vSplit, ucsSixHost)
                JsonItem(oSshSettings, "Port") = At(vSplit, ucsSixPort)
                JsonItem(oSshSettings, "User ID") = At(vSplit, ucsSixUser)
                JsonItem(oSshSettings, "Password") = CryptRC4(pvFromHexDump(At(vSplit, ucsSixPass)), LCase$(At(vSplit, ucsSixHost)) & STR_RC4_SALT & LCase$(At(vSplit, ucsSixUser)))
                JsonItem(oSshSettings, "Key File") = At(vSplit, ucsSixKeyFile)
                JsonItem(oSshSettings, "Local Port") = At(vSplit, ucsSixLocalPort)
                JsonItem(oSshSettings, "Extra Params") = At(vSplit, ucsSixExtraParams)
            End If
        End If
    End If
    If Not oSshSettings Is Nothing Then
        If Not m_oSshTunnel Is Nothing Then
            If m_oSshTunnel.KillProcess() Then
                Call Sleep(100)
            End If
            Set m_oSshTunnel = Nothing
            m_lSshLocalPort = 0
            m_dSshLastAccess = 0
        End If
        If LenB(pvGetConnProp(sLocalConnStr, "Data Source")) <> 0 Then
            STR_DATA_SOURCE = "Data Source"
            STR_FAILOVER_PARTNER = "Failover Partner"
        Else
            STR_DATA_SOURCE = "Server"
            STR_FAILOVER_PARTNER = "Failover_Partner"
        End If
        vSqlServer = Split(pvGetConnProp(sLocalConnStr, STR_DATA_SOURCE), ",")
        vMirrorServer = Split(pvGetConnProp(sLocalConnStr, STR_FAILOVER_PARTNER), ",")
        lLocalPort = C_Lng(JsonItem(oSshSettings, "Local Port"))
        If Not pvMakeSshTunnel( _
                JsonItem(oSshSettings, "Host"), _
                C_Lng(JsonItem(oSshSettings, "Port")), _
                JsonItem(oSshSettings, "User ID"), _
                JsonItem(oSshSettings, "Password"), _
                JsonItem(oSshSettings, "Key File"), _
                At(vSqlServer, 0), _
                C_Lng(At(vSqlServer, 1, 1433)), _
                At(vMirrorServer, 0), _
                C_Lng(At(vMirrorServer, 1, 1433)), _
                JsonItem(oSshSettings, "Extra Params"), _
                sSshError, oTunnel, lLocalPort) Then
            vSplit = Split(sSshError, vbCrLf)
            m_sLastError = Printf(STR_SSH_TUNNEL_ERROR, At(vSplit, 0) & IIf(UBound(vSplit) > 0, ". " & At(vSplit, UBound(vSplit)), vbNullString))
            GoTo QH
        End If
        If Not oTunnel Is Nothing Then
            sLocalConnStr = pvSetConnProp(sLocalConnStr, STR_DATA_SOURCE, "127.0.0.1," & lLocalPort)
            If LenB(pvGetConnProp(sLocalConnStr, STR_FAILOVER_PARTNER)) <> 0 Then
                sLocalConnStr = pvSetConnProp(sLocalConnStr, STR_FAILOVER_PARTNER, "127.0.0.1," & (lLocalPort + 1))
            End If
        End If
        Set m_oSshTunnel = oTunnel
        m_lSshLocalPort = lLocalPort
        m_dSshLastAccess = Now
    End If
    Set m_oConn = New ADODB.Connection
    m_oConn.CursorLocation = adUseClient
    m_oConn.Open sLocalConnStr
    DebugLog Printf(IIf(bReopen, STR_CONN_REOPEN, STR_CONN_OPEN) & " [" & MODULE_NAME & "." & FUNC_NAME & "]", C_Str(m_oConn))
    '--- start heart-beat timer
    Set m_pTimer = InitFireOnceTimerThunk(Me, pvAddressOfTimerProc.TimerProc, Delay:=m_lQueueTimeout + 3000)
    '--- success
    pvOpenConn = True
QH:
    If Not pvOpenConn Then
        Set m_pTimer = Nothing
        Set m_oConn = Nothing
        Set m_oSshTunnel = Nothing
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Private Function pvGetProfileString(IniFile As String, Section As String, Name As String, Optional DefaultValue As String) As String
    Const FUNC_NAME     As String = "pvGetProfileString"
    Dim sValue          As String
    
    On Error GoTo EH
    sValue = String$(1000, 0)
    Call GetPrivateProfileString(Section, Name, DefaultValue, sValue, Len(sValue), IniFile)
    pvGetProfileString = Left$(sValue, InStr(sValue, vbNullChar) - 1)
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Function pvFromHexDump(sText As String) As String
    Const FUNC_NAME     As String = "pvFromHexDump"
    Dim lIdx            As Long
    
    On Error GoTo EH
    For lIdx = 1 To Len(sText) Step 2
        pvFromHexDump = pvFromHexDump & Chr$(C_Lng("&H" & Mid$(sText, lIdx, 2)))
    Next
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Function pvMakeSshTunnel( _
            sSshSrv As String, _
            ByVal lSshPort As Long, _
            sUser As String, _
            sPass As String, _
            sKeyFile As String, _
            sSqlSrv As String, _
            ByVal lSqlPort As Long, _
            sMirrorSrv As String, _
            ByVal lMirrorPort As Long, _
            sExtraParams As String, _
            sError As String, _
            oTunnel As cExec, _
            lLocalPort As Long) As Boolean
    Const FUNC_NAME     As String = "pvMakeSshTunnel"
    Const DEF_TIMEOUT   As Long = 10
    Const DEF_MAXATTEMPTS As Long = 10
    Const MIN_PORT      As Long = 2000
    Const MAX_PORT      As Long = &HFFFF&
    Dim sParams         As String
    Dim sInput          As String
    Dim lAttempt        As Long
    Dim lMaxAttempts    As Long
    Dim dblTimer        As Double
    Dim lTimeout        As Double
    Dim lExtraTimeout   As Double
    Dim vElem           As Variant
    Dim sExeFile        As String
    Dim sLine           As String
    Dim bAcceptKeys     As Boolean
    
    On Error GoTo EH
    '--- terminate previous tunnel
    If Not oTunnel Is Nothing Then
        If oTunnel.KillProcess() Then
            Call Sleep(100)
        End If
        Set oTunnel = Nothing
    End If
    '--- init local vars
    lTimeout = DEF_TIMEOUT
    lMaxAttempts = DEF_MAXATTEMPTS
    bAcceptKeys = True
    sExeFile = GetEnvironmentVar("_UCS_SSH_PLINK")
    If LenB(sExeFile) = 0 Then
        sExeFile = LocateFile(PathCombine(App.Path, "plink.exe"))
    End If
    If LenB(sExeFile) = 0 Then
        sExeFile = PathCombine(App.Path, "plink.exe")
    End If
    #If DebugMode Then
        DebugPrint FUNC_NAME, MODULE_NAME, "lTimeout=" & lTimeout & ", lMaxAttempts=" & lMaxAttempts & ", bAcceptKeys=" & bAcceptKeys & ", sExeFile=" & sExeFile, 2
    #End If
    '--- check plink.exe exists
    If Not FileExists(sExeFile) Then
        sError = Printf(STR_SSH_TUNNEL_MISSING_FILE, sExeFile)
        GoTo QH
    End If
    '--- figure out free local port
    If lLocalPort <> 0 Then
        For lAttempt = 1 To lMaxAttempts
            If pvTestListen(lLocalPort) And pvTestListen(lLocalPort + 1) Then
                Exit For
            End If
            lLocalPort = lLocalPort + 1
            If lLocalPort >= MAX_PORT Then
                lLocalPort = MIN_PORT
            End If
            If lLocalPort = 8191 Or lLocalPort = 8192 Then
                lLocalPort = lLocalPort + 2
            End If
        Next
    Else
        For lAttempt = 1 To lMaxAttempts
            lLocalPort = MIN_PORT + Rnd * (MAX_PORT - MIN_PORT)
            If lLocalPort = 8191 Or lLocalPort = 8192 Then
                lLocalPort = lLocalPort + 2
            End If
            If pvTestListen(lLocalPort) And pvTestListen(lLocalPort + 1) Then
                Exit For
            End If
        Next
    End If
    If lAttempt > lMaxAttempts Then
        sError = STR_SSH_TUNNEL_NO_PORT_AVAILABLE
        GoTo QH
    End If
    '--- plink.exe options:
    '---   -ssh   use SSH protocol
    '---   -2     use version 2 of protocol
    '---   -C     compress w/ zlib
    '---   -N     no shell
    '---   -v     verbose
    '---   -P     server port
    '---   -l     (l)user
    '---   -pw    password
    '---   -i     keyfile
    '---   -L     local port tunnel [local_IP:]local_port:remote_IP:remote_port
    '---   -R     remote port tunnel [remote_IP:]remote_port:local_IP:local_port
    sParams = Printf("%1 -ssh -2 -C -N -v -l %2", ArgvQuote(sSshSrv), ArgvQuote(sUser)) & _
        IIf(LenB(sSqlSrv) <> 0, Printf(" -L %1", ArgvQuote(lLocalPort & ":" & sSqlSrv & ":" & lSqlPort)), vbNullString) & _
        IIf(LenB(sMirrorSrv) <> 0, Printf(" -L %1", ArgvQuote((lLocalPort + 1) & ":" & sMirrorSrv & ":" & lMirrorPort)), vbNullString) & _
        IIf(LenB(sKeyFile) <> 0, Printf(" -i %1", ArgvQuote(sKeyFile)), vbNullString) & _
        IIf(LenB(sPass) <> 0, Printf(" -pw %1", ArgvQuote(sPass)), vbNullString) & _
        IIf(lSshPort <> 0, Printf(" -P %1", lSshPort), vbNullString) & _
        IIf(LenB(sExtraParams) <> 0, " " & sExtraParams, vbNullString)
    #If DebugMode Then
        DebugPrint FUNC_NAME, MODULE_NAME, "sParams=" & sParams, 2
    #End If
    '--- shell plink.exe
    Set oTunnel = New cExec
    With oTunnel
        .Run sExeFile, sParams, StartHidden:=True, LimitFlags:=JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
        dblTimer = TimerEx
        Do While Not .AtEndOfError Or Not .AtEndOfOutput
            If TimerEx - dblTimer > lTimeout + lExtraTimeout Then
                sError = STR_SSH_TUNNEL_TIMEOUT & vbCrLf & sInput
                If .KillProcess() Then
                    Call Sleep(100)
                End If
                GoTo QH
            End If
            Call Sleep(1)
            sLine = .ReadPendingOutput & .ReadPendingError
            If LenB(sLine) Then
                #If DebugMode Then
                    DebugPrint FUNC_NAME, MODULE_NAME, "plink.exe output: " & sLine, 2
                #End If
                sInput = sInput & Format$(TimerEx - dblTimer, "0.000") & ": " & Replace(Left$(sLine, Len(sLine) - 1), vbCrLf, vbCrLf & Format$(TimerEx - dblTimer, "0.000") & ": ") & Right$(sLine, 1)
            End If
            If InStr(sInput, "Server version") > 0 And lExtraTimeout = 0 Then
                lExtraTimeout = lTimeout
                #If DebugMode Then
                    DebugPrint FUNC_NAME, MODULE_NAME, "lExtraTimeout=" & lExtraTimeout, 2
                #End If
            End If
            If InStr(sInput, "y/n") > 0 Then
                '--- find ssh server rsa key
                For Each vElem In Split(sInput, vbCrLf)
                    If InStr(vElem, "ssh-rsa ") > 0 Then
                        sError = vElem & vbCrLf
                        #If DebugMode Then
                            DebugPrint FUNC_NAME, MODULE_NAME, "SSH server key: " & vElem, 2
                        #End If
                    End If
                Next
                If bAcceptKeys Then
                    .WriteInput "y" & vbCrLf
                    sInput = Replace(sInput, "y/n", vbNullString)
                Else
                    sError = STR_SSH_TUNNEL_UNKNOWN_SERVER_KEY & vbCrLf & sError
                    GoTo QH
                End If
            ElseIf InStr(sInput, "Local port " & lLocalPort & " forwarding") > 0 Then
                If Not pvTestListen(lLocalPort) Then
                    pvMakeSshTunnel = True
                Else
                    sError = STR_SSH_TUNNEL_PORT_UNACCESSIBLE
                End If
                GoTo QH
            End If
        Loop
        sError = sInput
        #If DebugMode Then
            DebugPrint FUNC_NAME, MODULE_NAME, "GetExitCode=" & .GetExitCode, 2
        #End If
    End With
QH:
    '--- trim ending new lines
    Do While Right$(sError, 2) = vbCrLf
        sError = Left$(sError, Len(sError) - 2)
    Loop
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Function pvTestListen(ByVal lPort As Long) As Boolean
    Const FUNC_NAME     As String = "pvTestListen"
    Dim hSocket         As Long
    Dim uAddr           As SOCKADDR_IN
    Dim baWSAData()     As Byte
    
    On Error GoTo EH
    ReDim baWSAData(0 To 1000) As Byte
    If WSAStartup(&H202, baWSAData(0)) = 0 Then
        hSocket = SocketInit(AF_INET, SOCK_STREAM, IPPROTO_TCP)
        If hSocket <> INVALID_SOCKET Then
            With uAddr
                .Addr = SocketGetIP("127.0.0.1")
                .Family = AF_INET
                .Port = SocketnTons(lPort And &HFFFF&)
            End With
            If SocketBind(hSocket, uAddr, Len(uAddr)) <> -1 Then
                If SocketListen(hSocket, 5) <> -1 Then
                    pvTestListen = True
                End If
            End If
            Call SocketClose(hSocket)
        End If
        Call WSACleanup
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Function pvGetConnProp(ByVal sConnStr As String, sProp As String) As String
    Const FUNC_NAME     As String = "pvGetConnProp"
    Dim vElem           As Variant
    
    On Error GoTo EH
    For Each vElem In Split(sConnStr, ";")
        If LCase$(Left$(vElem, Len(sProp) + 1)) = LCase$(sProp & "=") Then
            pvGetConnProp = Mid$(vElem, InStr(vElem, "=") + 1)
            Exit Function
        End If
    Next
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Function pvSetConnProp(ByVal sConnStr As String, sProp As String, sValue As String) As String
    Const FUNC_NAME     As String = "pvSetConnProp"
    Dim vSplit          As Variant
    Dim lIdx            As Long
    
    On Error GoTo EH
    vSplit = Split(sConnStr, ";")
    For lIdx = 0 To UBound(vSplit)
        If LCase$(Left$(vSplit(lIdx), Len(sProp) + 1)) = LCase$(sProp & "=") Then
            If LenB(sValue) <> 0 Then
                vSplit(lIdx) = sProp & "=" & sValue
            Else
                vSplit(lIdx) = vbNullString
            End If
            Exit For
        End If
    Next
    If lIdx > UBound(vSplit) Then
        pvSetConnProp = sConnStr
        If LenB(sValue) <> 0 Then
            pvSetConnProp = pvSetConnProp & IIf(LenB(pvSetConnProp) <> 0 And Right$(pvSetConnProp, 1) <> ";", ";", vbNullString) & sProp & "=" & sValue
        End If
    Else
        For lIdx = 0 To UBound(vSplit)
            If LenB(vSplit(lIdx)) <> 0 Then
                pvSetConnProp = pvSetConnProp & IIf(LenB(pvSetConnProp) <> 0, ";", vbNullString) & vSplit(lIdx)
            End If
        Next
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

'=========================================================================
' MST
'=========================================================================

Private Function InitAddressOfMethod(pObj As Object, ByVal MethodParamCount As Long) As Object
    Const STR_THUNK     As String = "6AAAAABag+oFV4v6ge9QEMEAgcekEcEAuP9EJAS5+QcAAPOri8LB4AgFuQAAAKuLwsHoGAUAjYEAq7gIAAArq7hEJASLq7hJCIsEq7iBi1Qkq4tEJAzB4AIFCIkCM6uLRCQMweASBcDCCACriTrHQgQBAAAAi0QkCIsAiUIIi0QkEIlCDIHqUBDBAIvCBTwRwQCri8IFUBHBAKuLwgVgEcEAq4vCBYQRwQCri8IFjBHBAKuLwgWUEcEAq4vCBZwRwQCri8IFpBHBALn5BwAAq4PABOL6i8dfgcJQEMEAi0wkEIkRK8LCEAAPHwCLVCQE/0IEi0QkDIkQM8DCDABmkItUJAT/QgSLQgTCBAAPHwCLVCQE/0oEi0IEg/gAfgPCBABZWotCDGgAgAAAagBSUf/gZpC4AUAAgMIIALgBQACAwhAAuAFAAIDCGAC4AUAAgMIkAA==" ' 25.3.2019 14:01:08
    Const THUNK_SIZE    As Long = 16728
    Dim hThunk          As Long
    Dim lSize           As Long
    
    hThunk = VirtualAlloc(0, THUNK_SIZE, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
    If hThunk = 0 Then
        Exit Function
    End If
    Call CryptStringToBinary(STR_THUNK, Len(STR_THUNK), CRYPT_STRING_BASE64, hThunk, THUNK_SIZE)
    lSize = CallWindowProc(hThunk, ObjPtr(pObj), MethodParamCount, GetProcAddress(GetModuleHandle("kernel32"), "VirtualFree"), VarPtr(InitAddressOfMethod))
    Debug.Assert lSize = THUNK_SIZE
End Function

Private Function InitFireOnceTimerThunk(pObj As Object, ByVal pfnCallback As Long, Optional Delay As Long) As IUnknown
    Const STR_THUNK     As String = "6AAAAABag+oFgeogERkAV1aLdCQUg8YIgz4AdCqL+oHHBBMZAIvCBSgSGQCri8IFZBIZAKuLwgV0EhkAqzPAq7kIAAAA86WBwgQTGQBSahj/UhBai/iLwqu4AQAAAKszwKuri3QkFKWlg+8Yi0IMSCX/AAAAUItKDDsMJHULWIsPV/9RFDP/62P/QgyBYgz/AAAAjQTKjQTIjUyIMIB5EwB101jHAf80JLiJeQTHQQiJRCQEi8ItBBMZAAWgEhkAUMHgCAW4AAAAiUEMWMHoGAUA/+CQiUEQiU8MUf90JBRqAGoAiw//URiJRwiLRCQYiTheX7g0ExkALSARGQAFABQAAMIQAGaQi0QkCIM4AHUqg3gEAHUkgXgIwAAAAHUbgXgMAAAARnUSi1QkBP9CBItEJAyJEDPAwgwAuAJAAIDCDACQi1QkBP9CBItCBMIEAA8fAItUJAT/SgSLQgR1HYtCDMZAEwCLCv9yCGoA/1Eci1QkBIsKUv9RFDPAwgQAi1QkBIsKi0EohcB0J1L/0FqD+AF3SYsKUv9RLFqFwHU+iwpSavD/cSD/USRaqQAAAAh1K4sKUv9yCGoA/1EcWv9CBDPAUFT/chD/UhSLVCQIx0IIAAAAAFLodv///1jCFABmkA==" ' 27.3.2019 9:14:57
    Const THUNK_SIZE    As Long = 5652
    Static hThunk       As Long
    Dim aParams(0 To 9) As Long
    Dim lSize           As Long
    
    aParams(0) = ObjPtr(pObj)
    aParams(1) = pfnCallback
    #If ImplSelfContained Then
        If hThunk = 0 Then
            hThunk = pvThunkGlobalData("InitFireOnceTimerThunk")
        End If
    #End If
    If hThunk = 0 Then
        hThunk = VirtualAlloc(0, THUNK_SIZE, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
        If hThunk = 0 Then
            Exit Function
        End If
        Call CryptStringToBinary(STR_THUNK, Len(STR_THUNK), CRYPT_STRING_BASE64, hThunk, THUNK_SIZE)
        aParams(2) = GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemAlloc")
        aParams(3) = GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemFree")
        aParams(4) = GetProcAddress(GetModuleHandle("user32"), "SetTimer")
        aParams(5) = GetProcAddress(GetModuleHandle("user32"), "KillTimer")
        '--- for IDE protection
        Debug.Assert pvGetIdeOwner(aParams(6))
        If aParams(6) <> 0 Then
            aParams(7) = GetProcAddress(GetModuleHandle("user32"), "GetWindowLongA")
            aParams(8) = GetProcAddress(GetModuleHandle("vba6"), "EbMode")
            aParams(9) = GetProcAddress(GetModuleHandle("vba6"), "EbIsResetting")
        End If
        #If ImplSelfContained Then
            pvThunkGlobalData("InitFireOnceTimerThunk") = hThunk
        #End If
    End If
    lSize = CallWindowProc(hThunk, 0, Delay, VarPtr(aParams(0)), VarPtr(InitFireOnceTimerThunk))
    Debug.Assert lSize = THUNK_SIZE
End Function

Private Function pvGetIdeOwner(hIdeOwner As Long) As Boolean
    #If Not ImplNoIdeProtection Then
        Dim lProcessId      As Long
        
        Do
            hIdeOwner = FindWindowEx(0, hIdeOwner, "IDEOwner", vbNullString)
            Call GetWindowThreadProcessId(hIdeOwner, lProcessId)
        Loop While hIdeOwner <> 0 And lProcessId <> GetCurrentProcessId()
    #End If
    pvGetIdeOwner = True
End Function

#If ImplSelfContained Then
Private Property Get pvThunkGlobalData(sKey As String) As Long
    Dim sBuffer     As String
    
    sBuffer = String$(50, 0)
    Call GetEnvironmentVariable("_MST_GLOBAL" & GetCurrentProcessId() & "_" & sKey, sBuffer, Len(sBuffer) - 1)
    pvThunkGlobalData = Val(Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1))
End Property

Private Property Let pvThunkGlobalData(sKey As String, ByVal lValue As Long)
    Call SetEnvironmentVariable("_MST_GLOBAL" & GetCurrentProcessId() & "_" & sKey, lValue)
End Property
#End If

'=========================================================================
' Events
'=========================================================================

Private Sub m_oConn_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
    Const FUNC_NAME     As String = "m_oConn_ExecuteComplete"
    Dim sHandle         As String
    Dim sResponse       As String
    Dim lErrorCode      As Long
    Dim sPrinterId      As String
    Dim vErr            As Variant
    
    On Error GoTo EH
    If Not pRecordset Is Nothing Then
        If (pRecordset.State And adStateOpen) <> 0 Then
            If pRecordset.Fields.Count = LNG_QUEUE_FIELDS Then
                pvQueueLoopAsync
                sHandle = C_Str(pRecordset!Handle.Value)
                If C_Lng(pRecordset!Result.Value) = 0 And LenB(sHandle) <> 0 Then
                    If pvHandleRequest(C_Str(pRecordset!SvcName.Value), C_Str(pRecordset!Request.Value), sResponse, lErrorCode) Then
                        m_oConn.Execute Printf(STR_SEND_ON_CONV, Quote(sHandle), Quote(sResponse))
                    Else
                        m_oConn.Execute Printf(STR_END_CONV_WITH_ERROR, Quote(sHandle), Znl(lErrorCode, 500), _
                            Quote(Zn(sResponse, Printf(ERR_UNKNOWN, lErrorCode))))
                    End If
                ElseIf C_Lng(pRecordset!Result.Value) <> 99 Then
                    sResponse = C_Str(pRecordset!ErrorText.Value)
                    If LenB(sResponse) <> 0 Then
                        sPrinterId = At(Split(C_Str(pRecordset!SvcName.Value), "/"), 1)
                        DebugLog Printf(STR_QUEUE_ERROR & " [" & MODULE_NAME & "." & FUNC_NAME & "]", m_sQueueName, sPrinterId, sResponse)
                    End If
                End If
            End If
        End If
    End If
    Exit Sub
EH:
    vErr = Array(Err.Number, Err.Source, Err.Description)
    PrintError FUNC_NAME
    Resume NextLine
NextLine:
    On Error Resume Next
    If vErr(0) = ERR_DB_E_NOTABLE Then
        pvQueueSetup
    End If
    If LenB(sHandle) <> 0 Then
        m_oConn.Execute Printf(STR_END_CONV_WITH_ERROR, Quote(sHandle), Znl(lErrorCode, 500), _
            Quote(Zn(m_sLastError, Printf(ERR_UNKNOWN, lErrorCode))))
    End If
    pvQueueLoopAsync
End Sub

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Terminate()
    Terminate
End Sub

'=========================================================================
' IEndpoint interface
'=========================================================================

Private Sub IEndpoint_Terminate()
    Terminate
End Sub
