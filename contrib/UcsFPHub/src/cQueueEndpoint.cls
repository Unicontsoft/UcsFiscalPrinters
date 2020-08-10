VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cQueueEndpoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
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
Private Const MODULE_NAME As String = "cQueueEndpoint"
Implements IEndpoint

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

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function WSAStartup Lib "ws2_32" (ByVal wVersionRequired As Long, lpWSAData As Any) As Long
Private Declare Function WSACleanup Lib "ws2_32" () As Long
Private Declare Function SocketInit Lib "ws2_32" Alias "socket" (ByVal lAf As Long, ByVal lSocketType As Long, ByVal lProtocol As Long) As Long
Private Declare Function SocketBind Lib "ws2_32" Alias "bind" (ByVal lSocket As Long, ByRef uName As SOCKADDR_IN, ByRef lNameLen As Long) As Long
Private Declare Function SocketListen Lib "ws2_32" Alias "listen" (ByVal lSocket As Long, ByVal lBackLog As Long) As Long
Private Declare Function SocketClose Lib "ws2_32" Alias "closesocket" (ByVal lSocket As Long) As Long
Private Declare Function SocketGetIP Lib "ws2_32" Alias "inet_addr" (ByVal sConnPoint As String) As Long
Private Declare Function SocketnTons Lib "ws2_32" Alias "htons" (ByVal lNetShort As Long) As Integer

Private Type SOCKADDR_IN
    Family              As Integer
    Port                As Integer
    Addr                As Long
    Zero(7)             As Byte
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_CONN_OPEN             As String = "Връзка с %2 на %1 е отворена"
Private Const STR_CONN_REOPEN           As String = "Връзка с %2 на %1 е възстановена"
Private Const STR_HANDLER_ELAPSED       As String = "Принтер %1 изпълни заявка %2 за %3 сек."
Private Const STR_HANDLER_ERROR         As String = "Принтер %1 върна грешка %2"
Private Const STR_CROSS_PRINTER_REQUEST As String = "Заявка %1 не е към принтер на услуга %2"
Private Const STR_DATETIME_SETUP        As String = "Принтер %1 е с установен час %2 (предишен час %3)"
Private Const STR_DATETIME_ERROR        As String = "Принтер %1 върна грешка %2"
Private Const STR_QUEUE_PREFIX          As String = "UcsFpTargetQueue"
'--- errors
Private Const ERR_INVALID_SERVICE       As String = "Невалидна Service Broker услуга: %1"
Private Const ERR_INVALID_URL           As String = "Невалиден адрес: %1"
Private Const ERR_INVALID_XML           As String = "Невалиден XML"
'--- defaults
Private Const DEF_MSSQL_TIMEOUT         As Long = 15000
Private Const DEF_MYSQL_TIMEOUT         As Long = 5000
Private Const DEF_ASYNC_COMMAND_TIMEOUT As Long = 12& * 3600 '--- 12h
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
Private Const STR_SSH_TUNNEL_MISSING_PASSPHRASE As String = "Клиентският ключ изисква парола"

Private m_sLastError                As String
Private m_oConfig                   As Object
Private m_oPrinters                 As Object
Private m_pTimer                    As IUnknown
Private m_oConn                     As ADODB.Connection
Attribute m_oConn.VB_VarHelpID = -1
Private m_oRecvConn                 As ADODB.Connection
Private WithEvents m_oController    As cServiceController
Attribute m_oController.VB_VarHelpID = -1
Private m_oSshTunnel                As cExec
Private m_lSshLocalPort             As Long
Private m_dSshLastAccess            As Date
Private m_oQueue                    As IQueue
Private m_sSqlInstance              As String
Private m_sSqlDatabase              As String

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
    #If USE_DEBUG_LOG <> 0 Then
        DebugLog MODULE_NAME, sFunction & "(" & m_sSqlInstance & "/" & m_sSqlDatabase & ")(" & Erl & ")", Err.Description & " &H" & Hex$(Err.Number), vbLogEventTypeError
    #Else
        Debug.Print "Critical error: " & Err.Description & " [" & MODULE_NAME & "." & sFunction & "]"
    #End If
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
    Dim sQueueName      As String
    Dim lQueueTimeout   As Long

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
        dClock = m_oConn.Execute("SELECT GETDATE()", Options:=adCmdText).Fields(0).Value
        pvSyncDateTime dClock, C_Lng(JsonItem(oConfig, "SyncDateTimeAdjustTolerance"))
    End If
    '--- init Service Broker queue
    sQueueName = GetErrorComputerName() & "/" & Right$(Replace(GetMacAddress(), ":", vbNullString), 6) & "/" & m_sSqlInstance & "/" & m_sSqlDatabase
    sQueueName = STR_QUEUE_PREFIX & "/" & Zn(C_Str(JsonItem(oConfig, "QueueName")), sQueueName)
    Select Case LCase$(JsonItem(oConfig, "Binding"))
    Case "mssqlservicebroker"
        lQueueTimeout = Znl(C_Lng(JsonItem(oConfig, "QueueTimeout")), DEF_MSSQL_TIMEOUT)
        If pvInitMssqlQueue(sQueueName, lQueueTimeout, RetVal:=m_oQueue) Is Nothing Then
            GoTo QH
        End If
    Case "mysqlmessagequeue"
        lQueueTimeout = Znl(C_Lng(JsonItem(oConfig, "QueueTimeout")), DEF_MYSQL_TIMEOUT)
        If pvInitMysqlQueue(sQueueName, lQueueTimeout, RetVal:=m_oQueue) Is Nothing Then
            GoTo QH
        End If
    End Select
    '--- success
    Init = True
QH:
    If LenB(m_sLastError) <> 0 Then
        DebugLog MODULE_NAME, FUNC_NAME, m_sLastError, vbLogEventTypeError
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Private Function pvInitMssqlQueue(sQueueName As String, ByVal lQueueTimeout As Long, Optional RetVal As cMssqlQueue) As cMssqlQueue
    Set RetVal = New cMssqlQueue
    If RetVal.Init(m_oConn, m_oRecvConn, JsonItem(m_oPrinters, "*/DeviceSerialNo"), sQueueName, lQueueTimeout, Me) Then
        m_sLastError = RetVal.LastError
        Set pvInitMssqlQueue = RetVal
    End If
End Function

Private Function pvInitMysqlQueue(sQueueName As String, ByVal lQueueTimeout As Long, Optional RetVal As cMysqlQueue) As cMysqlQueue
    Set RetVal = New cMysqlQueue
    If RetVal.Init(m_oConn, m_oRecvConn, JsonItem(m_oPrinters, "*/DeviceSerialNo"), sQueueName, lQueueTimeout, Me) Then
        m_sLastError = RetVal.LastError
        Set pvInitMysqlQueue = RetVal
    End If
End Function

Public Sub Terminate()
    If Not m_oQueue Is Nothing Then
        m_oQueue.Terminate
        Set m_oQueue = Nothing
    End If
    Set m_pTimer = Nothing
    Set m_oConn = Nothing
    Set m_oRecvConn = Nothing
    Set m_oSshTunnel = Nothing
    Set m_oController = Nothing
    Set m_oConfig = Nothing
    Set m_oPrinters = Nothing
End Sub

'= friend ================================================================

Friend Function frHandleRequest(sSvcName As String, sRequest As String, sResponse As String, lErrorCode As Long) As Boolean
    Const FUNC_NAME     As String = "frHandleRequest"
    Dim vJson           As Variant
    Dim sPrinterID      As String
    Dim sError          As String
    Dim sRawUrl         As String
    Dim dblStartTimerEx As Double
    Dim vSplit          As Variant
    
    On Error GoTo EH
    dblStartTimerEx = TimerEx
    If IsLogDebugEnabled Then
        DebugLog MODULE_NAME, FUNC_NAME, "sRequest=" & Replace(sRequest, vbCrLf, "^p") & ", sSvcName=" & sSvcName, vbLogEventTypeDebug
    End If
    vSplit = Split(sSvcName, "/")
    If UBound(vSplit) = 1 Then
        If Not IsOnlyDigits(At(vSplit, 1)) Then
            sPrinterID = At(vSplit, 1)
        End If
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
    If LenB(sPrinterID) <> 0 And LCase$("/printers/" & sPrinterID) <> LCase$(Left$(sRawUrl, Len("/printers/" & sPrinterID))) Then
        DebugLog MODULE_NAME, FUNC_NAME, Printf(STR_CROSS_PRINTER_REQUEST, sRawUrl, sSvcName)
    End If
    vSplit = Split(sRawUrl, "?", Limit:=2)
    If Not m_oController.ServiceRequest(At(vSplit, 0), At(vSplit, 1), sRequest, sResponse) Then
        lErrorCode = m_oController.LastCode
        sResponse = m_oController.LastError
        GoTo QH
    End If
    DebugLog MODULE_NAME, FUNC_NAME, Printf(STR_HANDLER_ELAPSED, Zn(sPrinterID, "N/A"), sRawUrl, Format(TimerEx - dblStartTimerEx, FORMAT_BASE_3))
    '--- success
    frHandleRequest = True
QH:
    If Not frHandleRequest Then
        DebugLog MODULE_NAME, FUNC_NAME, Printf(STR_HANDLER_ERROR, sPrinterID, sResponse)
    End If
    If IsLogDebugEnabled Then
        DebugLog MODULE_NAME, FUNC_NAME, "sResponse=" & Replace(sResponse, vbCrLf, "^p") & IIf(lErrorCode <> 0, ", lErrorCode=" & lErrorCode, vbNullString), vbLogEventTypeDebug
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    sResponse = m_sLastError
    GoTo QH
End Function

Friend Function frReopenConn(oConn As Object, Optional oRecvConn As Object) As Boolean
    If pvOpenConn(m_oConfig) Then
        Set oConn = m_oConn
        Set oRecvConn = m_oRecvConn
        '--- success
        frReopenConn = True
    End If
End Function

'= private ===============================================================

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
                        If Not IsEmpty(JsonItem(C_Obj(vJson), "PreviousDateTime")) Then
                            DebugLog MODULE_NAME, FUNC_NAME, Printf(STR_DATETIME_SETUP, vElem, JsonItem(C_Obj(vJson), "DeviceDateTime"), JsonItem(C_Obj(vJson), "PreviousDateTime"))
                        End If
                    Else
                        DebugLog MODULE_NAME, FUNC_NAME, Printf(STR_DATETIME_ERROR, vElem, JsonItem(C_Obj(vJson), "ErrorText")), vbLogEventTypeError
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
    Dim sAppName        As String
    
    On Error GoTo EH
    bReopen = Not m_oConn Is Nothing
    sLocalConnStr = JsonItem(oConfig, "ConnectString")
    Set oSshSettings = ParseConnectString(JsonItem(oConfig, "SshSettings"))
    sDbUser = C_Str(JsonItem(oConfig, "User"))
    If LenB(sDbUser) <> 0 Then
        sDbPass = C_Str(JsonItem(oConfig, "Password"))
        If LenB(pvFromHexDump(sDbPass)) <> 0 Then
            sSqlServer = Zn(C_Str(JsonItem(oConfig, "Server")), "(local)")
            sDbPass = CryptRC4(pvFromHexDump(sDbPass), LCase$(sSqlServer) & STR_RC4_SALT & LCase$(sDbUser))
        End If
        sLocalConnStr = pvSetConnProp(sLocalConnStr, "User ID", sDbUser)
        sLocalConnStr = pvSetConnProp(sLocalConnStr, "Password", sDbPass)
    End If
    sIniFile = C_Str(JsonItem(oConfig, "IniFile"))
    If LenB(sIniFile) <> 0 Then
        If LenB(sLocalConnStr) = 0 Then
            sLocalConnStr = pvGetProfileString(sIniFile, DEF_SECTION, PRM_CONNSTR, vbNullString)
            sDbUser = pvGetProfileString(sIniFile, DEF_SECTION, PRM_DBUSER, vbNullString)
            If LenB(sDbUser) <> 0 Then
                sDbPass = pvGetProfileString(sIniFile, DEF_SECTION, PRM_DBPASS, vbNullString)
                If LenB(pvFromHexDump(sDbPass)) <> 0 Then
                    sSqlServer = pvGetProfileString(sIniFile, DEF_SECTION, PRM_SQLSERVER, "(local)")
                    sDbPass = CryptRC4(pvFromHexDump(sDbPass), LCase$(sSqlServer) & STR_RC4_SALT & LCase$(sDbUser))
                End If
                sLocalConnStr = pvSetConnProp(sLocalConnStr, "Password", sDbPass)
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
        If LenB(pvGetConnProp(sLocalConnStr, "Data Source")) <> 0 Then
            sLocalConnStr = pvSetConnProp(sLocalConnStr, "Application Name", "{APP_NAME}")
        Else
            sLocalConnStr = pvSetConnProp(sLocalConnStr, "APP", "{APP_NAME}")
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
        m_dSshLastAccess = GetCurrentNow
    End If
    If InStr(sLocalConnStr, "{APP_NAME}") > 0 Then
        sAppName = App.Title & " " & STR_VERSION & _
            " [" & GetCurrentProcessUser(False) & "@" & LCase$(GetErrorComputerName()) & "] " & _
            Format$(GetErrorProcessCreationTime(), FORMAT_DATETIME_LOG)
        sLocalConnStr = Replace(sLocalConnStr, "{APP_NAME}", sAppName)
    End If
    If IsLogDataDumpEnabled Then '-- prevent password dump
        DebugLog MODULE_NAME, FUNC_NAME, "sLocalConnStr=" & sLocalConnStr, vbLogEventTypeDebug
    End If
    Set m_oConn = New ADODB.Connection
    m_oConn.CursorLocation = adUseClient
    m_oConn.Open sLocalConnStr
    If Not IsEmpty(JsonItem(oConfig, "CommandTimeout")) Then
        m_oConn.CommandTimeout = C_Lng(JsonItem(oConfig, "CommandTimeout"))
    Else
        m_oConn.CommandTimeout = m_oConn.ConnectionTimeout * 2
    End If
    Select Case LCase$(JsonItem(oConfig, "Binding"))
    Case "mssqlservicebroker"
        With m_oConn.Execute("SELECT @@SERVERNAME, DB_NAME()", Options:=adCmdText)
            m_sSqlInstance = .Fields.Item(0).Value
            m_sSqlDatabase = .Fields.Item(1).Value
        End With
    Case "mysqlmessagequeue"
        With m_oConn.Execute("SELECT @@HOSTNAME, DATABASE();", Options:=adCmdText)
            m_sSqlInstance = .Fields.Item(0).Value
            m_sSqlDatabase = .Fields.Item(1).Value
        End With
    End Select
    Set m_oRecvConn = New ADODB.Connection
    m_oRecvConn.CursorLocation = adUseClient
    m_oRecvConn.Open sLocalConnStr
    If Not IsEmpty(JsonItem(oConfig, "AsyncCommandTimeout")) Then
        m_oRecvConn.CommandTimeout = C_Lng(JsonItem(oConfig, "AsyncCommandTimeout"))
    Else
        m_oRecvConn.CommandTimeout = DEF_ASYNC_COMMAND_TIMEOUT
    End If
    DebugLog MODULE_NAME, FUNC_NAME, Printf(IIf(bReopen, STR_CONN_REOPEN, STR_CONN_OPEN), m_sSqlInstance, m_sSqlDatabase)
    '--- success
    pvOpenConn = True
QH:
    If Not pvOpenConn Then
        Set m_oConn = Nothing
        Set m_oRecvConn = Nothing
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
    Dim lIdx            As Long
    
    On Error GoTo QH
    For lIdx = 1 To Len(sText) Step 2
        If Not IsNumeric("&H" & Mid$(sText, lIdx, 2)) Then
            GoTo QH
        End If
        pvFromHexDump = pvFromHexDump & Chr$(CLng("&H" & Mid$(sText, lIdx, 2)))
    Next
    Exit Function
QH:
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
    Dim dblTimerEx      As Double
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
    '--- check plink.exe exists
    If FileExists(sExeFile) Then
        sExeFile = ArgvQuote(sExeFile)
    ElseIf LenB(sExeFile) = 0 Then
        sError = Printf(STR_SSH_TUNNEL_MISSING_FILE, "plink.exe")
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
            If lLocalPort = DEF_LISTEN_PORT - 1 Or lLocalPort = DEF_LISTEN_PORT Then
                lLocalPort = lLocalPort + 2
            End If
        Next
    Else
        For lAttempt = 1 To lMaxAttempts
            lLocalPort = MIN_PORT + Rnd * (MAX_PORT - MIN_PORT)
            If lLocalPort = DEF_LISTEN_PORT - 1 Or lLocalPort = DEF_LISTEN_PORT Then
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
    sParams = Printf("%1 -C -N -v -l %2", ArgvQuote(sSshSrv), ArgvQuote(sUser)) & _
        IIf(LenB(sSqlSrv) <> 0, Printf(" -L %1", ArgvQuote(lLocalPort & ":" & sSqlSrv & ":" & lSqlPort)), vbNullString) & _
        IIf(LenB(sMirrorSrv) <> 0, Printf(" -L %1", ArgvQuote((lLocalPort + 1) & ":" & sMirrorSrv & ":" & lMirrorPort)), vbNullString) & _
        IIf(LenB(sKeyFile) <> 0, Printf(" -i %1", ArgvQuote(sKeyFile)), vbNullString) & _
        IIf(LenB(sPass) <> 0, Printf(" -pw %1", ArgvQuote(sPass)), vbNullString) & _
        IIf(lSshPort <> 0, Printf(" -P %1", lSshPort), vbNullString) & _
        IIf(LenB(sExtraParams) <> 0, " " & sExtraParams, vbNullString)
    If IsLogDataDumpEnabled Then '-- prevent password dump
        DebugLog MODULE_NAME, FUNC_NAME, "sParams=" & sParams, vbLogEventTypeDebug
    End If
    '--- shell plink.exe
    Set oTunnel = New cExec
    With oTunnel
        .Run vbNullString, sExeFile & " " & sParams, StartHidden:=True, LimitFlags:=JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
        dblTimerEx = TimerEx
        Do While Not .AtEndOfError Or Not .AtEndOfOutput
            If TimerEx > dblTimerEx + lTimeout + lExtraTimeout Then
                sError = STR_SSH_TUNNEL_TIMEOUT & vbCrLf & sInput
                If .KillProcess() Then
                    Call Sleep(100)
                End If
                GoTo QH
            End If
            Call Sleep(1)
            sLine = .ReadPendingOutput & .ReadPendingError
            If Len(sLine) > 1 Then
                If IsLogDebugEnabled Then
                    DebugLog MODULE_NAME, FUNC_NAME, Left$(sLine, Len(sLine) - 2), vbLogEventTypeDebug
                End If
                sInput = sInput & Format$(TimerEx - dblTimerEx, FORMAT_BASE_3) & ": " & Replace(Left$(sLine, Len(sLine) - 1), vbCrLf, vbCrLf & Format$(TimerEx - dblTimerEx, FORMAT_BASE_3) & ": ") & Right$(sLine, 1)
            End If
            If InStr(sInput, "Server version") > 0 And lExtraTimeout = 0 Then
                lExtraTimeout = lTimeout
            End If
            If InStr(sInput, "y/n") > 0 Then
                '--- find ssh server rsa key
                For Each vElem In Split(sInput, vbCrLf)
                    If InStr(vElem, "ssh-rsa ") > 0 Then
                        sError = vElem & vbCrLf
                        If IsLogDebugEnabled Then
                            DebugLog MODULE_NAME, FUNC_NAME, "SSH server key: " & vElem, vbLogEventTypeDebug
                        End If
                    End If
                Next
                If bAcceptKeys Then
                    .WriteInput "y" & vbCrLf
                    sInput = Replace(sInput, "y/n", "{y}/n")
                Else
                    sError = STR_SSH_TUNNEL_UNKNOWN_SERVER_KEY & vbCrLf & sError
                    GoTo QH
                End If
            ElseIf InStr(sInput, "Press Return to begin session") > 0 Then
                .WriteInput vbCrLf
                sInput = Replace(sInput, "Press Return", "{P}ress Return")
            ElseIf InStr(sInput, "Passphrase for key") > 0 Then
                If LenB(sPass) <> 0 Then
                    .WriteInput sPass & vbCrLf
                    sInput = Replace(sInput, "Passphrase for", "{P}assphrase for")
                Else
                    sError = STR_SSH_TUNNEL_MISSING_PASSPHRASE
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
        If IsLogDebugEnabled Then
            DebugLog MODULE_NAME, FUNC_NAME, "GetExitCode=" & .GetExitCode, vbLogEventTypeDebug
        End If
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
    pvGetConnProp = C_Str(JsonItem(ParseConnectString(sConnStr), sProp))
End Function

Private Function pvSetConnProp(ByVal sConnStr As String, sProp As String, sValue As String) As String
    Dim oMap            As Object
    
    Set oMap = ParseConnectString(sConnStr)
    JsonItem(oMap, sProp) = Zn(sValue, Empty)
    pvSetConnProp = ToConnectString(oMap)
End Function

'=========================================================================
' Events
'=========================================================================

Private Sub m_oController_BeforeCommitReceipt(oCtx As Object, Cancel As Boolean)
    Const FUNC_NAME     As String = "m_oController_BeforeCommitReceipt"
    
    On Error GoTo EH
    If Not m_oQueue.IsConversationAlive() Then
        Cancel = True
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
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
