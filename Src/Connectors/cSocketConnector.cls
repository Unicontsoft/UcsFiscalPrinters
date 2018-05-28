VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSocketConnector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================================================
' $Header: /UcsFiscalPrinter/Src/Connectors/cSocketConnector.cls 3     28.05.18 16:37 Wqw $
'
'   Unicontsoft Fiscal Printers Project
'   Copyright (c) 2008-2018 Unicontsoft
'
'   Network socket connector
'
' $Log: /UcsFiscalPrinter/Src/Connectors/cSocketConnector.cls $
' 
' 3     28.05.18 16:37 Wqw
' REF: timeout on create
'
' 2     14.05.18 12:39 Wqw
' REF: impl w/o ext dep on async socket class
'
' 1     25.04.18 10:39 Wqw
' Initial implementation
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "cSerialPortConnector"
Implements IDeviceConnector

Public Enum UcsConnectorSocketTypeEnum
    ucsSckStream = 1
    ucsSckDatagram = 2
End Enum

Public Enum UcsConnectorSocketOptionNameEnum
    ucsSsoDebug = &H1                       ' Debugging is enabled.
    ucsSsoAcceptConnection = &H2            ' Socket is listening.
    ucsSsoReuseAddress = &H4                ' The socket can be bound to an address which is already in use. Not applicable for ATM sockets.
    ucsSsoKeepAlive = &H8                   ' Keep-alives are being sent. Not supported on ATM sockets.
    ucsSsoDontRoute = &H10                  ' Routing is disabled. Not supported on ATM sockets.
    ucsSsoBroadcast = &H20                  ' Socket is configured for the transmission of broadcast messages.
    ucsSsoUseLoopback = &H40                ' Bypass hardware when possible.
    ucsSsoLinger = &H80                     ' Linger on close if unsent data is present.
    ucsSsoOutOfBandInline = &H100           ' Receives out-of-band data in the normal data stream.
    ucsSsoDontLinger = Not ucsSsoLinger     ' Close socket gracefully without lingering.
    ucsSsoExclusiveAddressUse = Not ucsSsoReuseAddress ' Enables a socket to be bound for exclusive access.
    ucsSsoSendBuffer = &H1001               ' Buffer size for sends.
    ucsSsoReceiveBuffer = &H1002            ' Buffer size for receives.
    ucsSsoSendLowWater = &H1003             ' Specifies the total per-socket buffer space reserved for receives.
    ucsSsoReceiveLowWater = &H1004          ' Receive low water mark.
    ucsSsoSendTimeout = &H1005              ' Sends time-out (available in Microsoft implementation of Windows Sockets 2).
    ucsSsoReceiveTimeout = &H1006           ' Receives time-out (available in Microsoft implementation of Windows Sockets 2).
    ucsSsoError = &H1007                    ' Get error status and clear.
    ucsSsoType = &H1008                     ' Get socket type.
'    ucsSsoGroupId = &H2001                  ' Reserved.
'    ucsSsoGroupPriority = &H2002            ' Reserved.
    ucsSsoMaxMsgSize = &H2003               ' Maximum size of a message for message-oriented socket types (for example, SOCK_DGRAM). Has no meaning for stream oriented sockets.
    ucsSsoProtocolInfo = &H2004             ' Description of protocol information for protocol that is bound to this socket.
    ucsSsoReuseUnicastPort = &H3007         ' Defer ephemeral port allocation for outbound connections
    ucsSsoMaxConnections = &H7FFFFFFF       ' Maximum queue length specifiable by listen.
End Enum

Public Enum UcsConnectorSocketOptionLevelEnum
    ucsSolSocket = &HFFFF&                  ' SOL_SOCKET
End Enum

'=========================================================================
' API
'=========================================================================

Private Const SOCKET_ERROR                  As Long = -1
Private Const INVALID_SOCKET                As Long = -1
Private Const AF_INET                       As Long = 2
'--- for ws_gethostbyname
Private Const INADDR_ANY                    As Long = 0
Private Const INADDR_NONE                   As Long = -1
'--- for ioctlsocket
Private Const FIONREAD                      As Long = &H4004667F
'--- for WSAGetLastError
Private Const WSABASEERR                    As Long = 10000
'Private Const WSAEINTR                      As Long = (WSABASEERR + 4)
Private Const WSAEINVAL                     As Long = (WSABASEERR + 22)
'Private Const WSAEMFILE                     As Long = (WSABASEERR + 24)
'Private Const WSAEWOULDBLOCK                As Long = (WSABASEERR + 35)
Private Const WSAEALREADY                   As Long = (WSABASEERR + 37)
'Private Const WSAENOTCONN                   As Long = (WSABASEERR + 57)
'Private Const WSANOTINITIALISED             As Long = (WSABASEERR + 93)
'--- for FormatMessage
Private Const FORMAT_MESSAGE_FROM_SYSTEM    As Long = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Args As Any) As Long
Private Declare Function ws_socket Lib "ws2_32" Alias "socket" (ByVal af As Long, ByVal s_type As Long, ByVal Protocol As Long) As Long
Private Declare Function ws_closesocket Lib "ws2_32" Alias "closesocket" (ByVal s As Long) As Long
Private Declare Function ws_connect Lib "ws2_32" Alias "connect" (ByVal s As Long, ByRef Name As SOCKADDR_IN, ByVal namelen As Long) As Long
Private Declare Function ws_shutdown Lib "ws2_32" Alias "shutdown" (ByVal s As Long, ByVal How As Long) As Long
Private Declare Function ws_bind Lib "ws2_32" Alias "bind" (ByVal s As Long, ByRef Name As SOCKADDR_IN, ByRef namelen As Long) As Long
Private Declare Function ws_htonl Lib "ws2_32" Alias "htonl" (ByVal hostlong As Long) As Long
Private Declare Function ws_htons Lib "ws2_32" Alias "htons" (ByVal hostshort As Long) As Integer
Private Declare Function ws_ntohs Lib "ws2_32" Alias "ntohs" (ByVal netshort As Long) As Integer
Private Declare Function ws_inet_ntoa Lib "ws2_32" Alias "inet_ntoa" (ByVal inn As Long) As Long
Private Declare Function ws_inet_addr Lib "ws2_32" Alias "inet_addr" (ByVal cp As String) As Long
Private Declare Function ws_gethostbyname Lib "ws2_32" Alias "gethostbyname" (ByVal host_name As String) As Long
Private Declare Function ws_getsockopt Lib "ws2_32" Alias "getsockopt" (ByVal s As Long, ByVal Level As Long, ByVal optname As Long, optval As Any, optlen As Long) As Long
Private Declare Function ws_setsockopt Lib "ws2_32" Alias "setsockopt" (ByVal s As Long, ByVal Level As Long, ByVal optname As Long, optval As Any, ByVal optlen As Long) As Long
Private Declare Function ws_ioctlsocket Lib "ws2_32" Alias "ioctlsocket" (ByVal s As Long, ByVal cmd As Long, argp As Long) As Long
Private Declare Function ws_recv Lib "ws2_32" Alias "recv" (ByVal s As Long, ByRef buf As Any, ByVal BufLen As Long, ByVal Flags As Long) As Long
Private Declare Function ws_recvfrom Lib "ws2_32" Alias "recvfrom" (ByVal s As Long, ByRef buf As Any, ByVal BufLen As Long, ByVal Flags As Long, ByRef from As SOCKADDR_IN, ByRef fromlen As Long) As Long
Private Declare Function ws_send Lib "ws2_32" Alias "send" (ByVal s As Long, ByRef buf As Any, ByVal BufLen As Long, ByVal Flags As Long) As Long
Private Declare Function ws_sendto Lib "ws2_32" Alias "sendto" (ByVal s As Long, ByRef buf As Any, ByVal BufLen As Long, ByVal Flags As Long, ByRef toaddr As SOCKADDR_IN, ByVal tolen As Long) As Long

Private Type SOCKADDR_IN
    sin_family          As Integer
    sin_port            As Integer
    sin_addr            As Long
    sin_zero(1 To 8)    As Byte
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Private m_hSocket               As Long
Private m_lResolvTimeout        As Long
Private m_lLastError            As Long

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunc As String)
    Debug.Print MODULE_NAME & "." & sFunc & ": " & Err.Description
    OutputDebugLog MODULE_NAME, sFunc & "(" & Erl & ")", "Run-time error: " & Err.Description
End Sub

Private Sub RaiseError(sFunc As String)
    Debug.Print MODULE_NAME & "." & sFunc & ": " & Err.Description
    OutputDebugLog MODULE_NAME, sFunc & "(" & Erl & ")", "Run-time error: " & Err.Description
    Err.Raise Err.Number, MODULE_NAME & "." & sFunc & "(" & Erl & ")" & vbCrLf & Err.Source, Err.Description
End Sub

Private Sub DebugLog(sFunc As String, sText As String)
    OutputDebugLog MODULE_NAME, sFunc, sText
End Sub

'=========================================================================
' Properties
'=========================================================================

Property Get LastError() As Long
    LastError = m_lLastError
End Property

Private Property Let pvLastError(ByVal lValue As Long)
    If lValue <> 0 Then
        m_lLastError = lValue
    Else
        m_lLastError = 0
    End If
End Property

Property Get SockOpt(ByVal OptionName As UcsConnectorSocketOptionNameEnum, Optional ByVal Level As UcsConnectorSocketOptionLevelEnum = ucsSolSocket) As Long
    Call ws_getsockopt(m_hSocket, Level, OptionName, SockOpt, 4)
    pvLastError = Err.LastDllError
End Property

Property Let SockOpt(ByVal OptionName As UcsConnectorSocketOptionNameEnum, Optional ByVal Level As UcsConnectorSocketOptionLevelEnum = ucsSolSocket, ByVal Value As Long)
    Call ws_setsockopt(m_hSocket, Level, OptionName, Value, 4)
    pvLastError = Err.LastDllError
End Property

Property Get AvailableBytes() As Long
    If ws_ioctlsocket(m_hSocket, FIONREAD, AvailableBytes) = SOCKET_ERROR Then
        pvLastError = Err.LastDllError
        AvailableBytes = SOCKET_ERROR
    End If
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function Create( _
            Optional ByVal SocketPort As Long, _
            Optional ByVal SocketType As UcsConnectorSocketTypeEnum = ucsSckStream, _
            Optional SocketAddress As String, _
            Optional ByVal Timeout As Long) As Boolean
    Const FUNC_NAME     As String = "Create"
    
    On Error GoTo EH
    If m_hSocket <> INVALID_SOCKET Then
        pvLastError = WSAEALREADY
        GoTo QH
    End If
    m_hSocket = ws_socket(AF_INET, SocketType, 0)
    If m_hSocket = INVALID_SOCKET Then
        pvLastError = Err.LastDllError
        GoTo QH
    End If
    m_lResolvTimeout = IIf(Timeout <> 0, Timeout, 3000)
    If Not Bind(SocketAddress, SocketPort) Then
        Close_
        GoTo QH
    End If
    '--- success
    Create = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Public Function Bind(Optional SocketAddress As String, Optional ByVal SocketPort As Long) As Boolean
    Const FUNC_NAME     As String = "Bind"
    Dim uAddr           As SOCKADDR_IN
    
    On Error GoTo EH
    If Not pvToSockAddr(SocketAddress, SocketPort, uAddr) Then
        GoTo QH
    End If
    If ws_bind(m_hSocket, uAddr, LenB(uAddr)) = SOCKET_ERROR Then
        pvLastError = Err.LastDllError
        GoTo QH
    End If
    '--- success
    Bind = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Public Sub Close_()
    '--- note: used in terminate -> no error handling
    If m_hSocket <> INVALID_SOCKET Then
        Call ws_shutdown(m_hSocket, 1)
        Call ws_closesocket(m_hSocket)
        m_hSocket = INVALID_SOCKET
    End If
End Sub

Public Function Connect(HostAddress As String, ByVal HostPort As Long) As Boolean
    Const FUNC_NAME     As String = "Connect"
    Dim uAddr           As SOCKADDR_IN
    
    On Error GoTo EH
    If pvToSockAddr(HostAddress, HostPort, uAddr) Then
        If SockOpt(ucsSsoType) = ucsSckStream Then
            If ws_connect(m_hSocket, uAddr, LenB(uAddr)) = SOCKET_ERROR Then
                pvLastError = Err.LastDllError
                GoTo QH
            End If
        End If
    End If
    '--- success
    Connect = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Public Function SyncReceiveArray( _
            Buffer() As Byte, _
            Optional HostAddress As String = STR_CHR1, _
            Optional HostPort As Long, _
            Optional ByVal Timeout As Long, _
            Optional ByVal Flags As Long = 0) As Boolean
    Const FUNC_NAME     As String = "SyncReceiveArray"
    Dim lAvailable      As Long
    Dim lResult         As Long
    
    On Error GoTo EH
    lAvailable = AvailableBytes
    If lAvailable <= 0 Then
        lAvailable = SockOpt(ucsSsoReceiveBuffer)
    End If
    If lAvailable <= 0 Then
        lAvailable = 4096
    End If
    ReDim Buffer(0 To lAvailable - 1) As Byte
    Do
        If Not SyncReceive(VarPtr(Buffer(0)), lAvailable, lResult, HostAddress, HostPort, Timeout, Flags) Then
            GoTo QH
        End If
    Loop While lResult = 0
    If UBound(Buffer) + 1 <> lResult Then
        ReDim Preserve Buffer(0 To lResult - 1) As Byte
    End If
    '--- success
    SyncReceiveArray = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Public Function SyncReceive( _
            ByVal BufPtr As Long, _
            ByVal BufLen As Long, _
            Received As Long, _
            Optional HostAddress As String = STR_CHR1, _
            Optional HostPort As Long, _
            Optional ByVal Timeout As Long, _
            Optional ByVal Flags As Long = 0) As Boolean
    Const FUNC_NAME     As String = "SyncReceive"
    Dim uAddr           As SOCKADDR_IN
    Dim lResult         As Long
    
    On Error GoTo EH
    If Timeout <> 0 Then
        SockOpt(ucsSsoReceiveTimeout) = Timeout
    End If
    Do
        If HostAddress <> STR_CHR1 Then
            lResult = ws_recvfrom(m_hSocket, ByVal BufPtr, BufLen, Flags, uAddr, LenB(uAddr))
        Else
            lResult = ws_recv(m_hSocket, ByVal BufPtr, BufLen, Flags)
        End If
        If lResult >= 0 Then
            Received = lResult
            Exit Do
        End If
        pvLastError = Err.LastDllError
        GoTo QH
    Loop
    If uAddr.sin_family <> 0 Then
        pvFromSockAddr uAddr, HostAddress, HostPort
    End If
    '--- success
    SyncReceive = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Public Function SyncSendArray( _
            Buffer() As Byte, _
            Optional HostAddress As String, _
            Optional ByVal HostPort As Long, _
            Optional ByVal Timeout As Long, _
            Optional ByVal Flags As Long = 0) As Boolean
    Const FUNC_NAME     As String = "SyncSendArray"
    
    On Error GoTo EH
    If UBound(Buffer) < 0 Then
        SyncSendArray = True
    Else
        SyncSendArray = SyncSend(VarPtr(Buffer(0)), UBound(Buffer) + 1, HostAddress, HostPort, Timeout, Flags)
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Public Function SyncSend( _
            ByVal BufPtr As Long, _
            ByVal BufLen As Long, _
            Optional HostAddress As String, _
            Optional ByVal HostPort As Long, _
            Optional ByVal Timeout As Long, _
            Optional ByVal Flags As Long = 0) As Boolean
    Const FUNC_NAME     As String = "SyncSend"
    Dim uAddr           As SOCKADDR_IN
    Dim lResult         As Long
    
    On Error GoTo EH
    If Timeout <> 0 Then
        SockOpt(ucsSsoSendTimeout) = Timeout
    End If
    If LenB(HostAddress) <> 0 Then
        If Not pvToSockAddr(HostAddress, HostPort, uAddr) Then
            GoTo QH
        End If
    End If
    Do
        Do
            If uAddr.sin_family <> 0 Then
                lResult = ws_sendto(m_hSocket, ByVal BufPtr, BufLen, Flags, uAddr, LenB(uAddr))
            Else
                lResult = ws_send(m_hSocket, ByVal BufPtr, BufLen, Flags)
            End If
            If lResult >= 0 Then
                Exit Do
            End If
            pvLastError = Err.LastDllError
            GoTo QH
        Loop
        If lResult = BufLen Then
            Exit Do
        End If
        BufPtr = UnsignedAdd(BufPtr, lResult)
        BufLen = BufLen - lResult
    Loop
    SyncSend = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Public Function GetErrorDescription(ByVal ErrorNumber As Long) As String
    Const FUNC_NAME     As String = "GetErrorDescription"
    Dim lSize           As Long
    
    On Error GoTo EH
    GetErrorDescription = Space$(2000)
    lSize = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0&, ErrorNumber, 0, GetErrorDescription, Len(GetErrorDescription) + 1, 0)
    If lSize > 2 Then
        If Mid$(GetErrorDescription, lSize - 1, 2) = vbCrLf Then
            lSize = lSize - 2
        End If
    End If
    GetErrorDescription = Left$(GetErrorDescription, lSize)
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

'= private ===============================================================

Private Function pvToSockAddr( _
            Address As String, _
            ByVal Port As Long, _
            uAddr As SOCKADDR_IN) As Boolean
    Dim lPtr            As Long
    
    uAddr.sin_family = AF_INET
    If LenB(Address) = 0 Then
        uAddr.sin_addr = ws_htonl(INADDR_ANY)
    Else
        uAddr.sin_addr = ws_inet_addr(Address)
        If uAddr.sin_addr = INADDR_NONE Then
            '--- ToDo: use m_lResolvTimeout
            lPtr = ws_gethostbyname(Address)
            If lPtr = 0 Then
                pvLastError = WSAEINVAL
                GoTo QH
            End If
            Call CopyMemory(lPtr, ByVal UnsignedAdd(lPtr, 12), 4)
            Call CopyMemory(lPtr, ByVal lPtr, 4)
            Call CopyMemory(uAddr.sin_addr, ByVal lPtr, 4)
        End If
    End If
    uAddr.sin_port = ws_htons(Port)
    '--- success
    pvToSockAddr = True
QH:
End Function

Private Sub pvFromSockAddr(uAddr As SOCKADDR_IN, Address As String, Port As Long)
    Port = pvToLong(ws_ntohs(uAddr.sin_port))
    Address = pvToString(ws_inet_ntoa(uAddr.sin_addr))
End Sub

Private Function pvToLong(ByVal nLoWord As Integer) As Long
    Call CopyMemory(pvToLong, nLoWord, 2)
End Function

Private Function pvToString(ByVal lPtr As Long) As String
    If lPtr <> 0 Then
        pvToString = String$(lstrlen(lPtr), 0)
        Call lstrcpy(pvToString, lPtr)
    End If
End Function

#If Not ImplUseShared Then
Private Function UnsignedAdd(ByVal lUnsignedPtr As Long, ByVal lSignedOffset As Long) As Long
    '--- note: safely add *signed* offset to *unsigned* ptr for *unsigned* retval w/o overflow in LARGEADDRESSAWARE processes
    UnsignedAdd = ((lUnsignedPtr Xor &H80000000) + lSignedOffset) Xor &H80000000
End Function
#End If

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Initialize()
    m_hSocket = INVALID_SOCKET
End Sub

Private Sub Class_Terminate()
    Close_
End Sub

'=========================================================================
' IDeviceConnector interface
'=========================================================================

Public Function Init(Device As String, Optional ByVal Timeout As Long) As Boolean
    Const FUNC_NAME     As String = "Init"
    Dim vSplit          As Variant
    
    On Error GoTo EH
    vSplit = Split(Device, ":")
    If Not Create(SocketType:=ucsSckStream, Timeout:=Timeout) Then
        DebugLog FUNC_NAME, "Create failed, LastError=" & LastError
        GoTo QH
    End If
    If Not Connect(At(vSplit, 0), At(vSplit, 1)) Then
        DebugLog FUNC_NAME, "Connect failed, Device=" & Device & ", LastError=" & LastError
        GoTo QH
    End If
    '--- success
    Init = True
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Function IDeviceConnector_ReadData(Result() As Byte, Optional ByVal Timeout As Long) As Boolean
    Const FUNC_NAME     As String = "IDeviceConnector_ReadData"
    
    On Error GoTo EH
    If Not SyncReceiveArray(Result, Timeout:=Timeout) Then
        DebugLog FUNC_NAME, "SyncReceiveArray failed, LastError=" & LastError
        GoTo QH
    End If
    '--- success
    IDeviceConnector_ReadData = True
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Function IDeviceConnector_WriteData(Data() As Byte, Optional ByVal Timeout As Long) As Boolean
    Const FUNC_NAME     As String = "IDeviceConnector_WriteData"
    
    On Error GoTo EH
    If Not SyncSendArray(Data, Timeout:=Timeout) Then
        DebugLog FUNC_NAME, "SyncSendArray failed, LastError=" & LastError
        GoTo QH
    End If
    '--- success
    IDeviceConnector_WriteData = True
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Function IDeviceConnector_GetLastError(Optional ErrCategory As UcsConnectorErrorsEnum) As String
    IDeviceConnector_GetLastError = GetErrorDescription(m_lLastError)
End Function

