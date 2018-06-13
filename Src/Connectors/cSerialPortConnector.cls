VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSerialPortConnector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================================================
' $Header: /UcsFiscalPrinter/Src/Connectors/cSerialPortConnector.cls 4     13.06.18 12:19 Wqw $
'
'   Unicontsoft Fiscal Printers Project
'   Copyright (c) 2008-2018 Unicontsoft
'
'   COM port connector
'
' $Log: /UcsFiscalPrinter/Src/Connectors/cSerialPortConnector.cls $
' 
' 4     13.06.18 12:19 Wqw
' REF: clear persistent on terminate
'
' 3     28.05.18 16:37 Wqw
' REF: impl connector get last error
'
' 2     14.05.18 12:39 Wqw
' REF: read data
'
' 1     25.04.18 10:39 Wqw
' Initial implementation
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "cSerialPortConnector"
Implements IDeviceConnector

'=========================================================================
' API
'=========================================================================

'--- for CreateFile
Private Const INVALID_HANDLE_VALUE          As Long = -1
Private Const ERROR_ACCESS_DENIED           As Long = 5
'--- for WaitCommEvent
Private Const EV_RXCHAR                     As Long = &H1                '  Any Character received
Private Const EV_ERR                        As Long = &H80
'--- for WaitForMultipleObjects
Private Const WAIT_OBJECT_0                 As Long = 0
Private Const WAIT_TIMEOUT                  As Long = 258
'--- error codes
Private Const ERROR_IO_PENDING              As Long = 997
Private Const ERROR_HANDLE_EOF              As Long = 38
'--- for PurgeComm
Private Const PURGE_TXCLEAR                 As Long = &H4     '  Kill the transmit queue if there.
Private Const PURGE_RXCLEAR                 As Long = &H8     '  Kill the typeahead buffer if there.

Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function SetCommTimeouts Lib "kernel32" (ByVal hFile As Long, lpCommTimeouts As COMMTIMEOUTS) As Long
Private Declare Function BuildCommDCB Lib "kernel32" Alias "BuildCommDCBA" (ByVal lpDef As String, lpDCB As DCB) As Long
Private Declare Function SetCommState Lib "kernel32" (ByVal hCommDev As Long, lpDCB As DCB) As Long
Private Declare Function SetCommMask Lib "kernel32" (ByVal hFile As Long, ByVal dwEvtMask As Long) As Long
Private Declare Function WaitCommEvent Lib "kernel32" (ByVal hFile As Long, lpEvtMask As Long, lpOverlapped As OVERLAPPED) As Long
Private Declare Function CreateEvent Lib "kernel32" Alias "CreateEventA" (ByVal lpEventAttributes As Long, ByVal bManualReset As Long, ByVal bInitialState As Long, ByVal lpName As Long) As Long
Private Declare Function WaitForMultipleObjects Lib "kernel32" (ByVal nCount As Long, lpHandles As Long, ByVal bWaitAll As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CancelIo Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function PurgeComm Lib "kernel32" (ByVal hFile As Long, ByVal dwFlags As Long) As Long
Private Declare Function GetOverlappedResult Lib "kernel32" (ByVal hFile As Long, lpOverlapped As OVERLAPPED, lpNumberOfBytesTransferred As Long, ByVal bWait As Long) As Long

Private Type COMMTIMEOUTS
    ReadIntervalTimeout         As Long
    ReadTotalTimeoutMultiplier  As Long
    ReadTotalTimeoutConstant    As Long
    WriteTotalTimeoutMultiplier As Long
    WriteTotalTimeoutConstant   As Long
End Type

Private Type DCB
    DCBlength                   As Long
    BaudRate                    As Long
    fBinary                     As Long
    fParity                     As Long
    fOutxCtsFlow                As Long
    fOutxDsrFlow                As Long
    fDtrControl                 As Long
    fDsrSensitivity             As Long
    fTXContinueOnXoff           As Long
    fOutX                       As Long
    fInX                        As Long
    fErrorChar                  As Long
    fNull                       As Long
    fRtsControl                 As Long
    fAbortOnError               As Long
    fDummy2                     As Long
    wReserved                   As Integer
    XonLim                      As Integer
    XoffLim                     As Integer
    ByteSize                    As Byte
    Parity                      As Byte
    StopBits                    As Byte
    XonChar                     As Byte
    XoffChar                    As Byte
    ErrorChar                   As Byte
    EofChar                     As Byte
    EvtChar                     As Byte
End Type

Private Type OVERLAPPED
    Internal                    As Long
    InternalHigh                As Long
    offset                      As Long
    OffsetHigh                  As Long
    hEvent                      As Long
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_INTERNAL              As String = "No device info set|CreateFile failed: %1|SetCommTimeouts failed: %1|BuildCommDCB failed: %1|SetCommState failed: %1|WriteFile failure: %1|Timeout waiting for response|ReadFile failed: %1|WaitCommEvent failed: %1"
Private Const DEF_TIMEOUT               As Long = 3000
Private Const DEF_COMM_EVENT_TIMEOUT    As Long = 250

Private m_hComm                 As Long
Private m_lTimeout              As Long
Private m_sDevice               As String
Private m_sLastError            As String
Private m_vLocalizedText(0 To [_ucsFscLciMax] - 1) As Variant
Private m_lCommEventTimeout     As Long
Private m_bClearPersistent      As Boolean

Private Enum UcsInternalErrors
    ucsErrNoDeviceInfoSet
    ucsErrCreateFileFailed
    ucsErrSetCommTimeoutsFailed
    ucsErrBuildCommDCBFailed
    ucsErrSetCommStateFailed
    ucsErrWriteFileFailure
    ucsErrTimeoutWaitingForResponse
    ucsErrReadFileFailed
    ucsErrWaitCommEventFailed
End Enum

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

Property Get LocalizedText(ByVal eIdx As UcsFiscalLocalizedIndexesEnum) As String
    LocalizedText = Join(m_vLocalizedText(eIdx), "|")
End Property

Property Let LocalizedText(ByVal eIdx As UcsFiscalLocalizedIndexesEnum, sValue As String)
    m_vLocalizedText(eIdx) = SplitOrReindex(sValue, "|")
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function Init(Device As String, Optional ByVal Timeout As Long) As Boolean
    Const FUNC_NAME     As String = "Init"
    Dim vSplit          As Variant
    
    On Error GoTo EH
    DebugLog FUNC_NAME, "Device=" & Device & ", Timeout=" & Timeout
    PortWrapper.CloseComPort m_hComm
    m_hComm = INVALID_HANDLE_VALUE
    m_lTimeout = IIf(Timeout <= 0, DEF_TIMEOUT, Timeout)
    vSplit = Split(Device, ",")
    m_sDevice = Trim$(At(vSplit, 0, "COM1")) & "," & C_Lng(At(vSplit, 1, "9600")) & "," & At(vSplit, 2) & "," & C_Lng(At(vSplit, 3, "8")) & "," & _
        IIf(UCase$(At(vSplit, 4, "N")) = "Y", "Y", "N") & "," & C_Lng(At(vSplit, 5, "1"))
    m_hComm = pvOpenPort(m_sDevice, m_lTimeout, m_sLastError)
    If LenB(m_sLastError) <> 0 Then
        GoTo QH
    End If
    '--- success
    Init = True
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function WriteData(Data() As Byte, Optional ByVal Timeout As Long) As Boolean
    Const FUNC_NAME     As String = "pvSendCommand"
    Dim lWritten        As Long
    Dim uOver           As OVERLAPPED
    Dim lResult         As Long
    Dim lLastErr        As Long
    
    On Error GoTo EH
    uOver.hEvent = CreateEvent(0, 1, 0, 0)
    Call PurgeComm(m_hComm, PURGE_TXCLEAR Or PURGE_RXCLEAR)
    lResult = WriteFile(m_hComm, Data(0), UBound(Data) + 1, lWritten, uOver)
    lLastErr = Err.LastDllError
    If lResult = 0 Then
        If lLastErr <> ERROR_IO_PENDING Then
            m_bClearPersistent = True
            m_sLastError = Printf(pvInternal(ucsErrWriteFileFailure), GetApiErr(lLastErr))
            DebugLog FUNC_NAME, "WriteFile, lResult=" & lResult & ", Err.LastDllError=" & lLastErr
            GoTo QH
        End If
    End If
    lResult = WaitForMultipleObjects(1, uOver.hEvent, 0, Timeout)
    If lResult <> WAIT_OBJECT_0 And Timeout <> 0 Then
        Call CancelIo(m_hComm)
        If lResult = WAIT_TIMEOUT Then
            m_bClearPersistent = True
            m_sLastError = pvInternal(ucsErrTimeoutWaitingForResponse)
        End If
        DebugLog FUNC_NAME, "CancelIo after WriteFile, lResult=" & lResult
        GoTo QH
    End If
    '--- success
    WriteData = True
QH:
    Call CloseHandle(uOver.hEvent)
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Public Function ReadData(Result() As Byte, Optional ByVal Timeout As Long) As Boolean
    Const FUNC_NAME     As String = "pvSendCommand"
    Dim uOver           As OVERLAPPED
    Dim lResult         As Long
    Dim lMask           As Long
    Dim lActual         As Long
    Dim lRetryCount     As Long
    Dim lIterTimeout    As Long
    Dim lLastErr        As Long
    
    On Error GoTo EH
    uOver.hEvent = CreateEvent(0, 1, 0, 0)
    ReDim Result(0 To 65535) As Byte
    Call SetCommMask(m_hComm, EV_RXCHAR Or EV_ERR)
    '--- async read response
    Do
        lIterTimeout = LimitLong(m_lCommEventTimeout, , Timeout - lRetryCount * m_lCommEventTimeout)
        lActual = 0
        lResult = ReadFile(m_hComm, Result(0), UBound(Result) + 1, lActual, uOver)
        lLastErr = Err.LastDllError
        If Timeout <> 0 Then
            If lResult <> 0 And lActual = 0 Then
WaitEvent:
                lResult = WaitCommEvent(m_hComm, lMask, uOver)
                lLastErr = Err.LastDllError
                If lResult = 0 Then
                    Select Case lLastErr
                    Case ERROR_IO_PENDING
                        lResult = WaitForMultipleObjects(1, uOver.hEvent, 0, lIterTimeout)
                        Select Case lResult
                        Case WAIT_OBJECT_0
                            If GetOverlappedResult(m_hComm, uOver, lActual, 0) = 0 Then
                                DebugLog FUNC_NAME, "GetOverlappedResult after WaitCommEvent, Err.LastDllError=" & Err.LastDllError
                                lRetryCount = lRetryCount + 1
                                If lRetryCount * m_lCommEventTimeout >= Timeout Then
                                    GoTo TimeoutReached
                                End If
                            End If
                        Case WAIT_TIMEOUT
                            lRetryCount = lRetryCount + 1
                        End Select
                    Case Else
                        m_bClearPersistent = True
                        m_sLastError = Printf(pvInternal(ucsErrWaitCommEventFailed), GetApiErr(lLastErr))
                        DebugLog FUNC_NAME, "WaitCommEvent, Err.LastDllError=" & lLastErr
                        GoTo QH
                    End Select
                End If
                lResult = ReadFile(m_hComm, Result(0), UBound(Result) + 1, lActual, uOver)
                lLastErr = Err.LastDllError
            End If
            If lResult = 0 Then
                Select Case lLastErr
                Case ERROR_IO_PENDING
                    lResult = WaitForMultipleObjects(1, uOver.hEvent, 0, lIterTimeout)
                    Select Case lResult
                    Case WAIT_OBJECT_0
                        If GetOverlappedResult(m_hComm, uOver, lActual, 0) = 0 Then
                            DebugLog FUNC_NAME, "GetOverlappedResult after ReadFile, Err.LastDllError=" & Err.LastDllError
                            lRetryCount = lRetryCount + 1
                            If lRetryCount * m_lCommEventTimeout >= Timeout Then
                                GoTo TimeoutReached
                            End If
                        End If
                    Case WAIT_TIMEOUT
                        lRetryCount = lRetryCount + 1
                    End Select
                Case ERROR_HANDLE_EOF
                    GoTo WaitEvent
                Case Else
                    m_bClearPersistent = True
                    m_sLastError = Printf(pvInternal(ucsErrReadFileFailed), GetApiErr(lLastErr))
                    DebugLog FUNC_NAME, "ReadFile, Err.LastDllError=" & lLastErr
                    GoTo QH
                End Select
            End If
            If lResult <> 0 And lRetryCount * m_lCommEventTimeout >= Timeout Then
TimeoutReached:
                Call CancelIo(m_hComm)
                DebugLog FUNC_NAME, "CancelIo after ReadFile, lResult=" & lResult & ", lIterTimeout= " & lIterTimeout & ", m_lCommEventTimeout=" & m_lCommEventTimeout & ", Timeout=" & Timeout
                m_bClearPersistent = True
                m_sLastError = pvInternal(ucsErrTimeoutWaitingForResponse)
                GoTo QH
            End If
        End If
        If lActual > 0 Then
            ReDim Preserve Result(0 To lActual + 1) As Byte
            '--- success
            ReadData = True
            Exit Do
        ElseIf Timeout = 0 Then
            Exit Do
        End If
    Loop
QH:
    Call CloseHandle(uOver.hEvent)
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Public Function GetLastError() As String
    GetLastError = m_sLastError
End Function

'= private ===============================================================

Private Function pvOpenPort(sDevice As String, ByVal lTimeout As Long, Optional Error As String) As Long
    Const FUNC_NAME     As String = "pvOpenPort"
    Dim vSplit          As Variant
    Dim hComm           As Long
    Dim uTimeout        As COMMTIMEOUTS
    Dim uDCB            As DCB
    Dim lLastErr        As Long
    Dim dblTimer        As Double
    
    On Error GoTo EH
    vSplit = Split(sDevice, ",")
    If LenB(At(vSplit, 0)) = 0 Then
        Error = pvInternal(ucsErrNoDeviceInfoSet)
        GoTo QH
    End If
    hComm = PortWrapper.OpenComPort(At(vSplit, 0), At(vSplit, 2) = "P")
    If hComm = INVALID_HANDLE_VALUE Then
        lLastErr = Err.LastDllError
        Error = Printf(pvInternal(ucsErrCreateFileFailed), GetApiErr(lLastErr))
        If lLastErr = ERROR_ACCESS_DENIED And lTimeout <> 0 Then
            dblTimer = DateTimer + lTimeout / 1000#
            Do While DateTimer < dblTimer
                Call Sleep(100)
                hComm = PortWrapper.OpenComPort(At(vSplit, 0), At(vSplit, 2) = "P")
                If hComm <> INVALID_HANDLE_VALUE Then
                    Error = vbNullString
                    Exit Do
                End If
            Loop
        End If
        If LenB(Error) <> 0 Then
            GoTo QH
        End If
    End If
    With uTimeout
        .ReadIntervalTimeout = -1
        .ReadTotalTimeoutConstant = 0
        .ReadTotalTimeoutMultiplier = 0
        .WriteTotalTimeoutConstant = 500
        .WriteTotalTimeoutMultiplier = 1
    End With
    If SetCommTimeouts(hComm, uTimeout) = 0 Then
        Error = Printf(pvInternal(ucsErrSetCommTimeoutsFailed), GetApiErr(Err.LastDllError))
        PortWrapper.CloseComPort hComm, ClearPersistent:=True
        GoTo QH
    End If
    If BuildCommDCB("baud=" & At(vSplit, 1, 9600) & " data=" & At(vSplit, 3, 8) & " parity=" & At(vSplit, 4, "N") & " stop=" & At(vSplit, 5, 1), uDCB) = 0 Then
        Error = Printf(pvInternal(ucsErrBuildCommDCBFailed), GetApiErr(Err.LastDllError))
        PortWrapper.CloseComPort hComm, ClearPersistent:=True
        GoTo QH
    End If
    If SetCommState(hComm, uDCB) = 0 Then
        Error = Printf(pvInternal(ucsErrSetCommStateFailed), GetApiErr(Err.LastDllError))
        PortWrapper.CloseComPort hComm, ClearPersistent:=True
        GoTo QH
    End If
    pvOpenPort = hComm
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Function pvInternal(ByVal lIdx As UcsInternalErrors) As String
    Const FUNC_NAME     As String = "pvInternal"
    
    On Error GoTo EH
    pvInternal = At(m_vLocalizedText(ucsFscLciInternalErrors), lIdx)
    If LenB(pvInternal) = 0 Then
        pvInternal = At(Split(STR_INTERNAL, "|"), lIdx)
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Initialize()
    m_hComm = INVALID_HANDLE_VALUE
    m_lCommEventTimeout = DEF_COMM_EVENT_TIMEOUT
End Sub

Private Sub Class_Terminate()
    PortWrapper.CloseComPort m_hComm, m_bClearPersistent
    m_hComm = INVALID_HANDLE_VALUE
End Sub

'=========================================================================
' IDeviceConnector interface
'=========================================================================

Private Function IDeviceConnector_GetLastError(Optional ErrCategory As UcsConnectorErrorsEnum) As String
    IDeviceConnector_GetLastError = GetLastError()
End Function

Private Function IDeviceConnector_Init(Device As String, Optional ByVal Timeout As Long) As Boolean
    IDeviceConnector_Init = Init(Device, Timeout)
End Function

Private Function IDeviceConnector_ReadData(Result() As Byte, Optional ByVal Timeout As Long) As Boolean
    IDeviceConnector_ReadData = ReadData(Result, Timeout)
End Function

Private Function IDeviceConnector_WriteData(Data() As Byte, Optional ByVal Timeout As Long) As Boolean
    IDeviceConnector_WriteData = WriteData(Data, Timeout)
End Function

