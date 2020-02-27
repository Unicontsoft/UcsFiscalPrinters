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
'
' UcsFP20 (c) 2008-2019 by Unicontsoft
'
' Unicontsoft Fiscal Printers Component 2.0
'
' This project is licensed under the terms of the MIT license
' See the LICENSE file in the project root for more information
'
'=========================================================================
'
' COM port connector
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "cSerialPortConnector"
Implements IDeviceConnector

#Const ImplUseShared = (USE_SHARED <> 0)

#If Not ImplUseShared Then

    Public Enum UcsFiscalLocalizedIndexesEnum
        ucsFscLciInternalErrors
        [_ucsFscLciMax]
    End Enum
    
    Private Const STR_CONNECTOR_ERRORS      As String = "No device info set|CreateFile failed: %1|SetCommTimeouts failed: %1|BuildCommDCB failed: %1|SetCommState failed: %1|WriteFile failure: %1|Timeout waiting for response|ReadFile failed: %1|WaitCommEvent failed: %1"
    Private Const vbLogEventTypeDebug       As Long = vbLogEventTypeInformation + 1
    
    Private Logger                  As Object
    Private PortWrapper             As New cPortWrapper

#End If
    
'=========================================================================
' API
'=========================================================================

'--- for CreateFile
Private Const INVALID_HANDLE_VALUE          As Long = -1
Private Const ERROR_ACCESS_DENIED           As Long = 5
'--- for WaitCommEvent
Private Const EV_RXCHAR                     As Long = &H1
Private Const EV_ERR                        As Long = &H80
'--- for WaitForMultipleObjects
Private Const WAIT_OBJECT_0                 As Long = 0
Private Const WAIT_TIMEOUT                  As Long = 258
'--- error codes
Private Const ERROR_IO_PENDING              As Long = 997
Private Const ERROR_HANDLE_EOF              As Long = 38
'--- for PurgeComm
Private Const PURGE_TXCLEAR                 As Long = &H4
Private Const PURGE_RXCLEAR                 As Long = &H8
'--- for VariantChangeType
Private Const VARIANT_ALPHABOOL             As Long = 2
'--- for FormatMessage
Private Const FORMAT_MESSAGE_FROM_SYSTEM    As Long = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
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
#If Not ImplUseShared Then
    Private Declare Function VariantChangeType Lib "oleaut32" (vDest As Variant, Src As Variant, ByVal wFlags As Integer, ByVal vt As Long) As Long
    Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
    Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
    Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Args As Any) As Long
#End If

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
    Offset                      As Long
    OffsetHigh                  As Long
    hEvent                      As Long
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const DEF_TIMEOUT               As Long = 3000
Private Const DEF_COMM_EVENT_TIMEOUT    As Long = 500
Private Const DEF_SPEED                 As Long = 115200

Private m_hComm                 As Long
Private m_lTimeout              As Long
Private m_sDevice               As String
Private m_sLastError            As String
Private m_vLocalizedText(0 To [_ucsFscLciMax] - 1) As Variant
Private m_lCommEventTimeout     As Long
Private m_bClearPersistent      As Boolean
'--- overlapped I/O buffers
Private m_lWritten              As Long
Private m_lActual               As Long
Private m_lMask                 As Long
Private m_baBuffer()            As Byte

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunction As String)
    m_sLastError = Err.Description
    Debug.Print "Critical error: " & Err.Description & " [" & MODULE_NAME & "." & sFunction & "]"
    If Not Logger Is Nothing Then
        Logger.Log vbLogEventTypeError, MODULE_NAME, sFunction & "(" & Erl & ")", Err.Description
    End If
End Sub

Private Sub RaiseError(sFunction As String)
    m_sLastError = Err.Description
    Debug.Print "Critical error: " & Err.Description & " [" & MODULE_NAME & "." & sFunction & "]"
    If Not Logger Is Nothing Then
        Logger.Log vbLogEventTypeError, MODULE_NAME, sFunction & "(" & Erl & ")", Err.Description
    End If
    Err.Raise Err.Number, MODULE_NAME & "." & sFunction & "(" & Erl & ")" & vbCrLf & Err.Source, Err.Description
End Sub

Private Sub DebugLog(sFunction As String, sText As String)
    If Not Logger Is Nothing Then
        Logger.Log vbLogEventTypeDebug, MODULE_NAME, sFunction, sText
    End If
End Sub

'=========================================================================
' Properties
'=========================================================================

Property Get Device() As String
    Device = m_sDevice
End Property

Property Get LocalizedText(ByVal eIdx As UcsFiscalLocalizedIndexesEnum) As String
    LocalizedText = Join(m_vLocalizedText(eIdx), "|")
End Property

Property Let LocalizedText(ByVal eIdx As UcsFiscalLocalizedIndexesEnum, sValue As String)
    m_vLocalizedText(eIdx) = SplitOrReindex(sValue, "|")
End Property

Property Get hComm() As Long
    hComm = m_hComm
End Property

Property Let hComm(ByVal lValue As Long)
    m_hComm = lValue
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
    m_sDevice = Trim$(At(vSplit, 0, "COM1")) & "," & C_Lng(At(vSplit, 1, DEF_SPEED)) & "," & At(vSplit, 2) & "," & C_Lng(At(vSplit, 3, "8")) & "," & _
        IIf(UCase$(At(vSplit, 4, "N")) = "Y", "Y", "N") & "," & C_Lng(At(vSplit, 5, "1"))
    m_hComm = pvOpenPort(m_sDevice, m_lTimeout, m_sLastError)
    If LenB(m_sLastError) <> 0 Then
        GoTo QH
    End If
    Call SetCommMask(m_hComm, EV_RXCHAR Or EV_ERR)
    '--- success
    Init = True
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function WriteData(Data() As Byte, Optional ByVal Timeout As Long) As Boolean
    Const FUNC_NAME     As String = "WriteData"
    Dim dblTimerEx      As Long
    Dim lOffset         As Long
    Dim uWriteOver      As OVERLAPPED
    Dim lResult         As Long
    Dim lLastErr        As Long
    Dim bPendingWrite   As Boolean
    
    On Error GoTo EH
    dblTimerEx = TimerEx
    uWriteOver.hEvent = CreateEvent(0, 1, 0, 0)
    Call PurgeComm(m_hComm, PURGE_TXCLEAR Or PURGE_RXCLEAR)
RetryWrite:
    m_lWritten = 0
    lResult = WriteFile(m_hComm, Data(lOffset), UBound(Data) + 1 - lOffset, m_lWritten, uWriteOver)
    lLastErr = Err.LastDllError
    If lResult = 0 Then
        If lLastErr = ERROR_IO_PENDING Then
            bPendingWrite = True
            lResult = WaitForMultipleObjects(1, uWriteOver.hEvent, 0, Timeout)
            lLastErr = Err.LastDllError
            If lResult <> WAIT_OBJECT_0 Then
                m_bClearPersistent = True
                If lResult = WAIT_TIMEOUT Then
                    m_sLastError = pvInternal(ucsErrTimeoutWaitingForResponse)
                    DebugLog FUNC_NAME, "WriteFile timeout, Timeout=" & Timeout & ", m_sLastError=" & m_sLastError & ", Elapsed=" & Format$(TimerEx - dblTimerEx, "0.000")
                Else
                    m_sLastError = GetErrorDescription(lLastErr)
                    DebugLog FUNC_NAME, "WaitForMultipleObjects failure after WriteFile, lResult=" & lResult & ", Timeout=" & Timeout & ", m_sLastError=" & m_sLastError & ", Elapsed=" & Format$(TimerEx - dblTimerEx, "0.000")
                End If
                GoTo QH
            End If
            bPendingWrite = False
            m_lWritten = 0
            Call GetOverlappedResult(m_hComm, uWriteOver, m_lWritten, 0)
        Else
            m_bClearPersistent = True
            m_sLastError = Printf(pvInternal(ucsErrWriteFileFailure), GetErrorDescription(lLastErr))
            DebugLog FUNC_NAME, "WriteFile, lResult=" & lResult & ", Err.LastDllError=" & lLastErr
            GoTo QH
        End If
    End If
    lOffset = lOffset + m_lWritten
    If m_lWritten > 0 And lOffset < UBound(Data) + 1 Then
        GoTo RetryWrite
    End If
    '--- success
    WriteData = True
QH:
    If bPendingWrite Then
        Call CancelIo(m_hComm)
    End If
    Call CloseHandle(uWriteOver.hEvent)
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Public Function ReadData(Result() As Byte, Optional ByVal Timeout As Long) As Boolean
    Const FUNC_NAME     As String = "ReadData"
    Dim dblTimerEx      As Double
    Dim uReadOver       As OVERLAPPED
    Dim uWaitOver       As OVERLAPPED
    Dim lResult         As Long
    Dim lLastErr        As Long
    Dim lRetryCount     As Long
    Dim lActual         As Long
    Dim bPendingRead    As Boolean
    Dim bPendingWait    As Boolean
    
    On Error GoTo EH
    dblTimerEx = TimerEx
    uReadOver.hEvent = CreateEvent(0, 1, 0, 0)
    uWaitOver.hEvent = CreateEvent(0, 1, 0, 0)
    If UBound(m_baBuffer) < 0 Then
        ReDim m_baBuffer(0 To 65535) As Byte
    End If
    GoTo InLoop
    Do
        lRetryCount = lRetryCount + 1
        m_lMask = 0
        lResult = WaitCommEvent(m_hComm, m_lMask, uWaitOver)
        lLastErr = Err.LastDllError
        If lResult = 0 Then
            If lLastErr = ERROR_IO_PENDING Then
                bPendingWait = True
                lResult = WaitForMultipleObjects(1, uWaitOver.hEvent, 0, Timeout)
                lLastErr = Err.LastDllError
                If lResult <> WAIT_OBJECT_0 Then
                    m_bClearPersistent = True
                    If lResult = WAIT_TIMEOUT Then
                        m_sLastError = pvInternal(ucsErrTimeoutWaitingForResponse)
                        DebugLog FUNC_NAME, "WaitCommEvent timeout, Timeout=" & Timeout & ", m_sLastError=" & m_sLastError & ", Elapsed=" & Format$(TimerEx - dblTimerEx, "0.000")
                    Else
                        m_sLastError = GetErrorDescription(lLastErr)
                        DebugLog FUNC_NAME, "WaitForMultipleObjects failure after WaitCommEvent, lResult=" & lResult & ", Timeout=" & Timeout & ", m_sLastError=" & m_sLastError & ", Elapsed=" & Format$(TimerEx - dblTimerEx, "0.000")
                    End If
                    GoTo QH
                End If
                bPendingWait = False
            Else
                m_bClearPersistent = True
                m_sLastError = Printf(pvInternal(ucsErrWaitCommEventFailed), GetErrorDescription(lLastErr))
                DebugLog FUNC_NAME, "WaitCommEvent, lResult=" & lResult & ", m_sLastError=" & m_sLastError
                GoTo QH
            End If
        End If
InLoop:
        m_lActual = 0
        lResult = ReadFile(m_hComm, m_baBuffer(0), UBound(m_baBuffer) + 1, m_lActual, uReadOver)
        lLastErr = Err.LastDllError
        If lResult = 0 Then
            If lLastErr = ERROR_IO_PENDING Then
                bPendingRead = True
                lResult = WaitForMultipleObjects(1, uReadOver.hEvent, 0, Timeout)
                lLastErr = Err.LastDllError
                If lResult <> WAIT_OBJECT_0 Then
                    m_bClearPersistent = True
                    If lResult = WAIT_TIMEOUT Then
                        m_sLastError = pvInternal(ucsErrTimeoutWaitingForResponse)
                        DebugLog FUNC_NAME, "ReadFile timeout, Timeout=" & Timeout & ", m_sLastError=" & m_sLastError & ", Elapsed=" & Format$(TimerEx - dblTimerEx, "0.000")
                    Else
                        m_sLastError = GetErrorDescription(lLastErr)
                        DebugLog FUNC_NAME, "WaitForMultipleObjects failure after ReadFile, lResult=" & lResult & ", Timeout=" & Timeout & ", m_sLastError=" & m_sLastError & ", Elapsed=" & Format$(TimerEx - dblTimerEx, "0.000")
                    End If
                    GoTo QH
                End If
                bPendingRead = False
                m_lActual = 0
                Call GetOverlappedResult(m_hComm, uReadOver, m_lActual, 0)
            ElseIf lLastErr = ERROR_HANDLE_EOF Then
                GoTo Continue
            Else
                m_bClearPersistent = True
                m_sLastError = Printf(pvInternal(ucsErrReadFileFailed), GetErrorDescription(lLastErr))
                DebugLog FUNC_NAME, "ReadFile, lResult=" & lResult & ", m_sLastError=" & m_sLastError & ", Elapsed=" & Format$(TimerEx - dblTimerEx, "0.000")
                GoTo QH
            End If
        End If
        If m_lActual > 0 Or Timeout = 0 Then
            lActual = m_lActual
            GoTo QH
        End If
        If TimerEx > dblTimerEx + Timeout / 1000# Then
            m_bClearPersistent = True
            m_sLastError = pvInternal(ucsErrTimeoutWaitingForResponse)
            DebugLog FUNC_NAME, "ReadData timeout, lRetryCount=" & lRetryCount & ", Timeout=" & Timeout & ", m_sLastError=" & m_sLastError & ", Elapsed=" & Format$(TimerEx - dblTimerEx, "0.000")
            GoTo QH
        End If
        Call Sleep(1)
Continue:
    Loop
QH:
    If IsLogDataDumpEnabled Then
        DebugLog FUNC_NAME, "Will return, lActual=" & lActual & ", lRetryCount=" & lRetryCount & ", bPendingRead=" & bPendingRead & ", bPendingWait=" & bPendingWait & ", Elapsed=" & Format$(TimerEx - dblTimerEx, "0.000")
    End If
    If lActual > 0 Then
        ReDim Result(0 To lActual - 1) As Byte
        Call CopyMemory(Result(0), m_baBuffer(0), lActual)
        '--- success
        ReadData = True
    Else
        Result = vbNullString
    End If
    If bPendingRead Or bPendingWait Then
        Call CancelIo(m_hComm)
    End If
    Call CloseHandle(uReadOver.hEvent)
    Call CloseHandle(uWaitOver.hEvent)
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
    hComm = PortWrapper.OpenComPort(At(vSplit, 0), C_Bool(At(vSplit, 2)))
    If hComm = INVALID_HANDLE_VALUE Then
        lLastErr = Err.LastDllError
        Error = Printf(pvInternal(ucsErrCreateFileFailed), GetErrorDescription(lLastErr))
        If lLastErr = ERROR_ACCESS_DENIED And lTimeout <> 0 Then
            dblTimer = TimerEx + lTimeout / 1000#
            Do While TimerEx < dblTimer
                Call Sleep(100)
                hComm = PortWrapper.OpenComPort(At(vSplit, 0), C_Bool(At(vSplit, 2)))
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
        Error = Printf(pvInternal(ucsErrSetCommTimeoutsFailed), GetErrorDescription(Err.LastDllError))
        PortWrapper.CloseComPort hComm, ClearPersistent:=True
        GoTo QH
    End If
    If BuildCommDCB("baud=" & At(vSplit, 1, DEF_SPEED) & " data=" & At(vSplit, 3, 8) & " parity=" & At(vSplit, 4, "N") & " stop=" & At(vSplit, 5, 1), uDCB) = 0 Then
        Error = Printf(pvInternal(ucsErrBuildCommDCBFailed), GetErrorDescription(Err.LastDllError))
        PortWrapper.CloseComPort hComm, ClearPersistent:=True
        GoTo QH
    End If
    If SetCommState(hComm, uDCB) = 0 Then
        Error = Printf(pvInternal(ucsErrSetCommStateFailed), GetErrorDescription(Err.LastDllError))
        PortWrapper.CloseComPort hComm, ClearPersistent:=True
        GoTo QH
    End If
    pvOpenPort = hComm
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Function pvInternal(ByVal lIdx As UcsConnectorErrorsEnum) As String
    Const FUNC_NAME     As String = "pvInternal"
    
    On Error GoTo EH
    pvInternal = At(m_vLocalizedText(ucsFscLciInternalErrors), lIdx)
    If LenB(pvInternal) = 0 Then
        pvInternal = At(Split(STR_CONNECTOR_ERRORS, "|"), lIdx)
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

#If Not ImplUseShared Then

Public Function SplitOrReindex(Expression As String, Delimiter As String) As Variant
    SplitOrReindex = Split(Expression, Delimiter)
End Function

Private Property Get At(vData As Variant, ByVal lIdx As Long, Optional sDefault As String) As String
    On Error GoTo QH
    At = sDefault
    If IsArray(vData) Then
        If lIdx < LBound(vData) Then
            '--- lIdx = -1 for last element
            lIdx = UBound(vData) + 1 + lIdx
        End If
        If LBound(vData) <= lIdx And lIdx <= UBound(vData) Then
            At = C_Str(vData(lIdx))
        End If
    End If
QH:
End Property

Private Function C_Lng(Value As Variant) As Long
    Dim vDest           As Variant
    
    If VarType(Value) = vbLong Then
        C_Lng = Value
    ElseIf VariantChangeType(vDest, Value, 0, vbLong) = 0 Then
        C_Lng = vDest
    End If
End Function

Private Function C_Str(Value As Variant) As String
    Dim vDest           As Variant
    
    If VarType(Value) = vbString Then
        C_Str = Value
    ElseIf VariantChangeType(vDest, Value, VARIANT_ALPHABOOL, vbString) = 0 Then
        C_Str = vDest
    End If
End Function

Private Function C_Bool(Value As Variant) As Boolean
    Dim vDest           As Variant
    
    If VarType(Value) = vbBoolean Then
        C_Bool = Value
    ElseIf VariantChangeType(vDest, Value, VARIANT_ALPHABOOL, vbBoolean) = 0 Then
        C_Bool = vDest
    End If
End Function

Private Property Get TimerEx() As Double
    Dim cFreq           As Currency
    Dim cValue          As Currency
    
    Call QueryPerformanceFrequency(cFreq)
    Call QueryPerformanceCounter(cValue)
    TimerEx = cValue / cFreq
End Property

Private Function GetErrorDescription(ByVal ErrorCode As Long) As String
    Dim lSize           As Long
   
    GetErrorDescription = Space$(2000)
    lSize = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0, ErrorCode, 0, GetErrorDescription, Len(GetErrorDescription), 0)
    If lSize > 2 Then
        If Mid$(GetErrorDescription, lSize - 1, 2) = vbCrLf Then
            lSize = lSize - 2
        End If
    End If
    GetErrorDescription = Left$(GetErrorDescription, lSize)
End Function

Private Function Printf(ByVal sText As String, ParamArray A() As Variant) As String
    Const LNG_PRIVATE   As Long = &HE1B6 '-- U+E000 to U+F8FF - Private Use Area (PUA)
    Dim lIdx            As Long
    
    For lIdx = UBound(A) To LBound(A) Step -1
        sText = Replace(sText, "%" & (lIdx - LBound(A) + 1), Replace(A(lIdx), "%", ChrW$(LNG_PRIVATE)))
    Next
    Printf = Replace(sText, ChrW$(LNG_PRIVATE), "%")
End Function

Private Property Get IsLogDataDumpEnabled() As Boolean

End Property

#End If

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Initialize()
    m_hComm = INVALID_HANDLE_VALUE
    m_lCommEventTimeout = DEF_COMM_EVENT_TIMEOUT
    m_lTimeout = DEF_TIMEOUT
    m_baBuffer = vbNullString
End Sub

Private Sub Class_Terminate()
    If m_hComm <> INVALID_HANDLE_VALUE Then
        Call CancelIo(m_hComm)
        PortWrapper.CloseComPort m_hComm, m_bClearPersistent
        m_hComm = INVALID_HANDLE_VALUE
    End If
End Sub

'=========================================================================
' IDeviceConnector interface
'=========================================================================

Private Function IDeviceConnector_GetLastError(Optional ErrNo As UcsConnectorErrorsEnum) As String
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

