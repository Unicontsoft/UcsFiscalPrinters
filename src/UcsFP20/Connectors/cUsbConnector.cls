VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cUsbConnector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================================================
'
' UcsFP20 (c) 2008-2026 by Unicontsoft
'
' Unicontsoft Fiscal Printers Component 2.0
'
' This project is licensed under the terms of the MIT license
' See the LICENSE file in the project root for more information
'
'=========================================================================
'
' USB connector
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "cUsbConnector"
Implements IDeviceConnector

#Const ImplUseShared = (USE_SHARED <> 0)

#If Not ImplUseShared Then

    Public Enum UcsFiscalLocalizedIndexesEnum
        ucsFscLciInternalErrors
        [_ucsFscLciMax]
    End Enum
    
    Private Const STR_CONNECTOR_ERRORS      As String = "No device info set|%1 failed: %2|Timeout waiting for response"
    Private Const vbLogEventTypeDebug       As Long = vbLogEventTypeInformation + 1
    
    Private Logger                  As Object

#End If
    
'=========================================================================
' API
'=========================================================================

'--- for CreateFile
Private Const GENERIC_READ                  As Long = &H80000000
Private Const GENERIC_WRITE                 As Long = &H40000000
Private Const OPEN_EXISTING                 As Long = 3
Private Const INVALID_HANDLE_VALUE          As Long = -1
Private Const FILE_FLAG_OVERLAPPED          As Long = &H40000000
'--- for WaitForMultipleObjects
Private Const WAIT_OBJECT_0                 As Long = 0
Private Const WAIT_TIMEOUT                  As Long = 258
'--- error codes
Private Const ERROR_IO_PENDING              As Long = 997
Private Const ERROR_HANDLE_EOF              As Long = 38
'--- for VariantChangeType
Private Const VARIANT_ALPHABOOL             As Long = 2
'--- for FormatMessage
Private Const FORMAT_MESSAGE_FROM_SYSTEM    As Long = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200
'--- for SetupDiGetClassDevs
Private Const DIGCF_PRESENT                 As Long = &H2
Private Const DIGCF_ALLCLASSES              As Long = &H4
Private Const DIGCF_DEVICEINTERFACE         As Long = &H10
'--- for SetupDiGetDeviceRegistryProperty
Private Const SPDRP_FRIENDLYNAME            As Long = &HC

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function CreateEvent Lib "kernel32" Alias "CreateEventA" (ByVal lpEventAttributes As Long, ByVal bManualReset As Long, ByVal bInitialState As Long, ByVal lpName As Long) As Long
Private Declare Function WaitForMultipleObjects Lib "kernel32" (ByVal nCount As Long, lpHandles As Long, ByVal bWaitAll As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CancelIo Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetOverlappedResult Lib "kernel32" (ByVal hFile As Long, lpOverlapped As OVERLAPPED, lpNumberOfBytesTransferred As Long, ByVal bWait As Long) As Long
Private Declare Function CM_Locate_DevNode Lib "cfgmgr32" Alias "CM_Locate_DevNodeW" (pdnDevInst As Long, ByVal pDeviceID As Long, ByVal ulFlags As Long) As Long
Private Declare Function CM_Get_Parent Lib "cfgmgr32" (pdnDevInst As Long, ByVal dnDevInst As Long, ByVal ulFlags As Long) As Long
Private Declare Function CM_Get_Device_ID Lib "cfgmgr32" Alias "CM_Get_Device_IDW" (ByVal dnDevInst As Long, ByVal Buffer As Long, ByVal BufferLen As Long, ByVal ulFlags As Long) As Long
Private Declare Function SetupDiGetClassDevs Lib "setupapi" Alias "SetupDiGetClassDevsW" (ByVal ClassGuid As Long, ByVal Enumerator As Long, ByVal hwndParent As Long, ByVal Flags As Long) As Long
Private Declare Function SetupDiEnumDeviceInfo Lib "setupapi" (ByVal DeviceInfoSet As Long, ByVal MemberIndex As Long, ByRef DeviceInfoData As SP_DEVINFO_DATA) As Long
Private Declare Function SetupDiGetDeviceRegistryProperty Lib "setupapi" Alias "SetupDiGetDeviceRegistryPropertyW" (ByVal DeviceInfoSet As Long, ByRef DeviceInfoData As SP_DEVINFO_DATA, ByVal Proprty As Long, ByRef PropertyRegDataType As Long, ByVal PropertyBuffer As Long, ByVal PropertyBufferSize As Long, ByRef RequiredSize As Long) As Long
Private Declare Function SetupDiDestroyDeviceInfoList Lib "setupapi" (ByVal DeviceInfoSet As Long) As Long
Private Declare Function SetupDiEnumDeviceInterfaces Lib "setupapi" (ByVal DeviceInfoSet As Long, ByVal DeviceInfoData As Long, ByVal InterfaceClassGuid As Long, ByVal MemberIndex As Long, ByRef DeviceInterfaceData As SP_DEVICE_INTERFACE_DATA) As Long
Private Declare Function SetupDiGetDeviceInterfaceDetail Lib "setupapi" Alias "SetupDiGetDeviceInterfaceDetailW" (ByVal DeviceInfoSet As Long, ByRef DeviceInterfaceData As SP_DEVICE_INTERFACE_DATA, ByVal pDetailData As Long, ByVal DetailDataSize As Long, ByRef RequiredSize As Long, ByVal DeviceInfoData As Long) As Long
Private Declare Function SysReAllocString Lib "oleaut32" (ByVal pBSTR As Long, Optional ByVal pszStrPtr As Long) As Long
#If Not ImplUseShared Then
    Private Declare Function VariantChangeType Lib "oleaut32" (vDest As Variant, Src As Variant, ByVal wFlags As Integer, ByVal vt As Long) As Long
    Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
    Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
    Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Args As Any) As Long
#End If

Private Type OVERLAPPED
    Internal                    As Long
    InternalHigh                As Long
    Offset                      As Long
    OffsetHigh                  As Long
    hEvent                      As Long
End Type

Private Type SP_DEVINFO_DATA
    cbSize                      As Long
    ClassGuid(0 To 3)           As Long
    DevInst                     As Long
    Reserved                    As Long
End Type

Private Type SP_DEVICE_INTERFACE_DATA
    cbSize                      As Long
    InterfaceClassGuid(0 To 3)  As Long
    Flags                       As Long
    Reserved                    As Long
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const DEF_TIMEOUT               As Long = 3000

Private m_hFile                 As Long
Private m_lTimeout              As Long
Private m_sDevicePath           As String
Private m_sLastError            As String
Private m_vLocalizedText(0 To [_ucsFscLciMax] - 1) As Variant
Private m_bClearPersistent      As Boolean
'--- overlapped I/O buffers
Private m_lWritten              As Long
Private m_lActual               As Long
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
    Device = m_sDevicePath
End Property

Property Get LocalizedText(ByVal eIdx As UcsFiscalLocalizedIndexesEnum) As String
    LocalizedText = Join(m_vLocalizedText(eIdx), "|")
End Property

Property Let LocalizedText(ByVal eIdx As UcsFiscalLocalizedIndexesEnum, sValue As String)
    m_vLocalizedText(eIdx) = SplitOrReindex(sValue, "|")
End Property

Property Get hFile() As Long
    hFile = m_hFile
End Property

Property Let hFile(ByVal lValue As Long)
    m_hFile = lValue
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function Init(Device As String, Optional ByVal Timeout As Long) As Boolean
    Const FUNC_NAME     As String = "Init"
    Dim dblTimerEx      As Double
    Dim lRetryCount     As Long
    
    On Error GoTo EH
    DebugLog FUNC_NAME, "Device=" & Device & ", Timeout=" & Timeout
    Call CloseHandle(m_hFile)
    m_hFile = INVALID_HANDLE_VALUE
    m_lTimeout = IIf(Timeout <= 0, DEF_TIMEOUT, Timeout)
    m_sDevicePath = pvGetUsbDevicePath(Device)
    dblTimerEx = TimerEx
    Do
        lRetryCount = lRetryCount + 1
        m_sLastError = vbNullString
        m_hFile = pvOpenFile(m_sDevicePath, m_sLastError)
        If LenB(m_sLastError) = 0 Then
            Exit Do
        End If
        DebugLog FUNC_NAME, "pvOpenFile failed, lRetryCount=" & lRetryCount & ", m_sLastError=" & m_sLastError
        If TimerEx > dblTimerEx + m_lTimeout / 1000 Then
            GoTo QH
        End If
        Call Sleep(100)
    Loop
    '--- success
    Init = True
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function WriteData(Data() As Byte, Optional ByVal Timeout As Long, Optional IgnoreError As Boolean) As Boolean
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
RetryWrite:
    m_lWritten = 0
    lResult = WriteFile(m_hFile, Data(lOffset), UBound(Data) + 1 - lOffset, m_lWritten, uWriteOver)
    lLastErr = Err.LastDllError
    If lResult = 0 Then
        If lLastErr = ERROR_IO_PENDING Then
            bPendingWrite = True
            lResult = WaitForMultipleObjects(1, uWriteOver.hEvent, 0, Timeout)
            lLastErr = Err.LastDllError
            If lResult <> WAIT_OBJECT_0 Then
                If lResult = WAIT_TIMEOUT Then
                    m_sLastError = pvInternal(ucsErrTimeoutWaitingForResponse)
                    If Not IgnoreError Then
                        DebugLog FUNC_NAME, "WriteFile timeout, Timeout=" & Timeout & ", m_sLastError=" & m_sLastError & ", Elapsed=" & Format$(TimerEx - dblTimerEx, "0.000")
                        m_bClearPersistent = True
                    End If
                Else
                    m_sLastError = GetErrorDescription(lLastErr)
                    If Not IgnoreError Then
                        DebugLog FUNC_NAME, "WaitForMultipleObjects failure after WriteFile, lResult=" & lResult & ", Timeout=" & Timeout & ", m_sLastError=" & m_sLastError & ", Elapsed=" & Format$(TimerEx - dblTimerEx, "0.000")
                        m_bClearPersistent = True
                    End If
                End If
                GoTo QH
            End If
            bPendingWrite = False
            m_lWritten = 0
            Call GetOverlappedResult(m_hFile, uWriteOver, m_lWritten, 0)
        Else
            m_sLastError = Printf(pvInternal(ucsErrApiCallFailed), "WriteFile", GetErrorDescription(lLastErr))
            If Not IgnoreError Then
                DebugLog FUNC_NAME, "WriteFile, lResult=" & lResult & ", Err.LastDllError=" & lLastErr
                m_bClearPersistent = True
            End If
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
        Call CancelIo(m_hFile)
        Call GetOverlappedResult(m_hFile, uWriteOver, m_lWritten, 1)
    End If
    Call CloseHandle(uWriteOver.hEvent)
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Public Function ReadData(Result() As Byte, Optional ByVal Timeout As Long, Optional IgnoreError As Boolean) As Boolean
    Const FUNC_NAME     As String = "ReadData"
    Dim dblTimerEx      As Double
    Dim uReadOver       As OVERLAPPED
    Dim lResult         As Long
    Dim lLastErr        As Long
    Dim lRetryCount     As Long
    Dim lActual         As Long
    Dim bPendingRead    As Boolean
    
    On Error GoTo EH
    dblTimerEx = TimerEx
    uReadOver.hEvent = CreateEvent(0, 1, 0, 0)
    If UBound(m_baBuffer) < 0 Then
        ReDim m_baBuffer(0 To 65535) As Byte
    End If
    Do
        m_lActual = 0
        lResult = ReadFile(m_hFile, m_baBuffer(0), UBound(m_baBuffer) + 1, m_lActual, uReadOver)
        lLastErr = Err.LastDllError
        If lResult = 0 Then
            If lLastErr = ERROR_IO_PENDING Then
                bPendingRead = True
                lResult = WaitForMultipleObjects(1, uReadOver.hEvent, 0, Timeout)
                lLastErr = Err.LastDllError
                If lResult <> WAIT_OBJECT_0 Then
                    If lResult = WAIT_TIMEOUT Then
                        m_sLastError = pvInternal(ucsErrTimeoutWaitingForResponse)
                        If Not IgnoreError Then
                            DebugLog FUNC_NAME, "ReadFile timeout, Timeout=" & Timeout & ", m_sLastError=" & m_sLastError & ", Elapsed=" & Format$(TimerEx - dblTimerEx, "0.000")
                            m_bClearPersistent = True
                        End If
                    Else
                        m_sLastError = GetErrorDescription(lLastErr)
                        If Not IgnoreError Then
                            DebugLog FUNC_NAME, "WaitForMultipleObjects failure after ReadFile, lResult=" & lResult & ", Timeout=" & Timeout & ", m_sLastError=" & m_sLastError & ", Elapsed=" & Format$(TimerEx - dblTimerEx, "0.000")
                            m_bClearPersistent = True
                        End If
                    End If
                    GoTo QH
                End If
                bPendingRead = False
                m_lActual = 0
                Call GetOverlappedResult(m_hFile, uReadOver, m_lActual, 0)
            ElseIf lLastErr = ERROR_HANDLE_EOF Then
                GoTo Continue
            Else
                m_sLastError = Printf(pvInternal(ucsErrApiCallFailed), "ReadFile", GetErrorDescription(lLastErr))
                If Not IgnoreError Then
                    DebugLog FUNC_NAME, "ReadFile, lResult=" & lResult & ", m_sLastError=" & m_sLastError & ", Elapsed=" & Format$(TimerEx - dblTimerEx, "0.000")
                    m_bClearPersistent = True
                End If
                GoTo QH
            End If
        End If
        If m_lActual > 0 Or Timeout = 0 Then
            lActual = m_lActual
            GoTo QH
        End If
        If TimerEx > dblTimerEx + Timeout / 1000# Then
            m_sLastError = pvInternal(ucsErrTimeoutWaitingForResponse)
            If Not IgnoreError Then
                DebugLog FUNC_NAME, "ReadData timeout, lRetryCount=" & lRetryCount & ", Timeout=" & Timeout & ", m_sLastError=" & m_sLastError & ", Elapsed=" & Format$(TimerEx - dblTimerEx, "0.000")
                m_bClearPersistent = True
            End If
            GoTo QH
        End If
        Call Sleep(1)
Continue:
        lRetryCount = lRetryCount + 1
    Loop
QH:
    If IsLogDataDumpEnabled Then
        DebugLog FUNC_NAME, "Will return, lActual=" & lActual & ", lRetryCount=" & lRetryCount & ", bPendingRead=" & bPendingRead & ", Elapsed=" & Format$(TimerEx - dblTimerEx, "0.000")
    End If
    If lActual > 0 Then
        ReDim Result(0 To lActual - 1) As Byte
        Call CopyMemory(Result(0), m_baBuffer(0), lActual)
        '--- success
        ReadData = True
    Else
        Result = vbNullString
    End If
    If bPendingRead Then
        Call CancelIo(m_hFile)
        If bPendingRead Then
            Call GetOverlappedResult(m_hFile, uReadOver, m_lActual, 1)
        End If
    End If
    Call CloseHandle(uReadOver.hEvent)
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Public Function GetLastError() As String
    GetLastError = m_sLastError
End Function

'= private ===============================================================

Private Function pvOpenFile(sDevice As String, Optional Error As String) As Long
    Const FUNC_NAME     As String = "pvOpenFile"
    Dim hFile           As Long
    Dim lApiError       As Long
    
    On Error GoTo EH
    If LenB(sDevice) = 0 Then
        Error = pvInternal(ucsErrNoDeviceInfoSet)
        GoTo QH
    End If
    hFile = CreateFile(sDevice, GENERIC_READ Or GENERIC_WRITE, 0, 0, OPEN_EXISTING, FILE_FLAG_OVERLAPPED, 0)
    If hFile = INVALID_HANDLE_VALUE Then
        lApiError = Err.LastDllError
        Error = Printf(pvInternal(ucsErrApiCallFailed), "CreateFile", GetErrorDescription(lApiError))
        GoTo QH
    End If
    pvOpenFile = hFile
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Function pvGetUsbDevicePath(sDevice As String) As String
    Dim vElem           As Variant

    If Left$(sDevice, 2) = "\\" Then
        pvGetUsbDevicePath = sDevice
    End If
    If LenB(pvGetUsbDevicePath) = 0 Then
        '--- try mapping "USB001" port name
        pvGetUsbDevicePath = pvGetPortDevicePath(sDevice)
    End If
    If LenB(pvGetUsbDevicePath) = 0 Then
        '--- try mapping "HPRT Alpha (203 dpi)" printer name
        For Each vElem In pvGetPrinterDeviceIDs(sDevice)
            Do While LenB(vElem) <> 0
                pvGetUsbDevicePath = pvGetInterfaceDevicePath(vElem)
                If LenB(pvGetUsbDevicePath) <> 0 Then
                    Exit For
                End If
                vElem = pvGetParentDeviceID(vElem)
            Loop
        Next
    End If
End Function

Private Function pvGetPortDevicePath(ByVal sPortName As String) As String
    Const HKLM          As Long = &H80000002
    Dim sKey            As String
    
    sKey = "SYSTEM\CurrentControlSet\Control\Print\Monitors\USB Monitor\Ports\" & sPortName
    With GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
        .GetStringValue HKLM, sKey, "Device Path", pvGetPortDevicePath
    End With
End Function

Private Function pvGetPrinterDeviceIDs(ByVal sPrinterName As String) As Collection
    Dim hDev            As Long
    Dim uDevData        As SP_DEVINFO_DATA
    Dim sBuffer         As String
    Dim lIdx            As Long
    
    Set pvGetPrinterDeviceIDs = New Collection
    hDev = SetupDiGetClassDevs(0, 0, 0, DIGCF_ALLCLASSES)
    If hDev = -1 Then
        GoTo QH
    End If
    For lIdx = 0 To 9999
        uDevData.cbSize = LenB(uDevData)
        If SetupDiEnumDeviceInfo(hDev, lIdx, uDevData) = 0 Then
            Exit For
        End If
        sBuffer = String$(1024, 0)
        If SetupDiGetDeviceRegistryProperty(hDev, uDevData, SPDRP_FRIENDLYNAME, 0, StrPtr(sBuffer), LenB(sBuffer) - 2, 0) <> 0 Then
            sBuffer = Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1)
            If StrComp(sBuffer, sPrinterName, vbTextCompare) = 0 Then
                sBuffer = String$(1024, 0)
                If CM_Get_Device_ID(uDevData.DevInst, StrPtr(sBuffer), Len(sBuffer) - 1, 0) = 0 Then
                    pvGetPrinterDeviceIDs.Add Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1)
                End If
            End If
        End If
    Next
QH:
    If hDev <> 0 Then
        Call SetupDiDestroyDeviceInfoList(hDev)
    End If
End Function

Private Function pvGetParentDeviceID(ByVal sDeviceID As String) As String
    Dim hDevInst        As Long
    Dim hParentInst     As Long
    Dim sBuffer         As String
    
    If LenB(sDeviceID) = 0 Then
        GoTo QH
    End If
    If CM_Locate_DevNode(hDevInst, StrPtr(sDeviceID), 0) <> 0 Then
        GoTo QH
    End If
    If CM_Get_Parent(hParentInst, hDevInst, 0) <> 0 Then
        GoTo QH
    End If
    sBuffer = String$(1024, 0)
    If CM_Get_Device_ID(hParentInst, StrPtr(sBuffer), Len(sBuffer) - 1, 0) <> 0 Then
        GoTo QH
    End If
    pvGetParentDeviceID = Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1)
QH:
End Function

Private Function pvGetInterfaceDevicePath(ByVal sDeviceID As String) As String
    Dim uGuid(0 To 3)   As Long
    Dim hDev            As Long
    Dim uDevData        As SP_DEVINFO_DATA
    Dim uIfData         As SP_DEVICE_INTERFACE_DATA
    Dim lSize           As Long
    Dim baBuffer()      As Byte
    
    If LenB(sDeviceID) = 0 Then
        GoTo QH
    End If
    '--- GUID_DEVINTERFACE_USBPRINT = {28D78FAD-5A12-11D1-AE5B-0000F803A8C2}
    uGuid(0) = &H28D78FAD: uGuid(1) = &H11D15A12: uGuid(2) = &H5BAE&: uGuid(3) = &HC2A803F8
    hDev = SetupDiGetClassDevs(VarPtr(uGuid(0)), StrPtr(sDeviceID), 0, DIGCF_PRESENT Or DIGCF_DEVICEINTERFACE)
    If hDev = -1 Then
        GoTo QH
    End If
    uDevData.cbSize = LenB(uDevData)
    If SetupDiEnumDeviceInfo(hDev, 0, uDevData) = 0 Then
        GoTo QH
    End If
    uIfData.cbSize = LenB(uIfData)
    If SetupDiEnumDeviceInterfaces(hDev, VarPtr(uDevData), VarPtr(uGuid(0)), 0, uIfData) = 0 Then
        GoTo QH
    End If
    Call SetupDiGetDeviceInterfaceDetail(hDev, uIfData, 0, 0, lSize, 0)
    If lSize = 0 Then
        GoTo QH
    End If
    ReDim baBuffer(0 To lSize - 1) As Byte
    baBuffer(0) = 6
    If SetupDiGetDeviceInterfaceDetail(hDev, uIfData, VarPtr(baBuffer(0)), UBound(baBuffer) + 1, 0, 0) = 0 Then
        GoTo QH
    End If
    '--- success
    pvGetInterfaceDevicePath = pvToString(VarPtr(baBuffer(4)))
QH:
    If hDev <> 0 Then
        Call SetupDiDestroyDeviceInfoList(hDev)
    End If
End Function

Private Function pvToString(ByVal lPtr As Long) As String
    Call SysReAllocString(VarPtr(pvToString), lPtr)
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

Private Function SplitOrReindex(Expression As String, Delimiter As String) As Variant
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

Private Function C_Str(Value As Variant) As String
    Dim vDest           As Variant
    
    If VarType(Value) = vbString Then
        C_Str = Value
    ElseIf VariantChangeType(vDest, Value, VARIANT_ALPHABOOL, vbString) = 0 Then
        C_Str = vDest
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
    Const LNG_PRIVATE   As Long = &HE1B6& '-- U+E000 to U+F8FF - Private Use Area (PUA)
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
    m_hFile = INVALID_HANDLE_VALUE
    m_lTimeout = DEF_TIMEOUT
    m_baBuffer = vbNullString
End Sub

Private Sub Class_Terminate()
    If m_hFile <> INVALID_HANDLE_VALUE Then
        Call CloseHandle(m_hFile)
        m_hFile = INVALID_HANDLE_VALUE
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

Private Function IDeviceConnector_ReadData(Result() As Byte, Optional ByVal Timeout As Long, Optional IgnoreError As Boolean) As Boolean
    IDeviceConnector_ReadData = ReadData(Result, Timeout, IgnoreError)
End Function

Private Function IDeviceConnector_WriteData(Data() As Byte, Optional ByVal Timeout As Long, Optional IgnoreError As Boolean) As Boolean
    IDeviceConnector_WriteData = WriteData(Data, Timeout, IgnoreError)
End Function

