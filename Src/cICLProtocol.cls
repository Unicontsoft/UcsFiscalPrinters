VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cICLProtocol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=========================================================================
' $Header: /UcsFiscalPrinter/Src/cICLProtocol.cls 18    14.02.11 18:14 Wqw $
'
'   Unicontsoft Fiscal Printers Project
'   Copyright (c) 2008-2011 Unicontsoft
'
'   Protocol za fiscalni printer na Datecs i Daisy po ICL protocol
'
' $Log: /UcsFiscalPrinter/Src/cICLProtocol.cls $
' 
' 18    14.02.11 18:14 Wqw
' REF: constanti za internal msgs
'
' 17    26.11.10 13:20 Wqw
' REF: impl poddryjka daisy FP
'
' 16    25.11.10 13:49 Wqw
' ADD: Consts. REF: err handling na SendCommand prowerqwa i za memory
' error
'
' 15    17.05.10 16:16 Wqw
' REF: debug log
'
' 14    2.02.09 13:52 Wqw
' REF: pvReadData probwa da chete pyrwo posle wait-wa za comm event
'
' 13    19.01.09 17:36 Wqw
' REF: exception podtiska samo pri print error
'
' 12    17.12.08 18:27 Wqw
' REF: DebugLog prowerqwa dali log file-a syshtestwuwa
'
' 11    2.12.08 0:15 Wqw
' REF: logging
'
' 10    1.12.08 23:02 Wqw
' REF: logging
'
' 9     1.12.08 16:14 Wqw
' REF: impl debug log
'
' 8     1.12.08 0:42 Wqw
' REF: UcsIclCommandsEnum, STR_ERRORS_ONLY
'
' 7     11.11.08 17:20 Wqw
' REF: Private pvGetFlagsText -> Friend frGetFlagsText
'
' 6     7.11.08 13:02 Wqw
' REF: enum ports pyrwo enumerira wsichki serial ports
'
' 5     4.11.08 14:19 Wqw
' REF: polzwa overlapped I/O pri dostyp do com port-a
'
' 4     30.10.08 22:49 Wqw
' REF: STR_STATUSES
'
' 3     14.10.08 19:09 Wqw
' REF: impl logo i door open
'
' 2     19.06.08 10:48 Wqw
' REF: public class weche
'
' 1     18.06.08 16:45 Wqw
' Initial implementation
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "cICLProtocol"

'=========================================================================
' Public events
'=========================================================================

Event CommandComplete(ByVal lCmd As Long, sData As String, sResult As String)

'=========================================================================
' Public enums
'=========================================================================

Public Enum UcsIclStatusBitsEnum
    ucsStbSyntaxError = 2 ^ 0
    ucsStbInvalidCommand = 2 ^ 1
    ucsStbClockNotSet = 2 ^ 2
    ucsStbDeviceError = 2 ^ 4
    ucsStbPrintingError = 2 ^ 5
    ucsStbSumOverflow = 2 ^ (8 + 0)
    ucsStbInvalidFiscalMode = 2 ^ (8 + 1)
    ucsStbRamCleared = 2 ^ (8 + 2)
    ucsStbRamCorrupted = 2 ^ (8 + 4)
    ucsStbOutOfPaper = 2 ^ (16 + 0)
    ucsStbFiscalPrinting = 2 ^ (16 + 3)
    ucsStbNonFiscalPrinting = 2 ^ (16 + 5)
End Enum

Public Enum UcsIclDipSwitchesEnum
    ucsDipCenterHeaderFooter = 2 ^ 0
    ucsDipPrintHeader = 2 ^ 1
    ucsDipSwitch_1_3 = 2 ^ 2
    ucsDipSwitch_1_4 = 2 ^ 3
    ucsDipZeroMemory = 2 ^ 4
    ucsDipTransparentDisplay = 2 ^ 5
    ucsDipShowToDisplay = 2 ^ 6
End Enum

Public Enum UcsIclFiscalMemoryEnum
    ucsFmeWriteError = 2 ^ 0
    ucsFmeNoMemoryModule = 2 ^ 2
    ucsFmeLowMemory = 2 ^ 3
    ucsFmeOutOfMemory = 2 ^ 4
    ucsFmeMemoryError = 2 ^ 5
    ucsFmeReadOnly = 2 ^ (8 + 0)
    ucsFmeFormatted = 2 ^ (8 + 1)
    ucsFmeFiscalMode = 2 ^ (8 + 3)
    ucsFmeTaxRatesSet = 2 ^ (8 + 4)
    ucsFmeIdentNoSet = 2 ^ (8 + 5)
End Enum

Public Enum UcsIclCommandsEnum
    '--- init
    ucsFpcInitHeaderFooter = 43
    ucsFpcInitDateTime = 61
    ucsFpcInitInvoiceNo = 66
    ucsFpcInitDecimals = 83
    ucsFpcInitPaymentType = 85
    ucsFpcInitDepartment = 87
    ucsFpcInitTaxRates = 96
    ucsFpcInitBulstat = 98
    ucsFpcInitOperatorPassword = 101
    ucsFpcInitOperatorName = 102
    ucsFpcInitOperatorReset = 104
    ucsFpcInitItem = 107
    ucsFpcInitLogo = 115
    ucsFpcInitOpenDrawer = 106
    ucsFpcInitEcTape = 120
    ucsFpcInitDaisyText = 149
    '--- transaction
    ucsFpcNonFiscalOpen = 38
    ucsFpcNonFiscalClose = 39
    ucsFpcNonFiscalText = 42
    ucsFpcFiscalOpen = 48
    ucsFpcFiscalSell = 49
    ucsFpcFiscalPayment = 53
    ucsFpcFiscalText = 54
    ucsFpcFiscalClose = 56
    ucsFpcFiscalCgInfo = 57
    ucsFpcFiscalCancel = 60
    ucsFpcAdminCashDebitCredit = 70
    ucsFpcPrintDuplicate = 109
    '--- info
    ucsFpcInfoDateTime = 62
    ucsFpcInfoLastFiscalReceipt = 64
    ucsFpcInfoStatus = 74
    ucsFpcInfoTransaction = 76
    ucsFpcInfoDepartment = 88
    ucsFpcInfoDiagnostics = 90
    ucsFpcInfoTaxRates = 97
    ucsFpcInfoBulstat = 99
    ucsFpcInfoSums = 103
    ucsFpcInfoOperator = 112
    ucsFpcInfoLastReceipt = 113
    ucsFpcInfoDaisyConsts = 128
    '--- print
    ucsFpcPrintDailyReport = 69
    ucsFpcPrintDiagnostics = 71
    ucsFpcPrintReportByNumberDetailed = 73
    ucsFpcPrintReportByDateShort = 79
    ucsFpcPrintReportByDateDetailed = 94
    ucsFpcPrintReportByNumberShort = 95
    ucsFpcPrintReportByOperators = 105
    ucsFpcPrintDailyReportItems = 108
    ucsFpcPrintDailyReportDepartments = 117
    ucsFpcPrintDailyReportItemsDepartments = 118
    ucsFpcPrintReportTaxRates = 50
End Enum

'=========================================================================
' API
'=========================================================================

'--- for CreateFile
Private Const GENERIC_READ                  As Long = &H80000000
Private Const GENERIC_WRITE                 As Long = &H40000000
Private Const OPEN_EXISTING                 As Long = 3
Private Const FILE_FLAG_OVERLAPPED          As Long = &H40000000
Private Const INVALID_HANDLE_VALUE          As Long = -1
'--- for WaitCommEvent
Private Const EV_RXCHAR                     As Long = &H1                '  Any Character received
'--- for WaitForMultipleObjects
Private Const WAIT_OBJECT_0                 As Long = 0
Private Const WAIT_TIMEOUT                  As Long = 258
'--- error codes
Private Const ERROR_IO_PENDING              As Long = 997

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function SetCommTimeouts Lib "kernel32" (ByVal hFile As Long, lpCommTimeouts As COMMTIMEOUTS) As Long
Private Declare Function BuildCommDCB Lib "kernel32" Alias "BuildCommDCBA" (ByVal lpDef As String, lpDCB As DCB) As Long
Private Declare Function SetCommState Lib "kernel32" (ByVal hCommDev As Long, lpDCB As DCB) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function SetCommMask Lib "kernel32" (ByVal hFile As Long, ByVal dwEvtMask As Long) As Long
Private Declare Function WaitCommEvent Lib "kernel32" (ByVal hFile As Long, lpEvtMask As Long, lpOverlapped As OVERLAPPED) As Long
Private Declare Function CreateEvent Lib "kernel32" Alias "CreateEventA" (ByVal lpEventAttributes As Long, ByVal bManualReset As Long, ByVal bInitialState As Long, ByVal lpName As Long) As Long
Private Declare Function WaitForMultipleObjects Lib "kernel32" (ByVal nCount As Long, lpHandles As Long, ByVal bWaitAll As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CancelIo Lib "kernel32" (ByVal hFile As Long) As Long

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

Private Const STR_INTERNAL              As String = "No device info set|CreateFile failed: |SetCommTimeouts failed: |BuildCommDCB failed: |SetCommState failed: |WriteFile failure: |Timeout waiting for response|ReadFile failed: |Invalid message format or checksum (NAK)|Already connected|Not connected|Invalid response length|Response missing postamble|Response missing preamble|WaitCommEvent failed: "
Private Const STR_STATUSES              As String = "Syntax error|Invalid command|Clock not set|3|Feed error|5|6|7|Overflow|Command not allowed|Memory reset|3|Memory corrupt|5|6|7|Out of paper|1|2|Fiscal printing|4|Nonfiscal printing|6|7"
Private Const STR_ERRORS_ONLY           As String = "Syntax error|Invalid command|2|3|Feed error|5|6|7|1|Command not allowed|Memory reset|3|Memory corrupt|5|6|7|Out of paper|1|2|3|4|5|6|7"
Private Const STR_DIP_SWITCHES          As String = "Center header and footer|Enable header|Sw1.3|Sw1.4|Zero memory|Transparent display|No display|7"
Private Const STR_MEMORY                As String = "Write error|1|No memory module|Memory low|Out of memory|Fiscal memory error|6|7|Read only|Memory formatted|2|Fiscal mode|Tax rates set|Tax number set|6|7"
Private Const DEF_TIMEOUT               As Long = 3000

Private m_sDevice                   As String
Private m_hComm                     As Long
Private m_lTimeout                  As Long
Private m_lSequence                 As Long
Private m_eStatus                   As UcsIclStatusBitsEnum
Private m_eDip                      As UcsIclDipSwitchesEnum
Private m_eMemory                   As UcsIclFiscalMemoryEnum
Private m_sLastError                As String
Private m_bExceptions               As Boolean
Private m_sLocalizedText(0 To 5)    As String

Private Enum UcsInternalErrors
    ucsErrNoDeviceInfoSet
    ucsErrCreateFileFailed
    ucsErrSetCommTimeoutsFailed
    ucsErrBuildCommDCBFailed
    ucsErrSetCommStateFailed
    ucsErrWriteFileFailure
    ucsErrTimeoutWaitingForResponse
    ucsErrReadFileFailed
    ucsErrInvalidMessageFormatOrChecksum
    ucsErrAlreadyConnected
    ucsErrNotConnected
    ucsErrInvalidResponseLength
    ucsErrResponseMissingPostamble
    ucsErrResponseMissingPreamble
    ucsErrWaitCommEventFailed
End Enum

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunc As String)
    Debug.Print MODULE_NAME & "." & sFunc & ": " & Error
    OutputDebugLog MODULE_NAME, sFunc, "Run-time error: " & Error
End Sub

Private Sub DebugLog(sFunc As String, sText As String)
    OutputDebugLog MODULE_NAME, sFunc, sText
End Sub

'=========================================================================
' Properties
'=========================================================================

Property Get Timeout() As Long
    Timeout = m_lTimeout
End Property

Property Let Timeout(ByVal lValue As Long)
    m_lTimeout = lValue
End Property

Property Get Status(ByVal eValue As UcsIclStatusBitsEnum) As Boolean
    Status = (m_eStatus And eValue) <> 0
End Property

Property Get StatusText() As String
    StatusText = frGetFlagsText(m_eStatus, LocalizedText(ucsFscLciPrinterStatuses))
End Property

Property Get ErrorText() As String
    ErrorText = frGetFlagsText(m_eStatus, LocalizedText(ucsFscLciPrinterErrors))
End Property

Property Get Dip(ByVal eValue As UcsIclDipSwitchesEnum) As Boolean
    Dip = (m_eDip And eValue) <> 0
End Property

Property Get DipText() As String
    DipText = frGetFlagsText(m_eDip, LocalizedText(ucsFscLciDipSwitchesTexts))
End Property

Property Get Memory(ByVal eValue As UcsIclFiscalMemoryEnum) As Boolean
    Memory = (m_eMemory And eValue) <> 0
End Property

Property Get MemoryText() As String
    MemoryText = frGetFlagsText(m_eMemory, LocalizedText(ucsFscLciMemoryTexts))
End Property

Property Get LastError() As String
    LastError = m_sLastError
End Property

Property Get Exceptions() As Boolean
    Exceptions = m_bExceptions
End Property

Property Let Exceptions(ByVal bValue As Boolean)
    m_bExceptions = bValue
End Property

Property Get Device() As String
    Device = m_sDevice
End Property

Property Get IsConnected() As Boolean
    IsConnected = (m_hComm <> INVALID_HANDLE_VALUE)
End Property

Property Get LocalizedText(ByVal eIdx As UcsFiscalLocalizedIndexesEnum) As String
    LocalizedText = m_sLocalizedText(eIdx)
    If LenB(LocalizedText) = 0 Then
        Select Case eIdx
        Case ucsFscLciInternalErrors
            LocalizedText = STR_INTERNAL
        Case ucsFscLciPrinterStatuses
            LocalizedText = STR_STATUSES
        Case ucsFscLciPrinterErrors
            LocalizedText = STR_ERRORS_ONLY
        Case ucsFscLciDipSwitchesTexts
            LocalizedText = STR_DIP_SWITCHES
        Case ucsFscLciMemoryTexts
            LocalizedText = STR_MEMORY
        End Select
    End If
End Property

Property Let LocalizedText(ByVal eIdx As UcsFiscalLocalizedIndexesEnum, sValue As String)
    m_sLocalizedText(eIdx) = sValue
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function EnumPorts() As Variant
    Dim hComm           As Long
    Dim sError          As String
    Dim sResult         As String
    Dim vPort           As Variant
    Dim vRet            As Variant
    Dim lCount          As Long
    
    ReDim vRet(0 To 255) As Variant
    For Each vPort In EnumSerialPorts
        sError = vbNullString
        hComm = pvOpenPort(vPort & ",9600,8,N,1", sError)
        If hComm <> INVALID_HANDLE_VALUE And LenB(sError) = 0 Then
            If pvWriteData(hComm, ucsFpcInfoStatus, &H20) Then
                sError = vbNullString
                sResult = pvReadData(hComm, 500, sError)
                If LenB(sResult) <> 0 And LenB(sError) = 0 Then
                    If Asc(sResult) = 1 Then
                        vRet(lCount) = vPort
                        lCount = lCount + 1
                    End If
                End If
            End If
            Call CloseHandle(hComm)
        End If
    Next
    If lCount = 0 Then
        EnumPorts = Split("")
    Else
        ReDim Preserve vRet(0 To lCount - 1) As Variant
        EnumPorts = vRet
    End If
End Function

Public Function Init(sDevice As String, Optional ByVal lTimeout As Long) As Boolean
    Const FUNC_NAME     As String = "Init"
    Dim vSplit          As Variant
    
    DebugLog FUNC_NAME, "sDevice=" & sDevice & "; lTimeout=" & lTimeout
    m_sLastError = vbNullString
    If m_hComm <> INVALID_HANDLE_VALUE Then
        Call CloseHandle(m_hComm)
        m_hComm = INVALID_HANDLE_VALUE
    End If
    If lTimeout <= 0 Then
        m_lTimeout = DEF_TIMEOUT
    Else
        m_lTimeout = lTimeout
    End If
    vSplit = Split(sDevice, ",")
    m_sDevice = Trim(At(vSplit, 0, "COM1")) & "," & C_Lng(At(vSplit, 1, "9600")) & "," & C_Lng(At(vSplit, 2, "8")) & "," & _
        IIf(UCase(At(vSplit, 3, "N")) = "Y", "Y", "N") & "," & C_Lng(At(vSplit, 4, "1"))
    m_hComm = pvOpenPort(m_sDevice, m_sLastError)
    If LenB(m_sLastError) <> 0 Then
        GoTo QH
    End If
    '--- success
    Init = True
QH:
    If LenB(m_sLastError) <> 0 Then
        Err.Raise vbObjectError, , m_sLastError
    End If
End Function

Public Function Connect() As Boolean
    Const FUNC_NAME     As String = "Connect"
    
    DebugLog FUNC_NAME, "m_hComm=" & m_hComm
    m_sLastError = vbNullString
    If m_hComm <> INVALID_HANDLE_VALUE Then
        m_sLastError = pvInternal(ucsErrAlreadyConnected)
        GoTo QH
    End If
    m_hComm = pvOpenPort(m_sDevice, m_sLastError)
    '--- success
    Connect = True
QH:
    If LenB(m_sLastError) <> 0 Then
        Err.Raise vbObjectError, , m_sLastError
    End If
End Function

Public Function Disconnect() As Boolean
    Const FUNC_NAME     As String = "Disconnect"
    
    DebugLog FUNC_NAME, "m_hComm=" & m_hComm
    m_sLastError = vbNullString
    If m_hComm = INVALID_HANDLE_VALUE Then
        m_sLastError = pvInternal(ucsErrNotConnected)
        GoTo QH
    End If
    Call CloseHandle(m_hComm)
    m_hComm = INVALID_HANDLE_VALUE
    '--- success
    Disconnect = True
QH:
    If LenB(m_sLastError) <> 0 Then
        Err.Raise vbObjectError, , m_sLastError
    End If
End Function

Public Function SendCommand(ByVal eCmd As UcsIclCommandsEnum, Optional sData As String) As String
    Const FUNC_NAME     As String = "SendCommand"
    Dim sText           As String
    Dim lIdx            As Long
    Dim lSize           As Long
        
    DebugLog FUNC_NAME, "eCmd=" & eCmd & ", sData=" & sData
    m_sLastError = vbNullString
    If Not pvWriteData(m_hComm, eCmd, m_lSequence, sData, m_sLastError) Then
        GoTo QH
    End If
    '--- increment sequence
    m_lSequence = m_lSequence + 1
    If m_lSequence > &H7F Then
        m_lSequence = &H20
    End If
    '--- check result
    sText = pvReadData(m_hComm, m_lTimeout, m_sLastError)
    If LenB(m_sLastError) = 0 Then
        If Left(sText, 1) = Chr$(1) Then
            lIdx = InStr(sText, Chr$(4))
            If lIdx > 2 Then
                lSize = Asc(Mid$(sText, 2, 1)) - &H20 - 11
                '--- check message length
                If lIdx - 5 = lSize Then
                    pvSetStatus Mid$(sText, lIdx + 1, 6)
                    SendCommand = pvTranscode(Mid$(sText, 5, lSize), False)
                    GoTo QH
                Else
                    m_sLastError = pvInternal(ucsErrInvalidResponseLength)
                End If
            Else
                m_sLastError = pvInternal(ucsErrResponseMissingPostamble)
            End If
        Else
            m_sLastError = pvInternal(ucsErrResponseMissingPreamble)
        End If
    End If
    pvSetStatus vbNullString '--- clear
QH:
    DebugLog FUNC_NAME, "RetVal=" & SendCommand & IIf(LenB(m_sLastError) <> 0, ", m_sLastError=" & m_sLastError, vbNullString)
    RaiseEvent CommandComplete(eCmd, sData, SendCommand)
    If LenB(m_sLastError) <> 0 Then
        Err.Raise vbObjectError, , m_sLastError
    End If
    If Exceptions Then
        If Status(ucsStbPrintingError) Then
            DebugLog FUNC_NAME, "ErrorText=" & ErrorText
            Err.Raise vbObjectError, , ErrorText
        End If
        If Memory(ucsFmeMemoryError) Then
            DebugLog FUNC_NAME, "MemoryText=" & ErrorText
            Err.Raise vbObjectError, , MemoryText
        End If
    End If
End Function

Public Function ReadData(Optional ByVal lTimeout As Long) As String
    Const FUNC_NAME     As String = "ReadData"
    
    m_sLastError = vbNullString
    ReadData = pvReadData(m_hComm, IIf(lTimeout > 0, lTimeout, m_lTimeout), m_sLastError)
    DebugLog FUNC_NAME, "RetVal=" & ReadData
    If LenB(m_sLastError) <> 0 Then
        DebugLog FUNC_NAME, "m_sLastError=" & m_sLastError
        Err.Raise vbObjectError, , m_sLastError
    End If
End Function

'= private ===============================================================

Private Function pvOpenPort(sDevice As String, Optional sError As String) As Long
    Const FUNC_NAME     As String = "pvOpenPort"
    Dim vSplit          As Variant
    Dim uTimeout        As COMMTIMEOUTS
    Dim uDCB            As DCB
    
    On Error GoTo EH
    vSplit = Split(sDevice, ",")
    If LenB(At(vSplit, 0)) = 0 Then
        sError = pvInternal(ucsErrNoDeviceInfoSet)
        GoTo QH
    End If
    pvOpenPort = CreateFile("\\.\" & At(vSplit, 0), GENERIC_READ Or GENERIC_WRITE, 0, 0, OPEN_EXISTING, FILE_FLAG_OVERLAPPED, 0)
    If pvOpenPort = INVALID_HANDLE_VALUE Then
        sError = pvInternal(ucsErrCreateFileFailed) & GetApiErr(Err.LastDllError)
        GoTo QH
    End If
    With uTimeout
        .ReadIntervalTimeout = -1
        .ReadTotalTimeoutConstant = 0
        .ReadTotalTimeoutMultiplier = 0
        .WriteTotalTimeoutConstant = 500
        .WriteTotalTimeoutMultiplier = 1
    End With
    If SetCommTimeouts(pvOpenPort, uTimeout) = 0 Then
        sError = pvInternal(ucsErrSetCommTimeoutsFailed) & GetApiErr(Err.LastDllError)
        GoTo QH
    End If
    If BuildCommDCB("baud=" & At(vSplit, 1, 9600) & " data=" & At(vSplit, 2, 8) & " parity=" & At(vSplit, 3, "N") & " stop=" & At(vSplit, 4, 1), uDCB) = 0 Then
        sError = pvInternal(ucsErrBuildCommDCBFailed) & GetApiErr(Err.LastDllError)
        GoTo QH
    End If
    If SetCommState(pvOpenPort, uDCB) = 0 Then
        sError = pvInternal(ucsErrSetCommStateFailed) & GetApiErr(Err.LastDllError)
        GoTo QH
    End If
QH:
    pvSetStatus vbNullString '--- clear
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function pvWriteData(ByVal hComm As Long, ByVal eCmd As UcsIclCommandsEnum, ByVal lSequence As Long, Optional sData As String, Optional sError As String) As Boolean
    Const FUNC_NAME     As String = "pvWriteData"
    Dim sText           As String
    Dim baText()        As Byte
    Dim lIdx            As Long
    Dim lSum            As Long
    Dim lWritten        As Long
    Dim uOver           As OVERLAPPED
    Dim lSize           As Long
    Dim lResult         As Long
    
    On Error GoTo EH
    uOver.hEvent = CreateEvent(0, 1, 0, 0)
    '--- prepare command and transcode cyrillic
    sText = pvTranscode(sData, True)
    sText = Chr$(1) & Chr$(&H20 + 4 + Len(sText)) & Chr$(lSequence) & Chr$(eCmd) & sText & Chr$(5)
    '--- calc checksum
    baText = StrConv(sText, vbFromUnicode)
    For lIdx = 1 To UBound(baText)
        lSum = lSum + baText(lIdx)
    Next
    '--- append checksum and terminator
    sText = sText & Chr$((lSum \ &H1000 Mod &H10) + &H30) & Chr$((lSum \ &H100 Mod &H10) + &H30) & Chr$((lSum \ &H10 Mod &H10) + &H30) & Chr$((lSum \ &H1 Mod &H10) + &H30) & Chr$(3)
    lSize = Len(sText)
    baText = StrConv(sText, vbFromUnicode)
    lResult = WriteFile(hComm, baText(0), lSize, lWritten, VarPtr(uOver))
    If lResult = 0 Then
        If Err.LastDllError <> ERROR_IO_PENDING Then
            sError = pvInternal(ucsErrWriteFileFailure) & GetApiErr(Err.LastDllError)
            GoTo QH
        End If
    End If
    lResult = WaitForMultipleObjects(1, uOver.hEvent, 0, m_lTimeout)
    If lResult = WAIT_OBJECT_0 Then
        pvWriteData = True
    Else
        Call CancelIo(hComm)
    End If
QH:
    Call CloseHandle(uOver.hEvent)
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function pvReadData(ByVal hComm As Long, ByVal lTimeout As Long, Optional sError As String) As String
    Const FUNC_NAME     As String = "pvReadData"
    Dim sBuffer         As String
    Dim dwRead          As Long
    Dim lResult         As Long
    Dim lMask           As Long
    Dim uOver           As OVERLAPPED
    
    On Error GoTo EH
    Call SetCommMask(hComm, EV_RXCHAR)
    uOver.hEvent = CreateEvent(0, 1, 0, 0)
    GoTo InLoop
    Do
        lResult = WaitCommEvent(hComm, lMask, uOver)
        If lResult = 0 Then
            If Err.LastDllError <> ERROR_IO_PENDING Then
                DebugLog FUNC_NAME, "WaitCommEvent, Err.LastDllError=" & Err.LastDllError
                sError = pvInternal(ucsErrWaitCommEventFailed) & GetApiErr(Err.LastDllError)
                GoTo QH
            End If
            lResult = WaitForMultipleObjects(1, uOver.hEvent, 0, lTimeout)
        Else
            DebugLog FUNC_NAME, "WaitCommEvent, lResult=" & lResult
            lResult = WAIT_OBJECT_0
        End If
        If lResult = WAIT_OBJECT_0 And (lMask And EV_RXCHAR) <> 0 Then
InLoop:
            sBuffer = String(1000, 0)
            lResult = ReadFile(hComm, ByVal StrPtr(sBuffer), Len(sBuffer), dwRead, VarPtr(uOver))
            If lResult = 0 Then
                If Err.LastDllError <> ERROR_IO_PENDING Then
                    DebugLog FUNC_NAME, "ReadFile, Err.LastDllError=" & Err.LastDllError
                    sError = pvInternal(ucsErrReadFileFailed) & GetApiErr(Err.LastDllError)
                    GoTo QH
                End If
                lResult = WaitForMultipleObjects(1, uOver.hEvent, 0, lTimeout)
            Else
                lResult = WAIT_OBJECT_0
            End If
            If lResult = WAIT_OBJECT_0 Then
                sBuffer = Left(StrConv(sBuffer, vbUnicode), dwRead)
                '--- check for NAK (&H15)
                If InStr(sBuffer, Chr$(&H15)) > 0 Then
                    DebugLog FUNC_NAME, "NAK, sBuffer=" & sBuffer
                    sError = pvInternal(ucsErrInvalidMessageFormatOrChecksum)
                    GoTo QH
                End If
                '--- remove SYN (&H16)
                pvReadData = pvReadData & Replace(sBuffer, Chr$(&H16), vbNullString)
            End If
        End If
        If lResult <> WAIT_OBJECT_0 Then
            If lResult = WAIT_TIMEOUT Then
                sError = pvInternal(ucsErrTimeoutWaitingForResponse)
            End If
            Call CancelIo(hComm)
            Exit Do
        End If
    Loop While Right(pvReadData, 1) <> Chr$(&H3) '--- until Terminator (&H3) found
QH:
    Call CloseHandle(uOver.hEvent)
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function pvTranscode(sText As String, ByVal bToPrinter As Boolean) As String
    Const FUNC_NAME     As String = "pvTranscode"
    Dim lIdx            As Long
    Dim lChar           As Long
    
    On Error GoTo EH
    pvTranscode = sText
    For lIdx = 1 To Len(pvTranscode)
        lChar = Asc(Mid$(pvTranscode, lIdx, 1))
        If lChar < &H7F Then
            '--- do nothing
        ElseIf lChar >= &HC0 And lChar <= &HFF And bToPrinter Then
            Mid$(pvTranscode, lIdx, 1) = Chr$(lChar - &H40)
        ElseIf lChar >= &H80 And lChar < &HC0 And Not bToPrinter Then
            Mid$(pvTranscode, lIdx, 1) = Chr$(lChar + &H40)
        Else
            Mid$(pvTranscode, lIdx, 1) = "?"
        End If
    Next
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Sub pvSetStatus(sStatus As String)
    Const FUNC_NAME     As String = "pvSetStatus"
    Dim baStatus()      As Byte
    
    On Error GoTo EH
    baStatus = StrConv(sStatus, vbFromUnicode)
    If UBound(baStatus) >= 2 Then
        Call CopyMemory(m_eStatus, baStatus(0), 3)
    Else
        m_eStatus = 0
    End If
    If UBound(baStatus) >= 3 Then
        Call CopyMemory(m_eDip, baStatus(3), 1)
    Else
        m_eDip = 0
    End If
    If UBound(baStatus) >= 5 Then
        Call CopyMemory(m_eMemory, baStatus(4), 2)
    Else
        m_eMemory = 0
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Friend Function frGetFlagsText(ByVal lFlags As Long, sTexts As String) As String
    Const FUNC_NAME     As String = "frGetFlagsText"
    Dim lIdx            As Long
    Dim vSplit          As Variant
    
    On Error GoTo EH
    vSplit = Split(sTexts, "|")
    For lIdx = 0 To UBound(vSplit)
        If (lFlags And (2 ^ lIdx)) <> 0 Then
            If Len(vSplit(lIdx)) > 1 Then
                If LenB(frGetFlagsText) <> 0 Then
                    frGetFlagsText = frGetFlagsText & ", "
                End If
                frGetFlagsText = frGetFlagsText & vSplit(lIdx)
            End If
        End If
    Next
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function pvInternal(ByVal lIdx As UcsInternalErrors) As String
    pvInternal = At(Split(LocalizedText(ucsFscLciInternalErrors), "|"), lIdx, At(Split(STR_INTERNAL, "|"), lIdx))
End Function

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Initialize()
    m_hComm = INVALID_HANDLE_VALUE
    m_lTimeout = 3000
    m_lSequence = &H20
    m_bExceptions = True
End Sub

Private Sub Class_Terminate()
    If m_hComm <> INVALID_HANDLE_VALUE Then
        Call CloseHandle(m_hComm)
        m_hComm = INVALID_HANDLE_VALUE
    End If
End Sub
