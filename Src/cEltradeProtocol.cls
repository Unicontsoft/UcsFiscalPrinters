VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEltradeProtocol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=========================================================================
' $Header: /UcsFiscalPrinter/Src/cEltradeProtocol.cls 25    19.10.12 0:36 Wqw $
'
'   Unicontsoft Fiscal Printers Project
'   Copyright (c) 2008-2012 Unicontsoft
'
'   Protocol for Eltrade cash registers
'
' $Log: /UcsFiscalPrinter/Src/cEltradeProtocol.cls $
' 
' 25    19.10.12 0:36 Wqw
' REF: wrong public scope of private functions
'
' 24    9.10.12 15:20 Wqw
' REF: parms of IDeviceProtocol_CashDebitCredit
'
' 23    5.10.12 14:20 Wqw
' ADD: Function CashDebitCredit
'
' 22    6.08.12 18:39 Wqw
' REF: impl AutodetectDevices
'
' 21    23.03.12 15:28 Wqw
' ADD: GetTotalsByVatGroups. REF: cancel receipt before printing reports
'
' 20    8.12.11 15:49 Wqw
' REF: params of GetDefaultPassword
'
' 19    13.07.11 18:14 Wqw
' REF: while loop in convert extra rows
'
' 18    4.07.11 15:48 Wqw
' REF: err handling
'
' 17    17.06.11 11:35 Wqw
' REF: impl edge case -- 30th PLU w/ discount
'
' 16    10.05.11 15:14 Wqw
' REF: type mismatch in RunZReport
'
' 15    4.05.11 19:48 Wqw
' ADD: Function RunPeriodReport
'
' 14    4.05.11 15:04 Wqw
' REF: invoice number off by one error
'
' 13    11.04.11 14:51 Wqw
' REF: default timeout, param timeout na send command, retry w wait
' status
'
' 12    29.03.11 14:59 Wqw
' REF: impl retry na get status w pvWaitStatus
'
' 11    18.03.11 16:51 Wqw
' REF: impl extra checks for first PLU programming
'
' 10    8.03.11 13:07 Wqw
' REF: impl out-of-bounds discounts
'
' 9     7.03.11 19:22 Wqw
' REF: w WaitDevice ako wyrne status busy izchakwa 100 ms
'
' 8     7.03.11 18:24 Wqw
' REF: default time-out
'
' 7     23.02.11 17:09 Wqw
' DEL: dead const
'
' 6     22.02.11 17:38 Wqw
' REF: error handling
'
' 5     22.02.11 15:19 Wqw
' REF: impl debug logging
'
' 4     22.02.11 13:51 Wqw
' ADD: Function GetDeviceProtocol
'
' 3     22.02.11 10:05 Wqw
' ADD: Function GetDeviceModel
'
' 2     21.02.11 13:47 Wqw
' ADD: Function CancelReceipt. REF: impl localization support, rename bon
' to receipt, SendInput is public
'
' 1     14.02.11 18:13 Wqw
' Initial implementation
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "cEltradeProtocol"
Implements IDeviceProtocol

'=========================================================================
' Public events
'=========================================================================

Event CommandComplete(ByVal lCmd As Long, sData As String, sResult As String)

'=========================================================================
' Public enums
'=========================================================================

Public Enum UcsEltradeCommandsEnum
    '--- transaction
    ucsEltCmdKeyboardInput = &H2D
    ucsEltCmdCommentLine = &H2E
    '--- info
    ucsEltCmdInfoFpNumber = &H10
    ucsEltCmdInfoHeadingLines = &H11
    ucsEltCmdInfoYesNoParams = &H13
    ucsEltCmdInfoFiscalNumber = &H14
    ucsEltCmdInfoBulstat = &H15
    ucsEltCmdInfoItem = &H16
    ucsEltCmdInfoDepartment = &H17
    ucsEltCmdInfoVatGroups = &H19
    ucsEltCmdInfoOperator = &H1A
    ucsEltCmdInfoPaymentTypes = &H1B
    ucsEltCmdInfoKeyFunctions = &H1D
    ucsEltCmdInfoTurnoverVatGroups = &H22
    ucsEltCmdInfoTurnoverOperator = &H23
    ucsEltCmdInfoStatus = &H2C
    ucsEltCmdInfoEcrParams = &H50
    ucsEltCmdInfoDateTime = &H52
    ucsEltCmdInfoInvoiceNo = &H56
    ucsEltCmdInfoInvoiceData = &H58
    ucsEltCmdInfoDeviceStatus = &H70
    '--- init
    ucsEltCmdInitKeylock = &H39
    ucsEltCmdInitHeadingLines = &H41
    ucsEltCmdInitYesNoParams = &H43
    ucsEltCmdInitVatGroups = &H44
    ucsEltCmdInitItem = &H46
    ucsEltCmdInitDepartment = &H47
    ucsEltCmdInitOperator = &H4A
    ucsEltCmdInitPaymentTypes = &H4B
    ucsEltCmdInitKeyFunctions = &H4D
    ucsEltCmdInitDateTime = &H53
    ucsEltCmdInitInvoiceNo = &H57
    ucsEltCmdInitInvoiceData = &H59
'    '--- print
'    ucsEltCmdPrintXReport = &H24
'    ucsEltCmdPrintZReportByItem = &H30
'    ucsEltCmdPrintZReportByItemNo = &H31
'    ucsEltCmdPrintZReportByDep = &H32
'    ucsEltCmdPrintZReportByVatGrp = &H33
'    ucsEltCmdPrintZReportByOper = &H34
    ucsEltCmdPrintZReport = &H35
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
'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function SetCommMask Lib "kernel32" (ByVal hFile As Long, ByVal dwEvtMask As Long) As Long
Private Declare Function WaitCommEvent Lib "kernel32" (ByVal hFile As Long, lpEvtMask As Long, lpOverlapped As OVERLAPPED) As Long
Private Declare Function CreateEvent Lib "kernel32" Alias "CreateEventA" (ByVal lpEventAttributes As Long, ByVal bManualReset As Long, ByVal bInitialState As Long, ByVal lpName As Long) As Long
Private Declare Function WaitForMultipleObjects Lib "kernel32" (ByVal nCount As Long, lpHandles As Long, ByVal bWaitAll As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CancelIo Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function FlushFileBuffers Lib "kernel32" (ByVal hFile As Long) As Long

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

Private Const STR_INTERNAL_ERRORS       As String = "No device info set|CreateFile failed: |SetCommTimeouts failed: |BuildCommDCB failed: |SetCommState failed: |WriteFile failure: |Timeout waiting for response|ReadFile failed: |Already connected|Not connected|Invalid response header|WaitCommEvent failed: |No receipt started|Out of paper|Invalid 'Command' parameter|Invalid command for duplicate receipt|Printing error|Fatal error|Cannot cancel previous receipt|Error in line %1: %2"
Private Const STR_RECEIPT_TEXTS         As String = "TOTAL|SUBTOTAL|Sales %1|unit price|quantity|value*%1|CHANGE|total|VAT*%1|netto|Invoice #|Seller:|   * * * * * * * *|Received by:|Tax # |VAT # |%1 items|Resp:|Recv:|SURCHARGE|DISCOUNT"
Private Const DEF_TIMEOUT               As Long = 3000
Private Const LNG_DELAY                 As Long = 100
Private Const MAX_RECEIPT_ROWS          As Long = 30
Private Const MAX_PAYMENT_LEN           As Long = 24
Private Const MIN_DISCOUNT              As Double = -99.99
Private Const MAX_DISCOUNT              As Double = 99.99

Private m_lTimeout                  As Long
Private m_sDevice                   As String
Private m_hComm                     As Long
Private m_lCashDesk                 As Long
Private m_sLastError                As String
Private m_sLocalizedText(0 To [_ucsFscLciMax] - 1) As String
Private m_uRow()                    As UcsRowData
Private m_lRowCount                 As Long
Private m_sHeadingLines             As String
Private m_sPaymentTypes             As String
Private m_lRowChars                 As Long
Private m_bFiscalPaymentExecuted    As Boolean
Private m_sCommandLog               As String

Private Enum UcsInternalErrorsEnum
    ucsErrNoDeviceInfoSet
    ucsErrCreateFileFailed
    ucsErrSetCommTimeoutsFailed
    ucsErrBuildCommDCBFailed
    ucsErrSetCommStateFailed
    ucsErrWriteFileFailure
    ucsErrTimeoutWaitingForResponse
    ucsErrReadFileFailed
    ucsErrAlreadyConnected
    ucsErrNotConnected
    ucsErrInvalidResponseHeader
    ucsErrWaitCommEventFailed
    ucsErrNoReceiptStarted
    ucsErrOutOfPaper
    ucsErrInvalidCommandParameter
    ucsErrInvalidCommandForDuplicateReceipt
    ucsErrPrintingError
    ucsErrFatalError
    ucsErrCannotCancelReceipt
    ucsErrErrorInLine
End Enum

Private Enum UcsReceiptTextsEnum
    ucsTxtTotal
    ucsTxtSubTotal
    ucsTxtPluSell
    ucsTxtPluPrice
    ucsTxtPluQuantity
    ucsTxtPluSum
    ucsTxtPmtChange
    ucsTxtPmtTotal
    ucsTxtPmtVatGroup
    ucsTxtPmtNeto
    ucsTxtHeaderInvoiceNo
    ucsTxtFooterSeller
    ucsTxtFooterDivider
    ucsTxtFooterReceiver
    ucsTxtFooterTaxNo
    ucsTxtFooterVatNo
    ucsTxtFooterPluCount
    ucsTxtInvMol
    ucsTxtInvReceiver
    ucsTxtSurcharge
    ucsTxtDiscount
End Enum

Private Enum UcsRowTypeEnum
    ucsRowInit = 1
    ucsRowPlu
    ucsRowLine
    ucsRowDiscount
    ucsRowPayment
End Enum

Private Type UcsRowData
    RowType             As UcsRowTypeEnum
    InitReceiptType     As UcsFiscalReceiptTypeEnum
    InitPLUNumbering    As Boolean
    InitOperatorName    As String
    InitInvData         As Variant
    PluNo               As Long
    PluName             As String
    PluPrice            As Double
    PluQuantity         As Double
    PluVatGroup         As Long
    LineText            As String
    LineCommand         As String
    LineWordWrap        As Boolean
    DiscType            As UcsFiscalDiscountTypeEnum
    DiscValue           As Double
    PmtType             As UcsFiscalPaymentTypeEnum
    PmtName             As String
    PmtAmount           As Double
    PmtRate             As Double
    PrintRowType        As UcsFiscalReceiptTypeEnum
End Type

Private Type UcsContext
    Keys                As String
    Status              As String
    ItemChars           As Long
    PluNo               As Long
    LastRow             As UcsRowData
    GrpTotal(1 To 8)    As Double
    Paid                As Double
    PluCount            As Long
    Finished            As Boolean
End Type

Private Enum UcsCharsEnum
    ucsChrD7 = &H7
    ucsChrPlus = &HA
    ucsChrMinus = &HB
    ucsChrRFN = &HB
    ucsChrPLU = &HF
    ucsChrX = &H1A
    ucsChrVoid = &H1B
    ucsChrClear = &H1C
    ucsChrDot = &H1E
    ucsChrSubtotal = &H1F
    ucsChrTotal = &H20
    ucsChrPy1 = &HA
    ucsChrPy2 = &H20 ' &HB
    ucsChrPy3 = &H20 ' &HC
End Enum

Private Enum UcsLastCommandEnum
    ucsLcdOutOfReceipt = 0
    ucsLcdItem = 5
    ucsLcdVoid = 6
    ucsLcdPayment = 7
    ucsLcdPlus = &H45
    ucsLcdMinus = &H85
End Enum

Private Enum UcsKeylockPosition
    ucsKlcEcrUnlock = 0
    ucsKlcRegistration = 1
    ucsKlcXReport = 2
    ucsKlcZReport = 4
End Enum

Private Enum UcsInvDataIndex
    ucsInvNumber
    ucsInvTaxNumber
    ucsInvVatNumber
    ucsInvCompany
    ucsInvCity
    ucsInvAddress
    ucsInvMol
    ucsInvReceiver
End Enum

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunc As String)
    m_sLastError = Err.Description
    Debug.Print MODULE_NAME & "." & sFunc & ": " & Err.Description
    OutputDebugLog MODULE_NAME, sFunc & "(" & Erl & ")", "Run-time error: " & Err.Description
End Sub

Private Sub RaiseError(sFunc As String)
    m_sLastError = Err.Description
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

Property Get IsConnected() As Boolean
    IsConnected = (m_hComm <> INVALID_HANDLE_VALUE)
End Property

Property Get LocalizedText(ByVal eIdx As UcsFiscalLocalizedIndexesEnum) As String
    LocalizedText = m_sLocalizedText(eIdx)
    If LenB(LocalizedText) = 0 Then
        Select Case eIdx
        Case ucsFscLciInternalErrors
            LocalizedText = STR_INTERNAL_ERRORS
        Case ucsFscLciReceiptTexts
            LocalizedText = STR_RECEIPT_TEXTS
        End Select
    End If
End Property

Property Let LocalizedText(ByVal eIdx As UcsFiscalLocalizedIndexesEnum, sValue As String)
    m_sLocalizedText(eIdx) = sValue
End Property

Property Get LastError() As String
    LastError = m_sLastError
End Property

Property Get Device() As String
    Device = m_sDevice
End Property

Private Property Get pvRowChars() As Long
    If m_lRowChars = 0 Then
        '--- get chars per row
        m_lRowChars = pvPeek(SendCommand(ucsEltCmdInfoEcrParams), 3, 1)
        If m_lRowChars = 0 Then
            m_lRowChars = Len(SendCommand(ucsEltCmdInfoHeadingLines)) / 8 - 1
        End If
    End If
    pvRowChars = m_lRowChars
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function AutodetectDevices(vPorts As Variant) As Variant
    Const FUNC_NAME     As String = "AutodetectDevices"
    Const STR_SPEEDS    As String = "9600" ' |38400|115200
    Dim vRetVal         As Variant
    Dim vSpeed          As Variant
    Dim lIdx            As Long
    Dim sError          As String
    Dim hComm           As Long
    Dim sProtocol       As String
    Dim sModel          As String
    
    On Error GoTo EH
    vRetVal = vPorts
    For Each vSpeed In Split(STR_SPEEDS, "|")
        For lIdx = 0 To UBound(vPorts)
            If LenB(At(vPorts, lIdx)) <> 0 Then
                sError = vbNullString
                hComm = pvOpenPort(At(vPorts, lIdx) & "," & vSpeed, sError:=sError)
                If hComm <> INVALID_HANDLE_VALUE And LenB(sError) = 0 Then
                    If LenB(pvSendCommand(hComm, 10, ucsEltCmdInfoStatus)) = 0 Then
                        '--- do nothing
                    ElseIf pvGetDeviceInfo(hComm, 100, sProtocol, sModel) Then
                        vRetVal(lIdx) = Array(At(vPorts, lIdx), vSpeed, sProtocol, sModel)
                    End If
                End If
                Call CloseHandle(hComm)
            End If
        Next
    Next
    AutodetectDevices = vRetVal
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function Init( _
            sDevice As String, _
            Optional ByVal lTimeout As Long, _
            Optional ByVal lCashDesk As Long) As Boolean
    Const FUNC_NAME     As String = "Init"
    Dim vSplit          As Variant
    
    On Error GoTo EH
    DebugLog FUNC_NAME, "sDevice=" & sDevice & ", lTimeout=" & lTimeout
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
    m_lCashDesk = lCashDesk
    vSplit = Split(sDevice, ",")
    m_sDevice = Trim$(At(vSplit, 0, "COM1")) & "," & C_Lng(At(vSplit, 1, "9600")) & "," & C_Lng(At(vSplit, 2, "8")) & "," & _
        IIf(UCase$(At(vSplit, 3, "N")) = "Y", "Y", "N") & "," & C_Lng(At(vSplit, 4, "1"))
    m_hComm = pvOpenPort(m_sDevice, m_sLastError)
    If LenB(m_sLastError) <> 0 Then
        DebugLog FUNC_NAME, "m_hComm=" & m_hComm & ", m_sLastError=" & m_sLastError
        Err.Raise vbObjectError, , m_sLastError
    End If
    '--- success
    Init = True
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function Connect() As Boolean
    Const FUNC_NAME     As String = "Connect"
    
    On Error GoTo EH
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
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function Disconnect() As Boolean
    Const FUNC_NAME     As String = "Disconnect"
    
    On Error GoTo EH
    DebugLog FUNC_NAME, "m_hComm=" & m_hComm
    m_sLastError = vbNullString
    If m_hComm = INVALID_HANDLE_VALUE Then
        m_sLastError = pvInternal(ucsErrNotConnected)
        Err.Raise vbObjectError, , m_sLastError
    End If
    Call CloseHandle(m_hComm)
    m_hComm = INVALID_HANDLE_VALUE
    '--- cleanup cache
    m_sHeadingLines = vbNullString
    m_sPaymentTypes = vbNullString
    m_lRowChars = 0
    '--- success
    Disconnect = True
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function SendCommand( _
            ByVal eCmd As UcsEltradeCommandsEnum, _
            Optional sData As String, _
            Optional eRetCmd As UcsEltradeCommandsEnum, _
            Optional ByVal lTimeout As Long) As String
    Const FUNC_NAME     As String = "SendCommand"
    
    On Error GoTo EH
    m_sLastError = vbNullString
    SendCommand = pvSendCommand(m_hComm, Znl(lTimeout, m_lTimeout), eCmd, sData, eRetCmd, m_sLastError)
    '--- append internal log
    m_sCommandLog = m_sCommandLog & pvDumpHex(Chr$(eCmd)) & IIf(LenB(sData) <> 0, "<-" & pvDumpHex(sData), vbNullString) & IIf(LenB(SendCommand) <> 0, "->" & pvDumpHex(SendCommand), vbNullString) & vbCrLf
    If LenB(m_sLastError) <> 0 Then
        m_sCommandLog = m_sCommandLog & m_sLastError & vbCrLf
    End If
    '--- raise event & error
    RaiseEvent CommandComplete(eCmd, sData, SendCommand)
    If LenB(m_sLastError) <> 0 Then
        Err.Raise vbObjectError, , m_sLastError
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function SendInput(sText As String) As String
    Const FUNC_NAME     As String = "SendInput"
    Dim lChunk          As Long
    Dim lIdx            As Long
    Dim sData           As String
    
    On Error GoTo EH
    For lChunk = 0 To 100
        If Len(sText) <= lChunk * 30 Then
            Exit For
        End If
        sData = Mid$(sText, 1 + lChunk * 30, 30)
        For lIdx = 1 To LimitLong(m_lTimeout \ LNG_DELAY, 1)
            SendInput = SendCommand(ucsEltCmdKeyboardInput, Chr$(Len(sData)) & sData)
            If pvPeek(SendInput, 0, 2) <> 0 Then
                Exit For
            End If
            WaitDevice LNG_DELAY
        Next
    Next
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function WaitDevice(Optional ByVal lWait As Long) As Boolean
    Const FUNC_NAME     As String = "WaitDevice"
    Dim sResult         As String
    Dim eRetCmd         As UcsEltradeCommandsEnum
    
    On Error GoTo EH
    Do
        If lWait > 0 Then
            Call Sleep(lWait)
        End If
        lWait = 100
        sResult = pvSendCommand(m_hComm, m_lTimeout, ucsEltCmdInfoDeviceStatus, vbNullString, eRetCmd, m_sLastError)
        If LenB(m_sLastError) <> 0 Then
            Err.Raise vbObjectError, , m_sLastError
        End If
        If eRetCmd = ucsEltCmdInfoDeviceStatus - 1 Then
            Select Case Asc(sResult & Chr$(1))
            Case 1
                Err.Raise vbObjectError, , pvInternal(ucsErrOutOfPaper)
            Case 2
                Err.Raise vbObjectError, , pvInternal(ucsErrPrintingError)
            Case Else
                Err.Raise vbObjectError, , pvInternal(ucsErrFatalError)
            End Select
        End If
    Loop While eRetCmd <> ucsEltCmdInfoDeviceStatus
    '--- success
    WaitDevice = True
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function StartReceipt( _
            ByVal ReceiptType As UcsFiscalReceiptTypeEnum, _
            ByVal PLUNumbering As Boolean, _
            OperatorName As String, _
            Optional InvNumber As String, _
            Optional InvTaxNumber As String, _
            Optional InvVatNumber As String, _
            Optional InvCompany As String, _
            Optional InvCity As String, _
            Optional InvAddress As String, _
            Optional InvMol As String, _
            Optional InvReceiver As String) As Boolean
    Const FUNC_NAME     As String = "StartReceipt"
    Dim sCity           As String
    Dim sAddress        As String
    Dim sMol            As String
    Dim sReceiver       As String
    
    On Error GoTo EH
    ReDim m_uRow(0 To 10) As UcsRowData
    m_lRowCount = 0
    m_bFiscalPaymentExecuted = False
    With m_uRow(pvAddRow())
        .RowType = ucsRowInit
        .InitReceiptType = LimitLong(ReceiptType, 1, [_ucsFscRetMax] - 1)
        .InitPLUNumbering = PLUNumbering
        .InitOperatorName = OperatorName
        SplitCgAddress Trim$(InvCity) & vbCrLf & Trim$(InvAddress), sCity, sAddress, pvRowChars - 2
        If .InitReceiptType = ucsFscRetInvoiceFiscal Or .InitReceiptType = ucsFscRetInvoiceNonfiscal Then
            If LenB(InvMol) <> 0 Then
                sMol = pvText(ucsTxtInvMol) & InvMol
            End If
            If LenB(InvReceiver) <> 0 Then
                sReceiver = pvText(ucsTxtInvReceiver) & InvReceiver
            End If
        End If
        .InitInvData = Array(InvNumber, InvTaxNumber, InvVatNumber, InvCompany, sCity, sAddress, sMol, sReceiver)
    End With
    '--- success
    StartReceipt = True
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function AddPLU( _
            ByVal Numb As Long, _
            Name As String, _
            Price As Double, _
            Optional Quantity As Double = 1, _
            Optional VatGroup As Long = 2, _
            Optional ByVal BeforeIndex As Long) As Boolean
    Const FUNC_NAME     As String = "AddPLU"
    Dim uRow            As UcsRowData
    Dim bNegative       As Boolean
    
    On Error GoTo EH
    '--- sanity check
    If m_lRowCount = 0 Then
        Err.Raise vbObjectError, , pvInternal(ucsErrNoReceiptStarted)
    End If
    With uRow
        .RowType = ucsRowPlu
        .PluNo = Numb
        .PluName = RTrim$(Name)
        bNegative = (Round(Price, 2) * Round(Quantity, 3) < 0)
        .PluPrice = IIf(bNegative, -1, 1) * Round(Abs(Price), 2)
        .PluQuantity = Round(Abs(Quantity), 3)
        .PluVatGroup = LimitLong(VatGroup, 1, 8)
        .PrintRowType = m_uRow(0).InitReceiptType
    End With
    pvInsertRow BeforeIndex, uRow
    '--- success
    AddPLU = True
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function AddLine( _
            Line As String, _
            Optional Command As String, _
            Optional ByVal WordWrap As Boolean, _
            Optional ByVal BeforeIndex As Long) As Boolean
    Const FUNC_NAME     As String = "AddLine"
    Dim uRow            As UcsRowData
    
    On Error GoTo EH
    '--- sanity check
    If m_lRowCount = 0 Then
        Err.Raise vbObjectError, , pvInternal(ucsErrNoReceiptStarted)
    End If
    If InStr("DBWIDTH|CONDENSED", UCase$(Command)) = 0 Then
        Err.Raise vbObjectError, , pvInternal(ucsErrInvalidCommandParameter)
    End If
    With uRow
        .RowType = ucsRowLine
        .LineText = RTrim$(Line)
        .LineCommand = Command
        .LineWordWrap = WordWrap
    End With
    If BeforeIndex = 0 Then
        '--- find first payment row
        For BeforeIndex = 0 To m_lRowCount - 1
            If m_uRow(BeforeIndex).RowType = ucsRowPayment Then
                Exit For
            End If
        Next
    End If
    pvInsertRow BeforeIndex, uRow
    '--- success
    AddLine = True
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function AddDiscount( _
            ByVal DiscType As UcsFiscalDiscountTypeEnum, _
            Value As Double) As Boolean
    Const FUNC_NAME     As String = "AddDiscount"
    
    On Error GoTo EH
    '--- sanity check
    If m_lRowCount = 0 Then
        Err.Raise vbObjectError, , pvInternal(ucsErrNoReceiptStarted)
    End If
    Select Case DiscType
    Case ucsFscDstTotal
        AddPLU 1, IIf(Value > 0, pvText(ucsTxtSurcharge), pvText(ucsTxtDiscount)), Value
    Case Else
        With m_uRow(pvAddRow())
            .RowType = ucsRowDiscount
            .DiscType = LimitLong(DiscType, 1, [_ucsFscDstMax] - 1)
            .DiscValue = Round(Value, 2)
            .PrintRowType = m_uRow(0).InitReceiptType
        End With
    End Select
    '--- success
    AddDiscount = True
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function AddPayment( _
            ByVal Number As UcsFiscalPaymentTypeEnum, _
            Name As String, _
            Amount As Double, _
            Rate As Double) As Boolean
    Const FUNC_NAME     As String = "AddPayment"
    
    On Error GoTo EH
    '--- sanity check
    If m_lRowCount = 0 Then
        Err.Raise vbObjectError, , pvInternal(ucsErrNoReceiptStarted)
    End If
    With m_uRow(pvAddRow())
        .RowType = ucsRowPayment
        '--- note: only cash and card payments supported yet
        .PmtType = IIf(Number = ucsFscPmtCard, ucsFscPmtCard, ucsFscPmtCash)
        .PmtName = Left$(Name, MAX_PAYMENT_LEN)
        .PmtAmount = Round(Amount, 2)
        If Rate <= 0 Then
            .PmtRate = 1
        Else
            .PmtRate = Rate
        End If
        .PrintRowType = m_uRow(0).InitReceiptType
    End With
    '--- success
    AddPayment = True
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function EndReceipt() As Boolean
    Const FUNC_NAME     As String = "EndReceipt"
    Dim uCtx            As UcsContext
    Dim lRow            As Long
    Dim sData           As String
    Dim sResult         As String
    Dim lTotal          As Long
    Dim lIdx            As Long
    Dim lPlu            As Long
    Dim dblTotal        As Double
    Dim vSplit          As Variant
    Dim lRetry          As Long
    
    On Error GoTo EH
    '--- sanity check
    If m_lRowCount = 0 Then
        Err.Raise vbObjectError, , pvInternal(ucsErrNoReceiptStarted)
    End If
    '--- check device ready
    WaitDevice
    '--- fix fiscal receipts with for more than 30 PLUs
    If m_uRow(0).InitReceiptType = ucsFscRetFiscal Or m_uRow(0).InitReceiptType = ucsFscRetInvoiceFiscal Then
        pvConvertExtraRows
    End If
    '--- append final payment (total)
    With m_uRow(pvAddRow())
        .RowType = ucsRowPayment
        .PrintRowType = m_uRow(0).InitReceiptType
    End With
    '--- ring sale
    pvSetKeylock ucsKlcRegistration
    For lRow = 0 To m_lRowCount - 1
        With m_uRow(lRow)
        Select Case .RowType
        Case ucsRowInit
            uCtx.Paid = -1
            uCtx.PluNo = -.InitPLUNumbering
            '--- cancel previous receipt
            If Not CancelReceipt() Then
                On Error GoTo 0
                Err.Raise vbObjectError, , pvInternal(ucsErrCannotCancelReceipt)
            End If
            '--- get chars per item name
            sResult = SendCommand(ucsEltCmdInfoItem, pvPoke(1, 2))
            uCtx.ItemChars = IIf(Len(sResult) >= 41, 18, 12)
            '--- fix cash payment rate to 1.0
            If GetPaymentRate(1) <> 100 Then
                Mid$(m_sPaymentTypes, 1, 2) = pvPoke(100, 2)
                sResult = SendCommand(ucsEltCmdInitPaymentTypes, m_sPaymentTypes)
            End If
            '--- setup operator name
            lIdx = LimitLong(pvPeek(SendCommand(ucsEltCmdInfoStatus), 10, 2), 1, 10)
            sResult = SendCommand(ucsEltCmdInfoOperator, pvPoke(lIdx, 2))
            sData = sResult
            Mid$(sData, 3, 12) = pvSPoke(.InitOperatorName, 12)
            If sData <> sResult Then
                sResult = SendCommand(ucsEltCmdInitOperator, sData)
            End If
            '--- invoice specific init
            Select Case .InitReceiptType
            Case ucsFscRetInvoiceNonfiscal
                If Len(.InitInvData(ucsInvNumber)) > 10 Then
                    pvPrintLine .InitInvData(ucsInvNumber)
                ElseIf LenB(.InitInvData(ucsInvNumber)) <> 0 Then
                    pvPrintLine AlignText(pvText(ucsTxtHeaderInvoiceNo), pvSPoke(.InitInvData(ucsInvNumber), 10, vbRightJustify, Asc("0")), pvRowChars - 2)
                End If
            Case ucsFscRetInvoiceFiscal
                '--- setup InvoiceNo
                If C_Lng(.InitInvData(ucsInvNumber)) > 0 Then
                    sData = pvPoke(C_Lng(.InitInvData(ucsInvNumber)) - 1, 6)
                    sData = Mid$(sData, 3, 4) & Mid$(sData, 1, 2) & pvPoke(1, 2)
                    sResult = SendCommand(ucsEltCmdInitInvoiceNo, sData)
                End If
                '--- setup contragent
                sData = vbNullString
                For lIdx = 3 To 7
                    sData = sData & IIf(LenB(.InitInvData(lIdx)) <> 0, Chr$(&H80), Chr$(0)) & _
                        pvSPoke(.InitInvData(lIdx), pvRowChars)
                Next
                sResult = SendCommand(ucsEltCmdInitInvoiceData, sData)
                WaitDevice
                '--- setup TaxNo (bulstat)
                sData = Chr$(ucsChrClear) & Chr$(ucsChrX) & Chr$(ucsChrDot)
                sData = sData & pvDumpHex(pvSPoke(.InitInvData(ucsInvTaxNumber), 13), vbNullString)
                '--- setup VatNo
                sData = sData & Chr$(ucsChrDot)
                sData = sData & pvDumpHex(pvSPoke(.InitInvData(ucsInvVatNumber), 13), vbNullString)
                SendInput sData
                pvWaitStatus bInvoice:=True
            End Select
        Case ucsRowPlu
            uCtx.LastRow = m_uRow(lRow)
            uCtx.PluCount = uCtx.PluCount + 1
            dblTotal = Round(.PluQuantity * .PluPrice, 2)
            uCtx.GrpTotal(.PluVatGroup) = Round(uCtx.GrpTotal(.PluVatGroup) + dblTotal, 2)
            '--- on invoices print empty line before first PLU
            If uCtx.PluCount = 1 Then
                Select Case .PrintRowType
                Case ucsFscRetInvoiceFiscal, ucsFscRetInvoiceNonfiscal
                    pvPrintLine vbNullString
                End Select
            End If
            Select Case .PrintRowType
            Case ucsFscRetNonfiscal
                If .PluQuantity <> 1 Then
                    pvPrintLine AlignText("    " & Format$(.PluQuantity, "0.000") & " *", Format$(.PluPrice, "0.00"), pvRowChars - 4)
                End If
                pvPrintLine AlignText(Left$(.PluName, uCtx.ItemChars), Format$(dblTotal, "0.00") & " " & Chr$(191 + .PluVatGroup), pvRowChars - 2)
                If Len(.PluName) > uCtx.ItemChars Then
                    pvPrintLine Trim$(Mid$(.PluName, uCtx.ItemChars + 1))
                End If
            Case ucsFscRetInvoiceNonfiscal
                pvPrintLine .PluName
                pvPrintLine AlignText(pvText(ucsTxtPluPrice), Format$(.PluPrice, "0.00"), pvRowChars - 2)
                pvPrintLine AlignText(pvText(ucsTxtPluQuantity), Format$(.PluQuantity, "0.000"), pvRowChars - 2)
                pvPrintLine AlignText(Printf(pvText(ucsTxtPluSum), Chr$(191 + .PluVatGroup)), Format$(dblTotal, "0.00"), pvRowChars - 2)
            Case ucsFscRetFiscal, ucsFscRetInvoiceFiscal
                If uCtx.PluNo > 0 Then
                    lPlu = uCtx.PluNo
                    uCtx.PluNo = uCtx.PluNo + 1
                Else
                    lPlu = LimitLong(.PluNo, 1, 200)
                End If
                If .PluPrice < 0 Then
                    '--- program keys: keycode 0B to return
                    uCtx.Keys = SendCommand(ucsEltCmdInfoKeyFunctions)
                    sData = uCtx.Keys
                    Mid$(sData, Len(sData) \ 2, 1) = Chr$(&H20) '-- function &H20 = [RFN]
                    sResult = SendCommand(ucsEltCmdInitKeyFunctions, sData)
                End If
                If .PrintRowType = ucsFscRetInvoiceFiscal Then
                    vSplit = WrapText(.PluName, pvRowChars - 2)
                    sData = Trim$(vSplit(UBound(vSplit)))
                    If Len(sData) > uCtx.ItemChars Then
                        sData = vbNullString
                    Else
                        vSplit(UBound(vSplit)) = vbNullString
                    End If
                Else
                    sData = .PluName
                    vSplit = Split(vbNullString)
                End If
                '--- setup PLU name, price and VAT group
                If uCtx.ItemChars = 12 Then
                    sData = pvPoke(lPlu, 2) & pvPoke(Floor(Abs(.PluPrice) * 100), 4) & pvPoke(.PluQuantity * 1000, 4) & pvPoke(0, 1) & _
                        pvPoke(.PluVatGroup - 1, 1) & pvPoke(0, 4) & pvSPoke(sData, 12) & _
                        pvPoke(0, 1) & pvPoke(0, 1) & pvPoke(0, 1) & pvPoke(0, 4)
                Else
                    sData = pvPoke(lPlu, 2) & pvPoke(Floor(Abs(.PluPrice) * 100), 4) & pvPoke(.PluQuantity * 1000, 4) & pvPoke(0, 1) & _
                        pvPoke(.PluVatGroup - 1, 1) & pvPoke(0, 4) & pvPoke(0, 1) & pvSPoke(sData, 18) & _
                        pvPoke(0, 1) & pvPoke(0, 1) & pvPoke(0, 4)
                End If
                sResult = SendCommand(ucsEltCmdInitItem, sData)
                '--- extra checks on first PLU
                If uCtx.PluCount = 1 Then
                    For lIdx = 1 To LimitLong(m_lTimeout \ LNG_DELAY, 1)
                        sResult = SendCommand(ucsEltCmdInfoItem, pvPoke(lPlu, 2))
                        If uCtx.ItemChars = 12 Then
                            Mid$(sResult, 13, 4) = pvPoke(0, 4) '--- turnover
                            Mid$(sResult, 32, 4) = pvPoke(0, 4) '--- quantity sold
                        Else
                            Mid$(sResult, 13, 4) = pvPoke(0, 4) '--- turnover
                            Mid$(sResult, 38, 4) = pvPoke(0, 4) '--- quantity sold
                        End If
                        If sData = sResult Then
                            Exit For
                        End If
                        DebugLog FUNC_NAME, "PLU diff error, lRetry=" & lRetry & ", sResult=" & pvDumpHex(sResult) & ", sData=" & pvDumpHex(sData)
                        sResult = SendCommand(ucsEltCmdInitItem, sData)
                        WaitDevice LNG_DELAY
                    Next
                End If
                '--- on invoices print extra chars in name before PLU
                If .PrintRowType = ucsFscRetInvoiceFiscal Then
                    For lIdx = 0 To UBound(vSplit)
                        If LenB(vSplit(lIdx)) <> 0 Then
                            pvPrintLine vSplit(lIdx)
                        End If
                    Next
                End If
                '--- print PLU (store in fiscal memory)
                lTotal = pvPeek(SendCommand(ucsEltCmdInfoStatus), 0, 4)
                sData = Chr$(ucsChrClear)
                If .PluPrice < 0 Then
                    sData = sData & Chr$(ucsChrRFN)
                End If
                If .PluQuantity <> 1 Then
                    sData = sData & Replace(.PluQuantity, g_sDecimalSeparator, Chr$(ucsChrDot))
                    sData = sData & Chr$(ucsChrX)
                End If
                sData = sData & lPlu & Chr$(ucsChrPLU)
                SendInput sData
                pvWaitStatus lTotal, ucsLcdItem
                '--- on receipts print extra chars in name after PLU
                If .PrintRowType = ucsFscRetFiscal Then
                    If Len(.PluName) > uCtx.ItemChars Then
                        pvPrintLine Trim$(Mid$(.PluName, uCtx.ItemChars + 1))
                    End If
                End If
                If .PluPrice < 0 Then
                    '--- restore keys
                    sResult = SendCommand(ucsEltCmdInitKeyFunctions, uCtx.Keys)
                End If
            End Select
        Case ucsRowLine
            If Not uCtx.Finished Then
                pvPrintLine .LineText, .LineCommand, .LineWordWrap
            End If
        Case ucsRowDiscount
            If .DiscType = ucsFscDstPlu Then
                dblTotal = Round(uCtx.LastRow.PluQuantity * uCtx.LastRow.PluPrice, 2)
                dblTotal = Round(dblTotal * .DiscValue / 100, 2)
                If uCtx.LastRow.PluVatGroup > 0 Then
                    uCtx.GrpTotal(uCtx.LastRow.PluVatGroup) = Round(uCtx.GrpTotal(uCtx.LastRow.PluVatGroup) + dblTotal, 2)
                End If
            End If
            Select Case .PrintRowType
            Case ucsFscRetNonfiscal, ucsFscRetInvoiceNonfiscal
                If .DiscType = ucsFscDstPlu Then
                    If uCtx.LastRow.PluVatGroup > 0 Then
                        pvPrintLine AlignText(" " & Format$(.DiscValue, "0.00") & "%", Format$(dblTotal, "0.00") & " " & Chr$(191 + uCtx.LastRow.PluVatGroup), pvRowChars - 2)
                    End If
                ElseIf .DiscType = ucsFscDstSubtotal Then
                    pvPrintLine AlignText(pvText(ucsTxtSubTotal), Format$(SumArray(uCtx.GrpTotal), "0.00"), pvRowChars - 4)
                    For lIdx = 1 To UBound(uCtx.GrpTotal)
                        If uCtx.GrpTotal(lIdx) <> 0 Then
                            dblTotal = Round(.DiscValue * uCtx.GrpTotal(lIdx) / 100, 2)
                            pvPrintLine AlignText(" " & Format$(.DiscValue, "0.00") & "%", Format$(dblTotal, "0.00") & " " & Chr$(191 + lIdx), pvRowChars - 2)
                        End If
                    Next
                End If
            Case ucsFscRetFiscal, ucsFscRetInvoiceFiscal
                If .DiscType = ucsFscDstPlu Or .DiscType = ucsFscDstSubtotal Then
                    '--- program keys: keycode 0A to surcharge (+%), keycode 0B to discount (-%)
                    uCtx.Keys = SendCommand(ucsEltCmdInfoKeyFunctions)
                    sData = uCtx.Keys
                    Mid$(sData, Len(sData) \ 2 - 1, 2) = Chr$(&HC) & Chr$(&HD) '--- function &HC = [+%], function &HD = [-%]
                    Mid$(sData, Len(sData) - 1, 2) = Chr$(&HC) & Chr$(&HD)
                    sResult = SendCommand(ucsEltCmdInitKeyFunctions, sData)
                    '--- send discount/surcharge
                    lTotal = pvPeek(SendCommand(ucsEltCmdInfoStatus), 0, 4)
                    sData = IIf(.DiscType = ucsFscDstSubtotal, Chr$(ucsChrSubtotal), vbNullString)
                    If .DiscValue <> 0 Then
                        sData = sData & Replace(Abs(.DiscValue), g_sDecimalSeparator, Chr$(ucsChrDot))
                        sData = sData & IIf(.DiscValue < 0, Chr$(ucsChrMinus), Chr$(ucsChrPlus))
                    End If
                    If LenB(sData) <> 0 Then
                        SendInput sData
                        pvWaitStatus lTotal, ucsLcdPlus, ucsLcdMinus
                    End If
                    '--- restore keys
                    sResult = SendCommand(ucsEltCmdInitKeyFunctions, uCtx.Keys)
                End If
            End Select
            '--- for subtotal discount, modify GrpTotals after print
            If .DiscType = ucsFscDstSubtotal Then
                For lIdx = 1 To UBound(uCtx.GrpTotal)
                    If uCtx.GrpTotal(lIdx) <> 0 Then
                        dblTotal = Round(.DiscValue * uCtx.GrpTotal(lIdx) / 100, 2)
                        uCtx.GrpTotal(lIdx) = Round(uCtx.GrpTotal(lIdx) + dblTotal, 2)
                    End If
                Next
            End If
        Case ucsRowPayment
            Select Case .PrintRowType
            Case ucsFscRetNonfiscal, ucsFscRetInvoiceNonfiscal
                If .PmtType <> 0 Or uCtx.PluCount > 0 Then
                    If uCtx.Paid < 0 Then
                        pvPrintLine String$(pvRowChars - 2, "-")
                        pvPrintLine AlignText(pvText(ucsTxtTotal), Format$(SumArray(uCtx.GrpTotal), "0.00"), pvRowChars \ 2 - 1), "DBWIDTH"
                        pvPrintLine String$(pvRowChars - 2, "-")
                        uCtx.Paid = 0
                        If .PrintRowType = ucsFscRetInvoiceNonfiscal Then
                            sResult = SendCommand(ucsEltCmdInfoVatGroups)
                            '--- invoice totals by VAT groups
                            For lIdx = 1 To UBound(uCtx.GrpTotal)
                                If uCtx.GrpTotal(lIdx) <> 0 Then
                                    pvPrintLine AlignText(pvText(ucsTxtPmtTotal), Format$(uCtx.GrpTotal(lIdx), "0.00"), pvRowChars - 2)
                                    '--- total w/o VAT
                                    dblTotal = Round(uCtx.GrpTotal(lIdx) / (1 + pvPeek(sResult, (lIdx - 1) * 2, 2) / 10000#), 2)
                                    pvPrintLine AlignText(Printf(pvText(ucsTxtPmtVatGroup), Chr$(191 + lIdx)), Format$(uCtx.GrpTotal(lIdx) - dblTotal, "0.00"), pvRowChars - 2)
                                    pvPrintLine Chr$(191 + lIdx) & "=" & Format$(pvPeek(sResult, (lIdx - 1) * 2, 2) / 100#, "0.00")
                                    pvPrintLine AlignText(pvText(ucsTxtPmtNeto), Format$(dblTotal, "0.00"), pvRowChars - 2)
                                End If
                            Next
                        End If
                    End If
                    If .PmtType = 0 Then
                        .PmtAmount = Round(SumArray(uCtx.GrpTotal) - uCtx.Paid, 2)
                        .PmtRate = 1
                        .PmtName = GetPaymentName(1)
                    ElseIf LenB(.PmtName) = 0 Then
                        .PmtName = GetPaymentName(.PmtType)
                    End If
                    If .PmtAmount >= 0 Then
                        pvPrintLine AlignText(.PmtName, Format$(Round(.PmtAmount * .PmtRate, 2), "0.00"), pvRowChars - 2)
                        uCtx.Paid = uCtx.Paid + Round(.PmtAmount * .PmtRate, 2)
                    End If
                    If .PmtType <> 0 Then
                        dblTotal = Round(uCtx.Paid - SumArray(uCtx.GrpTotal), 2)
                        If dblTotal > 0 Then
                            pvPrintLine AlignText(pvText(ucsTxtPmtChange), Format$(Round(dblTotal, 2), "0.00"), pvRowChars - 2)
                        End If
                        If dblTotal >= 0 Then
                            uCtx.Finished = True
                        End If
                    End If
                    If .PmtType = 0 Then
                        If .PrintRowType = ucsFscRetInvoiceNonfiscal Then
                            With m_uRow(0)
                                '--- invoice contragent info
                                pvPrintLine vbNullString
                                pvPrintLine pvText(ucsTxtFooterSeller)
                                pvPrintLine pvText(ucsTxtFooterDivider)
                                pvPrintLine pvText(ucsTxtFooterReceiver)
                                For lIdx = 3 To 7
                                    If LenB(.InitInvData(lIdx)) <> 0 Then
                                        pvPrintLine .InitInvData(lIdx)
                                    End If
                                Next
                                pvPrintLine pvText(ucsTxtFooterTaxNo) & .InitInvData(ucsInvTaxNumber)
                                If LenB(.InitInvData(ucsInvVatNumber)) <> 0 Then
                                    pvPrintLine pvText(ucsTxtFooterVatNo) & .InitInvData(ucsInvVatNumber)
                                End If
                                pvPrintLine pvText(ucsTxtFooterDivider)
                            End With
                        End If
                        pvPrintLine AlignText(vbNullString, Printf(pvText(ucsTxtFooterPluCount), Format$(uCtx.PluCount, "00")), pvRowChars - 2)
                        For lIdx = 1 To 2
                            sData = GetFooterText(lIdx)
                            If LenB(Trim$(sData)) <> 0 Then
                                If Left$(sData, 1) = " " Then
                                    sData = Mid$(sData, 2)
                                End If
                                pvPrintLine sData, bWordWrap:=False
                            End If
                        Next
                    End If
                End If
            Case ucsFscRetFiscal, ucsFscRetInvoiceFiscal
                If .PmtType <> 0 Or Not uCtx.Finished Then
                    If LenB(.PmtName) <> 0 Then
                        sData = m_sPaymentTypes
                        Mid$(sData, 1 + (.PmtType - 1) * 2, 2) = pvPoke(Round(IIf(.PmtType <> 1, .PmtRate, 1) * 100), 2)
                        Mid$(sData, 9 + (.PmtType - 1) * 6, 6) = pvSPoke(.PmtName, 6)
                        If sData <> m_sPaymentTypes Then
                            sResult = SendCommand(ucsEltCmdInitPaymentTypes, sData)
                            m_sPaymentTypes = SendCommand(ucsEltCmdInfoPaymentTypes)
                        End If
                    End If
                    lTotal = pvPeek(SendCommand(ucsEltCmdInfoStatus), 0, 4)
                    sData = Replace(Round(.PmtAmount, 2), g_sDecimalSeparator, Chr$(ucsChrDot))
                    Select Case .PmtType
                    Case 2
                        sData = sData & Chr$(ucsChrPy1)
                    Case 3, 4
                        On Error GoTo 0
                        Err.Raise vbObjectError, , Printf("Payment type %1 not implemented yet", .PmtType)
                    Case 0
                        sData = Chr$(ucsChrTotal)
                    Case Else
                        sData = sData & Chr$(ucsChrTotal)
                    End Select
                    SendInput sData
                    sResult = pvWaitStatus(lTotal, ucsLcdPayment, ucsLcdOutOfReceipt)
                    If pvPeek(sResult, 4, 2) = ucsLcdOutOfReceipt Then
                        uCtx.Finished = True
                    End If
                    m_bFiscalPaymentExecuted = True
                End If
            End Select
        End Select
        If LenB(m_sLastError) <> 0 Then
            On Error GoTo 0
            Err.Raise vbObjectError, , m_sLastError
        End If
        End With
    Next
    '--- eject non-fiscal receipt
    Select Case m_uRow(0).InitReceiptType
    Case ucsFscRetNonfiscal, ucsFscRetInvoiceNonfiscal
        pvSetKeylock ucsKlcXReport
    End Select
    '--- show clock
    pvSetKeylock ucsKlcEcrUnlock
    SendInput Chr$(ucsChrClear)
    '--- success
    EndReceipt = True
    Exit Function
EH:
    PrintError FUNC_NAME
    If uCtx.PluCount > 0 Then
        m_sLastError = Printf(pvInternal(ucsErrErrorInLine), uCtx.PluCount, m_sLastError)
    End If
    Err.Raise vbObjectError, , m_sLastError
End Function

Public Function CopyLastReceipt() As Boolean
    Const FUNC_NAME     As String = "CopyLastReceipt"
    Dim sResult         As String
    Dim lIdx            As Long
    Dim lReceipt        As Long
    
    On Error GoTo EH
    '--- check device ready
    WaitDevice
    pvSetKeylock ucsKlcRegistration
    lReceipt = pvPeek(SendCommand(ucsEltCmdInfoStatus), 6, 2)
    SendInput Chr$(ucsChrClear) & "1 "
    For lIdx = 1 To LimitLong(m_lTimeout \ LNG_DELAY, 1)
        sResult = SendCommand(ucsEltCmdInfoStatus)
        If pvPeek(sResult, 6, 2) <> lReceipt Then
            Exit For
        End If
        If pvPeek(sResult, 8, 2) = 0 Then '--- oper no = 0 -> failure
            m_sLastError = pvInternal(ucsErrInvalidCommandForDuplicateReceipt)
            GoTo QH
        End If
        WaitDevice LNG_DELAY
    Next
    pvSetKeylock ucsKlcEcrUnlock
    SendInput Chr$(ucsChrClear)
    WaitDevice
    '--- success
    CopyLastReceipt = True
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetClock() As Date
    Const FUNC_NAME     As String = "GetClock"
    Dim sResult         As String
    
    On Error GoTo EH
    sResult = pvFromPbcd(SendCommand(ucsEltCmdInfoDateTime))
    GetClock = DateSerial(Mid$(sResult, 5, 2), Mid$(sResult, 7, 2), 2000 + Mid$(sResult, 9, 2)) _
        + TimeSerial(Mid$(sResult, 3, 2), Mid$(sResult, 1, 2), 0)
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Sub SetClock(ByVal NewDate As Date)
    Const FUNC_NAME     As String = "SetClock"
    Dim sData           As String
    
    On Error GoTo EH
    sData = Chr$("&H" & Day(NewDate)) & Chr$("&H" & Month(NewDate)) & Chr$("&H" & (Year(NewDate) Mod 100)) & _
        Chr$("&H" & Weekday(NewDate, vbMonday)) & Chr$("&H" & Minute(NewDate)) & Chr$("&H" & Hour(NewDate))
    Call SendCommand(ucsEltCmdInitDateTime, sData)
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Public Function RunXReport(ByVal ReportType As UcsFiscalReportsTypeEnum) As Boolean
    Const FUNC_NAME     As String = "RunXReport"
    
    If Not CancelReceipt() Then
        Err.Raise vbObjectError, , pvInternal(ucsErrCannotCancelReceipt)
    End If
    On Error GoTo EH
    Select Case ReportType
    Case ucsFscRstDaily
        RunXReport = pvRunReport(ucsKlcXReport, Chr$(ucsChrTotal) & Chr$(ucsChrClear) & Chr$(ucsChrTotal))
    Case ucsFscRstDailyByDepartment
        RunXReport = pvRunReport(ucsKlcXReport, "2" & Chr$(ucsChrD7) & Chr$(ucsChrTotal))
    Case ucsFscRstDailyByItems
        RunXReport = pvRunReport(ucsKlcXReport, "3" & Chr$(ucsChrD7) & Chr$(ucsChrTotal))
    End Select
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function RunZReport(ByVal ReportType As UcsFiscalReportsTypeEnum) As Boolean
    Const FUNC_NAME     As String = "RunZReport"
    
    If Not CancelReceipt() Then
        Err.Raise vbObjectError, , pvInternal(ucsErrCannotCancelReceipt)
    End If
    On Error GoTo EH
    Select Case ReportType
    Case ucsFscRstDaily
'        RunZReport = pvRunReport(ucsKlcZReport, Chr$(ucsChrTotal))
        Call SendCommand(ucsEltCmdPrintZReport)
        RunZReport = True
    Case ucsFscRstDailyByDepartment
        RunZReport = pvRunReport(ucsKlcZReport, "2" & Chr$(ucsChrD7) & Chr$(ucsChrTotal))
    Case ucsFscRstDailyByItems
        RunZReport = pvRunReport(ucsKlcZReport, "3" & Chr$(ucsChrD7) & Chr$(ucsChrTotal))
    End Select
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function RunPeriodReport(ByVal ReportType As UcsFiscalReportsTypeEnum, ByVal StartDate As Date, ByVal EndDate As Date) As Boolean
    Const FUNC_NAME     As String = "RunPeriodReport"
    
    If Not CancelReceipt() Then
        Err.Raise vbObjectError, , pvInternal(ucsErrCannotCancelReceipt)
    End If
    On Error GoTo EH
    Select Case ReportType
    Case ucsFscRstPeriodDetailed, ucsFscRstPeriodShort
        RunPeriodReport = pvRunReport(ucsKlcXReport, _
            IIf(ReportType = ucsFscRstPeriodDetailed, "1", "11") & Chr$(ucsChrSubtotal) & _
            Format$(StartDate, "ddmmyy") & Chr$(ucsChrX) & _
            Format$(EndDate, "ddmmyy") & Chr$(ucsChrVoid))
    End Select
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetLastReceiptNumber() As Long
    Const FUNC_NAME     As String = "GetLastReceiptNumber"
    
    On Error GoTo EH
    GetLastReceiptNumber = pvPeek(SendCommand(ucsEltCmdInfoStatus), 6, 2)
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetLastInvoiceNumber() As String
    Const FUNC_NAME     As String = "GetLastInvoiceNumber"
    Dim sResult         As String
    
    On Error GoTo EH
    sResult = SendCommand(ucsEltCmdInfoInvoiceNo)
    GetLastInvoiceNumber = pvPeek(Mid$(sResult, 5, 2) & sResult, 0, 6)
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetSerialNumber() As String
    Const FUNC_NAME     As String = "GetSerialNumber"
    
    On Error GoTo EH
    GetSerialNumber = SendCommand(ucsEltCmdInfoFpNumber)
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetFiscalNumber() As String
    Const FUNC_NAME     As String = "GetFiscalNumber"
    
    On Error GoTo EH
    GetFiscalNumber = pvFromPbcd(SendCommand(ucsEltCmdInfoFiscalNumber))
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetTaxNumber() As String
    Const FUNC_NAME     As String = "GetTaxNumber"
    
    On Error GoTo EH
    GetTaxNumber = pvSPeek(SendCommand(ucsEltCmdInfoBulstat), 0, 13)
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetTaxCaption() As String
    Const FUNC_NAME     As String = "GetTaxCaption"
    
    On Error GoTo EH
    If LenB(m_sHeadingLines) = 0 Then
        m_sHeadingLines = SendCommand(ucsEltCmdInfoHeadingLines)
    End If
    GetTaxCaption = pvSPeek(m_sHeadingLines, 7 * (pvRowChars + 1) + 1, pvRowChars \ 2)
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetHeaderText(ByVal Index As Long) As String
    Const FUNC_NAME     As String = "GetHeaderText"
    
    On Error GoTo EH
    If LenB(m_sHeadingLines) = 0 Then
        m_sHeadingLines = SendCommand(ucsEltCmdInfoHeadingLines)
    End If
    If Index >= 1 And Index <= 5 Then
        GetHeaderText = pvSPeek(m_sHeadingLines, (Index - 1) * (pvRowChars + 1) + 1, pvRowChars)
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetFooterText(ByVal Index As Long) As String
    Const FUNC_NAME     As String = "GetFooterText"
    
    On Error GoTo EH
    If LenB(m_sHeadingLines) = 0 Then
        m_sHeadingLines = SendCommand(ucsEltCmdInfoHeadingLines)
    End If
    If Index >= 1 And Index <= 2 Then
        GetFooterText = pvSPeek(m_sHeadingLines, (Index + 4) * (pvRowChars + 1) + 1, pvRowChars)
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetPaymentName(ByVal Index As Long) As String
    Const FUNC_NAME     As String = "GetPaymentName"
    
    On Error GoTo EH
    If LenB(m_sPaymentTypes) = 0 Then
        m_sPaymentTypes = SendCommand(ucsEltCmdInfoPaymentTypes)
    End If
    If Index >= 1 And Index <= 4 Then
        GetPaymentName = pvSPeek(m_sPaymentTypes, 8 + (Index - 1) * 6, 6)
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetPaymentRate(ByVal Index As Long) As Long
    Const FUNC_NAME     As String = "GetPaymentRate"
    
    On Error GoTo EH
    If LenB(m_sPaymentTypes) = 0 Then
        m_sPaymentTypes = SendCommand(ucsEltCmdInfoPaymentTypes)
    End If
    If Index >= 1 And Index <= 4 Then
        GetPaymentRate = pvPeek(m_sPaymentTypes, 2 * (Index - 1), 2)
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function CancelReceipt() As Boolean
    Const FUNC_NAME     As String = "CancelReceipt"
    Dim lTotal          As Long
    Dim lIdx            As Long
    Dim sResult         As String
    
    On Error GoTo EH
    '--- cancel outstanding receipt
    lTotal = pvPeek(SendCommand(ucsEltCmdInfoStatus), 0, 4)
    If lTotal <> 0 Then
        SendInput Chr$(ucsChrClear)
        For lIdx = 1 To MAX_RECEIPT_ROWS
            '--- void last row
            SendInput Chr$(ucsChrVoid)
            sResult = pvWaitStatus(lTotal, ucsLcdVoid, ucsLcdOutOfReceipt)
            If pvPeek(sResult, 0, 4) = 0 Then
                Exit For
            End If
            If pvPeek(sResult, 4, 2) <> ucsLcdVoid Then
                SendInput Chr$(ucsChrClear) & Replace(Round(pvPeek(sResult, 0, 4) / 100#, 2), g_sDecimalSeparator, Chr$(ucsChrDot))
                Exit For
            End If
        Next
        SendInput Chr$(ucsChrTotal)
        pvWaitStatus -1, ucsLcdOutOfReceipt
    End If
    '--- success (or failure)
    CancelReceipt = (pvPeek(SendCommand(ucsEltCmdInfoStatus), 0, 4) = 0)
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetDeviceProtocol() As String
    Const FUNC_NAME     As String = "GetDeviceProtocol"
    
    On Error GoTo EH
    pvGetDeviceInfo m_hComm, m_lTimeout, GetDeviceProtocol, vbNullString
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetDeviceModel() As String
    Const FUNC_NAME     As String = "GetDeviceModel"
    
    On Error GoTo EH
    pvGetDeviceInfo m_hComm, m_lTimeout, vbNullString, GetDeviceModel
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Function GetTotalsByVatGroups() As Variant
    Const FUNC_NAME     As String = "GetTotalsByVatGroups"
    Dim sResult         As String
    Dim vRetVal         As Variant
    Dim lIdx            As Long
    
    On Error GoTo EH
    sResult = SendCommand(ucsEltCmdInfoTurnoverVatGroups)
    If Len(sResult) \ 4 > 0 Then
        ReDim vRetVal(0 To Len(sResult) \ 4 - 1) As Double
        For lIdx = 0 To UBound(vRetVal)
            vRetVal(lIdx) = C_Dbl(pvPeek(sResult, lIdx * 4, 4)) / 100
        Next
        GetTotalsByVatGroups = vRetVal
    Else
        GetTotalsByVatGroups = EmptyDoubleArray
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
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
        .WriteTotalTimeoutConstant = 1000
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
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function pvSendCommand( _
            ByVal hComm As Long, _
            ByVal lTimeout As Long, _
            ByVal eCmd As UcsEltradeCommandsEnum, _
            Optional sData As String, _
            Optional eRetCmd As UcsEltradeCommandsEnum, _
            Optional sError As String) As String
    Const FUNC_NAME     As String = "pvSendCommand"
    Dim sText           As String
    Dim baText()        As Byte
    Dim lIdx            As Long
    Dim lSum            As Long
    Dim lActual         As Long
    Dim lResult         As Long
    Dim uOver           As OVERLAPPED
    Dim sBuffer         As String
    Dim lSize           As Long
    Dim lMask           As Long
    
    On Error GoTo EH
    DebugLog FUNC_NAME, "eCmd=0x" & Hex(eCmd) & ", sData=" & pvDumpHex(sData)
    uOver.hEvent = CreateEvent(0, 1, 0, 0)
    sText = Chr$(&HAA) & Chr$(&H55) & Chr$(m_lCashDesk) & Chr$(eCmd) & Chr$(Len(sData)) & sData
    '--- calc checksum
    baText = StrConv(sText, vbFromUnicode)
    For lIdx = 3 To UBound(baText)
        lSum = lSum + baText(lIdx)
    Next
    '--- append checksum
    sText = sText & Chr$((256 - lSum Mod 256) Mod 256)
    lSize = Len(sText)
    baText = StrConv(sText, vbFromUnicode)
    '--- async send
    lResult = WriteFile(hComm, baText(0), lSize, lActual, VarPtr(uOver))
    If lResult = 0 Then
        If Err.LastDllError <> ERROR_IO_PENDING Then
            DebugLog FUNC_NAME, "WriteFile, Err.LastDllError=" & Err.LastDllError
            sError = pvInternal(ucsErrWriteFileFailure) & GetApiErr(Err.LastDllError)
            GoTo QH
        End If
    End If
    lResult = WaitForMultipleObjects(1, uOver.hEvent, 0, lTimeout)
    If lResult <> WAIT_OBJECT_0 Then
        DebugLog FUNC_NAME, "CancelIo after WriteFile, lResult=" & lResult
        Call CancelIo(hComm)
        Call Sleep(100)
        Call FlushFileBuffers(hComm)
        GoTo QH
    End If
    '--- async read response
    sText = vbNullString
    sBuffer = String$(1000, 0)
    Call SetCommMask(hComm, EV_RXCHAR)
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
            lResult = ReadFile(hComm, ByVal StrPtr(sBuffer), Len(sBuffer), lActual, VarPtr(uOver))
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
                sText = sText & Left$(StrConv(sBuffer, vbUnicode), lActual)
            End If
        End If
        If lResult <> WAIT_OBJECT_0 Then
            DebugLog FUNC_NAME, "CancelIo after ReadFile, lResult=" & lResult
            If lResult = WAIT_TIMEOUT Then
                sError = pvInternal(ucsErrTimeoutWaitingForResponse)
            End If
            Call CancelIo(hComm)
            Call Sleep(100)
            Call FlushFileBuffers(hComm)
            GoTo QH
        End If
        '--- check if response complete
        If Len(sText) >= 5 Then
            baText = StrConv(sText, vbFromUnicode)
            If baText(0) <> &HAA Or baText(1) <> &H55 Or baText(2) <> m_lCashDesk Then
                DebugLog FUNC_NAME, "Invalid response header, sText=" & pvDumpHex(sText)
                sError = pvInternal(ucsErrInvalidResponseHeader)
                Call CancelIo(hComm)
                Call Sleep(100)
                Call FlushFileBuffers(hComm)
                GoTo QH
            End If
            '--- 5 = sizeof header, baText(4) = size of response, 1 - size of checksum
            If Len(sText) >= 5 + baText(4) + 1 Then
                eRetCmd = baText(3)
                pvSendCommand = Mid$(sText, 6, baText(4))
                DebugLog FUNC_NAME, "eRetCmd=0x" & Hex(eRetCmd) & ", RetVal=" & pvDumpHex(pvSendCommand)
                Exit Do
            End If
        End If
    Loop
QH:
    Call CloseHandle(uOver.hEvent)
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function pvWaitStatus( _
            Optional ByVal lTotal As Long = -1, _
            Optional ByVal lLastCommand As UcsLastCommandEnum = -1, _
            Optional ByVal lAltCommand As UcsLastCommandEnum = -1, _
            Optional ByVal bInvoice As Boolean) As String
    Const FUNC_NAME     As String = "pvWaitStatus"
    Dim lIdx            As Long
    Dim bPulse          As Boolean
    Dim lRetry          As Long
    
    On Error GoTo EH
    For lIdx = 1 To LimitLong(m_lTimeout \ LNG_DELAY, 1)
        On Error Resume Next
        For lRetry = 1 To LimitLong(m_lTimeout \ 1000, 2)
            pvWaitStatus = SendCommand(ucsEltCmdInfoStatus, lTimeout:=1000)
            If Err.Number = 0 Then
                Exit For
            End If
            DebugLog FUNC_NAME, "info status, lRetry=" & lRetry & ", m_sLastError=" & m_sLastError
        Next
        On Error GoTo 0
        If lTotal <> -1 And Not bPulse Then
            If pvPeek(pvWaitStatus, 0, 4) <> lTotal Then
                Exit For
            End If
            If lTotal = 0 Then
                If lLastCommand <> -1 Then
                    If Not pvPeek(pvWaitStatus, 4, 2) = lLastCommand Then
                        bPulse = True
                    End If
                End If
                If lAltCommand <> -1 Then
                    If Not pvPeek(pvWaitStatus, 4, 2) = lAltCommand Then
                        bPulse = True
                    End If
                End If
                If bInvoice Then
                    If Not pvPeek(pvWaitStatus, 18, 2) <> 0 Then
                        bPulse = True
                    End If
                End If
            End If
        Else
            If lLastCommand <> -1 Then
                If pvPeek(pvWaitStatus, 4, 2) = lLastCommand Then
                    Exit For
                End If
            End If
            If lAltCommand <> -1 Then
                If pvPeek(pvWaitStatus, 4, 2) = lAltCommand Then
                    Exit For
                End If
            End If
            If bInvoice Then
                If pvPeek(pvWaitStatus, 18, 2) <> 0 Then
                    Exit For
                End If
            End If
        End If
        WaitDevice LNG_DELAY
    Next
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Function pvPrintLine( _
            ByVal sText As String, _
            Optional sCommand As String, _
            Optional ByVal bWordWrap As Boolean = True) As String
    Const FUNC_NAME     As String = "pvPrintLine"
    Dim lIdx            As Long
    Dim vSplit          As Variant
    Dim sData           As String
    
    On Error GoTo EH
    If bWordWrap Then
        vSplit = WrapText(sText, IIf(UCase$(sCommand) = "DBWIDTH", pvRowChars \ 2 - 1, pvRowChars - 2))
    Else
        vSplit = Array(sText)
    End If
    For lIdx = 0 To UBound(vSplit)
        Select Case UCase$(sCommand)
        Case "DBWIDTH"
            sData = sData & Chr$(&H90)
        Case "CONDENSED"
            sData = sData & Chr$(&H84)
        Case Else
            sData = sData & Chr$(&H80)
        End Select
        sData = sData & pvSPoke(vSplit(lIdx), pvRowChars - 1)
        If lIdx >= 7 Then
            Exit For
        End If
    Next
    pvPrintLine = SendCommand(ucsEltCmdCommentLine, sData)
    WaitDevice
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Function pvRunReport(ByVal lKeylock As UcsKeylockPosition, sData As String) As Boolean
    Const FUNC_NAME     As String = "pvRunReport"
    Dim lReceipt        As Long
    Dim lIdx            As Long
    
    On Error GoTo EH
    WaitDevice
    pvSetKeylock lKeylock
    lReceipt = pvPeek(SendCommand(ucsEltCmdInfoStatus), 6, 2)
    SendInput sData
    For lIdx = 1 To LimitLong(m_lTimeout \ LNG_DELAY, 1)
        If pvPeek(SendCommand(ucsEltCmdInfoStatus), 6, 2) <> lReceipt Then
            pvRunReport = True
            Exit For
        End If
        WaitDevice LNG_DELAY
    Next
    pvSetKeylock ucsKlcEcrUnlock
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Function pvSetKeylock(ByVal ePosition As UcsKeylockPosition) As Boolean
    Const FUNC_NAME     As String = "pvSetKeylock"
    
    On Error GoTo EH
    If pvPeek(SendCommand(ucsEltCmdInfoStatus), 8, 2) <> ePosition Then
        Call SendCommand(ucsEltCmdInitKeylock, Chr$(ePosition))
        WaitDevice
        If pvPeek(SendCommand(ucsEltCmdInfoStatus), 8, 2) <> ePosition Then
            Call SendCommand(ucsEltCmdInitKeylock, Chr$(ePosition))
        End If
        '--- success
        pvSetKeylock = True
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Function pvInternal(ByVal lIdx As UcsInternalErrorsEnum) As String
    pvInternal = At(Split(m_sLocalizedText(ucsFscLciInternalErrors), "|"), lIdx, At(Split(STR_INTERNAL_ERRORS, "|"), lIdx))
End Function

Private Function pvText(ByVal lIdx As UcsReceiptTextsEnum) As String
    pvText = At(Split(m_sLocalizedText(ucsFscLciReceiptTexts), "|"), lIdx, At(Split(STR_RECEIPT_TEXTS, "|"), lIdx))
End Function

Private Function pvAddRow() As Long
    Const FUNC_NAME     As String = "pvAddRow"
    
    On Error GoTo EH
    If m_lRowCount > UBound(m_uRow) Then
        ReDim Preserve m_uRow(0 To 2 * UBound(m_uRow)) As UcsRowData
    End If
    pvAddRow = m_lRowCount
    m_lRowCount = m_lRowCount + 1
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Sub pvInsertRow(ByVal lRow As Long, uRow As UcsRowData)
    Const FUNC_NAME     As String = "pvInsertRow"
    Dim lIdx            As Long
    
    On Error GoTo EH
    If lRow = 0 Or lRow >= m_lRowCount Then
        m_uRow(pvAddRow()) = uRow
    Else
        '--- shift rows down and insert new row
        For lIdx = pvAddRow() To lRow + 1 Step -1
            m_uRow(lIdx) = m_uRow(lIdx - 1)
        Next
        m_uRow(lRow) = uRow
    End If
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Private Function pvPeek(sText As String, ByVal lOffset As Long, Optional ByVal lSize As Long = 1) As Currency
    Dim sTemp           As String
    
    sTemp = StrReverse(Mid$(sText, lOffset + 1, lSize))
    Call CopyMemory(pvPeek, ByVal sTemp, Len(sTemp))
    pvPeek = pvPeek * 10000@
End Function

Private Function pvPoke(ByVal cValue As Currency, Optional ByVal lSize As Long = 1) As String
    Dim sTemp           As String
    
    cValue = cValue / 10000@
    sTemp = String$(8, 0)
    Call CopyMemory(ByVal sTemp, cValue, 8)
    pvPoke = Right$(StrReverse(sTemp), lSize)
End Function

Private Function pvSPeek(sText As String, ByVal lOffset As Long, ByVal lSize As Long) As String
    pvSPeek = pvFromMik(RTrim$(Replace$(Mid$(sText, lOffset + 1, lSize), Chr$(0), vbNullString)))
End Function

Private Function pvSPoke(ByVal sText As String, ByVal lSize As Long, Optional ByVal eAlign As AlignmentConstants = vbLeftJustify, Optional ByVal lFillChar As Long = 32) As String
    Select Case eAlign
    Case vbLeftJustify
        pvSPoke = Left$(sText, lSize)
        pvSPoke = pvSPoke & String$(lSize - Len(pvSPoke), lFillChar)
    Case vbRightJustify
        pvSPoke = Right$(sText, lSize)
        pvSPoke = String$(lSize - Len(pvSPoke), lFillChar) & pvSPoke
    Case vbCenter
        pvSPoke = Left$(sText, lSize)
        pvSPoke = String$((lSize - Len(pvSPoke) + 1) \ 2, lFillChar) & pvSPoke
        pvSPoke = pvSPoke & String$(lSize - Len(pvSPoke), lFillChar)
    End Select
    pvSPoke = pvToMik(pvSPoke)
End Function

Private Function pvFromPbcd(sText As String) As String
    Const FUNC_NAME     As String = "pvFromPbcd"
    Dim baText()        As Byte
    Dim lIdx            As Long
    
    On Error GoTo EH
    baText = StrConv(sText, vbFromUnicode)
    For lIdx = 0 To UBound(baText)
        pvFromPbcd = pvFromPbcd & Right$("0" & Hex(baText(lIdx)), 2)
    Next
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Function pvFromMik(sText As String) As String
    Const FUNC_NAME     As String = "pvFromMik"
    Dim lIdx            As Long
    Dim lChar           As Long
    
    On Error GoTo EH
    pvFromMik = sText
    For lIdx = 1 To Len(sText)
        lChar = Asc(Mid$(sText, lIdx, 1))
        If lChar >= &H80 And lChar < &H80 + 64 Then
            Mid$(pvFromMik, lIdx, 1) = Chr$(lChar + 64)
        End If
    Next
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Function pvToMik(sText As String) As String
    Const FUNC_NAME     As String = "pvToMik"
    Dim lIdx            As Long
    Dim lChar           As Long
    
    On Error GoTo EH
    pvToMik = sText
    For lIdx = 1 To Len(sText)
        lChar = Asc(Mid$(sText, lIdx, 1))
        If lChar >= &HC0 And lChar < &HC0 + 64 Then
            Mid$(pvToMik, lIdx, 1) = Chr$(lChar - 64)
        End If
    Next
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Function pvDumpHex(ByVal sText As String, Optional sSeparator As String = " ") As String
    Const FUNC_NAME     As String = "pvDumpHex"
    Dim lIdx            As Long
    
    On Error GoTo EH
    For lIdx = 1 To Len(sText)
        pvDumpHex = pvDumpHex & sSeparator & Right$("0" & Hex(Asc(Mid$(sText, lIdx, 1))), 2)
    Next
    pvDumpHex = Mid$(pvDumpHex, 1 + Len(sSeparator), Len(pvDumpHex))
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Sub pvConvertExtraRows()
    Const FUNC_NAME     As String = "pvConvertExtraRows"
    Dim uCtx            As UcsContext
    Dim lIdx            As Long
    Dim lRow            As Long
    Dim lCount          As Long
    Dim lTotal          As Long
    Dim dblTotal        As Double
    Dim uSum            As UcsContext
    Dim dblDiscount     As Double
    Dim bIsLastDiscount As Boolean
    
    On Error GoTo EH
    '--- convert out-of-range discounts to PLU rows
    '--- note: m_lRowCount may change in loop on AddPLU
    Do While lRow < m_lRowCount
        '--- note: 'With' locks m_uRow array and fails if auto-grow needed in AddPLU
'        With m_uRow(lRow)
            If m_uRow(lRow).RowType = ucsRowPlu Then
                uCtx.LastRow = m_uRow(lRow)
            ElseIf m_uRow(lRow).RowType = ucsRowDiscount And (m_uRow(lRow).DiscValue < MIN_DISCOUNT Or m_uRow(lRow).DiscValue > MAX_DISCOUNT) Then
                If m_uRow(lRow).DiscType = ucsFscDstPlu And uCtx.LastRow.PluVatGroup > 0 Then
                    dblTotal = Round(uCtx.LastRow.PluQuantity * uCtx.LastRow.PluPrice, 2)
                    dblDiscount = Limit(m_uRow(lRow).DiscValue, MIN_DISCOUNT, MAX_DISCOUNT)
                    If Round(dblTotal * m_uRow(lRow).DiscValue / 100#, 2) = Round(dblTotal * dblDiscount / 100#, 2) Then
                        m_uRow(lRow).DiscValue = dblDiscount
                    Else
                        With m_uRow(lRow)
                            .RowType = ucsRowPlu
                            .PluNo = 1
                            .PluName = Format$(.DiscValue, "0.00") & "%"
                            .PluPrice = Round(dblTotal * .DiscValue / 100#, 2)
                            .PluQuantity = 1
                            .PluVatGroup = uCtx.LastRow.PluVatGroup
                        End With
                    End If
                ElseIf m_uRow(lRow).DiscType = ucsFscDstSubtotal Then
                    pvGetSubtotals lRow, uSum
                    dblDiscount = Limit(m_uRow(lRow).DiscValue, MIN_DISCOUNT, MAX_DISCOUNT)
                    lCount = 0
                    For lIdx = 1 To UBound(uSum.GrpTotal)
                        If Round(uSum.GrpTotal(lIdx) * m_uRow(lRow).DiscValue / 100#, 2) <> Round(uSum.GrpTotal(lIdx) * dblDiscount / 100#, 2) Then
                            lCount = lCount + 1
                        End If
                    Next
                    If lCount = 0 Then
                        m_uRow(lRow).DiscValue = dblDiscount
                    Else
                        m_uRow(lRow).DiscValue = 0
                        For lIdx = UBound(uSum.GrpTotal) To 1 Step -1
                            If uSum.GrpTotal(lIdx) <> 0 Then
                                AddPLU lIdx, Format$(m_uRow(lRow).DiscValue, "0.00") & "%", Round(uSum.GrpTotal(lIdx) * m_uRow(lRow).DiscValue / 100#, 2), 1, lIdx, lRow + 1
                            End If
                        Next
                    End If
                End If
            End If
'        End With
        lRow = lRow + 1
    Loop
    '--- count PLU rows and mark different VAT groups
    lCount = 0
    For lRow = 0 To m_lRowCount - 1
        With m_uRow(lRow)
            If .RowType = ucsRowPlu Then
                bIsLastDiscount = False
                lCount = lCount + 1
                uCtx.GrpTotal(.PluVatGroup) = 1
            ElseIf .RowType = ucsRowDiscount And .DiscType = ucsFscDstPlu Then
                bIsLastDiscount = True
            End If
        End With
    Next
    '--- firmware bug: cannot print PLU discount on last 30th row
    If lCount > MAX_RECEIPT_ROWS Or (lCount = MAX_RECEIPT_ROWS And bIsLastDiscount) Then
        '--- count different VAT groups in PLUs
        For lRow = 1 To UBound(uCtx.GrpTotal)
            If uCtx.GrpTotal(lRow) <> 0 Then
                lTotal = lTotal + 1
                uCtx.GrpTotal(lRow) = 0
            End If
        Next
        '--- set extra rows to nonfiscal printing and calc GrpTotal by VAT groups
        lCount = 0
        For lRow = 0 To m_lRowCount - 1
            With m_uRow(lRow)
                If .RowType = ucsRowPlu Then
                    lCount = lCount + 1
                    If lCount > MAX_RECEIPT_ROWS - lTotal Then
                        .PrintRowType = IIf(.PrintRowType = ucsFscRetFiscal, ucsFscRetNonfiscal, ucsFscRetInvoiceNonfiscal)
                        uCtx.GrpTotal(.PluVatGroup) = Round(uCtx.GrpTotal(.PluVatGroup) + Round(.PluQuantity * .PluPrice, 2), 2)
                        uCtx.LastRow = m_uRow(lRow)
                    End If
                ElseIf .RowType = ucsRowDiscount And .DiscType = ucsFscDstPlu Then
                    If lCount > MAX_RECEIPT_ROWS - lTotal Then
                        .PrintRowType = IIf(.PrintRowType = ucsFscRetFiscal, ucsFscRetNonfiscal, ucsFscRetInvoiceNonfiscal)
                        dblTotal = Round(uCtx.LastRow.PluQuantity * uCtx.LastRow.PluPrice, 2)
                        If uCtx.LastRow.PluVatGroup > 0 Then
                            uCtx.GrpTotal(uCtx.LastRow.PluVatGroup) = Round(uCtx.GrpTotal(uCtx.LastRow.PluVatGroup) + Round(dblTotal * .DiscValue / 100#, 2), 2)
                        End If
                    End If
                ElseIf .RowType = ucsRowDiscount And .DiscType = ucsFscDstSubtotal Then
                    If lCount > MAX_RECEIPT_ROWS - lTotal Then
                        .PrintRowType = IIf(.PrintRowType = ucsFscRetFiscal, ucsFscRetNonfiscal, ucsFscRetInvoiceNonfiscal)
                        pvGetSubtotals lRow, uSum
                        For lIdx = 1 To UBound(uCtx.GrpTotal)
                            uCtx.GrpTotal(lIdx) = Round(uCtx.GrpTotal(lIdx) + Round(uSum.GrpTotal(lIdx) * .DiscValue / 100#, 2), 2)
                        Next
                    End If
                End If
            End With
        Next
        '--- find first payment row
        For lRow = 0 To m_lRowCount - 1
            If m_uRow(lRow).RowType = ucsRowPayment Then
                '--- move backward before any text lines
                Do While m_uRow(lRow - 1).RowType = ucsRowLine
                    lRow = lRow - 1
                Loop
                Exit For
            End If
        Next
        '--- append fiscal rows for GrpTotal by VAT groups
        For lIdx = 1 To UBound(uCtx.GrpTotal)
            If uCtx.GrpTotal(lIdx) <> 0 Then
                AddPLU 190 + lIdx, Printf(pvText(ucsTxtPluSell), Chr$(191 + lIdx)), uCtx.GrpTotal(lIdx), 1, lIdx, lRow
                lRow = lRow + 1
            End If
        Next
    End If
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Private Sub pvGetSubtotals(ByVal lRow As Long, uCtx As UcsContext)
    Const FUNC_NAME     As String = "pvGetSubtotals"
    Dim lIdx            As Long
    Dim lJdx            As Long
    Dim dblTotal        As Double
    Dim uEmpty          As UcsContext
    
    On Error GoTo EH
    uCtx = uEmpty
    For lIdx = 0 To lRow - 1
        With m_uRow(lIdx)
        If .RowType = ucsRowPlu Then
            dblTotal = Round(.PluQuantity * .PluPrice, 2)
            uCtx.GrpTotal(.PluVatGroup) = Round(uCtx.GrpTotal(.PluVatGroup) + dblTotal, 2)
            uCtx.LastRow = m_uRow(lIdx)
        ElseIf .RowType = ucsRowDiscount Then
            If .DiscType = ucsFscDstPlu Then
                dblTotal = Round(uCtx.LastRow.PluQuantity * uCtx.LastRow.PluPrice, 2)
                dblTotal = Round(dblTotal * .DiscValue / 100#, 2)
                If uCtx.LastRow.PluVatGroup > 0 Then
                    uCtx.GrpTotal(uCtx.LastRow.PluVatGroup) = Round(uCtx.GrpTotal(uCtx.LastRow.PluVatGroup) + dblTotal, 2)
                End If
            ElseIf .DiscType = ucsFscDstSubtotal Then
                For lJdx = 1 To UBound(uCtx.GrpTotal)
                    dblTotal = Round(uCtx.GrpTotal(lJdx) * .DiscValue / 100#, 2)
                    uCtx.GrpTotal(lJdx) = Round(uCtx.GrpTotal(lJdx) + dblTotal, 2)
                Next
            End If
        End If
        End With
    Next
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Private Function pvGetDeviceInfo(ByVal hComm As Long, ByVal lTimeout As Long, sProtocol As String, sModel As String) As Boolean
    Const FUNC_NAME     As String = "GetDeviceModel"
    Dim sResult         As String
    
    On Error GoTo EH
    sResult = pvSendCommand(hComm, lTimeout, ucsEltCmdInfoEcrParams)
    If LenB(sResult) Then
        sProtocol = STR_PROTOCOL_ELTRADE_ECR
        Select Case pvPeek(sResult, 2)
        Case 1
            sModel = "A100"
        Case 2
            sModel = "A300"
        Case 4
            sModel = "A500"
        Case 8
            sModel = "A800"
        Case 16
            sModel = "A600"
        Case Else
            sModel = "Model 0x" & Hex(pvPeek(sResult, 2))
        End Select
        Select Case pvPeek(sResult, 1)
        Case &H30, &H70, &HF0
            sResult = "70a"
        Case &HF4
            sResult = "70x"
        Case Else
            sResult = "Unknown 0x" & Hex(pvPeek(sResult, 1))
        End Select
        sModel = sModel & " (" & sResult & ")"
        '--- success
        pvGetDeviceInfo = True
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Initialize()
    m_hComm = INVALID_HANDLE_VALUE
    m_lTimeout = DEF_TIMEOUT
End Sub

Private Sub Class_Terminate()
    If m_hComm <> INVALID_HANDLE_VALUE Then
        Call CloseHandle(m_hComm)
        m_hComm = INVALID_HANDLE_VALUE
    End If
End Sub

'=========================================================================
' IDeviceProtocol interface
'=========================================================================

Private Function IDeviceProtocol_AddDiscount(ByVal DiscType As UcsFiscalDiscountTypeEnum, Value As Double) As Boolean
    IDeviceProtocol_AddDiscount = AddDiscount(DiscType, Value)
End Function

Private Function IDeviceProtocol_AddLine(Line As String, Optional Command As String, Optional ByVal WordWrap As Boolean = True) As Boolean
    IDeviceProtocol_AddLine = AddLine(Line, Command, WordWrap)
End Function

Private Function IDeviceProtocol_AddPayment(ByVal Number As UcsFiscalPaymentTypeEnum, Name As String, Amount As Double, Optional Rate As Double) As Boolean
    IDeviceProtocol_AddPayment = AddPayment(Number, Name, Amount, Rate)
End Function

Private Function IDeviceProtocol_AddPLU(Name As String, Price As Double, Quantity As Double, VatGroup As Long) As Boolean
    IDeviceProtocol_AddPLU = AddPLU(0, Name, Price, Quantity, VatGroup)
End Function

Private Function IDeviceProtocol_AutodetectDevices(Ports As Variant) As Variant
    IDeviceProtocol_AutodetectDevices = AutodetectDevices(Ports)
End Function

Private Function IDeviceProtocol_CancelReceipt() As Boolean
    IDeviceProtocol_CancelReceipt = CancelReceipt()
End Function

Private Function IDeviceProtocol_CashDebitCredit(OperatorCode As String, OperatorPassword As String, ByVal Value As Double) As Variant
    '--- not implemented
End Function

Private Function IDeviceProtocol_Connect() As Boolean
    IDeviceProtocol_Connect = Connect()
End Function

Private Function IDeviceProtocol_CopyLastReceipt() As Boolean
    IDeviceProtocol_CopyLastReceipt = CopyLastReceipt()
End Function

Private Function IDeviceProtocol_Disconnect() As Boolean
    IDeviceProtocol_Disconnect = Disconnect()
End Function

Private Function IDeviceProtocol_EndReceipt() As Boolean
    IDeviceProtocol_EndReceipt = EndReceipt()
End Function

Private Function IDeviceProtocol_GetCharsPerLine() As Long
    IDeviceProtocol_GetCharsPerLine = pvRowChars - 2
End Function

Private Function IDeviceProtocol_GetClock() As Date
    IDeviceProtocol_GetClock = GetClock()
End Function

Private Function IDeviceProtocol_GetCommandLog() As String
    IDeviceProtocol_GetCommandLog = m_sCommandLog
End Function

Private Function IDeviceProtocol_GetDefaultPassword(OperatorCode As String) As String
    '--- not implemented
End Function

Private Function IDeviceProtocol_GetDeviceModel() As String
    IDeviceProtocol_GetDeviceModel = GetDeviceModel()
End Function

Private Function IDeviceProtocol_GetDeviceProtocol() As String
    IDeviceProtocol_GetDeviceProtocol = GetDeviceProtocol()
End Function

Private Function IDeviceProtocol_GetFiscalNumber() As String
    IDeviceProtocol_GetFiscalNumber = GetFiscalNumber()
End Function

Private Function IDeviceProtocol_GetFooterText(ByVal Index As Long) As String
    IDeviceProtocol_GetFooterText = GetFooterText(Index)
End Function

Private Function IDeviceProtocol_GetHeaderText(ByVal Index As Long) As String
    IDeviceProtocol_GetHeaderText = GetHeaderText(Index)
End Function

Private Function IDeviceProtocol_GetLastReceiptNumber() As Long
    IDeviceProtocol_GetLastReceiptNumber = GetLastReceiptNumber()
End Function

Private Function IDeviceProtocol_GetLastError() As String
    IDeviceProtocol_GetLastError = LastError
End Function

Private Function IDeviceProtocol_GetLastInvoiceNumber() As String
    IDeviceProtocol_GetLastInvoiceNumber = GetLastInvoiceNumber()
End Function

Private Function IDeviceProtocol_GetPaymentExecuted() As Boolean
    IDeviceProtocol_GetPaymentExecuted = m_bFiscalPaymentExecuted
End Function

Private Function IDeviceProtocol_GetPaymentName(ByVal Index As Long) As String
    IDeviceProtocol_GetPaymentName = GetPaymentName(Index)
End Function

Private Function IDeviceProtocol_GetSerialNumber() As String
    IDeviceProtocol_GetSerialNumber = GetSerialNumber()
End Function

Private Function IDeviceProtocol_GetTaxCaption() As String
    IDeviceProtocol_GetTaxCaption = GetTaxCaption()
End Function

Private Function IDeviceProtocol_GetTaxNumber() As String
    IDeviceProtocol_GetTaxNumber = GetTaxNumber()
End Function

Private Function IDeviceProtocol_GetTotalsByVatGroups() As Variant
    IDeviceProtocol_GetTotalsByVatGroups = GetTotalsByVatGroups()
End Function

Private Function IDeviceProtocol_Init(Device As String, Optional ByVal Timeout As Long, Optional ByVal CashDeskNo As Long) As Boolean
    IDeviceProtocol_Init = Init(Device, Timeout, CashDeskNo)
End Function

Private Function IDeviceProtocol_OpenDrawer() As Boolean
    '--- not implemented
End Function

Private Function IDeviceProtocol_RunPeriodReport(ByVal ReportType As UcsFiscalReportsTypeEnum, ByVal StartDate As Date, ByVal EndDate As Date) As Boolean
    IDeviceProtocol_RunPeriodReport = RunPeriodReport(ReportType, StartDate, EndDate)
End Function

Private Function IDeviceProtocol_RunXReport(ByVal ReportType As UcsFiscalReportsTypeEnum) As Boolean
    IDeviceProtocol_RunXReport = RunXReport(ReportType)
End Function

Private Function IDeviceProtocol_RunZReport(ByVal ReportType As UcsFiscalReportsTypeEnum) As Boolean
    IDeviceProtocol_RunZReport = RunZReport(ReportType)
End Function

Private Sub IDeviceProtocol_SetClock(ByVal NewDate As Date)
    SetClock NewDate
End Sub

Private Sub IDeviceProtocol_SetLocalizedText(ByVal Index As UcsFiscalLocalizedIndexesEnum, Text As String)
    LocalizedText(Index) = Text
End Sub

Private Function IDeviceProtocol_StartReceipt(ByVal ReceiptType As UcsFiscalReceiptTypeEnum, OperatorCode As String, OperatorName As String, OperatorPassword As String, Optional InvNumber As String, Optional InvTaxNumber As String, Optional InvVatNumber As String, Optional InvCompany As String, Optional InvCity As String, Optional InvAddress As String, Optional InvMol As String, Optional InvReceiver As String) As Boolean
    IDeviceProtocol_StartReceipt = StartReceipt(ReceiptType, False, OperatorName, InvNumber, InvTaxNumber, InvVatNumber, InvCompany, InvCity, InvAddress, InvMol, InvReceiver)
End Function
