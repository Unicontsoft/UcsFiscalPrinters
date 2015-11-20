VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cZekaProtocol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=========================================================================
' $Header: /UcsFiscalPrinter/Src/cZekaProtocol.cls 41    20.11.15 16:39 Wqw $
'
'   Unicontsoft Fiscal Printers Project
'   Copyright (c) 2008-2015 Unicontsoft
'
'   Protocol handler for Tremol/Zeka cash registers & FP
'
' $Log: /UcsFiscalPrinter/Src/cZekaProtocol.cls $
' 
' 41    20.11.15 16:39 Wqw
' REF: optional param of EndReceipt impl of IDeviceProtocol interface
'
' 40    14.04.15 18:23 Wqw
' REF: impl ecr check and default password
'
' 39    6.04.15 15:16 Wqw
' REF: non-fiscal invoice total vat rounding
'
' 38    17.02.15 19:23 Wqw
' REF: err hanlding on port open
'
' 37    29.01.15 11:44 Wqw
' REF: open port on error clears persistent comm handle
'
' 36    28.01.15 15:34 Wqw
' REF: error messages
'
' 35    23.01.15 17:57 Wqw
' REF: uses persistent port global object
'
' 34    23.01.15 17:11 Wqw
' REF: uses persistent com port connection
'
' 33    6.01.15 17:49 Wqw
' REF: impl timeout on port open
'
' 32    17.12.14 16:09 Wqw
' REF: uses to ascii for string to byte array conversion, retries on
' handle eof (RDP serial ports)
'
' 31    31.07.14 16:06 Wqw
' REF: uses epsilon for floating point comparisons
'
' 30    25.04.14 22:53 Wqw
' REF: separate conf name constants
'
' 29    31.03.14 17:33 Wqw
' ADD: Sub SetLocalizedCommand, ClearLocalizedCommand
'
' 28    4.07.13 16:44 Wqw
' REF: auto detect reply text config
'
' 27    26.06.13 21:50 Wqw
' REF: if invalid hex dump response in err message
'
' 26    26.06.13 10:13 Wqw
' REF: err handling
'
' 25    25.06.13 17:37 Wqw
' REF: trim bogus prefix in response from device
'
' 24    25.06.13 12:12 Wqw
' REF: impl noop command
'
' 23    20.06.13 12:49 Wqw
' REF: handle custom payment types in AddPayment and GetPaymentName
'
' 22    18.06.13 17:18 Wqw
' ADD: Function GetTotalsByPayments, GetTotalsStats
'
' 21    14.03.13 16:28 Wqw
' REF: default operator passwords. reads single/multiple operators FP
' setting
'
' 20    26.02.13 19:38 Wqw
' REF: reduce logging on wait device
'
' 19    8.01.13 15:35 Wqw
' REF: wait device uses fast status
'
' 18    4.01.13 12:21 Wqw
' REF: err handling
'
' 17    3.01.13 17:15 Wqw
' REF: turn on delay printing by default
'
' 16    3.01.13 16:40 Wqw
' REF: impl resume ticket and command configs
'
' 15    16.11.12 18:52 Wqw
' REF: maintains fiscal payment executed
'
' 14    8.11.12 17:07 Wqw
' REF: pad function params
'
' 13    8.11.12 17:06 Wqw
' REF: default oper code/pass in CashDebitCredit
'
' 12    19.10.12 0:36 Wqw
' REF: config defaults
'
' 11    10.10.12 14:59 Wqw
' REF: GetTotalsByVatGroups retval array size fixed
'
' 10    9.10.12 15:20 Wqw
' ADD: config values DetailedReceipt & DelayPrint
'
' 9     8.10.12 11:16 Wqw
' REF: correct registers in CashDebitCredit
'
' 8     5.10.12 14:20 Wqw
' ADD: Function CashDebitCredit
'
' 7     29.08.12 15:05 Wqw
' REF: impl config values instead of constants
'
' 6     6.08.12 18:39 Wqw
' REF: impl AutodetectDevices
'
' 5     23.03.12 15:26 Wqw
' ADD: GetTotalsByVatGroups. REF: cancel receipt before printing reports
'
' 4     8.12.11 15:48 Wqw
' REF: params na GetDefaultPassword
'
' 3     11.11.11 20:27 Wqw
' REF: zero quantity is handled separately in PLU print
'
' 2     13.10.11 10:54 Wqw
' REF: Init params
'
' 1     9.08.11 23:26 Wqw
' Initial implementation
'=========================================================================
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "cZekaProtocol"
Implements IDeviceProtocol

#If False Then
    Private Type ucsFscRetInvoiceNonfiscal
    End Type
#End If

'=========================================================================
' Public events
'=========================================================================

Event CommandComplete(ByVal lCmd As Long, sData As String, sResult As String)

'=========================================================================
' Public enums
'=========================================================================

Public Enum UcsZekaCommandsEnum
    ucsZekCmdNoop = -1
    '--- fast status
    ucsZekCmdFastIsOnline = &H4
    ucsZekCmdFastIsReady = &H5
    ucsZekCmdFastStatus = &H9
    [_ucsZekCmdFastMax] = &H1F
    '--- init
    ucsZekCmdInitOpenDrawer = &H2A
    ucsZekCmdInitTaxRates = &H42
    ucsZekCmdInitDecimals = &H43
    ucsZekCmdInitPaymentType = &H44
    ucsZekCmdInitParameters = &H45
    ucsZekCmdInitDepartment = &H47
    ucsZekCmdInitDateTime = &H48
    ucsZekCmdInitHeaderFooter = &H49
    ucsZekCmdInitOperator = &H4A
    ucsZekCmdInitItem = &H4B
    ucsZekCmdInitLogo = &H4C
    ucsZekCmdInitLogoByNum = &H4D
    ucsZekCmdInitInvoiceNo = &H50
    ucsZekCmdInitClient = &H52
    '--- transaction
    ucsZekCmdNonFiscalOpen = &H2E
    ucsZekCmdNonFiscalClose = &H2F
    ucsZekCmdFiscalOpen = &H30
    ucsZekCmdFiscalSell = &H31
    ucsZekCmdFiscalSubtotal = &H33
    ucsZekCmdFiscalPayment = &H35
    ucsZekCmdFiscalPayAndClose = &H36
    ucsZekCmdPrintText = &H37
    ucsZekCmdFiscalClose = &H38
    ucsZekCmdFiscalCancel = &H39
    ucsZekCmdPrintDuplicate = &H3A
    ucsZekCmdAdminCashDebitCredit = &H3B
    '--- info
    ucsZekCmdInfoStatus = &H20
    ucsZekCmdInfoVersion = &H21
    ucsZekCmdInfoLogo = &H23
    ucsZekCmdInfoClient = &H53
    ucsZekCmdInfoDiagnostics = &H60
    ucsZekCmdInfoBulstat = &H61
    ucsZekCmdInfoTaxRates = &H62
    ucsZekCmdInfoDecimals = &H63
    ucsZekCmdInfoPaymentTypes = &H64
    ucsZekCmdInfoParameters = &H65
    ucsZekCmdInfoDepartment = &H67
    ucsZekCmdInfoDateTime = &H68
    ucsZekCmdInfoHeaderFooter = &H69
    ucsZekCmdInfoOperator = &H6A
    ucsZekCmdInfoItem = &H6B
    ucsZekCmdInfoTotalsByVatGroups = &H6D
    ucsZekCmdInfoRegisters = &H6E
    ucsZekCmdInfoInvoiceNo = &H70
    ucsZekCmdInfoLastReceipt = &H71
    ucsZekCmdInfoTransaction = &H72
    ucsZekCmdInfoReportLastDate = &H73
    '--- print
    ucsZekCmdPrintDiagnostics = &H22
    ucsZekCmdPrintReportDailyDepartments = &H76
    ucsZekCmdPrintReportSpecial = &H77
    ucsZekCmdPrintReportByNumberDetailed = &H78
    ucsZekCmdPrintReportByNumberShort = &H79
    ucsZekCmdPrintReportByDateDetailed = &H7A
    ucsZekCmdPrintReportByDateShort = &H7B
    ucsZekCmdPrintReportDaily = &H7C
    ucsZekCmdPrintReportByOperators = &H7D
    ucsZekCmdPrintReportDailyItems = &H7E
End Enum

'=========================================================================
' API
'=========================================================================

'--- for CreateFile
Private Const INVALID_HANDLE_VALUE          As Long = -1
Private Const ERROR_ACCESS_DENIED           As Long = 5
'--- for WaitCommEvent
Private Const EV_RXCHAR                     As Long = &H1                '  Any Character received
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

Private Const STR_INTERNAL              As String = "No device info set|CreateFile failed: %1|SetCommTimeouts failed: %1|BuildCommDCB failed: %1|SetCommState failed: %1|WriteFile failure: %1|Timeout waiting for response|ReadFile failed: %1|Invalid message format or checksum (NACK)|Already connected|Not connected|Invalid response length: %1|Response missing preamble: %1|WaitCommEvent failed: %1|No receipt started|Error in line %1: %2|Cannot cancel previous receipt|Invalid report type|Device not online|Device not ready| or cannot set operator password| or invalid operator password"
Private Const STR_STATUSES              As String = "Unknown error|Invalid command|Illegal command|Command denied because of uncommited report|Syntax error|Input register overflow|Input register is zero|Missing transaction for void|Insufficient subtotal|Tax groups conflict"
Private Const STR_ERRORS                As String = "Unknown error|Out of paper|Daily registers overflow|Clock not adjusted|Fiscal receipt open|Receipt not fully paid|Non-fiscal receipt open|Receipt already paid out|Read-only fiscal memory|Bad password or command not allowed|Display missing|24 hours without daily report|Printer overheated|Power down|Electronic journal overflow|Not enough conditions met"
Private Const STR_RECEIPT_TEXTS         As String = "Cash|Card|Cheque|Surcharge%1|Discount%1|Total|Subtotal|Invoice No %1|Sales %1|Unit Price|Quantity|Value*%1|Sum Group %1|VAT*%1=%2|Netto Sum|Total sum|%1 items|1 item|Change|Seller: |Receiver: |Buyer: |Tax No: |VAT No: |EIC;Citizen #;Foreigner #;Official #|VOID"
Private Const DEF_AUTODETECTTIMEOUT     As Long = 50
Private Const DEF_DEVICEINFOTIMEOUT     As Long = 1000
Private Const DEF_AUTODETECTSPEEDS      As String = "115200|38400|9600|19200|57600"
Private Const DEF_TIMEOUT               As Long = 2000
Private Const DEF_PRINT_TIMEOUT         As Long = 30000
Private Const DEF_OPERNO                As String = "1"
Private Const DEF_ECR_PASSWORD          As String = "0"
Private Const DEF_FP_PASSWORD           As String = "0000"
Private Const DEF_MAX_RECEIPT_ROWS      As Long = 120
Private Const DEF_MIN_DISCOUNT          As Double = -100
Private Const DEF_MAX_DISCOUNT          As Double = 100
Private Const DEF_PING_TIMEOUT          As Long = 200
Private Const DEF_LINE_WIDTH            As Long = 32
Private Const DEF_MAX_OPERCODE          As Long = 2
Private Const DEF_MAX_OPERNAME          As Long = 20
Private Const DEF_MAX_OPERPASSWORD      As Long = 4
Private Const DEF_MAX_COMPANY           As Long = 26
Private Const DEF_MAX_MOL               As Long = 16
Private Const DEF_MAX_VATNUMBER         As Long = 13
Private Const DEF_MAX_TAXNUMBER         As Long = 13
Private Const DEF_MAX_CITY              As Long = 30
Private Const DEF_MAX_INVOICENO         As Long = 10
Private Const DEF_MAX_PLUNAME           As Long = 36
Private Const DEF_MAX_COMMENT           As Long = 34
Private Const DEF_DELIM                 As String = ";"
Private Const LNG_STATUS_READY          As Long = &H40
Private Const LNG_STATUS_IS_BUSY        As Long = &H41
Private Const LNG_STATUS_OUT_OF_PAPER   As Long = &H42
Private Const LNG_STATUS_OUT_OF_PAPER_LEGACY As Long = &H7
Private Const LNG_STATUS_OUT_OF_PAPER_AND_IS_BUSY As Long = &H43
Private Const LNG_STATUS_OVERHEATED     As Long = &H44

Private m_sDevice                   As String
Private m_hComm                     As Long
Private m_lTimeout                  As Long
Private m_lCashDeskNo               As Long
Private m_lSequence                 As Long
Private m_eLastErrNumber            As UcsFiscalErrorsEnum
Private m_sLastError                As String
Private m_uRow()                    As UcsRowData
Private m_lRowCount                 As Long
Private m_sCommandLog               As String
Private m_vInfoDiagnostics          As Variant
Private m_eError                    As Long
Private m_eStatus                   As Long
Private m_sPaymentTypes             As String
Private m_uConfig                   As UcsConfigValues
Private m_uCtxPrint                 As UcsContext

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
    ucsErrResponseMissingPreamble
    ucsErrWaitCommEventFailed
    ucsErrNoReceiptStarted
    ucsErrErrorInLine
    ucsErrCannotCancelReceipt
    ucsErrInvalidReportType
    ucsErrDeviceNotOnline
    ucsErrDeviceNotReady
    ucsErrPasswordNotSet
    ucsErrInvalidPassword
End Enum

Private Enum UcsDeviceErrorsAndStatuses
    '--- device errors
    ucsErrOutOfPaper = 1
    ucsErrPrinterOverheated = 13
    '--- device statuses
    ucsErrIllegalCommand = 2
End Enum

Private Enum UcsReceiptTextsEnum
    ucsTxtPaymentCash
    ucsTxtPaymentCard
    ucsTxtPaymentCheque
    ucsTxtSurcharge
    ucsTxtDiscount
    ucsTxtTotal
    ucsTxtSubTotal
    ucsTxtHeaderInvoiceNo
    ucsTxtPluSales
    ucsTxtPluPrice
    ucsTxtPluQuantity
    ucsTxtPluTotal
    ucsTxtGroupTotal
    ucsTxtGroupVat
    ucsTxtGroupNetto
    ucsTxtFooterTotal
    ucsTxtFooterArticles
    ucsTxtFooterArticleSingle
    ucsTxtFooterChange
    ucsTxtFooterSeller
    ucsTxtFooterReceiver
    ucsTxtFooterBuyer
    ucsTxtFooterTaxNo
    ucsTxtFooterVatNo
    ucsTxtBulstatCaption
    ucsTxtPluVoid
End Enum

Private Enum UcsProtocolMarkers
    ZFP_STX = &H2               '// start transaction
    ZFP_PING = &H4              '// ping!
    ZFP_BUSY = &H5              '// is busy
    ZFP_ACK = &H6
    ZFP_ETX = &HA               '// end transaction
    ZFP_RETRY = &HE
    ZFP_NACK = &H15
    ZFP_BASEDATA = &H20
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
    InitOperatorCode    As String
    InitOperatorName    As String
    InitOperatorPassword As String
    InitInvData         As Variant
    InitCashDeskNo      As Long
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
    PrintRowType        As UcsFiscalReceiptTypeEnum
End Type

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

Private Type UcsContext
    GrpTotal(1 To 8)    As Double
    Paid                As Double
    PluCount            As Long
    PmtPrinted          As Boolean
    ChangePrinted       As Boolean
    Row                 As Long
End Type

Private Type UcsAutodetect
    hComm               As Long
    WriteOver           As OVERLAPPED
    ReadOver            As OVERLAPPED
    Buffer              As String
End Type

Private Type UcsConfigValues
    IsEcr               As VbTriState
    RowChars            As Long
    MaxReceiptRows      As Long
    MinDiscount         As Double
    MaxDiscount         As Double
    PingTimeout         As Long
    DetailedReceipt     As Boolean
    DelayPrint          As Boolean
    MaxOperCode         As Long
    MaxOperName         As Long
    MaxOperPassword     As Long
    MaxCompany          As Long
    MaxMol              As Long
    MaxVatNumber        As Long
    MaxTaxNumber        As Long
    MaxCity             As Long
    MaxInvoiceNo        As Long
    MaxPluName          As Long
    MaxComment          As Long
    ConfigCommands      As Collection
    LocalizedCommands   As Collection
    LocalizedText(0 To [_ucsFscLciMax] - 1) As Variant
End Type

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

Property Get Timeout() As Long
    Timeout = m_lTimeout
End Property

Property Let Timeout(ByVal lValue As Long)
    m_lTimeout = lValue
End Property

Property Get StatusNo() As Long
    StatusNo = m_eStatus
End Property

Property Get StatusText() As String
    If m_eStatus > 0 Then
        StatusText = At(m_uConfig.LocalizedText(ucsFscLciPrinterStatuses), m_eStatus, At(m_uConfig.LocalizedText(ucsFscLciPrinterStatuses), 0))
    End If
End Property

Property Get ErrorNo() As Long
    ErrorNo = m_eError
End Property

Property Get ErrorText() As String
    If m_eError > 0 Then
        ErrorText = At(m_uConfig.LocalizedText(ucsFscLciPrinterErrors), m_eError, At(m_uConfig.LocalizedText(ucsFscLciPrinterErrors), 0))
    End If
End Property

Property Get LastErrNumber() As UcsFiscalErrorsEnum
    LastErrNumber = m_eLastErrNumber
End Property

Property Get LastError() As String
    LastError = m_sLastError
End Property

Property Get Device() As String
    Device = m_sDevice
End Property

Property Get IsConnected() As Boolean
    IsConnected = (m_hComm <> INVALID_HANDLE_VALUE)
End Property

Property Get LocalizedText(ByVal eIdx As UcsFiscalLocalizedIndexesEnum) As String
    LocalizedText = Join(m_uConfig.LocalizedText(eIdx), "|")
End Property

Property Let LocalizedText(ByVal eIdx As UcsFiscalLocalizedIndexesEnum, sValue As String)
    m_uConfig.LocalizedText(eIdx) = Split(sValue, "|")
End Property

Property Get IsEcr() As Boolean
    Const FUNC_NAME     As String = "IsEcr [get]"
    Const CONF_NAME     As String = "IsEcr"
    Dim eCmd            As UcsIclCommandsEnum
    Dim sParam          As String
    Dim lTimeout        As Long
    Dim lTypeIndex      As Long
    Dim vInfoVersion    As Variant
    
    On Error GoTo EH
    If m_uConfig.IsEcr = vbUseDefault Then
        pvGetCommandConfig CONF_NAME, "InfoVersion", eCmd, ucsZekCmdInfoVersion, sParam, vbNullString, _
            "Timeout", lTimeout, 100, _
            "TypeIndex", lTypeIndex, 0
        vInfoVersion = Split(pvSendCommand(m_hComm, lTimeout, eCmd, sParam), DEF_DELIM)
        m_uConfig.IsEcr = IIf(At(vInfoVersion, lTypeIndex) = "1", vbTrue, vbFalse)
    End If
    IsEcr = (m_uConfig.IsEcr = vbTrue)
    Exit Property
EH:
    RaiseError FUNC_NAME
End Property

Private Property Get pvRowChars() As Long
    pvRowChars = pvZfplibValue(m_uConfig.RowChars, "FPLineWidth", DEF_LINE_WIDTH)
    pvRowChars = pvRowChars - 2
End Property

Private Property Get pvPingTimeout() As Long
    pvPingTimeout = pvZfplibValue(m_uConfig.PingTimeout, "PingTimeout", DEF_PING_TIMEOUT)
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function AutodetectDevices(vPorts As Variant) As Variant
    Const FUNC_NAME     As String = "AutodetectDevices"
    Dim lIdx            As Long
    Dim sProbeText      As String
    Dim sReplyText      As String
    Dim lTimeout        As Long
    Dim lDeviceInfoTimeout As Long
    Dim vRetVal()       As Variant
    Dim vSpeed          As Variant
    Dim aPorts()        As UcsAutodetect
    Dim hEvents()       As Long
    Dim lEventsCount    As Long
    Dim lResult         As Long
    Dim lActual         As Long
    Dim sProtocol       As String
    Dim sModel          As String
    Dim sError          As String
    Dim vSpeeds         As Variant
    
    On Error GoTo EH
    sProbeText = GetConfigValue(STR_PROTOCOL_ZEKA_FP, "ProbeText", Chr$(ucsZekCmdFastIsReady))
    sReplyText = GetConfigValue(STR_PROTOCOL_ZEKA_FP, "ReplyText", Chr$(ucsZekCmdFastIsReady))
    lTimeout = GetConfigNumber(STR_PROTOCOL_ZEKA_FP, "AutodetectTimeout", DEF_AUTODETECTTIMEOUT)
    lDeviceInfoTimeout = GetConfigNumber(STR_PROTOCOL_ZEKA_FP, "DeviceInfoTimeout", DEF_DEVICEINFOTIMEOUT)
    AssignVariant vSpeeds, GetConfigValue(STR_PROTOCOL_ZEKA_FP, "AutodetectSpeeds", Split(DEF_AUTODETECTSPEEDS, "|"))
    If IsObject(vSpeeds) Then
        vSpeeds = vSpeeds.Items()
    End If
    vRetVal = vPorts
    For Each vSpeed In vSpeeds
        ReDim aPorts(0 To UBound(vPorts)) As UcsAutodetect
        ReDim hEvents(0 To UBound(vPorts)) As Long
        lEventsCount = 0
        For lIdx = 0 To UBound(aPorts)
            With aPorts(lIdx)
                If Not IsArray(vRetVal(lIdx)) And LenB(At(vPorts, lIdx)) <> 0 Then
                    sError = vbNullString
                    .hComm = pvOpenPort(At(vPorts, lIdx) & "," & vSpeed, 0, Error:=sError)
                    If .hComm <> INVALID_HANDLE_VALUE And LenB(sError) = 0 Then
                        .WriteOver.hEvent = CreateEvent(0, 1, 0, 0)
                        .ReadOver.hEvent = CreateEvent(0, 1, 0, 0)
                    Else
                        PortWrapper.CloseComPort .hComm
                        .hComm = 0
                    End If
                End If
            End With
        Next
        For lIdx = 0 To UBound(aPorts)
            With aPorts(lIdx)
                If .hComm <> 0 Then
                    Call PurgeComm(.hComm, PURGE_TXCLEAR Or PURGE_RXCLEAR)
                    lResult = WriteFile(.hComm, ByVal sProbeText, Len(sProbeText), lActual, .WriteOver)
                    If lResult = 0 And Err.LastDllError = ERROR_IO_PENDING Then
                        hEvents(lEventsCount) = .WriteOver.hEvent
                        lEventsCount = lEventsCount + 1
                    End If
                End If
            End With
        Next
        If lEventsCount > 0 Then
'            lResult = WaitForMultipleObjects(lEventsCount, hEvents(0), 1, lTimeout)
            Call Sleep(lTimeout)
            lEventsCount = 0
        End If
        For lIdx = 0 To UBound(aPorts)
            With aPorts(lIdx)
                If .hComm <> 0 Then
                    .Buffer = String$(1000, 0)
                    lResult = ReadFile(.hComm, ByVal StrPtr(.Buffer), LenB(.Buffer), lActual, .ReadOver)
                    If lResult = 0 And Err.LastDllError = ERROR_IO_PENDING Then
                        hEvents(lEventsCount) = .ReadOver.hEvent
                        lEventsCount = lEventsCount + 1
                    Else
                        .Buffer = Left$(StrConv(.Buffer, vbUnicode), lActual)
                    End If
                End If
            End With
        Next
        If lEventsCount > 0 Then
            lResult = WaitForMultipleObjects(lEventsCount, hEvents(0), 1, lTimeout)
            lEventsCount = 0
        End If
        For lIdx = 0 To UBound(aPorts)
            With aPorts(lIdx)
                If .hComm <> 0 Then
                    If InStr(.Buffer, Chr$(0)) > 0 Then
                        .Buffer = StrConv(.Buffer, vbUnicode)
                        .Buffer = Left$(.Buffer, InStr(.Buffer, Chr$(0)) - 1)
                    End If
                    If LenB(.Buffer) <> 0 Then
                        If Right$(.Buffer, Len(sReplyText)) = sReplyText Then
                            If pvGetDeviceInfo(.hComm, lDeviceInfoTimeout, Empty, sProtocol, sModel) Then
                                vRetVal(lIdx) = Array(At(vPorts, lIdx), vSpeed, sProtocol, sModel)
                            End If
                        Else
                            DebugLog FUNC_NAME, "Unknown reply: 0x" & ToHexDump(.Buffer)
                        End If
                    End If
                End If
            End With
        Next
        For lIdx = 0 To UBound(aPorts)
            With aPorts(lIdx)
                PortWrapper.CloseComPort .hComm
                .hComm = 0
                If .WriteOver.hEvent <> 0 Then
                    Call CloseHandle(.WriteOver.hEvent)
                End If
                If .ReadOver.hEvent <> 0 Then
                    Call CloseHandle(.ReadOver.hEvent)
                End If
            End With
        Next
    Next
    AutodetectDevices = vRetVal
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function Init( _
            sDevice As String, _
            Optional ByVal Timeout As Long, _
            Optional ByVal CashDeskNo As Long) As Boolean
    Const FUNC_NAME     As String = "Init"
    Const CONF_NAME     As String = FUNC_NAME
    Dim vSplit          As Variant
    Dim sSerial         As String
    Dim eCmd            As UcsZekaCommandsEnum
    Dim sParam          As String
    Dim lSerialIndex    As Long
    
    On Error GoTo EH
    DebugLog FUNC_NAME, "sDevice=" & sDevice & ", Timeout=" & Timeout
    pvSetLastError vbNullString
    PortWrapper.CloseComPort m_hComm
    m_hComm = INVALID_HANDLE_VALUE
    m_lTimeout = IIf(Timeout <= 0, DEF_TIMEOUT, Timeout)
    m_lCashDeskNo = IIf(CashDeskNo <= 0, 1, CashDeskNo)
    vSplit = Split(sDevice, ",")
    m_sDevice = Trim$(At(vSplit, 0, "COM1")) & "," & C_Lng(At(vSplit, 1, "9600")) & "," & At(vSplit, 2) & "," & C_Lng(At(vSplit, 3, "8")) & "," & _
        IIf(UCase$(At(vSplit, 4, "N")) = "Y", "Y", "N") & "," & C_Lng(At(vSplit, 5, "1"))
    m_hComm = pvOpenPort(m_sDevice, m_lTimeout, m_sLastError)
    If LenB(m_sLastError) <> 0 Then
        GoTo QH
    End If
    If Not WaitDevice(-1) Then
        GoTo QH
    End If
    pvGetCommandConfig CONF_NAME, "InfoDiagnostics", eCmd, ucsZekCmdInfoDiagnostics, sParam, vbNullString, _
        "SerialIndex", lSerialIndex, 0
    m_vInfoDiagnostics = Split(pvSendCommand(m_hComm, m_lTimeout, eCmd, sParam, m_sLastError), DEF_DELIM)
    If LenB(m_sLastError) <> 0 Then
        GoTo QH
    End If
    '--- load config values
    sSerial = At(m_vInfoDiagnostics, lSerialIndex)
    With m_uConfig
        .IsEcr = vbUseDefault
        .RowChars = GetConfigNumber(sSerial, "RowChars", 0)
        .MaxReceiptRows = GetConfigNumber(sSerial, "MaxReceiptRows", DEF_MAX_RECEIPT_ROWS)
        .MinDiscount = GetConfigNumber(sSerial, "MinDiscount", DEF_MIN_DISCOUNT)
        .MaxDiscount = GetConfigNumber(sSerial, "MaxDiscount", DEF_MAX_DISCOUNT)
        .PingTimeout = GetConfigNumber(sSerial, "PingTimeout", 0)
        .DetailedReceipt = C_Bool(GetConfigValue(sSerial, "DetailedReceipt", False))
        .DelayPrint = C_Bool(GetConfigValue(sSerial, "DelayPrint", True))
        .MaxOperCode = GetConfigNumber(sSerial, "MaxOperCode", DEF_MAX_OPERCODE)
        .MaxOperName = GetConfigNumber(sSerial, "MaxOperName", DEF_MAX_OPERNAME)
        .MaxOperPassword = GetConfigNumber(sSerial, "MaxOperPassword", DEF_MAX_OPERPASSWORD)
        .MaxCompany = GetConfigNumber(sSerial, "MaxCompany", DEF_MAX_COMPANY)
        .MaxMol = GetConfigNumber(sSerial, "MaxMol", DEF_MAX_MOL)
        .MaxVatNumber = GetConfigNumber(sSerial, "MaxVatNumber", DEF_MAX_VATNUMBER)
        .MaxTaxNumber = GetConfigNumber(sSerial, "MaxTaxNumber", DEF_MAX_TAXNUMBER)
        .MaxCity = GetConfigNumber(sSerial, "MaxCity", DEF_MAX_CITY)
        .MaxInvoiceNo = GetConfigNumber(sSerial, "MaxInvoiceNo", DEF_MAX_INVOICENO)
        .MaxPluName = GetConfigNumber(sSerial, "MaxPluName", DEF_MAX_PLUNAME)
        .MaxComment = GetConfigNumber(sSerial, "MaxComment", DEF_MAX_COMMENT)
        Set .ConfigCommands = GetConfigCollection(sSerial, "Commands")
    End With
    '--- success
    Init = True
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function Connect() As Boolean
    Const FUNC_NAME     As String = "Connect"
    
    On Error GoTo EH
    DebugLog FUNC_NAME, "m_hComm=" & m_hComm
    pvSetLastError vbNullString
    If m_hComm <> INVALID_HANDLE_VALUE Then
        pvSetLastError pvInternal(ucsErrAlreadyConnected)
        GoTo QH
    End If
    m_hComm = pvOpenPort(m_sDevice, m_lTimeout, m_sLastError)
    '--- success
    Connect = True
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function Disconnect() As Boolean
    Const FUNC_NAME     As String = "Disconnect"
    
    On Error GoTo EH
    DebugLog FUNC_NAME, "m_hComm=" & m_hComm
    pvSetLastError vbNullString
    If m_hComm = INVALID_HANDLE_VALUE Then
        pvSetLastError pvInternal(ucsErrNotConnected)
        GoTo QH
    End If
    PortWrapper.CloseComPort m_hComm
    m_hComm = INVALID_HANDLE_VALUE
    '--- success
    Disconnect = True
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function SendCommand(ByVal eCmd As UcsZekaCommandsEnum, Optional Data As String) As String
    Const FUNC_NAME     As String = "SendCommand"
    
    On Error GoTo EH
    pvSetLastError vbNullString
    If eCmd = ucsZekCmdNoop Then
        Exit Function
    End If
    If Not WaitDevice() Then
        Exit Function
    End If
    SendCommand = pvSendCommand(m_hComm, m_lTimeout, eCmd, Data, m_sLastError)
    '--- raise event
    RaiseEvent CommandComplete(eCmd, Data, SendCommand)
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function WaitDevice(Optional ByVal Timeout As Long) As Boolean
    Const FUNC_NAME     As String = "WaitDevice"
    Const CONF_NAME     As String = FUNC_NAME
    Dim sError          As String
    Dim eCmd            As UcsZekaCommandsEnum
    Dim sParam          As String
    Dim sStatus         As String
    Dim dblTimer        As Double
    Dim sCurrent        As String
    Dim lIter           As Long
    Dim sDelay          As String
    
    On Error GoTo EH
    pvSetLastError vbNullString
    If Timeout >= 0 Then
        Timeout = Znl(Timeout, m_lTimeout)
        dblTimer = DateTimer
        Do While DateTimer - dblTimer < Timeout / 1000
            lIter = lIter + 1
            pvGetCommandConfig CONF_NAME, "FastIsOnline", eCmd, ucsZekCmdFastIsOnline, sParam, vbNullString, _
                "Status", sStatus, Chr$(ucsZekCmdFastIsOnline)
            If pvSendCommand(m_hComm, pvPingTimeout, eCmd, sParam, sError) = sStatus Then
                lIter = lIter + 1
                pvGetCommandConfig CONF_NAME, "FastStatus", eCmd, ucsZekCmdFastStatus, sParam, vbNullString, _
                    "ReadyStatus", sStatus, LNG_STATUS_READY, _
                    "Delay", sDelay, 40
                sCurrent = pvSendCommand(m_hComm, pvPingTimeout, eCmd, sParam, sError)
                If LenB(sCurrent) <> 0 Then
                    Select Case Asc(sCurrent)
                    Case C_Lng(sStatus) '--- LNG_STATUS_READY
                        WaitDevice = True
                        If lIter > 2 Then
                            DebugLog FUNC_NAME, "Ready after " & lIter
                        End If
                        GoTo QH
                    Case LNG_STATUS_OUT_OF_PAPER, LNG_STATUS_OUT_OF_PAPER_LEGACY, LNG_STATUS_OUT_OF_PAPER_AND_IS_BUSY
                        pvSetLastError At(m_uConfig.LocalizedText(ucsFscLciPrinterErrors), ucsErrOutOfPaper)
                        GoTo QH
                    Case LNG_STATUS_OVERHEATED
                        pvSetLastError At(m_uConfig.LocalizedText(ucsFscLciPrinterErrors), ucsErrPrinterOverheated)
                        GoTo QH
                    End Select
                End If
                Call Sleep(C_Lng(sDelay))
            End If
        Loop
    End If
    pvGetCommandConfig CONF_NAME, "FastIsOnline", eCmd, ucsZekCmdFastIsOnline, sParam, vbNullString, _
        "Status", sStatus, Chr$(ucsZekCmdFastIsOnline)
    If pvSendCommand(m_hComm, pvPingTimeout, eCmd, sParam, sError) <> sStatus Then
        pvSetLastError pvInternal(ucsErrDeviceNotOnline)
        GoTo QH
    End If
    If Timeout >= 0 Then
        pvGetCommandConfig CONF_NAME, "FastStatus", eCmd, ucsZekCmdFastStatus, sParam, vbNullString, _
            "ReadyStatus", sStatus, LNG_STATUS_READY
        sCurrent = pvSendCommand(m_hComm, pvPingTimeout, eCmd, sParam, sError)
        If LenB(sCurrent) = 0 Then
            pvSetLastError pvInternal(ucsErrDeviceNotReady)
            GoTo QH
        End If
        Select Case Asc(sCurrent)
        Case C_Lng(sStatus) '--- LNG_STATUS_READY
            WaitDevice = True
            GoTo QH
        Case LNG_STATUS_IS_BUSY
            pvSetLastError pvInternal(ucsErrDeviceNotReady)
        Case LNG_STATUS_OUT_OF_PAPER, LNG_STATUS_OUT_OF_PAPER_LEGACY, LNG_STATUS_OUT_OF_PAPER_AND_IS_BUSY
            pvSetLastError At(m_uConfig.LocalizedText(ucsFscLciPrinterErrors), ucsErrOutOfPaper)
        Case LNG_STATUS_OVERHEATED
            pvSetLastError At(m_uConfig.LocalizedText(ucsFscLciPrinterErrors), ucsErrPrinterOverheated)
        Case Else
            pvSetLastError "Unknown status: " & Hex(Asc(sCurrent))
        End Select
    Else
        WaitDevice = True
        GoTo QH
    End If
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function StartReceipt( _
            ByVal ReceiptType As UcsFiscalReceiptTypeEnum, _
            Optional OperatorCode As String, _
            Optional OperatorName As String, _
            Optional OperatorPassword As String, _
            Optional InvNumber As String, _
            Optional InvTaxNumber As String, _
            Optional InvVatNumber As String, _
            Optional InvCompany As String, _
            Optional InvCity As String, _
            Optional InvAddress As String, _
            Optional InvMol As String, _
            Optional InvReceiver As String) As Boolean
    Const FUNC_NAME     As String = "StartReceipt"
    Dim uCtxEmpty       As UcsContext
    Dim sCity           As String
    Dim sAddress        As String
    
    On Error GoTo EH
    m_uCtxPrint = uCtxEmpty
    ReDim m_uRow(0 To 10) As UcsRowData
    m_lRowCount = 0
    With m_uRow(pvAddRow())
        .RowType = ucsRowInit
        .InitReceiptType = LimitLong(ReceiptType, 1, [_ucsFscRetMax] - 1)
        .InitOperatorCode = Zn(SafeText(OperatorCode), DEF_OPERNO)
        .InitOperatorName = SafeText(OperatorName)
        .InitOperatorPassword = SafeText(OperatorPassword)
        If LenB(.InitOperatorPassword) = 0 Then
            .InitOperatorPassword = IDeviceProtocol_GetDefaultPassword(.InitOperatorCode)
        End If
        .InitCashDeskNo = m_lCashDeskNo
        SplitCgAddress Trim$(SafeText(InvCity)) & vbCrLf & Trim$(SafeText(InvAddress)), sCity, sAddress, 30
        .InitInvData = Array(SafeText(InvNumber), SafeText(InvTaxNumber), SafeText(InvVatNumber), _
            SafeText(InvCompany), sCity, sAddress, SafeText(InvMol), SafeText(InvReceiver))
    End With
    '--- success
    StartReceipt = True
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function AddPLU( _
            Name As String, _
            ByVal Price As Double, _
            Optional ByVal Quantity As Double = 1, _
            Optional ByVal VatGroup As Long = 2, _
            Optional ByVal BeforeIndex As Long) As Boolean
    Const FUNC_NAME     As String = "AddPLU"
    Dim uRow            As UcsRowData
    Dim bNegative       As Boolean
    
    On Error GoTo EH
    '--- sanity check
    If m_lRowCount = 0 Then
        pvSetLastError pvInternal(ucsErrNoReceiptStarted)
        GoTo QH
    End If
    With uRow
        .RowType = ucsRowPlu
        .PluName = RTrim$(SafeText(Name))
        bNegative = (Round(Price, 2) * Round(Quantity, 3) < -DBL_EPSILON)
        .PluPrice = IIf(bNegative, -1, 1) * Round(Abs(Price), 2)
        .PluQuantity = Round(Abs(Quantity), 3)
        .PluVatGroup = LimitLong(VatGroup, 1, 8)
        .PrintRowType = m_uRow(0).InitReceiptType
    End With
    pvInsertRow BeforeIndex, uRow
    '--- success
    AddPLU = True
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function AddLine( _
            Line As String, _
            Optional ByVal WordWrap As Boolean, _
            Optional ByVal BeforeIndex As Long) As Boolean
    Const FUNC_NAME     As String = "AddLine"
    Dim uRow            As UcsRowData
    
    On Error GoTo EH
    '--- sanity check
    If m_lRowCount = 0 Then
        pvSetLastError pvInternal(ucsErrNoReceiptStarted)
        GoTo QH
    End If
    With uRow
        .RowType = ucsRowLine
        .LineText = RTrim$(SafeText(Line))
        .LineWordWrap = WordWrap
        .PrintRowType = m_uRow(0).InitReceiptType
    End With
    pvInsertRow BeforeIndex, uRow
    '--- success
    AddLine = True
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function AddDiscount( _
            ByVal DiscType As UcsFiscalDiscountTypeEnum, _
            ByVal Value As Double) As Boolean
    Const FUNC_NAME     As String = "AddDiscount"
    Dim lIdx            As Long
    
    On Error GoTo EH
    '--- sanity check
    If m_lRowCount = 0 Then
        pvSetLastError pvInternal(ucsErrNoReceiptStarted)
        GoTo QH
    End If
    Select Case DiscType
    Case ucsFscDstTotal
        AddPLU Printf(IIf(Value > DBL_EPSILON, pvText(ucsTxtSurcharge), pvText(ucsTxtDiscount)), vbNullString), Value
    Case ucsFscDstPlu
        For lIdx = m_lRowCount - 1 To 0 Step -1
            With m_uRow(lIdx)
                If .RowType = ucsRowPlu Then
                    .DiscType = DiscType
                    .DiscValue = Round(Value, 2)
                    Exit For
                End If
            End With
        Next
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
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function AddPayment( _
            ByVal Number As UcsFiscalPaymentTypeEnum, _
            Name As String, _
            ByVal Amount As Double) As Boolean
    Const FUNC_NAME     As String = "AddPayment"
    
    On Error GoTo EH
    '--- sanity check
    If m_lRowCount = 0 Then
        pvSetLastError pvInternal(ucsErrNoReceiptStarted)
        GoTo QH
    End If
    If Number < 0 Then
        '--- custom payment types: 4 & 5
        Number = 3 - Number
    End If
    With m_uRow(pvAddRow())
        .RowType = ucsRowPayment
        .PmtType = LimitLong(Number, 1, 8)
        .PmtName = SafeText(Name)
        .PmtAmount = Round(Amount, 2)
        .PrintRowType = m_uRow(0).InitReceiptType
    End With
    '--- success
    AddPayment = True
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function EndReceipt(sResumeTicket As String) As Boolean
    Const FUNC_NAME     As String = "EndReceipt"
    Dim vSplit          As Variant
    Dim lIdx            As Long
    Dim lPos            As Long
    
    On Error GoTo EH
    '--- sanity check
    If m_lRowCount = 0 Then
        pvSetLastError pvInternal(ucsErrNoReceiptStarted)
        GoTo QH
    End If
    '--- restore context
    vSplit = Split(sResumeTicket, CHR1)
    With m_uCtxPrint
        For lIdx = LBound(.GrpTotal) To UBound(.GrpTotal)
            .GrpTotal(lIdx) = C_Dbl(At(vSplit, lPos)): lPos = lPos + 1
        Next
        .Paid = C_Dbl(At(vSplit, lPos)): lPos = lPos + 1
        .PluCount = C_Lng(At(vSplit, lPos)): lPos = lPos + 1
        .PmtPrinted = C_Bool(At(vSplit, lPos)): lPos = lPos + 1
        .ChangePrinted = C_Bool(At(vSplit, lPos)): lPos = lPos + 1
        .Row = C_Lng(At(vSplit, lPos)): lPos = lPos + 1
    End With
    '--- fix fiscal receipts with for more than m_uConfig.MaxReceiptRows PLUs
    If m_uRow(0).InitReceiptType = ucsFscRetFiscal Or m_uRow(0).InitReceiptType = ucsFscRetInvoiceFiscal Then
        pvConvertExtraRows
    End If
    '--- append final payment (total)
    With m_uRow(pvAddRow())
        .RowType = ucsRowPayment
        .PrintRowType = m_uRow(0).InitReceiptType
    End With
    EndReceipt = pvPrintReceipt(m_uCtxPrint, m_uRow, m_lRowCount)
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetResumeTicket() As String
    Const FUNC_NAME     As String = "GetResumeTicket"
    Dim lIdx            As Long
    
    On Error GoTo EH
    '--- sanity check
    If m_lRowCount = 0 Then
        pvSetLastError pvInternal(ucsErrNoReceiptStarted)
        GoTo QH
    End If
    '--- only fiscal receipts need resume ticket if payment processed
    Select Case m_uRow(0).InitReceiptType
    Case ucsFscRetFiscal, ucsFscRetInvoiceFiscal
        With m_uCtxPrint
            If .PmtPrinted Then
                For lIdx = LBound(.GrpTotal) To UBound(.GrpTotal)
                    GetResumeTicket = GetResumeTicket & .GrpTotal(lIdx) & CHR1
                Next
                GetResumeTicket = GetResumeTicket & .Paid & CHR1 & .PluCount & CHR1 & -.PmtPrinted & CHR1 & -.ChangePrinted & CHR1 & .Row
            End If
        End With
    End Select
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function CopyLastReceipt() As Boolean
    Const FUNC_NAME     As String = "CopyLastReceipt"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsZekaCommandsEnum
    Dim sParam          As String
    Dim sPrevReceiptNo  As String
    Dim sReceiptNo      As String
    Dim lPrintTimeout   As Long
    
    On Error GoTo EH
    sPrevReceiptNo = GetLastReceiptNumber()
    pvGetCommandConfig CONF_NAME, "PrintDuplicate", eCmd, ucsZekCmdPrintDuplicate, sParam, vbNullString, _
        "PrintTimeout", lPrintTimeout, 0
    SendCommand eCmd, sParam
    If Not LenB(m_sLastError) = 0 Then
        GoTo QH
    End If
    If Not WaitDevice(Znl(lPrintTimeout, DEF_PRINT_TIMEOUT)) Then
        GoTo QH
    End If
    sReceiptNo = GetLastReceiptNumber()
    If LenB(sReceiptNo) = 0 Or sReceiptNo = sPrevReceiptNo Then
        pvSetLastError At(m_uConfig.LocalizedText(ucsFscLciPrinterStatuses), ucsErrIllegalCommand)
        GoTo QH
    End If
    '--- success
    CopyLastReceipt = True
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetClock() As Date
    Const FUNC_NAME     As String = "GetClock"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsZekaCommandsEnum
    Dim sParam          As String
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoDateTime", eCmd, ucsZekCmdInfoDateTime, sParam, vbNullString
    GetClock = C_Date(SendCommand(eCmd, sParam))
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function SetClock(ByVal NewDate As Date) As Boolean
    Const FUNC_NAME     As String = "SetClock"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsZekaCommandsEnum
    Dim sParam          As String
    Dim sFormat         As String
    Dim lPrintTimeout   As Long
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InitDateTime", eCmd, ucsZekCmdInitDateTime, sParam, "%1", _
        "Format", sFormat, "dd-MM-yy hh:mm:ss", _
        "PrintTimeout", lPrintTimeout, 0
    sParam = Printf(sParam, Format$(NewDate, sFormat))
    SendCommand eCmd, sParam
    If LenB(m_sLastError) = 0 Then
        '--- success (or failure)
        SetClock = WaitDevice(Znl(lPrintTimeout, DEF_PRINT_TIMEOUT))
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function RunXReport(ByVal ReportType As UcsFiscalReportsTypeEnum) As Boolean
    Const FUNC_NAME     As String = "RunXReport"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsZekaCommandsEnum
    Dim sParam          As String
    Dim lPrintTimeout   As Long
    
    If Not CancelReceipt() Then
        GoTo QH
    End If
    On Error GoTo EH
    Select Case ReportType
    Case ucsFscRstDaily
        pvGetCommandConfig CONF_NAME, "PrintReportDaily", eCmd, ucsZekCmdPrintReportDaily, sParam, "X", _
            "PrintTimeout", lPrintTimeout, 0
    Case ucsFscRstDailyByItems
        pvGetCommandConfig CONF_NAME, "PrintReportDailyItems", eCmd, ucsZekCmdPrintReportDailyItems, sParam, "X", _
            "PrintTimeout", lPrintTimeout, 0
    Case ucsFscRstDailyByDepartment
        pvGetCommandConfig CONF_NAME, "PrintReportDailyDepartments", eCmd, ucsZekCmdPrintReportDailyDepartments, sParam, "X", _
            "PrintTimeout", lPrintTimeout, 0
    Case Else
        pvSetLastError pvInternal(ucsErrInvalidReportType)
        GoTo QH
    End Select
    SendCommand eCmd, sParam
    If LenB(m_sLastError) = 0 Then
        '--- success (or failure)
        RunXReport = WaitDevice(Znl(lPrintTimeout, DEF_PRINT_TIMEOUT))
    End If
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function RunZReport(ByVal ReportType As UcsFiscalReportsTypeEnum) As Boolean
    Const FUNC_NAME     As String = "RunZReport"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsZekaCommandsEnum
    Dim sParam          As String
    Dim lPrintTimeout   As Long
    
    If Not CancelReceipt() Then
        GoTo QH
    End If
    On Error GoTo EH
    Select Case ReportType
    Case ucsFscRstDaily
        pvGetCommandConfig CONF_NAME, "PrintReportDaily", eCmd, ucsZekCmdPrintReportDaily, sParam, "Z", _
            "PrintTimeout", lPrintTimeout, 0
    Case ucsFscRstDailyByItems
        pvGetCommandConfig CONF_NAME, "PrintReportDailyItems", eCmd, ucsZekCmdPrintReportDailyItems, sParam, "Z", _
            "PrintTimeout", lPrintTimeout, 0
    Case ucsFscRstDailyByDepartment
        pvGetCommandConfig CONF_NAME, "PrintReportDailyDepartments", eCmd, ucsZekCmdPrintReportDailyDepartments, sParam, "Z", _
            "PrintTimeout", lPrintTimeout, 0
    Case Else
        pvSetLastError pvInternal(ucsErrInvalidReportType)
        GoTo QH
    End Select
    SendCommand eCmd, sParam
    If LenB(m_sLastError) = 0 Then
        '--- success (or failure)
        RunZReport = WaitDevice(Znl(lPrintTimeout, DEF_PRINT_TIMEOUT))
    End If
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function RunPeriodReport(ByVal ReportType As UcsFiscalReportsTypeEnum, ByVal StartDate As Date, ByVal EndDate As Date) As Boolean
    Const FUNC_NAME     As String = "RunPeriodReport"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsZekaCommandsEnum
    Dim sParam          As String
    Dim sFormat         As String
    Dim lPrintTimeout   As Long
    
    If Not CancelReceipt() Then
        GoTo QH
    End If
    On Error GoTo EH
    Select Case ReportType
    Case ucsFscRstPeriodShort
        pvGetCommandConfig CONF_NAME, "PrintReportByDateShort", eCmd, ucsZekCmdPrintReportByDateShort, sParam, "%1;%2", _
            "Format", sFormat, "ddmmyy", _
            "PrintTimeout", lPrintTimeout, 0
    Case ucsFscRstPeriodDetailed
        pvGetCommandConfig CONF_NAME, "PrintReportByDateDetailed", eCmd, ucsZekCmdPrintReportByDateDetailed, sParam, "%1;%2", _
            "Format", sFormat, "ddmmyy", _
            "PrintTimeout", lPrintTimeout, 0
    Case Else
        pvSetLastError pvInternal(ucsErrInvalidReportType)
        GoTo QH
    End Select
    sParam = Printf(sParam, Format$(StartDate, sFormat), Format$(EndDate, sFormat))
    SendCommand eCmd, sParam
    If LenB(m_sLastError) = 0 Then
        '--- success (or failure)
        RunPeriodReport = WaitDevice(Znl(lPrintTimeout, DEF_PRINT_TIMEOUT))
    End If
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetLastReceiptNumber() As String
    Const FUNC_NAME     As String = "GetLastReceiptNumber"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsZekaCommandsEnum
    Dim sParam          As String
    Dim lIndex          As Long
    Dim vSplit          As Variant
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoLastReceipt", eCmd, ucsZekCmdInfoLastReceipt, sParam, vbNullString, _
        "Index", lIndex, 0
    vSplit = Split(SendCommand(eCmd, sParam), DEF_DELIM)
    If LenB(m_sLastError) = 0 Then
        GetLastReceiptNumber = At(vSplit, lIndex)
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetLastInvoiceNumber() As String
    Const FUNC_NAME     As String = "GetLastInvoiceNumber"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsZekaCommandsEnum
    Dim sParam          As String
    Dim lIndex          As Long
    Dim vSplit          As Variant
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoInvoiceNo", eCmd, ucsZekCmdInfoInvoiceNo, sParam, vbNullString, _
        "Index", lIndex, 0
    vSplit = Split(SendCommand(eCmd, sParam), DEF_DELIM)
    If LenB(m_sLastError) = 0 And IsNumeric(At(vSplit, lIndex)) Then
        GetLastInvoiceNumber = C_Dbl(At(vSplit, lIndex)) - 1
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetSerialNumber() As String
    Const FUNC_NAME     As String = "GetSerialNumber"
    Const CONF_NAME     As String = FUNC_NAME
    Dim lIndex          As Long
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoDiagnostics", 0, 0, vbNullString, vbNullString, _
        "Index", lIndex, 0
    GetSerialNumber = At(m_vInfoDiagnostics, lIndex)
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetFiscalNumber() As String
    Const FUNC_NAME     As String = "GetFiscalNumber"
    Const CONF_NAME     As String = FUNC_NAME
    Dim lIndex          As Long
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoDiagnostics", 0, 0, vbNullString, vbNullString, _
        "Index", lIndex, 1
    GetFiscalNumber = At(m_vInfoDiagnostics, lIndex)
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetTaxNumber() As String
    Const FUNC_NAME     As String = "GetTaxNumber"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsZekaCommandsEnum
    Dim sParam          As String
    Dim lIndex          As Long
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoBulstat", eCmd, ucsZekCmdInfoBulstat, sParam, vbNullString, _
        "Index", lIndex, 0
    GetTaxNumber = At(Split(SendCommand(eCmd, sParam), DEF_DELIM), lIndex)
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetTaxCaption() As String
    Const FUNC_NAME     As String = "GetTaxCaption"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsZekaCommandsEnum
    Dim sParam          As String
    Dim lIndex          As Long
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoBulstat", eCmd, ucsZekCmdInfoBulstat, sParam, vbNullString, _
        "Index", lIndex, 1
    lIndex = C_Lng(At(Split(SendCommand(eCmd, sParam), DEF_DELIM), lIndex))
    GetTaxCaption = At(Split(pvText(ucsTxtBulstatCaption), DEF_DELIM), lIndex)
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetHeaderText(ByVal Index As Long) As String
    Const FUNC_NAME     As String = "GetHeaderText"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsZekaCommandsEnum
    Dim sParam          As String
    Dim lIndex          As Long
    
    On Error GoTo EH
    If Index >= 1 And Index <= 6 Then
        pvGetCommandConfig CONF_NAME, "InfoHeaderFooter" & Index, eCmd, ucsZekCmdInfoHeaderFooter, sParam, C_Str(Index), _
            "Index", lIndex, 1
        GetHeaderText = At(Split(SendCommand(eCmd, sParam), DEF_DELIM), lIndex)
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetFooterText(ByVal Index As Long) As String
    Const FUNC_NAME     As String = "GetFooterText"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsZekaCommandsEnum
    Dim sParam          As String
    Dim lIndex          As Long
    
    On Error GoTo EH
    If Index >= 1 And Index <= 1 Then
        pvGetCommandConfig CONF_NAME, "InfoHeaderFooter" & Index, eCmd, ucsZekCmdInfoHeaderFooter, sParam, C_Str(7 + Index), _
            "Index", lIndex, 1
        GetFooterText = At(Split(SendCommand(eCmd, sParam), DEF_DELIM), lIndex)
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetPaymentName(ByVal Index As Long) As String
    Const FUNC_NAME     As String = "GetPaymentName"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsZekaCommandsEnum
    Dim sParam          As String
    Dim lIndex          As Long
    
    On Error GoTo EH
    If Index < 0 Then
        '--- custom payment types: 4 & 5
        Index = 3 - Index
    End If
    If Index >= 1 And Index <= 3 Then
        GetPaymentName = pvText(ucsTxtPaymentCash + Index - 1)
    ElseIf Index >= 4 And Index <= 5 Then
        pvGetCommandConfig CONF_NAME, "InfoPaymentTypes" & Index, eCmd, ucsZekCmdInfoPaymentTypes, sParam, vbNullString, _
            "Index", lIndex, Index - 1
        If LenB(m_sPaymentTypes) = 0 Then
            m_sPaymentTypes = SendCommand(eCmd, sParam)
        End If
        GetPaymentName = Trim$(At(Split(m_sPaymentTypes, DEF_DELIM), lIndex))
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function OpenDrawer() As Boolean
    Const FUNC_NAME     As String = "OpenDrawer"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsZekaCommandsEnum
    Dim sParam          As String
    Dim lPrintTimeout   As Long
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InitOpenDrawer", eCmd, ucsZekCmdInitOpenDrawer, sParam, vbNullString, _
        "PrintTimeout", lPrintTimeout, 0
    SendCommand eCmd, sParam
    If LenB(m_sLastError) = 0 Then
        '--- success (or failure)
        OpenDrawer = WaitDevice(Znl(lPrintTimeout, DEF_PRINT_TIMEOUT))
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function CancelReceipt() As Boolean
    Const FUNC_NAME     As String = "CancelReceipt"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eTransCmd       As UcsZekaCommandsEnum
    Dim sTransParam     As String
    Dim lTransIndex     As Long
    Dim lTransStatus    As Long
    Dim eCmd            As UcsZekaCommandsEnum
    Dim sParam          As String
    Dim lPrintTimeout   As Long
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoTransaction", eTransCmd, ucsZekCmdInfoTransaction, sTransParam, vbNullString, _
        "Index", lTransIndex, 0, _
        "Status", lTransStatus, 0
    If C_Lng(At(Split(SendCommand(eTransCmd, sTransParam), DEF_DELIM), lTransIndex)) <> lTransStatus Then
        pvGetCommandConfig CONF_NAME, "NonFiscalClose", eCmd, ucsZekCmdNonFiscalClose, sParam, vbNullString, _
            "PrintTimeout", lPrintTimeout, 0
        SendCommand eCmd, sParam
        If LenB(m_sLastError) <> 0 Then
            pvGetCommandConfig CONF_NAME, "FiscalCancel", eCmd, ucsZekCmdFiscalCancel, sParam, vbNullString
            SendCommand eCmd, sParam
            If LenB(m_sLastError) <> 0 Then
                pvGetCommandConfig CONF_NAME, "FiscalPayment", eCmd, ucsZekCmdFiscalPayment, sParam, "0;0;-"
                SendCommand eCmd, sParam
                pvGetCommandConfig CONF_NAME, "FiscalCancel", eCmd, ucsZekCmdFiscalCancel, sParam, vbNullString
                SendCommand eCmd, sParam
            End If
            If LenB(m_sLastError) <> 0 Then
                pvGetCommandConfig CONF_NAME, "FiscalPayAndClose", eCmd, ucsZekCmdFiscalPayAndClose, sParam, vbNullString
                SendCommand eCmd, sParam
                If LenB(m_sLastError) <> 0 Then
                    GoTo QH
                End If
            End If
        End If
        If Not WaitDevice(Znl(lPrintTimeout, DEF_PRINT_TIMEOUT)) Then
            GoTo QH
        End If
        '--- success
        CancelReceipt = C_Lng(At(Split(SendCommand(eTransCmd, sTransParam), DEF_DELIM), lTransIndex)) = lTransStatus
    Else
        CancelReceipt = True
    End If
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetDeviceProtocol() As String
    Const FUNC_NAME     As String = "GetDeviceProtocol"
    
    On Error GoTo EH
    pvGetDeviceInfo m_hComm, m_lTimeout, m_vInfoDiagnostics, GetDeviceProtocol, vbNullString
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetDeviceModel() As String
    Const FUNC_NAME     As String = "GetDeviceModel"
    
    On Error GoTo EH
    pvGetDeviceInfo m_hComm, m_lTimeout, m_vInfoDiagnostics, vbNullString, GetDeviceModel
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetDeviceStatus(CurrentStatus As String) As Boolean
    Const FUNC_NAME     As String = "GetDeviceStatus"
    
    On Error GoTo EH
    WaitDevice -1
    CurrentStatus = m_sLastError
    '--- success
    GetDeviceStatus = True
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetTotalsByPayments() As Variant
    Const FUNC_NAME     As String = "GetTotalsByPayments"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsZekaCommandsEnum
    Dim sParam          As String
    Dim vPayments       As Variant
    Dim vChanges        As Variant
    Dim vRetVal         As Variant
    Dim lIdx            As Long
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoRegisters4", eCmd, ucsZekCmdInfoRegisters, sParam, "4"
    vPayments = Split(SendCommand(eCmd, sParam), DEF_DELIM)
    pvGetCommandConfig CONF_NAME, "InfoRegisters6", eCmd, ucsZekCmdInfoRegisters, sParam, "6"
    vChanges = Split(SendCommand(eCmd, sParam), DEF_DELIM)
    If UBound(vPayments) > 0 Then
        ReDim vRetVal(0 To 4) As Double
        For lIdx = 0 To UBound(vRetVal)
            vRetVal(lIdx) = ParseSum(At(vPayments, lIdx + 1)) - ParseSum(At(vChanges, lIdx + 1))
        Next
        GetTotalsByPayments = vRetVal
    Else
        GetTotalsByPayments = EmptyVariantArray
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetTotalsByVatGroups() As Variant
    Const FUNC_NAME     As String = "GetTotalsByVatGroups"
    Const CONF_NAME     As String = FUNC_NAME
    Dim vVatGroups      As Variant
    Dim vTaxRates       As Variant
    Dim vRetVal         As Variant
    Dim lIdx            As Long
    Dim eCmd            As UcsZekaCommandsEnum
    Dim sParam          As String
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoTotalsByVatGroups", eCmd, ucsZekCmdInfoTotalsByVatGroups, sParam, vbNullString
    vVatGroups = Split(SendCommand(eCmd, sParam), DEF_DELIM)
    pvGetCommandConfig CONF_NAME, "InfoTaxRates", eCmd, ucsZekCmdInfoTaxRates, sParam, vbNullString
    vTaxRates = Split(Replace(SendCommand(eCmd, sParam), "%", vbNullString), ";")
    If UBound(vVatGroups) > 0 Then
        ReDim vRetVal(0 To 7) As Variant
        For lIdx = 0 To UBound(vRetVal)
            vRetVal(lIdx) = Array(ParseSum(At(vVatGroups, lIdx)), C_Dbl(At(vTaxRates, lIdx)))
        Next
        GetTotalsByVatGroups = vRetVal
    Else
        GetTotalsByVatGroups = EmptyVariantArray
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetTotalsStats() As Variant
    Const FUNC_NAME     As String = "GetTotalsStats"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsZekaCommandsEnum
    Dim sParam          As String
    Dim lIndex          As Long
    Dim vSplit          As Variant
    Dim lReceipts       As Long
    Dim dLastDate       As Date
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoRegisters1", eCmd, ucsZekCmdInfoRegisters, sParam, "1", _
        "Index", lIndex, 1
    vSplit = Split(SendCommand(eCmd, sParam), DEF_DELIM)
    lReceipts = C_Lng(At(vSplit, lIndex))
'    pvGetCommandConfig CONF_NAME, "InfoReportLastDate", eCmd, ucsZekCmdInfoReportLastDate, sParam, vbNullString, _
'        "Index", lIndex, 0
    pvGetCommandConfig CONF_NAME, "InfoRegisters4", eCmd, ucsZekCmdInfoRegisters, sParam, "5", _
        "Index", lIndex, 4
    vSplit = Split(SendCommand(eCmd, sParam), DEF_DELIM)
    dLastDate = C_Date(At(vSplit, lIndex))
    GetTotalsStats = Array(lReceipts, IIf(dLastDate = 0 Or dLastDate = #1/1/2008#, Empty, dLastDate))
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function CashDebitCredit( _
        Optional ByVal OperatorCode As String, _
        Optional ByVal OperatorPassword As String, _
        Optional ByVal PmtNumber As UcsFiscalPaymentTypeEnum, _
        Optional ByVal Value As Double, _
        Optional Comment As String) As Variant
    Const FUNC_NAME     As String = "CashDebitCredit"
    Const CONF_NAME     As String = FUNC_NAME
    Dim vSplit          As Variant
    Dim vRetVal         As Variant
    Dim eCmd            As UcsZekaCommandsEnum
    Dim sParam          As String
    Dim sFormat         As String
    Dim lIndex          As Long
    Dim lPrintTimeout   As Long
    
    On Error GoTo EH
    If Abs(Value) > DBL_EPSILON Then
        pvGetCommandConfig CONF_NAME, "InfoParameters", eCmd, ucsZekCmdInfoParameters, sParam, vbNullString, _
            "Index", lIndex, 9
        vSplit = Split(SendCommand(eCmd, sParam), DEF_DELIM)
        If C_Bool(At(vSplit, lIndex)) Then
            '--- use single operator
            OperatorCode = DEF_OPERNO
            OperatorPassword = IDeviceProtocol_GetDefaultPassword(DEF_OPERNO)
        End If
        pvGetCommandConfig CONF_NAME, "AdminCashDebitCredit", eCmd, ucsZekCmdAdminCashDebitCredit, sParam, "%1;%2;%3;%4@%5", _
            "Format", sFormat, "0.00", _
            "PrintTimeout", lPrintTimeout, 0
        sParam = Printf(sParam, Pad(Zn(OperatorCode, DEF_OPERNO), -m_uConfig.MaxOperCode), _
            Pad(Zn(OperatorPassword, IDeviceProtocol_GetDefaultPassword(Zn(OperatorCode, DEF_OPERNO))), m_uConfig.MaxOperPassword), _
            LimitLong(PmtNumber - 1, 0, 7), _
            SafeFormat(Value, sFormat), _
            Left$(Comment, m_uConfig.MaxComment))
        SendCommand eCmd, sParam
        If LenB(LastError) <> 0 Then
            vRetVal = EmptyDoubleArray
            GoTo QH
        End If
        If Not WaitDevice(Znl(lPrintTimeout, DEF_PRINT_TIMEOUT)) Then
            vRetVal = EmptyDoubleArray
            GoTo QH
        End If
    End If
    pvGetCommandConfig CONF_NAME, "InfoRegisters0", eCmd, ucsZekCmdInfoRegisters, sParam, "0"
    vSplit = Split(SendCommand(eCmd, sParam), DEF_DELIM)
    If LenB(LastError) <> 0 Then
        vRetVal = EmptyDoubleArray
        GoTo QH
    End If
    ReDim vRetVal(0 To 2) As Double
    vRetVal(0) = C_Dbl(At(vSplit, PmtNumber))
    pvGetCommandConfig CONF_NAME, "InfoRegisters2", eCmd, ucsZekCmdInfoRegisters, sParam, "2"
    vSplit = Split(SendCommand(eCmd, sParam), DEF_DELIM)
    vRetVal(1) = C_Dbl(At(vSplit, PmtNumber))
    pvGetCommandConfig CONF_NAME, "InfoRegisters3", eCmd, ucsZekCmdInfoRegisters, sParam, "3"
    vSplit = Split(SendCommand(eCmd, sParam), DEF_DELIM)
    vRetVal(2) = C_Dbl(At(vSplit, PmtNumber))
QH:
    CashDebitCredit = vRetVal
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Sub SetLocalizedCommand( _
            sFunction As String, _
            sKey As String, _
            Optional ByVal Cmd As UcsZekaCommandsEnum, _
            Optional Param As String, _
            Optional More As Variant)
    Dim sMerged         As String
    Dim lIdx            As Long
    
    If m_uConfig.LocalizedCommands Is Nothing Then
        Set m_uConfig.LocalizedCommands = New Collection
    End If
    sMerged = "\" & sFunction & IIf(LenB(sKey) <> 0, "\" & sKey, vbNullString)
    With m_uConfig.LocalizedCommands
        If Cmd <> 0 Then
            .Add Cmd, sMerged
        End If
        If LenB(Param) <> 0 Then
            .Add Param, sMerged & "Param"
        End If
        If IsArray(More) Then
            For lIdx = 0 To UBound(More) Step 2
                .Add More(lIdx + 1), sMerged & More(lIdx)
            Next
        End If
    End With
End Sub

Public Sub ClearLocalizedCommand(sFunction As String, sKey As String, Optional More As Variant)
    Dim sMerged         As String
    Dim lIdx            As Long
    
    sMerged = "\" & sFunction & IIf(LenB(sKey) <> 0, "\" & sKey, vbNullString)
    If SearchCollection(m_uConfig.LocalizedCommands, sMerged) Then
        m_uConfig.LocalizedCommands.Remove sMerged
    End If
    If SearchCollection(m_uConfig.LocalizedCommands, sMerged & "Param") Then
        m_uConfig.LocalizedCommands.Remove sMerged & "Param"
    End If
    If IsArray(More) Then
        For lIdx = 0 To UBound(More)
            If SearchCollection(m_uConfig.LocalizedCommands, sMerged & More(lIdx)) Then
                m_uConfig.LocalizedCommands.Remove sMerged & More(lIdx)
            End If
        Next
    End If
End Sub

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

Private Function pvSendCommand( _
            ByVal hComm As Long, _
            ByVal lTimeout As Long, _
            ByVal eCmd As UcsZekaCommandsEnum, _
            Optional Data As String, _
            Optional Error As String) As String
    Const FUNC_NAME     As String = "pvSendCommand"
    Dim sText           As String
    Dim baText()        As Byte
    Dim lIdx            As Long
    Dim lSum            As Long
    Dim lWritten        As Long
    Dim uOver           As OVERLAPPED
    Dim lSize           As Long
    Dim lResult         As Long
    Dim sBuffer         As String
    Dim lMask           As Long
    Dim lActual         As Long
    Dim bLogging        As Boolean
    
    On Error GoTo EH
    bLogging = (eCmd > [_ucsZekCmdFastMax])
    If bLogging Then
        DebugLog FUNC_NAME, "eCmd=0x" & Hex(eCmd) & IIf(LenB(Data) <> 0, ", Data=" & Data, vbNullString)
    End If
    uOver.hEvent = CreateEvent(0, 1, 0, 0)
RetrySend:
    '--- prepare command
    If eCmd <= [_ucsZekCmdFastMax] Then
        sText = Chr$(eCmd)
    ElseIf eCmd = ucsZekCmdInitLogo Or eCmd = ucsZekCmdInitLogoByNum Then
        sText = Chr$(ZFP_STX) & Chr$(&H39) & Chr$(&H37) & Chr$(eCmd) & Data
        lTimeout = 60000
    Else
        sText = Chr$(ZFP_STX) & Chr$(ZFP_BASEDATA + 3 + Len(Data)) & Chr$(m_lSequence) & Chr$(eCmd) & Data
        '--- increment sequence
        m_lSequence = IIf(m_lSequence >= &H9F, ZFP_BASEDATA, m_lSequence + 1)
        '--- calc checksum
        baText = ToAscii(sText)
        For lIdx = 1 To Len(sText) - 1
            lSum = lSum Xor baText(lIdx)
        Next
        '--- append checksum and terminator
        sText = sText & Chr$((lSum \ &H10 Mod &H10) + &H30) & Chr$((lSum \ &H1 Mod &H10) + &H30) & Chr$(ZFP_ETX)
    End If
    '--- async send
    Call PurgeComm(hComm, PURGE_TXCLEAR Or PURGE_RXCLEAR)
    baText = ToAscii(sText)
    lResult = WriteFile(hComm, baText(0), Len(sText), lWritten, uOver)
    If lResult = 0 Then
        If Err.LastDllError <> ERROR_IO_PENDING Then
            Error = Printf(pvInternal(ucsErrWriteFileFailure), GetApiErr(Err.LastDllError))
            GoTo QH
        End If
    End If
    lResult = WaitForMultipleObjects(1, uOver.hEvent, 0, lTimeout)
    If lResult <> WAIT_OBJECT_0 Then
        Call CancelIo(hComm)
        DebugLog FUNC_NAME, "eCmd=0x" & Hex(eCmd) & ", CancelIo after WriteFile, lResult=" & lResult & ", lTimeout=" & lTimeout
        If lResult = WAIT_TIMEOUT Then
            Error = pvInternal(ucsErrTimeoutWaitingForResponse)
        End If
        GoTo QH
    End If
    '--- async read response
    sText = vbNullString
    Call SetCommMask(hComm, EV_RXCHAR)
    GoTo InLoop
    Do
        lResult = WaitCommEvent(hComm, lMask, uOver)
        If lResult = 0 Then
            If Err.LastDllError <> ERROR_IO_PENDING Then
                DebugLog FUNC_NAME, "eCmd=0x" & Hex(eCmd) & ", WaitCommEvent, Err.LastDllError=" & Err.LastDllError
                Error = Printf(pvInternal(ucsErrWaitCommEventFailed), GetApiErr(Err.LastDllError))
                GoTo QH
            End If
            lResult = WaitForMultipleObjects(1, uOver.hEvent, 0, lTimeout)
        Else
            DebugLog FUNC_NAME, "eCmd=0x" & Hex(eCmd) & ", WaitCommEvent, lResult=" & lResult
            lResult = WAIT_OBJECT_0
        End If
        If lResult = WAIT_OBJECT_0 And (lMask And EV_RXCHAR) <> 0 Then
InLoop:
            Do
                sBuffer = String$(1000, 0)
                lActual = 0
                lResult = ReadFile(hComm, ByVal sBuffer, Len(sBuffer), lActual, uOver)
                If lResult = 0 Then
                    If Err.LastDllError = ERROR_HANDLE_EOF Then
                        lResult = WAIT_OBJECT_0
                    Else
                        If Err.LastDllError <> ERROR_IO_PENDING Then
                            DebugLog FUNC_NAME, "ReadFile, eCmd=0x" & Hex(eCmd) & ", Err.LastDllError=" & Err.LastDllError
                            Error = Printf(pvInternal(ucsErrReadFileFailed), GetApiErr(Err.LastDllError))
                            GoTo QH
                        End If
                        lResult = WaitForMultipleObjects(1, uOver.hEvent, 0, lTimeout)
                    End If
                Else
                    lResult = WAIT_OBJECT_0
                End If
                If lActual = 0 Then
                    Exit Do
                End If
                sText = sText & Left$(sBuffer, lActual)
            Loop
        End If
        If lResult <> WAIT_OBJECT_0 Then
            Call CancelIo(hComm)
            If bLogging Then
                DebugLog FUNC_NAME, "CancelIo after ReadFile, lResult=" & lResult & ", lTimeout=" & lTimeout
            End If
            If lResult = WAIT_TIMEOUT Then
                Error = pvInternal(ucsErrTimeoutWaitingForResponse)
            End If
            GoTo QH
        End If
RetryParse:
        If LenB(sText) <> 0 Then
            Select Case Asc(sText)
            Case ZFP_STX            ' &H2
                '--- check if response complete
                If Len(sText) > 6 And Asc(Right$(sText, 1)) = ZFP_ETX Then
                    lSize = Asc(Mid$(sText, 2, 1)) - ZFP_BASEDATA - 3
                    '--- check message length
                    If Len(sText) - 7 = lSize Then
                        pvSendCommand = Mid$(sText, 5, lSize)
                        If bLogging Then
                            DebugLog FUNC_NAME, "RetVal=" & pvSendCommand
                        End If
                    Else
                        Error = Printf(pvInternal(ucsErrInvalidResponseLength), "0x" & ToHexDump(sText))
                    End If
                    Exit Do
                End If
            Case ZFP_ACK            ' &H6
                If Len(sText) > 6 And Asc(Right$(sText, 1)) = ZFP_ETX Then
                    lSize = Len(sText) - 5
                    '--- parse respone
                    sText = Mid$(sText, 3, lSize)
                    If Len(sText) > 1 Then
                        m_eError = Asc(Mid$(sText, 1)) - &H30
                        m_eStatus = Asc(Mid$(sText, 2)) - &H30
                        If m_eError <> 0 Then
                            Error = ErrorText
                        Else
                            Error = StatusText
                        End If
                    Else
                        Error = Printf(pvInternal(ucsErrInvalidResponseLength), "0x" & ToHexDump(sText))
                    End If
                    Exit Do
                End If
            Case ZFP_RETRY          ' &HE
                Call CancelIo(hComm)
                DebugLog FUNC_NAME, "eCmd=0x" & Hex(eCmd) & ", Retry requested, sText=" & sText
                GoTo RetrySend
            Case ZFP_NACK           ' &H15
                Error = Printf(pvInternal(ucsErrInvalidMessageFormatOrChecksum), "0x" & ToHexDump(sText))
                Exit Do
            Case Else
                If eCmd <= [_ucsZekCmdFastMax] Then
                    pvSendCommand = Right$(sText, 1)
                Else
                    lIdx = InStr(sText, Chr$(ZFP_STX))
                    lIdx = LimitLong(InStr(sText, Chr$(ZFP_ACK)), , Znl(lIdx, 2147483647))
                    If lIdx > 1 Then
                        DebugLog FUNC_NAME, "Trim bogus symbols, Prefix=0x" & ToHexDump(Left$(sText, lIdx - 1))
                        sText = Mid$(sText, lIdx)
                        GoTo RetryParse
                    End If
                    Error = Printf(pvInternal(ucsErrResponseMissingPreamble), "0x" & ToHexDump(sText))
                End If
                Exit Do
            End Select
        End If
    Loop
QH:
    Call CloseHandle(uOver.hEvent)
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function pvInternal(ByVal lIdx As UcsInternalErrors) As String
    Const FUNC_NAME     As String = "pvInternal"
    
    On Error GoTo EH
    pvInternal = At(m_uConfig.LocalizedText(ucsFscLciInternalErrors), lIdx)
    If LenB(pvInternal) = 0 Then
        pvInternal = At(Split(STR_INTERNAL, "|"), lIdx)
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Function pvText(ByVal lIdx As UcsReceiptTextsEnum) As String
    Const FUNC_NAME     As String = "pvText"
    
    On Error GoTo EH
    pvText = At(m_uConfig.LocalizedText(ucsFscLciReceiptTexts), lIdx)
    If LenB(pvText) = 0 Then
        pvText = At(Split(STR_RECEIPT_TEXTS, "|"), lIdx)
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
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

Private Function pvPrintLine(vText As Variant) As Boolean
    Const FUNC_NAME     As String = "pvPrintLine"
    Const CONF_NAME     As String = "PrintLine"
    Dim vElem           As Variant
    Dim eCmd            As UcsZekaCommandsEnum
    Dim sParam          As String
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "PrintText", eCmd, ucsZekCmdPrintText, sParam, "%1"
    If IsArray(vText) Then
        For Each vElem In vText
            If LenB(C_Str(vElem)) = 0 Then
                SendCommand eCmd, Printf(sParam, " ")
            Else
                SendCommand eCmd, Printf(sParam, C_Str(vElem))
            End If
            If LenB(m_sLastError) <> 0 Then
                Exit Function
            End If
        Next
    Else
        If LenB(C_Str(vText)) = 0 Then
            SendCommand eCmd, Printf(sParam, " ")
        Else
            SendCommand eCmd, Printf(sParam, C_Str(vText))
        End If
        If LenB(m_sLastError) <> 0 Then
            Exit Function
        End If
    End If
    '--- success
    pvPrintLine = True
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Function pvPrintReceipt(uCtx As UcsContext, uRow() As UcsRowData, ByVal lRowCount As Long) As Boolean
    Const FUNC_NAME     As String = "pvPrintReceipt"
    Const CONF_NAME     As String = "PrintReceipt"
    Dim sData           As String
    Dim dblTotal        As Double
    Dim dblDisc         As Double
    Dim vSplit          As Variant
    Dim lIdx            As Long
    Dim eCmd            As UcsZekaCommandsEnum
    Dim sParam          As String
    Dim lIndex          As Long
    Dim lPassIndex      As Long
    Dim sFormatPrice    As String
    Dim sFormatQty      As String
    Dim sFormatDisc     As String
    Dim sFormatAmount   As String
    
    On Error GoTo EH
    '--- ring sale
    Do While uCtx.Row < lRowCount
        With uRow(uCtx.Row)
        Select Case .RowType
        Case ucsRowInit
            pvGetCommandConfig CONF_NAME, "InfoParameters", eCmd, ucsZekCmdInfoParameters, sParam, vbNullString, _
                "Index", lIndex, 9
            vSplit = Split(SendCommand(eCmd, sParam), DEF_DELIM)
            If C_Bool(At(vSplit, lIndex)) Then
                '--- use single operator
                .InitOperatorCode = DEF_OPERNO
                .InitOperatorPassword = IDeviceProtocol_GetDefaultPassword(DEF_OPERNO)
            End If
            pvGetCommandConfig CONF_NAME, "InfoOperator", eCmd, ucsZekCmdInfoOperator, sParam, "%1", _
                "Index", lIndex, 1, _
                "PassIndex", lPassIndex, 2
            sParam = Printf(sParam, Pad(.InitOperatorCode, -m_uConfig.MaxOperCode))
            vSplit = Split(SendCommand(eCmd, sParam), DEF_DELIM)
            If LenB(m_sLastError) <> 0 Then
                GoTo QH
            End If
            Select Case .InitReceiptType
            Case ucsFscRetFiscal, ucsFscRetInvoiceFiscal
                '--- init operator name (if needed)
                If LenB(.InitOperatorName) = 0 Then
                    .InitOperatorName = Trim$(At(vSplit, lIndex))
                End If
                If Trim$(At(vSplit, lIndex)) <> Left$(Trim$(.InitOperatorName), m_uConfig.MaxOperName) Or LenB(Trim$(At(vSplit, lPassIndex))) = 0 Then
                    pvGetCommandConfig CONF_NAME, "InitOperator", eCmd, ucsZekCmdInitOperator, sParam, "%1;%2;%3"
                    sParam = Printf(sParam, Pad(.InitOperatorCode, -m_uConfig.MaxOperCode), _
                        Pad(.InitOperatorName, m_uConfig.MaxOperName), _
                        Pad(.InitOperatorPassword, m_uConfig.MaxOperPassword))
                    SendCommand eCmd, sParam
                    '--- can return "Command denied because of uncommited report"
                    If LenB(Trim$(At(vSplit, lPassIndex))) = 0 And LenB(m_sLastError) <> 0 Then
                        pvSetLastError m_sLastError & pvInternal(ucsErrPasswordNotSet), ucsFerPasswordNotSet
                        GoTo QH
                    End If
                End If
                '--- setup invoice no
                If .InitReceiptType = ucsFscRetInvoiceFiscal And LenB(.InitInvData(ucsInvNumber)) <> 0 Then
                    pvGetCommandConfig CONF_NAME, "InitInvoiceNo", eCmd, ucsZekCmdInitInvoiceNo, sParam, "%1;%1"
                    sParam = Printf(sParam, Pad(.InitInvData(ucsInvNumber), -m_uConfig.MaxInvoiceNo))
                    SendCommand eCmd, sParam
                    If LenB(m_sLastError) <> 0 Then
                        GoTo QH
                    End If
                End If
                '--- setup fiscal receipt
                pvGetCommandConfig CONF_NAME, "FiscalOpen" & .InitReceiptType, eCmd, ucsZekCmdFiscalOpen, _
                    sParam, "%1;%2;%3;0;%4" & IIf(.InitReceiptType = ucsFscRetInvoiceFiscal, ";%5;%6;%7;%8;%9", vbNullString)
                sParam = Printf(sParam, Pad(.InitOperatorCode, -m_uConfig.MaxOperCode), _
                    Pad(.InitOperatorPassword, m_uConfig.MaxOperPassword), _
                    IIf(m_uConfig.DetailedReceipt, 1, 0), _
                    IIf(m_uConfig.DelayPrint, 2, 0) Or IIf(.InitReceiptType = ucsFscRetInvoiceFiscal, 1, 0), _
                    Pad(.InitInvData(ucsInvCompany), m_uConfig.MaxCompany), _
                    Pad(.InitInvData(ucsInvMol), m_uConfig.MaxMol), _
                    Pad(.InitInvData(ucsInvVatNumber), m_uConfig.MaxVatNumber), _
                    Pad(.InitInvData(ucsInvTaxNumber), m_uConfig.MaxTaxNumber), _
                    Pad(.InitInvData(ucsInvCity), m_uConfig.MaxCity))
                SendCommand eCmd, sParam
                If LenB(m_sLastError) <> 0 Then
                    pvSetLastError m_sLastError, ucsFerInvalidPassword
                    GoTo QH
                End If
            Case ucsFscRetNonfiscal, ucsFscRetInvoiceNonfiscal
                If LenB(.InitOperatorName) = 0 Then
                    .InitOperatorName = Trim$(At(vSplit, lIndex))
                End If
                '--- setup non-fiscal receipt
                pvGetCommandConfig CONF_NAME, "NonFiscalOpen" & .InitReceiptType, eCmd, ucsZekCmdNonFiscalOpen, sParam, "%1;%2"
                sParam = Printf(sParam, Pad(.InitOperatorCode, -m_uConfig.MaxOperCode), Pad(.InitOperatorPassword, m_uConfig.MaxOperPassword))
                SendCommand eCmd, sParam
                If LenB(m_sLastError) <> 0 Then
                    GoTo QH
                End If
                If .InitReceiptType = ucsFscRetInvoiceNonfiscal Then
                    sData = .InitInvData(ucsInvNumber)
                    If Len(sData) <= 10 Then
                        sData = Printf(pvText(ucsTxtHeaderInvoiceNo), sData)
                    End If
                    pvPrintLine CenterText(sData, pvRowChars)
                End If
            End Select
        Case ucsRowPlu
            uCtx.PluCount = uCtx.PluCount + 1
            dblTotal = Round(.PluPrice * .PluQuantity, 2)
            dblDisc = Round(dblTotal * IIf(.DiscType = ucsFscDstPlu, .DiscValue, 0) / 100#, 2)
            uCtx.GrpTotal(.PluVatGroup) = Round(uCtx.GrpTotal(.PluVatGroup) + dblTotal + dblDisc, 2)
            '--- on invoices print empty line before first PLU
            If uCtx.PluCount = 1 Then
                Select Case .PrintRowType
                Case ucsFscRetInvoiceFiscal, ucsFscRetInvoiceNonfiscal
                    If Not pvPrintLine(vbNullString) Then
                        GoTo QH
                    End If
                End Select
            End If
            Select Case .PrintRowType
            Case ucsFscRetFiscal, ucsFscRetInvoiceFiscal
                pvGetCommandConfig CONF_NAME, "FiscalSell" & .InitReceiptType, eCmd, ucsZekCmdFiscalSell, sParam, "%1", _
                    "FormatPrice", sFormatPrice, "0.00", _
                    "FormatQty", sFormatQty, "0.000", _
                    "FormatDisc", sFormatDisc, "0.00"
                vSplit = WrapText(.PluName, pvRowChars)
                For lIdx = 0 To UBound(vSplit) - 1
                    If Not pvPrintLine(C_Str(vSplit(lIdx))) Then
                        GoTo QH
                    End If
                Next
                sData = Pad(At(vSplit, UBound(vSplit)), m_uConfig.MaxPluName)
                sData = sData & ";" & Chr$(191 + .PluVatGroup) & ";"
                If Abs(.PluQuantity) > DBL_EPSILON Then
                    sData = sData & SafeFormat(.PluPrice, sFormatPrice)
                    If .PluQuantity <> 1 Then
                        sData = sData & "*" & SafeFormat(.PluQuantity, sFormatQty)
                    End If
                    If .DiscType = ucsFscDstPlu And Abs(.PluPrice) > DBL_EPSILON Then
                        sData = sData & "," & SafeFormat(.DiscValue, sFormatDisc)
                    End If
                Else
                    sData = sData & SafeFormat(0, sFormatPrice)
                End If
                sParam = Printf(sParam, sData)
                SendCommand eCmd, sParam
            Case ucsFscRetNonfiscal, ucsFscRetInvoiceNonfiscal
                pvGetCommandConfig CONF_NAME, "NonFiscalSell" & .PrintRowType, 0, 0, vbNullString, vbNullString, _
                    "FormatPrice", sFormatPrice, "0.00", _
                    "FormatQty", sFormatQty, "0.000", _
                    "FormatDisc", sFormatDisc, "0.00", _
                    "FormatAmount", sFormatAmount, "0.00"
                If .PluPrice < -DBL_EPSILON Then
                    pvPrintLine CenterText(pvText(ucsTxtPluVoid), pvRowChars)
                End If
                If .PrintRowType = ucsFscRetNonfiscal Then
                    pvPrintLine WrapText(.PluName, pvRowChars)
                    If .PluQuantity <> 1 Then
                        pvPrintLine AlignText(vbNullString, SafeFormat(.PluQuantity, sFormatQty) & " x " & SafeFormat(Abs(.PluPrice), sFormatPrice), pvRowChars)
                    End If
                    pvPrintLine AlignText(Printf(pvText(ucsTxtPluTotal), Chr$(191 + .PluVatGroup)), SafeFormat(Abs(dblTotal), sFormatAmount), pvRowChars)
                Else
                    vSplit = WrapText(.PluName, pvRowChars)
                    For lIdx = 0 To UBound(vSplit)
                        pvPrintLine C_Str(vSplit(lIdx))
                    Next
                    pvPrintLine AlignText(pvText(ucsTxtPluQuantity), SafeFormat(.PluQuantity, sFormatQty), pvRowChars)
                    pvPrintLine AlignText(pvText(ucsTxtPluPrice), SafeFormat(Abs(.PluPrice), sFormatPrice), pvRowChars)
                    pvPrintLine AlignText(Printf(pvText(ucsTxtPluTotal), Chr$(191 + .PluVatGroup)), SafeFormat(Abs(dblTotal), sFormatAmount), pvRowChars)
                End If
                If .DiscType = ucsFscDstPlu Then
                    pvPrintLine AlignText(Printf(IIf(IIf(.PluPrice < -DBL_EPSILON, -1, 1) * dblDisc > DBL_EPSILON, pvText(ucsTxtSurcharge), pvText(ucsTxtDiscount)), "*" & Chr$(191 + .PluVatGroup) & "  " & SafeFormat(Abs(.DiscValue), sFormatDisc) & "%"), _
                        SafeFormat(IIf(.PluPrice < -DBL_EPSILON, -1, 1) * dblDisc, sFormatAmount), pvRowChars)
                End If
                If .PluPrice < -DBL_EPSILON Then
                    pvPrintLine CenterText(String(pvRowChars - 10, "-"), pvRowChars)
                End If
            End Select
        Case ucsRowDiscount
            Select Case .PrintRowType
            Case ucsFscRetFiscal, ucsFscRetInvoiceFiscal
                If .DiscType = ucsFscDstSubtotal Then
                    pvGetCommandConfig CONF_NAME, "FiscalSubtotal" & .DiscType, eCmd, ucsZekCmdFiscalSubtotal, sParam, "1;0,%1", _
                        "FormatDisc", sFormatDisc, "0.00"
                    sParam = Printf(sParam, SafeFormat(.DiscValue, sFormatDisc))
                    SendCommand eCmd, sParam
                End If
            Case ucsFscRetNonfiscal, ucsFscRetInvoiceNonfiscal
                If .DiscType = ucsFscDstSubtotal Then
                    pvGetCommandConfig CONF_NAME, "NonFiscalSubtotal" & .DiscType, 0, 0, vbNullString, vbNullString, _
                        "FormatDisc", sFormatDisc, "0.00", _
                        "FormatAmount", sFormatAmount, "0.00"
                    pvPrintLine AlignText(pvText(ucsTxtSubTotal), SafeFormat(SumArray(uCtx.GrpTotal), sFormatAmount), pvRowChars)
                    For lIdx = 1 To UBound(uCtx.GrpTotal)
                        If Abs(uCtx.GrpTotal(lIdx)) > DBL_EPSILON Then
                            dblTotal = Round(uCtx.GrpTotal(lIdx) * .DiscValue / 100#, 2)
                            pvPrintLine AlignText(Printf(IIf(dblTotal > DBL_EPSILON, pvText(ucsTxtSurcharge), pvText(ucsTxtDiscount)), "*" & Chr$(191 + lIdx) & "  " & SafeFormat(Abs(.DiscValue), sFormatDisc) & "%"), _
                                SafeFormat(dblTotal, sFormatAmount), pvRowChars)
                        End If
                    Next
                End If
            End Select
            If .DiscType = ucsFscDstSubtotal Then
                For lIdx = 1 To UBound(uCtx.GrpTotal)
                    If Abs(uCtx.GrpTotal(lIdx)) > DBL_EPSILON Then
                        dblTotal = Round(uCtx.GrpTotal(lIdx) * .DiscValue / 100#, 2)
                        uCtx.GrpTotal(lIdx) = Round(uCtx.GrpTotal(lIdx) + dblTotal, 2)
                    End If
                Next
            End If
        Case ucsRowLine
            If .LineWordWrap Then
                vSplit = WrapText(.LineText, pvRowChars)
            Else
                vSplit = Array(Left$(.LineText, pvRowChars))
            End If
            Select Case .PrintRowType
            Case ucsFscRetFiscal, ucsFscRetInvoiceFiscal
                pvPrintLine vSplit
            Case ucsFscRetNonfiscal, ucsFscRetInvoiceNonfiscal
                If uCtx.PmtPrinted And Not uCtx.ChangePrinted Then
                    dblTotal = Round(SumArray(uCtx.GrpTotal) - uCtx.Paid, 2)
                    If dblTotal > DBL_EPSILON Then
                        pvPrintLine AlignText(GetPaymentName(1), SafeFormat(dblTotal, "0.00"), pvRowChars)
                    ElseIf dblTotal < -DBL_EPSILON Then
                        pvPrintLine AlignText(pvText(ucsTxtFooterChange), SafeFormat(-dblTotal, "0.00"), pvRowChars)
                    End If
                    uCtx.ChangePrinted = True
                End If
                If uCtx.PluCount = 0 Then
                    For lIdx = 0 To UBound(vSplit)
                        vSplit(lIdx) = CenterText(Trim$(vSplit(lIdx)), pvRowChars)
                    Next
                End If
                pvPrintLine vSplit
            End Select
        Case ucsRowPayment
            Select Case .PrintRowType
            Case ucsFscRetFiscal, ucsFscRetInvoiceFiscal
                pvGetCommandConfig CONF_NAME, "InfoTransaction", eCmd, ucsZekCmdInfoTransaction, sParam, vbNullString, _
                    "Index", lIndex, 0
                vSplit = Split(SendCommand(eCmd, sParam), DEF_DELIM)
                If LenB(m_sLastError) <> 0 Then
                    GoTo QH
                End If
                If .PmtType <> 0 Then
                    pvGetCommandConfig CONF_NAME, "FiscalPayment" & .PmtType, eCmd, ucsZekCmdFiscalPayment, sParam, (.PmtType - 1) & ";0;%1", _
                        "FormatAmount", sFormatAmount, "0.00"
                    sParam = Printf(sParam, SafeFormat(.PmtAmount, sFormatAmount))
                ElseIf C_Lng(At(vSplit, lIndex)) <> 0 Then
                    pvGetCommandConfig CONF_NAME, "FiscalPayAndClose", eCmd, ucsZekCmdFiscalPayAndClose, sParam, vbNullString
                End If
                SendCommand eCmd, sParam
                If LenB(m_sLastError) <> 0 Then
                    GoTo QH
                End If
                uCtx.PmtPrinted = True
            Case ucsFscRetNonfiscal, ucsFscRetInvoiceNonfiscal
                pvGetCommandConfig CONF_NAME, "NonFiscalFooter", 0, 0, vbNullString, vbNullString, _
                    "FormatAmount", sFormatAmount, "0.00"
                If Not uCtx.PmtPrinted Then
                    If uCtx.PluCount > 0 Or .PmtType <> 0 Then
                        pvPrintLine vbNullString
                        pvPrintLine AlignText(pvText(ucsTxtFooterTotal), SafeFormat(SumArray(uCtx.GrpTotal), sFormatAmount), pvRowChars)
                        If .PrintRowType = ucsFscRetInvoiceNonfiscal Then
                            '--- print invoice totals by VAT group
                            pvGetCommandConfig CONF_NAME, "InfoTaxRates", eCmd, ucsZekCmdInfoTaxRates, sParam, vbNullString
                            vSplit = Split(Replace(SendCommand(eCmd, sParam), "%", vbNullString), DEF_DELIM)
                            If LenB(m_sLastError) <> 0 Then
                                GoTo QH
                            End If
                            For lIdx = 1 To UBound(uCtx.GrpTotal)
                                If Abs(uCtx.GrpTotal(lIdx)) > DBL_EPSILON Then
                                    '--- first, round VAT
                                    dblTotal = Round(uCtx.GrpTotal(lIdx) - uCtx.GrpTotal(lIdx) / (1 + C_Dbl(At(vSplit, lIdx - 1)) / 100#), 2)
                                    '--- then, calc TaxBase = Total - VAT
                                    dblTotal = Round(uCtx.GrpTotal(lIdx) - dblTotal, 2)
                                    pvPrintLine AlignText(Printf(pvText(ucsTxtGroupTotal), Chr$(191 + lIdx)), SafeFormat(uCtx.GrpTotal(lIdx), sFormatAmount), pvRowChars)
                                    pvPrintLine AlignText(Printf(pvText(ucsTxtGroupVat), Chr$(191 + lIdx), SafeFormat(C_Dbl(At(vSplit, lIdx - 1)), sFormatAmount) & "%"), _
                                        SafeFormat(Round(uCtx.GrpTotal(lIdx) - dblTotal, 2), sFormatAmount), pvRowChars)
                                    pvPrintLine AlignText(pvText(ucsTxtGroupNetto), SafeFormat(dblTotal, sFormatAmount), pvRowChars)
                                End If
                            Next
                        End If
                        uCtx.PmtPrinted = True
                    End If
                End If
                If .PmtType <> 0 Then
                    If LenB(.PmtName) = 0 Then
                        .PmtName = GetPaymentName(.PmtType)
                    End If
                    pvPrintLine AlignText(.PmtName, SafeFormat(.PmtAmount, sFormatAmount), pvRowChars)
                    uCtx.Paid = Round(uCtx.Paid + .PmtAmount, 2)
                Else
                    If uCtx.PmtPrinted And Not uCtx.ChangePrinted Then
                        dblTotal = Round(SumArray(uCtx.GrpTotal) - uCtx.Paid, 2)
                        If dblTotal > DBL_EPSILON Then
                            pvPrintLine AlignText(GetPaymentName(1), SafeFormat(dblTotal, sFormatAmount), pvRowChars)
                        ElseIf dblTotal < -DBL_EPSILON Then
                            pvPrintLine AlignText(pvText(ucsTxtFooterChange), SafeFormat(-dblTotal, sFormatAmount), pvRowChars)
                        End If
                        uCtx.ChangePrinted = True
                    End If
                    If .PrintRowType = ucsFscRetInvoiceNonfiscal Then
                        With uRow(0)
                            pvPrintLine AlignText(pvText(ucsTxtFooterSeller), Chr$(1) & ".", pvRowChars)
                            pvPrintLine AlignText(vbNullString, "/" & .InitOperatorName & "/", pvRowChars)
                            pvPrintLine CenterText(String(pvRowChars - 10, "-"), pvRowChars)
                            pvPrintLine WrapText(pvText(ucsTxtFooterReceiver) & .InitInvData(ucsInvCompany), pvRowChars)
                            If LenB(.InitInvData(ucsInvMol)) <> 0 Then
                                pvPrintLine WrapText(pvText(ucsTxtFooterBuyer) & .InitInvData(ucsInvMol), pvRowChars)
                            End If
                            If LenB(.InitInvData(ucsInvCity)) <> 0 Then
                                pvPrintLine .InitInvData(ucsInvCity)
                            End If
                            If LenB(.InitInvData(ucsInvAddress)) <> 0 Then
                                pvPrintLine .InitInvData(ucsInvAddress)
                            End If
                            pvPrintLine AlignText(pvText(ucsTxtFooterTaxNo), .InitInvData(ucsInvTaxNumber), pvRowChars)
                            If LenB(.InitInvData(ucsInvVatNumber)) <> 0 Then
                                pvPrintLine AlignText(pvText(ucsTxtFooterVatNo), .InitInvData(ucsInvVatNumber), pvRowChars)
                            End If
                            pvPrintLine CenterText(String(pvRowChars - 10, "-"), pvRowChars)
                        End With
                    End If
                    sData = Trim$(GetFooterText(1))
                    If LenB(sData) <> 0 Then
                        pvPrintLine CenterText(sData, pvRowChars)
                    End If
                    If uCtx.PluCount > 0 Then
                        sData = IIf(uCtx.PluCount = 1, pvText(ucsTxtFooterArticleSingle), Printf(pvText(ucsTxtFooterArticles), uCtx.PluCount))
                        pvPrintLine AlignText(vbNullString, sData, pvRowChars)
                    End If
                    pvGetCommandConfig CONF_NAME, "NonFiscalClose" & .PrintRowType, eCmd, ucsZekCmdNonFiscalClose, sParam, vbNullString
                    SendCommand eCmd, sParam
                    If LenB(m_sLastError) <> 0 Then
                        GoTo QH
                    End If
                End If
            End Select
            If .PmtType = 0 And m_uConfig.DelayPrint Then
                If Not WaitDevice(lRowCount * m_lTimeout) Then
                    GoTo QH
                End If
            End If
        End Select
        End With
        If LenB(m_sLastError) <> 0 Then
            GoTo QH
        End If
        uCtx.Row = uCtx.Row + 1
    Loop
    '--- success
    pvPrintReceipt = True
QH:
    If LenB(m_sLastError) <> 0 And uCtx.PluCount > 0 And Not uCtx.PmtPrinted Then
        pvSetLastError Printf(pvInternal(ucsErrErrorInLine), uCtx.PluCount, m_sLastError), m_eLastErrNumber
    End If
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
    
    On Error GoTo EH
    '--- convert out-of-range discounts to PLU rows
    '--- note: m_lRowCount may change in loop on AddPLU
    Do While lRow < m_lRowCount
        '--- note: 'With' locks m_uRow array and fails if auto-grow needed in AddPLU
'        With m_uRow(lRow)
            If m_uRow(lRow).RowType = ucsRowPlu Then
                If (m_uRow(lRow).DiscValue < m_uConfig.MinDiscount Or m_uRow(lRow).DiscValue > m_uConfig.MaxDiscount) Then
                    dblTotal = Round(m_uRow(lRow).PluQuantity * m_uRow(lRow).PluPrice, 2)
                    dblDiscount = Limit(m_uRow(lRow).DiscValue, m_uConfig.MinDiscount, m_uConfig.MaxDiscount)
                    If Round(dblTotal * m_uRow(lRow).DiscValue / 100#, 2) = Round(dblTotal * dblDiscount / 100#, 2) Then
                        m_uRow(lRow).DiscValue = dblDiscount
                    Else
                        dblDiscount = m_uRow(lRow).DiscValue
                        m_uRow(lRow).DiscType = 0
                        m_uRow(lRow).DiscValue = 0
                        AddPLU Printf(IIf(dblTotal * dblDiscount > DBL_EPSILON, pvText(ucsTxtSurcharge), pvText(ucsTxtDiscount)), "*" & Chr$(191 + m_uRow(lRow).PluVatGroup) & "  " & SafeFormat(Abs(dblDiscount), "0.00") & "%"), _
                            Round(dblTotal * dblDiscount / 100#, 2), 1, m_uRow(lRow).PluVatGroup, lRow + 1
                    End If
                ElseIf m_uRow(lRow).DiscType = ucsFscDstPlu And m_uRow(lRow).PluPrice < -DBL_EPSILON Then
                    '--- convert PLU discount on void rows
                    dblTotal = Round(m_uRow(lRow).PluQuantity * m_uRow(lRow).PluPrice, 2)
                    dblDiscount = m_uRow(lRow).DiscValue
                    m_uRow(lRow).DiscType = 0
                    m_uRow(lRow).DiscValue = 0
                    AddPLU Printf(IIf(dblTotal * dblDiscount > DBL_EPSILON, pvText(ucsTxtSurcharge), pvText(ucsTxtDiscount)), "*" & Chr$(191 + m_uRow(lRow).PluVatGroup) & "  " & SafeFormat(Abs(dblDiscount), "0.00") & "%"), _
                            Round(dblTotal * dblDiscount / 100#, 2), 1, m_uRow(lRow).PluVatGroup, lRow + 1
                End If
            ElseIf m_uRow(lRow).RowType = ucsRowDiscount Then
                If (m_uRow(lRow).DiscValue < m_uConfig.MinDiscount Or m_uRow(lRow).DiscValue > m_uConfig.MaxDiscount) And m_uRow(lRow).DiscType = ucsFscDstSubtotal Then
                    pvGetSubtotals lRow, uSum
                    dblDiscount = Limit(m_uRow(lRow).DiscValue, m_uConfig.MinDiscount, m_uConfig.MaxDiscount)
                    lCount = 0
                    For lIdx = 1 To UBound(uSum.GrpTotal)
                        If Round(uSum.GrpTotal(lIdx) * m_uRow(lRow).DiscValue / 100#, 2) <> Round(uSum.GrpTotal(lIdx) * dblDiscount / 100#, 2) Then
                            lCount = lCount + 1
                        End If
                    Next
                    If lCount = 0 Then
                        m_uRow(lRow).DiscValue = dblDiscount
                    Else
                        dblDiscount = m_uRow(lRow).DiscValue
                        m_uRow(lRow).DiscValue = 0
                        For lIdx = UBound(uSum.GrpTotal) To 1 Step -1
                            If Abs(uSum.GrpTotal(lIdx)) > DBL_EPSILON Then
                                AddPLU Printf(IIf(uSum.GrpTotal(lIdx) * dblDiscount > DBL_EPSILON, pvText(ucsTxtSurcharge), pvText(ucsTxtDiscount)), "*" & Chr$(191 + lIdx) & "  " & SafeFormat(Abs(dblDiscount), "0.00") & "%"), _
                                    Round(uSum.GrpTotal(lIdx) * dblDiscount / 100#, 2), 1, lIdx, lRow + 1
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
                lCount = lCount + 1
                uCtx.GrpTotal(.PluVatGroup) = 1
            End If
        End With
    Next
    If lCount > m_uConfig.MaxReceiptRows Then
        '--- count different VAT groups in PLUs
        For lRow = 1 To UBound(uCtx.GrpTotal)
            If Abs(uCtx.GrpTotal(lRow)) > DBL_EPSILON Then
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
                    If lCount > m_uConfig.MaxReceiptRows - lTotal Then
                        .PrintRowType = IIf(.PrintRowType = ucsFscRetFiscal, ucsFscRetNonfiscal, ucsFscRetInvoiceNonfiscal)
                        dblTotal = Round(.PluQuantity * .PluPrice, 2)
                        If .DiscType = ucsFscDstPlu Then
                            dblTotal = Round(dblTotal + Round(dblTotal * .DiscValue / 100#, 2), 2)
                        End If
                        If .PluVatGroup > 0 Then
                            uCtx.GrpTotal(.PluVatGroup) = Round(uCtx.GrpTotal(.PluVatGroup) + dblTotal, 2)
                        End If
                    End If
                ElseIf .RowType = ucsRowDiscount And .DiscType = ucsFscDstSubtotal Then
                    If lCount > m_uConfig.MaxReceiptRows - lTotal Then
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
                Exit For
            End If
        Next
        '--- append fiscal rows for GrpTotal by VAT groups
        For lIdx = 1 To UBound(uCtx.GrpTotal)
            If Abs(uCtx.GrpTotal(lIdx)) > DBL_EPSILON Then
                AddPLU Printf(pvText(ucsTxtPluSales), Chr$(191 + lIdx)), uCtx.GrpTotal(lIdx), 1, lIdx, lRow
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
            If .DiscType = ucsFscDstPlu Then
                dblTotal = Round(dblTotal + Round(dblTotal * .DiscValue / 100#, 2), 2)
            End If
            If .PluVatGroup > 0 Then
                uCtx.GrpTotal(.PluVatGroup) = Round(uCtx.GrpTotal(.PluVatGroup) + dblTotal, 2)
            End If
        ElseIf .RowType = ucsRowDiscount Then
            If .DiscType = ucsFscDstSubtotal Then
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

Private Function pvZfplibValue(lValue As Long, sRegValue As String, ByVal lDefValue As Long) As Long
    If lValue = 0 Then
        lValue = C_Lng(RegReadString(HKEY_LOCAL_MACHINE, "Software\Tremol\ZFPLib", sRegValue))
        If lValue = 0 Then
            lValue = lDefValue
        End If
    End If
    pvZfplibValue = lValue
End Function

Private Function pvGetDeviceInfo(ByVal hComm As Long, ByVal lTimeout As Long, vResult As Variant, sProtocol As String, sModel As String, Optional Error As String) As Boolean
    Const FUNC_NAME     As String = "pvGetDeviceInfo"
    Const CONF_NAME     As String = "GetDeviceInfo"
    Dim vSplit          As Variant
    Dim eCmd            As UcsZekaCommandsEnum
    Dim sParam          As String
    Dim lSerialIndex    As Long
    Dim lRegDateIndex   As Long
    Dim lModelIndex     As Long
    Dim sSuffix         As String
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoDiagnostics", eCmd, ucsZekCmdInfoDiagnostics, sParam, vbNullString, _
        "SerialIndex", lSerialIndex, 0
    If Not IsArray(vResult) Then
        vResult = Split(pvSendCommand(hComm, lTimeout, eCmd, sParam, Error), DEF_DELIM)
    End If
    If LenB(At(vResult, lSerialIndex)) <> 0 Then
        sProtocol = STR_PROTOCOL_ZEKA_FP
        pvGetCommandConfig CONF_NAME, "InfoVersion", eCmd, ucsZekCmdInfoVersion, sParam, vbNullString, _
            "RegDateIndex", lRegDateIndex, 3, _
            "ModelIndex", lModelIndex, 4, _
            "Suffix", sSuffix, "Ê.Ñ."
        vSplit = Split(pvSendCommand(hComm, lTimeout, eCmd, sParam), DEF_DELIM)
        sModel = Trim$(At(vSplit, lRegDateIndex) & " " & At(vSplit, lModelIndex))
        If InStr(sModel, sSuffix) > 0 Then
            sModel = Trim$(Left$(sModel, InStr(sModel, sSuffix) - 1))
        End If
        '--- success
        pvGetDeviceInfo = True
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Sub pvGetCommandConfig( _
            sFunc As String, _
            sKey As String, _
            eCmd As UcsZekaCommandsEnum, _
            ByVal eCmdDefault As UcsZekaCommandsEnum, _
            sParam As String, _
            sParamDefault As String, _
            ParamArray A() As Variant)
    Dim lIdx            As Long
    
    eCmd = GetConfigForCommand(m_uConfig.ConfigCommands, m_uConfig.LocalizedCommands, sFunc, sKey, eCmdDefault)
    sParam = GetConfigForCommand(m_uConfig.ConfigCommands, m_uConfig.LocalizedCommands, sFunc, sKey & "Param", sParamDefault)
    For lIdx = 0 To UBound(A) Step 3
        A(lIdx + 1) = GetConfigForCommand(m_uConfig.ConfigCommands, m_uConfig.LocalizedCommands, sFunc, sKey & A(lIdx + 0), A(lIdx + 2))
    Next
End Sub

Private Sub pvSetLastError(sError As String, Optional ByVal ErrNum As UcsFiscalErrorsEnum = -1)
    If ErrNum < 0 Then
        m_eLastErrNumber = IIf(LenB(sError) = 0, ucsFerNone, ucsFerGeneralError)
    Else
        m_eLastErrNumber = ErrNum
    End If
    m_sLastError = sError
End Sub

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Initialize()
    m_hComm = INVALID_HANDLE_VALUE
    m_lTimeout = DEF_TIMEOUT
    m_lSequence = ZFP_BASEDATA
    LocalizedText(ucsFscLciInternalErrors) = STR_INTERNAL
    LocalizedText(ucsFscLciPrinterStatuses) = STR_STATUSES
    LocalizedText(ucsFscLciPrinterErrors) = STR_ERRORS
    LocalizedText(ucsFscLciReceiptTexts) = STR_RECEIPT_TEXTS
    Set m_uConfig.ConfigCommands = GetConfigCollection(STR_PROTOCOL_ZEKA_FP, "Commands")
End Sub

Private Sub Class_Terminate()
    PortWrapper.CloseComPort m_hComm
    m_hComm = INVALID_HANDLE_VALUE
End Sub

'=========================================================================
' IDeviceProtocol interface
'=========================================================================

Private Function IDeviceProtocol_AddDiscount(ByVal DiscType As UcsFiscalDiscountTypeEnum, ByVal Value As Double) As Boolean
    IDeviceProtocol_AddDiscount = AddDiscount(DiscType, Value)
End Function

Private Function IDeviceProtocol_AddLine(Line As String, Optional Command As String, Optional ByVal WordWrap As Boolean = True) As Boolean
    IDeviceProtocol_AddLine = AddLine(Line, WordWrap)
End Function

Private Function IDeviceProtocol_AddPayment(ByVal Number As UcsFiscalPaymentTypeEnum, Name As String, ByVal Amount As Double, Optional ByVal Rate As Double) As Boolean
    IDeviceProtocol_AddPayment = AddPayment(Number, Name, Amount)
End Function

Private Function IDeviceProtocol_AddPLU(Name As String, ByVal Price As Double, ByVal Quantity As Double, ByVal VatGroup As Long) As Boolean
    IDeviceProtocol_AddPLU = AddPLU(Name, Price, Quantity, VatGroup)
End Function

Private Function IDeviceProtocol_AutodetectDevices(Ports As Variant) As Variant
    IDeviceProtocol_AutodetectDevices = AutodetectDevices(Ports)
End Function

Private Function IDeviceProtocol_CancelReceipt() As Boolean
    IDeviceProtocol_CancelReceipt = CancelReceipt()
End Function

Private Function IDeviceProtocol_CashDebitCredit(OperatorCode As String, OperatorPassword As String, ByVal Value As Double) As Variant
    IDeviceProtocol_CashDebitCredit = CashDebitCredit(OperatorCode, OperatorPassword, ucsFscPmtCash, Value)
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

Private Function IDeviceProtocol_EndReceipt(Optional ResumeTicket As String) As Boolean
    IDeviceProtocol_EndReceipt = EndReceipt(ResumeTicket)
End Function

Private Function IDeviceProtocol_GetCharsPerLine() As Long
    IDeviceProtocol_GetCharsPerLine = pvRowChars
End Function

Private Function IDeviceProtocol_GetClock() As Date
    IDeviceProtocol_GetClock = GetClock()
End Function

Private Function IDeviceProtocol_GetCommandLog() As String
    IDeviceProtocol_GetCommandLog = m_sCommandLog
End Function

Private Function IDeviceProtocol_GetDefaultPassword(OperatorCode As String) As String
    If IsEcr Then
        If C_Lng(OperatorCode) > 1 Then
            IDeviceProtocol_GetDefaultPassword = OperatorCode
        Else
            IDeviceProtocol_GetDefaultPassword = DEF_ECR_PASSWORD
        End If
    Else
        IDeviceProtocol_GetDefaultPassword = DEF_FP_PASSWORD
    End If
End Function

Private Function IDeviceProtocol_GetDeviceModel() As String
    IDeviceProtocol_GetDeviceModel = GetDeviceModel()
End Function

Private Function IDeviceProtocol_GetDeviceProtocol() As String
    IDeviceProtocol_GetDeviceProtocol = GetDeviceProtocol()
End Function

Private Function IDeviceProtocol_GetDeviceStatus(CurrentStatus As String) As Boolean
    IDeviceProtocol_GetDeviceStatus = GetDeviceStatus(CurrentStatus)
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

Private Function IDeviceProtocol_GetLastReceiptNumber() As String
    IDeviceProtocol_GetLastReceiptNumber = GetLastReceiptNumber()
End Function

Private Function IDeviceProtocol_GetLastError(Optional ErrCategory As UcsFiscalErrorsEnum) As String
    ErrCategory = LastErrNumber
    IDeviceProtocol_GetLastError = LastError
End Function

Private Function IDeviceProtocol_GetLastInvoiceNumber() As String
    IDeviceProtocol_GetLastInvoiceNumber = GetLastInvoiceNumber()
End Function

Private Function IDeviceProtocol_GetPaymentName(ByVal Index As Long) As String
    IDeviceProtocol_GetPaymentName = GetPaymentName(Index)
End Function

Private Function IDeviceProtocol_GetResumeTicket() As String
    IDeviceProtocol_GetResumeTicket = GetResumeTicket()
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

Private Function IDeviceProtocol_GetTotalsByPayments() As Variant
    IDeviceProtocol_GetTotalsByPayments = GetTotalsByPayments()
End Function

Private Function IDeviceProtocol_GetTotalsByVatGroups() As Variant
    IDeviceProtocol_GetTotalsByVatGroups = GetTotalsByVatGroups()
End Function

Private Function IDeviceProtocol_GetTotalsStats() As Variant
    IDeviceProtocol_GetTotalsStats = GetTotalsStats()
End Function

Private Function IDeviceProtocol_Init(Device As String, Optional ByVal Timeout As Long, Optional ByVal CashDeskNo As Long) As Boolean
    IDeviceProtocol_Init = Init(Device, Timeout, CashDeskNo)
End Function

Private Property Get IDeviceProtocol_IsConnected() As Boolean
    IDeviceProtocol_IsConnected = IsConnected
End Property

Private Function IDeviceProtocol_OpenDrawer() As Boolean
    IDeviceProtocol_OpenDrawer = OpenDrawer()
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

Private Function IDeviceProtocol_SetClock(ByVal NewDate As Date) As Boolean
    IDeviceProtocol_SetClock = SetClock(NewDate)
End Function

Private Sub IDeviceProtocol_SetLocalizedText(ByVal Index As UcsFiscalLocalizedIndexesEnum, Text As String)
    LocalizedText(Index) = Text
End Sub

Private Function IDeviceProtocol_StartReceipt(ByVal ReceiptType As UcsFiscalReceiptTypeEnum, OperatorCode As String, OperatorName As String, OperatorPassword As String, Optional InvNumber As String, Optional InvTaxNumber As String, Optional InvVatNumber As String, Optional InvCompany As String, Optional InvCity As String, Optional InvAddress As String, Optional InvMol As String, Optional InvReceiver As String) As Boolean
    IDeviceProtocol_StartReceipt = StartReceipt(ReceiptType, OperatorCode, OperatorName, OperatorPassword, InvNumber, InvTaxNumber, InvVatNumber, InvCompany, InvCity, InvAddress, InvMol, InvReceiver)
End Function

