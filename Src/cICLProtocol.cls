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
' $Header: /UcsFiscalPrinter/Src/cICLProtocol.cls 46    31.10.12 17:30 Wqw $
'
'   Unicontsoft Fiscal Printers Project
'   Copyright (c) 2008-2012 Unicontsoft
'
'   Protocol for Datecs & Daisy fiscal printers (ICL protocol)
'
' $Log: /UcsFiscalPrinter/Src/cICLProtocol.cls $
' 
' 46    31.10.12 17:30 Wqw
' REF: x and z reports clear sums by operators
'
' 45    19.10.12 0:36 Wqw
' REF: config defaults
'
' 44    10.10.12 14:58 Wqw
' REF: GetTotalsByVatGroups/CashDebitCredit retval array size fixed
'
' 43    9.10.12 15:20 Wqw
' REF: parms of IDeviceProtocol_CashDebitCredit
'
' 42    5.10.12 14:20 Wqw
' ADD: Function CashDebitCredit
'
' 41    29.08.12 15:05 Wqw
' REF: impl config values instead of constants
'
' 40    10.08.12 18:12 Wqw
' REF: KLEN device check
'
' 39    10.08.12 18:06 Wqw
' REF: workaround for KLEN devices that don't support custom PLUs in
' voids and zero price in fiscal rows
'
' 38    6.08.12 18:39 Wqw
' REF: impl AutodetectDevices
'
' 37    5.05.12 18:44 Wqw
' REF: remove caret in hdr/ftr, simpler transcode
'
' 36    23.03.12 15:28 Wqw
' ADD: GetTotalsByVatGroups. REF: cancel outstanding receipt before
' printing reports, no exceptions on info functions
'
' 35    8.12.11 15:50 Wqw
' ADD: Property IsEcr, Function GetDeviceModel. REF: better support for
' ECRs
'
' 34    6.12.11 17:48 Wqw
' REF: always try to transcode from printer
'
' 33    13.10.11 10:54 Wqw
' REF: Init params
'
' 32    9.08.11 23:22 Wqw
' REF: discount on negative price handled in pvConvertExtraRows
'
' 31    18.07.11 13:43 Wqw
' REF: operator name size syntax error
'
' 30    15.07.11 16:48 Wqw
' REF: final payment ot zero sum receipt
'
' 29    13.07.11 19:45 Wqw
' REF: encoding auto-sensing error tollerance in send message
'
' 28    13.07.11 18:14 Wqw
' REF: final payment if anything outstanding, while loop in convert extra
' rows
'
' 27    4.07.11 15:48 Wqw
' REF: err handling
'
' 26    17.06.11 13:23 Wqw
' REF: in SendCommand sets last error on printer/memory error
'
' 25    17.06.11 11:34 Wqw
' REF: auto-sense charset, merge pvRead/WriteData
'
' 24    1.04.11 14:41 Wqw
' REF: signal timeout in pvWriteData
'
' 23    8.03.11 13:07 Wqw
' REF: impl out-of-bounds discounts
'
' 22    22.02.11 17:38 Wqw
' REF: calls FlushFileBuffers after CancelIo
'
' 21    22.02.11 13:52 Wqw
' ADD: GetDeviceProtocol
'
' 20    22.02.11 10:05 Wqw
' ADD: Function GetDeviceModel
'
' 19    21.02.11 13:44 Wqw
' REF: impl IDeviceProtocol
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
Implements IDeviceProtocol

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
    ucsFpcFiscalSubtotal = 51
    ucsFpcFiscalPayment = 53
    ucsFpcFiscalText = 54
    ucsFpcFiscalClose = 56
    ucsFpcFiscalCgInfo = 57
    ucsFpcFiscalCancel = 60
    ucsFpcAdminCashDebitCredit = 70
    ucsFpcPrintDuplicate = 109
    '--- info
    ucsFpcInfoDateTime = 62
    ucsFpcInfoLastFiscalBlock = 64
    ucsFpcInfoTotalsByVatGroups = 65
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
    '--- ECR
    ucsFpcEcrReadRow = &HFF84&
    ucsFpcEcrWriteRow = &HFF86&
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
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
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
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
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

Private Const STR_INTERNAL              As String = "No device info set|CreateFile failed: |SetCommTimeouts failed: |BuildCommDCB failed: |SetCommState failed: |WriteFile failure: |Timeout waiting for response|ReadFile failed: |Invalid message format or checksum (NAK)|Already connected|Not connected|Invalid response length|Response missing postamble|Response missing preamble|WaitCommEvent failed: |No receipt started| or invalid operator password|Error in line %1: %2|Cannot cancel previous receipt"
Private Const STR_STATUSES              As String = "Syntax error|Invalid command|Clock not set|3|Feed error|5|6|7|Overflow|Command not allowed|Memory reset|3|Memory corrupt|5|6|7|Out of paper|1|2|Fiscal printing|4|Nonfiscal printing|6|7"
Private Const STR_ERRORS_ONLY           As String = "Syntax error|Invalid command|2|3|Feed error|5|6|7|1|Command not allowed|Memory reset|3|Memory corrupt|5|6|7|Out of paper|1|2|3|4|5|6|7"
Private Const STR_DIP_SWITCHES          As String = "Center header and footer|Enable header|Sw1.3|Sw1.4|Zero memory|Transparent display|No display|7"
Private Const STR_MEMORY                As String = "Write error|1|No memory module|Memory low|Out of memory|Fiscal memory error|6|7|Read only|Memory formatted|2|Fiscal mode|Tax rates set|Tax number set|6|7"
Private Const STR_RECEIPT_TEXTS         As String = "CASH|CARD|CHEQUE|CREDIT|SURCHARGE %1|DISCOUNT %1|TOTAL|SUBTOTAL|INVOICE No %1|Sales %1|UNIT PRICE|QUANTITY|VALUE|SUM GROUP %1|VAT %1=%2|NETTO SUM|TOTAL|%1 ITEMS|1 ITEM|CHANGE|SELLER: |RECEIVER: |BUYER: |TAX No: |VAT No: "
Private Const DEF_TIMEOUT               As Long = 3000
Private Const DEF_PASSWORD              As String = "0000"
Private Const DEF_COMMENT_LEN           As Long = 30
Private Const DEF_MAX_RECEIPT_ROWS      As Long = 120
Private Const DEF_MIN_DISCOUNT          As Double = -99
Private Const DEF_MAX_DISCOUNT          As Double = 99
Private Const DEF_MAX_PAYMENT_LEN       As Long = 24
Private Const MIK_OFFSET                As Long = &H80
'Private Const WIN_OFFSET                As Long = &HC0
Private Const ucsFscDstPluAbs           As Long = ucsFscDstPlu + 100
Private Const ucsFscDstSubtotalAbs      As Long = ucsFscDstSubtotal + 100

Private m_sDevice                   As String
Private m_hComm                     As Long
Private m_lTimeout                  As Long
Private m_lCashDeskNo               As Long
Private m_lSequence                 As Long
Private m_eStatus                   As UcsIclStatusBitsEnum
Private m_eDip                      As UcsIclDipSwitchesEnum
Private m_eMemory                   As UcsIclFiscalMemoryEnum
Private m_sLastError                As String
Private m_bExceptions               As Boolean
Private m_uRow()                    As UcsRowData
Private m_lRowCount                 As Long
Private m_bFiscalPaymentExecuted    As Boolean
Private m_sCommandLog               As String
Private m_vInfoDiagnostics          As Variant
Private m_uConfig                   As UcsConfigValues

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
    ucsErrNoReceiptStarted
    ucsErrInvalidOperatorPassword
    ucsErrErrorInLine
    ucsErrCannotCancelReceipt
End Enum

Private Enum UcsReceiptTextsEnum
    ucsTxtPaymentCash
    ucsTxtPaymentCard
    ucsTxtPaymentCheque
    ucsTxtPaymentCredit
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
End Type

Private Enum UcsCharset
    ucsCrsAuto
    ucsCrsWindows
    ucsCrsMik
End Enum

Private Enum UcsTriState
    ucsTri_Undefined = -2       ' Undefined value. Usually value is inherited from parent definition.
    ucsTriFalse = 0             ' (Default) Represents False.
    ucsTriTrue = -1             ' Represents True.
End Enum

Private Type UcsAutodetect
    hComm               As Long
    WriteOver           As OVERLAPPED
    ReadOver            As OVERLAPPED
    Buffer              As String
End Type

Private Type UcsConfigValues
    RowChars            As Long
    AbsoluteDiscount    As Boolean
    NegativePrices      As Boolean
    MaxReceiptRows      As Long
    MinDiscount         As Double
    MaxDiscount         As Double
    MaxPaymentLen       As Long
    Charset             As UcsCharset
    IsEcr               As UcsTriState
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

Property Get Status(ByVal eValue As UcsIclStatusBitsEnum) As Boolean
    Status = (m_eStatus And eValue) <> 0
End Property

Property Get StatusText() As String
    StatusText = frGetFlagsText(m_eStatus, m_uConfig.LocalizedText(ucsFscLciPrinterStatuses))
End Property

Property Get ErrorText() As String
    ErrorText = frGetFlagsText(m_eStatus, m_uConfig.LocalizedText(ucsFscLciPrinterErrors))
End Property

Property Get Dip(ByVal eValue As UcsIclDipSwitchesEnum) As Boolean
    Dip = (m_eDip And eValue) <> 0
End Property

Property Get DipText() As String
    DipText = frGetFlagsText(m_eDip, m_uConfig.LocalizedText(ucsFscLciDipSwitchesTexts))
End Property

Property Get Memory(ByVal eValue As UcsIclFiscalMemoryEnum) As Boolean
    Memory = (m_eMemory And eValue) <> 0
End Property

Property Get MemoryText() As String
    MemoryText = frGetFlagsText(m_eMemory, m_uConfig.LocalizedText(ucsFscLciMemoryTexts))
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
    LocalizedText = Join(m_uConfig.LocalizedText(eIdx), "|")
End Property

Property Let LocalizedText(ByVal eIdx As UcsFiscalLocalizedIndexesEnum, sValue As String)
    m_uConfig.LocalizedText(eIdx) = Split(sValue, "|")
End Property

Property Get IsEcr() As Boolean
    If m_uConfig.IsEcr = ucsTri_Undefined Then
        m_uConfig.IsEcr = IIf(LenB(pvSendCommand(m_hComm, 100, ucsFpcEcrReadRow, "2;1;")) <> 0, ucsTriTrue, ucsTriFalse)
    End If
    IsEcr = (m_uConfig.IsEcr = ucsTriTrue)
End Property

Private Property Get pvRowChars() As Long
    Dim bPrevExceptions As Boolean
    
    If m_uConfig.RowChars = 0 Then
        If Left$(GetSerialNumber(), 2) = "DY" Then '--- daisy
            bPrevExceptions = pvSetExceptions(False)
            m_uConfig.RowChars = At(Split(SendCommand(ucsFpcInfoDaisyConsts), ","), 10)
            Exceptions = bPrevExceptions
        End If
        If m_uConfig.RowChars <= 0 Then
            m_uConfig.RowChars = DEF_COMMENT_LEN
        End If
    End If
    pvRowChars = m_uConfig.RowChars
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function AutodetectDevices(vPorts As Variant) As Variant
    Const FUNC_NAME     As String = "AutodetectDevices"
    Const STR_SPEEDS    As String = "115200|38400|9600|19200|57600"
    Dim lIdx            As Long
    Dim sText           As String
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
    
    On Error GoTo EH
    sText = Chr$(1) & Chr$(&H24) & Chr$(&H20) & Chr(&H20) & Chr$(5) & Chr$(&H30) & Chr$(&H30) & Chr$(&H30) & Chr$(&H30) & Chr$(3)
    vRetVal = vPorts
    For Each vSpeed In Split(STR_SPEEDS, "|")
        ReDim aPorts(0 To UBound(vPorts)) As UcsAutodetect
        ReDim hEvents(0 To UBound(vPorts)) As Long
        lEventsCount = 0
        For lIdx = 0 To UBound(aPorts)
            With aPorts(lIdx)
                If Not IsArray(vRetVal(lIdx)) And LenB(At(vPorts, lIdx)) <> 0 Then
                    sError = vbNullString
                    .hComm = pvOpenPort(At(vPorts, lIdx) & "," & vSpeed, sError:=sError)
                    If .hComm <> INVALID_HANDLE_VALUE And LenB(sError) = 0 Then
                        .WriteOver.hEvent = CreateEvent(0, 1, 0, 0)
                        .ReadOver.hEvent = CreateEvent(0, 1, 0, 0)
                    Else
                        Call CloseHandle(.hComm)
                        .hComm = 0
                    End If
                End If
            End With
        Next
        For lIdx = 0 To UBound(aPorts)
            With aPorts(lIdx)
                If .hComm <> 0 Then
                    lResult = WriteFile(.hComm, ByVal sText, Len(sText), lActual, .WriteOver)
                    If lResult = 0 And Err.LastDllError = ERROR_IO_PENDING Then
                        hEvents(lEventsCount) = .WriteOver.hEvent
                        lEventsCount = lEventsCount + 1
                    End If
                End If
            End With
        Next
        If lEventsCount > 0 Then
'            lResult = WaitForMultipleObjects(lEventsCount, hEvents(0), 1, 100)
            Call Sleep(100)
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
            lResult = WaitForMultipleObjects(lEventsCount, hEvents(0), 1, 100)
            lEventsCount = 0
        End If
        For lIdx = 0 To UBound(aPorts)
            With aPorts(lIdx)
                If .hComm <> 0 Then
                    If InStr(.Buffer, Chr$(0)) > 0 Then
                        .Buffer = StrConv(.Buffer, vbUnicode)
                        .Buffer = Left$(.Buffer, InStr(.Buffer, Chr$(0)) - 1)
                    End If
                    If .Buffer = Chr$(&H15) Then '--- NAK (&H15)
                        If pvGetDeviceInfo(.hComm, 1000, Empty, sProtocol, sModel) Then
                            vRetVal(lIdx) = Array(At(vPorts, lIdx), vSpeed, sProtocol, sModel)
                        End If
                    End If
                End If
            End With
        Next
        For lIdx = 0 To UBound(aPorts)
            With aPorts(lIdx)
                If .hComm <> 0 Then
                    Call CloseHandle(.hComm)
                    Call CloseHandle(.WriteOver.hEvent)
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
            Optional ByVal lTimeout As Long, _
            Optional ByVal lCashDeskNo As Long) As Boolean
    Const FUNC_NAME     As String = "Init"
    Dim vSplit          As Variant
    Dim sSerial         As String
    Dim bHasSwitchesInfo As Boolean
    
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
    If lCashDeskNo <= 0 Then
        m_lCashDeskNo = 1
    Else
        m_lCashDeskNo = lCashDeskNo
    End If
    vSplit = Split(sDevice, ",")
    m_sDevice = Trim$(At(vSplit, 0, "COM1")) & "," & C_Lng(At(vSplit, 1, "9600")) & "," & C_Lng(At(vSplit, 2, "8")) & "," & _
        IIf(UCase$(At(vSplit, 3, "N")) = "Y", "Y", "N") & "," & C_Lng(At(vSplit, 4, "1"))
    m_hComm = pvOpenPort(m_sDevice, m_sLastError)
    If LenB(m_sLastError) <> 0 Then
        GoTo QH
    End If
    m_vInfoDiagnostics = Split(pvSendCommand(m_hComm, m_lTimeout, ucsFpcInfoDiagnostics, sError:=m_sLastError), ",")
    If LenB(m_sLastError) <> 0 Then
        GoTo QH
    End If
    '--- load config values
    sSerial = At(m_vInfoDiagnostics, 4)
    bHasSwitchesInfo = Len(At(m_vInfoDiagnostics, 3)) > 2 '--- heuristic check for recent version of firmware
    With m_uConfig
        .AbsoluteDiscount = C_Bool(GetConfigValue(sSerial, "AbsoluteDiscount", bHasSwitchesInfo))
        .NegativePrices = C_Bool(GetConfigValue(sSerial, "NegativePrices", Not bHasSwitchesInfo))
        .RowChars = C_Lng(GetConfigValue(sSerial, "RowChars", 0))
        .MaxReceiptRows = C_Lng(GetConfigValue(sSerial, "MaxReceiptRows", DEF_MAX_RECEIPT_ROWS))
        If .MaxReceiptRows <= 0 Then
            .MaxReceiptRows = DEF_MAX_RECEIPT_ROWS
        End If
        .MinDiscount = C_Dbl(GetConfigValue(sSerial, "MinDiscount", DEF_MIN_DISCOUNT))
        If .MinDiscount >= 0 Then
            .MinDiscount = DEF_MIN_DISCOUNT
        End If
        .MaxDiscount = C_Dbl(GetConfigValue(sSerial, "MaxDiscount", DEF_MAX_DISCOUNT))
        If .MaxDiscount <= 0 Then
            .MaxDiscount = DEF_MAX_DISCOUNT
        End If
        .MaxPaymentLen = C_Lng(GetConfigValue(sSerial, "MaxPaymentLen", 0))
        If .MaxPaymentLen <= 0 Then
            .MaxPaymentLen = DEF_MAX_PAYMENT_LEN
        End If
        .Charset = ucsCrsAuto
        .IsEcr = ucsTri_Undefined
    End With
    '--- success
    Init = True
QH:
    If LenB(m_sLastError) <> 0 Then
        Err.Raise vbObjectError, , m_sLastError
    End If
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
    If LenB(m_sLastError) <> 0 Then
        GoTo QH
    End If
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
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function SendCommand(ByVal eCmd As UcsIclCommandsEnum, Optional sData As String) As String
    Const FUNC_NAME     As String = "SendCommand"
    Dim vResult         As Variant
    Dim lChar           As Long
    
    On Error GoTo EH
    m_sLastError = vbNullString
    '--- auto-sense encoding
    If m_uConfig.Charset = ucsCrsAuto Then
        vResult = Split(pvSendCommand(m_hComm, m_lTimeout, ucsFpcInitDecimals, sError:=m_sLastError), ",")
        If LenB(m_sLastError) <> 0 Then
            Err.Raise vbObjectError, , m_sLastError
        End If
        '--- first char of currency name
        lChar = Asc(Zn(Trim$(At(vResult, 2)), " "))
        m_uConfig.Charset = IIf(lChar >= MIK_OFFSET And lChar < MIK_OFFSET + &H40, ucsCrsMik, ucsCrsWindows)
    End If
    SendCommand = pvSendCommand(m_hComm, m_lTimeout, eCmd, sData, m_sLastError)
    RaiseEvent CommandComplete(eCmd, sData, SendCommand)
    If LenB(m_sLastError) <> 0 Then
        Err.Raise vbObjectError, , m_sLastError
    End If
    If Exceptions Then
        If Status(ucsStbPrintingError) Then
            m_sLastError = ErrorText
            Err.Raise vbObjectError, , m_sLastError
        End If
        If Memory(ucsFmeMemoryError) Then
            m_sLastError = MemoryText
            Err.Raise vbObjectError, , m_sLastError
        End If
    End If
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
    Dim sCity           As String
    Dim sAddress        As String
    
    On Error GoTo EH
    ReDim m_uRow(0 To 10) As UcsRowData
    m_lRowCount = 0
    m_bFiscalPaymentExecuted = False
    With m_uRow(pvAddRow())
        .RowType = ucsRowInit
        .InitReceiptType = LimitLong(ReceiptType, 1, [_ucsFscRetMax] - 1)
        .InitOperatorCode = Zn(OperatorCode, "1")
        .InitOperatorName = OperatorName
        .InitOperatorPassword = Zn(OperatorPassword, DEF_PASSWORD)
        SplitCgAddress Trim$(InvCity) & vbCrLf & Trim$(InvAddress), sCity, sAddress, pvRowChars - 2
        .InitInvData = Array(InvNumber, InvTaxNumber, InvVatNumber, InvCompany, sCity, sAddress, InvMol, InvReceiver)
    End With
    '--- success
    StartReceipt = True
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function AddPLU( _
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
        .PluName = RTrim$(pvSafeText(Name))
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
            Optional ByVal WordWrap As Boolean, _
            Optional ByVal BeforeIndex As Long) As Boolean
    Const FUNC_NAME     As String = "AddLine"
    Dim uRow            As UcsRowData
    
    On Error GoTo EH
    '--- sanity check
    If m_lRowCount = 0 Then
        Err.Raise vbObjectError, , pvInternal(ucsErrNoReceiptStarted)
    End If
    With uRow
        .RowType = ucsRowLine
        .LineText = RTrim$(pvSafeText(Line))
        .LineWordWrap = WordWrap
        .PrintRowType = m_uRow(0).InitReceiptType
    End With
    pvInsertRow BeforeIndex, uRow
    '--- success
    AddLine = True
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function AddDiscount( _
            ByVal DiscType As UcsFiscalDiscountTypeEnum, _
            Value As Double, _
            Optional ByVal BeforeIndex As Long) As Boolean
    Const FUNC_NAME     As String = "AddDiscount"
    Dim uRow            As UcsRowData
    Dim lIdx            As Long
    
    On Error GoTo EH
    '--- sanity check
    If m_lRowCount = 0 Then
        Err.Raise vbObjectError, , pvInternal(ucsErrNoReceiptStarted)
    End If
    Select Case DiscType
    Case ucsFscDstTotal
        AddPLU Printf(IIf(Value > 0, pvText(ucsTxtSurcharge), pvText(ucsTxtDiscount)), vbNullString), Value, BeforeIndex:=BeforeIndex
    Case ucsFscDstPlu
        For lIdx = IIf(BeforeIndex <> 0, BeforeIndex, m_lRowCount) - 1 To 0 Step -1
            With m_uRow(lIdx)
                If .RowType = ucsRowPlu Then
                    .DiscType = DiscType
                    .DiscValue = Round(Value, 2)
                    Exit For
                End If
            End With
        Next
    Case Else
        With uRow
            .RowType = ucsRowDiscount
            .DiscType = DiscType ' LimitLong(DiscType, 1, [_ucsFscDstMax] - 1)
            .DiscValue = Round(Value, 2)
            .PrintRowType = m_uRow(0).InitReceiptType
        End With
        pvInsertRow BeforeIndex, uRow
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
            Amount As Double) As Boolean
    Const FUNC_NAME     As String = "AddPayment"
    
    On Error GoTo EH
    '--- sanity check
    If m_lRowCount = 0 Then
        Err.Raise vbObjectError, , pvInternal(ucsErrNoReceiptStarted)
    End If
    With m_uRow(pvAddRow())
        .RowType = ucsRowPayment
        .PmtType = LimitLong(Number, 1, 8)
        .PmtName = pvSafeText(Name)
        .PmtAmount = Round(Amount, 2)
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
    Dim bPrevExceptions As Boolean
    Dim lRow            As Long
    Dim lIdx            As Long
    Dim vSplit          As Variant
    Dim sType           As String
    Dim bInLogin        As Boolean
    Dim sData           As String
    Dim uCtx            As UcsContext
    Dim dblTotal        As Double
    Dim dblDisc         As Double
    Dim vTemp           As Variant
    
    '--- sanity check
    If m_lRowCount = 0 Then
        Err.Raise vbObjectError, , pvInternal(ucsErrNoReceiptStarted)
    End If
    If Not CancelReceipt() Then
        Err.Raise vbObjectError, , pvInternal(ucsErrCannotCancelReceipt)
    End If
    On Error GoTo EH
    '--- fix fiscal receipts with for more than 99 PLUs
    If m_uRow(0).InitReceiptType = ucsFscRetFiscal Or m_uRow(0).InitReceiptType = ucsFscRetInvoiceFiscal Then
        pvConvertExtraRows
    End If
    '--- append final payment (total)
    With m_uRow(pvAddRow())
        .RowType = ucsRowPayment
        .PrintRowType = m_uRow(0).InitReceiptType
    End With
    '--- ring sale
    For lRow = 0 To m_lRowCount - 1
        With m_uRow(lRow)
        Select Case .RowType
        Case ucsRowInit
            Select Case .InitReceiptType
            Case ucsFscRetFiscal, ucsFscRetInvoiceFiscal
                bPrevExceptions = pvSetExceptions(False)
                sData = At(Split(SendCommand(ucsFpcInfoOperator, .InitOperatorCode), ","), 5)
                If LenB(.InitOperatorName) <> 0 Then
                    If .InitOperatorName <> sData Then
                        SendCommand ucsFpcInitOperatorName, .InitOperatorCode & "," & .InitOperatorPassword & "," & Left$(.InitOperatorName, 24)
                        If Status(ucsStbPrintingError) Then
                            SendCommand ucsFpcInitOperatorName, .InitOperatorCode & "," & .InitOperatorPassword & "," & Left$(.InitOperatorName, LimitLong(Len(sData), 10))
                        End If
                    End If
                Else
                    .InitOperatorName = RTrim$(sData)
                End If
                If LenB(.InitInvData(ucsInvNumber)) <> 0 Then
                    SendCommand ucsFpcInitInvoiceNo, .InitInvData(ucsInvNumber) & "," & .InitInvData(ucsInvNumber)
                End If
                Exceptions = bPrevExceptions
                bInLogin = True
                SendCommand ucsFpcFiscalOpen, .InitOperatorCode & "," & .InitOperatorPassword & "," & m_lCashDeskNo & _
                    IIf(.InitReceiptType = ucsFscRetInvoiceFiscal, ",I", vbNullString)
                bInLogin = False
            Case ucsFscRetNonfiscal, ucsFscRetInvoiceNonfiscal
                If LenB(.InitOperatorName) <> 0 Then
                    .InitOperatorName = At(Split(SendCommand(ucsFpcInfoOperator, .InitOperatorCode), ","), 5)
                End If
                SendCommand ucsFpcNonFiscalOpen
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
            If .DiscType = ucsFscDstPlu Then
                dblDisc = Round(dblTotal * .DiscValue / 100#, 2)
            ElseIf .DiscType = ucsFscDstPluAbs Then
                dblDisc = Round(.DiscValue, 2)
            Else
                dblDisc = 0
            End If
            uCtx.GrpTotal(.PluVatGroup) = Round(uCtx.GrpTotal(.PluVatGroup) + dblTotal + dblDisc, 2)
            Select Case .PrintRowType
            Case ucsFscRetFiscal, ucsFscRetInvoiceFiscal
                vSplit = WrapText(.PluName, pvRowChars - 5)
                For lIdx = 0 To UBound(vSplit) - 2
                    pvPrintLine C_Str(vSplit(lIdx))
                Next
                sData = At(vSplit, UBound(vSplit))
                If UBound(vSplit) > 0 Then
                    sData = At(vSplit, UBound(vSplit) - 1) & vbLf & sData
                End If
                sData = sData & vbTab & Chr$(191 + .PluVatGroup) & pvSafeFormat(.PluPrice, "0.00")
                If .PluQuantity <> 1 Then
                    sData = sData & "*" & pvSafeFormat(.PluQuantity, "0.000")
                End If
                Select Case .DiscType
                Case ucsFscDstPlu, ucsFscDstPluAbs
                    sData = sData & IIf(.DiscType = ucsFscDstPlu, ",", ";") & pvSafeFormat(.DiscValue, "0.00")
                End Select
                SendCommand ucsFpcFiscalSell, sData
            Case ucsFscRetNonfiscal, ucsFscRetInvoiceNonfiscal
                If uCtx.PluCount = 1 Then
                    pvPrintLine AlignText(Format$(m_uRow(0).InitOperatorCode, "0000") & " " & m_uRow(0).InitOperatorName, Format$(m_lCashDeskNo, "00"), pvRowChars)
                End If
                If .PrintRowType = ucsFscRetNonfiscal Then
                    vSplit = WrapText(.PluName, pvRowChars - 7)
                    If .PluQuantity <> 1 Then
                        If UBound(vSplit) = 0 Then
                            ReDim Preserve vSplit(0 To 1) As Variant
                            vSplit(1) = vSplit(0)
                            vSplit(0) = vbNullString
                        End If
                        sData = " " & IIf(.PluPrice < 0, "-", vbNullString) & pvSafeFormat(.PluQuantity, "0.000") & " x " & pvSafeFormat(Abs(.PluPrice), "0.00")
                        lIdx = UBound(vSplit) - 1
                        If Len(vSplit(lIdx)) + Len(sData) > pvRowChars - 2 Then
                            vTemp = WrapText(vSplit(lIdx), pvRowChars - Len(sData) - 3)
                            vSplit(lIdx) = Trim$(At(vTemp, 0))
                            vSplit(UBound(vSplit)) = Trim$(At(vTemp, 1) & " " & vSplit(UBound(vSplit)))
                        End If
                        vSplit(lIdx) = AlignText(vSplit(lIdx), sData, pvRowChars - 2)
                    End If
                    vSplit(UBound(vSplit)) = AlignText(vSplit(UBound(vSplit)), " " & pvSafeFormat(dblTotal, "0.00") & " " & Chr$(191 + .PluVatGroup), pvRowChars)
                    pvPrintLine vSplit
                Else
                    vSplit = WrapText(.PluName, pvRowChars - 2)
                    For lIdx = 0 To UBound(vSplit)
                        pvPrintLine C_Str(vSplit(lIdx))
                    Next
                    pvPrintLine AlignText(pvText(ucsTxtPluPrice), pvSafeFormat(.PluPrice, "0.00"), pvRowChars - 2)
                    pvPrintLine AlignText(pvText(ucsTxtPluQuantity), pvSafeFormat(.PluQuantity, "0.000"), pvRowChars - 2)
                    pvPrintLine AlignText(pvText(ucsTxtPluTotal), pvSafeFormat(dblTotal, "0.00") & " " & Chr$(191 + .PluVatGroup), pvRowChars)
                End If
                Select Case .DiscType
                Case ucsFscDstPlu, ucsFscDstPluAbs
                    sData = AlignText(Printf(IIf(dblDisc > 0, pvText(ucsTxtSurcharge), pvText(ucsTxtDiscount)), IIf(.DiscType = ucsFscDstPlu, pvSafeFormat(Abs(.DiscValue), "0.00") & " %", vbNullString)), _
                        pvSafeFormat(dblDisc, "0.00") & " " & Chr$(191 + .PluVatGroup), pvRowChars)
                    pvPrintLine sData
                End Select
            End Select
        Case ucsRowDiscount
            Select Case .PrintRowType
            Case ucsFscRetFiscal, ucsFscRetInvoiceFiscal
                Select Case .DiscType
                Case ucsFscDstSubtotal
                    SendCommand ucsFpcFiscalSubtotal, "10," & pvSafeFormat(.DiscValue, "0.00")
                Case ucsFscDstSubtotalAbs
                    SendCommand ucsFpcFiscalSubtotal, "10;" & pvSafeFormat(.DiscValue, "0.00")
                End Select
            Case ucsFscRetNonfiscal, ucsFscRetInvoiceNonfiscal
                Select Case .DiscType
                Case ucsFscDstSubtotal
                    pvPrintLine AlignText(pvText(ucsTxtSubTotal), pvSafeFormat(SumArray(uCtx.GrpTotal), "0.00"), pvRowChars - 2)
                    For lIdx = 1 To UBound(uCtx.GrpTotal)
                        If uCtx.GrpTotal(lIdx) <> 0 Then
                            dblTotal = Round(uCtx.GrpTotal(lIdx) * .DiscValue / 100#, 2)
                            pvPrintLine AlignText(Printf(IIf(dblTotal > 0, pvText(ucsTxtSurcharge), pvText(ucsTxtDiscount)), pvSafeFormat(Abs(.DiscValue), "0.00") & " %"), _
                                pvSafeFormat(dblTotal, "0.00") & " " & Chr$(191 + lIdx), pvRowChars)
                        End If
                    Next
                Case ucsFscDstSubtotalAbs
                    '--- ToDo: impl absolute subtotal discount
                End Select
            End Select
            Select Case .DiscType
            Case ucsFscDstSubtotal
                For lIdx = 1 To UBound(uCtx.GrpTotal)
                    If uCtx.GrpTotal(lIdx) <> 0 Then
                        dblTotal = Round(uCtx.GrpTotal(lIdx) * .DiscValue / 100#, 2)
                        uCtx.GrpTotal(lIdx) = Round(uCtx.GrpTotal(lIdx) + dblTotal, 2)
                    End If
                Next
            Case ucsFscDstSubtotalAbs
            End Select
        Case ucsRowLine
            If .LineWordWrap Then
                vSplit = WrapText(.LineText, pvRowChars)
            Else
                vSplit = Array(Left$(.LineText, pvRowChars))
            End If
            Select Case .PrintRowType
            Case ucsFscRetFiscal, ucsFscRetInvoiceFiscal
                pvPrintLine vSplit
            Case Else
                If uCtx.PmtPrinted And Not uCtx.ChangePrinted Then
                    dblTotal = Round(SumArray(uCtx.GrpTotal) - uCtx.Paid, 2)
                    If dblTotal > 0 Then
                        pvPrintLine AlignText(GetPaymentName(1), pvSafeFormat(dblTotal, "0.00"), pvRowChars - 2)
                    ElseIf dblTotal < 0 Then
                        pvPrintLine AlignText(pvText(ucsTxtFooterChange), pvSafeFormat(-dblTotal, "0.00"), pvRowChars - 2)
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
                If .PmtType <> 0 Then
                    sType = Mid$("PDCNIJKL", .PmtType, 1)
                    If .PmtType > 4 Then
                        bPrevExceptions = pvSetExceptions(False)
                        SendCommand ucsFpcInitPaymentType, sType & "," & Left$(.PmtName, m_uConfig.MaxPaymentLen)
                        If Status(ucsStbPrintingError) Then
                            SendCommand ucsFpcInitDaisyText, "P" & (56 + .PmtType) & "," & Left$(.PmtName, m_uConfig.MaxPaymentLen)
                        End If
                        Exceptions = bPrevExceptions
                    End If
                    SendCommand ucsFpcFiscalPayment, vbTab & sType & pvSafeFormat(.PmtAmount, "0.00")
                    uCtx.PmtPrinted = True
                Else
                    '--- check if anything left for payment
                    vSplit = Split(SendCommand(ucsFpcInfoTransaction, "T"), ",")
                    If C_Dbl(At(vSplit, 2)) > C_Dbl(At(vSplit, 3)) Or LenB(At(vSplit, 2)) = 0 Or Not uCtx.PmtPrinted Then
                        bPrevExceptions = pvSetExceptions(False)
                        SendCommand ucsFpcFiscalPayment, vbTab
                        Exceptions = bPrevExceptions
                    End If
                    If .PrintRowType = ucsFscRetInvoiceFiscal Then
                        With m_uRow(0)
                            SendCommand ucsFpcFiscalCgInfo, _
                                Left$(.InitInvData(ucsInvTaxNumber), 13) & vbTab & _
                                Left$(.InitOperatorName, 30) & vbTab & _
                                Left$(Zn(C_Str(.InitInvData(ucsInvReceiver)), .InitInvData(ucsInvMol)), 25) & vbTab & _
                                Left$(.InitInvData(ucsInvCompany), 27) & vbTab & _
                                Left$(.InitInvData(ucsInvVatNumber), 12) & vbTab & _
                                .InitInvData(ucsInvCity) & IIf(LenB(.InitInvData(ucsInvAddress)) <> 0, _
                                    vbLf & .InitInvData(ucsInvAddress), vbNullString)
                        End With
                    End If
                    SendCommand ucsFpcFiscalClose
                End If
            Case ucsFscRetNonfiscal, ucsFscRetInvoiceNonfiscal
                If Not uCtx.PmtPrinted Then
                    If uCtx.PluCount > 0 Or .PmtType <> 0 Then
                        pvPrintLine String$(pvRowChars, "-")
                        pvPrintLine AlignText(pvText(ucsTxtFooterTotal), pvSafeFormat(SumArray(uCtx.GrpTotal), "0.00"), pvRowChars - 2)
                        If .PrintRowType = ucsFscRetInvoiceNonfiscal Then
                            '--- print invoice totals by VAT group
                            vSplit = Split(SendCommand(ucsFpcInfoTaxRates), ",")
                            For lIdx = 1 To UBound(uCtx.GrpTotal)
                                If uCtx.GrpTotal(lIdx) <> 0 Then
                                    dblTotal = Round(uCtx.GrpTotal(lIdx) / (1 + C_Dbl(At(vSplit, lIdx - 1)) / 100#), 2)
                                    pvPrintLine AlignText(Printf(pvText(ucsTxtGroupTotal), Chr$(191 + lIdx)), pvSafeFormat(uCtx.GrpTotal(lIdx), "0.00"), pvRowChars - 2)
                                    pvPrintLine AlignText(Printf(pvText(ucsTxtGroupVat), Chr$(191 + lIdx), pvSafeFormat(C_Dbl(At(vSplit, lIdx - 1)), "0.00") & "%"), _
                                        pvSafeFormat(Round(uCtx.GrpTotal(lIdx) - dblTotal, 2), "0.00"), pvRowChars - 2)
                                    pvPrintLine AlignText(pvText(ucsTxtGroupNetto), pvSafeFormat(dblTotal, "0.00"), pvRowChars - 2)
                                End If
                            Next
                            pvPrintLine String$(pvRowChars, "-")
                        End If
                        uCtx.PmtPrinted = True
                    End If
                End If
                If .PmtType <> 0 Then
                    If LenB(.PmtName) = 0 Then
                        .PmtName = GetPaymentName(.PmtType)
                    End If
                    pvPrintLine AlignText(.PmtName, pvSafeFormat(.PmtAmount, "0.00"), pvRowChars - 2)
                    uCtx.Paid = Round(uCtx.Paid + .PmtAmount, 2)
                Else
                    If uCtx.PmtPrinted And Not uCtx.ChangePrinted Then
                        dblTotal = Round(SumArray(uCtx.GrpTotal) - uCtx.Paid, 2)
                        If dblTotal > 0 Then
                            pvPrintLine AlignText(GetPaymentName(1), pvSafeFormat(dblTotal, "0.00"), pvRowChars - 2)
                        ElseIf dblTotal < 0 Then
                            pvPrintLine AlignText(pvText(ucsTxtFooterChange), pvSafeFormat(-dblTotal, "0.00"), pvRowChars - 2)
                        End If
                        uCtx.ChangePrinted = True
                    End If
                    If .PrintRowType = ucsFscRetInvoiceNonfiscal Then
                        With m_uRow(0)
                            pvPrintLine AlignText(pvText(ucsTxtFooterSeller), Chr$(1) & ".", pvRowChars)
                            pvPrintLine AlignText(vbNullString, "/" & .InitOperatorName & "/", pvRowChars)
                            pvPrintLine String$(pvRowChars, "-")
                            pvPrintLine AlignText(pvText(ucsTxtFooterReceiver), Chr$(1) & ".", pvRowChars)
                            pvPrintLine AlignText(vbNullString, "/" & Zn(C_Str(.InitInvData(ucsInvReceiver)), .InitInvData(ucsInvMol)) & "/", pvRowChars)
                            pvPrintLine pvText(ucsTxtFooterBuyer) & .InitInvData(ucsInvCompany)
                            pvPrintLine pvText(ucsTxtFooterTaxNo) & .InitInvData(ucsInvTaxNumber)
                            If LenB(.InitInvData(ucsInvVatNumber)) <> 0 Then
                                pvPrintLine pvText(ucsTxtFooterVatNo) & .InitInvData(ucsInvVatNumber)
                            End If
                            pvPrintLine .InitInvData(ucsInvCity)
                            pvPrintLine .InitInvData(ucsInvAddress)
                        End With
                    End If
                    sData = GetFooterText(1)
                    If LenB(sData) <> 0 Then
                        pvPrintLine CenterText(sData, pvRowChars)
                    End If
                    sData = GetFooterText(2)
                    If LenB(sData) <> 0 Then
                        pvPrintLine CenterText(sData, pvRowChars)
                    End If
                    If uCtx.PluCount > 0 Then
                        sData = IIf(uCtx.PluCount = 1, pvText(ucsTxtFooterArticleSingle), Printf(pvText(ucsTxtFooterArticles), uCtx.PluCount))
                        pvPrintLine CenterText(sData, pvRowChars)
                    End If
                    SendCommand ucsFpcNonFiscalClose
                End If
            End Select
        End Select
        End With
    Next
    '--- success
    EndReceipt = True
    Exit Function
EH:
    m_sLastError = Err.Description
    PrintError FUNC_NAME
    If bInLogin Then
        m_sLastError = m_sLastError & pvInternal(ucsErrInvalidOperatorPassword)
    End If
    If uCtx.PluCount > 0 Then
        m_sLastError = Printf(pvInternal(ucsErrErrorInLine), uCtx.PluCount, m_sLastError)
    End If
    Err.Raise vbObjectError, , m_sLastError
    Resume
End Function

Public Function CopyLastReceipt() As Boolean
    On Error GoTo EH_Cancel
    SendCommand ucsFpcPrintDuplicate, "1"
    '--- success
    CopyLastReceipt = True
EH_Cancel:
End Function

Public Function GetClock() As Date
    Const FUNC_NAME     As String = "GetClock"
    Dim bPrevExceptions As Boolean
    Dim vSplit          As Variant
    
    On Error GoTo EH
    bPrevExceptions = pvSetExceptions(False)
    vSplit = Split(SendCommand(ucsFpcInfoDateTime), " ")
    Exceptions = bPrevExceptions
    GetClock = C_Date(At(vSplit, 0)) + C_Date(At(vSplit, 1))
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Sub SetClock(ByVal NewDate As Date)
    Const FUNC_NAME     As String = "SetClock"
    
    On Error GoTo EH
    SendCommand ucsFpcInitDateTime, Format$(NewDate, "dd-MM-yy") & " " & Format$(NewDate, "hh:mm:ss")
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Public Function RunXReport(ByVal ReportType As UcsFiscalReportsTypeEnum) As Boolean
    Const FUNC_NAME     As String = "RunXReport"
    Dim bPrevExceptions As Boolean
    Dim eCmd            As UcsIclCommandsEnum
    
    If Not CancelReceipt() Then
        Err.Raise vbObjectError, , pvInternal(ucsErrCannotCancelReceipt)
    End If
    On Error GoTo EH
    Select Case ReportType
    Case ucsFscRstDaily
        eCmd = ucsFpcPrintDailyReport
    Case ucsFscRstDailyByItems
        eCmd = ucsFpcPrintDailyReportItems
    Case ucsFscRstDailyByDepartment
        eCmd = ucsFpcPrintDailyReportDepartments
    Case ucsFscRstDailyByItemsAndDepartment
        eCmd = ucsFpcPrintDailyReportItemsDepartments
    End Select
    If eCmd <> 0 Then
        bPrevExceptions = pvSetExceptions(False)
        SendCommand eCmd, "2"
        If Status(ucsStbPrintingError) Then
            '--- daisy: print by departaments
            If eCmd = ucsFpcPrintDailyReportDepartments Then
                SendCommand ucsFpcPrintDailyReport, "9"
            ElseIf eCmd = ucsFpcPrintDailyReportItemsDepartments Then
                SendCommand ucsFpcPrintDailyReportItems, "9"
            Else
                Exceptions = bPrevExceptions
                m_sLastError = ErrorText
                Err.Raise vbObjectError, , m_sLastError
            End If
        End If
        Exceptions = bPrevExceptions
        '--- success
        RunXReport = True
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function RunZReport(ByVal ReportType As UcsFiscalReportsTypeEnum) As Boolean
    Const FUNC_NAME     As String = "RunZReport"
    Dim bPrevExceptions As Boolean
'    Dim lIdx            As Long
'    Dim lCount          As Long
    Dim eCmd            As UcsIclCommandsEnum
    
    If Not CancelReceipt() Then
        Err.Raise vbObjectError, , pvInternal(ucsErrCannotCancelReceipt)
    End If
    On Error GoTo EH
    Select Case ReportType
    Case ucsFscRstDaily
        eCmd = ucsFpcPrintDailyReport
    Case ucsFscRstDailyByItems
        eCmd = ucsFpcPrintDailyReportItems
    Case ucsFscRstDailyByDepartment
        eCmd = ucsFpcPrintDailyReportDepartments
    Case ucsFscRstDailyByItemsAndDepartment
        eCmd = ucsFpcPrintDailyReportItemsDepartments
    End Select
    If eCmd <> 0 Then
        bPrevExceptions = pvSetExceptions(False)
        '--- clear electronic control tape (if present)
'        lCount = C_Lng(At(Split(SendCommand(ucsFpcInitEcTape, "I"), ","), 1))
'        If Not Status(ucsStbPrintingError) Then
'            For lIdx = 1 To lCount
'                SendCommand ucsFpcInitEcTape, IIf(lIdx = 1, "PS", "CS")
'            Next
'            SendCommand ucsFpcInitEcTape, "E"
'        End If
        SendCommand eCmd, "0"
        If Status(ucsStbPrintingError) Then
            '--- daisy: print by departments
            If eCmd = ucsFpcPrintDailyReportDepartments Then
                SendCommand ucsFpcPrintDailyReport, "8"
            ElseIf eCmd = ucsFpcPrintDailyReportItemsDepartments Then
                SendCommand ucsFpcPrintDailyReportItems, "8"
            Else
                Exceptions = bPrevExceptions
                m_sLastError = ErrorText
                Err.Raise vbObjectError, , m_sLastError
            End If
        End If
        Exceptions = bPrevExceptions
        '--- success
        RunZReport = True
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function RunPeriodReport(ByVal ReportType As UcsFiscalReportsTypeEnum, ByVal StartDate As Date, ByVal EndDate As Date) As Boolean
    Const FUNC_NAME     As String = "RunPeriodReport"
    Dim eCmd            As UcsIclCommandsEnum
    
    If Not CancelReceipt() Then
        Err.Raise vbObjectError, , pvInternal(ucsErrCannotCancelReceipt)
    End If
    On Error GoTo EH
    Select Case ReportType
    Case ucsFscRstPeriodShort
        eCmd = ucsFpcPrintReportByDateShort
    Case ucsFscRstPeriodDetailed
        eCmd = ucsFpcPrintReportByDateDetailed
    End Select
    If eCmd <> 0 Then
        SendCommand eCmd, Format$(StartDate, "ddmmyy") & "," & Format$(EndDate, "ddmmyy")
        '--- success
        RunPeriodReport = True
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetLastReceiptNumber() As Long
    Const FUNC_NAME     As String = "GetLastReceiptNumber"
    Dim bPrevExceptions As Boolean
    
    On Error GoTo EH
    bPrevExceptions = pvSetExceptions(False)
    GetLastReceiptNumber = SendCommand(ucsFpcInfoLastReceipt)
    Exceptions = bPrevExceptions
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetLastInvoiceNumber() As String
    Const FUNC_NAME     As String = "GetLastInvoiceNumber"
    Dim bPrevExceptions As Boolean
    
    On Error GoTo EH
    bPrevExceptions = pvSetExceptions(False)
    GetLastInvoiceNumber = At(Split(SendCommand(ucsFpcInitInvoiceNo), ","), 0)
    Exceptions = bPrevExceptions
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetSerialNumber() As String
    Const FUNC_NAME     As String = "GetSerialNumber"
    
    On Error GoTo EH
    GetSerialNumber = At(m_vInfoDiagnostics, 4)
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetFiscalNumber() As String
    Const FUNC_NAME     As String = "GetFiscalNumber"
    
    On Error GoTo EH
    GetFiscalNumber = At(m_vInfoDiagnostics, 5)
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetTaxNumber() As String
    Const FUNC_NAME     As String = "GetTaxNumber"
    Dim bPrevExceptions As Boolean
    
    On Error GoTo EH
    bPrevExceptions = pvSetExceptions(False)
    GetTaxNumber = At(Split(SendCommand(ucsFpcInfoBulstat), ","), 0)
    Exceptions = bPrevExceptions
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetTaxCaption() As String
    Const FUNC_NAME     As String = "GetTaxCaption"
    Dim bPrevExceptions As Boolean
    
    On Error GoTo EH
    bPrevExceptions = pvSetExceptions(False)
    GetTaxCaption = At(Split(SendCommand(ucsFpcInfoBulstat), ","), 1)
    Exceptions = bPrevExceptions
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetHeaderText(ByVal Index As Long) As String
    Const FUNC_NAME     As String = "GetHeaderText"
    Dim bPrevExceptions As Boolean
    
    On Error GoTo EH
    If Index >= 1 And Index <= 6 Then
        bPrevExceptions = pvSetExceptions(False)
        GetHeaderText = Replace(SendCommand(ucsFpcInitHeaderFooter, "I" & (Index - 1)), "^", vbNullString)
        Exceptions = bPrevExceptions
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetFooterText(ByVal Index As Long) As String
    Const FUNC_NAME     As String = "GetFooterText"
    Dim bPrevExceptions As Boolean
    
    On Error GoTo EH
    If Index >= 1 And Index <= 2 Then
        bPrevExceptions = pvSetExceptions(False)
        GetFooterText = Replace(SendCommand(ucsFpcInitHeaderFooter, "I" & (Index + 5)), "^", vbNullString)
        Exceptions = bPrevExceptions
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetPaymentName(ByVal Index As Long) As String
    Const FUNC_NAME     As String = "GetPaymentName"
    Dim bPrevExceptions As Boolean
    
    On Error GoTo EH
    If Index > 4 Then
        bPrevExceptions = pvSetExceptions(False)
        GetPaymentName = SendCommand(ucsFpcInitPaymentType, Chr$(68 + Index))
        Exceptions = bPrevExceptions
        If Status(ucsStbPrintingError) Then
            GetPaymentName = SendCommand(ucsFpcInitDaisyText, "R" & (56 + Index))
        End If
    Else
        GetPaymentName = pvText(ucsTxtPaymentCash + Index - 1)
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function OpenDrawer() As Boolean
    On Error GoTo EH_Cancel
    SendCommand ucsFpcInitOpenDrawer
    '--- success
    OpenDrawer = True
EH_Cancel:
End Function

Public Function CancelReceipt() As Boolean
    Const FUNC_NAME     As String = "CancelReceipt"
    Dim bPrevExceptions As Boolean
    
    On Error GoTo EH
    bPrevExceptions = pvSetExceptions(False)
    If Left$(SendCommand(ucsFpcInfoTransaction), 1) = "1" Then
        If Status(ucsStbFiscalPrinting) Then
            '--- note: FP-3530 can cancel always, FP-550 sometimes can not
            SendCommand ucsFpcFiscalCancel
            If Status(ucsStbFiscalPrinting) Then
                '--- note: this helps if sum payed is less than total receipt sum
                SendCommand ucsFpcFiscalPayment, vbTab
                '--- note: when printing invoice, if no contragent info set then cancel fails!
                SendCommand ucsFpcFiscalCgInfo, "0000000000"
                '--- for FP-550
                SendCommand ucsFpcFiscalClose
            End If
        Else
            SendCommand ucsFpcNonFiscalClose
        End If
    End If
    Exceptions = bPrevExceptions
    '--- success (or failure)
    CancelReceipt = (Left$(SendCommand(ucsFpcInfoTransaction), 1) <> "1")
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

Public Function GetDeviceProtocol() As String
    Const FUNC_NAME     As String = "GetDeviceProtocol"
    
    On Error GoTo EH
    pvGetDeviceInfo m_hComm, m_lTimeout, m_vInfoDiagnostics, GetDeviceProtocol, vbNullString
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetTotalsByVatGroups() As Variant
    Const FUNC_NAME     As String = "GetTotalsByVatGroups"
    Dim bPrevExceptions As Boolean
    Dim vSplit          As Variant
    Dim vRetVal         As Variant
    Dim lIdx            As Long
    
    On Error GoTo EH
    bPrevExceptions = pvSetExceptions(False)
    vSplit = Split(SendCommand(ucsFpcInfoTotalsByVatGroups), ",")
    Exceptions = bPrevExceptions
    If UBound(vSplit) > 0 Then
        ReDim vRetVal(0 To 7) As Double
        For lIdx = 0 To UBound(vRetVal)
            If InStr(At(vSplit, lIdx + 1), ".") > 0 Then
                vRetVal(lIdx) = C_Dbl(Replace(At(vSplit, lIdx + 1), ".", g_sDecimalSeparator))
            Else
                vRetVal(lIdx) = C_Dbl(At(vSplit, lIdx + 1)) / 100#
            End If
        Next
        GetTotalsByVatGroups = vRetVal
    Else
        GetTotalsByVatGroups = EmptyDoubleArray
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function CashDebitCredit(ByVal dblValue As Double) As Variant
    Const FUNC_NAME     As String = "CashDebitCredit"
    Dim vSplit          As Variant
    Dim vRetVal         As Variant
    Dim lIdx            As Long
    
    On Error GoTo EH
    vSplit = Split(SendCommand(ucsFpcAdminCashDebitCredit, IIf(dblValue <> 0, pvSafeFormat(dblValue, "0.00"), vbNullString)), ",")
    If UBound(vSplit) > 0 Then
        ReDim vRetVal(0 To 2) As Double
        For lIdx = 0 To UBound(vRetVal)
            If InStr(At(vSplit, lIdx + 1), ".") > 0 Then
                vRetVal(lIdx) = C_Dbl(Replace(At(vSplit, lIdx + 1), ".", g_sDecimalSeparator))
            Else
                vRetVal(lIdx) = C_Dbl(At(vSplit, lIdx + 1)) / 100#
            End If
        Next
        CashDebitCredit = vRetVal
    Else
        CashDebitCredit = EmptyDoubleArray
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
    RaiseError FUNC_NAME
End Function

Private Function pvSendCommand( _
            ByVal hComm As Long, _
            ByVal lTimeout As Long, _
            ByVal eCmd As UcsIclCommandsEnum, _
            Optional sData As String, _
            Optional sError As String) As String
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
    
    On Error GoTo EH
    DebugLog FUNC_NAME, "eCmd=" & eCmd & ", sData=" & sData
    uOver.hEvent = CreateEvent(0, 1, 0, 0)
    '--- prepare command and transcode cyrillic
    sText = pvTranscode(sData, True)
    Do While eCmd >= &H100
        sText = Chr$(eCmd Mod &H100) & sText
        eCmd = eCmd \ &H100
    Loop
    sText = Chr$(1) & Chr$(&H20 + 4 + Len(sText)) & Chr$(m_lSequence) & Chr$(eCmd) & sText & Chr$(5)
    '--- increment sequence
    m_lSequence = IIf(m_lSequence >= &H7F, &H20, m_lSequence + 1)
    '--- calc checksum
    baText = StrConv(sText, vbFromUnicode)
    For lIdx = 1 To UBound(baText)
        lSum = lSum + baText(lIdx)
    Next
    '--- append checksum and terminator
    sText = sText & Chr$((lSum \ &H1000 Mod &H10) + &H30) & Chr$((lSum \ &H100 Mod &H10) + &H30) & Chr$((lSum \ &H10 Mod &H10) + &H30) & Chr$((lSum \ &H1 Mod &H10) + &H30) & Chr$(3)
    '--- async send
    lResult = WriteFile(hComm, ByVal sText, Len(sText), lWritten, uOver)
    If lResult = 0 Then
        If Err.LastDllError <> ERROR_IO_PENDING Then
            sError = pvInternal(ucsErrWriteFileFailure) & GetApiErr(Err.LastDllError)
            GoTo QH
        End If
    End If
    lResult = WaitForMultipleObjects(1, uOver.hEvent, 0, lTimeout)
    If lResult <> WAIT_OBJECT_0 Then
        DebugLog FUNC_NAME, "CancelIo after WriteFile, lResult=" & lResult
        If lResult = WAIT_TIMEOUT Then
            sError = pvInternal(ucsErrTimeoutWaitingForResponse)
        End If
        Call CancelIo(hComm)
        Call Sleep(100)
        Call FlushFileBuffers(hComm)
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
            sBuffer = String$(1000, 0)
            lResult = ReadFile(hComm, ByVal sBuffer, Len(sBuffer), lActual, uOver)
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
                sBuffer = Left$(sBuffer, lActual)
                '--- check for NAK (&H15)
                If InStr(sBuffer, Chr$(&H15)) > 0 Then
                    DebugLog FUNC_NAME, "NAK, sBuffer=" & sBuffer
                    sError = pvInternal(ucsErrInvalidMessageFormatOrChecksum)
                    GoTo QH
                End If
                '--- remove SYN (&H16)
                sText = sText & Replace(sBuffer, Chr$(&H16), vbNullString)
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
            Exit Do
        End If
        '--- check if response complete (3 = Terminator)
        If Right$(sText, 1) = Chr$(3) Then
            If Left$(sText, 1) = Chr$(1) Then
                lIdx = InStr(sText, Chr$(4))
                If lIdx > 2 Then
                    lSize = Asc(Mid$(sText, 2, 1)) - &H20 - 11
                    '--- check message length
                    If lIdx - 5 = lSize Then
                        pvSetStatus Mid$(sText, lIdx + 1, 6)
                        pvSendCommand = pvTranscode(Mid$(sText, 5, lSize), False)
                        DebugLog FUNC_NAME, "RetVal=" & pvSendCommand & IIf(Status(ucsStbPrintingError), ", ErrorText=" & ErrorText, vbNullString) & IIf(Memory(ucsFmeMemoryError), ", MemoryText=" & ErrorText, vbNullString)
                        GoTo QH
                    Else
                        sError = pvInternal(ucsErrInvalidResponseLength)
                    End If
                Else
                    sError = pvInternal(ucsErrResponseMissingPostamble)
                End If
            Else
                sError = pvInternal(ucsErrResponseMissingPreamble)
            End If
            pvSetStatus vbNullString '--- clear
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

Private Function pvTranscode(sText As String, ByVal bToPrinter As Boolean) As String
    Const FUNC_NAME     As String = "pvTranscode"
    Dim lIdx            As Long
    Dim lChar           As Long
    
    On Error GoTo EH
    #If bToPrinter Then '--- touch arg
    #End If
    pvTranscode = sText
    For lIdx = 1 To Len(pvTranscode)
        lChar = Asc(Mid$(pvTranscode, lIdx, 1))
'        If bToPrinter And m_uConfig.Charset = ucsCrsMik And lChar >= WIN_OFFSET And lChar < WIN_OFFSET + &H40 Then
'            Mid$(pvTranscode, lIdx, 1) = Chr$(lChar + MIK_OFFSET - WIN_OFFSET)
'        ElseIf m_uConfig.Charset <> ucsCrsAuto And lChar >= MIK_OFFSET And lChar < MIK_OFFSET + &H40 Then
'            Mid$(pvTranscode, lIdx, 1) = Chr$(lChar + WIN_OFFSET - MIK_OFFSET)
'        End If
        If m_uConfig.Charset = ucsCrsMik Then
            Mid$(pvTranscode, lIdx, 1) = Chr$(IIf(lChar >= &H80, lChar Xor &H40, lChar))
        End If
    Next
    Exit Function
EH:
    RaiseError FUNC_NAME
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
    RaiseError FUNC_NAME
End Sub

Friend Function frGetFlagsText(ByVal lFlags As Long, vTexts As Variant) As String
    Const FUNC_NAME     As String = "frGetFlagsText"
    Dim lIdx            As Long
    
    On Error GoTo EH
    For lIdx = 0 To UBound(vTexts)
        If (lFlags And (2 ^ lIdx)) <> 0 Then
            If Len(vTexts(lIdx)) > 1 Then
                If LenB(frGetFlagsText) <> 0 Then
                    frGetFlagsText = frGetFlagsText & ", "
                End If
                frGetFlagsText = frGetFlagsText & vTexts(lIdx)
            End If
        End If
    Next
    Exit Function
EH:
    RaiseError FUNC_NAME
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

Private Function pvSetExceptions(ByVal bValue As Boolean) As Boolean
    pvSetExceptions = Exceptions
    Exceptions = bValue
End Function

Private Function pvSafeText(sText As String) As String
    Const FUNC_NAME     As String = "pvSafeText"
    Dim lIdx            As Long
    
    On Error GoTo EH
    pvSafeText = sText
    For lIdx = 0 To 31
        pvSafeText = Replace(pvSafeText, Chr$(lIdx), vbNullString)
    Next
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Function pvSafeFormat(Expression As Variant, Optional Fmt As Variant, Optional sDecimal As String = ".") As String
    Const FUNC_NAME     As String = "pvSafeFormat"
    
    On Error GoTo EH
    pvSafeFormat = Replace(Format$(Expression, Fmt), g_sDecimalSeparator, sDecimal)
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Function pvPrintLine(vText As Variant)
    Const FUNC_NAME     As String = "pvPrintLine"
    Dim vElem               As Variant
    Dim eCmd                As UcsIclCommandsEnum
    
    On Error GoTo EH
    Select Case m_uRow(0).InitReceiptType
    Case ucsFscRetFiscal, ucsFscRetInvoiceFiscal
        eCmd = ucsFpcFiscalText
    Case Else
        eCmd = ucsFpcNonFiscalText
    End Select
    If IsArray(vText) Then
        For Each vElem In vText
            SendCommand eCmd, C_Str(vElem)
        Next
    Else
        SendCommand eCmd, C_Str(vText)
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
    Dim dblDiscTotal    As Double
    Dim dblPrice        As Double
    Dim vSplit          As Variant
    
    On Error GoTo EH
    '--- convert out-of-range discounts to PLU rows
    '--- note: m_lRowCount may change in loop on AddPLU
    Do While lRow < m_lRowCount
        '--- note: 'With' locks m_uRow array and fails if auto-grow needed in AddPLU
'        With m_uRow(lRow)
            If m_uRow(lRow).RowType = ucsRowPlu Then
                dblPrice = m_uRow(lRow).PluPrice
                dblTotal = Round(m_uRow(lRow).PluQuantity * dblPrice, 2)
                dblDiscTotal = Round(dblTotal * m_uRow(lRow).DiscValue / 100#, 2)
                If Not m_uConfig.NegativePrices And dblPrice <= 0 Then
                    vSplit = WrapText(m_uRow(lRow).PluName, pvRowChars - 5)
                    lIdx = LimitLong(UBound(vSplit), , 1)
                    vSplit(lIdx) = AlignText(vSplit(lIdx), pvSafeFormat(dblTotal + dblDiscTotal, "0.00"), pvRowChars)
                    m_uRow(lRow).RowType = ucsRowLine
                    m_uRow(lRow).LineText = vSplit(0)
                    If lIdx > 0 Then
                        AddLine At(vSplit, 1), False, lRow + 1
                        lRow = lRow + 1
                    End If
                    If dblPrice < 0 Then
                        AddDiscount ucsFscDstSubtotalAbs, dblTotal + dblDiscTotal, lRow + 1
                    End If
                ElseIf (m_uRow(lRow).DiscValue < m_uConfig.MinDiscount Or m_uRow(lRow).DiscValue > m_uConfig.MaxDiscount) Then
                    dblDiscount = Limit(m_uRow(lRow).DiscValue, m_uConfig.MinDiscount, m_uConfig.MaxDiscount)
                    If m_uConfig.AbsoluteDiscount Then
                        m_uRow(lRow).DiscType = ucsFscDstPluAbs
                        m_uRow(lRow).DiscValue = dblDiscTotal
                    ElseIf dblDiscTotal = Round(dblTotal * dblDiscount / 100#, 2) Then
                        m_uRow(lRow).DiscValue = dblDiscount
                    Else
                        dblDiscount = m_uRow(lRow).DiscValue
                        m_uRow(lRow).DiscType = 0
                        m_uRow(lRow).DiscValue = 0
                        AddPLU Printf(IIf(dblDiscTotal > 0, pvText(ucsTxtSurcharge), pvText(ucsTxtDiscount)), pvSafeFormat(Abs(dblDiscount), "0.00") & " %"), _
                            dblDiscTotal, 1, m_uRow(lRow).PluVatGroup, lRow + 1
                    End If
                ElseIf m_uRow(lRow).DiscType = ucsFscDstPlu And dblPrice < 0 Then
                    '--- convert PLU discount on void rows
                    If m_uConfig.AbsoluteDiscount Then
                        m_uRow(lRow).DiscType = ucsFscDstPluAbs
                        m_uRow(lRow).DiscValue = dblDiscTotal
                    Else
                        dblDiscount = m_uRow(lRow).DiscValue
                        m_uRow(lRow).DiscType = 0
                        m_uRow(lRow).DiscValue = 0
                        AddPLU Printf(IIf(dblTotal * dblDiscount > 0, pvText(ucsTxtSurcharge), pvText(ucsTxtDiscount)), pvSafeFormat(Abs(dblDiscount), "0.00") & " %"), _
                                dblDiscTotal, 1, m_uRow(lRow).PluVatGroup, lRow + 1
                    End If
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
                            If uSum.GrpTotal(lIdx) <> 0 Then
                                AddPLU Printf(IIf(uSum.GrpTotal(lIdx) * dblDiscount > 0, pvText(ucsTxtSurcharge), pvText(ucsTxtDiscount)), pvSafeFormat(Abs(dblDiscount), "0.00") & " %"), _
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
                    If lCount > m_uConfig.MaxReceiptRows - lTotal Then
                        .PrintRowType = IIf(.PrintRowType = ucsFscRetFiscal, ucsFscRetNonfiscal, ucsFscRetInvoiceNonfiscal)
                        dblTotal = Round(.PluQuantity * .PluPrice, 2)
                        If .DiscType = ucsFscDstPlu Then
                            dblTotal = Round(dblTotal + Round(dblTotal * .DiscValue / 100#, 2), 2)
                        ElseIf .DiscType = ucsFscDstPluAbs Then
                            dblTotal = Round(dblTotal + .DiscValue, 2)
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
            If uCtx.GrpTotal(lIdx) <> 0 Then
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
            Select Case .DiscType
            Case ucsFscDstPlu
                dblTotal = Round(dblTotal + Round(dblTotal * .DiscValue / 100#, 2), 2)
            Case ucsFscDstPluAbs
                dblTotal = Round(dblTotal + .DiscValue, 2)
            End Select
            If .PluVatGroup > 0 Then
                uCtx.GrpTotal(.PluVatGroup) = Round(uCtx.GrpTotal(.PluVatGroup) + dblTotal, 2)
            End If
        ElseIf .RowType = ucsRowDiscount Then
            Select Case .DiscType
            Case ucsFscDstSubtotal
                For lJdx = 1 To UBound(uCtx.GrpTotal)
                    dblTotal = Round(uCtx.GrpTotal(lJdx) * .DiscValue / 100#, 2)
                    uCtx.GrpTotal(lJdx) = Round(uCtx.GrpTotal(lJdx) + dblTotal, 2)
                Next
            Case ucsFscDstSubtotalAbs
                For lJdx = 1 To UBound(uCtx.GrpTotal)
                    If uCtx.GrpTotal(lJdx) <> 0 Then
                        uCtx.GrpTotal(lJdx) = Round(uCtx.GrpTotal(lJdx) - .DiscValue, 2)
                        Exit For
                    End If
                Next
            End Select
        End If
        End With
    Next
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Private Function pvGetDeviceInfo(ByVal hComm As Long, ByVal lTimeout As Long, vResult As Variant, sProtocol As String, sModel As String, Optional sError As String) As Boolean
    Const FUNC_NAME     As String = "pvGetDeviceInfo"
    
    On Error GoTo EH
    If Not IsArray(vResult) Then
        vResult = Split(pvSendCommand(hComm, lTimeout, ucsFpcInfoDiagnostics, sError:=sError), ",")
    End If
    If UBound(vResult) >= 4 And LenB(sError) = 0 Then
        '--- check serial
        If Left$(At(vResult, 4), 2) = "DY" Then
            sProtocol = STR_PROTOCOL_DAISY_ECR
        ElseIf LenB(At(vResult, 4)) <> 0 Then
            sProtocol = STR_PROTOCOL_DATECS_FP
        End If
        sModel = At(vResult, 0)
        If Len(At(vResult, 3)) > 2 Then
            sModel = sModel & " " & At(vResult, 1)
        End If
        '--- success
        pvGetDeviceInfo = True
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Initialize()
    m_hComm = INVALID_HANDLE_VALUE
    m_lTimeout = DEF_TIMEOUT
    m_lSequence = &H20
    m_bExceptions = True
    LocalizedText(ucsFscLciInternalErrors) = STR_INTERNAL
    LocalizedText(ucsFscLciPrinterStatuses) = STR_STATUSES
    LocalizedText(ucsFscLciPrinterErrors) = STR_ERRORS_ONLY
    LocalizedText(ucsFscLciDipSwitchesTexts) = STR_DIP_SWITCHES
    LocalizedText(ucsFscLciMemoryTexts) = STR_MEMORY
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
    IDeviceProtocol_AddLine = AddLine(Line, WordWrap)
End Function

Private Function IDeviceProtocol_AddPayment(ByVal Number As UcsFiscalPaymentTypeEnum, Name As String, Amount As Double, Optional Rate As Double) As Boolean
    IDeviceProtocol_AddPayment = AddPayment(Number, Name, Amount)
End Function

Private Function IDeviceProtocol_AddPLU(Name As String, Price As Double, Quantity As Double, VatGroup As Long) As Boolean
    IDeviceProtocol_AddPLU = AddPLU(Name, Price, Quantity, VatGroup)
End Function

Private Function IDeviceProtocol_AutodetectDevices(Ports As Variant) As Variant
    IDeviceProtocol_AutodetectDevices = AutodetectDevices(Ports)
End Function

Private Function IDeviceProtocol_CancelReceipt() As Boolean
    IDeviceProtocol_CancelReceipt = CancelReceipt()
End Function

Private Function IDeviceProtocol_CashDebitCredit(OperatorCode As String, OperatorPassword As String, ByVal Value As Double) As Variant
    IDeviceProtocol_CashDebitCredit = CashDebitCredit(Value)
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
        IDeviceProtocol_GetDefaultPassword = OperatorCode
    Else
        IDeviceProtocol_GetDefaultPassword = DEF_PASSWORD
    End If
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

Private Sub IDeviceProtocol_SetClock(ByVal NewDate As Date)
    SetClock NewDate
End Sub

Private Sub IDeviceProtocol_SetLocalizedText(ByVal Index As UcsFiscalLocalizedIndexesEnum, Text As String)
    LocalizedText(Index) = Text
End Sub

Private Function IDeviceProtocol_StartReceipt(ByVal ReceiptType As UcsFiscalReceiptTypeEnum, OperatorCode As String, OperatorName As String, OperatorPassword As String, Optional InvNumber As String, Optional InvTaxNumber As String, Optional InvVatNumber As String, Optional InvCompany As String, Optional InvCity As String, Optional InvAddress As String, Optional InvMol As String, Optional InvReceiver As String) As Boolean
    IDeviceProtocol_StartReceipt = StartReceipt(ReceiptType, OperatorCode, OperatorName, OperatorPassword, InvNumber, InvTaxNumber, InvVatNumber, InvCompany, InvCity, InvAddress, InvMol, InvReceiver)
End Function

