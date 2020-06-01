VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cIslProtocol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=========================================================================
'
' UcsFP20 (c) 2008-2020 by Unicontsoft
'
' Unicontsoft Fiscal Printers Component 2.0
'
' This project is licensed under the terms of the MIT license
' See the LICENSE file in the project root for more information
'
'=========================================================================
'
' Implementation of ISL protocol for Datecs (legacy) A and B models, Daisy and Incotex
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "cIslProtocol"
Implements IDeviceProtocol

'=========================================================================
' Public events
'=========================================================================

Event CommandComplete(ByVal lCmd As Long, sData As String, sResult As String)

'=========================================================================
' Public enums
'=========================================================================

Public Enum UcsIslStatusBitsEnum
    ucsIslStbSyntaxError = 2 ^ 0
    ucsIslStbInvalidCommand = 2 ^ 1
    ucsIslStbClockNotSet = 2 ^ 2
    ucsIslStbDeviceError = 2 ^ 4
    ucsIslStbPrintingError = 2 ^ 5
    ucsIslStbSumOverflow = 2 ^ (8 + 0)
    ucsIslStbInvalidFiscalMode = 2 ^ (8 + 1)
    ucsIslStbRamCleared = 2 ^ (8 + 2)
    ucsIslStbRamCorrupted = 2 ^ (8 + 4)
    ucsIslStbOutOfPaper = 2 ^ (16 + 0)
    ucsIslStbFiscalPrinting = 2 ^ (16 + 3)
    ucsIslStbNonFiscalPrinting = 2 ^ (16 + 5)
End Enum

Public Enum UcsIslDipSwitchesEnum
    ucsIslDipCenterHeaderFooter = 2 ^ 0
    ucsIslDipPrintHeader = 2 ^ 1
    ucsIslDipSwitch_1_3 = 2 ^ 2
    ucsIslDipSwitch_1_4 = 2 ^ 3
    ucsIslDipZeroMemory = 2 ^ 4
    ucsIslDipTransparentDisplay = 2 ^ 5
    ucsIslDipShowToDisplay = 2 ^ 6
End Enum

Public Enum UcsIslFiscalMemoryEnum
    ucsIslFmeWriteError = 2 ^ 0
    ucsIslFmeNoMemoryModule = 2 ^ 2
    ucsIslFmeLowMemory = 2 ^ 3
    ucsIslFmeOutOfMemory = 2 ^ 4
    ucsIslFmeMemoryError = 2 ^ 5
    ucsIslFmeReadOnly = 2 ^ (8 + 0)
    ucsIslFmeFormatted = 2 ^ (8 + 1)
    ucsIslFmeFiscalMode = 2 ^ (8 + 3)
    ucsIslFmeTaxRatesSet = 2 ^ (8 + 4)
    ucsIslFmeIdentNoSet = 2 ^ (8 + 5)
End Enum

Public Enum UcsIslCommandsEnum
    ucsIslCmdNoop = -1
    '--- init
    ucsIslCmdInitHeaderFooter = 43
    ucsIslCmdInitDateTime = 61
    ucsIslCmdInitInvoiceNo = 66
    ucsIslCmdInitDecimals = 83
    ucsIslCmdInitPaymentType = 85
    ucsIslCmdInitDepartment = 87
    ucsIslCmdInitTaxRates = 96
    ucsIslCmdInitBulstat = 98
    ucsIslCmdInitOperatorPassword = 101
    ucsIslCmdInitOperatorName = 102
    ucsIslCmdInitOperatorReset = 104
    ucsIslCmdInitItem = 107
    ucsIslCmdInitLogo = 115
    ucsIslCmdInitOpenDrawer = 106
    ucsIslCmdInitEcTape = 119
    ucsIslCmdExtendedInitText = 149                '-- command above &H7F
    ucsIslCmdExtendedInitSetting = 150             '-- command above &H7F
    ucsIslCmdExtendedInitPayment = 151             '-- command above &H7F
    ucsIslCmdExtendedInitInvoiceNo = 180           '-- command above &H7F
    ucsIslCmdExtendedInitEJ = 195                  '-- command above &H7F
    '--- transaction
    ucsIslCmdNonFiscalOpen = 38
    ucsIslCmdNonFiscalClose = 39
    ucsIslCmdNonFiscalText = 42
    ucsIslCmdLineFeed = 44
    ucsIslCmdFiscalReversal = 46
    ucsIslCmdFiscalOpen = 48
    ucsIslCmdFiscalSell = 49
    ucsIslCmdFiscalSubtotal = 51
    ucsIslCmdFiscalSellAndDisplay = 52
    ucsIslCmdFiscalPayment = 53
    ucsIslCmdFiscalText = 54
    ucsIslCmdFiscalNoPrintOpen = 55
    ucsIslCmdFiscalClose = 56
    ucsIslCmdFiscalCgInfo = 57
    ucsIslCmdFiscalCancel = 60
    ucsIslCmdFiscalServiceDeposit = 70
    ucsIslCmdExtendedFiscalCancel = 130            '-- command above &H7F
    ucsIslCmdExtendedFiscalDepartmentSell = 138    '-- command above &H7F
    ucsIslCmdExtendedFiscalOpen = 144              '-- command above &H7F
    '--- info
    ucsIslCmdInfoServiceContract = 34
    ucsIslCmdInfoNap = 37
    ucsIslCmdInfoDateTime = 62
    ucsIslCmdInfoLastFiscalBlock = 64
    ucsIslCmdInfoTotalsByTaxGroups = 65
    ucsIslCmdInfoDailyStats = 67
    ucsIslCmdInfoFreeMemoryRecords = 68
    ucsIslCmdInfoStatus = 74
    ucsIslCmdInfoTransaction = 76
    ucsIslCmdInfoLastFiscalDateTime = 86
    ucsIslCmdInfoDepartment = 88
    ucsIslCmdInfoDiagnostics = 90
    ucsIslCmdInfoTaxRates = 97
    ucsIslCmdInfoBulstat = 99
    ucsIslCmdInfoSums = 103
    ucsIslCmdInfoTotalsByPayments = 110
    ucsIslCmdInfoOperator = 112
    ucsIslCmdInfoLastReceipt = 113
    ucsIslCmdInfoFiscalRecord = 114
    ucsIslCmdInfoQRCode = 116
    ucsIslCmdExtendedInfoConsts = 128              '-- command above &H7F
    ucsIslCmdExtendedInfoDepartment = 131          '-- command above &H7F
    ucsIslCmdExtendedInfoTotalsByPayments = 173    '-- command above &H7F
    '--- print
    ucsIslCmdPrintReportTaxRates = 50
    ucsIslCmdPrintDailyReport = 69
    ucsIslCmdPrintDiagnostics = 71
    ucsIslCmdPrintReportByNumberDetailed = 73
    ucsIslCmdPrintReportByDateShort = 79
    ucsIslCmdPrintBarcode = 84
    ucsIslCmdPrintSeparator = 92
    ucsIslCmdPrintReportByDateDetailed = 94
    ucsIslCmdPrintReportByNumberShort = 95
    ucsIslCmdPrintReportByOperators = 105
    ucsIslCmdPrintDailyReportItems = 108
    ucsIslCmdPrintDuplicate = 109
    ucsIslCmdPrintReportByItems = 111
    ucsIslCmdPrintDailyReportDepartments = 117
    ucsIslCmdPrintDailyReportItemsDepartments = 118
    '--- paper
    ucsIslCmdPaperLinefeed = 44
    ucsIslCmdPaperCut = 45
    '--- display
    ucsIslCmdDisplayClear = 33
    ucsIslCmdDisplaySecondLine = 35
    ucsIslCmdDisplayFirstLine = 47
    ucsIslCmdDisplayClock = 63
    ucsIslCmdDisplayText = 100
    '--- misc
    ucsIslCmdMiscRingSound = 80
    '--- ECR
    ucsIslCmdEcrReadRow = &HFF84&
    ucsIslCmdEcrWriteRow = &HFF86&
End Enum

Public Enum UcsIslConfigIndexEnum
    ucsIslCfgDatecsModelA
    ucsIslCfgDatecsModelB
    ucsIslCfgDaisy
    ucsIslCfgIncotex
    ucsIslCfgEltrade
    [_ucsIslCfgMax]
End Enum

'=========================================================================
' API
'=========================================================================

'--- error codes
Private Const ERROR_IO_PENDING              As Long = 997
'--- for WaitCommEvent
Private Const EV_RXCHAR                     As Long = &H1
Private Const EV_ERR                        As Long = &H80

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function SetCommMask Lib "kernel32" (ByVal hFile As Long, ByVal dwEvtMask As Long) As Long
Private Declare Function WaitCommEvent Lib "kernel32" (ByVal hFile As Long, lpEvtMask As Long, lpOverlapped As Any) As Long
Private Declare Function CreateEvent Lib "kernel32" Alias "CreateEventA" (ByVal lpEventAttributes As Long, ByVal bManualReset As Long, ByVal bInitialState As Long, ByVal lpName As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function WaitForMultipleObjects Lib "kernel32" (ByVal nCount As Long, lpHandles As Long, ByVal bWaitAll As Long, ByVal dwMilliseconds As Long) As Long

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

Private Const STR_INTERNAL              As String = "Timeout waiting for response|Invalid response length: %1|Invalid response checksum: %1|Response missing postamble: %1|Response missing preamble: %1|No receipt started| or invalid operator password|Error in line %1: %2|Cannot cancel previous receipt|Invalid report type|Error %1|Memory %1|Unsupported payment type %1|Barcode type %1 not supported|Cannot print|Wrong invoice number %1 generated|Unsupported discount type %1"
Private Const STR_DIP_SWITCHES          As String = "Center header and footer|Enable header|Sw1.3|Sw1.4|Zero memory|Transparent display|No display|7"
Private Const STR_RECEIPT_TEXTS         As String = "CASH|CARD|CHEQUE|CARD|SURCHARGE %1|DISCOUNT %1|SUBTOTAL|Sales %1|EIC"
Private Const STR_STATUSES_A            As String = "Syntax error in the received data|Invalid command code received|The clock is not set|No customer display is connected|Printing unit fault|5|The printer cover is open|7|The command resulted in an overflow of some amount fields|The command is not allowed in the current fiscal mode|The RAM has been reset|Low battery (the real-time clock is in RESET status)|A refund (storno) receipt is open|A service receipt with 90-degree rotated text printing is open|The built-in tax terminal is not responding|7|No paper|Low paper|End of the EJ (less than 1 MB free in the EJ)|A fiscal receipt is open|The end of the EJ is near (less than 10 MB free in the EJ)|A service receipt is open|The end of the EJ is very near (only certain receipts are allowed)"
Private Const STR_ERRORS_A              As String = "Syntax error in the received data|Invalid command code received|2|3|Printing unit fault|5|6|7|0|The command is not allowed in the current fiscal mode|The RAM has been reset|Low battery (the real-time clock is in RESET status)|4|5|6|7|No paper"
Private Const STR_MEMORY_A              As String = "Fiscal memory store error|BULSTAT UIC is set|Unique Printer ID and Fiscal Memory ID are set|There is space for less than 50 records remaining in the FM|The fiscal memory is full|FM error|The printing head is overheated|7|The fiscal memory is set in READONLY mode (locked)|The fiscal memory is formatted|The last fiscal memory store operation is not successful|The printer is in fiscal mode|The tax rates are set at least once|Fiscal memory read error"
Private Const STR_STATUSES_B            As String = "Syntax error in the received data|Invalid command code received|The clock is not set|No customer display is connected|4|5|6|7|The command resulted in an overflow of some amount fields|The command is not allowed in the current fiscal mode|2|3|4|Unsend documents for more than duration set|The built-in tax terminal is not responding|7|No paper|Low paper|End of the EJ (less than 1 MB free in the EJ)|A fiscal receipt is open|The end of the EJ is near (less than 10 MB free in the EJ)|A service receipt is open"
Private Const STR_ERRORS_B              As String = "Syntax error in the received data|Invalid command code received|2|3|4|5|6|7|0|The command is not allowed in the current fiscal mode|2|3|4|5|6|7|No paper"
Private Const STR_MEMORY_B              As String = "Fiscal memory store error|BULSTAT UIC is set|Unique Printer ID and Fiscal Memory ID are set|There is space for less than 50 records remaining in the FM|The fiscal memory is full|FM error|6|7|0|FM is formatted|2|Device is fiscalized|VAT are set at least once"
Private Const STR_STATUSES_DAISY        As String = "Syntax error|Invalid command|Date/time not set|No customer display is connected|Error in printer|5|6|7|The command resulted in an overflow of some amount fields|The command is not allowed in the current fiscal mode|RAM reset|3|4|Cutter error|Wrong password|7|No paper|Low paper|End of the EJ|A fiscal receipt is open|The end of the EJ is near|A service receipt is open|Printing enabled"
Private Const STR_ERRORS_DAISY          As String = "Syntax error|Invalid command|2|3|Error in printer|5|6|7|0|The command is not allowed in the current fiscal mode|RAM reset|3|4|5|6|7|No paper"
Private Const STR_MEMORY_DAISY          As String = "Fiscal memory store error|Tax terminal error|Wrong record in FM|There is space for less than 50 records remaining in the FM|The fiscal memory is full|5|6|7|FM full|1|2|Device is fiscalized|VAT are set at least once|Device and FM serial numbers set|FM ready"
Private Const STR_EXTENDED_DAISY        As String = "1|Overflow after operation|3|No more sales in this bill|4|No more payments in this bill|5|Attempt to perform zero transaction|6|Attempt to perform a sale, after payment is initiated|7|Operation not allowed|8|TAX group with disabled sales|9|Invoice No range not set|11|More than one decimal point|12|More than one character '+' or '-'|13|Symbol '+' or '-' is not in the first position|14|Illegal symbol, e.g. barcode contains not only numbers|15|Extra numbers after decimal point|16|Extra characters|20|With current settings cannot process this command from PC|21|Value outside limits|22|Review system parameter 8|23|Attempt to ""deep"" void after discount/surcharge on subtotal|24|Attempt to ""deep"" void of non-existent transaction|25|Attempt to make a payment without having sales|26|Attempt to sell an item with quantity exceeding its stock|41|Barcode is incorrect (wrong checksum)|" & _
                                                    "42|Attempt to sell without barcode|43|Attempt to program with weight barcode|44|Attempt to sell without programmed barcode|45|Attempt to program already existing barcode|66|Incorrect password|71|!!! Incorrect data in fiscal memory !!!!|72|!!! Error saving in fiscal memory !!!!|76|Information needed from NRA servers|90|Needs a reset with monthly report|91|Needs a reset with daily z-report|92|Needs a reset with report by operators|93|Needs a reset with report by articles|94|Can not reprogram this field|81|Daily z-report overflow|83|Report by operators overflow|84|Report by articles overflow|84|Monthly report overflow|88|Electronic journal overflow|102|No communication between fiscal unit and tax terminal|104|Incorrect communication between fiscal unit and tax terminal|110|SIM card replaced|111|Error in communication between tax terminal and NRA servers|" & _
                                                    "113|NRA server rejects data sent|117|Failed to connect tax terminal with mobile operator network|118|The operation is denied|119|Incorrect value|120|Missing value"
Private Const STR_STATUSES_INCO         As String = "Syntax error in the received data|Invalid command code received|The clock is not set|3|4|5|Printer cover is open|7|0|Command not allowed|FM is reset|Invoices range not set|4|Wrong password 3 times|6|7|End of paper|1|2|A fiscal receipt is open|4|A service receipt is open"
Private Const STR_ERRORS_INCO           As String = "Syntax error in the received data|Invalid command code received|2|3|4|5|Printer cover is open|7|0|Command not allowed|FM is reset|Invoices range not set|4|5|6|7|End of paper"
Private Const STR_MEMORY_INCO           As String = "Fiscal memory store error|1|Wrong record|There is space for less than 50 records remaining in the FM|The fiscal memory is full|FM error|6|7|FM is in READONLY mode|1|2|Device is fiscalized|VAT are set at least once|Unique Printer ID and Fiscal Memory ID are set"
Private Const STR_EXTENDED_INCO         As String = "1|Multiplication overflow|2|Daily report overflow|3|Invalid operation|4|The amount of discount if more than subtotal|5|No sum|6|Fiscal receipt overflow|7|Quantity overflow|8|Price or total overflow|9|No value|10|Custom prices are not allowed for this department/item|11|Fraction quantity is not allowed|12|Maximum number of operation for single fiscal receipt reached. Current receipt has to be committed|13|The percent value is outside 0.00 – 99.99 range|14|The operation is not allow as it leads to a negative result|15|Error accessing item|16|Available quantity is less than entered|18|There are no rows in the fiscal receipt|19|Period report overflow|21|Mot enough available sum|22|Operators report overflow|24|Fiscal memory overflow|25|Invalid range for fiscal memory report|26|Items/groups report overflow|28|Cannot print duplicate|29|Invoices will be available after fiscalization|" & _
                                                    "32|Barcode already used by another item|33|Item/group33|Item/group report required|34|Groups report overflow|36|Insufficient availability|37|Invalid (inactive) tax group|42|Payment type is disabled or does not allow change|43|Invoices range depleted|73|Battery error|75|Total/price overflow programed limit (parameter 2- field 9; parameter 4 – field 6 and 7 for EFSA)|76|Real-time clock error|77|EJ media communication error|78|EJ overflow|79|EJ media not found"
Private Const DEF_SEP                   As String = ","
Private Const DEF_AUTODETECTTIMEOUT     As Long = 51
Private Const DEF_DEVICEINFOTIMEOUT     As Long = 100
Private Const DEF_AUTODETECTSPEEDS      As String = "115200|38400|9600|19200|57600"
Private Const DEF_CONNECTIONTIMEOUT     As Long = 3000
Private Const DEF_TIMEOUT               As Long = 15000
Private Const DEF_OPERNO                As String = "1"
Private Const DEF_PASSWORD              As String = "0000"
Private Const DEF_SOCKET_PORT           As Long = 4999
Private Const DEF_ROW_CHARS             As Long = 42
Private Const DEF_MIN_DISCOUNT          As Double = -99
Private Const DEF_MAX_DISCOUNT          As Double = 99
Private Const DEF_MAX_RECEIPT_ROWS      As Long = 120
Private Const DEF_MAX_ITEM_LINES        As Long = 2
Private Const DEF_CODEPAGE              As Long = 1251
Private Const DEF_MAX_CGINFO            As String = "14|24|25|25|14|35|35"  '-- see UcsMaxCgInfoEnum below
Private Const DEF_MAX_OPERATOR_NAME     As Long = 24
Private Const DEF_SEQUENCE              As Long = &H20
Private Const LNG_MAX_SEQUENCE          As Long = &H7F
Private Const FORMAT_DATE_EXTENDED      As String = "dd\-mm\-yy hh\:nn\:ss"
Private Const FORMAT_DATE_SHORT         As String = "ddmmyyhhnnss"
Private Const FORMAT_DATE_REPORT        As String = "ddmmyy"
Private Const FORMAT_DATE_XML          As String = "yyyy\-mm\-dd\Thh\:nn\:ss"
Private Const MAX_HEADER_LINES          As Long = 6
Private Const MAX_FOOTER_LINES          As Long = 2

Private m_sDevice                   As String
Private m_oConnector                As IDeviceConnector
Private m_lConnectionTimeout        As Long
Private m_lCommandTimeout           As Long
Private m_lSequence                 As Long
Private m_eStatus                   As UcsIslStatusBitsEnum
Private m_eDip                      As UcsIslDipSwitchesEnum
Private m_eMemory                   As UcsIslFiscalMemoryEnum
Private m_vDeviceNumbers            As Variant
Private m_uPrintData                As UcsProtocolPrintData
Private m_uConfig                   As UcsConfigValues
Private m_uLConfig(0 To [_ucsIslCfgMax] - 1) As UcsLocalizedConfigType

Private Enum UcsInternalErrors
    [__ucsErrTimeoutWaitingForResponse]
    ucsErrInvalidResponseLength
    ucsErrInvalidResponseChecksum
    ucsErrResponseMissingPostamble
    ucsErrResponseMissingPreamble
    ucsErrNoReceiptStarted
    ucsErrInvalidOperatorPassword
    ucsErrErrorInLine
    ucsErrCannotCancelReceipt
    ucsErrInvalidReportType
    ucsErrUnknownError
    ucsErrUnknownMemory
    ucsErrUnsupportedPaymentType
    ucsErrBarcodeNotSupported
    ucsErrCannotPrint
    ucsErrWrongInvoiceNo
    ucsErrUnsupportedDiscType
End Enum

Private Enum UcsReceiptTextsEnum
    ucsTxtPaymentCash
    ucsTxtPaymentCredit
    ucsTxtPaymentCheque
    ucsTxtPaymentCard
    ucsTxtSurcharge
    ucsTxtDiscount
    ucsTxtSubTotal
    ucsTxtPluSales
    ucsTxtTaxNo
End Enum

Private Enum UcsModelGroupEnum
    ucsModFpModelA
    ucsModEcrModelB
    ucsModDaisy
    ucsModIncotex
    ucsModEltrade
End Enum

Private Enum UcsProtocolMarkers
    SYM_PREAMBLE = &H1
    SYM_TERMINATOR = &H3
    SYM_SEPARATOR = &H4
    SYM_POSTAMBLE = &H5
    SYM_NAK = &H15
    SYM_SYN = &H16
    SYM_BASEDATA = &H20
End Enum

Private Enum UcsMaxCgInfoEnum
    ucsMciCgTaxNo
    ucsMciOperatorName
    ucsMciCgPrsReceive
    ucsMciCgName
    ucsMciCgVatNo
    ucsMciCgCity
    ucsMciCgAddress
    [_ucsMciLast]
End Enum

Private Type UcsAutodetect
    Connector           As cSerialPortConnector
    WaitOver            As OVERLAPPED
    EventMask           As Long
End Type

Private Type UcsConfigValues
    DefaultPassword     As String
    CodePage            As Long
    MaxCgInfo(0 To [_ucsMciLast] - 1) As Long
    MaxOperatorName     As Long
    DefPmtTypes         As String
    ConfigCommands      As Collection
    LocalizedCommands   As Collection
    PmtMapping          As Variant
End Type

Private Type UcsLocalizedConfigType
    LocalizedText(0 To [_ucsFscLciMax] - 1) As Variant
End Type

Private Type UcsGetDeviceInfoType
    Protocol            As String
    Model               As String
    Firmware            As String
    DeviceSerialNo      As String
    FiscalMemoryNo      As String
    ModelGroup          As UcsModelGroupEnum
    Error               As String
End Type

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunction As String)
    Debug.Print "Critical error: " & Err.Description & " [" & MODULE_NAME & "." & sFunction & "]"
    Logger.Log vbLogEventTypeError, MODULE_NAME, sFunction & "(" & Erl & ")", Err.Description
End Sub

Private Sub RaiseError(sFunction As String)
    Debug.Print "Critical error: " & Err.Description & " [" & MODULE_NAME & "." & sFunction & "]"
    Logger.Log vbLogEventTypeError, MODULE_NAME, sFunction & "(" & Erl & ")", Err.Description
    Err.Raise Err.Number, MODULE_NAME & "." & sFunction & "(" & Erl & ")" & vbCrLf & Err.Source, Err.Description
End Sub

Private Sub DebugLog(sFunction As String, sText As String)
    Logger.Log vbLogEventTypeDebug, MODULE_NAME, sFunction, sText
End Sub

Private Sub DebugDataDump(sFunction As String, sPrefix As String, sData As String)
    Logger.DataDump MODULE_NAME, sFunction, sPrefix, sData
End Sub

'=========================================================================
' Properties
'=========================================================================

Property Get ConnectionTimeout() As Long
    ConnectionTimeout = m_lConnectionTimeout
End Property

Property Let ConnectionTimeout(ByVal lValue As Long)
    m_lConnectionTimeout = lValue
End Property

Property Get CommandTimeout() As Long
    CommandTimeout = m_lCommandTimeout
End Property

Property Let CommandTimeout(ByVal lValue As Long)
    m_lCommandTimeout = lValue
End Property

Property Get Status(ByVal eValue As UcsIslStatusBitsEnum) As Boolean
    Status = (m_eStatus And eValue) <> 0
End Property

Property Get StatusText() As String
    StatusText = pvGetFlagsText(m_eStatus, pvLocalizedTextArray(ucsFscLciPrinterStatuses))
End Property

Property Get ErrorText() As String
    ErrorText = pvGetFlagsText(m_eStatus, pvLocalizedTextArray(ucsFscLciPrinterErrors))
    If (IsDaisy Or IsIncotex) And (m_eDip And &H7F) <> 0 Then
        ErrorText = IIf(LenB(ErrorText) <> 0, ErrorText & vbCrLf, vbNullString) & _
            Zn(At(pvLocalizedTextArray(ucsFscLciExtendedErrors), m_eDip And &H7F), Printf(pvInternal(ucsErrUnknownError), m_eDip And &H7F))
    End If
End Property

Property Get Dip(ByVal eValue As UcsIslDipSwitchesEnum) As Boolean
    Dip = (m_eDip And eValue) <> 0
End Property

Property Get DipText() As String
    If Not IsDaisy Then
        DipText = pvGetFlagsText(m_eDip, pvLocalizedTextArray(ucsFscLciDipSwitchesTexts))
    End If
End Property

Property Get Memory(ByVal eValue As UcsIslFiscalMemoryEnum) As Boolean
    Memory = (m_eMemory And eValue) <> 0
End Property

Property Get MemoryText() As String
    MemoryText = pvGetFlagsText(m_eMemory, pvLocalizedTextArray(ucsFscLciMemoryTexts))
End Property

Property Get LastErrNo() As UcsFiscalErrorsEnum
    LastErrNo = m_uPrintData.LastErrNo
End Property

Property Get LastError() As String
    LastError = m_uPrintData.LastError
End Property

Property Get Device() As String
    Device = m_sDevice
End Property

Property Get IsConnected() As Boolean
    IsConnected = Not m_oConnector Is Nothing
End Property

Property Get LocalizedText(ByVal eIdx As UcsFiscalLocalizedIndexesEnum, Optional ByVal ConfigIndex As Long) As String
    If IsArray(m_uLConfig(ConfigIndex).LocalizedText(eIdx)) Then
        LocalizedText = Join(m_uLConfig(ConfigIndex).LocalizedText(eIdx), "|")
    End If
End Property

Property Let LocalizedText(ByVal eIdx As UcsFiscalLocalizedIndexesEnum, Optional ByVal ConfigIndex As Long, sValue As String)
    m_uLConfig(ConfigIndex).LocalizedText(eIdx) = SplitOrReindex(sValue, "|")
    Select Case eIdx
    Case ucsFscLciInternalErrors
        m_uPrintData.LocalizedText.ErrNoReceiptStarted = pvInternal(ucsErrNoReceiptStarted)
    Case ucsFscLciReceiptTexts
        m_uPrintData.LocalizedText.TxtSurcharge = pvText(ucsTxtSurcharge)
        m_uPrintData.LocalizedText.TxtDiscount = pvText(ucsTxtDiscount)
        m_uPrintData.LocalizedText.TxtPluSales = pvText(ucsTxtPluSales)
    End Select
End Property

Property Get IsFP() As Boolean
    IsFP = (pvModelGroup = ucsModFpModelA)
End Property

Property Get IsEcr() As Boolean
    IsEcr = (pvModelGroup = ucsModEcrModelB)
End Property

Property Get IsDaisy() As Boolean
    IsDaisy = (pvModelGroup = ucsModDaisy)
End Property

Property Get IsIncotex() As Boolean
    IsIncotex = (pvModelGroup = ucsModIncotex)
End Property

Property Get IsEltrade() As Boolean
    IsEltrade = (pvModelGroup = ucsModEltrade)
End Property

Property Get DefPmtTypes() As String
    Const FUNC_NAME     As String = "DefPmtTypes [get]"
    
    On Error GoTo EH
    If LenB(m_uConfig.DefPmtTypes) = 0 Then
        If IsDaisy Then
            m_uConfig.DefPmtTypes = "PNCDB"
        ElseIf IsIncotex Then
            m_uConfig.DefPmtTypes = "PCNDB"
        ElseIf IsEltrade Then
            m_uConfig.DefPmtTypes = "PNCDIJKLMNQR"
        ElseIf IsEcr Then
            m_uConfig.DefPmtTypes = "PNCDIJ"
        Else
            m_uConfig.DefPmtTypes = "PNCDIJKL  mnopqrs"
        End If
    End If
    DefPmtTypes = m_uConfig.DefPmtTypes
    Exit Property
EH:
    RaiseError FUNC_NAME
End Property

'= private ===============================================================

Private Property Get pvRowChars() As Long
    Const FUNC_NAME     As String = "pvRowChars"
    Const CONF_NAME     As String = "RowChars"
    Dim sDaisyPrefix    As String
    Dim sIncotexPrefix  As String
    Dim sEltradePrefix  As String
    Dim eCmd            As UcsIslCommandsEnum
    Dim sParam          As String
    Dim lIndex          As Long
    Dim lOffset         As Long
    Dim sResult         As String
    Dim vSplit          As Variant
    
    On Error GoTo EH
    If m_uPrintData.Config.RowChars = 0 Then
        pvGetCommandConfig CONF_NAME, "ExtendedInfoConsts", eCmd, ucsIslCmdExtendedInfoConsts, sParam, vbNullString, _
            "Index", lIndex, 8, _
            "DaisyPrefix", sDaisyPrefix, "DY", _
            "IncotexPrefix", sIncotexPrefix, "IN", _
            "EltradePrefix", sEltradePrefix, "ED"
        If Left$(GetDeviceSerialNo(), Len(sDaisyPrefix)) = sDaisyPrefix And LenB(sDaisyPrefix) <> 0 Or _
                Left$(GetDeviceSerialNo(), Len(sIncotexPrefix)) = sIncotexPrefix And LenB(sIncotexPrefix) <> 0 Then
            vSplit = Split(SendCommand(eCmd, sParam, LogName:=FUNC_NAME), DEF_SEP)
            m_uPrintData.Config.RowChars = C_Lng(At(vSplit, lIndex + 0))
            If m_uPrintData.Config.CommentChars = 0 Then
                m_uPrintData.Config.CommentChars = C_Lng(At(vSplit, lIndex + 1))
            End If
            If m_uPrintData.Config.ItemChars = 0 Then
                m_uPrintData.Config.ItemChars = C_Lng(At(vSplit, lIndex + 2))
            End If
            If m_uPrintData.Config.MaxReceiptRows = DEF_MAX_RECEIPT_ROWS And C_Lng(At(vSplit, lIndex + 15)) > 0 Then
                m_uPrintData.Config.MaxReceiptRows = C_Lng(At(vSplit, lIndex + 15))
            End If
        ElseIf Left$(GetDeviceSerialNo(), Len(sEltradePrefix)) = sEltradePrefix And LenB(sEltradePrefix) <> 0 Then
            pvGetCommandConfig CONF_NAME, "InitItem", eCmd, ucsIslCmdInitItem, sParam, "I", _
                "Index", lIndex, 3, _
                "Offset", lOffset, 12
            vSplit = Split(SendCommand(eCmd, sParam, LogName:=FUNC_NAME), DEF_SEP)
            If C_Lng(At(vSplit, lIndex)) <> 0 Then
                m_uPrintData.Config.RowChars = C_Lng(At(vSplit, lIndex)) + lOffset
            End If
        Else
            pvGetCommandConfig CONF_NAME, "InitLogo", eCmd, ucsIslCmdInitLogo, sParam, "R0"
            sResult = SendCommand(eCmd, sParam, LogName:=FUNC_NAME)
            '--- check for paper width above 57mm (432px) e.g. 80mm paper is 576px wide
            Select Case Len(sResult) * 4
'            Case 384
'                m_uPrintData.Config.RowChars = DEF_ROW_CHARS
            Case 400
                m_uPrintData.Config.RowChars = Len(sResult) * 4 / 10
            Case Is > 432
                m_uPrintData.Config.RowChars = Len(sResult) * 4 / 12
            End Select
        End If
        If m_uPrintData.Config.RowChars <= 0 Then
            m_uPrintData.Config.RowChars = DEF_ROW_CHARS
        End If
    End If
    pvRowChars = m_uPrintData.Config.RowChars
    Exit Property
EH:
    RaiseError FUNC_NAME
End Property

Private Property Get pvCommentChars() As Long
    pvCommentChars = m_uPrintData.Config.CommentChars
    If pvCommentChars = 0 Then
        pvCommentChars = pvRowChars - IIf(IsEcr, 6, 2)
    End If
End Property

Private Property Get pvItemChars() As Long
    pvItemChars = m_uPrintData.Config.ItemChars
    If pvItemChars = 0 Then
        pvItemChars = pvRowChars - IIf(IsEcr, 20, 8)
    End If
End Property

Property Get pvLocalizedTextArray(ByVal eIdx As UcsFiscalLocalizedIndexesEnum) As Variant
    Dim lCurrentIdx     As Long
    
    If Not IsArray(m_vDeviceNumbers) Then
        lCurrentIdx = C_Lng(At(Split(C_Str(m_vDeviceNumbers), "|"), 2))
    Else
        lCurrentIdx = C_Lng(At(m_vDeviceNumbers, 6))
    End If
    pvLocalizedTextArray = m_uLConfig(lCurrentIdx).LocalizedText(eIdx)
    '--- fallback to default localization config
    If Not IsArray(pvLocalizedTextArray) And lCurrentIdx <> 0 Then
        pvLocalizedTextArray = m_uLConfig(0).LocalizedText(eIdx)
    End If
End Property

Private Property Get pvModelGroup() As UcsModelGroupEnum
    Const FUNC_NAME     As String = "pvModelGroup"

    On Error GoTo EH
    If m_oConnector Is Nothing Then
        GoTo QH
    End If
    If Not IsArray(m_vDeviceNumbers) Then
        pvModelGroup = C_Lng(Zn(At(Split(C_Str(m_vDeviceNumbers), "|"), 2), -1))
    Else
        pvModelGroup = -1
    End If
    If pvModelGroup < 0 Then
        pvModelGroup = pvRawGetDeviceInfo(m_oConnector, m_lCommandTimeout, m_vDeviceNumbers).ModelGroup
    End If
QH:
    Exit Property
EH:
    RaiseError FUNC_NAME
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function AutodetectDevices(vPorts As Variant) As Variant
    Const FUNC_NAME     As String = "AutodetectDevices"
    Dim lIdx            As Long
    Dim baProbe()       As Byte
    Dim sReplyText      As String
    Dim lTimeout        As Long
    Dim lDeviceInfoTimeout As Long
    Dim vRetVal         As Variant
    Dim vSpeed          As Variant
    Dim aPorts()        As UcsAutodetect
    Dim hEvents()       As Long
    Dim lEventsCount    As Long
    Dim lResult         As Long
    Dim vSpeeds         As Variant
    Dim sDefProbe       As String
    Dim baBuffer()      As Byte
    Dim sBuffer         As String
    
    On Error GoTo EH
    vRetVal = vPorts
    If Not IsArray(vPorts) Then
        GoTo QH
    End If
    If UBound(vPorts) < 0 Then
        GoTo QH
    End If
    sDefProbe = Chr$(SYM_PREAMBLE) & Chr$(&H25) & Chr$(&H23) & Chr$(&H4A) & Chr$(&H58) & Chr$(SYM_POSTAMBLE) & Chr$(&H30) & Chr$(&H30) & Chr$(&H3E) & Chr$(&H3F) & Chr$(SYM_TERMINATOR)
    baProbe = ToAscii(GetConfigValue(STR_PROTOCOL_DATECS_FP, "ProbeText", sDefProbe), m_uConfig.CodePage)
    sReplyText = GetConfigValue(STR_PROTOCOL_DATECS_FP, "ReplyText", Chr$(SYM_PREAMBLE)) ' PREAMBLE (&H1)
    lTimeout = GetConfigNumber(STR_PROTOCOL_DATECS_FP, "AutodetectTimeout", DEF_AUTODETECTTIMEOUT)
    lDeviceInfoTimeout = GetConfigNumber(STR_PROTOCOL_DATECS_FP, "DeviceInfoTimeout", DEF_DEVICEINFOTIMEOUT)
    AssignVariant vSpeeds, GetConfigValue(STR_PROTOCOL_DATECS_FP, "AutodetectSpeeds", Split(DEF_AUTODETECTSPEEDS, "|"))
    If IsObject(vSpeeds) Then
        vSpeeds = vSpeeds.Items()
    End If
    For Each vSpeed In vSpeeds
        ReDim aPorts(0 To UBound(vPorts)) As UcsAutodetect
        ReDim hEvents(0 To UBound(vPorts)) As Long
        lEventsCount = 0
        For lIdx = 0 To UBound(aPorts)
            With aPorts(lIdx)
                If Not IsArray(vRetVal(lIdx)) And LenB(At(vPorts, lIdx)) <> 0 Then
                    Set .Connector = InitDeviceConnector(vRetVal(lIdx) & "," & vSpeed, lTimeout, LocalizedText(ucsFscLciConnectorErrors))
                    If Not .Connector Is Nothing Then
                        .Connector.WriteData baProbe, 1
                        .WaitOver.hEvent = CreateEvent(0, 1, 0, 0)
                        Call SetCommMask(.Connector.hComm, EV_RXCHAR Or EV_ERR)
                        lResult = WaitCommEvent(.Connector.hComm, .EventMask, .WaitOver)
                        If lResult = 0 And Err.LastDllError = ERROR_IO_PENDING Then
                            hEvents(lEventsCount) = .WaitOver.hEvent
                            lEventsCount = lEventsCount + 1
                        End If
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
                If Not .Connector Is Nothing And .EventMask <> 0 Then
                    If .Connector.ReadData(baBuffer, lTimeout) And UBound(baBuffer) >= 0 Then
                        sBuffer = Replace(Replace(StrConv(baBuffer, vbUnicode), Chr$(SYM_NAK), vbNullString), Chr$(SYM_SYN), vbNullString)
                        If LenB(sBuffer) = 0 Then
                            If .Connector.ReadData(baBuffer, lTimeout) And UBound(baBuffer) >= 0 Then
                                sBuffer = Replace(Replace(StrConv(baBuffer, vbUnicode), Chr$(SYM_NAK), vbNullString), Chr$(SYM_SYN), vbNullString)
                            End If
                        End If
                        If Left$(sBuffer, Len(sReplyText)) = sReplyText Then
                            DebugLog FUNC_NAME, "Auto-detecting " & .Connector.Device & " reply 0x" & ToHexDump(sBuffer)
                            With pvRawGetDeviceInfo(.Connector, lDeviceInfoTimeout, Empty)
                                If LenB(.Protocol) <> 0 Then
                                    vRetVal(lIdx) = Array(At(vPorts, lIdx), vSpeed, .Protocol, .Model, .Firmware, .DeviceSerialNo, .FiscalMemoryNo, .ModelGroup)
                                    DebugLog FUNC_NAME, "Auto-detected vSpeed=" & vSpeed & ", lIdx=" & lIdx & ", .Protocol=" & .Protocol & ", .Model=" & .Model & ", .Firmware=" & .Firmware & ", .DeviceSerialNo=" & .DeviceSerialNo & ", FiscalMemoryNo=" & .FiscalMemoryNo & ", ModelGroup=" & .ModelGroup
                                End If
                            End With
                        Else
                            DebugLog FUNC_NAME, "Discarding vSpeed=" & vSpeed & ", lIdx=" & lIdx & ", " & .Connector.Device & " reply 0x" & ToHexDump(sBuffer)
                        End If
                    End If
                    Set .Connector = Nothing
                    If .WaitOver.hEvent <> 0 Then
                        Call CloseHandle(.WaitOver.hEvent)
                    End If
                End If
            End With
        Next
    Next
QH:
    AutodetectDevices = vRetVal
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function Init(sDeviceString As String) As Boolean
    Const FUNC_NAME     As String = "Init"
    Dim oOptions        As Object
    Dim sDeviceSerialNo As String
    Dim lIdx            As Long
    Dim vSplit          As Variant
    
    On Error GoTo EH
    DebugLog FUNC_NAME, "sDeviceString=" & sDeviceString
    pvSetLastError vbNullString
    Set oOptions = ParseDeviceString(sDeviceString)
    m_lConnectionTimeout = C_Lng(JsonItem(oOptions, "ConnectionTimeout"))
    If m_lConnectionTimeout <= 0 Then
        m_lConnectionTimeout = DEF_CONNECTIONTIMEOUT
    End If
    m_lCommandTimeout = C_Lng(JsonItem(oOptions, "Timeout"))
    If m_lCommandTimeout <= 0 Then
        m_lCommandTimeout = DEF_TIMEOUT
    End If
    m_sDevice = ToConnectorDevice(oOptions, DEF_SOCKET_PORT, Me)
    If Not Connect() Then
        GoTo QH
    End If
    sDeviceSerialNo = ToDeviceSerialNo(JsonItem(oOptions, "DeviceSerialNo"))
    If LenB(sDeviceSerialNo) = 0 Then
        With pvRawGetDeviceInfo(m_oConnector, m_lCommandTimeout, m_vDeviceNumbers)
            If LenB(.Protocol) = 0 Then
                pvSetLastError .Error
                GoTo QH
            End If
            JsonItem(oOptions, "DeviceSerialNo") = .DeviceSerialNo
            JsonItem(oOptions, "FiscalMemoryNo") = .FiscalMemoryNo
            JsonItem(oOptions, "ModelGroup") = .ModelGroup
            sDeviceSerialNo = .DeviceSerialNo
        End With
    Else
        m_vDeviceNumbers = sDeviceSerialNo & "|" & Trim$(JsonItem(oOptions, "FiscalMemoryNo")) & "|" & Trim$(JsonItem(oOptions, "ModelGroup"))
    End If
    sDeviceString = ToDeviceString(oOptions)
    '--- load config values
    With m_uPrintData.Config
        .RowChars = GetConfigNumber(sDeviceSerialNo, "RowChars", C_Lng(JsonItem(oOptions, "RowChars")))
        .CommentChars = GetConfigNumber(sDeviceSerialNo, "CommentChars", C_Lng(JsonItem(oOptions, "CommentChars")))
        .ItemChars = GetConfigNumber(sDeviceSerialNo, "ItemChars", C_Lng(JsonItem(oOptions, "ItemChars")))
        .AbsoluteDiscount = C_Bool(GetConfigValue(sDeviceSerialNo, "AbsoluteDiscount", JsonBoolItem(oOptions, "AbsoluteDiscount", True)))
        .MinDiscount = GetConfigNumber(sDeviceSerialNo, "MinDiscount", Zndbl(C_Dbl(JsonItem(oOptions, "MinDiscount")), DEF_MIN_DISCOUNT))
        .MaxDiscount = GetConfigNumber(sDeviceSerialNo, "MaxDiscount", Zndbl(C_Dbl(JsonItem(oOptions, "MaxDiscount")), DEF_MAX_DISCOUNT))
        .MaxReceiptRows = GetConfigNumber(sDeviceSerialNo, "MaxReceiptRows", Znl(C_Lng(JsonItem(oOptions, "MaxReceiptRows")), DEF_MAX_RECEIPT_ROWS))
    End With
    With m_uConfig
        .DefaultPassword = GetConfigValue(sDeviceSerialNo, "DefaultPassword", JsonItem(oOptions, "DefaultPassword"))
        .CodePage = GetConfigNumber(sDeviceSerialNo, "CodePage", Znl(C_Lng(JsonItem(oOptions, "CodePage")), DEF_CODEPAGE))
        vSplit = Split(DEF_MAX_CGINFO, "|")
        For lIdx = LBound(.MaxCgInfo) To UBound(.MaxCgInfo)
            .MaxCgInfo(lIdx) = GetConfigNumber(sDeviceSerialNo, "MaxCgInfo" & lIdx, C_Lng(vSplit(lIdx)))
        Next
        .MaxOperatorName = GetConfigNumber(sDeviceSerialNo, "MaxOperatorName", Znl(C_Lng(JsonItem(oOptions, "MaxOperatorName")), DEF_MAX_OPERATOR_NAME))
        Set .ConfigCommands = GetConfigCollection(sDeviceSerialNo, "Commands")
    End With
    '--- more dependant on m_uConfig being populated
    With m_uPrintData.Config
        .NegativePrices = C_Bool(GetConfigValue(sDeviceSerialNo, "NegativePrices", JsonBoolItem(oOptions, "NegativePrices", Not IsFP And Not IsIncotex)))
        .MaxItemLines = GetConfigNumber(sDeviceSerialNo, "MaxItemLines", Znl(C_Lng(JsonItem(oOptions, "MaxItemLines")), IIf(IsEltrade, 0, DEF_MAX_ITEM_LINES)))
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
    DebugLog FUNC_NAME, "m_sDevice=" & m_sDevice & ", m_lConnectionTimeout=" & m_lConnectionTimeout
    Set m_oConnector = InitDeviceConnector(m_sDevice, m_lConnectionTimeout, LocalizedText(ucsFscLciConnectorErrors), m_uPrintData.LastError)
    If m_oConnector Is Nothing Then
        DebugLog FUNC_NAME, "InitDeviceConnector failed, m_uPrintData.LastError=" & m_uPrintData.LastError
        GoTo QH
    End If
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
    DebugLog FUNC_NAME, "TypeOf m_oConnector=" & TypeName(m_oConnector)
    Set m_oConnector = Nothing
    '--- success
    Disconnect = True
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function SendCommand(ByVal eCmd As UcsIslCommandsEnum, Data As String, Optional LogName As String) As String
    Const FUNC_NAME     As String = "SendCommand"
    Dim sError          As String
    
    On Error GoTo EH
    pvSetLastError vbNullString
    If eCmd = ucsIslCmdNoop Then
        Exit Function
    End If
    SendCommand = pvRawSendCommand(m_oConnector, m_lCommandTimeout, eCmd, Data, LogName, Error:=sError)
    If LenB(sError) <> 0 Then
        pvSetLastError sError
        GoTo QH
    End If
    If Status(ucsIslStbPrintingError) Then
        pvSetLastError Zn(ErrorText, Printf(pvInternal(ucsErrUnknownError), "&H" & Hex$(m_eStatus)))
        GoTo QH
    End If
    If Memory(ucsIslFmeMemoryError) Then
        pvSetLastError Zn(MemoryText, Printf(pvInternal(ucsErrUnknownMemory), "&H" & Hex$(m_eMemory)))
        GoTo QH
    End If
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function PrintDuplicate( _
            Optional DuplicateNo As String, _
            Optional ReceiptNo As String, _
            Optional ReceiptDate As Date) As Boolean
    Const FUNC_NAME     As String = "PrintDuplicate"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsIslCommandsEnum
    Dim sParam          As String
    Dim sPrevReceiptNo  As String
    
    On Error GoTo EH
    If Not CancelReceipt() Then
        pvSetLastError pvInternal(ucsErrCannotCancelReceipt)
        GoTo QH
    End If
    If LenB(DuplicateNo) = 0 Then
        pvGetCommandConfig CONF_NAME, "PrintDuplicate", eCmd, ucsIslCmdPrintDuplicate, sParam, "1"
    Else
        If IsDaisy Then
            sParam = "R21" & DEF_SEP & "%1" & DEF_SEP & "%1"
            pvGetCommandConfig CONF_NAME, "ExtendedInitEJ" & sParam, eCmd, ucsIslCmdExtendedInitEJ, sParam, sParam
        Else
            sParam = "P" & DEF_SEP & "%1"
            pvGetCommandConfig CONF_NAME, "InitEcTape" & sParam, eCmd, ucsIslCmdInitEcTape, sParam, sParam
        End If
        sParam = Printf(sParam, DuplicateNo)
    End If
    sPrevReceiptNo = GetLastReceiptNo()
    ReceiptDate = GetClock()
    If LenB(m_uPrintData.LastError) <> 0 Then
        GoTo QH
    End If
    SendCommand eCmd, sParam, LogName:=FUNC_NAME
    ReceiptNo = GetLastReceiptNo()
    If LenB(ReceiptNo) = 0 Or ReceiptNo = sPrevReceiptNo Then
        If LenB(LastError) = 0 Then
            pvSetLastError pvInternal(ucsErrCannotPrint)
        End If
        GoTo QH
    End If
    '--- success
    PrintDuplicate = True
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function PrintDailyXReport( _
            ByVal ReportType As UcsFiscalReportsTypeEnum, _
            Optional ReceiptNo As String, _
            Optional ReceiptDate As Date) As Boolean
    Const FUNC_NAME     As String = "PrintDailyXReport"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsIslCommandsEnum
    Dim sParam          As String
    Dim sPrevReceiptNo  As String
    
    On Error GoTo EH
    If Not CancelReceipt() Then
        pvSetLastError pvInternal(ucsErrCannotCancelReceipt)
        GoTo QH
    End If
    Select Case ReportType
    Case ucsFscRstDaily
        pvGetCommandConfig CONF_NAME, "PrintDailyReport" & ReportType, eCmd, ucsIslCmdPrintDailyReport, sParam, "2"
    Case ucsFscRstDailyByItems
        pvGetCommandConfig CONF_NAME, "PrintDailyReport" & ReportType, eCmd, ucsIslCmdPrintDailyReportItems, sParam, "2"
    Case ucsFscRstDailyByDepartments
        If IsDaisy Then
            pvGetCommandConfig CONF_NAME, "PrintDailyReport" & ReportType, eCmd, ucsIslCmdPrintDailyReport, sParam, "9"
        Else
            pvGetCommandConfig CONF_NAME, "PrintDailyReport" & ReportType, eCmd, ucsIslCmdPrintDailyReportDepartments, sParam, "2"
        End If
    Case ucsFscRstDailyByItemsAndDepartments
        If IsDaisy Then
            pvGetCommandConfig CONF_NAME, "PrintDailyReport" & ReportType, eCmd, ucsIslCmdPrintDailyReportItems, sParam, "9"
        Else
            pvGetCommandConfig CONF_NAME, "PrintDailyReport" & ReportType, eCmd, ucsIslCmdPrintDailyReportItemsDepartments, sParam, "2"
        End If
    Case ucsFscRstDailyByOperators
        pvGetCommandConfig CONF_NAME, "PrintDailyReport" & ReportType, eCmd, ucsIslCmdPrintReportByOperators, sParam, vbNullString
    Case Else
        pvSetLastError pvInternal(ucsErrInvalidReportType)
        GoTo QH
    End Select
    sPrevReceiptNo = GetLastReceiptNo()
    ReceiptDate = GetClock()
    If LenB(m_uPrintData.LastError) <> 0 Then
        GoTo QH
    End If
    SendCommand eCmd, sParam, LogName:=FUNC_NAME
    ReceiptNo = GetLastReceiptNo()
    If LenB(ReceiptNo) = 0 Or ReceiptNo = sPrevReceiptNo Then
        If LenB(LastError) = 0 Then
            pvSetLastError pvInternal(ucsErrCannotPrint)
        End If
        GoTo QH
    End If
    '--- success
    PrintDailyXReport = True
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function PrintDailyZReport( _
            ByVal ReportType As UcsFiscalReportsTypeEnum, _
            Optional ReceiptNo As String, _
            Optional ReceiptDate As Date) As Boolean
    Const FUNC_NAME     As String = "PrintDailyZReport"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsIslCommandsEnum
    Dim sParam          As String
    Dim sPrevReceiptNo  As String
    
    On Error GoTo EH
    If Not CancelReceipt() Then
        pvSetLastError pvInternal(ucsErrCannotCancelReceipt)
        GoTo QH
    End If
    Select Case ReportType
    Case ucsFscRstDaily
        pvGetCommandConfig CONF_NAME, "PrintDailyReport" & ReportType, eCmd, ucsIslCmdPrintDailyReport, sParam, "0"
    Case ucsFscRstDailyByItems
        pvGetCommandConfig CONF_NAME, "PrintDailyReport" & ReportType, eCmd, ucsIslCmdPrintDailyReportItems, sParam, "0"
    Case ucsFscRstDailyByDepartments
        If IsDaisy Then
            pvGetCommandConfig CONF_NAME, "PrintDailyReport" & ReportType, eCmd, ucsIslCmdPrintDailyReport, sParam, "8"
        Else
            pvGetCommandConfig CONF_NAME, "PrintDailyReport" & ReportType, eCmd, ucsIslCmdPrintDailyReportDepartments, sParam, "0"
        End If
    Case ucsFscRstDailyByItemsAndDepartments
        If IsDaisy Then
            pvGetCommandConfig CONF_NAME, "PrintDailyReport" & ReportType, eCmd, ucsIslCmdPrintDailyReportItems, sParam, "8"
        Else
            pvGetCommandConfig CONF_NAME, "PrintDailyReport" & ReportType, eCmd, ucsIslCmdPrintDailyReportItemsDepartments, sParam, "0"
        End If
    Case ucsFscRstDailyByOperators
        pvGetCommandConfig CONF_NAME, "PrintDailyReport" & ReportType, eCmd, ucsIslCmdPrintReportByOperators, sParam, vbNullString
    Case Else
        pvSetLastError Printf(pvInternal(ucsErrInvalidReportType), ReportType)
        GoTo QH
    End Select
    sPrevReceiptNo = GetLastReceiptNo()
    ReceiptDate = GetClock()
    If LenB(m_uPrintData.LastError) <> 0 Then
        GoTo QH
    End If
    SendCommand eCmd, sParam, LogName:=FUNC_NAME
    ReceiptNo = GetLastReceiptNo()
    If LenB(ReceiptNo) = 0 Or ReceiptNo = sPrevReceiptNo Then
        If LenB(LastError) = 0 Then
            pvSetLastError pvInternal(ucsErrCannotPrint)
        End If
        GoTo QH
    End If
    If IsDaisy And (ReportType = ucsFscRstDailyByOperators Or ReportType = ucsFscRstDaily) Then
        pvGetCommandConfig CONF_NAME, "InitOperatorReset", eCmd, ucsIslCmdInitOperatorReset, sParam, "0" & DEF_SEP & "9999"
        pvRawSendCommand m_oConnector, m_lCommandTimeout, eCmd, sParam, FUNC_NAME
    End If
    '--- success
    PrintDailyZReport = True
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function PrintMonthlyReport( _
            ByVal ReportType As UcsFiscalReportsTypeEnum, _
            ByVal StartDate As Date, _
            ByVal EndDate As Date, _
            Optional ReceiptNo As String, _
            Optional ReceiptDate As Date) As Boolean
    Const FUNC_NAME     As String = "PrintMonthlyReport"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsIslCommandsEnum
    Dim sFormat         As String
    Dim sParam          As String
    Dim sPrevReceiptNo  As String
    
    On Error GoTo EH
    If Not CancelReceipt() Then
        pvSetLastError pvInternal(ucsErrCannotCancelReceipt)
        GoTo QH
    End If
    Select Case ReportType
    Case ucsFscRstMonthlyShort
        pvGetCommandConfig CONF_NAME, "PrintReportByDateShort", eCmd, ucsIslCmdPrintReportByDateShort, sParam, "%1" & DEF_SEP & "%2", _
            "Format", sFormat, FORMAT_DATE_REPORT
    Case ucsFscRstMonthlyDetailed
        pvGetCommandConfig CONF_NAME, "PrintReportByDateDetailed", eCmd, ucsIslCmdPrintReportByDateDetailed, sParam, "%1" & DEF_SEP & "%2", _
            "Format", sFormat, FORMAT_DATE_REPORT
    Case Else
        pvSetLastError Printf(pvInternal(ucsErrInvalidReportType), ReportType)
        GoTo QH
    End Select
    sPrevReceiptNo = GetLastReceiptNo()
    ReceiptDate = GetClock()
    If LenB(m_uPrintData.LastError) <> 0 Then
        GoTo QH
    End If
    sParam = Printf(sParam, Format$(StartDate, sFormat), Format$(EndDate, sFormat))
    SendCommand eCmd, sParam, LogName:=FUNC_NAME
    ReceiptNo = GetLastReceiptNo()
    If LenB(ReceiptNo) = 0 Or ReceiptNo = sPrevReceiptNo Then
        If LenB(LastError) = 0 Then
            pvSetLastError pvInternal(ucsErrCannotPrint)
        End If
        GoTo QH
    End If
    '--- success
    PrintMonthlyReport = True
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function PrintServiceDeposit( _
            OperatorCode As String, _
            ByVal Amount As Double, _
            Optional ReceiptNo As String, _
            Optional ReceiptDate As Date, _
            Optional ReceiptAmount As Double) As Variant
    Const FUNC_NAME     As String = "PrintServiceDeposit"
    Const CONF_NAME     As String = FUNC_NAME
    Dim vSplit          As Variant
    Dim vRetVal(0 To 2) As Double
    Dim lIdx            As Long
    Dim eCmd            As UcsIslCommandsEnum
    Dim sParam          As String
    Dim sFormat         As String
    Dim bPrint          As Boolean
    Dim sPrevReceiptNo  As String
    Dim lIndex          As Long
    
    On Error GoTo EH
    If Not CancelReceipt() Then
        pvSetLastError pvInternal(ucsErrCannotCancelReceipt)
        GoTo QH
    End If
    bPrint = Abs(Amount) > DBL_EPSILON
    If bPrint Then
        sPrevReceiptNo = GetLastReceiptNo()
        ReceiptDate = GetClock()
        If LenB(m_uPrintData.LastError) <> 0 Then
            GoTo QH
        End If
    End If
    sParam = "%1" & IIf(IsIncotex And LenB(OperatorCode) <> 0, DEF_SEP & "%2", vbNullString)
    sFormat = IIf(bPrint Or IsIncotex Or IsEltrade, FORMAT_BASE_2, "\")
    pvGetCommandConfig CONF_NAME, "FiscalServiceDeposit" & -bPrint, eCmd, ucsIslCmdFiscalServiceDeposit, sParam, sParam, _
        "Format", sFormat, sFormat, _
        "Index", lIndex, 1
    sParam = Printf(sParam, SafeFormat(Amount, sFormat), OperatorCode)
    vSplit = Split(SendCommand(eCmd, sParam, LogName:=FUNC_NAME), DEF_SEP)
    If UBound(vSplit) <= 0 Then
        GoTo QH
    End If
    If bPrint Then
        ReceiptNo = GetLastReceiptNo()
        ReceiptAmount = Amount
        If LenB(ReceiptNo) = 0 Or ReceiptNo = sPrevReceiptNo Then
            If LenB(LastError) = 0 Then
                pvSetLastError pvInternal(ucsErrCannotPrint)
            End If
            GoTo QH
        End If
    End If
    For lIdx = 0 To UBound(vRetVal)
        vRetVal(lIdx) = ParseSum(At(vSplit, lIdx + lIndex))
    Next
    If IsEltrade Then
        '--- eltrade idiosyncrasy: negative total on service withdraws
        vRetVal(2) = -vRetVal(2)
    End If
    '--- success
    PrintServiceDeposit = vRetVal
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetClock() As Date
    Const FUNC_NAME     As String = "GetClock"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsIslCommandsEnum
    Dim sParam          As String
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoDateTime", eCmd, ucsIslCmdInfoDateTime, sParam, vbNullString
    GetClock = ParseExtendedDate(SendCommand(eCmd, sParam, LogName:=FUNC_NAME))
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function SetClock(ByVal NewDate As Date) As Boolean
    Const FUNC_NAME     As String = "SetClock"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsIslCommandsEnum
    Dim sParam          As String
    Dim sFormat         As String
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InitDateTime", eCmd, ucsIslCmdInitDateTime, sParam, "%1", _
        "Format", sFormat, FORMAT_DATE_EXTENDED
    sParam = Printf(sParam, Format$(NewDate, sFormat))
    SendCommand eCmd, sParam, LogName:=FUNC_NAME
    '--- success
    SetClock = True
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetLastReceiptNo() As String
    Const FUNC_NAME     As String = "GetLastReceiptNo"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsIslCommandsEnum
    Dim sParam          As String
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoLastReceipt", eCmd, ucsIslCmdInfoLastReceipt, sParam, vbNullString
    GetLastReceiptNo = pvRawSendCommand(m_oConnector, m_lCommandTimeout, eCmd, sParam, FUNC_NAME)
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetLastInvoiceNo() As String
    Const FUNC_NAME     As String = "GetLastInvoiceNo"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsIslCommandsEnum
    Dim sParam          As String
    Dim lIndex          As Long
    Dim lStatus         As Long
    Dim lOffset         As Long
    
    On Error GoTo EH
    lOffset = -1
    If IsFP Then
        '--- datecs FP idiosyncrasy: current invoice number is not incremented until fiscal transaction is closed
        pvGetCommandConfig CONF_NAME, "InfoTransaction", eCmd, ucsIslCmdInfoTransaction, sParam, vbNullString, _
            "Index", lIndex, 0, _
            "Status", lStatus, 0
        If C_Lng(At(Split(pvRawSendCommand(m_oConnector, m_lCommandTimeout, eCmd, sParam, FUNC_NAME), DEF_SEP), lIndex)) <> lStatus Then
            lOffset = 0
        End If
    End If
    pvGetCommandConfig CONF_NAME, "InfoSums", eCmd, ucsIslCmdInfoSums, sParam, vbNullString, _
        "Index", lIndex, 10, _
        "Offset", lOffset, lOffset
    GetLastInvoiceNo = At(Split(pvRawSendCommand(m_oConnector, m_lCommandTimeout, eCmd, sParam, FUNC_NAME), DEF_SEP), lIndex)
    If lOffset <> 0 And IsNumeric(GetLastInvoiceNo) Then
        GetLastInvoiceNo = Val(GetLastInvoiceNo) + lOffset
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetLastQRCodeInfo(Optional LastReceiptDate As Date) As String
    Const FUNC_NAME     As String = "GetLastQRCodeInfo"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsDatecsCommandsEnum
    Dim sParam          As String
    Dim lIndex          As Long
    Dim lDateIndex      As Long
    Dim sDelimiter      As String
    Dim vSplit          As Variant
    
    On Error GoTo EH
    If IsDaisy Then
        pvGetCommandConfig CONF_NAME, "InfoQRCode", eCmd, ucsIslCmdInfoQRCode, sParam, vbNullString, _
            "Index", lIndex, 1, _
            "DateIndex", lDateIndex, 2, _
            "Delimiter", sDelimiter, "*"
        vSplit = Split(SendCommand(eCmd, sParam, LogName:=FUNC_NAME), sDelimiter)
        If UBound(vSplit) <= 0 Then
            GoTo QH
        End If
        LastReceiptDate = ParseExtendedDate(At(vSplit, lDateIndex) & " " & At(vSplit, lDateIndex + 1))
        '--- success
        GetLastQRCodeInfo = RTrim$(At(vSplit, lIndex))
    End If
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetDeviceSerialNo() As String
    Const FUNC_NAME     As String = "GetDeviceSerialNo"
    
    On Error GoTo EH
    If Not IsArray(m_vDeviceNumbers) Then
        GetDeviceSerialNo = At(Split(C_Str(m_vDeviceNumbers), "|"), 0)
    End If
    If LenB(GetDeviceSerialNo) = 0 Then
        GetDeviceSerialNo = pvRawGetDeviceInfo(m_oConnector, m_lCommandTimeout, m_vDeviceNumbers).DeviceSerialNo
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetFiscalMemoryNo() As String
    Const FUNC_NAME     As String = "GetFiscalMemoryNo"
    
    On Error GoTo EH
    If Not IsArray(m_vDeviceNumbers) Then
        GetFiscalMemoryNo = At(Split(C_Str(m_vDeviceNumbers), "|"), 1)
    End If
    If LenB(GetFiscalMemoryNo) = 0 Then
        GetFiscalMemoryNo = pvRawGetDeviceInfo(m_oConnector, m_lCommandTimeout, m_vDeviceNumbers).FiscalMemoryNo
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetDeviceProtocol() As String
    Const FUNC_NAME     As String = "GetDeviceProtocol"
    
    On Error GoTo EH
    GetDeviceProtocol = pvRawGetDeviceInfo(m_oConnector, m_lCommandTimeout, m_vDeviceNumbers).Protocol
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetDeviceModel() As String
    Const FUNC_NAME     As String = "GetDeviceModel"
    
    On Error GoTo EH
    GetDeviceModel = pvRawGetDeviceInfo(m_oConnector, m_lCommandTimeout, m_vDeviceNumbers).Model
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetDeviceFirmware() As String
    Const FUNC_NAME     As String = "GetDeviceFirmware"
    
    On Error GoTo EH
    GetDeviceFirmware = pvRawGetDeviceInfo(m_oConnector, m_lCommandTimeout, m_vDeviceNumbers).Firmware
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetTaxNo() As String
    Const FUNC_NAME     As String = "GetTaxNo"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsIslCommandsEnum
    Dim sParam          As String
    Dim lIndex          As Long
    Dim vSplit          As Variant
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoBulstat", eCmd, ucsIslCmdInfoBulstat, sParam, vbNullString, _
        "Index", lIndex, IIf(IsIncotex, -1, 0)
    vSplit = Split(SendCommand(eCmd, sParam, LogName:=FUNC_NAME), DEF_SEP)
    GetTaxNo = RTrim$(At(vSplit, lIndex))
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetTaxCaption() As String
    Const FUNC_NAME     As String = "GetTaxCaption"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsIslCommandsEnum
    Dim sParam          As String
    Dim lIndex          As Long
    
    On Error GoTo EH
    If IsIncotex Then
        pvGetCommandConfig CONF_NAME, "ExtendedInitText", eCmd, ucsIslCmdExtendedInitText, sParam, "R14", _
            "Index", lIndex, 1
    Else
        pvGetCommandConfig CONF_NAME, "InfoBulstat", eCmd, ucsIslCmdInfoBulstat, sParam, vbNullString, _
            "Index", lIndex, 1
    End If
    GetTaxCaption = RTrim$(At(Split(SendCommand(eCmd, sParam, LogName:=FUNC_NAME), DEF_SEP), lIndex))
    If IsNumeric(GetTaxCaption) Or LenB(GetTaxCaption) = 0 Then
        GetTaxCaption = pvText(ucsTxtTaxNo)
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetHeaderText(ByVal Index As Long) As String
    Const FUNC_NAME     As String = "GetHeaderText"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsIslCommandsEnum
    Dim sParam          As String
    
    On Error GoTo EH
    If Index >= 1 And Index <= MAX_HEADER_LINES Then
        pvGetCommandConfig CONF_NAME, "InitHeaderFooter" & Index, eCmd, ucsIslCmdInitHeaderFooter, sParam, "I" & (Index - 1)
        GetHeaderText = RTrim$(pvStripBold(SendCommand(eCmd, sParam, LogName:=FUNC_NAME)))
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetFooterText(ByVal Index As Long) As String
    Const FUNC_NAME     As String = "GetFooterText"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsIslCommandsEnum
    Dim sParam          As String
    
    On Error GoTo EH
    If Index >= 1 And Index <= MAX_FOOTER_LINES Then
        pvGetCommandConfig CONF_NAME, "InitHeaderFooter" & Index, eCmd, ucsIslCmdInitHeaderFooter, sParam, "I" & (Index + 5)
        GetFooterText = RTrim$(pvStripBold(SendCommand(eCmd, sParam, LogName:=FUNC_NAME)))
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetPaymentName(ByVal PmtType As UcsFiscalPaymentTypeEnum) As String
    Const FUNC_NAME     As String = "GetPaymentName"
    
    On Error GoTo EH
    If PmtType >= ucsFscPmtCash And PmtType <= ucsFscPmtCash Then
        GetPaymentName = pvText(ucsTxtPaymentCash + PmtType - 1)
    Else
        GetPaymentName = pvToDevicePmtName(PmtType)
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function OpenDrawer() As Boolean
    Const FUNC_NAME     As String = "OpenDrawer"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsIslCommandsEnum
    Dim sParam          As String
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InitOpenDrawer", eCmd, ucsIslCmdInitOpenDrawer, sParam, vbNullString
    SendCommand eCmd, sParam, LogName:=FUNC_NAME
    '--- success
    OpenDrawer = True
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function CancelReceipt() As Boolean
    Const FUNC_NAME     As String = "CancelReceipt"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eTransCmd       As UcsIslCommandsEnum
    Dim sTransParam     As String
    Dim lTransIndex     As Long
    Dim lTransStatus    As Long
    Dim eCmd            As UcsIslCommandsEnum
    Dim sParam          As String
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoTransaction", eTransCmd, ucsIslCmdInfoTransaction, sTransParam, vbNullString, _
        "Index", lTransIndex, 0, _
        "Status", lTransStatus, 0
    If C_Lng(At(Split(SendCommand(eTransCmd, sTransParam, LogName:=FUNC_NAME), DEF_SEP), lTransIndex, lTransStatus + 1)) = lTransStatus Then
        '--- nothing to do
        CancelReceipt = True
        GoTo QH
    End If
    If Status(ucsIslStbFiscalPrinting) Then
        If IsDaisy Or IsIncotex Then
            pvGetCommandConfig CONF_NAME, "ExtendedFiscalCancel", eCmd, ucsIslCmdExtendedFiscalCancel, sParam, vbNullString
        Else
            pvGetCommandConfig CONF_NAME, "FiscalCancel", eCmd, ucsIslCmdFiscalCancel, sParam, vbNullString
        End If
        SendCommand eCmd, sParam, LogName:=FUNC_NAME
        If Status(ucsIslStbFiscalPrinting) Then
            '--- note: this helps if total sum paid is less than total receipt amount
            pvGetCommandConfig CONF_NAME, "FiscalPayment", eCmd, ucsIslCmdFiscalPayment, sParam, vbTab
            SendCommand eCmd, sParam, LogName:=FUNC_NAME
            '--- note: when printing an extended receipt, cancel fails if no contragent info set
            pvGetCommandConfig CONF_NAME, "FiscalCgInfo", eCmd, ucsIslCmdFiscalCgInfo, sParam, "0000000000" & IIf(IsEcr, vbTab & "0", vbNullString)
            SendCommand eCmd, sParam, LogName:=FUNC_NAME
            pvGetCommandConfig CONF_NAME, "FiscalClose", eCmd, ucsIslCmdFiscalClose, sParam, vbNullString
            SendCommand eCmd, sParam, LogName:=FUNC_NAME
        End If
    Else
        pvGetCommandConfig CONF_NAME, "NonFiscalClose", eCmd, ucsIslCmdNonFiscalClose, sParam, vbNullString
        SendCommand eCmd, sParam, LogName:=FUNC_NAME
    End If
    '--- success (or failure)
    CancelReceipt = C_Lng(At(Split(SendCommand(eTransCmd, sTransParam, LogName:=FUNC_NAME), DEF_SEP), lTransIndex)) = lTransStatus
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetDeviceStatus(StatusText As String) As UcsFiscalStatusCodeEnum
    Const FUNC_NAME     As String = "GetDeviceStatus"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsIslCommandsEnum
    Dim sParam          As String
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoStatus", eCmd, ucsIslCmdInfoStatus, sParam, "X"
    pvRawSendCommand m_oConnector, m_lCommandTimeout, eCmd, sParam, FUNC_NAME, Error:=m_uPrintData.LastError
    If LenB(m_uPrintData.LastError) <> 0 Then
        '--- do nothing
    ElseIf Status(ucsIslStbPrintingError) Then
        pvSetLastError Zn(ErrorText, Printf(pvInternal(ucsErrUnknownError), "&H" & Hex$(m_eStatus)))
    ElseIf Memory(ucsIslFmeMemoryError) Then
        pvSetLastError Zn(MemoryText, Printf(pvInternal(ucsErrUnknownMemory), "&H" & Hex$(m_eMemory)))
    End If
    StatusText = m_uPrintData.LastError
    pvSetLastError vbNullString
    GetDeviceStatus = IIf(LenB(StatusText) <> 0, ucsFscStaFailed, ucsFscStaReady)
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetTotalsByPaymentTypes() As Variant
    Const FUNC_NAME     As String = "GetTotalsByPaymentTypes"
    Const CONF_NAME     As String = FUNC_NAME
    Dim vPmtSum         As Variant
    Dim vRevSum         As Variant
    Dim vDepSum         As Variant
    Dim vRetVal(0 To MAX_PMT_TYPE - MIN_PMT_TYPE) As Variant
    Dim lIdx            As Long
    Dim eCmd            As UcsIslCommandsEnum
    Dim sParam          As String
    Dim lPmtIndex       As Long
    Dim lOffset         As Long
    Dim lRevOffset      As Long
    Dim vTemp           As Variant
    
    On Error GoTo EH
    If IsEcr Then
        pvGetCommandConfig CONF_NAME, "InfoTotalsByPayments0", eCmd, ucsIslCmdInfoTotalsByPayments, sParam, "0", _
            "Offset", lOffset, -10
        vPmtSum = Split(SendCommand(eCmd, sParam, LogName:=FUNC_NAME), DEF_SEP)
        pvGetCommandConfig CONF_NAME, "InfoTotalsByPayments1", eCmd, ucsIslCmdInfoTotalsByPayments, sParam, "1", _
            "RevOffset", lRevOffset, -10
        vRevSum = Split(SendCommand(eCmd, sParam, LogName:=FUNC_NAME), DEF_SEP)
    ElseIf IsIncotex Then
        pvGetCommandConfig CONF_NAME, "InfoTotalsByPayments", eCmd, ucsIslCmdExtendedInfoTotalsByPayments, sParam, vbNullString, _
            "Offset", lOffset, 0
        vPmtSum = Split(SendCommand(eCmd, sParam, LogName:=FUNC_NAME), DEF_SEP)
    Else
        pvGetCommandConfig CONF_NAME, "InfoTotalsByPayments", eCmd, ucsIslCmdInfoTotalsByPayments, sParam, IIf(IsDaisy, "A", IIf(IsEltrade, "11", "*")), _
            "Offset", lOffset, 0
        vPmtSum = Split(SendCommand(eCmd, sParam, LogName:=FUNC_NAME), DEF_SEP)
        If IsFP And UBound(vPmtSum) > 0 Then
            vPmtSum = ConcatArrays(Split(SendCommand(eCmd, vbNullString, LogName:=FUNC_NAME), DEF_SEP), vPmtSum)
            '--- datecs FP idiosyncrasy: cheque and debit cards are swapped
            vTemp = vPmtSum(2)
            vPmtSum(2) = vPmtSum(3)
            vPmtSum(3) = vTemp
        End If
    End If
    If UBound(vPmtSum) <= 0 Then
        GoTo QH
    End If
    For lIdx = 0 To UBound(vRetVal)
        lPmtIndex = pvToDevicePmtIndex(lIdx + 1)
        If lPmtIndex >= 0 Then
            If IsArray(vRevSum) Then
                vRetVal(lIdx) = Array(pvToDevicePmtName(lIdx + 1), ParseSum(At(vPmtSum, lPmtIndex + lOffset)), ParseSum(At(vRevSum, lPmtIndex + lRevOffset)))
            Else
                vRetVal(lIdx) = Array(pvToDevicePmtName(lIdx + 1), ParseSum(At(vPmtSum, lPmtIndex + lOffset)))
            End If
        End If
    Next
    If IsEltrade And IsArray(vRetVal(0)) Then
        '--- eltrade idiosyncrasy: service deposit/withdraw is added to cash total
        pvGetCommandConfig CONF_NAME, "FiscalServiceDeposit", eCmd, ucsIslCmdFiscalServiceDeposit, sParam, "0"
        vDepSum = Split(SendCommand(eCmd, sParam, LogName:=FUNC_NAME), DEF_SEP)
        vRetVal(0)(1) = vRetVal(0)(1) - ParseSum(At(vDepSum, 2)) - ParseSum(At(vDepSum, 3))
    End If
    '--- success
    GetTotalsByPaymentTypes = vRetVal
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetTotalsByTaxGroups() As Variant
    Const FUNC_NAME     As String = "GetTotalsByTaxGroups"
    Const CONF_NAME     As String = FUNC_NAME
    Dim vTaxRates       As Variant
    Dim vPmtSum         As Variant
    Dim vRevSum         As Variant
    Dim vRetVal(0 To MAX_TAX_GROUP - 1) As Variant
    Dim lIdx            As Long
    Dim eCmd            As UcsIslCommandsEnum
    Dim sParam          As String
    Dim lOffset         As Long
    Dim lRevOffset      As Long
    Dim sFormatAmount   As String
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoTaxRates", eCmd, ucsIslCmdInfoTaxRates, sParam, vbNullString
    vTaxRates = Split(Replace(SendCommand(eCmd, sParam, LogName:=FUNC_NAME), "%", vbNullString), DEF_SEP)
    If IsDaisy Then
        pvGetCommandConfig CONF_NAME, "InfoTotalsByTaxGroups", eCmd, ucsIslCmdInfoTotalsByTaxGroups, sParam, "T", _
            "Offset", lOffset, 0, _
            "RevOffset", lRevOffset, 8
        vPmtSum = Split(SendCommand(eCmd, sParam, LogName:=FUNC_NAME), DEF_SEP)
        vRevSum = vPmtSum
    ElseIf IsIncotex Then
        pvGetCommandConfig CONF_NAME, "InfoTotalsByTaxGroups", eCmd, ucsIslCmdInfoTotalsByTaxGroups, sParam, vbNullString, _
            "Offset", lOffset, 1, _
            "RevOffset", lRevOffset, 9, _
            "FormatAmount", sFormatAmount, FORMAT_BASE_2
        vPmtSum = Split(SendCommand(eCmd, sParam, LogName:=FUNC_NAME), DEF_SEP)
        '--- incotex idiosyncrasy: can return netto sums w/o VAT only
        For lIdx = 0 To UBound(vTaxRates)
            If lOffset + lIdx <= UBound(vPmtSum) Then
                vPmtSum(lOffset + lIdx) = Format$(ParseSum(At(vPmtSum, lOffset + lIdx)) * (1# + ParseSum(At(vTaxRates, lIdx)) / 100#), sFormatAmount)
            End If
        Next
        '--- incotex idiosyncrasy: no reversals info
        vRevSum = vPmtSum
    ElseIf IsEltrade Then
        pvGetCommandConfig CONF_NAME, "InfoTotalsByTaxGroups", eCmd, ucsIslCmdInfoTotalsByTaxGroups, sParam, "0", _
            "Offset", lOffset, 0, _
            "RevOffset", lRevOffset, 8
        vPmtSum = Split(SendCommand(eCmd, sParam, LogName:=FUNC_NAME), DEF_SEP)
        '--- eltrade idiosyncrasy: no reversals info
        vRevSum = vPmtSum
    Else
        pvGetCommandConfig CONF_NAME, "InfoTotalsByTaxGroups0", eCmd, ucsIslCmdInfoTotalsByTaxGroups, sParam, "0", _
            "Offset", lOffset, 0
        vPmtSum = Split(SendCommand(eCmd, sParam, LogName:=FUNC_NAME), DEF_SEP)
        pvGetCommandConfig CONF_NAME, "InfoTotalsByTaxGroups5", eCmd, ucsIslCmdInfoTotalsByTaxGroups, sParam, "5", _
            "RevOffset", lRevOffset, 0
        vRevSum = Split(SendCommand(eCmd, sParam, LogName:=FUNC_NAME), DEF_SEP)
    End If
    If UBound(vPmtSum) <= 0 And UBound(vRevSum) <= 0 Then
        GoTo QH
    End If
    For lIdx = 0 To UBound(vRetVal)
        vRetVal(lIdx) = Array(ParseSum(At(vTaxRates, lIdx)), ParseSum(At(vPmtSum, lIdx + lOffset)), ParseSum(At(vRevSum, lIdx + lRevOffset)))
    Next
    '--- success
    GetTotalsByTaxGroups = vRetVal
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetTotalsStats() As Variant
    Const FUNC_NAME     As String = "GetTotalsStats"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsIslCommandsEnum
    Dim sParam          As String
    Dim lIndex          As Long
    Dim vSplit          As Variant
    Dim lReceipts       As Long
    Dim dLastDate       As Date
    Dim lIdx            As Long
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoTotalsByPayments", eCmd, ucsIslCmdInfoTotalsByPayments, sParam, vbNullString, _
        "Index", lIndex, -1
    vSplit = Split(SendCommand(eCmd, sParam, LogName:=FUNC_NAME), DEF_SEP)
    If UBound(vSplit) >= 1 Then
        lReceipts = C_Lng(At(vSplit, lIndex))
    Else
        pvGetCommandConfig CONF_NAME, "InfoOperator", eCmd, ucsIslCmdInfoOperator, sParam, "%1", _
            "Index", lIndex, 0
        lReceipts = 0
        For lIdx = 1 To 100
            vSplit = Split(SendCommand(eCmd, Printf(sParam, lIdx), LogName:=FUNC_NAME), DEF_SEP)
            If UBound(vSplit) <= 0 Then
                Exit For
            End If
            lReceipts = lReceipts + C_Lng(At(vSplit, lIndex))
        Next
    End If
    pvGetCommandConfig CONF_NAME, "InfoLastFiscalDateTime", eCmd, ucsIslCmdInfoLastFiscalDateTime, sParam, "T"
    dLastDate = ParseExtendedDate(SendCommand(eCmd, sParam, LogName:=FUNC_NAME))
    If dLastDate = 0 Then
        pvGetCommandConfig CONF_NAME, "InfoLastFiscalBlock", eCmd, ucsIslCmdInfoLastFiscalBlock, sParam, vbNullString, _
            "Index", lIndex, -1
        vSplit = Split(SendCommand(eCmd, sParam, LogName:=FUNC_NAME), DEF_SEP)
        dLastDate = pvParseDdmmyyDate(At(vSplit, lIndex))
    End If
    GetTotalsStats = Array(lReceipts, IIf(dLastDate = 0, Empty, dLastDate))
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Sub SetLocalizedCommand( _
            sFunction As String, _
            sKey As String, _
            Optional ByVal cmd As UcsIslCommandsEnum, _
            Optional Param As String, _
            Optional More As Variant)
    Const FUNC_NAME     As String = "SetLocalizedCommand"
    Dim sMerged         As String
    Dim lIdx            As Long
    
    On Error GoTo EH
    If m_uConfig.LocalizedCommands Is Nothing Then
        Set m_uConfig.LocalizedCommands = New Collection
    End If
    sMerged = "\" & sFunction & IIf(LenB(sKey) <> 0, "\" & sKey, vbNullString)
    With m_uConfig.LocalizedCommands
        If cmd <> 0 Then
            .Add cmd, sMerged
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
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Public Sub ClearLocalizedCommand(sFunction As String, sKey As String, Optional More As Variant)
    Const FUNC_NAME     As String = "ClearLocalizedCommand"
    Dim sMerged         As String
    Dim lIdx            As Long
    
    On Error GoTo EH
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
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Public Function GetOperatorName(OperatorCode As String) As String
    Const FUNC_NAME     As String = "GetOperatorName"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsIslCommandsEnum
    Dim sParam          As String
    Dim lIndex          As Long
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoOperator", eCmd, ucsIslCmdInfoOperator, sParam, "%1", _
        "Index", lIndex, 5
    sParam = Printf(sParam, OperatorCode)
    GetOperatorName = RTrim$(At(Split(Replace(SendCommand(eCmd, sParam, LogName:=FUNC_NAME), vbTab, DEF_SEP), DEF_SEP), lIndex))
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetDefaultPassword(OperatorCode As String) As String
    If LenB(m_uConfig.DefaultPassword) <> 0 Then
        GetDefaultPassword = m_uConfig.DefaultPassword
    ElseIf IsEcr Or IsDaisy Or IsIncotex Then
        GetDefaultPassword = OperatorCode
    Else
        GetDefaultPassword = DEF_PASSWORD
    End If
End Function

Public Function GetDepartmentName(DepartmentCode As String, Optional TaxGroup As Long) As String
    Const FUNC_NAME     As String = "GetDepartmentName"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsIslCommandsEnum
    Dim sParam          As String
    Dim lIndex          As Long
    Dim lTaxIndex       As Long
    Dim vSplit          As Variant
    
    On Error GoTo EH
    If IsDaisy Then
        pvGetCommandConfig CONF_NAME, "InfoOperator", eCmd, ucsIslCmdExtendedInfoDepartment, sParam, "R%1", _
            "Index", lIndex, 2, _
            "TaxIndex", lTaxIndex, 3
        sParam = Printf(sParam, DepartmentCode)
        vSplit = Split(Replace(SendCommand(eCmd, sParam, LogName:=FUNC_NAME), vbLf, DEF_SEP), DEF_SEP)
        If UBound(vSplit) > 0 Then
            GetDepartmentName = RTrim$(At(vSplit, lIndex))
            TaxGroup = Limit(Asc(At(vSplit, lTaxIndex) & " ") - 191, 0, 8)
        End If
    Else
        pvGetCommandConfig CONF_NAME, "InfoOperator", eCmd, ucsIslCmdInfoDepartment, sParam, "%1", _
            "Index", lIndex, 5, _
            "TaxIndex", lTaxIndex, 0
        sParam = Printf(sParam, DepartmentCode)
        vSplit = Split(Mid$(SendCommand(eCmd, sParam, LogName:=FUNC_NAME), 2), DEF_SEP)
        If UBound(vSplit) > 0 Then
            GetDepartmentName = RTrim$(At(vSplit, lIndex))
            TaxGroup = Limit(Asc(At(vSplit, lTaxIndex) & " ") - 191, 0, 8)
        End If
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetTaxRates() As Variant
    Const FUNC_NAME     As String = "GetTaxRates"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsIslCommandsEnum
    Dim sParam          As String
    Dim vTaxRates       As Variant
    Dim vRetVal(0 To MAX_TAX_GROUP - 1) As Double
    Dim lIdx            As Long
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoTaxRates", eCmd, ucsIslCmdInfoTaxRates, sParam, vbNullString
    vTaxRates = Split(Replace(SendCommand(eCmd, sParam, LogName:=FUNC_NAME), "%", vbNullString), DEF_SEP)
    If UBound(vRetVal) <= 0 Then
        GoTo QH
    End If
    For lIdx = 0 To UBound(vRetVal)
        vRetVal(lIdx) = ParseSum(At(vTaxRates, lIdx, -1))
        If vRetVal(lIdx) >= 100 Then
            vRetVal(lIdx) = -1
        End If
    Next
    '--- success
    GetTaxRates = vRetVal
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

'= private ===============================================================

Private Function pvRawSendCommand( _
            oConnector As IDeviceConnector, _
            ByVal lTimeout As Long, _
            ByVal eCmd As UcsIslCommandsEnum, _
            Data As String, _
            LogName As String, _
            Optional Error As String) As String
    Const FUNC_NAME     As String = "pvRawSendCommand"
    Dim sSend           As String
    Dim sText           As String
    Dim baText()        As Byte
    Dim lIdx            As Long
    Dim lSize           As Long
    Dim sBuffer         As String
    Dim dblTimerEx      As Double
    
    On Error GoTo EH
    DebugLog Zn(LogName, FUNC_NAME), "eCmd=" & eCmd & IIf(LenB(Data) <> 0, ", Data=" & Data, vbNullString)
    '--- prepare command
    sSend = Data
    Do While eCmd >= &H100
        sSend = Chr$(eCmd Mod &H100) & sSend
        eCmd = eCmd \ &H100
    Loop
    sSend = Chr$(SYM_PREAMBLE) & Chr$(SYM_BASEDATA + 4 + Len(sSend)) & Chr$(m_lSequence) & Chr$(eCmd) & sSend & Chr$(SYM_POSTAMBLE)
    '--- increment sequence
    m_lSequence = IIf(m_lSequence >= LNG_MAX_SEQUENCE, DEF_SEQUENCE, m_lSequence + 1)
    '--- append checksum and terminator
    sSend = sSend & pvCheckSum(sSend) & Chr$(SYM_TERMINATOR)
    '--- send request
    dblTimerEx = TimerEx
RetrySend:
    If TimerEx > dblTimerEx + lTimeout / 1000# Then
        Error = pvConnectorError(ucsErrTimeoutWaitingForResponse)
        GoTo QH
    End If
    DebugDataDump Zn(LogName, FUNC_NAME), "[SEND ] ", sSend
    baText = ToAscii(sSend, m_uConfig.CodePage)
    If Not oConnector.WriteData(baText, lTimeout) Then
        Error = oConnector.GetLastError()
        DebugLog Zn(LogName, FUNC_NAME), "WriteData failed, Error=" & Error
        GoTo QH
    End If
    '--- read response
    sText = vbNullString
    Do
        If TimerEx > dblTimerEx + lTimeout / 1000# Then
            Error = pvConnectorError(ucsErrTimeoutWaitingForResponse)
            GoTo QH
        End If
        If Not oConnector.ReadData(baText, lTimeout) Then
            Error = oConnector.GetLastError()
            DebugLog Zn(LogName, FUNC_NAME), "ReadData failed, Error=" & Error
            GoTo QH
        End If
        sBuffer = FromAscii(baText, m_uConfig.CodePage)
        DebugDataDump Zn(LogName, FUNC_NAME), "[RECV ] ", sBuffer
        '--- check for NAK (&H15)
        If InStr(sBuffer, Chr$(SYM_NAK)) > 0 Then
            DebugLog Zn(LogName, FUNC_NAME), "Will retry send on NAK received, sBuffer=0x" & ToHexDump(sBuffer)
            GoTo RetrySend
        End If
        '--- remove SYN (&H16)
        sText = sText & Replace(sBuffer, Chr$(SYM_SYN), vbNullString)
        '--- restart timeout timer
        If InStr(sBuffer, Chr$(SYM_SYN)) > 0 Then
            dblTimerEx = TimerEx
        End If
    Loop While Right$(sText, 1) <> Chr$(SYM_TERMINATOR)
RetryParse:
    '--- response complete w/ TERMINATOR (&H3)
    Debug.Assert Right$(sText, 1) = Chr$(SYM_TERMINATOR)
    If Left$(sText, 1) = Chr$(SYM_PREAMBLE) Then
        lIdx = InStr(sText, Chr$(SYM_POSTAMBLE))
        If lIdx <= 2 Then
            Error = Printf(pvInternal(ucsErrResponseMissingPostamble), "0x" & ToHexDump(sText))
            GoTo QH
        End If
        lSize = Asc(Mid$(sText, 2, 1)) - SYM_BASEDATA
        '--- check message length
        If lSize < 4 Or lSize >= lIdx Then
            Error = Printf(pvInternal(ucsErrInvalidResponseLength), "0x" & ToHexDump(sText))
            GoTo QH
        End If
        '--- check sum
        If pvCheckSum(Mid$(sText, 1, lSize + 1)) <> Mid$(sText, lSize + 2, 4) Then
            Error = Printf(pvInternal(ucsErrInvalidResponseChecksum), "0x" & ToHexDump(sText))
            GoTo QH
        End If
        '--- check for status info
        lIdx = InStrRev(sText, Chr$(SYM_SEPARATOR), lIdx)
        If lIdx > 4 Then
            pvSetStatus Mid$(sText, lIdx + 1, 6)
            lSize = lIdx - 1
        Else
            pvSetStatus vbNullString '--- clear
        End If
        pvRawSendCommand = Mid$(sText, 5, lSize - 4)
        If Status(ucsIslStbPrintingError) Then
            Error = Zn(ErrorText, Printf(pvInternal(ucsErrUnknownError), "&H" & Hex$(m_eStatus)))
        End If
        DebugLog Zn(LogName, FUNC_NAME), "RetVal=" & pvRawSendCommand & _
            IIf(Status(ucsIslStbPrintingError), ", ErrorText=" & Error, vbNullString) & _
            IIf(Memory(ucsIslFmeMemoryError), ", MemoryText=" & Zn(MemoryText, Printf(pvInternal(ucsErrUnknownMemory), "&H" & Hex$(m_eMemory))), vbNullString)
    Else
        '--- check for GSM tax terminal interferences: "AT+CPIN?\r\n" or "AT+CIMI\r\n"
        lIdx = InStr(sText, Chr$(SYM_PREAMBLE))
        If lIdx > 1 Then
            DebugLog Zn(LogName, FUNC_NAME), "Trim bogus symbols, Prefix=0x" & ToHexDump(Left$(sText, lIdx - 1))
            sText = Mid$(sText, lIdx)
            GoTo RetryParse
        End If
        Error = Printf(pvInternal(ucsErrResponseMissingPreamble), "0x" & ToHexDump(sText))
        pvSetStatus vbNullString '--- clear
    End If
QH:
    RaiseEvent CommandComplete(eCmd, Data, pvRawSendCommand)
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function pvCheckSum(sText As String) As String
    Dim baText()        As Byte
    Dim lIdx            As Long
    Dim lSum            As Long
    
    baText = ToAscii(sText, m_uConfig.CodePage)
    For lIdx = 1 To UBound(baText)
        lSum = lSum + baText(lIdx)
    Next
    pvCheckSum = pvToWireLong(lSum)
End Function

Private Function pvToWireLong(ByVal lValue As Long) As String
    pvToWireLong = Chr$((lValue \ &H1000 Mod &H10) + &H30) & Chr$((lValue \ &H100 Mod &H10) + &H30) & Chr$((lValue \ &H10 Mod &H10) + &H30) & Chr$((lValue \ &H1 Mod &H10) + &H30)
End Function

Private Sub pvSetStatus(sStatus As String)
    Const FUNC_NAME     As String = "pvSetStatus"
    Dim baStatus()      As Byte
    Dim lIdx            As Long
    
    On Error GoTo EH
    baStatus = ToAscii(sStatus, m_uConfig.CodePage)
    For lIdx = 0 To UBound(baStatus)
        If (baStatus(lIdx) And &H80) = 0 Then
            baStatus(lIdx) = 0
        End If
    Next
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

Private Function pvGetFlagsText(ByVal lFlags As Long, vTexts As Variant) As String
    Const FUNC_NAME     As String = "pvGetFlagsText"
    Dim lIdx            As Long
    
    On Error GoTo EH
    For lIdx = 0 To UBound(vTexts)
        If (lFlags And (2 ^ lIdx)) <> 0 And Len(vTexts(lIdx)) > 1 Then
            pvGetFlagsText = IIf(LenB(pvGetFlagsText) <> 0, pvGetFlagsText & ", ", vbNullString) & vTexts(lIdx)
        End If
    Next
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Function pvInternal(ByVal lIdx As UcsInternalErrors) As String
    Const FUNC_NAME     As String = "pvInternal"
    
    On Error GoTo EH
    pvInternal = At(pvLocalizedTextArray(ucsFscLciInternalErrors), lIdx)
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
    pvText = At(pvLocalizedTextArray(ucsFscLciReceiptTexts), lIdx)
    If LenB(pvText) = 0 Then
        pvText = At(Split(STR_RECEIPT_TEXTS, "|"), lIdx)
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Function pvConnectorError(ByVal lIdx As UcsConnectorErrorsEnum) As String
    Const FUNC_NAME     As String = "pvConnectorError"
    
    On Error GoTo EH
    pvConnectorError = At(pvLocalizedTextArray(ucsFscLciConnectorErrors), lIdx)
    If LenB(pvConnectorError) = 0 Then
        pvConnectorError = At(Split(STR_CONNECTOR_ERRORS, "|"), lIdx)
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Function pvPrintLine(vText As Variant) As Boolean
    Const FUNC_NAME     As String = "pvPrintLine"
    Const CONF_NAME     As String = "PrintLine"
    Dim vElem               As Variant
    Dim eCmd                As UcsIslCommandsEnum
    Dim sParam              As String
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "FiscalText", eCmd, ucsIslCmdFiscalText, sParam, "%1"
    If IsArray(vText) Then
        For Each vElem In vText
            SendCommand eCmd, Zn(Printf(sParam, C_Str(vElem)), " "), LogName:=FUNC_NAME
            If LenB(m_uPrintData.LastError) <> 0 Then
                GoTo QH
            End If
        Next
    Else
        SendCommand eCmd, Zn(Printf(sParam, C_Str(vText)), " "), LogName:=FUNC_NAME
        If LenB(m_uPrintData.LastError) <> 0 Then
            GoTo QH
        End If
    End If
    '--- success
    pvPrintLine = True
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Function pvPrintReceipt(uCtx As UcsPpdExecuteContext, uRow() As UcsPpdRowData, ByVal lRowCount As Long) As Boolean
    Const FUNC_NAME     As String = "pvPrintReceipt"
    Const CONF_NAME     As String = "PrintReceipt"
    Dim lIdx            As Long
    Dim vSplit          As Variant
    Dim sType           As String
    Dim bInLogin        As Boolean
    Dim sData           As String
    Dim dblTotal        As Double
    Dim dblDisc         As Double
    Dim eCmd            As UcsIslCommandsEnum
    Dim sParam          As String
    Dim lIndex          As Long
    Dim lStatusClosed   As Long
    Dim lMinName        As Long
    Dim sFormatPrice    As String
    Dim sFormatQty      As String
    Dim sFormatDisc     As String
    Dim sFormatAmount   As String
    Dim sFormatDate     As String
    Dim lAmountIndex    As Long
    Dim lPaidIndex      As Long
    Dim sOperParamMax   As String
    Dim sOperParamMin   As String
    Dim lEndIndex       As Long
    Dim sRevType        As String
    Dim sCodeType       As String
    Dim sCodeData       As String
    Dim sUnit           As String
    Dim dLastDate       As Date
    
    On Error GoTo EH
    '--- ring sale
    Do While uCtx.Row < lRowCount
        With uRow(uCtx.Row)
        Select Case .RowType
        Case ucsRowInit
            pvGetCommandConfig CONF_NAME, "InfoOperator", eCmd, ucsIslCmdInfoOperator, sParam, "%1", _
                "Index", lIndex, 5
            sParam = Printf(sParam, .InitOperatorCode)
            vSplit = Split(SendCommand(eCmd, sParam, LogName:=FUNC_NAME), DEF_SEP)
            If UBound(vSplit) > 0 Then
                sData = At(vSplit, lIndex)
                If LenB(.InitOperatorName) <> 0 Then
                    If .InitOperatorName <> sData Then
                        sParam = "%1" & DEF_SEP & "%2" & DEF_SEP & "%3" & IIf(IsDaisy, vbTab & "1,1,1", vbNullString)
                        pvGetCommandConfig CONF_NAME, "InitOperatorName", eCmd, ucsIslCmdInitOperatorName, sParam, sParam, _
                            "MinName", lMinName, 10
                        sOperParamMax = Printf(sParam, .InitOperatorCode, .InitOperatorPassword, Left$(.InitOperatorName, m_uConfig.MaxOperatorName))
                        sOperParamMin = Printf(sParam, .InitOperatorCode, .InitOperatorPassword, Left$(.InitOperatorName, Clamp(Len(sData), lMinName)))
                        SendCommand eCmd, sOperParamMax, LogName:=FUNC_NAME
                        If LenB(m_uPrintData.LastError) <> 0 And sOperParamMin <> sOperParamMax Then
                            SendCommand eCmd, sOperParamMin, LogName:=FUNC_NAME
                            '--- skip last error check
                        End If
                    End If
                Else
                    .InitOperatorName = RTrim$(sData)
                End If
            End If
            '--- for operator errors: if linked receipt in current Z-report -> print service deposit for non-cash payments
            '---                      else -> print service deposit for total receipt amount
            '--- for eltrade: if linked receipt in current Z-report -> print service withdraw for cash payment
            If LenB(.InitRevData(ucsRevReceiptNo)) <> 0 And .InitRevData(ucsRevType) = ucsFscRevOperatorError Then
                If .InitRevData(ucsRevFiscalMemoryNo) = GetFiscalMemoryNo() Then
                    pvGetCommandConfig CONF_NAME, "InfoLastFiscalDateTime", eCmd, ucsIslCmdInfoLastFiscalDateTime, sParam, "T"
                    dLastDate = ParseExtendedDate(SendCommand(eCmd, sParam, LogName:=FUNC_NAME))
                    If dLastDate = 0 Then
                        pvGetCommandConfig CONF_NAME, "InfoLastFiscalBlock", eCmd, ucsIslCmdInfoLastFiscalBlock, sParam, vbNullString, _
                            "Index", lIndex, -1
                        vSplit = Split(SendCommand(eCmd, sParam, LogName:=FUNC_NAME), DEF_SEP)
                        dLastDate = pvParseDdmmyyDate(At(vSplit, lIndex))
                    End If
                Else
                    dLastDate = .InitRevData(ucsRevReceiptDate)
                End If
                If IsEltrade Then
                    If .InitRevData(ucsRevReceiptDate) > dLastDate Then
                        vSplit = PpdGetTotalsByPaymentTypes(uRow, lRowCount)
                        dblTotal = C_Dbl(At(vSplit, ucsFscPmtCash))
                    End If
                    If Abs(dblTotal) > DBL_EPSILON Then
                        PrintServiceDeposit .InitOperatorCode, -dblTotal
                        If LenB(m_uPrintData.LastError) <> 0 Then
                            GoTo QH
                        End If
                    End If
                Else
                    vSplit = PpdGetTotalsByPaymentTypes(uRow, lRowCount)
                    dblTotal = C_Dbl(At(vSplit, 0)) - IIf(.InitRevData(ucsRevReceiptDate) > dLastDate, C_Dbl(At(vSplit, ucsFscPmtCash)), 0)
                    If Abs(dblTotal) > DBL_EPSILON Then
                        PrintServiceDeposit .InitOperatorCode, dblTotal
                        If LenB(m_uPrintData.LastError) <> 0 Then
                            GoTo QH
                        End If
                    End If
                End If
            End If
            '--- setup invoice no
            If LenB(.InitInvData(ucsInvDocNo)) <> 0 Then
                Select Case .InitReceiptType
                Case ucsFscRcpInvoice, ucsFscRcpCreditNote
                    If IsDaisy Then
                        pvGetCommandConfig CONF_NAME, "ExtendedInitSetting", eCmd, ucsIslCmdExtendedInitSetting, sParam, "P%1" & DEF_SEP & "%2", _
                            "StartIndex", lIndex, 19, _
                            "EndIndex", lEndIndex, 20
                        If lIndex > 0 Then
                            SendCommand eCmd, Printf(sParam, lIndex, .InitInvData(ucsInvDocNo)), LogName:=FUNC_NAME
                        End If
                        If lEndIndex > 0 Then
                            SendCommand eCmd, Printf(sParam, lEndIndex, .InitInvData(ucsInvDocNo)), LogName:=FUNC_NAME
                        End If
                    ElseIf IsIncotex Then
                        pvGetCommandConfig CONF_NAME, "ExtendedInitInvoiceNo", eCmd, ucsIslCmdExtendedInitInvoiceNo, sParam, "%1" & vbLf & "%1"
                        sParam = Printf(sParam, .InitInvData(ucsInvDocNo))
                        SendCommand eCmd, sParam, LogName:=FUNC_NAME
                    ElseIf IsEltrade Then
                        sParam = IIf(.InitReceiptType = ucsFscRcpCreditNote, "S", vbNullString) & "%1" & DEF_SEP & "%1"
                        pvGetCommandConfig CONF_NAME, "InitInvoiceNo", eCmd, ucsIslCmdInitInvoiceNo, sParam, sParam
                        sParam = Printf(sParam, .InitInvData(ucsInvDocNo))
                        SendCommand eCmd, sParam, LogName:=FUNC_NAME
                    Else
                        pvGetCommandConfig CONF_NAME, "InitInvoiceNo", eCmd, ucsIslCmdInitInvoiceNo, sParam, "%1" & DEF_SEP & "%1"
                        sParam = Printf(sParam, .InitInvData(ucsInvDocNo))
                        SendCommand eCmd, sParam, LogName:=FUNC_NAME
                    End If
                    If LenB(m_uPrintData.LastError) <> 0 Then
                        GoTo QH
                    End If
                End Select
            End If
            '--- setup fiscal receipt
            If IsDaisy Then
                Select Case .InitReceiptType
                Case ucsFscRcpReversal, ucsFscRcpCreditNote
                    sRevType = Switch( _
                        .InitRevData(ucsRevType) = ucsFscRevTaxBaseReduction, "2", _
                        .InitRevData(ucsRevType) = ucsFscRevOperatorError, "1", _
                        True, "0")
                    '--- %1={ClerkNum},%2={Password},%3={UnicSaleNum}%4=({Refund} | {Credit}{InvLink},)%5={Reason},%6={DocLink},%7={DocLinkDT}{Tab}%8={FiskMem}
                    sParam = "%1" & DEF_SEP & "%2" & DEF_SEP & "%3" & vbTab & "%4%5" & DEF_SEP & "%6" & DEF_SEP & "%7" & vbTab & "%8"
                    If .InitDisablePrinting Then
                        pvGetCommandConfig CONF_NAME, "FiscalNoPrintReversal" & .PrintRowType, eCmd, ucsIslCmdFiscalNoPrintOpen, sParam, sParam, _
                            "FormatDate", sFormatDate, FORMAT_DATE_EXTENDED
                    Else
                        pvGetCommandConfig CONF_NAME, "FiscalReversal" & .PrintRowType, eCmd, ucsIslCmdFiscalOpen, sParam, sParam, _
                            "FormatDate", sFormatDate, FORMAT_DATE_EXTENDED
                    End If
                    sParam = Printf(sParam, .InitOperatorCode, .InitOperatorPassword, _
                        .InitUniqueSaleNo, _
                        IIf(.InitReceiptType = ucsFscRcpReversal, "R", "C" & .InitRevData(ucsRevInvoiceNo) & DEF_SEP), _
                        sRevType, _
                        .InitRevData(ucsRevReceiptNo), _
                        Format$(.InitRevData(ucsRevReceiptDate), sFormatDate), _
                        .InitRevData(ucsRevFiscalMemoryNo))
                Case Else
                    '--- %1={ClerkNum},%2={Password},%3={UnicSaleNum}[{Tab}{Invoice}]
                    sParam = "%1" & DEF_SEP & "%2" & DEF_SEP & "%3" & IIf(.InitReceiptType = ucsFscRcpInvoice, vbTab & "I", vbNullString)
                    If .InitDisablePrinting Then
                        pvGetCommandConfig CONF_NAME, "FiscalNoPrintOpen" & .PrintRowType, eCmd, ucsIslCmdFiscalNoPrintOpen, sParam, sParam
                    Else
                        pvGetCommandConfig CONF_NAME, "FiscalOpen" & .PrintRowType, eCmd, ucsIslCmdFiscalOpen, sParam, sParam
                    End If
                    sParam = Printf(sParam, .InitOperatorCode, .InitOperatorPassword, .InitUniqueSaleNo)
                End Select
            ElseIf IsIncotex Then
                If Not pvPrintCgDetails(uRow(0)) Then
                    GoTo QH
                End If
                '--- %1=<OpNum>,%2=<UNPstring>,%3=<RevDocNo>[,%4=<F1>[%5=(<F2><RevInvoiceNo>,<dd-mm-yy hh:mm:ss>,origDevFMNo)]]
                Select Case .InitReceiptType
                Case ucsFscRcpReversal, ucsFscRcpCreditNote
                    sRevType = Switch( _
                        .InitRevData(ucsRevType) = ucsFscRevTaxBaseReduction, "V", _
                        .InitRevData(ucsRevType) = ucsFscRevOperatorError, "R", _
                        True, "S")
                    sParam = "%1" & DEF_SEP & "%2" & DEF_SEP & "%3" & DEF_SEP & IIf(.InitReceiptType = ucsFscRcpCreditNote, "I", "%4") & "%5"
                Case Else
                    sParam = "%1" & DEF_SEP & "%2" & DEF_SEP & "0" & IIf(.InitReceiptType = ucsFscRcpInvoice, DEF_SEP & "I", vbNullString)
                End Select
                pvGetCommandConfig CONF_NAME, "FiscalOpen" & .PrintRowType, eCmd, ucsIslCmdFiscalOpen, sParam, sParam, _
                    "FormatDate", sFormatDate, FORMAT_DATE_EXTENDED
                sParam = Printf(sParam, .InitOperatorCode, _
                    Zn(.InitUniqueSaleNo, "0"), _
                    .InitRevData(ucsRevReceiptNo), _
                    sRevType, _
                    sRevType & Zn(C_Str(.InitRevData(ucsRevInvoiceNo)), "0") & _
                        DEF_SEP & Format$(.InitRevData(ucsRevReceiptDate), sFormatDate) & _
                        DEF_SEP & .InitRevData(ucsRevFiscalMemoryNo))
            ElseIf IsEltrade Then
                '--- %1=<OperName>,%2=<UNP>[,Type[,%3=(<FMIN>,<Reason>,<num>,<time>)[,%4=<inv>]]]
                Select Case .InitReceiptType
                Case ucsFscRcpReversal, ucsFscRcpCreditNote
                    sRevType = Switch( _
                        .InitRevData(ucsRevType) = ucsFscRevTaxBaseReduction, "T", _
                        .InitRevData(ucsRevType) = ucsFscRevOperatorError, "O", _
                        True, "R")
                    sParam = "%1" & DEF_SEP & "%2" & DEF_SEP & "S" & DEF_SEP & "%3" & IIf(.InitReceiptType = ucsFscRcpCreditNote, DEF_SEP & "%4", vbNullString)
                Case Else
                    sParam = "%1" & DEF_SEP & "%2" & IIf(.InitReceiptType = ucsFscRcpInvoice, DEF_SEP & "I", vbNullString)
                End Select
                pvGetCommandConfig CONF_NAME, "ExtendedFiscalOpen" & .PrintRowType, eCmd, ucsIslCmdExtendedFiscalOpen, sParam, sParam, _
                    "FormatDate", sFormatDate, FORMAT_DATE_XML
                sParam = Printf(sParam, .InitOperatorName, _
                    .InitUniqueSaleNo, _
                    .InitRevData(ucsRevFiscalMemoryNo) & _
                        DEF_SEP & sRevType & _
                        DEF_SEP & .InitRevData(ucsRevReceiptNo) & _
                        DEF_SEP & Format$(.InitRevData(ucsRevReceiptDate), sFormatDate), _
                    .InitRevData(ucsRevInvoiceNo))
            ElseIf IsEcr Then
                Select Case .InitReceiptType
                Case ucsFscRcpReversal, ucsFscRcpCreditNote
                    '--- %1=<OpCode>,%2=<OpPwd>,%3=<NSale>,<TillNmb>,%4=(<DocType>,<DocNumber>,<DocDateTime>,<FMNumber>)[,<Invoice>,%5=(<InvNumber>,<Reason>)]
                    sParam = "%1" & DEF_SEP & "%2" & DEF_SEP & "%3" & DEF_SEP & "1" & DEF_SEP & "%4" & IIf(.InitReceiptType = ucsFscRcpCreditNote, DEF_SEP & "I" & DEF_SEP & "%5", vbNullString)
                    pvGetCommandConfig CONF_NAME, "FiscalReversal" & .PrintRowType, eCmd, ucsIslCmdFiscalReversal, sParam, sParam, _
                        "FormatDate", sFormatDate, FORMAT_DATE_SHORT
                    sParam = Printf(sParam, .InitOperatorCode, .InitOperatorPassword, _
                        .InitUniqueSaleNo, _
                        .InitRevData(ucsRevType) & _
                            DEF_SEP & .InitRevData(ucsRevReceiptNo) & _
                            DEF_SEP & Format$(.InitRevData(ucsRevReceiptDate), sFormatDate) & _
                            DEF_SEP & .InitRevData(ucsRevFiscalMemoryNo), _
                        .InitRevData(ucsRevInvoiceNo) & _
                            DEF_SEP & .InitRevData(ucsRevReason))
                Case Else
                    '--- %1=<OpCode>,%2=<OpPwd>%3=([,<NSale>]),<TillNmb>[,<Invoice>]
                    sParam = "%1" & DEF_SEP & "%2%3" & DEF_SEP & "1" & IIf(.InitReceiptType = ucsFscRcpInvoice, DEF_SEP & "I", vbNullString)
                    pvGetCommandConfig CONF_NAME, "FiscalOpen" & .PrintRowType, eCmd, ucsIslCmdFiscalOpen, sParam, sParam
                    sParam = Printf(sParam, .InitOperatorCode, .InitOperatorPassword, _
                        IIf(LenB(.InitUniqueSaleNo) <> 0, DEF_SEP & .InitUniqueSaleNo, vbNullString))
                End Select
            Else
                Select Case .InitReceiptType
                Case ucsFscRcpReversal, ucsFscRcpCreditNote
                    sRevType = Switch( _
                        .InitRevData(ucsRevType) = ucsFscRevTaxBaseReduction, "T", _
                        .InitRevData(ucsRevType) = ucsFscRevOperatorError, "E", _
                        True, "R")
                    '--- %1=<OpNum>,%2=<Password>,<TillNum>[,<Invoice>%3=<InvNum>][,<UNP>],%4=(<StType><DocNo>)%5=([,<StUNP>,<StDT>,<StFMIN>])[#<RevType>]
                    sParam = "%1" & DEF_SEP & "%2" & DEF_SEP & "1" & IIf(.InitReceiptType = ucsFscRcpCreditNote, DEF_SEP & "I" & DEF_SEP & "%3", vbNullString) & DEF_SEP & "%4%5"
                    pvGetCommandConfig CONF_NAME, "FiscalReversal" & .PrintRowType, eCmd, ucsIslCmdFiscalReversal, sParam, sParam, _
                        "FormatDate", sFormatDate, FORMAT_DATE_SHORT
                    sParam = Printf(sParam, .InitOperatorCode, .InitOperatorPassword, _
                        .InitRevData(ucsRevInvoiceNo), _
                        sRevType & .InitRevData(ucsRevReceiptNo), _
                        IIf(LenB(.InitUniqueSaleNo) <> 0, _
                            DEF_SEP & .InitUniqueSaleNo & _
                            DEF_SEP & Format$(.InitRevData(ucsRevReceiptDate), sFormatDate) & _
                            DEF_SEP & .InitRevData(ucsRevFiscalMemoryNo), vbNullString))
                Case Else
                    '--- %1=<OpNum>,%2=<Password>,<TillNum>[,<Invoice>]%3=[,<UNP>]
                    sParam = "%1" & DEF_SEP & "%2" & DEF_SEP & "1" & IIf(.InitReceiptType = ucsFscRcpInvoice, DEF_SEP & "I", vbNullString) & "%3"
                    pvGetCommandConfig CONF_NAME, "FiscalOpen" & .PrintRowType, eCmd, ucsIslCmdFiscalOpen, sParam, sParam
                    sParam = Printf(sParam, .InitOperatorCode, .InitOperatorPassword, _
                        IIf(LenB(.InitUniqueSaleNo) <> 0, DEF_SEP & .InitUniqueSaleNo, vbNullString))
                End Select
            End If
            bInLogin = Not IsIncotex And Not IsEltrade
            SendCommand eCmd, sParam, LogName:=FUNC_NAME
            If LenB(m_uPrintData.LastError) <> 0 Then
                GoTo QH
            End If
            bInLogin = False
            '--- assert invoice number generated
            Select Case .InitReceiptType
            Case ucsFscRcpInvoice, ucsFscRcpCreditNote
                uCtx.InvoiceNo = GetLastInvoiceNo()
                If LenB(.InitInvData(ucsInvDocNo)) <> 0 And LenB(uCtx.InvoiceNo) <> 0 Then
                    If Val(.InitInvData(ucsInvDocNo)) <> Val(uCtx.InvoiceNo) Then
                        pvSetLastError Printf(pvInternal(ucsErrWrongInvoiceNo), uCtx.InvoiceNo)
                        GoTo QH
                    End If
                End If
            End Select
        Case ucsRowPlu
            uCtx.PluCount = uCtx.PluCount + 1
            dblTotal = Round(.PluPrice * .PluQuantity, DEF_PRICE_SCALE)
            If .DiscType = ucsFscDscPlu Then
                dblDisc = Round(dblTotal * .DiscValue / 100#, DEF_PRICE_SCALE)
            ElseIf .DiscType = ucsFscDscPluAbs Then
                dblDisc = Round(.DiscValue, DEF_PRICE_SCALE)
            Else
                dblDisc = 0
            End If
            uCtx.GrpTotal(.PluTaxGroup) = Round(uCtx.GrpTotal(.PluTaxGroup) + dblTotal + dblDisc, DEF_PRICE_SCALE)
            If .PluDepartmentNo <> 0 And (IsDaisy Or IsIncotex) Then
                '--- Daisy: {[Sign]Dept}{@Price}[*{QTY}][,Percent][$Netto]
                '--- Incotex: <[Sign]Dept>[*<QTY>][,<Percent>][@Price]
                pvGetCommandConfig CONF_NAME, "ExtendedFiscalDepartmentSell" & .PrintRowType, eCmd, ucsIslCmdExtendedFiscalDepartmentSell, sParam, "%1", _
                    "FormatPrice", sFormatPrice, FORMAT_BASE_2, _
                    "FormatQty", sFormatQty, FORMAT_BASE_3, _
                    "FormatDisc", sFormatDisc, FORMAT_BASE_2
                sData = IIf(.PluPrice < 0, "-", vbNullString) & .PluDepartmentNo
                If IsDaisy Then
                    sData = sData & "@" & SafeFormat(Abs(.PluPrice), sFormatPrice)
                End If
                If Abs(.PluQuantity - 1) > DBL_EPSILON Then
                    sData = sData & "*" & SafeFormat(.PluQuantity, sFormatQty)
                End If
                Select Case .DiscType
                Case ucsFscDscPlu
                    If Abs(.PluPrice) > DBL_EPSILON Then
                        sData = sData & DEF_SEP & SafeFormat(.DiscValue, sFormatDisc)
                    End If
                Case ucsFscDscPluAbs
                    sData = sData & IIf(IsDaisy, "$", ";") & SafeFormat(.DiscValue, sFormatPrice)
                End Select
                If IsIncotex Then
                    sData = sData & "@" & SafeFormat(Abs(.PluPrice), sFormatPrice)
                End If
                sParam = Printf(sParam, sData)
            Else
                '--- FP: [<L1>][<Lf><L2>]<Tab><TaxCd><[Sign]Price>[*<Qwan>[#UN]][,Perc|;Abs]
                '--- FP: [<L1>][<Lf><L2>]<Tab><Dept><Tab><[Sign]Price>[*<Qwan>[#UN]][,Perc|;Abs]
                '--- Daisy: [{Text1}][{CR}{Text2}]{Tab}{TaxGr}{[Sign]Price}[*{QTY}][,Percent][$Netto]
                '--- Incotex: [<Text1>][<LF><Text2>]<Tabulator><TaxGr><[Sign]Price>[*<QTY>][,<Percent>|;<Abs>][@MUnitString]
                pvGetCommandConfig CONF_NAME, "FiscalSell" & .PrintRowType, eCmd, ucsIslCmdFiscalSell, sParam, "%1", _
                    "FormatPrice", sFormatPrice, FORMAT_BASE_2, _
                    "FormatQty", sFormatQty, FORMAT_BASE_3, _
                    "FormatDisc", sFormatDisc, FORMAT_BASE_2
                sUnit = IIf(Not IsFP And Not IsIncotex, .PluUnitOfMeasure, vbNullString)
                vSplit = WrapText(.PluItemName & IIf(LenB(sUnit) <> 0, " (" & sUnit & ")", vbNullString), _
                    Clamp(pvCommentChars, , pvItemChars))
                If m_uPrintData.Config.MaxItemLines > 0 Then
                    If UBound(vSplit) >= m_uPrintData.Config.MaxItemLines Then
                        ReDim Preserve vSplit(0 To m_uPrintData.Config.MaxItemLines - 1) As Variant
                    End If
                Else
                    For lIdx = 0 To UBound(vSplit) - 2
                        If Not pvPrintLine(C_Str(vSplit(lIdx))) Then
                            GoTo QH
                        End If
                    Next
                    If UBound(vSplit) >= 2 Then
                        vSplit(0) = vSplit(UBound(vSplit) - 1)
                        vSplit(1) = vSplit(UBound(vSplit))
                        ReDim Preserve vSplit(0 To 1) As Variant
                    End If
                End If
                sData = Join(vSplit, vbLf)
                If .PluDepartmentNo <> 0 Then
                    sData = sData & vbTab & .PluDepartmentNo & vbTab
                Else
                    sData = sData & vbTab & Chr$(191 + .PluTaxGroup)
                End If
                sData = sData & SafeFormat(.PluPrice, sFormatPrice)
                sUnit = IIf(IsFP, .PluUnitOfMeasure, vbNullString)
                If Abs(.PluQuantity - 1) > DBL_EPSILON Or LenB(sUnit) <> 0 Then
                    sData = sData & "*" & SafeFormat(.PluQuantity, sFormatQty) & IIf(LenB(sUnit) <> 0, "#" & sUnit, vbNullString)
                End If
                Select Case .DiscType
                Case ucsFscDscPlu
                    If Abs(.PluPrice) > DBL_EPSILON Then
                        sData = sData & DEF_SEP & SafeFormat(.DiscValue, sFormatDisc)
                    End If
                Case ucsFscDscPluAbs
                    sData = sData & IIf(IsDaisy, "$", ";") & SafeFormat(.DiscValue, sFormatPrice)
                End Select
                sUnit = IIf(IsIncotex, .PluUnitOfMeasure, vbNullString)
                If LenB(sUnit) <> 0 Then
                    sData = sData & "@" & sUnit
                End If
                sParam = Printf(sParam, sData)
            End If
            SendCommand eCmd, sParam, LogName:=FUNC_NAME
            If LenB(m_uPrintData.LastError) <> 0 Then
                GoTo QH
            End If
        Case ucsRowDiscount
            Select Case .PrintRowType
            Case ucsFscRcpNonfiscal
                Select Case .DiscType
                Case ucsFscDscSubtotal
                    pvGetCommandConfig CONF_NAME, "NonFiscalSubtotal" & .DiscType, 0, 0, vbNullString, vbNullString, _
                        "FormatDisc", sFormatDisc, FORMAT_BASE_2, _
                        "FormatAmount", sFormatAmount, FORMAT_BASE_2
                    If Not pvPrintLine(AlignText(pvText(ucsTxtSubTotal), SafeFormat(SumArray(uCtx.GrpTotal), sFormatAmount), pvCommentChars)) Then
                        GoTo QH
                    End If
                    For lIdx = 1 To UBound(uCtx.GrpTotal)
                        If Abs(uCtx.GrpTotal(lIdx)) > DBL_EPSILON Then
                            dblTotal = Round(uCtx.GrpTotal(lIdx) * .DiscValue / 100#, DEF_PRICE_SCALE)
                            If Not pvPrintLine(AlignText(Printf(IIf(dblTotal > DBL_EPSILON, pvText(ucsTxtSurcharge), pvText(ucsTxtDiscount)), "*" & Chr$(191 + lIdx) & "  " & SafeFormat(Abs(.DiscValue), sFormatDisc) & "%"), _
                                    SafeFormat(dblTotal, sFormatAmount), pvCommentChars)) Then
                                GoTo QH
                            End If
                        End If
                    Next
                Case Else
                    pvSetLastError Printf(pvInternal(ucsErrUnsupportedDiscType), .DiscType)
                    GoTo QH
                End Select
            Case Else
                Select Case .DiscType
                Case ucsFscDscSubtotal, ucsFscDscSubtotalAbs
                    sParam = "10" & IIf(.DiscType = ucsFscDscSubtotal, DEF_SEP, IIf(IsDaisy, "$", ";")) & "%1"
                Case Else
                    pvSetLastError Printf(pvInternal(ucsErrUnsupportedDiscType), .DiscType)
                    GoTo QH
                End Select
                pvGetCommandConfig CONF_NAME, "FiscalSubtotal" & .DiscType, eCmd, ucsIslCmdFiscalSubtotal, sParam, sParam, _
                    "FormatDisc", sFormatDisc, FORMAT_BASE_2
                sParam = Printf(sParam, SafeFormat(.DiscValue, sFormatDisc))
                SendCommand eCmd, sParam, LogName:=FUNC_NAME
                If LenB(m_uPrintData.LastError) <> 0 Then
                    GoTo QH
                End If
            End Select
            Select Case .DiscType
            Case ucsFscDscSubtotal
                For lIdx = 1 To UBound(uCtx.GrpTotal)
                    If Abs(uCtx.GrpTotal(lIdx)) > DBL_EPSILON Then
                        dblTotal = Round(uCtx.GrpTotal(lIdx) * .DiscValue / 100#, DEF_PRICE_SCALE)
                        uCtx.GrpTotal(lIdx) = Round(uCtx.GrpTotal(lIdx) + dblTotal, DEF_PRICE_SCALE)
                    End If
                Next
            Case ucsFscDscSubtotalAbs
                '--- ToDo: fix for multiple tax groups
                For lIdx = 1 To UBound(uCtx.GrpTotal)
                    If Abs(uCtx.GrpTotal(lIdx)) > DBL_EPSILON Then
                        uCtx.GrpTotal(lIdx) = Round(uCtx.GrpTotal(lIdx) + .DiscValue, DEF_PRICE_SCALE)
                        Exit For
                    End If
                Next
            End Select
        Case ucsRowLine
            If .LineWordWrap Then
                vSplit = WrapText(.LineText, pvCommentChars)
            Else
                vSplit = Array(Left$(.LineText, pvCommentChars))
            End If
            If Not pvPrintLine(vSplit) Then
                GoTo QH
            End If
        Case ucsRowBarcode
            If .BarcodeHeight > 0 And Not IsDaisy Then
                pvGetCommandConfig CONF_NAME, "InitHeaderFooter", eCmd, ucsIslCmdInitHeaderFooter, sParam, "B" & DEF_SEP & "%1"
                sParam = Printf(sParam, Clamp(.BarcodeHeight * 6, 24, 240))
                SendCommand eCmd, sParam, LogName:=FUNC_NAME
                '--- skip last error check
            End If
            sCodeData = .BarcodeText
            Select Case .BarcodeType
            Case ucsFscBrcEan8
                sCodeType = "1"
                sCodeData = Pad(sCodeData, 7, "0")
            Case ucsFscBrcEan13
                sCodeType = IIf(IsIncotex, "3", "2")
                sCodeData = Pad(sCodeData, 12, "0")
            Case ucsFscBrcCode128
                If IsIncotex Then
                    pvSetLastError Printf(pvInternal(ucsErrBarcodeNotSupported), .BarcodeType)
                    GoTo QH
                End If
                sCodeType = "3"
            Case ucsFscBrcQRcode
                If IsDaisy Or IsIncotex Or IsEcr Then
                    pvSetLastError Printf(pvInternal(ucsErrBarcodeNotSupported), .BarcodeType)
                    GoTo QH
                End If
                sCodeType = IIf(IsEltrade, "6", "Q")
            End Select
            If IsDaisy Then
                sParam = "%1" & DEF_SEP & "%2" & vbTab & "C" & DEF_SEP & "0" & DEF_SEP & "%3" & DEF_SEP & "0"
            ElseIf IsIncotex Then
                sParam = "%1" & DEF_SEP & "%2" & DEF_SEP & "0"
            ElseIf IsEltrade Then
                sParam = "%1" & DEF_SEP & "%2"
            ElseIf IsEcr Then
                sParam = "%1" & ";" & "%2"
            Else
                sParam = "%1" & "C;" & "%2"
            End If
            pvGetCommandConfig CONF_NAME, "PrintBarcode", eCmd, ucsIslCmdPrintBarcode, sParam, sParam
            sParam = Printf(sParam, sCodeType, sCodeData, .BarcodeHeight)
            SendCommand eCmd, sParam, LogName:=FUNC_NAME
            If LenB(m_uPrintData.LastError) <> 0 Then
                GoTo QH
            End If
        Case ucsRowPayment
            pvGetCommandConfig CONF_NAME, "InfoTransaction", eCmd, ucsIslCmdInfoTransaction, sParam, "T", _
                "Index", lIndex, 0, _
                "StatusClosed", lStatusClosed, 0, _
                "AmountIndex", lAmountIndex, 2, _
                "PaidIndex", lPaidIndex, 3
            vSplit = Split(SendCommand(eCmd, sParam, LogName:=FUNC_NAME), DEF_SEP)
            If UBound(vSplit) <= 0 Then
                GoTo QH
            End If
            If C_Lng(At(vSplit, lIndex)) = lStatusClosed Then
                pvSetLastError pvInternal(ucsErrCannotPrint)
                GoTo QH
            End If
            If .PmtType <> 0 Then
                If Abs(.PmtAmount) < DBL_EPSILON Then
                    eCmd = ucsIslCmdNoop
                ElseIf .PrintRowType = ucsFscRcpReversal Or .PrintRowType = ucsFscRcpCreditNote Then
                    eCmd = ucsIslCmdNoop
                Else
                    lIndex = pvToDevicePmtIndex(.PmtType)
                    If lIndex < 0 Then
                        pvSetLastError Printf(pvInternal(ucsErrUnsupportedPaymentType), .PmtType)
                        GoTo QH
                    End If
                    pvGetCommandConfig CONF_NAME, "PaymentTypes", 0, 0, sParam, DefPmtTypes
                    sType = Mid$(sParam, lIndex + 1, 1)
                    pvGetCommandConfig CONF_NAME, "FiscalPayment" & .PmtType, eCmd, ucsIslCmdFiscalPayment, sParam, vbTab & sType & "%1", _
                        "FormatAmount", sFormatAmount, FORMAT_BASE_2
                    sParam = Printf(sParam, SafeFormat(.PmtAmount, sFormatAmount))
                End If
            Else
                '--- check if anything left for payment
                If ParseSum(At(vSplit, lAmountIndex)) > ParseSum(At(vSplit, lPaidIndex)) _
                        Or LenB(At(vSplit, lAmountIndex)) = 0 _
                        Or Not uCtx.PmtPrinted Then
                    pvGetCommandConfig CONF_NAME, "FiscalPayment" & .PmtType, eCmd, ucsIslCmdFiscalPayment, sParam, vbTab
                    SendCommand eCmd, sParam, LogName:=FUNC_NAME
                    If LenB(m_uPrintData.LastError) = 0 Then
                        uCtx.PmtPrinted = True
                    End If
                End If
                If Not IsIncotex Then
                    If Not pvPrintCgDetails(uRow(0)) Then
                        GoTo QH
                    End If
                End If
                uCtx.ReceiptNo = GetLastReceiptNo()
                If LenB(uCtx.ReceiptNo) <> 0 And Not IsEltrade Then
                    uCtx.ReceiptNo = Val(uCtx.ReceiptNo) + 1
                End If
                uCtx.ReceiptDate = GetClock()
                uCtx.ReceiptAmount = ParseSum(At(vSplit, lAmountIndex))
                Select Case .PrintRowType
                Case ucsFscRcpReversal, ucsFscRcpCreditNote
                    uCtx.ReceiptAmount = -uCtx.ReceiptAmount
                End Select
                If uCtx.PmtPrinted Or SumArray(uCtx.GrpTotal) > DBL_EPSILON Then
                    pvGetCommandConfig CONF_NAME, "FiscalClose" & .PrintRowType, eCmd, ucsIslCmdFiscalClose, sParam, vbNullString
                Else
                    pvGetCommandConfig CONF_NAME, "FiscalCancel" & .PrintRowType, eCmd, ucsIslCmdFiscalCancel, sParam, vbNullString
                End If
            End If
            SendCommand eCmd, sParam, LogName:=FUNC_NAME
            If LenB(m_uPrintData.LastError) <> 0 Then
                GoTo QH
            End If
            If eCmd <> ucsIslCmdNoop Then
                uCtx.PmtPrinted = True
            End If
        End Select
        End With
        uCtx.Row = uCtx.Row + 1
    Loop
    '--- success
    pvPrintReceipt = True
QH:
    If bInLogin Then
        pvSetLastError m_uPrintData.LastError & pvInternal(ucsErrInvalidOperatorPassword), ucsFerInvalidPassword
    End If
    If LenB(m_uPrintData.LastError) <> 0 And uCtx.PluCount > 0 And Not uCtx.PmtPrinted Then
        pvSetLastError Printf(pvInternal(ucsErrErrorInLine), uCtx.PluCount, m_uPrintData.LastError), m_uPrintData.LastErrNo
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Function pvPrintCgDetails(uFirstRow As UcsPpdRowData) As Boolean
    Const FUNC_NAME     As String = "pvPrintCgDetails"
    Const CONF_NAME     As String = "PrintCgDetails"
    Dim eCmd            As UcsIslCommandsEnum
    Dim sParam          As String
    Dim sFormatAddr     As String
    Dim sTaxNoSuffix    As String
    Dim sTaxNoPrefix    As String
    Dim lMaxSize        As Long
    
    On Error GoTo EH
    With uFirstRow
        Select Case .InitReceiptType
        Case ucsFscRcpInvoice, ucsFscRcpCreditNote
            If IsDaisy Then
                sParam = "%1#%5#%2#%3#%4#%6"
            ElseIf IsIncotex Then
                sParam = "%5#%2#%3#%4#%1#%6"
            Else
                sParam = "%1#%2#%3#%4#%5#%6"
            End If
            If LenB(.InitInvData(ucsInvCgAddress)) <> 0 Then
                sFormatAddr = "%1" & vbLf & "%2"
            Else
                sFormatAddr = "%1"
            End If
            pvGetCommandConfig CONF_NAME, "FiscalCgInfo" & .PrintRowType, eCmd, ucsIslCmdFiscalCgInfo, sParam, Replace(sParam, "#", vbTab), _
                "FormatAddr", sFormatAddr, sFormatAddr
            sFormatAddr = Printf(sFormatAddr, Left$(.InitInvData(ucsInvCgCity), m_uConfig.MaxCgInfo(ucsMciCgCity)), _
                Left$(.InitInvData(ucsInvCgAddress), m_uConfig.MaxCgInfo(ucsMciCgAddress)))
            If IsEcr Then
                sTaxNoSuffix = vbTab & .InitInvData(ucsInvCgTaxNoType)
            ElseIf IsEltrade Then
                If .InitInvData(ucsInvCgTaxNoType) = ucsFscTxnCitizenNo Then
                    sTaxNoPrefix = "#"
                End If
            ElseIf IsFP Then
                sTaxNoPrefix = At(Split("|#|*|^", "|"), .InitInvData(ucsInvCgTaxNoType))
            End If
            lMaxSize = m_uConfig.MaxCgInfo(ucsMciCgTaxNo) - Len(sTaxNoPrefix) - Len(sTaxNoSuffix)
            sParam = Printf(sParam, sTaxNoPrefix & Left$(.InitInvData(ucsInvCgTaxNo), lMaxSize) & sTaxNoSuffix, _
                Left$(.InitOperatorName, m_uConfig.MaxCgInfo(ucsMciOperatorName)), _
                Left$(.InitInvData(ucsInvCgPrsReceive), m_uConfig.MaxCgInfo(ucsMciCgPrsReceive)), _
                Left$(.InitInvData(ucsInvCgName), m_uConfig.MaxCgInfo(ucsMciCgName)), _
                Left$(.InitInvData(ucsInvCgVatNo), m_uConfig.MaxCgInfo(ucsMciCgVatNo)), _
                sFormatAddr)
            SendCommand eCmd, sParam, LogName:=FUNC_NAME
            If LenB(m_uPrintData.LastError) <> 0 Then
                GoTo QH
            End If
        End Select
    End With
    '--- success
    pvPrintCgDetails = True
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Function pvRawGetDeviceInfo( _
            oConnector As IDeviceConnector, _
            ByVal lTimeout As Long, _
            vResult As Variant) As UcsGetDeviceInfoType
    Const FUNC_NAME     As String = "pvRawGetDeviceInfo"
    Const CONF_NAME     As String = "GetDeviceInfo"
    Dim eCmd            As UcsIslCommandsEnum
    Dim sParam          As String
    Dim sDaisyPrefix    As String
    Dim sIncotexPrefix  As String
    Dim sEltradePrefix  As String
    Dim lNameIndex      As Long
    Dim lSerialIndex    As Long
    Dim vSplit          As Variant
    Dim lEcrTimeout     As Long
    Dim uRetVal         As UcsGetDeviceInfoType
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoDiagnostics", eCmd, ucsIslCmdInfoDiagnostics, sParam, "0", _
        "DaisyPrefix", sDaisyPrefix, "DY", _
        "IncotexPrefix", sIncotexPrefix, "IN", _
        "EltradePrefix", sEltradePrefix, "ED", _
        "NameIndex", lNameIndex, 0, _
        "SerialIndex", lSerialIndex, 4
    If Not IsArray(vResult) Then
        vSplit = Split(pvRawSendCommand(oConnector, lTimeout, eCmd, vbNullString, FUNC_NAME, Error:=uRetVal.Error), DEF_SEP)
        If LenB(At(vSplit, lNameIndex)) = 0 Then
            uRetVal.Error = vbNullString
            vSplit = Split(pvRawSendCommand(oConnector, lTimeout, eCmd, sParam, FUNC_NAME, Error:=uRetVal.Error), DEF_SEP)
        End If
        If UBound(vSplit) >= lSerialIndex Then
            vResult = vSplit
            If Left$(At(vResult, lSerialIndex), Len(sIncotexPrefix)) = sIncotexPrefix And LenB(sIncotexPrefix) <> 0 Then
                vResult(0) = "Incotex " & vResult(0)
            ElseIf Left$(At(vResult, lSerialIndex + 1), Len(sEltradePrefix)) = sEltradePrefix And LenB(sEltradePrefix) <> 0 Then
                pvEltradeMergeInfoArray vResult
            End If
            ReDim Preserve vResult(0 To 6) As String
            If Left$(At(vResult, lSerialIndex), Len(sDaisyPrefix)) = sDaisyPrefix And LenB(sDaisyPrefix) <> 0 Then
                vResult(6) = ucsModDaisy
            ElseIf Left$(At(vResult, lSerialIndex), Len(sIncotexPrefix)) = sIncotexPrefix And LenB(sIncotexPrefix) <> 0 Then
                vResult(6) = ucsModIncotex
            ElseIf Left$(At(vResult, lSerialIndex), Len(sEltradePrefix)) = sEltradePrefix And LenB(sEltradePrefix) <> 0 Then
                vResult(6) = ucsModEltrade
            Else
                pvGetCommandConfig CONF_NAME, "EcrReadRow", eCmd, ucsIslCmdEcrReadRow, sParam, "2;1;", _
                    "Timeout", lEcrTimeout, 100
                If LenB(pvRawSendCommand(oConnector, lEcrTimeout, eCmd, sParam, FUNC_NAME)) <> 0 Then
                    vResult(6) = ucsModEcrModelB
                Else
                    vResult(6) = ucsModFpModelA
                End If
            End If
        End If
    End If
    If LenB(At(vResult, lSerialIndex)) <> 0 Then
        '--- check serial
        If Left$(At(vResult, lSerialIndex), Len(sDaisyPrefix)) = sDaisyPrefix And LenB(sDaisyPrefix) <> 0 Then
            uRetVal.Protocol = STR_PROTOCOL_DAISY
        ElseIf Left$(At(vResult, lSerialIndex), Len(sIncotexPrefix)) = sIncotexPrefix And LenB(sIncotexPrefix) <> 0 Then
            uRetVal.Protocol = STR_PROTOCOL_INCOTEX
        ElseIf Left$(At(vResult, lSerialIndex), Len(sEltradePrefix)) = sEltradePrefix And LenB(sEltradePrefix) <> 0 Then
            uRetVal.Protocol = STR_PROTOCOL_ELTRADE
        ElseIf LenB(At(vResult, lSerialIndex)) <> 0 Then
            uRetVal.Protocol = STR_PROTOCOL_DATECS_FP
        End If
        vSplit = Split(RTrim$(At(vResult, lNameIndex)) & RTrim$(" " & At(vResult, lNameIndex + 1)), Limit:=2)
        uRetVal.Model = RTrim$(At(vSplit, 0))
        uRetVal.Firmware = RTrim$(At(vSplit, 1))
        uRetVal.DeviceSerialNo = RTrim$(At(vResult, lSerialIndex))
        uRetVal.FiscalMemoryNo = RTrim$(At(vResult, lSerialIndex + 1))
        uRetVal.ModelGroup = C_Lng(At(vResult, 6))
    End If
    pvRawGetDeviceInfo = uRetVal
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function pvStripBold(sText As String) As String
    Const FUNC_NAME     As String = "pvStripBold"
    
    On Error GoTo EH
    pvStripBold = preg_replace(GetConfigForCommand(m_uConfig.ConfigCommands, m_uConfig.LocalizedCommands, FUNC_NAME, "Pattern", "/[!^]/i"), _
        sText, GetConfigForCommand(m_uConfig.ConfigCommands, m_uConfig.LocalizedCommands, FUNC_NAME, "Replace", vbNullString))
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Function pvGetPmtMapping() As Object
    Const FUNC_NAME     As String = "pvGetPmtMapping"
    Const CONF_NAME     As String = "GetPmtMapping"
    Dim eCmd            As UcsIslCommandsEnum
    Dim sParam          As String
    Dim lIdx            As Long
    Dim vConfig         As Variant
    Dim lOffset         As Long
    Dim lTagNo          As Long
    Dim vSplit          As Variant
    Dim oRetVal         As Object
    
    On Error GoTo EH
    If IsDaisy Then
        pvGetCommandConfig CONF_NAME, "ExtendedInitPayment", eCmd, ucsIslCmdExtendedInitPayment, sParam, "R%1", _
            "Offset", lOffset, 3
        For lIdx = 0 To 4
            vSplit = Split(Replace(SendCommand(eCmd, Printf(sParam, lIdx), LogName:=FUNC_NAME), vbTab, DEF_SEP), DEF_SEP)
            If lIdx > 0 Then
                If LenB(At(vSplit, lOffset)) = 0 Then
                    GoTo QH
                End If
                lTagNo = C_Lng(At(vSplit, lOffset))
            End If
            If IsEmpty(JsonItem(oRetVal, "#" & lTagNo)) Then
                JsonItem(oRetVal, "#" & lTagNo) = Array(lIdx, At(vSplit, 1))
            End If
        Next
    ElseIf IsEcr Then
        pvGetCommandConfig CONF_NAME, "EcrReadRow", eCmd, ucsIslCmdEcrReadRow, sParam, "21;1;", _
            "Offset", lOffset, 9
        vConfig = Split(SendCommand(eCmd, sParam, LogName:=FUNC_NAME), ";")
        pvGetCommandConfig CONF_NAME, "EcrReadRow", eCmd, ucsIslCmdEcrReadRow, sParam, "3;%1;"
        For lIdx = 0 To 5
            If lIdx > 0 Then
                If LenB(At(vConfig, lOffset + lIdx)) = 0 Then
                    GoTo QH
                End If
                lTagNo = C_Lng(At(vConfig, lOffset + lIdx))
            End If
            If IsEmpty(JsonItem(oRetVal, "#" & lTagNo)) Then
                vSplit = Split(SendCommand(eCmd, Printf(sParam, lIdx + 1), LogName:=FUNC_NAME), ";")
                '--- 1 - disabled, 0 - not disabled
                If C_Lng(At(vSplit, 3)) = 0 Then
                    JsonItem(oRetVal, "#" & lTagNo) = Array(lIdx, At(vSplit, 2))
                End If
            End If
        Next
    End If
    Set pvGetPmtMapping = oRetVal
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Function pvToDevicePmtIndex(ByVal PmtType As UcsFiscalPaymentTypeEnum) As Long
    Const FUNC_NAME     As String = "pvToDevicePmtIndex"
    Dim vInfo           As Variant
    
    On Error GoTo EH
    pvToDevicePmtIndex = -1             ' unsupported
    If Not IsObject(m_uConfig.PmtMapping) Then
        Set m_uConfig.PmtMapping = pvGetPmtMapping()
    End If
    If Not m_uConfig.PmtMapping Is Nothing Then
        AssignVariant vInfo, JsonItem(C_Obj(m_uConfig.PmtMapping), "#" & (PmtType - ucsFscPmtCash))
        If IsObject(vInfo) Then
            pvToDevicePmtIndex = C_Lng(JsonItem(C_Obj(vInfo), 0))
        End If
    ElseIf IsDaisy Then
        Select Case PmtType
        Case ucsFscPmtCash
            pvToDevicePmtIndex = 0      ' P=SCash (В БРОЙ)
        Case ucsFscPmtCheque
            pvToDevicePmtIndex = 1      ' N=SCheck (С ЧЕК)
        Case ucsFscPmtCard
            pvToDevicePmtIndex = 2      ' C=SCards (КРЕДИТ)
'        Case ucsFscPmtCash
'            pvToDevicePmtIndex = 3      ' D=SCash (USD)
        Case ucsFscPmtEUR
            pvToDevicePmtIndex = 4      ' B=SCash (EURO)
        End Select
    ElseIf IsIncotex Then
        Select Case PmtType
        Case ucsFscPmtCash
            pvToDevicePmtIndex = 0      ' P=SCash (В БРОЙ)
        Case ucsFscPmtCard
            pvToDevicePmtIndex = 1      ' C=SCards (КР./ДЕБ.КАРТА)
        Case ucsFscPmtCheque
            pvToDevicePmtIndex = 2      ' N=SCheck (ЧЕК)
        Case ucsFscPmtCustom1
            pvToDevicePmtIndex = 3      ' D=SR1 (НЗОК)
        Case ucsFscPmtEUR
            pvToDevicePmtIndex = 4      ' B=SCash (ДОП.ВАЛУТА)
        End Select
    ElseIf IsEltrade Then
        Select Case PmtType
        Case ucsFscPmtCash
            pvToDevicePmtIndex = 0      ' P=SCash (В брой)
        Case ucsFscPmtCheque
            pvToDevicePmtIndex = 1      ' N=SCheck (С чек)
        Case ucsFscPmtCoupon
            pvToDevicePmtIndex = 2      ' C=ST (Талон)
        Case ucsFscPmtVoucher
            pvToDevicePmtIndex = 3      ' D=SOT (външни талони)
        Case ucsFscPmtPackaging
            pvToDevicePmtIndex = 4      ' I=SP (амбалаж)
        Case ucsFscPmtMaintenance
            pvToDevicePmtIndex = 5      ' J=SSelf (вътрешно обслужване)
        Case ucsFscPmtDamage
            pvToDevicePmtIndex = 6      ' K=SDmg (повреди)
        Case ucsFscPmtCard
            pvToDevicePmtIndex = 7      ' L=SCards (кредитни/дебитни карти)
        Case ucsFscPmtBank
            pvToDevicePmtIndex = 8      ' M=SW (банкови трансфери)
        Case ucsFscPmtCustom1
            pvToDevicePmtIndex = 9      ' Q=SR1 (НЗОК)
        Case ucsFscPmtCustom2
            pvToDevicePmtIndex = 10     ' R=SR2 (Резерв 2)
        End Select
    ElseIf IsEcr Then
        Select Case PmtType
        Case ucsFscPmtCash
            pvToDevicePmtIndex = 0      ' P=SCash (В БРОЙ)
'        Case ucsFscPmtCard
'            pvToDevicePmtIndex = 1      ' N=SCards (КРЕДИТ)
        Case ucsFscPmtCard
            pvToDevicePmtIndex = 2      ' C=SCards (ДЕБ.КАРТА)
        Case ucsFscPmtCustom1
            pvToDevicePmtIndex = 3      ' D=SR1 (НЗОК)
        Case ucsFscPmtVoucher
            pvToDevicePmtIndex = 4      ' J=SOT (ВАУЧЕР)
        Case ucsFscPmtCoupon
            pvToDevicePmtIndex = 5      ' I=ST (КУПОН)
        End Select
    Else
        Select Case PmtType
        Case ucsFscPmtCash
            pvToDevicePmtIndex = 0      ' P=SCash (В брой)
'        Case ucsFscPmtCard
'            pvToDevicePmtIndex = 1      ' N=SCards (С кредит)
        Case ucsFscPmtCheque
            pvToDevicePmtIndex = 2      ' C=SChecks (С чек)
        Case ucsFscPmtCard
            pvToDevicePmtIndex = 3      ' D=SCards (С дебитна карта)
        Case ucsFscPmtCustom1
            pvToDevicePmtIndex = 4      ' I=SR1 (Програмируем тип плащане 1)
'        Case ucsFscPmtCustom1
'            pvToDevicePmtIndex = 5      ' J=SR1 (Програмируем тип плащане 2)
'        Case ucsFscPmtCustom1
'            pvToDevicePmtIndex = 6      ' K=SR1 (Програмируем тип плащане 3)
        Case ucsFscPmtCustom2
            pvToDevicePmtIndex = 7      ' L=SR2 (Програмируем тип плащане 4)
'        Case ucsFscPmtCash
'            pvToDevicePmtIndex = 8      ' unused
'        Case ucsFscPmtCash
'            pvToDevicePmtIndex = 9      ' unused
        Case ucsFscPmtCoupon
            pvToDevicePmtIndex = 10     ' m=ST (Талони)
        Case ucsFscPmtVoucher
            pvToDevicePmtIndex = 11     ' n=SOT (Външни талони)
        Case ucsFscPmtPackaging
            pvToDevicePmtIndex = 12     ' o=SP (Амбалаж)
        Case ucsFscPmtMaintenance
            pvToDevicePmtIndex = 13     ' p=SSelf (Вътрешно обслужване)
        Case ucsFscPmtDamage
            pvToDevicePmtIndex = 14     ' q=SDmg (Повреди)
        Case ucsFscPmtBank
            pvToDevicePmtIndex = 15     ' r=SW (Банкови трансфери)
        End Select
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Function pvToDevicePmtName(ByVal PmtType As UcsFiscalPaymentTypeEnum) As String
    Const FUNC_NAME     As String = "pvToDevicePmtName"
    Const CONF_NAME     As String = "GetPaymentName"
    Dim eCmd            As UcsIslCommandsEnum
    Dim sParam          As String
    Dim lIndex          As Long
    
    On Error GoTo EH
    lIndex = pvToDevicePmtIndex(PmtType)
    If lIndex >= 0 Then
        If IsDaisy Then
            pvGetCommandConfig CONF_NAME, "ExtendedInitText" & PmtType, eCmd, ucsIslCmdExtendedInitText, sParam, "R" & (60 + lIndex)
        ElseIf IsIncotex Then
            If lIndex = 0 Then
                pvToDevicePmtName = pvText(ucsTxtPaymentCash)
                GoTo QH
            End If
            pvGetCommandConfig CONF_NAME, "ExtendedInitText" & PmtType, eCmd, ucsIslCmdExtendedInitText, sParam, "R" & (9 + lIndex)
        Else
            pvGetCommandConfig CONF_NAME, "InitPaymentType" & PmtType, eCmd, ucsIslCmdInitPaymentType, sParam, Mid$(DefPmtTypes, lIndex + 1, 1)
        End If
        pvToDevicePmtName = RTrim$(pvStripBold(SendCommand(eCmd, sParam, LogName:=FUNC_NAME)))
        '--- check for failure (returns 'F')
        If pvToDevicePmtName = "F" Then
            If lIndex < 4 Then
                pvToDevicePmtName = pvText(ucsTxtPaymentCash + lIndex)
            Else
                pvToDevicePmtName = vbNullString
            End If
        ElseIf IsIncotex Then
            pvToDevicePmtName = At(Split(pvToDevicePmtName, DEF_SEP), 1)
        End If
    End If
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Sub pvGetCommandConfig( _
            sFunction As String, _
            sKey As String, _
            eCmd As UcsIslCommandsEnum, _
            ByVal eCmdDefault As UcsIslCommandsEnum, _
            sParam As String, _
            sParamDefault As String, _
            ParamArray A() As Variant)
    Const FUNC_NAME     As String = "pvGetCommandConfig"
    Dim lIdx            As Long
    
    On Error GoTo EH
    eCmd = GetConfigForCommand(m_uConfig.ConfigCommands, m_uConfig.LocalizedCommands, sFunction, sKey, eCmdDefault)
    sParam = GetConfigForCommand(m_uConfig.ConfigCommands, m_uConfig.LocalizedCommands, sFunction, sKey & "Param", sParamDefault)
    For lIdx = 0 To UBound(A) Step 3
        A(lIdx + 1) = GetConfigForCommand(m_uConfig.ConfigCommands, m_uConfig.LocalizedCommands, sFunction, sKey & A(lIdx + 0), A(lIdx + 2))
    Next
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Private Sub pvSetLastError(sError As String, Optional ByVal ErrNum As UcsFiscalErrorsEnum = -1)
    Const FUNC_NAME     As String = "pvSetLastError"
    
    On Error GoTo EH
    If ErrNum < 0 Then
        m_uPrintData.LastErrNo = IIf(LenB(sError) = 0, ucsFerNone, ucsFerGeneralError)
    Else
        m_uPrintData.LastErrNo = ErrNum
    End If
    m_uPrintData.LastError = sError
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Private Function pvParseDdmmyyDate(sText As String) As Date
    If Len(sText) = 6 And IsNumeric(sText) Then
        pvParseDdmmyyDate = DateSerial(2000 + C_Lng(Mid$(sText, 5, 2)), C_Lng(Mid$(sText, 3, 2)), C_Lng(Mid$(sText, 1, 2)))
    End If
QH:
End Function

Private Sub pvEltradeMergeInfoArray(vResult As Variant)
    Const FUNC_NAME     As String = "pvEltradeMergeInfoArray"
    
    On Error GoTo EH
    vResult(2) = At(vResult, 0) & DEF_SEP & At(vResult, 1) & " " & At(vResult, 2)
    vResult(1) = vbNullString
    vResult(0) = vbNullString
    vResult = Split(Mid$(Join(vResult, DEF_SEP), 3), DEF_SEP)
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Initialize()
    Dim eIdx            As UcsIslConfigIndexEnum
    
    m_lConnectionTimeout = DEF_CONNECTIONTIMEOUT
    m_lCommandTimeout = DEF_TIMEOUT
    m_lSequence = DEF_SEQUENCE
    LocalizedText(ucsFscLciInternalErrors) = STR_INTERNAL
    LocalizedText(ucsFscLciDipSwitchesTexts) = STR_DIP_SWITCHES
    LocalizedText(ucsFscLciReceiptTexts) = STR_RECEIPT_TEXTS
    eIdx = ucsIslCfgDatecsModelA
    LocalizedText(ucsFscLciPrinterStatuses, eIdx) = STR_STATUSES_A
    LocalizedText(ucsFscLciPrinterErrors, eIdx) = STR_ERRORS_A
    LocalizedText(ucsFscLciMemoryTexts, eIdx) = STR_MEMORY_A
    eIdx = ucsIslCfgDatecsModelB
    LocalizedText(ucsFscLciPrinterStatuses, eIdx) = STR_STATUSES_B
    LocalizedText(ucsFscLciPrinterErrors, eIdx) = STR_ERRORS_B
    LocalizedText(ucsFscLciMemoryTexts, eIdx) = STR_MEMORY_B
    eIdx = ucsIslCfgDaisy
    LocalizedText(ucsFscLciPrinterStatuses, eIdx) = STR_STATUSES_DAISY
    LocalizedText(ucsFscLciPrinterErrors, eIdx) = STR_ERRORS_DAISY
    LocalizedText(ucsFscLciMemoryTexts, eIdx) = STR_MEMORY_DAISY
    LocalizedText(ucsFscLciExtendedErrors, eIdx) = STR_EXTENDED_DAISY
    eIdx = ucsIslCfgIncotex
    LocalizedText(ucsFscLciPrinterStatuses, eIdx) = STR_STATUSES_INCO
    LocalizedText(ucsFscLciPrinterErrors, eIdx) = STR_ERRORS_INCO
    LocalizedText(ucsFscLciMemoryTexts, eIdx) = STR_MEMORY_INCO
    LocalizedText(ucsFscLciExtendedErrors, eIdx) = STR_EXTENDED_INCO
    Set m_uConfig.ConfigCommands = GetConfigCollection(STR_PROTOCOL_DATECS_FP, "Commands")
End Sub

'=========================================================================
' IDeviceProtocol interface
'=========================================================================

Private Function IDeviceProtocol_AddBarcode(ByVal BarcodeType As UcsFiscalBarcodeTypeEnum, Text As String, Optional ByVal Height As Long) As Boolean
    IDeviceProtocol_AddBarcode = PpdAddBarcode(m_uPrintData, BarcodeType, Text, Height)
End Function

Private Function IDeviceProtocol_AddDiscount(ByVal DiscType As UcsFiscalDiscountTypeEnum, ByVal Value As Double) As Boolean
    IDeviceProtocol_AddDiscount = PpdAddDiscount(m_uPrintData, DiscType, Value)
End Function

Private Function IDeviceProtocol_AddLine(Line As String, Optional Command As String, Optional ByVal WordWrap As Boolean = True) As Boolean
    IDeviceProtocol_AddLine = PpdAddLine(m_uPrintData, Line, WordWrap)
End Function

Private Function IDeviceProtocol_AddPayment(ByVal PmtType As UcsFiscalPaymentTypeEnum, Optional PmtName As String, Optional ByVal Amount As Double, Optional ByVal Rate As Double) As Boolean
    IDeviceProtocol_AddPayment = PpdAddPayment(m_uPrintData, PmtType, PmtName, Amount)
End Function

Private Function IDeviceProtocol_AddPLU(Name As String, ByVal Price As Double, Optional ByVal Quantity As Double = 1, Optional ByVal TaxGroup As Long = 2, Optional UnitOfMeasure As String, Optional ByVal DepartmentNo As Long) As Boolean
    IDeviceProtocol_AddPLU = PpdAddPLU(m_uPrintData, Name, Price, Quantity, TaxGroup, UnitOfMeasure, DepartmentNo)
End Function

Private Function IDeviceProtocol_AutodetectDevices(Ports As Variant) As Variant
    IDeviceProtocol_AutodetectDevices = AutodetectDevices(Ports)
End Function

Private Function IDeviceProtocol_CancelReceipt() As Boolean
    IDeviceProtocol_CancelReceipt = CancelReceipt()
End Function

Private Function IDeviceProtocol_PrintServiceDeposit(OperatorCode As String, OperatorPassword As String, ByVal Amount As Double, Optional ReceiptNo As String, Optional ReceiptDate As Date, Optional ReceiptAmount As Double) As Variant
    IDeviceProtocol_PrintServiceDeposit = PrintServiceDeposit(OperatorCode, Amount, ReceiptNo, ReceiptDate, ReceiptAmount)
End Function

Private Function IDeviceProtocol_Connect() As Boolean
    IDeviceProtocol_Connect = Connect()
End Function

Private Function IDeviceProtocol_PrintDuplicate(Optional DuplicateNo As String, Optional ReceiptNo As String, Optional ReceiptDate As Date) As Boolean
    IDeviceProtocol_PrintDuplicate = PrintDuplicate(DuplicateNo, ReceiptNo, ReceiptDate)
End Function

Private Function IDeviceProtocol_Disconnect() As Boolean
    IDeviceProtocol_Disconnect = Disconnect()
End Function

Private Function IDeviceProtocol_PrintReceipt(Optional ResumeToken As String, Optional ReceiptNo As String, Optional ReceiptDate As Date, Optional ReceiptAmount As Double, Optional InvoiceNo As String) As Boolean
    If PpdEndReceipt(m_uPrintData, ResumeToken) Then
        IDeviceProtocol_PrintReceipt = pvPrintReceipt(m_uPrintData.ExecCtx, m_uPrintData.Row, m_uPrintData.RowCount)
        With m_uPrintData.ExecCtx
            ReceiptNo = .ReceiptNo
            ReceiptDate = .ReceiptDate
            ReceiptAmount = .ReceiptAmount
            InvoiceNo = .InvoiceNo
        End With
    End If
End Function

Private Function IDeviceProtocol_GetCharsPerLine() As Long
    IDeviceProtocol_GetCharsPerLine = pvRowChars
End Function

Private Function IDeviceProtocol_GetCommentTextMaxLength() As Long
    IDeviceProtocol_GetCommentTextMaxLength = pvCommentChars
End Function

Private Function IDeviceProtocol_GetItemNameMaxLength() As Long
    IDeviceProtocol_GetItemNameMaxLength = pvItemChars
End Function

Private Function IDeviceProtocol_GetClock() As Date
    IDeviceProtocol_GetClock = GetClock()
End Function

Private Function IDeviceProtocol_GetCommandLog() As String
    
End Function

Private Function IDeviceProtocol_GetOperatorName(OperatorCode As String) As String
    IDeviceProtocol_GetOperatorName = GetOperatorName(OperatorCode)
End Function

Private Function IDeviceProtocol_GetDefaultPassword(OperatorCode As String) As String
    IDeviceProtocol_GetDefaultPassword = GetDefaultPassword(OperatorCode)
End Function

Private Function IDeviceProtocol_GetDeviceModel() As String
    IDeviceProtocol_GetDeviceModel = GetDeviceModel()
End Function

Private Function IDeviceProtocol_GetDeviceFirmware() As String
    IDeviceProtocol_GetDeviceFirmware = GetDeviceFirmware()
End Function

Private Function IDeviceProtocol_GetDeviceProtocol() As String
    IDeviceProtocol_GetDeviceProtocol = GetDeviceProtocol()
End Function

Private Function IDeviceProtocol_GetDeviceStatus(StatusText As String) As UcsFiscalStatusCodeEnum
    IDeviceProtocol_GetDeviceStatus = GetDeviceStatus(StatusText)
End Function

Private Function IDeviceProtocol_GetFiscalMemoryNo() As String
    IDeviceProtocol_GetFiscalMemoryNo = GetFiscalMemoryNo()
End Function

Private Function IDeviceProtocol_GetFooterText(ByVal Index As Long) As String
    IDeviceProtocol_GetFooterText = GetFooterText(Index)
End Function

Private Function IDeviceProtocol_GetHeaderText(ByVal Index As Long) As String
    IDeviceProtocol_GetHeaderText = GetHeaderText(Index)
End Function

Private Function IDeviceProtocol_GetLastError(Optional ErrNo As UcsFiscalErrorsEnum) As String
    ErrNo = LastErrNo
    IDeviceProtocol_GetLastError = LastError
End Function

Private Function IDeviceProtocol_GetLastReceiptNo() As String
    IDeviceProtocol_GetLastReceiptNo = GetLastReceiptNo()
End Function

Private Function IDeviceProtocol_GetLastInvoiceNo() As String
    IDeviceProtocol_GetLastInvoiceNo = GetLastInvoiceNo()
End Function

Private Function IDeviceProtocol_GetLastQRCodeInfo(Optional LastReceiptDate As Date) As String
    IDeviceProtocol_GetLastQRCodeInfo = GetLastQRCodeInfo(LastReceiptDate)
End Function

Private Function IDeviceProtocol_GetPaymentName(ByVal Index As Long) As String
    IDeviceProtocol_GetPaymentName = GetPaymentName(Index)
End Function

Private Function IDeviceProtocol_GetResumeToken() As String
    IDeviceProtocol_GetResumeToken = PpdGetResumeToken(m_uPrintData)
End Function

Private Function IDeviceProtocol_GetDeviceSerialNo() As String
    IDeviceProtocol_GetDeviceSerialNo = GetDeviceSerialNo()
End Function

Private Function IDeviceProtocol_GetTaxCaption() As String
    IDeviceProtocol_GetTaxCaption = GetTaxCaption()
End Function

Private Function IDeviceProtocol_GetTaxNo() As String
    IDeviceProtocol_GetTaxNo = GetTaxNo()
End Function

Private Function IDeviceProtocol_GetTotalsByPaymentTypes() As Variant
    IDeviceProtocol_GetTotalsByPaymentTypes = GetTotalsByPaymentTypes()
End Function

Private Function IDeviceProtocol_GetTotalsByTaxGroups() As Variant
    IDeviceProtocol_GetTotalsByTaxGroups = GetTotalsByTaxGroups()
End Function

Private Function IDeviceProtocol_GetTotalsStats() As Variant
    IDeviceProtocol_GetTotalsStats = GetTotalsStats()
End Function

Private Function IDeviceProtocol_Init(DeviceString As String) As Boolean
    IDeviceProtocol_Init = Init(DeviceString)
End Function

Private Property Get IDeviceProtocol_IsConnected() As Boolean
    IDeviceProtocol_IsConnected = IsConnected
End Property

Private Function IDeviceProtocol_OpenDrawer() As Boolean
    IDeviceProtocol_OpenDrawer = OpenDrawer()
End Function

Private Function IDeviceProtocol_PrintMonthlyReport(ByVal ReportType As UcsFiscalReportsTypeEnum, ByVal StartDate As Date, ByVal EndDate As Date, Optional ReceiptNo As String, Optional ReceiptDate As Date) As Boolean
    IDeviceProtocol_PrintMonthlyReport = PrintMonthlyReport(ReportType, StartDate, EndDate, ReceiptNo, ReceiptDate)
End Function

Private Function IDeviceProtocol_PrintDailyXReport(ByVal ReportType As UcsFiscalReportsTypeEnum, Optional ReceiptNo As String, Optional ReceiptDate As Date) As Boolean
    IDeviceProtocol_PrintDailyXReport = PrintDailyXReport(ReportType, ReceiptNo, ReceiptDate)
End Function

Private Function IDeviceProtocol_PrintDailyZReport(ByVal ReportType As UcsFiscalReportsTypeEnum, Optional ReceiptNo As String, Optional ReceiptDate As Date) As Boolean
    IDeviceProtocol_PrintDailyZReport = PrintDailyZReport(ReportType, ReceiptNo, ReceiptDate)
End Function

Private Function IDeviceProtocol_SetClock(ByVal NewDate As Date) As Boolean
    IDeviceProtocol_SetClock = SetClock(NewDate)
End Function

Private Sub IDeviceProtocol_SetLocalizedText(ByVal Index As UcsFiscalLocalizedIndexesEnum, Text As String, Optional ByVal ConfigIndex As Long)
    LocalizedText(Index, ConfigIndex) = Text
End Sub

Private Function IDeviceProtocol_StartReceipt(ByVal ReceiptType As UcsFiscalReceiptTypeEnum, OperatorCode As String, OperatorName As String, OperatorPassword As String, Optional UniqueSaleNo As String, Optional TableNo As String, Optional ByVal DisablePrinting As Boolean, Optional InvDocNo As String, Optional InvCgTaxNo As String, Optional ByVal InvCgTaxNoType As UcsFiscalTaxNoTypeEnum, Optional InvCgVatNo As String, Optional InvCgName As String, Optional InvCgCity As String, Optional InvCgAddress As String, Optional InvCgPrsReceive As String, Optional OwnData As String, Optional ByVal RevType As UcsFiscalReversalTypeEnum, Optional RevReceiptNo As String, Optional ByVal RevReceiptDate As Date, Optional RevFiscalMemoryNo As String, Optional RevInvoiceNo As String, Optional RevReason As String) As Boolean
    Dim sOperCode       As String
    Dim sOperPass       As String
    
    m_uPrintData.Config.RowChars = pvRowChars
    m_uPrintData.Config.CommentChars = pvCommentChars
    m_uPrintData.Config.ItemChars = pvItemChars
    sOperCode = Zn(OperatorCode, DEF_OPERNO)
    sOperPass = Zn(OperatorPassword, GetDefaultPassword(sOperCode))
    IDeviceProtocol_StartReceipt = PpdStartReceipt(m_uPrintData, ReceiptType, sOperCode, OperatorName, sOperPass, TableNo, UniqueSaleNo, DisablePrinting, InvDocNo, InvCgTaxNo, InvCgTaxNoType, InvCgVatNo, InvCgName, InvCgCity, InvCgAddress, InvCgPrsReceive, RevType, RevReceiptNo, RevReceiptDate, RevFiscalMemoryNo, RevInvoiceNo, RevReason, OwnData)
End Function

Private Function IDeviceProtocol_GetDepartmentName(DepartmentCode As String, Optional TaxGroup As Long) As String
    IDeviceProtocol_GetDepartmentName = GetDepartmentName(DepartmentCode, TaxGroup)
End Function

Private Function IDeviceProtocol_GetTaxRates() As Variant
    IDeviceProtocol_GetTaxRates = GetTaxRates()
End Function
