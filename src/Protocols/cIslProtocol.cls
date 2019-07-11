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
' UcsFP20 (c) 2008-2019 by Unicontsoft
'
' Unicontsoft Fiscal Printers Component 2.0
'
' This project is licensed under the terms of the MIT license
' See the LICENSE file in the project root for more information
'
'=========================================================================
'
' Protocol handler for Datecs, Daisy and Incotex fiscal printers (ISL protocol)
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

Public Enum UcsIslDipSwitchesEnum
    ucsDipCenterHeaderFooter = 2 ^ 0
    ucsDipPrintHeader = 2 ^ 1
    ucsDipSwitch_1_3 = 2 ^ 2
    ucsDipSwitch_1_4 = 2 ^ 3
    ucsDipZeroMemory = 2 ^ 4
    ucsDipTransparentDisplay = 2 ^ 5
    ucsDipShowToDisplay = 2 ^ 6
End Enum

Public Enum UcsIslFiscalMemoryEnum
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

Public Enum UcsIslCommandsEnum
    ucsFpcNoop = -1
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
    ucsFpcInitEcTape = 119
    ucsFpcExtendedInitText = 149                '-- command above &H7F
    ucsFpcExtendedInitSetting = 150             '-- command above &H7F
    ucsFpcExtendedInitCurrencyRate = 151        '-- command above &H7F
    ucsFpcExtendedInitInvoiceNo = 180           '-- command above &H7F
    '--- transaction
    ucsFpcNonFiscalOpen = 38
    ucsFpcNonFiscalClose = 39
    ucsFpcNonFiscalText = 42
    ucsFpcLineFeed = 44
    ucsFpcFiscalStorno = 46
    ucsFpcFiscalOpen = 48
    ucsFpcFiscalSell = 49
    ucsFpcFiscalSubtotal = 51
    ucsFpcFiscalSellAndDisplay = 52
    ucsFpcFiscalPayment = 53
    ucsFpcFiscalText = 54
    ucsFpcFiscalClose = 56
    ucsFpcFiscalCgInfo = 57
    ucsFpcFiscalCancel = 60
    ucsFpcAdminCashDebitCredit = 70
    ucsFpcExtendedFiscalCancel = 130            '-- command above &H7F
    '--- info
    ucsFpcInfoServiceContract = 34
    ucsFpcInfoNap = 37
    ucsFpcInfoDateTime = 62
    ucsFpcInfoLastFiscalBlock = 64
    ucsFpcInfoTotalsByVatGroups = 65
    ucsFpcInfoDailyStats = 67
    ucsFpcInfoFreeMemoryRecords = 68
    ucsFpcInfoStatus = 74
    ucsFpcInfoTransaction = 76
    ucsFpcInfoLastFiscalDateTime = 86
    ucsFpcInfoDepartment = 88
    ucsFpcInfoDiagnostics = 90
    ucsFpcInfoTaxRates = 97
    ucsFpcInfoBulstat = 99
    ucsFpcInfoSums = 103
    ucsFpcInfoTotalsByPayments = 110
    ucsFpcInfoOperator = 112
    ucsFpcInfoLastReceipt = 113
    ucsFpcInfoFiscalRecord = 114
    ucsFpcInfoMemoryDump = 116
    ucsFpcExtendedInfoConsts = 128              '-- command above &H7F
    '--- print
    ucsFpcPrintReportTaxRates = 50
    ucsFpcPrintDailyReport = 69
    ucsFpcPrintDiagnostics = 71
    ucsFpcPrintReportByNumberDetailed = 73
    ucsFpcPrintReportByDateShort = 79
    ucsFpcPrintBarcode = 84
    ucsFpcPrintSeparator = 92
    ucsFpcPrintReportByDateDetailed = 94
    ucsFpcPrintReportByNumberShort = 95
    ucsFpcPrintReportByOperators = 105
    ucsFpcPrintDailyReportItems = 108
    ucsFpcPrintDuplicate = 109
    ucsFpcPrintReportByItems = 111
    ucsFpcPrintDailyReportDepartments = 117
    ucsFpcPrintDailyReportItemsDepartments = 118
    '--- paper
    ucsFpcPaperLinefeed = 44
    ucsFpcPaperCut = 45
    '--- display
    ucsFpcDisplayClear = 33
    ucsFpcDisplaySecondLine = 35
    ucsFpcDisplayFirstLine = 47
    ucsFpcDisplayClock = 63
    ucsFpcDisplayText = 100
    '--- misc
    ucsFpcMiscRingSound = 80
    '--- ECR
    ucsFpcEcrReadRow = &HFF84&
    ucsFpcEcrWriteRow = &HFF86&
End Enum

'=========================================================================
' API
'=========================================================================

'--- error codes
Private Const ERROR_IO_PENDING              As Long = 997
'--- for PurgeComm
Private Const PURGE_TXCLEAR                 As Long = &H4     '  Kill the transmit queue if there.
Private Const PURGE_RXCLEAR                 As Long = &H8     '  Kill the typeahead buffer if there.

Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function CreateEvent Lib "kernel32" Alias "CreateEventA" (ByVal lpEventAttributes As Long, ByVal bManualReset As Long, ByVal bInitialState As Long, ByVal lpName As Long) As Long
Private Declare Function WaitForMultipleObjects Lib "kernel32" (ByVal nCount As Long, lpHandles As Long, ByVal bWaitAll As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function PurgeComm Lib "kernel32" (ByVal hFile As Long, ByVal dwFlags As Long) As Long
Private Declare Function GetOverlappedResult Lib "kernel32" (ByVal hFile As Long, lpOverlapped As OVERLAPPED, lpNumberOfBytesTransferred As Long, ByVal bWait As Long) As Long

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

Private Const STR_INTERNAL              As String = "No device info set|CreateFile failed: %1|SetCommTimeouts failed: %1|BuildCommDCB failed: %1|SetCommState failed: %1|WriteFile failure: %1|Timeout waiting for response|ReadFile failed: %1|Invalid message format or checksum (NAK)|Already connected|Not connected|Invalid response length: %1|Response missing postamble: %1|Response missing preamble: %1|WaitCommEvent failed: %1|No receipt started| or invalid operator password|Error in line %1: %2|Cannot cancel previous receipt|Invalid report type|Error %1|Memory %1"
Private Const STR_STATUSES              As String = "Syntax error|Invalid command|Clock not set|3|Feed error|5|6|7|Overflow|Command not allowed|Memory reset|3|Memory corrupt|5|6|7|Out of paper|1|2|Fiscal printing|4|Nonfiscal printing|6|7"
Private Const STR_ERRORS_ONLY           As String = "Syntax error|Invalid command|2|3|Feed error|5|6|7|1|Command not allowed|Memory reset|3|Memory corrupt|5|6|7|Out of paper|1|2|3|4|5|6|7"
Private Const STR_DIP_SWITCHES          As String = "Center header and footer|Enable header|Sw1.3|Sw1.4|Zero memory|Transparent display|No display|7"
Private Const STR_MEMORY                As String = "Write error|1|No memory module|Memory low|Out of memory|Fiscal memory error|6|7|Read only|Memory formatted|2|Fiscal mode|Tax rates set|Tax number set|6|7"
Private Const STR_EXT_ERRORS            As String = "1|Overflow after operation|3|No more sales in this bill|4|No more payments in this bill|5|Attempt to perform zero transaction|6|Attempt to perform a sale, after payment is initiated|7|Operation not allowed|8|TAX group with disabled sales|11|More than one decimal point|12|More than one character '+' or '-'|13|Symbol '+' or '-' is not in the first position|14|Illegal symbol, e.g. barcode contains not only numbers|15|Extra numbers after decimal point|16|Extra characters|20|With current settings cannot process this command from PC|21|Value outside limits|22|Review system parameter 8|23|Attempt to ""deep"" void after discount/surcharge on subtotal|24|Attempt to ""deep"" void of non-existent transaction|25|Attempt to make a payment without having sales|26|Attempt to sell an item with quantity exceeding its stock|41|Barcode is incorrect (wrong checksum)|42|Attempt to sell without barcode|" & _
                                                    "43|Attempt to program with weight barcode|44|Attempt to sell without programmed barcode|45|Attempt to program already existing barcode|66|Incorrect password|71|!!! Incorrect data in fiscal memory !!!!|72|!!! Error saving in fiscal memory !!!!|76|Information needed from NRA servers|90|Needs a reset with monthly report|91|Needs a reset with daily z-report|92|Needs a reset with report by operators|93|Needs a reset with report by articles|94|Can not reprogram this field|81|Daily z-report overflow|83|Report by operators overflow|84|Report by articles overflow|84|Monthly report overflow|88|Electronic journal overflow|102|No communication between fiscal unit and tax terminal|104|Incorrect communication between fiscal unit and tax terminal|110|SIM card replaced|111|Error in communication between tax terminal and NRA servers|113|NRA server rejects data sent|" & _
                                                    "117|Failed to connect tax terminal with mobile operator network|118|The operation is denied|119|Incorrect value|120|Missing value"
Private Const STR_RECEIPT_TEXTS         As String = "CASH|CARD|CHEQUE|CREDIT|SURCHARGE %1|DISCOUNT %1|TOTAL|SUBTOTAL|INVOICE No %1|Sales %1|UNIT PRICE|QUANTITY|VALUE|SUM GROUP %1|VAT %1=%2|NETTO SUM|TOTAL|%1 ITEMS|1 ITEM|CHANGE|SELLER: |RECEIVER: |BUYER: |TAX No: |VAT No: |EIC"
Private Const DEF_AUTODETECTTIMEOUT     As Long = 100
Private Const DEF_DEVICEINFOTIMEOUT     As Long = 1000
Private Const DEF_AUTODETECTSPEEDS      As String = "115200|38400|9600|19200|57600"
Private Const DEF_TIMEOUT               As Long = 3000
Private Const DEF_OPERNO                As String = "1"
Private Const DEF_PASSWORD              As String = "0000"
Private Const DEF_ROW_CHARS             As Long = 30
Private Const DEF_MAX_RECEIPT_ROWS      As Long = 120
Private Const DEF_MIN_DISCOUNT          As Double = -99
Private Const DEF_MAX_DISCOUNT          As Double = 99
Private Const DEF_MAX_PAYMENT_LEN       As Long = 24
Private Const DEF_MAX_CGINFO            As String = "14|24|25|25|14|35|35"
Private Const DEF_MAX_OPERATOR_NAME     As Long = 24
Private Const DEF_COMM_EVENT_TIMEOUT    As Long = 250
Private Const DEF_DELIM                 As String = ","
Private Const MIK_OFFSET                As Long = &H80
'Private Const WIN_OFFSET                As Long = &HC0
Private Const ucsFscDstPluAbs           As Long = ucsFscDstPlu + 100
Private Const ucsFscDstSubtotalAbs      As Long = ucsFscDstSubtotal + 100
Private Const ucsFscRetNonfiscal        As Long = ucsFscRetSale + 100
Private Const FORMAT_DATE_EXTENDED      As String = "dd\-mm\-yy hh\:nn\:ss"
Private Const FORMAT_DATE_SHORT         As String = "ddmmyyhhnnss"

Private m_sDevice                   As String
Private m_oConnector                As IDeviceConnector
Private m_lTimeout                  As Long
Private m_lSequence                 As Long
Private m_eStatus                   As UcsIslStatusBitsEnum
Private m_eDip                      As UcsIslDipSwitchesEnum
Private m_eMemory                   As UcsIslFiscalMemoryEnum
Private m_eLastErrNo            As UcsFiscalErrorsEnum
Private m_sLastError                As String
Private m_bExceptions               As Boolean
Private m_uRow()                    As UcsRowData
Private m_lRowCount                 As Long
Private m_sCommandLog               As String
Private m_vInfoDiagnostics          As Variant
Private m_uConfig                   As UcsConfigValues
Private m_uCtxPrint                 As UcsContext

Private Enum UcsInternalErrors
    ucsErrNoDeviceInfoSet
    ucsErrCreateFileFailed
    ucsErrSetCommTimeoutsFailed
    ucsErrBuildCommDCBFailed
    ucsErrSetCommStateFailed
    [__ucsErrWriteFileFailure]
    ucsErrTimeoutWaitingForResponse
    [__ucsErrReadFileFailed]
    [__ucsErrInvalidMessageFormatOrChecksum]
    [__ucsErrAlreadyConnected]
    [__ucsErrNotConnected]
    ucsErrInvalidResponseLength
    ucsErrResponseMissingPostamble
    ucsErrResponseMissingPreamble
    [__ucsErrWaitCommEventFailed]
    ucsErrNoReceiptStarted
    ucsErrInvalidOperatorPassword
    ucsErrErrorInLine
    ucsErrCannotCancelReceipt
    ucsErrInvalidReportType
    ucsErrUnknownError
    ucsErrUnknownMemory
End Enum

Private Enum UcsReceiptTextsEnum
    ucsTxtPaymentCash
    ucsTxtPaymentCard
    ucsTxtPaymentCheque
    ucsTxtPaymentCredit
    ucsTxtSurcharge
    ucsTxtDiscount
    [__ucsTxtTotal]
    [__ucsTxtSubTotal]
    [__ucsTxtHeaderInvoiceNo]
    ucsTxtPluSales
    [__ucsTxtPluPrice]
    [__ucsTxtPluQuantity]
    [__ucsTxtPluTotal]
    [__ucsTxtGroupTotal]
    [__ucsTxtGroupVat]
    [__ucsTxtGroupNetto]
    [__ucsTxtFooterTotal]
    [__ucsTxtFooterArticles]
    [__ucsTxtFooterArticleSingle]
    [__ucsTxtFooterChange]
    [__ucsTxtFooterSeller]
    [__ucsTxtFooterReceiver]
    [__ucsTxtFooterBuyer]
    [__ucsTxtFooterTaxNo]
    [__ucsTxtFooterVatNo]
    ucsTxtTaxNo
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
    InitUniqueSaleNo    As String
    InitInvData         As Variant
    InitStornoData      As Variant
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
    ucsInvDocNo
    ucsInvCgTaxNo
    ucsInvCgVatNo
    ucsInvCgName
    ucsInvCgCity
    ucsInvCgAddress
    ucsInvCgPrsReceive
    ucsInvCgTaxNoType
End Enum

Private Enum UcsStoreDataIndex
    ucsStoType
    ucsStoReceiptNo
    ucsStoReceiptDate
    ucsStoFiscalMemoryNo
    ucsStoInvoiceNo
    ucsStoReason
End Enum

Private Type UcsContext
    GrpTotal(1 To 8)    As Double
    Paid                As Double
    PluCount            As Long
    PmtPrinted          As Boolean
    ChangePrinted       As Boolean
    Row                 As Long
End Type

Private Enum UcsCharset
    ucsCrsAuto
    ucsCrsWindows
    ucsCrsMik
End Enum

Private Type UcsAutodetect
    Connector           As cSerialPortConnector
    WriteOver           As OVERLAPPED
    ReadOver            As OVERLAPPED
    Buffer              As String
End Type

Private Type UcsConfigValues
    Charset             As UcsCharset
    IsEcr               As VbTriState
    IsDaisy             As Boolean
    IsIncotex           As Boolean
    RowChars            As Long
    ItemChars           As Long
    AbsoluteDiscount    As Boolean
    NegativePrices      As Boolean
    MaxReceiptRows      As Long
    MinDiscount         As Double
    MaxDiscount         As Double
    MaxPaymentLen       As Long
    MaxCgInfo(0 To 6)   As Long
    MaxOperatorName     As Long
    CommEventTimeout    As Long
    DefPmtTypes         As String
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

Private Sub DebugDataDump(sFunc As String, sPrefix As String, sData As String)
    OutputDebugDataDump MODULE_NAME, sFunc, sPrefix, sData
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

Property Get Status(ByVal eValue As UcsIslStatusBitsEnum) As Boolean
    Status = (m_eStatus And eValue) <> 0
End Property

Property Get StatusText() As String
    StatusText = pvGetFlagsText(m_eStatus, m_uConfig.LocalizedText(ucsFscLciPrinterStatuses))
End Property

Property Get ErrorText() As String
    ErrorText = pvGetFlagsText(m_eStatus, m_uConfig.LocalizedText(ucsFscLciPrinterErrors))
    If IsDaisy And (m_eDip And &H7F) <> 0 Then
        ErrorText = IIf(LenB(ErrorText) <> 0, ErrorText & vbCrLf, vbNullString) & _
            Zn(At(m_uConfig.LocalizedText(ucsFscLciExtendedErrors), m_eDip And &H7F), Printf(pvInternal(ucsErrUnknownError), m_eDip And &H7F))
    End If
End Property

Property Get Dip(ByVal eValue As UcsIslDipSwitchesEnum) As Boolean
    Dip = (m_eDip And eValue) <> 0
End Property

Property Get DipText() As String
    If Not IsDaisy Then
        DipText = pvGetFlagsText(m_eDip, m_uConfig.LocalizedText(ucsFscLciDipSwitchesTexts))
    End If
End Property

Property Get Memory(ByVal eValue As UcsIslFiscalMemoryEnum) As Boolean
    Memory = (m_eMemory And eValue) <> 0
End Property

Property Get MemoryText() As String
    MemoryText = pvGetFlagsText(m_eMemory, m_uConfig.LocalizedText(ucsFscLciMemoryTexts))
End Property

Property Get LastErrNo() As UcsFiscalErrorsEnum
    LastErrNo = m_eLastErrNo
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
    IsConnected = Not m_oConnector Is Nothing
End Property

Property Get LocalizedText(ByVal eIdx As UcsFiscalLocalizedIndexesEnum) As String
    If IsArray(m_uConfig.LocalizedText(eIdx)) Then
        LocalizedText = Join(m_uConfig.LocalizedText(eIdx), "|")
    End If
End Property

Property Let LocalizedText(ByVal eIdx As UcsFiscalLocalizedIndexesEnum, sValue As String)
    m_uConfig.LocalizedText(eIdx) = SplitOrReindex(sValue, "|")
End Property

Property Get IsEcr() As Boolean
    Const FUNC_NAME     As String = "IsEcr [get]"
    Const CONF_NAME     As String = "IsEcr"
    Dim lTimeout        As Long
    Dim eCmd            As UcsIslCommandsEnum
    Dim sParam          As String
    
    On Error GoTo EH
    If m_uConfig.IsEcr = vbUseDefault Then
        pvGetCommandConfig CONF_NAME, "EcrReadRow", eCmd, ucsFpcEcrReadRow, sParam, "2;1;", _
            "Timeout", lTimeout, 100
        m_uConfig.IsEcr = IIf(LenB(pvSendCommand(m_oConnector, lTimeout, eCmd, sParam)) <> 0, vbTrue, vbFalse)
    End If
    IsEcr = (m_uConfig.IsEcr = vbTrue)
    Exit Property
EH:
    RaiseError FUNC_NAME
End Property

Property Get IsDaisy() As Boolean
    IsDaisy = m_uConfig.IsDaisy
End Property

Property Get IsIncotex() As Boolean
    IsIncotex = m_uConfig.IsIncotex
End Property

Property Get DefPmtTypes() As String
    Const FUNC_NAME     As String = "DefPmtTypes [get]"
    
    On Error GoTo EH
    If LenB(m_uConfig.DefPmtTypes) = 0 Then
        If IsDaisy Or IsIncotex Then
            m_uConfig.DefPmtTypes = "PCNCDBCC"
        ElseIf IsEcr Then
            '--- ToDo: use ucsFpcEcrReadRow to check for swapped [2 - card]<->[3 - cheque] payment types
            m_uConfig.DefPmtTypes = "PCDNIJKL"
        Else
            m_uConfig.DefPmtTypes = "PDCNIJKL"
        End If
    End If
    DefPmtTypes = m_uConfig.DefPmtTypes
    Exit Property
EH:
    RaiseError FUNC_NAME
End Property

Private Property Get pvRowChars() As Long
    Const FUNC_NAME     As String = "pvRowChars"
    Const CONF_NAME     As String = "RowChars"
    Dim bPrevExceptions As Boolean
    Dim sDaisyPrefix    As String
    Dim sIncotexPrefix  As String
    Dim eCmd            As UcsIslCommandsEnum
    Dim sParam          As String
    Dim lIndex          As Long
    Dim sResult         As String
    
    On Error GoTo EH
    If m_uConfig.RowChars = 0 Then
        pvGetCommandConfig CONF_NAME, "ExtendedInfoConsts", eCmd, ucsFpcExtendedInfoConsts, sParam, vbNullString, _
            "Index", lIndex, 9, _
            "DaisyPrefix", sDaisyPrefix, "DY", _
            "IncotexPrefix", sIncotexPrefix, "IN"
        If Left$(GetDeviceSerialNo(), Len(sDaisyPrefix)) = sDaisyPrefix And LenB(sDaisyPrefix) <> 0 Or _
                Left$(GetDeviceSerialNo(), Len(sIncotexPrefix)) = sIncotexPrefix And LenB(sIncotexPrefix) <> 0 Then
            bPrevExceptions = pvSetExceptions(False)
            m_uConfig.RowChars = C_Lng(At(Split(SendCommand(eCmd, sParam), DEF_DELIM), lIndex))
            Exceptions = bPrevExceptions
        Else
            pvGetCommandConfig CONF_NAME, "InitLogo", eCmd, ucsFpcInitLogo, sParam, "R0"
            bPrevExceptions = pvSetExceptions(False)
            sResult = SendCommand(eCmd, sParam)
            Exceptions = bPrevExceptions
            '--- check for paper width above 57mm (432px) e.g. 80mm paper is 576px wide
            If Len(sResult) * 4 > 432 Then
                m_uConfig.RowChars = Len(sResult) * 4 / 12 - 6
            End If
        End If
        If m_uConfig.RowChars <= 0 Then
            m_uConfig.RowChars = DEF_ROW_CHARS
        End If
    End If
    pvRowChars = m_uConfig.RowChars
    Exit Property
EH:
    RaiseError FUNC_NAME
End Property

Private Property Get pvItemChars() As Long
    pvItemChars = m_uConfig.ItemChars
    If pvItemChars = 0 Then
        pvItemChars = LimitLong(pvRowChars - 5, , 30)
    End If
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
    Dim baText()        As Byte
    
    On Error GoTo EH
    sProbeText = GetConfigValue(STR_PROTOCOL_DATECS_FP, "ProbeText", STR_CHR1 & Chr$(&H24) & Chr$(&H20) & Chr$(&H20) & Chr$(5) & Chr$(&H30) & Chr$(&H30) & Chr$(&H30) & Chr$(&H30) & Chr$(3))
    sReplyText = GetConfigValue(STR_PROTOCOL_DATECS_FP, "ReplyText", Chr$(&H15))  '--- NAK (&H15)
    lTimeout = GetConfigNumber(STR_PROTOCOL_DATECS_FP, "AutodetectTimeout", DEF_AUTODETECTTIMEOUT)
    lDeviceInfoTimeout = GetConfigNumber(STR_PROTOCOL_DATECS_FP, "DeviceInfoTimeout", DEF_DEVICEINFOTIMEOUT)
    AssignVariant vSpeeds, GetConfigValue(STR_PROTOCOL_DATECS_FP, "AutodetectSpeeds", Split(DEF_AUTODETECTSPEEDS, "|"))
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
                    Set .Connector = New cSerialPortConnector
                    If .Connector.Init(At(vPorts, lIdx) & "," & vSpeed) Then
                        .WriteOver.hEvent = CreateEvent(0, 1, 0, 0)
                        .ReadOver.hEvent = CreateEvent(0, 1, 0, 0)
                    Else
                        Set .Connector = Nothing
                    End If
                End If
            End With
        Next
        For lIdx = 0 To UBound(aPorts)
            With aPorts(lIdx)
                If Not .Connector Is Nothing Then
                    Call PurgeComm(.Connector.hComm, PURGE_TXCLEAR Or PURGE_RXCLEAR)
                    baText = ToAscii(sProbeText)
                    lResult = WriteFile(.Connector.hComm, baText(0), Len(sProbeText), lActual, .WriteOver)
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
                If Not .Connector Is Nothing Then
                    .Buffer = String$(1000, 0)
                    lActual = 0
                    lResult = ReadFile(.Connector.hComm, ByVal StrPtr(.Buffer), LenB(.Buffer), lActual, .ReadOver)
                    If lResult = 0 And Err.LastDllError = ERROR_IO_PENDING Then
                        hEvents(lEventsCount) = .ReadOver.hEvent
                        lEventsCount = lEventsCount + 1
                    Else
                        If lActual = 0 Then
                            Call GetOverlappedResult(.Connector.hComm, .ReadOver, lActual, 0)
                        End If
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
                If Not .Connector Is Nothing Then
                    If InStr(.Buffer, vbNullChar) > 0 Then
                        .Buffer = StrConv(.Buffer, vbUnicode)
                        .Buffer = Left$(.Buffer, InStr(.Buffer, vbNullChar) - 1)
                        '--- remove SYN (&H16)
                        .Buffer = Replace(.Buffer, Chr$(&H16), vbNullString)
                    End If
                    If LenB(.Buffer) <> 0 Then
'                        Debug.Print lIdx, ToHexDump(.Buffer), Timer
                        If Right$(.Buffer, Len(sReplyText)) = sReplyText Then
                            If pvGetDeviceInfo(.Connector, lDeviceInfoTimeout, Empty, sProtocol, sModel) Then
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
                Set .Connector = Nothing
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
            Optional ByVal RowChars As Long) As Boolean
    Const FUNC_NAME     As String = "Init"
    Const CONF_NAME     As String = FUNC_NAME
    Dim vSplit          As Variant
    Dim sSerial         As String
    Dim lIdx            As Long
    Dim eCmd            As UcsIslCommandsEnum
    Dim sParam          As String
    Dim eHasSwitchesInfo As VbTriState
    Dim sDaisyPrefix    As String
    Dim sIncotexPrefix  As String
    Dim lSerialIndex    As Long
    Dim lSwitchesIndex  As Long
    Dim sError          As String
    
    On Error GoTo EH
    DebugLog FUNC_NAME, "sDevice=" & sDevice & ", Timeout=" & Timeout
    pvSetLastError vbNullString
    m_lTimeout = IIf(Timeout <= 0, DEF_TIMEOUT, Timeout)
    vSplit = Split(sDevice, ",")
    m_sDevice = Trim$(At(vSplit, 0, "COM1")) & "," & C_Lng(At(vSplit, 1, "9600")) & "," & At(vSplit, 2) & "," & C_Lng(At(vSplit, 3, "8")) & "," & _
        IIf(UCase$(At(vSplit, 4, "N")) = "Y", "Y", "N") & "," & C_Lng(At(vSplit, 5, "1"))
    If Not Connect() Then
        GoTo QH
    End If
    pvGetCommandConfig CONF_NAME, "InfoDiagnostics", eCmd, ucsFpcInfoDiagnostics, sParam, "0", _
        "DaisyPrefix", sDaisyPrefix, "DY", _
        "IncotexPrefix", sIncotexPrefix, "IN", _
        "SerialIndex", lSerialIndex, 4, _
        "SwitchesIndex", lSwitchesIndex, 3
    m_vInfoDiagnostics = Split(pvSendCommand(m_oConnector, m_lTimeout, eCmd, sParam, sError), DEF_DELIM)
    If LenB(sError) <> 0 Then
        pvSetLastError sError
        GoTo QH
    End If
    '--- load config values
    sSerial = At(m_vInfoDiagnostics, lSerialIndex)
    eHasSwitchesInfo = GetConfigNumber(sSerial, "HasSwitchesInfo", vbUseDefault)
    If eHasSwitchesInfo = vbUseDefault Then
        '--- heuristic check for recent version of firmware
        eHasSwitchesInfo = IIf(Len(At(m_vInfoDiagnostics, lSwitchesIndex)) > 2, vbTrue, vbFalse)
    End If
    With m_uConfig
        .Charset = ucsCrsAuto
        .IsEcr = vbUseDefault
        .IsDaisy = (Left$(GetDeviceSerialNo(), Len(sDaisyPrefix)) = sDaisyPrefix) And LenB(sIncotexPrefix) <> 0
        .IsIncotex = (Left$(GetDeviceSerialNo(), Len(sIncotexPrefix)) = sIncotexPrefix) And LenB(sIncotexPrefix) <> 0
        .AbsoluteDiscount = C_Bool(GetConfigValue(sSerial, "AbsoluteDiscount", eHasSwitchesInfo = vbTrue))
        .NegativePrices = C_Bool(GetConfigValue(sSerial, "NegativePrices", eHasSwitchesInfo = vbFalse))
        .RowChars = GetConfigNumber(sSerial, "RowChars", RowChars)
        .ItemChars = GetConfigNumber(sSerial, "ItemChars", 0)
        .MaxReceiptRows = GetConfigNumber(sSerial, "MaxReceiptRows", DEF_MAX_RECEIPT_ROWS)
        .MinDiscount = GetConfigNumber(sSerial, "MinDiscount", DEF_MIN_DISCOUNT)
        .MaxDiscount = GetConfigNumber(sSerial, "MaxDiscount", DEF_MAX_DISCOUNT)
        .MaxPaymentLen = GetConfigNumber(sSerial, "MaxPaymentLen", DEF_MAX_PAYMENT_LEN)
        vSplit = Split(DEF_MAX_CGINFO, "|")
        For lIdx = LBound(.MaxCgInfo) To UBound(.MaxCgInfo)
            .MaxCgInfo(lIdx) = GetConfigNumber(sSerial, "MaxCgInfo" & lIdx, C_Lng(vSplit(lIdx)))
        Next
        .MaxOperatorName = GetConfigNumber(sSerial, "MaxOperatorName", DEF_MAX_OPERATOR_NAME)
        .CommEventTimeout = GetConfigNumber(sSerial, "CommEventTimeout", DEF_COMM_EVENT_TIMEOUT)
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
    DebugLog FUNC_NAME, "m_sDevice=" & m_sDevice & ", m_lTimeout=" & m_lTimeout
    Set m_oConnector = InitDeviceConnector(m_sDevice, m_lTimeout, LocalizedText(ucsFscLciConnectorErrors), m_sLastError)
    If m_oConnector Is Nothing Then
        DebugLog FUNC_NAME, "InitDeviceConnector failed, m_sLastError=" & m_sLastError
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
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function SendCommand(ByVal eCmd As UcsIslCommandsEnum, Optional Data As String) As String
    Const FUNC_NAME     As String = "SendCommand"
    Const CONF_NAME     As String = FUNC_NAME
    Dim vResult         As Variant
    Dim lChar           As Long
    Dim eDecimalsCmd    As UcsIslCommandsEnum
    Dim sDecimalsParam  As String
    Dim lDecimalsIndex  As Long
    Dim sError          As String
    
    On Error GoTo EH
    pvSetLastError vbNullString
    If eCmd = ucsFpcNoop Then
        Exit Function
    End If
    '--- auto-sense encoding
    If m_uConfig.Charset = ucsCrsAuto Then
        pvGetCommandConfig CONF_NAME, "InitDecimals", eDecimalsCmd, ucsFpcInitDecimals, sDecimalsParam, vbNullString, _
            "Index", lDecimalsIndex, 2
        vResult = Split(pvSendCommand(m_oConnector, m_lTimeout, eDecimalsCmd, sDecimalsParam, sError), DEF_DELIM)
        If LenB(sError) <> 0 Then
            pvSetLastError sError
            GoTo QH
        End If
        '--- first char of currency name
        lChar = Asc(Zn(Trim$(At(vResult, lDecimalsIndex)), " "))
        m_uConfig.Charset = IIf(lChar >= MIK_OFFSET And lChar < MIK_OFFSET + &H40, ucsCrsMik, ucsCrsWindows)
    End If
    SendCommand = pvSendCommand(m_oConnector, m_lTimeout, eCmd, Data, sError)
    RaiseEvent CommandComplete(eCmd, Data, SendCommand)
    If LenB(sError) <> 0 Then
        pvSetLastError sError
        GoTo QH
    End If
    If Status(ucsStbPrintingError) Then
        pvSetLastError Zn(ErrorText, Printf(pvInternal(ucsErrUnknownError), "0x" & Hex$(m_eStatus)))
        GoTo QH
    End If
    If Memory(ucsFmeMemoryError) Then
        pvSetLastError Zn(MemoryText, Printf(pvInternal(ucsErrUnknownMemory), "0x" & Hex$(m_eMemory)))
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
            Optional UniqueSaleNo As String, _
            Optional InvDocNo As String, _
            Optional InvCgTaxNo As String, _
            Optional ByVal InvCgTaxNoType As UcsFiscalTaxNoTypeEnum, _
            Optional InvCgVatNo As String, _
            Optional InvCgName As String, _
            Optional InvCgCity As String, _
            Optional InvCgAddress As String, _
            Optional InvCgPrsReceive As String, _
            Optional ByVal StornoType As UcsFiscalStornoTypeEnum, _
            Optional StornoReceiptNo As String, _
            Optional StornoReceiptDate As Date, _
            Optional StornoFiscalMemoryNo As String, _
            Optional StornoInvoiceNo As String, _
            Optional StornoReason As String) As Boolean
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
        .InitUniqueSaleNo = SafeText(UniqueSaleNo)
        SplitCgAddress Trim$(SafeText(InvCgCity)) & vbCrLf & Trim$(SafeText(InvCgAddress)), sCity, sAddress, pvRowChars - 2
        .InitInvData = Array(SafeText(InvDocNo), SafeText(InvCgTaxNo), SafeText(InvCgVatNo), _
            SafeText(InvCgName), sCity, sAddress, SafeText(InvCgPrsReceive), InvCgTaxNoType)
        .InitStornoData = Array(IIf(.InitReceiptType = ucsFscRetCreditNote, ucsFscRsnTaxBaseReduction, StornoType), _
            SafeText(StornoReceiptNo), StornoReceiptDate, SafeText(StornoFiscalMemoryNo), SafeText(StornoInvoiceNo), SafeText(StornoReason))
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
        Select Case m_uRow(0).InitReceiptType
        Case ucsFscRetStorno, ucsFscRetCreditNote
            bNegative = Not bNegative
        End Select
        .PluPrice = IIf(bNegative, -1, 1) * Round(Abs(Price), 2)
        .PluQuantity = Round(IIf(bNegative Or m_uConfig.NegativePrices, Abs(Quantity), Quantity), 3)
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
            ByVal Value As Double, _
            Optional ByVal BeforeIndex As Long) As Boolean
    Const FUNC_NAME     As String = "AddDiscount"
    Dim uRow            As UcsRowData
    Dim lIdx            As Long
    
    On Error GoTo EH
    '--- sanity check
    If m_lRowCount = 0 Then
        pvSetLastError pvInternal(ucsErrNoReceiptStarted)
        GoTo QH
    End If
    Select Case DiscType
    Case ucsFscDstTotal
        AddPLU Printf(IIf(Value > DBL_EPSILON, pvText(ucsTxtSurcharge), pvText(ucsTxtDiscount)), vbNullString), Value, BeforeIndex:=BeforeIndex
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
        '--- custom payment types: 5, 6, 7 & 8
        Number = 4 - Number
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
    vSplit = Split(sResumeTicket, STR_CHR1)
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
    If m_uRow(0).InitReceiptType = ucsFscRetSale Or m_uRow(0).InitReceiptType = ucsFscRetInvoice Then
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
    Case ucsFscRetSale, ucsFscRetInvoice
        With m_uCtxPrint
            If .PmtPrinted Then
                For lIdx = LBound(.GrpTotal) To UBound(.GrpTotal)
                    GetResumeTicket = GetResumeTicket & .GrpTotal(lIdx) & STR_CHR1
                Next
                GetResumeTicket = GetResumeTicket & .Paid & STR_CHR1 & .PluCount & STR_CHR1 & -.PmtPrinted & STR_CHR1 & -.ChangePrinted & STR_CHR1 & .Row
            End If
        End With
    End Select
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function CopyLastReceipt(Optional ReceiptNumber As String) As Boolean
    Const CONF_NAME     As String = "CopyLastReceipt"
    Dim eCmd            As UcsIslCommandsEnum
    Dim sParam          As String
    Dim sFormat         As String
    
    On Error GoTo EH_Cancel
    If LenB(ReceiptNumber) = 0 Then
        pvGetCommandConfig CONF_NAME, "PrintDuplicate", eCmd, ucsFpcPrintDuplicate, sParam, "1"
        SendCommand eCmd, sParam
    Else
        sParam = "P"
        pvGetCommandConfig CONF_NAME, "InitEcTape" & sParam, eCmd, ucsFpcInitEcTape, sParam, sParam, _
            "Format", sFormat, "%1,%2"
        SendCommand eCmd, Printf(sFormat, sParam, ReceiptNumber)
    End If
    '--- success
    CopyLastReceipt = True
EH_Cancel:
End Function

Public Function GetClock() As Date
    Const FUNC_NAME     As String = "GetClock"
    Const CONF_NAME     As String = FUNC_NAME
    Dim bPrevExceptions As Boolean
    Dim eCmd            As UcsIslCommandsEnum
    Dim sParam          As String
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoDateTime", eCmd, ucsFpcInfoDateTime, sParam, vbNullString
    bPrevExceptions = pvSetExceptions(False)
    GetClock = C_Date(SendCommand(eCmd, sParam))
    Exceptions = bPrevExceptions
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
    pvGetCommandConfig CONF_NAME, "InitDateTime", eCmd, ucsFpcInitDateTime, sParam, "%1", _
        "Format", sFormat, FORMAT_DATE_EXTENDED
    SendCommand eCmd, Printf(sParam, Format$(NewDate, sFormat))
    '--- success
    SetClock = True
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function RunXReport(ByVal ReportType As UcsFiscalReportsTypeEnum) As Boolean
    Const FUNC_NAME     As String = "RunXReport"
    Const CONF_NAME     As String = FUNC_NAME
    Dim bPrevExceptions As Boolean
    Dim eCmd            As UcsIslCommandsEnum
    Dim sParam          As String
    
    On Error GoTo EH
    If Not CancelReceipt() Then
        pvSetLastError pvInternal(ucsErrCannotCancelReceipt)
        GoTo QH
    End If
    Select Case ReportType
    Case ucsFscRstDaily
        pvGetCommandConfig CONF_NAME, "PrintDailyReport", eCmd, ucsFpcPrintDailyReport, sParam, "2"
    Case ucsFscRstDailyByItems
        pvGetCommandConfig CONF_NAME, "PrintDailyReportItems", eCmd, ucsFpcPrintDailyReportItems, sParam, "2"
    Case ucsFscRstDailyByDepartment
        If IsDaisy Then
            pvGetCommandConfig CONF_NAME, "ExtendedPrintDailyReport", eCmd, ucsFpcPrintDailyReport, sParam, "9"
        Else
            pvGetCommandConfig CONF_NAME, "PrintDailyReportDepartments", eCmd, ucsFpcPrintDailyReportDepartments, sParam, "2"
        End If
    Case ucsFscRstDailyByItemsAndDepartment
        If IsDaisy Then
            pvGetCommandConfig CONF_NAME, "ExtendedPrintDailyReportItems", eCmd, ucsFpcPrintDailyReportItems, sParam, "9"
        Else
            pvGetCommandConfig CONF_NAME, "PrintDailyReportItemsDepartments", eCmd, ucsFpcPrintDailyReportItemsDepartments, sParam, "2"
        End If
    Case Else
        pvSetLastError pvInternal(ucsErrInvalidReportType)
        GoTo QH
    End Select
    bPrevExceptions = pvSetExceptions(False)
    SendCommand eCmd, sParam
    Exceptions = bPrevExceptions
    '--- success
    RunXReport = True
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function RunZReport(ByVal ReportType As UcsFiscalReportsTypeEnum) As Boolean
    Const FUNC_NAME     As String = "RunZReport"
    Const CONF_NAME     As String = FUNC_NAME
    Dim bPrevExceptions As Boolean
    Dim lIdx            As Long
    Dim eCmd            As UcsIslCommandsEnum
    Dim sParam          As String
    Dim bEcClear        As Boolean
    Dim eEcCmd          As UcsIslCommandsEnum
    Dim sEcParam        As String
    Dim lEcIndex        As Long
    Dim lEcCount        As Long
    
    On Error GoTo EH
    If Not CancelReceipt() Then
        pvSetLastError pvInternal(ucsErrCannotCancelReceipt)
        GoTo QH
    End If
    Select Case ReportType
    Case ucsFscRstDaily
        pvGetCommandConfig CONF_NAME, "PrintDailyReport", eCmd, ucsFpcPrintDailyReport, sParam, "0", _
            "EcClear", bEcClear, False
    Case ucsFscRstDailyByItems
        pvGetCommandConfig CONF_NAME, "PrintDailyReportItems", eCmd, ucsFpcPrintDailyReportItems, sParam, "0", _
            "EcClear", bEcClear, False
    Case ucsFscRstDailyByDepartment
        If IsDaisy Then
            pvGetCommandConfig CONF_NAME, "ExtendedPrintDailyReport", eCmd, ucsFpcPrintDailyReport, sParam, "8", _
                "EcClear", bEcClear, False
        Else
            pvGetCommandConfig CONF_NAME, "PrintDailyReportDepartments", eCmd, ucsFpcPrintDailyReportDepartments, sParam, "0", _
                "EcClear", bEcClear, False
        End If
    Case ucsFscRstDailyByItemsAndDepartment
        If IsDaisy Then
            pvGetCommandConfig CONF_NAME, "ExtendedPrintDailyReportItems", eCmd, ucsFpcPrintDailyReportItems, sParam, "8", _
                "EcClear", bEcClear, False
        Else
            pvGetCommandConfig CONF_NAME, "PrintDailyReportItemsDepartments", eCmd, ucsFpcPrintDailyReportItemsDepartments, sParam, "0", _
                "EcClear", bEcClear, False
        End If
    Case Else
        pvSetLastError pvInternal(ucsErrInvalidReportType)
        GoTo QH
    End Select
    bPrevExceptions = pvSetExceptions(False)
    If bEcClear Then
        '--- clear electronic control tape (if present)
        sEcParam = "I"
        pvGetCommandConfig CONF_NAME, "InitEcTape" & sEcParam, eEcCmd, ucsFpcInitEcTape, sEcParam, sEcParam, _
            "Index", lEcIndex, 1
        lEcCount = C_Lng(At(Split(SendCommand(eEcCmd, sEcParam), DEF_DELIM), lEcIndex))
        If Not Status(ucsStbPrintingError) Then
            For lIdx = 1 To lEcCount
                sEcParam = IIf(lIdx = 1, "PS", "CS")
                pvGetCommandConfig CONF_NAME, "InitEcTape" & sEcParam, eEcCmd, ucsFpcInitEcTape, sEcParam, sEcParam
                SendCommand eEcCmd, sEcParam
            Next
            sEcParam = "E"
            pvGetCommandConfig CONF_NAME, "InitEcTape" & sEcParam, eEcCmd, ucsFpcInitEcTape, sEcParam, sEcParam
            SendCommand eEcCmd, sEcParam
        End If
    End If
    SendCommand eCmd, sParam
    Exceptions = bPrevExceptions
    If IsDaisy Then
        pvGetCommandConfig CONF_NAME, "ExtendedOperatorsReset", eCmd, ucsFpcInitOperatorReset, sParam, "0,9999"
        SendCommand eCmd, sParam
    End If
    '--- success
    RunZReport = True
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function RunPeriodReport(ByVal ReportType As UcsFiscalReportsTypeEnum, ByVal StartDate As Date, ByVal EndDate As Date) As Boolean
    Const FUNC_NAME     As String = "RunPeriodReport"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsIslCommandsEnum
    Dim sFormat         As String
    Dim sParam          As String
    
    On Error GoTo EH
    If Not CancelReceipt() Then
        pvSetLastError pvInternal(ucsErrCannotCancelReceipt)
        GoTo QH
    End If
    Select Case ReportType
    Case ucsFscRstPeriodShort
        pvGetCommandConfig CONF_NAME, "PrintReportByDateShort", eCmd, ucsFpcPrintReportByDateShort, sParam, "%1,%2", _
            "Format", sFormat, "ddmmyy"
    Case ucsFscRstPeriodDetailed
        pvGetCommandConfig CONF_NAME, "PrintReportByDateDetailed", eCmd, ucsFpcPrintReportByDateDetailed, sParam, "%1,%2", _
            "Format", sFormat, "ddmmyy"
    Case Else
        pvSetLastError pvInternal(ucsErrInvalidReportType)
        GoTo QH
    End Select
    sParam = Printf(sParam, Format$(StartDate, sFormat), Format$(EndDate, sFormat))
    SendCommand eCmd, sParam
    '--- success
    RunPeriodReport = True
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetLastReceiptNo() As String
    Const FUNC_NAME     As String = "GetLastReceiptNo"
    Const CONF_NAME     As String = FUNC_NAME
    Dim bPrevExceptions As Boolean
    Dim eCmd            As UcsIslCommandsEnum
    Dim sParam          As String
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoLastReceipt", eCmd, ucsFpcInfoLastReceipt, sParam, vbNullString
    bPrevExceptions = pvSetExceptions(False)
    GetLastReceiptNo = SendCommand(eCmd, sParam)
    Exceptions = bPrevExceptions
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetLastInvoiceNo() As String
    Const FUNC_NAME     As String = "GetLastInvoiceNo"
    Const CONF_NAME     As String = FUNC_NAME
    Dim bPrevExceptions As Boolean
    Dim eCmd            As UcsIslCommandsEnum
    Dim sParam          As String
    Dim lIndex          As Long
    Dim lOffset         As Long
    
    On Error GoTo EH
    If IsDaisy Then
        pvGetCommandConfig CONF_NAME, "ExtendedInitSetting", eCmd, ucsFpcExtendedInitSetting, sParam, "R17", _
            "Index", lIndex, 1
    ElseIf IsIncotex Then
        pvGetCommandConfig CONF_NAME, "InfoSums", eCmd, ucsFpcInfoSums, sParam, vbNullString, _
            "Index", lIndex, 10, _
            "Offset", lOffset, -1
    Else
        pvGetCommandConfig CONF_NAME, "InitInvoiceNo", eCmd, ucsFpcInitInvoiceNo, sParam, vbNullString, _
            "Index", lIndex, 0
    End If
    bPrevExceptions = pvSetExceptions(False)
    GetLastInvoiceNo = At(Split(SendCommand(eCmd, sParam), DEF_DELIM), lIndex)
    Exceptions = bPrevExceptions
    If lOffset <> 0 And IsNumeric(GetLastInvoiceNo) Then
        GetLastInvoiceNo = Val(GetLastInvoiceNo) + lOffset
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetDeviceSerialNo() As String
    Const FUNC_NAME     As String = "GetDeviceSerialNo"
    Const CONF_NAME     As String = FUNC_NAME
    Dim lIndex          As Long
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoDiagnostics", 0, 0, vbNullString, vbNullString, _
        "Index", lIndex, 4
    GetDeviceSerialNo = At(m_vInfoDiagnostics, lIndex)
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetFiscalMemoryNo() As String
    Const FUNC_NAME     As String = "GetFiscalMemoryNo"
    Const CONF_NAME     As String = FUNC_NAME
    Dim lIndex          As Long
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoDiagnostics", 0, 0, vbNullString, vbNullString, _
        "Index", lIndex, 5
    GetFiscalMemoryNo = At(m_vInfoDiagnostics, lIndex)
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetTaxNo() As String
    Const FUNC_NAME     As String = "GetTaxNo"
    Const CONF_NAME     As String = FUNC_NAME
    Dim bPrevExceptions As Boolean
    Dim eCmd            As UcsIslCommandsEnum
    Dim sParam          As String
    Dim lIndex          As Long
    Dim vSplit          As Variant
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoBulstat", eCmd, ucsFpcInfoBulstat, sParam, vbNullString, _
        "Index", lIndex, IIf(IsIncotex, -1, 0)
    bPrevExceptions = pvSetExceptions(False)
    vSplit = Split(SendCommand(eCmd, sParam), DEF_DELIM)
    GetTaxNo = At(vSplit, lIndex)
    Exceptions = bPrevExceptions
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetTaxCaption() As String
    Const FUNC_NAME     As String = "GetTaxCaption"
    Const CONF_NAME     As String = FUNC_NAME
    Dim bPrevExceptions As Boolean
    Dim eCmd            As UcsIslCommandsEnum
    Dim sParam          As String
    Dim lIndex          As Long
    
    On Error GoTo EH
    If IsIncotex Then
        pvGetCommandConfig CONF_NAME, "ExtendedInitText", eCmd, ucsFpcExtendedInitText, sParam, "R14", _
            "Index", lIndex, 1
    Else
        pvGetCommandConfig CONF_NAME, "InfoBulstat", eCmd, ucsFpcInfoBulstat, sParam, vbNullString, _
            "Index", lIndex, 1
    End If
    bPrevExceptions = pvSetExceptions(False)
    GetTaxCaption = At(Split(SendCommand(eCmd, sParam), DEF_DELIM), lIndex)
    Exceptions = bPrevExceptions
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
    Dim bPrevExceptions As Boolean
    Dim eCmd            As UcsIslCommandsEnum
    Dim sParam          As String
    
    On Error GoTo EH
    If Index >= 1 And Index <= 6 Then
        pvGetCommandConfig CONF_NAME, "InitHeaderFooter" & Index, eCmd, ucsFpcInitHeaderFooter, sParam, "I" & (Index - 1)
        bPrevExceptions = pvSetExceptions(False)
        GetHeaderText = pvStripBold(SendCommand(eCmd, sParam))
        Exceptions = bPrevExceptions
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetFooterText(ByVal Index As Long) As String
    Const FUNC_NAME     As String = "GetFooterText"
    Const CONF_NAME     As String = FUNC_NAME
    Dim bPrevExceptions As Boolean
    Dim eCmd            As UcsIslCommandsEnum
    Dim sParam          As String
    
    On Error GoTo EH
    If Index >= 1 And Index <= 2 Then
        pvGetCommandConfig CONF_NAME, "InitHeaderFooter" & Index, eCmd, ucsFpcInitHeaderFooter, sParam, "I" & (Index + 5)
        bPrevExceptions = pvSetExceptions(False)
        GetFooterText = pvStripBold(SendCommand(eCmd, sParam))
        Exceptions = bPrevExceptions
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetPaymentName(ByVal Index As Long) As String
    Const FUNC_NAME     As String = "GetPaymentName"
    Const CONF_NAME     As String = FUNC_NAME
    Dim bPrevExceptions As Boolean
    Dim eCmd            As UcsIslCommandsEnum
    Dim sParam          As String
    Dim lIndex          As Long
    
    On Error GoTo EH
    If Index < 0 Then
        '--- custom payment types: 5, 6, 7 & 8
        Index = 4 - Index
    End If
    If Index > 4 Then
        bPrevExceptions = pvSetExceptions(False)
        If IsDaisy Then
            pvGetCommandConfig CONF_NAME, "ExtendedInitText" & Index, eCmd, ucsFpcExtendedInitText, sParam, "R" & (58 + Index)
            GetPaymentName = Trim$(pvStripBold(SendCommand(eCmd, sParam)))
            GoTo QH
        ElseIf IsIncotex Then
            If Index >= 7 Then
                GoTo QH
            End If
            pvGetCommandConfig CONF_NAME, "ExtendedInitText" & Index, eCmd, ucsFpcExtendedInitText, sParam, "R" & (7 + Index), _
                "Index", lIndex, -1
        Else
            pvGetCommandConfig CONF_NAME, "InitPaymentType" & Index, eCmd, ucsFpcInitPaymentType, sParam, Chr$(68 + Index)
        End If
        GetPaymentName = Trim$(pvStripBold(At(Split(SendCommand(eCmd, sParam), DEF_DELIM), lIndex)))
        Exceptions = bPrevExceptions
        '--- check for failure (returns 'F')
        If Status(ucsStbPrintingError) And GetPaymentName = "F" Then
            GetPaymentName = vbNullString
        End If
    Else
        GetPaymentName = pvText(ucsTxtPaymentCash + Index - 1)
    End If
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function OpenDrawer() As Boolean
    Const FUNC_NAME     As String = "OpenDrawer"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsIslCommandsEnum
    Dim sParam          As String
    
    On Error GoTo EH_Cancel
    pvGetCommandConfig CONF_NAME, "InitOpenDrawer", eCmd, ucsFpcInitOpenDrawer, sParam, vbNullString
    SendCommand eCmd, sParam
    '--- success
    OpenDrawer = True
EH_Cancel:
End Function

Public Function CancelReceipt() As Boolean
    Const FUNC_NAME     As String = "CancelReceipt"
    Const CONF_NAME     As String = FUNC_NAME
    Dim bPrevExceptions As Boolean
    Dim eTransCmd       As UcsIslCommandsEnum
    Dim sTransParam     As String
    Dim lTransIndex     As Long
    Dim lTransStatus    As Long
    Dim eCmd            As UcsIslCommandsEnum
    Dim sParam          As String
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoTransaction", eTransCmd, ucsFpcInfoTransaction, sTransParam, vbNullString, _
        "Index", lTransIndex, 0, _
        "Status", lTransStatus, 0
    bPrevExceptions = pvSetExceptions(False)
    If C_Lng(At(Split(SendCommand(eTransCmd, sTransParam), DEF_DELIM), lTransIndex)) <> lTransStatus Then
        If Status(ucsStbFiscalPrinting) Then
            If IsDaisy Or IsIncotex Then
                pvGetCommandConfig CONF_NAME, "ExtendedFiscalCancel", eCmd, ucsFpcExtendedFiscalCancel, sParam, vbNullString
            Else
                '--- note: FP-3530 can cancel always, FP-550 sometimes can not
                pvGetCommandConfig CONF_NAME, "FiscalCancel", eCmd, ucsFpcFiscalCancel, sParam, vbNullString
            End If
            SendCommand eCmd, sParam
            If Status(ucsStbFiscalPrinting) Then
                '--- note: this helps if sum payed is less than total receipt sum
                pvGetCommandConfig CONF_NAME, "FiscalPayment", eCmd, ucsFpcFiscalPayment, sParam, vbTab
                SendCommand eCmd, sParam
                '--- note: when printing invoice, if no contragent info set then cancel fails!
                pvGetCommandConfig CONF_NAME, "FiscalCgInfo", eCmd, ucsFpcFiscalCgInfo, sParam, "0000000000" & IIf(IsEcr, vbTab & "0", vbNullString)
                SendCommand eCmd, sParam
                '--- for FP-550
                pvGetCommandConfig CONF_NAME, "FiscalClose", eCmd, ucsFpcFiscalClose, sParam, vbNullString
                SendCommand eCmd, sParam
            End If
        Else
            pvGetCommandConfig CONF_NAME, "NonFiscalClose", eCmd, ucsFpcNonFiscalClose, sParam, vbNullString
            SendCommand eCmd, sParam
        End If
        '--- success (or failure)
        CancelReceipt = C_Lng(At(Split(SendCommand(eTransCmd, sTransParam), DEF_DELIM), lTransIndex)) = lTransStatus
    Else
        CancelReceipt = True
    End If
    Exceptions = bPrevExceptions
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetDeviceModel() As String
    Const FUNC_NAME     As String = "GetDeviceModel"
    
    On Error GoTo EH
    pvGetDeviceInfo m_oConnector, m_lTimeout, m_vInfoDiagnostics, vbNullString, GetDeviceModel
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetDeviceProtocol() As String
    Const FUNC_NAME     As String = "GetDeviceProtocol"
    
    On Error GoTo EH
    pvGetDeviceInfo m_oConnector, m_lTimeout, m_vInfoDiagnostics, GetDeviceProtocol, vbNullString
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetDeviceStatus(CurrentStatus As String) As Boolean
    Const FUNC_NAME     As String = "GetDeviceStatus"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsIslCommandsEnum
    Dim sParam          As String
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoStatus", eCmd, ucsFpcInfoStatus, sParam, "X"
    pvSendCommand m_oConnector, m_lTimeout, eCmd, sParam, m_sLastError
    If LenB(m_sLastError) <> 0 Then
        '--- do nothing
    ElseIf Status(ucsStbPrintingError) Then
        pvSetLastError Zn(ErrorText, Printf(pvInternal(ucsErrUnknownError), "0x" & Hex$(m_eStatus))), DontRaise:=True
    ElseIf Memory(ucsFmeMemoryError) Then
        pvSetLastError Zn(MemoryText, Printf(pvInternal(ucsErrUnknownMemory), "0x" & Hex$(m_eMemory))), DontRaise:=True
    End If
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
    Dim bPrevExceptions As Boolean
    Dim vSplit          As Variant
    Dim vRetVal         As Variant
    Dim lIdx            As Long
    Dim lJdx            As Long
    Dim eCmd            As UcsIslCommandsEnum
    Dim sParam          As String
    Dim vIdxMap          As Variant
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoTotalsByPayments", eCmd, ucsFpcInfoTotalsByPayments, sParam, IIf(IsDaisy, "A", vbNullString)
    bPrevExceptions = pvSetExceptions(False)
    vSplit = Split(SendCommand(eCmd, sParam), DEF_DELIM)
    Exceptions = bPrevExceptions
    If UBound(vSplit) > 0 Then
        '--- WTF?!#@ happened w/ these indexes!
        If IsDaisy Then
            vIdxMap = Split("0|2|1|4|5|6|7|-1", "|")
        ElseIf IsIncotex Then
            vIdxMap = Split("0|1|2|4", "|")
        Else
            vIdxMap = Split("0|3|1|2", "|")
        End If
        ReDim vRetVal(0 To 7) As Double
        For lIdx = 0 To UBound(vRetVal)
            lJdx = At(vIdxMap, lIdx, C_Str(lIdx))
            If 0 <= lJdx And lIdx <= UBound(vSplit) - IIf(Not IsDaisy, 2, 3) Then
                vRetVal(lJdx) = ParseSum(At(vSplit, lIdx))
            End If
        Next
        GetTotalsByPayments = vRetVal
    Else
        GetTotalsByPayments = Array()
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetTotalsByVatGroups() As Variant
    Const FUNC_NAME     As String = "GetTotalsByVatGroups"
    Const CONF_NAME     As String = FUNC_NAME
    Dim bPrevExceptions As Boolean
    Dim vTotals         As Variant
    Dim vVatSums        As Variant
    Dim vTaxRates       As Variant
    Dim vRetVal         As Variant
    Dim lIdx            As Long
    Dim eCmd            As UcsIslCommandsEnum
    Dim sParam          As String
    
    On Error GoTo EH
    bPrevExceptions = pvSetExceptions(False)
    pvGetCommandConfig CONF_NAME, "InfoTotalsByVatGroups0", eCmd, ucsFpcInfoTotalsByVatGroups, sParam, IIf(IsDaisy, "T", vbNullString)
    vTotals = Split(SendCommand(eCmd, sParam), DEF_DELIM)
    pvGetCommandConfig CONF_NAME, "InfoTotalsByVatGroups1", eCmd, ucsFpcInfoTotalsByVatGroups, sParam, IIf(IsIncotex, vbNullString, "1")
    vVatSums = Split(SendCommand(eCmd, sParam), DEF_DELIM)
    pvGetCommandConfig CONF_NAME, "InfoTaxRates", eCmd, ucsFpcInfoTaxRates, sParam, vbNullString
    vTaxRates = Split(Replace(SendCommand(eCmd, sParam), "%", vbNullString), DEF_DELIM)
    Exceptions = bPrevExceptions
    If UBound(vTotals) > 0 Or UBound(vVatSums) > 0 Then
        ReDim vRetVal(0 To 7) As Variant
        For lIdx = 0 To UBound(vRetVal)
            If UBound(vVatSums) >= lIdx + 1 Then
                vRetVal(lIdx) = Array(ParseSum(At(vTotals, lIdx + 1, At(vVatSums, lIdx + 1))), C_Dbl(At(vTaxRates, lIdx)), ParseSum(At(vVatSums, lIdx + 1)))
            Else
                vRetVal(lIdx) = Array(ParseSum(At(vTotals, lIdx + 1)), C_Dbl(At(vTaxRates, lIdx)))
            End If
        Next
        GetTotalsByVatGroups = vRetVal
    Else
        GetTotalsByVatGroups = Array()
    End If
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
    Dim bPrevExceptions As Boolean
    Dim lIdx            As Long
    
    On Error GoTo EH
    If Not IsEcr Then
        pvGetCommandConfig CONF_NAME, "InfoDailyStats", eCmd, ucsFpcInfoDailyStats, sParam, vbNullString, _
            "Index", lIndex, -2
        bPrevExceptions = pvSetExceptions(False)
        vSplit = Split(SendCommand(eCmd, sParam), DEF_DELIM)
        Exceptions = bPrevExceptions
        If UBound(vSplit) >= lIndex Then
            lReceipts = C_Lng(At(vSplit, lIndex))
        Else
            vSplit = Empty
        End If
    End If
    If Not IsArray(vSplit) Then
        pvGetCommandConfig CONF_NAME, "InfoOperator", eCmd, ucsFpcInfoOperator, sParam, "%1", _
            "Index", lIndex, 0
        bPrevExceptions = pvSetExceptions(False)
        lReceipts = 0
        For lIdx = 1 To 100
            vSplit = Split(SendCommand(eCmd, Printf(sParam, lIdx)), DEF_DELIM)
            If Status(ucsStbPrintingError) Then
                Exit For
            End If
            lReceipts = lReceipts + C_Lng(At(vSplit, lIndex))
        Next
        Exceptions = bPrevExceptions
    End If
    If IsIncotex Then
        pvGetCommandConfig CONF_NAME, "InfoLastFiscalBlock", eCmd, ucsFpcInfoLastFiscalBlock, sParam, vbNullString, _
            "Index", lIndex, -1
        bPrevExceptions = pvSetExceptions(False)
        vSplit = Split(SendCommand(eCmd, sParam), DEF_DELIM)
        Exceptions = bPrevExceptions
        dLastDate = pvToDate(At(vSplit, lIndex))
    Else
        pvGetCommandConfig CONF_NAME, "InfoLastFiscalDateTime", eCmd, ucsFpcInfoLastFiscalDateTime, sParam, "T"
        bPrevExceptions = pvSetExceptions(False)
        dLastDate = C_Date(SendCommand(eCmd, sParam))
        Exceptions = bPrevExceptions
    End If
    GetTotalsStats = Array(lReceipts, IIf(dLastDate = 0, Empty, dLastDate))
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function CashDebitCredit(ByVal dblValue As Double) As Variant
    Const FUNC_NAME     As String = "CashDebitCredit"
    Const CONF_NAME     As String = FUNC_NAME
    Dim vSplit          As Variant
    Dim vRetVal         As Variant
    Dim lIdx            As Long
    Dim eCmd            As UcsIslCommandsEnum
    Dim sParam          As String
    Dim sFormat         As String
    Dim bPrevExceptions As Boolean
    Dim sIndexes        As String
    Dim vElem           As Variant
    Dim lOffset         As Long
    
    On Error GoTo EH
    sFormat = IIf(Abs(dblValue) > DBL_EPSILON, "0.00", "\")
    pvGetCommandConfig CONF_NAME, "AdminCashDebitCredit" & -(Abs(dblValue) > DBL_EPSILON), eCmd, ucsFpcAdminCashDebitCredit, sParam, "%1", _
        "Format", sFormat, sFormat
    sParam = Printf(sParam, SafeFormat(dblValue, sFormat))
    bPrevExceptions = pvSetExceptions(False)
    vSplit = Split(SendCommand(eCmd, sParam), DEF_DELIM)
    Exceptions = bPrevExceptions
    If UBound(vSplit) > 0 Then
        ReDim vRetVal(0 To 2) As Double
        For lIdx = 0 To UBound(vRetVal)
            vRetVal(lIdx) = ParseSum(At(vSplit, lIdx + 1))
        Next
        '--- sParam=4 -> P,CashInN,CashInS,CashOutN,CashOutS,CashSum
        pvGetCommandConfig CONF_NAME, "InfoNap", eCmd, ucsFpcInfoNap, sParam, "4", _
            "Indexes", sIndexes, "4,1,3"
        bPrevExceptions = pvSetExceptions(False)
        vSplit = Split(SendCommand(eCmd, sParam), DEF_DELIM)
        Exceptions = bPrevExceptions
        lIdx = 0
        If UBound(vSplit) > 0 Then
            lOffset = IIf(At(vSplit, 0) = "P", 1, 0)
            For Each vElem In Split(sIndexes, DEF_DELIM)
                If lIdx >= LBound(vRetVal) And lIdx <= UBound(vRetVal) Then
                    If C_Str(C_Lng(vElem)) = C_Str(vElem) Then
                        If LenB(At(vSplit, C_Lng(vElem) + lOffset)) <> 0 Then
                            vRetVal(lIdx) = ParseSum(At(vSplit, C_Lng(vElem) + lOffset))
                        End If
                    Else
                        vRetVal(lIdx) = vElem
                    End If
                End If
                lIdx = lIdx + 1
            Next
        End If
        CashDebitCredit = vRetVal
    Else
        CashDebitCredit = EmptyDoubleArray
    End If
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
    Dim sMerged         As String
    Dim lIdx            As Long
    
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

Public Function GetDefaultPassword(OperatorCode As String) As String
    If IsEcr Or IsDaisy Or IsIncotex Then
        GetDefaultPassword = OperatorCode
    Else
        GetDefaultPassword = DEF_PASSWORD
    End If
End Function

'= private ===============================================================

Private Function pvSendCommand( _
            oConnector As IDeviceConnector, _
            ByVal lTimeout As Long, _
            ByVal eCmd As UcsIslCommandsEnum, _
            Optional Data As String, _
            Optional Error As String) As String
    Const FUNC_NAME     As String = "pvSendCommand"
    Dim sSend           As String
    Dim sText           As String
    Dim baText()        As Byte
    Dim lIdx            As Long
    Dim lSum            As Long
    Dim lSize           As Long
    Dim sBuffer         As String
    Dim dblTimer        As Double
    
    On Error GoTo EH
    DebugLog FUNC_NAME, "eCmd=" & eCmd & IIf(LenB(Data) <> 0, ", Data=" & Data, vbNullString)
    '--- prepare command and transcode cyrillic
    sSend = pvTranscode(Data, True)
    Do While eCmd >= &H100
        sSend = Chr$(eCmd Mod &H100) & sSend
        eCmd = eCmd \ &H100
    Loop
    sSend = STR_CHR1 & Chr$(&H20 + 4 + Len(sSend)) & Chr$(m_lSequence) & Chr$(eCmd) & sSend & Chr$(5)
    '--- increment sequence
    m_lSequence = IIf(m_lSequence >= &H7F, &H20, m_lSequence + 1)
    '--- calc checksum
    baText = ToAscii(sSend)
    For lIdx = 1 To UBound(baText)
        lSum = lSum + baText(lIdx)
    Next
    '--- append checksum and terminator
    sSend = sSend & Chr$((lSum \ &H1000 Mod &H10) + &H30) & Chr$((lSum \ &H100 Mod &H10) + &H30) & Chr$((lSum \ &H10 Mod &H10) + &H30) & Chr$((lSum \ &H1 Mod &H10) + &H30) & Chr$(3)
    '--- send request
    dblTimer = TimerEx + lTimeout / 1000
RetrySend:
    DebugDataDump FUNC_NAME, "send ", sSend
    baText = ToAscii(sSend)
    If Not oConnector.WriteData(baText, lTimeout) Then
        Error = oConnector.GetLastError()
        DebugLog FUNC_NAME, "WriteData failed, Error=" & Error
        GoTo QH
    End If
    '--- read response
    sText = vbNullString
    Do While Right$(sText, 1) <> Chr$(3)
        If TimerEx > dblTimer Then
            Error = pvInternal(ucsErrTimeoutWaitingForResponse)
            DebugLog FUNC_NAME, "ReadData failed, Error=" & Error
            GoTo QH
        End If
        If Not oConnector.ReadData(baText, lTimeout) Then
            Error = oConnector.GetLastError()
            DebugLog FUNC_NAME, "ReadData failed, Error=" & Error
            GoTo QH
        End If
        sBuffer = FromAscii(baText)
        DebugDataDump FUNC_NAME, "recv ", sBuffer
        '--- check for NAK (&H15)
        If InStr(sBuffer, Chr$(&H15)) > 0 Then
            DebugLog FUNC_NAME, "Will retry send on NAK received, sBuffer=0x" & ToHexDump(sBuffer)
            GoTo RetrySend
        End If
        '--- remove SYN (&H16)
        sText = sText & Replace(sBuffer, Chr$(&H16), vbNullString)
    Loop
RetryParse:
    '--- response complete w/ 3 = Terminator
    Debug.Assert Right$(sText, 1) = Chr$(3)
    If Left$(sText, 1) = STR_CHR1 Then
        lIdx = InStr(sText, Chr$(5))
        If lIdx > 2 Then
            lSize = Asc(Mid$(sText, 2, 1)) - &H20
            '--- check message length
            If lSize >= 4 And lSize < lIdx Then
                '--- check for status info
                lIdx = InStrRev(sText, Chr$(4), lIdx)
                If lIdx > 4 Then
                    pvSetStatus Mid$(sText, lIdx + 1, 6)
                    lSize = lIdx - 1
                Else
                    pvSetStatus vbNullString '--- clear
                End If
                pvSendCommand = pvTranscode(Mid$(sText, 5, lSize - 4), False)
                DebugLog FUNC_NAME, "RetVal=" & pvSendCommand & _
                    IIf(Status(ucsStbPrintingError), ", ErrorText=" & Zn(ErrorText, Printf(pvInternal(ucsErrUnknownError), "0x" & Hex$(m_eStatus))), vbNullString) & _
                    IIf(Memory(ucsFmeMemoryError), ", MemoryText=" & Zn(MemoryText, Printf(pvInternal(ucsErrUnknownMemory), "0x" & Hex$(m_eMemory))), vbNullString)
                GoTo QH
            Else
                Error = Printf(pvInternal(ucsErrInvalidResponseLength), "0x" & ToHexDump(sText))
            End If
        Else
            Error = Printf(pvInternal(ucsErrResponseMissingPostamble), "0x" & ToHexDump(sText))
        End If
    Else
        '--- check for GSM tax terminal interferences: "AT+CPIN?\r\n" or "AT+CIMI\r\n"
        lIdx = InStr(sText, STR_CHR1)
        If lIdx > 1 Then
            DebugLog FUNC_NAME, "Trim bogus symbols, Prefix=0x" & ToHexDump(Left$(sText, lIdx - 1))
            sText = Mid$(sText, lIdx)
            GoTo RetryParse
        End If
        Error = Printf(pvInternal(ucsErrResponseMissingPreamble), "0x" & ToHexDump(sText))
    End If
    pvSetStatus vbNullString '--- clear
QH:
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
    Dim lIdx            As Long
    
    On Error GoTo EH
    baStatus = ToAscii(sStatus)
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
        If (lFlags And (2 ^ lIdx)) <> 0 Then
            If Len(vTexts(lIdx)) > 1 Then
                If LenB(pvGetFlagsText) <> 0 Then
                    pvGetFlagsText = pvGetFlagsText & ", "
                End If
                pvGetFlagsText = pvGetFlagsText & vTexts(lIdx)
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

Private Sub pvPrintLine(vText As Variant)
    Const FUNC_NAME     As String = "pvPrintLine"
    Const CONF_NAME     As String = "PrintLine"
    Dim vElem               As Variant
    Dim eCmd                As UcsIslCommandsEnum
    Dim sParam              As String
    
    On Error GoTo EH
    Select Case m_uRow(0).InitReceiptType
    Case ucsFscRetSale, ucsFscRetInvoice
        pvGetCommandConfig CONF_NAME, "FiscalText", eCmd, ucsFpcFiscalText, sParam, "%1"
    Case Else
        pvGetCommandConfig CONF_NAME, "NonFiscalText", eCmd, ucsFpcNonFiscalText, sParam, "%1"
    End Select
    If IsArray(vText) Then
        For Each vElem In vText
            SendCommand eCmd, Zn(Printf(sParam, C_Str(vElem)), " ")
        Next
    Else
        SendCommand eCmd, Zn(Printf(sParam, C_Str(vText)), " ")
    End If
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Private Function pvPrintReceipt(uCtx As UcsContext, uRow() As UcsRowData, ByVal lRowCount As Long) As Boolean
    Const FUNC_NAME     As String = "pvPrintReceipt"
    Const CONF_NAME     As String = "PrintReceipt"
    Dim bPrevExceptions As Boolean
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
    Dim lStatus         As Long
    Dim lMinName        As Long
    Dim sFormatPrice    As String
    Dim sFormatQty      As String
    Dim sFormatDisc     As String
    Dim sFormatAmount   As String
    Dim sFormatDate     As String
    Dim lAmountIndex    As Long
    Dim lPaidIndex      As Long
    Dim sFormatAddr     As String
    Dim sOperParamMax   As String
    Dim sOperParamMin   As String
    Dim lEndIndex       As Long
    Dim sStornoType     As String
    Dim sTaxNoPrefix    As String
    Dim sTaxNoSuffix    As String
    
    On Error GoTo EH
    '--- ring sale
    Do While uCtx.Row < lRowCount
        With uRow(uCtx.Row)
        Select Case .RowType
        Case ucsRowInit
            pvGetCommandConfig CONF_NAME, "InfoOperator", eCmd, ucsFpcInfoOperator, sParam, "%1", _
                "Index", lIndex, 5
            sParam = Printf(sParam, .InitOperatorCode)
            bPrevExceptions = pvSetExceptions(False)
            sData = At(Split(SendCommand(eCmd, sParam), DEF_DELIM), lIndex)
            If LenB(.InitOperatorName) <> 0 Then
                If .InitOperatorName <> sData Then
                    pvGetCommandConfig CONF_NAME, "InitOperatorName", eCmd, ucsFpcInitOperatorName, sParam, "%1,%2,%3", _
                        "MinName", lMinName, 10
                    sOperParamMax = Printf(sParam, .InitOperatorCode, .InitOperatorPassword, Left$(.InitOperatorName, m_uConfig.MaxOperatorName))
                    sOperParamMin = Printf(sParam, .InitOperatorCode, .InitOperatorPassword, Left$(.InitOperatorName, LimitLong(Len(sData), lMinName)))
                    SendCommand eCmd, sOperParamMax
                    If Status(ucsStbPrintingError) And sOperParamMin <> sOperParamMax Then
                        SendCommand eCmd, sOperParamMin
                    End If
                End If
            Else
                .InitOperatorName = RTrim$(sData)
            End If
            If LenB(.InitInvData(ucsInvDocNo)) <> 0 Then
                Select Case .InitReceiptType
                Case ucsFscRetInvoice, ucsFscRetCreditNote
                    If IsDaisy Then
                        pvGetCommandConfig CONF_NAME, "ExtendedInitSetting", eCmd, ucsFpcExtendedInitSetting, sParam, "P%1,%2", _
                            "StartIndex", lIndex, 17, _
                            "EndIndex", lEndIndex, 18
                        If lIndex > 0 Then
                            SendCommand eCmd, Printf(sParam, lIndex, .InitInvData(ucsInvDocNo))
                        End If
                        If lEndIndex > 0 Then
                            SendCommand eCmd, Printf(sParam, lEndIndex, .InitInvData(ucsInvDocNo))
                        End If
                    ElseIf IsIncotex Then
                        pvGetCommandConfig CONF_NAME, "ExtendedInitInvoiceNo", eCmd, ucsFpcExtendedInitInvoiceNo, sParam, "%1" & vbLf & "%1"
                        sParam = Printf(sParam, .InitInvData(ucsInvDocNo))
                        SendCommand eCmd, sParam
                    Else
                        pvGetCommandConfig CONF_NAME, "InitInvoiceNo", eCmd, ucsFpcInitInvoiceNo, sParam, "%1,%1"
                        sParam = Printf(sParam, .InitInvData(ucsInvDocNo))
                        SendCommand eCmd, sParam
                    End If
                End Select
            End If
            If IsDaisy And False Then
                '--- ToDo: impl daisy protocol
            ElseIf IsIncotex Then
                sStornoType = Switch( _
                    .InitStornoData(ucsStoType) = ucsFscRsnTaxBaseReduction, "V", _
                    .InitStornoData(ucsStoType) = ucsFscRsnOperatorError, "R", _
                    True, "S")
                '--- %1=<OpNum>,%2=<UNPstring>,%3=<RevDocNo>[,%4=<F1>[%5=(<F2><RevInvoiceNo>,<dd-mm-yy hh:mm:ss>,origDevFMNo)]]
                Select Case .InitReceiptType
                Case ucsFscRetStorno, ucsFscRetCreditNote
                    sParam = "%1,%2,%3,%4%5"
                Case Else
                    sParam = "%1,%2,0" & IIf(.InitReceiptType = ucsFscRetInvoice, ",I", vbNullString)
                End Select
                pvGetCommandConfig CONF_NAME, "FiscalOpen" & .InitReceiptType, eCmd, ucsFpcFiscalOpen, sParam, sParam, _
                    "FormatDate", sFormatDate, FORMAT_DATE_EXTENDED
                sParam = Printf(sParam, .InitOperatorCode, _
                    Zn(.InitUniqueSaleNo, "0"), _
                    .InitStornoData(ucsStoReceiptNo), _
                    sStornoType, _
                    sStornoType & Zn(At(.InitStornoData, ucsStoInvoiceNo), "0") & _
                        "," & Format$(.InitStornoData(ucsStoReceiptDate), sFormatDate) & _
                        "," & .InitStornoData(ucsStoFiscalMemoryNo))
            ElseIf IsEcr Then
                Select Case .InitReceiptType
                Case ucsFscRetStorno, ucsFscRetCreditNote
                    '--- %1=<OpCode>,%2=<OpPwd>,%3=<NSale>,<TillNmb>,%4=(<DocType>,<DocNumber>,<DocDateTime>,<FMNumber>)[,<Invoice>,%5=(<InvNumber>,<Reason>)]
                    sParam = "%1,%2,%3,1,%4" & IIf(.InitReceiptType = ucsFscRetCreditNote, ",I,%5", vbNullString)
                    pvGetCommandConfig CONF_NAME, "FiscalStorno" & .InitReceiptType, eCmd, ucsFpcFiscalStorno, sParam, sParam, _
                        "FormatDate", sFormatDate, FORMAT_DATE_SHORT
                    sParam = Printf(sParam, .InitOperatorCode, .InitOperatorPassword, _
                        .InitUniqueSaleNo, _
                        .InitStornoData(ucsStoType) & _
                            "," & .InitStornoData(ucsStoReceiptNo) & _
                            "," & Format$(.InitStornoData(ucsStoReceiptDate), sFormatDate) & _
                            "," & .InitStornoData(ucsStoFiscalMemoryNo), _
                        .InitStornoData(ucsStoInvoiceNo) & _
                            "," & .InitStornoData(ucsStoReason))
                Case Else
                    '--- %1=<OpCode>,%2=<OpPwd>%3=([,<NSale>]),<TillNmb>[,<Invoice>]
                    sParam = "%1,%2%3,1" & IIf(.InitReceiptType = ucsFscRetInvoice, ",I", vbNullString)
                    pvGetCommandConfig CONF_NAME, "FiscalOpen" & .InitReceiptType, eCmd, ucsFpcFiscalOpen, _
                        sParam, sParam
                    sParam = Printf(sParam, .InitOperatorCode, .InitOperatorPassword, _
                        IIf(LenB(.InitUniqueSaleNo) <> 0, "," & .InitUniqueSaleNo, vbNullString))
                End Select
            Else
                sStornoType = Switch( _
                    .InitStornoData(ucsStoType) = ucsFscRsnTaxBaseReduction, "T", _
                    .InitStornoData(ucsStoType) = ucsFscRsnOperatorError, "E", _
                    True, "R")
                Select Case .InitReceiptType
                Case ucsFscRetStorno, ucsFscRetCreditNote
                    '--- %1=<OpNum>,%2=<Password>,<TillNum>[,<Invoice>%3=<InvNum>][,<UNP>],%4=(<StTipe><DocNo>)%5=([,<StUNP>,<StDT>,<StFMIN>])[#<StornoType>]
                    sParam = "%1,%2,1" & IIf(.InitReceiptType = ucsFscRetCreditNote, ",I,%3", vbNullString) & ",%4"
                    pvGetCommandConfig CONF_NAME, "FiscalStorno" & .InitReceiptType, eCmd, ucsFpcFiscalStorno, sParam, sParam, _
                        "FormatDate", sFormatDate, FORMAT_DATE_SHORT
                    sParam = Printf(sParam, .InitOperatorCode, .InitOperatorPassword, _
                        .InitStornoData(ucsStoInvoiceNo), _
                        sStornoType & .InitStornoData(ucsStoReceiptNo), _
                        IIf(LenB(.InitUniqueSaleNo) <> 0, _
                            "," & .InitUniqueSaleNo & _
                            "," & Format$(.InitStornoData(ucsStoReceiptDate), sFormatDate) & _
                            "," & .InitStornoData(ucsStoFiscalMemoryNo), vbNullString))
                Case Else
                    '--- %1=<OpNum>,%2=<Password>,<TillNum>[,<Invoice>]%3=[,<UNP>]
                    sParam = "%1,%2,1" & IIf(.InitReceiptType = ucsFscRetInvoice, ",I", vbNullString) & "%3"
                    pvGetCommandConfig CONF_NAME, "FiscalOpen" & .InitReceiptType, eCmd, ucsFpcFiscalOpen, _
                        sParam, sParam
                    sParam = Printf(sParam, .InitOperatorCode, .InitOperatorPassword, _
                        IIf(LenB(.InitUniqueSaleNo) <> 0, "," & .InitUniqueSaleNo, vbNullString))
                End Select
            End If
            Exceptions = bPrevExceptions
            bInLogin = True
            SendCommand eCmd, sParam
            bInLogin = False
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
            pvGetCommandConfig CONF_NAME, "FiscalSell" & .PrintRowType, eCmd, ucsFpcFiscalSell, sParam, "%1", _
                "FormatPrice", sFormatPrice, "0.00", _
                "FormatQty", sFormatQty, "0.000", _
                "FormatDisc", sFormatDisc, "0.00"
            vSplit = WrapText(.PluName, pvItemChars)
            For lIdx = 0 To UBound(vSplit) - 2
                pvPrintLine C_Str(vSplit(lIdx))
            Next
            sData = At(vSplit, UBound(vSplit))
            If UBound(vSplit) > 0 Then
                sData = At(vSplit, UBound(vSplit) - 1) & vbLf & sData
            End If
            sData = sData & vbTab & Chr$(191 + .PluVatGroup) & SafeFormat(.PluPrice, sFormatPrice)
            If .PluQuantity <> 1 Then
                sData = sData & "*" & SafeFormat(.PluQuantity, sFormatQty)
            End If
            Select Case .DiscType
            Case ucsFscDstPlu, ucsFscDstPluAbs
                sData = sData & IIf(.DiscType = ucsFscDstPlu, ",", ";") & SafeFormat(.DiscValue, sFormatDisc)
            End Select
            sParam = Printf(sParam, sData)
            SendCommand eCmd, sParam
        Case ucsRowDiscount
            Select Case .DiscType
            Case ucsFscDstSubtotal
                pvGetCommandConfig CONF_NAME, "FiscalSubtotal" & .DiscType, eCmd, ucsFpcFiscalSubtotal, sParam, "10,%1", _
                    "FormatDisc", sFormatDisc, "0.00"
            Case ucsFscDstSubtotalAbs
                pvGetCommandConfig CONF_NAME, "FiscalSubtotal" & .DiscType, eCmd, ucsFpcFiscalSubtotal, sParam, "10;%1", _
                    "FormatDisc", sFormatDisc, "0.00"
            End Select
            sParam = Printf(sParam, SafeFormat(.DiscValue, sFormatDisc))
            SendCommand eCmd, sParam
            Select Case .DiscType
            Case ucsFscDstSubtotal
                For lIdx = 1 To UBound(uCtx.GrpTotal)
                    If Abs(uCtx.GrpTotal(lIdx)) > DBL_EPSILON Then
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
            pvPrintLine vSplit
        Case ucsRowPayment
            If .PmtType <> 0 Then
                pvGetCommandConfig CONF_NAME, "PaymentTypes", 0, 0, sParam, DefPmtTypes
                sType = Mid$(sParam, .PmtType, 1)
                If .PmtType > 4 Then
                    bPrevExceptions = pvSetExceptions(False)
                    If IsDaisy Then
                        pvGetCommandConfig CONF_NAME, "ExtendedInitText" & .PmtType, eCmd, ucsFpcExtendedInitText, sParam, "P" & (56 + .PmtType) & ",%1"
                    ElseIf IsIncotex Then
                        pvGetCommandConfig CONF_NAME, "ExtendedInitText" & .PmtType, eCmd, ucsFpcExtendedInitText, sParam, "P" & (7 + .PmtType) & ",%1"
                    Else
                        pvGetCommandConfig CONF_NAME, "InitPaymentType" & .PmtType, eCmd, ucsFpcInitPaymentType, sParam, sType & ",%1"
                    End If
                    sParam = Printf(sParam, Left$(.PmtName, m_uConfig.MaxPaymentLen))
                    SendCommand eCmd, sParam
                    Exceptions = bPrevExceptions
                End If
                pvGetCommandConfig CONF_NAME, "FiscalPayment" & .PmtType, eCmd, ucsFpcFiscalPayment, sParam, vbTab & sType & "%1", _
                    "FormatAmount", sFormatAmount, "0.00"
                sParam = Printf(sParam, SafeFormat(.PmtAmount, sFormatAmount))
                SendCommand eCmd, sParam
                uCtx.PmtPrinted = True
            Else
                '--- check if anything left for payment
                pvGetCommandConfig CONF_NAME, "InfoTransaction", eCmd, ucsFpcInfoTransaction, sParam, "T", _
                    "Index", lIndex, 0, _
                    "Status", lStatus, 0, _
                    "AmountIndex", lAmountIndex, 2, _
                    "PaidIndex", lPaidIndex, 3
                vSplit = Split(SendCommand(eCmd, sParam), DEF_DELIM)
                If C_Lng(At(vSplit, lIndex)) <> lStatus Then
                    If C_Dbl(At(vSplit, lAmountIndex)) > C_Dbl(At(vSplit, lPaidIndex)) _
                            Or LenB(At(vSplit, lAmountIndex)) = 0 _
                            Or Not uCtx.PmtPrinted Then
                        pvGetCommandConfig CONF_NAME, "FiscalPayment" & .PmtType, eCmd, ucsFpcFiscalPayment, sParam, vbTab
                        bPrevExceptions = pvSetExceptions(False)
                        SendCommand eCmd, sParam
                        Exceptions = bPrevExceptions
                        If Status(ucsStbPrintingError) = 0 Then
                            uCtx.PmtPrinted = True
                        End If
                    End If
                    Select Case .PrintRowType
                    Case ucsFscRetInvoice, ucsFscRetCreditNote
                        pvGetCommandConfig CONF_NAME, "FiscalCgInfo", eCmd, ucsFpcFiscalCgInfo, sParam, Replace("%1#%2#%3#%4#%5#%6", "#", vbTab), _
                            "FormatAddr", sFormatAddr, "%1" & vbLf & "%2"
                        With uRow(0)
                            If LenB(.InitInvData(ucsInvCgAddress)) <> 0 Then
                                sFormatAddr = Printf(sFormatAddr, Left$(.InitInvData(ucsInvCgCity), m_uConfig.MaxCgInfo(5)), _
                                    Left$(.InitInvData(ucsInvCgAddress), m_uConfig.MaxCgInfo(6)))
                            Else
                                sFormatAddr = Left$(.InitInvData(ucsInvCgCity), m_uConfig.MaxCgInfo(5))
                            End If
                            If Not IsEcr Then
                                sTaxNoPrefix = At(Split("|#|*|^", "|"), .InitInvData(ucsInvCgTaxNoType))
                            Else
                                sTaxNoSuffix = vbTab & .InitInvData(ucsInvCgTaxNoType)
                            End If
                            sParam = Printf(sParam, sTaxNoPrefix & Left$(.InitInvData(ucsInvCgTaxNo), m_uConfig.MaxCgInfo(0)) & sTaxNoSuffix, _
                                Left$(.InitOperatorName, m_uConfig.MaxCgInfo(1)), _
                                Left$(Zn(C_Str(.InitInvData(ucsInvCgPrsReceive)), .InitInvData(ucsInvCgPrsReceive)), m_uConfig.MaxCgInfo(2)), _
                                Left$(.InitInvData(ucsInvCgName), m_uConfig.MaxCgInfo(3)), _
                                Left$(.InitInvData(ucsInvCgVatNo), m_uConfig.MaxCgInfo(4)), _
                                sFormatAddr)
                        End With
                        SendCommand eCmd, sParam
                    End Select
                    pvGetCommandConfig CONF_NAME, "FiscalClose" & .PrintRowType, eCmd, ucsFpcFiscalClose, sParam, vbNullString
                    SendCommand eCmd, sParam
                End If
                pvGetCommandConfig CONF_NAME, "InfoStatus", eCmd, ucsFpcInfoStatus, sParam, "W"
                SendCommand eCmd, sParam
            End If
        End Select
        End With
        uCtx.Row = uCtx.Row + 1
    Loop
    '--- success
    pvPrintReceipt = True
    Exit Function
EH:
    pvSetLastError Err.Description, DontRaise:=True
    PrintError FUNC_NAME
    If bInLogin Then
        pvSetLastError m_sLastError & pvInternal(ucsErrInvalidOperatorPassword), ucsFerInvalidPassword, DontRaise:=True
    End If
    If uCtx.PluCount > 0 And Not uCtx.PmtPrinted Then
        pvSetLastError Printf(pvInternal(ucsErrErrorInLine), uCtx.PluCount, m_sLastError), m_eLastErrNo, DontRaise:=True
    End If
    Err.Raise vbObjectError, , m_sLastError
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
                    vSplit = WrapText(m_uRow(lRow).PluName, pvItemChars)
                    lIdx = LimitLong(UBound(vSplit), , 1)
                    vSplit(lIdx) = AlignText(vSplit(lIdx), SafeFormat(dblTotal + dblDiscTotal, "0.00") & " " & Chr$(191 + m_uRow(lRow).PluVatGroup), pvRowChars)
                    m_uRow(lRow).RowType = ucsRowLine
                    m_uRow(lRow).LineText = vSplit(0)
                    If lIdx > 0 Then
                        AddLine At(vSplit, 1), False, lRow + 1
                        lRow = lRow + 1
                    ElseIf lIdx = 0 And m_uRow(lRow).PluQuantity <> 1 Then
                        AddLine AlignText(vbNullString, SafeFormat(m_uRow(lRow).PluQuantity, "0.000") & " x " & SafeFormat(m_uRow(lRow).PluPrice, "0.00"), pvRowChars - 2), False, lRow
                    End If
                    If dblPrice < -DBL_EPSILON Then
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
                        AddPLU Printf(IIf(dblDiscTotal > DBL_EPSILON, pvText(ucsTxtSurcharge), pvText(ucsTxtDiscount)), SafeFormat(Abs(dblDiscount), "0.00") & " %"), _
                            dblDiscTotal, 1, m_uRow(lRow).PluVatGroup, lRow + 1
                    End If
                ElseIf m_uRow(lRow).DiscType = ucsFscDstPlu And dblPrice < -DBL_EPSILON Then
                    '--- convert PLU discount on void rows
                    If m_uConfig.AbsoluteDiscount Then
                        m_uRow(lRow).DiscType = ucsFscDstPluAbs
                        m_uRow(lRow).DiscValue = dblDiscTotal
                    Else
                        dblDiscount = m_uRow(lRow).DiscValue
                        m_uRow(lRow).DiscType = 0
                        m_uRow(lRow).DiscValue = 0
                        AddPLU Printf(IIf(dblTotal * dblDiscount > DBL_EPSILON, pvText(ucsTxtSurcharge), pvText(ucsTxtDiscount)), SafeFormat(Abs(dblDiscount), "0.00") & " %"), _
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
                            If Abs(uSum.GrpTotal(lIdx)) > DBL_EPSILON Then
                                AddPLU Printf(IIf(uSum.GrpTotal(lIdx) * dblDiscount > DBL_EPSILON, pvText(ucsTxtSurcharge), pvText(ucsTxtDiscount)), SafeFormat(Abs(dblDiscount), "0.00") & " %"), _
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
                        .PrintRowType = ucsFscRetNonfiscal
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
                        .PrintRowType = ucsFscRetNonfiscal
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
                    If Abs(uCtx.GrpTotal(lJdx)) > DBL_EPSILON Then
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

Private Function pvGetDeviceInfo(oConnector As IDeviceConnector, ByVal lTimeout As Long, vResult As Variant, sProtocol As String, sModel As String, Optional Error As String) As Boolean
    Const FUNC_NAME     As String = "pvGetDeviceInfo"
    Const CONF_NAME     As String = "GetDeviceInfo"
    Dim eCmd            As UcsIslCommandsEnum
    Dim sParam          As String
    Dim sDaisyPrefix    As String
    Dim sIncotexPrefix  As String
    Dim lNameIndex      As Long
    Dim lSwitchesIndex  As Long
    Dim lSerialIndex    As Long
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoDiagnostics", eCmd, ucsFpcInfoDiagnostics, sParam, "0", _
        "DaisyPrefix", sDaisyPrefix, "DY", _
        "IncotexPrefix", sIncotexPrefix, "IN", _
        "NameIndex", lNameIndex, 0, _
        "SwitchesIndex", lSwitchesIndex, 3, _
        "SerialIndex", lSerialIndex, 4
    If Not IsArray(vResult) Then
        vResult = Split(pvSendCommand(oConnector, lTimeout, eCmd, sParam, Error), DEF_DELIM)
        If LenB(At(vResult, 0)) = 0 Then
            sParam = vbNullString
            vResult = Split(pvSendCommand(oConnector, lTimeout, eCmd, sParam, Error), DEF_DELIM)
        End If
    End If
    If UBound(vResult) >= lSerialIndex And LenB(Error) = 0 Then
        '--- check serial
        If Left$(At(vResult, lSerialIndex), Len(sDaisyPrefix)) = sDaisyPrefix And LenB(sDaisyPrefix) <> 0 Then
            sProtocol = STR_PROTOCOL_DAISY_ECR
        ElseIf Left$(At(vResult, lSerialIndex), Len(sIncotexPrefix)) = sIncotexPrefix And LenB(sIncotexPrefix) <> 0 Then
            sProtocol = STR_PROTOCOL_INCOTEX_ECR
        ElseIf LenB(At(vResult, lSerialIndex)) <> 0 Then
            sProtocol = STR_PROTOCOL_DATECS_FP
        End If
        sModel = At(vResult, lNameIndex)
        '-- lSwitchesIndex in older models -> country index (6 = BG)
        If Len(At(vResult, lSwitchesIndex)) > 2 Then
            sModel = sModel & " " & At(vResult, lNameIndex + 1)
        End If
        '--- success
        pvGetDeviceInfo = True
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function pvStripBold(sText As String) As String
    Const FUNC_NAME     As String = "pvStripBold"
    
    pvStripBold = preg_replace(GetConfigForCommand(m_uConfig.ConfigCommands, m_uConfig.LocalizedCommands, FUNC_NAME, "Pattern", "/\^/i"), _
        sText, GetConfigForCommand(m_uConfig.ConfigCommands, m_uConfig.LocalizedCommands, FUNC_NAME, "Replace", vbNullString))
End Function

Private Sub pvGetCommandConfig( _
            sFunc As String, _
            sKey As String, _
            eCmd As UcsIslCommandsEnum, _
            ByVal eCmdDefault As UcsIslCommandsEnum, _
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

Private Sub pvSetLastError(sError As String, Optional ByVal ErrNum As UcsFiscalErrorsEnum = -1, Optional ByVal DontRaise As Boolean)
    If ErrNum < 0 Then
        m_eLastErrNo = IIf(LenB(sError) = 0, ucsFerNone, ucsFerGeneralError)
    Else
        m_eLastErrNo = ErrNum
    End If
    m_sLastError = sError
    If Exceptions And LenB(sError) <> 0 And Not DontRaise Then
        Err.Raise vbObjectError, , m_sLastError
    End If
End Sub

Private Function pvToDate(sText As String) As Date
    If Len(sText) = 6 And IsNumeric(sText) Then
        pvToDate = DateSerial(2000 + C_Lng(Mid$(sText, 5, 2)), C_Lng(Mid$(sText, 3, 2)), C_Lng(Mid$(sText, 1, 2)))
    End If
End Function

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Initialize()
    m_lTimeout = DEF_TIMEOUT
    m_lSequence = &H20
    m_bExceptions = True
    LocalizedText(ucsFscLciInternalErrors) = STR_INTERNAL
    LocalizedText(ucsFscLciPrinterStatuses) = STR_STATUSES
    LocalizedText(ucsFscLciPrinterErrors) = STR_ERRORS_ONLY
    LocalizedText(ucsFscLciDipSwitchesTexts) = STR_DIP_SWITCHES
    LocalizedText(ucsFscLciMemoryTexts) = STR_MEMORY
    LocalizedText(ucsFscLciExtendedErrors) = STR_EXT_ERRORS
    m_uConfig.CommEventTimeout = DEF_COMM_EVENT_TIMEOUT * 2
    Set m_uConfig.ConfigCommands = GetConfigCollection(STR_PROTOCOL_DATECS_FP, "Commands")
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
    IDeviceProtocol_CashDebitCredit = CashDebitCredit(Value)
End Function

Private Function IDeviceProtocol_Connect() As Boolean
    IDeviceProtocol_Connect = Connect()
End Function

Private Function IDeviceProtocol_CopyLastReceipt(Optional ReceiptNumber As String) As Boolean
    IDeviceProtocol_CopyLastReceipt = CopyLastReceipt(ReceiptNumber)
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
    IDeviceProtocol_GetDefaultPassword = GetDefaultPassword(OperatorCode)
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

Private Function IDeviceProtocol_GetFiscalMemoryNo() As String
    IDeviceProtocol_GetFiscalMemoryNo = GetFiscalMemoryNo()
End Function

Private Function IDeviceProtocol_GetFooterText(ByVal Index As Long) As String
    IDeviceProtocol_GetFooterText = GetFooterText(Index)
End Function

Private Function IDeviceProtocol_GetHeaderText(ByVal Index As Long) As String
    IDeviceProtocol_GetHeaderText = GetHeaderText(Index)
End Function

Private Function IDeviceProtocol_GetLastReceiptNo() As String
    IDeviceProtocol_GetLastReceiptNo = GetLastReceiptNo()
End Function

Private Function IDeviceProtocol_GetLastError(Optional ErrNo As UcsFiscalErrorsEnum) As String
    ErrNo = LastErrNo
    IDeviceProtocol_GetLastError = LastError
End Function

Private Function IDeviceProtocol_GetLastInvoiceNo() As String
    IDeviceProtocol_GetLastInvoiceNo = GetLastInvoiceNo()
End Function

Private Function IDeviceProtocol_GetPaymentName(ByVal Index As Long) As String
    IDeviceProtocol_GetPaymentName = GetPaymentName(Index)
End Function

Private Function IDeviceProtocol_GetResumeTicket() As String
    IDeviceProtocol_GetResumeTicket = GetResumeTicket()
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

Private Function IDeviceProtocol_GetTotalsByPayments() As Variant
    IDeviceProtocol_GetTotalsByPayments = GetTotalsByPayments()
End Function

Private Function IDeviceProtocol_GetTotalsByVatGroups() As Variant
    IDeviceProtocol_GetTotalsByVatGroups = GetTotalsByVatGroups()
End Function

Private Function IDeviceProtocol_GetTotalsStats() As Variant
    IDeviceProtocol_GetTotalsStats = GetTotalsStats()
End Function

Private Function IDeviceProtocol_Init(Device As String, Optional ByVal Timeout As Long, Optional ByVal RowChars As Long) As Boolean
    IDeviceProtocol_Init = Init(Device, Timeout, RowChars)
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

Private Function IDeviceProtocol_StartReceipt(ByVal ReceiptType As UcsFiscalReceiptTypeEnum, OperatorCode As String, OperatorName As String, OperatorPassword As String, Optional UniqueSaleNo As String, Optional TableNo As String, Optional InvDocNo As String, Optional InvCgTaxNo As String, Optional ByVal InvCgTaxNoType As UcsFiscalTaxNoTypeEnum, Optional InvCgVatNo As String, Optional InvCgName As String, Optional InvCgCity As String, Optional InvCgAddress As String, Optional InvCgPrsReceive As String, Optional OwnData As String, Optional ByVal StornoType As UcsFiscalStornoTypeEnum, Optional StornoReceiptNo As String, Optional ByVal StornoReceiptDate As Date, Optional StornoFiscalMemoryNo As String, Optional StornoInvoiceNo As String, Optional StornoReason As String) As Boolean
    IDeviceProtocol_StartReceipt = StartReceipt(ReceiptType, OperatorCode, OperatorName, OperatorPassword, UniqueSaleNo, InvDocNo, InvCgTaxNo, InvCgTaxNoType, InvCgVatNo, InvCgName, InvCgCity, InvCgAddress, InvCgPrsReceive, StornoType, StornoReceiptNo, StornoReceiptDate, StornoFiscalMemoryNo, StornoInvoiceNo, StornoReason)
End Function

