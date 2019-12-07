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
    [_ucsIslCfgMax]
End Enum

'=========================================================================
' API
'=========================================================================

'--- error codes
Private Const ERROR_IO_PENDING              As Long = 997
'--- for PurgeComm
Private Const PURGE_TXCLEAR                 As Long = &H4
Private Const PURGE_RXCLEAR                 As Long = &H8

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
    Offset                      As Long
    OffsetHigh                  As Long
    hEvent                      As Long
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_INTERNAL              As String = "Timeout waiting for response|Invalid response length: %1|Invalid response checksum: %1|Response missing postamble: %1|Response missing preamble: %1|No receipt started| or invalid operator password|Error in line %1: %2|Cannot cancel previous receipt|Invalid report type|Error %1|Memory %1|Unsupported payment type %1"
Private Const STR_STATUSES_A            As String = "Syntax error in the received data|Invalid command code received|The clock is not set|No customer display is connected|Printing unit fault|5|The printer cover is open|7|The command resulted in an overflow of some amount fields|The command is not allowed in the current fiscal mode|The RAM has been reset|Low battery (the real-time clock is in RESET status)|A refund (storno) receipt is open|A service receipt with 90-degree rotated text printing is open|The built-in tax terminal is not responding|7|No paper|Low paper|End of the EJ (less than 1 MB free in the EJ)|A fiscal receipt is open|The end of the EJ is near (less than 10 MB free in the EJ)|A service receipt is open|The end of the EJ is very near (only certain receipts are allowed)"
Private Const STR_ERRORS_A              As String = "Syntax error in the received data|Invalid command code received|2|3|Printing unit fault|5|6|7|0|The command is not allowed in the current fiscal mode|The RAM has been reset|Low battery (the real-time clock is in RESET status)|4|5|6|7|No paper"
Private Const STR_MEMORY_A              As String = "Fiscal memory store error|BULSTAT UIC is set|Unique Printer ID and Fiscal Memory ID are set|There is space for less than 50 records remaining in the FM|The fiscal memory is full|FM error|The printing head is overheated|7|The fiscal memory is set in READONLY mode (locked)|The fiscal memory is formatted|The last fiscal memory store operation is not successful|The printer is in fiscal mode|The tax rates are set at least once|Fiscal memory read error"
Private Const STR_STATUSES_B            As String = "Syntax error in the received data|Invalid command code received|The clock is not set|No customer display is connected|4|5|6|7|The command resulted in an overflow of some amount fields|The command is not allowed in the current fiscal mode|2|3|4|Unsend documents for more than duration set|The built-in tax terminal is not responding|7|No paper|Low paper|End of the EJ (less than 1 MB free in the EJ)|A fiscal receipt is open|The end of the EJ is near (less than 10 MB free in the EJ)|A service receipt is open"
Private Const STR_ERRORS_B              As String = "Syntax error in the received data|Invalid command code received|2|3|4|5|6|7|0|The command is not allowed in the current fiscal mode|2|3|4|5|6|7|No paper"
Private Const STR_MEMORY_B              As String = "Fiscal memory store error|BULSTAT UIC is set|Unique Printer ID and Fiscal Memory ID are set|There is space for less than 50 records remaining in the FM|The fiscal memory is full|FM error|6|7|0|FM is formatted|2|Device is fiscalized|VAT are set at least once"
Private Const STR_DIP_SWITCHES          As String = "Center header and footer|Enable header|Sw1.3|Sw1.4|Zero memory|Transparent display|No display|7"
Private Const STR_EXT_ERRORS            As String = "1|Overflow after operation|3|No more sales in this bill|4|No more payments in this bill|5|Attempt to perform zero transaction|6|Attempt to perform a sale, after payment is initiated|7|Operation not allowed|8|TAX group with disabled sales|9|Invoice No range not set|11|More than one decimal point|12|More than one character '+' or '-'|13|Symbol '+' or '-' is not in the first position|14|Illegal symbol, e.g. barcode contains not only numbers|15|Extra numbers after decimal point|16|Extra characters|20|With current settings cannot process this command from PC|21|Value outside limits|22|Review system parameter 8|23|Attempt to ""deep"" void after discount/surcharge on subtotal|24|Attempt to ""deep"" void of non-existent transaction|25|Attempt to make a payment without having sales|26|Attempt to sell an item with quantity exceeding its stock|41|Barcode is incorrect (wrong checksum)|42|Attempt to sell without barcode|" & _
                                                    "43|Attempt to program with weight barcode|44|Attempt to sell without programmed barcode|45|Attempt to program already existing barcode|66|Incorrect password|71|!!! Incorrect data in fiscal memory !!!!|72|!!! Error saving in fiscal memory !!!!|76|Information needed from NRA servers|90|Needs a reset with monthly report|91|Needs a reset with daily z-report|92|Needs a reset with report by operators|93|Needs a reset with report by articles|94|Can not reprogram this field|81|Daily z-report overflow|83|Report by operators overflow|84|Report by articles overflow|84|Monthly report overflow|88|Electronic journal overflow|102|No communication between fiscal unit and tax terminal|104|Incorrect communication between fiscal unit and tax terminal|110|SIM card replaced|111|Error in communication between tax terminal and NRA servers|113|NRA server rejects data sent|" & _
                                                    "117|Failed to connect tax terminal with mobile operator network|118|The operation is denied|119|Incorrect value|120|Missing value"
Private Const STR_RECEIPT_TEXTS         As String = "CASH|CARD|CHEQUE|CARD|SURCHARGE %1|DISCOUNT %1|SUBTOTAL|Sales %1|EIC"
Private Const DEF_AUTODETECTTIMEOUT     As Long = 100
Private Const DEF_DEVICEINFOTIMEOUT     As Long = 1000
Private Const DEF_AUTODETECTSPEEDS      As String = "115200|38400|9600|19200|57600"
Private Const DEF_TIMEOUT               As Long = 3000
Private Const DEF_OPERNO                As String = "1"
Private Const DEF_PASSWORD              As String = "0000"
Private Const DEF_SOCKET_PORT           As Long = 4999
Private Const DEF_SERIAL_PORT           As String = "COM1"
Private Const DEF_SERIAL_SPEED          As Long = 115200
Private Const DEF_ROW_CHARS             As Long = 42
Private Const DEF_MIN_DISCOUNT          As Double = -99
Private Const DEF_MAX_DISCOUNT          As Double = 99
Private Const DEF_MAX_RECEIPT_ROWS      As Long = 1000
Private Const DEF_MAX_PAYMENT_LEN       As Long = 24
Private Const DEF_MAX_CGINFO            As String = "14|24|25|25|14|35|35"
Private Const DEF_MAX_OPERATOR_NAME     As Long = 24
Private Const DEF_COMM_EVENT_TIMEOUT    As Long = 250
Private Const DEF_SEP                   As String = ","
Private Const DEF_SEQUENCE              As Long = &H20
Private Const LNG_MAX_SEQUENCE          As Long = &H7F
Private Const MIK_OFFSET                As Long = &H80
'Private Const WIN_OFFSET                As Long = &HC0
Private Const FORMAT_DATE_EXTENDED      As String = "dd\-mm\-yy hh\:nn\:ss"
Private Const FORMAT_DATE_SHORT         As String = "ddmmyyhhnnss"
Private Const FORMAT_DATE_REPORT        As String = "ddmmyy"

Private m_sDevice                   As String
Private m_oConnector                As IDeviceConnector
Private m_lTimeout                  As Long
Private m_lSequence                 As Long
Private m_eStatus                   As UcsIslStatusBitsEnum
Private m_eDip                      As UcsIslDipSwitchesEnum
Private m_eMemory                   As UcsIslFiscalMemoryEnum
Private m_sCommandLog               As String
Private m_vInfoDiagnostics          As Variant
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

Private Enum UcsCharset
    ucsCrsAuto
    ucsCrsWindows
    ucsCrsMik
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
    MaxPaymentLen       As Long
    MaxCgInfo(0 To 6)   As Long
    MaxOperatorName     As Long
    CommEventTimeout    As Long
    DefPmtTypes         As String
    ConfigCommands      As Collection
    LocalizedCommands   As Collection
    PmtMapping          As Variant
End Type

Private Type UcsLocalizedConfigType
    LocalizedText(0 To [_ucsFscLciMax] - 1) As Variant
End Type

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunction As String)
    Debug.Print "Critical error: " & Err.Description & " [" & MODULE_NAME & "." & sFunction & "]"
    OutputDebugLog MODULE_NAME, sFunction & "(" & Erl & ")", "Run-time error: " & Err.Description
End Sub

Private Sub RaiseError(sFunction As String)
    Debug.Print "Critical error: " & Err.Description & " [" & MODULE_NAME & "." & sFunction & "]"
    OutputDebugLog MODULE_NAME, sFunction & "(" & Erl & ")", "Run-time error: " & Err.Description
    Err.Raise Err.Number, MODULE_NAME & "." & sFunction & "(" & Erl & ")" & vbCrLf & Err.Source, Err.Description
End Sub

Private Sub DebugLog(sFunction As String, sText As String)
    OutputDebugLog MODULE_NAME, sFunction, sText
End Sub

Private Sub DebugDataDump(sFunction As String, sPrefix As String, sData As String)
    OutputDebugDataDump MODULE_NAME, sFunction, sPrefix, sData
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
    StatusText = pvGetFlagsText(m_eStatus, pvLocalizedTextArray(ucsFscLciPrinterStatuses))
End Property

Property Get ErrorText() As String
    ErrorText = pvGetFlagsText(m_eStatus, pvLocalizedTextArray(ucsFscLciPrinterErrors))
    If IsDaisy And (m_eDip And &H7F) <> 0 Then
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
End Property

Property Get IsEcr() As Boolean
    Const FUNC_NAME     As String = "IsEcr [get]"
    Const CONF_NAME     As String = "IsEcr"
    Dim lTimeout        As Long
    Dim eCmd            As UcsIslCommandsEnum
    Dim sParam          As String
    
    On Error GoTo EH
    If m_uConfig.IsEcr = vbUseDefault Then
        pvGetCommandConfig CONF_NAME, "EcrReadRow", eCmd, ucsIslCmdEcrReadRow, sParam, "2;1;", _
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
            m_uConfig.DefPmtTypes = "PNCDB"
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
    Dim eCmd            As UcsIslCommandsEnum
    Dim sParam          As String
    Dim lIndex          As Long
    Dim sResult         As String
    Dim vSplit          As Variant
    
    On Error GoTo EH
    If m_uPrintData.Config.RowChars = 0 Then
        pvGetCommandConfig CONF_NAME, "ExtendedInfoConsts", eCmd, ucsIslCmdExtendedInfoConsts, sParam, vbNullString, _
            "Index", lIndex, 8, _
            "DaisyPrefix", sDaisyPrefix, "DY", _
            "IncotexPrefix", sIncotexPrefix, "IN"
        If Left$(GetDeviceSerialNo(), Len(sDaisyPrefix)) = sDaisyPrefix And LenB(sDaisyPrefix) <> 0 Or _
                Left$(GetDeviceSerialNo(), Len(sIncotexPrefix)) = sIncotexPrefix And LenB(sIncotexPrefix) <> 0 Then
            vSplit = Split(SendCommand(eCmd, sParam), DEF_SEP)
            m_uPrintData.Config.RowChars = C_Lng(At(vSplit, lIndex + 0))
            If m_uPrintData.Config.CommentChars = 0 Then
                m_uPrintData.Config.CommentChars = C_Lng(At(vSplit, lIndex + 1))
            End If
            If m_uPrintData.Config.ItemChars = 0 Then
                m_uPrintData.Config.ItemChars = C_Lng(At(vSplit, lIndex + 2))
            End If
        Else
            pvGetCommandConfig CONF_NAME, "InitLogo", eCmd, ucsIslCmdInitLogo, sParam, "R0"
            sResult = SendCommand(eCmd, sParam)
            '--- check for paper width above 57mm (432px) e.g. 80mm paper is 576px wide
            If Len(sResult) * 4 > 432 Then
                m_uPrintData.Config.RowChars = Len(sResult) * 4 / 12
            End If
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
    
    lCurrentIdx = IIf(m_uConfig.IsEcr = vbTrue, 1, 0)
    pvLocalizedTextArray = m_uLConfig(lCurrentIdx).LocalizedText(eIdx)
    '--- fallback to default localization config
    If Not IsArray(pvLocalizedTextArray) And lCurrentIdx <> 0 Then
        pvLocalizedTextArray = m_uLConfig(0).LocalizedText(eIdx)
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
    Dim sReplyAlt     As String
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
    Dim sFirmware       As String
    Dim vSpeeds         As Variant
    Dim baText()        As Byte
    
    On Error GoTo EH
    sProbeText = GetConfigValue(STR_PROTOCOL_DATECS_X, "ProbeText", Chr$(SYM_PREAMBLE) & Chr$(&H25) & Chr$(SYM_BASEDATA) & Chr$(&H5A) & Chr$(&H30) & Chr$(SYM_POSTAMBLE) & Chr$(&H30) & Chr$(&H30) & Chr$(&H3D) & Chr$(&H34) & Chr$(SYM_TERMINATOR))
    sReplyText = GetConfigValue(STR_PROTOCOL_DATECS_X, "ReplyText", Chr$(SYM_SYN))  '--- SYN (&H16)
    sReplyAlt = GetConfigValue(STR_PROTOCOL_DATECS_X, "ReplyAlt", Chr$(SYM_PREAMBLE))
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
                    Set .Connector = New cSerialPortConnector
                    If .Connector.Init(At(vPorts, lIdx) & "," & vSpeed, lTimeout) Then
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
                    End If
                    If LenB(.Buffer) <> 0 Then
                        If Left$(.Buffer, Len(sReplyText)) = sReplyText Or Left$(.Buffer, Len(sReplyAlt)) = sReplyAlt Then
                            If pvGetDeviceInfo(.Connector, lDeviceInfoTimeout, Empty, sProtocol, sModel, sFirmware) Then
                                vRetVal(lIdx) = Array(At(vPorts, lIdx), vSpeed, sProtocol, sModel, sFirmware)
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

Public Function Init(sDeviceString As String) As Boolean
    Const FUNC_NAME     As String = "Init"
    Const CONF_NAME     As String = FUNC_NAME
    Dim oOptions        As Object
    Dim sSerial         As String
    Dim lIdx            As Long
    Dim eCmd            As UcsIslCommandsEnum
    Dim sParam          As String
    Dim sDaisyPrefix    As String
    Dim sIncotexPrefix  As String
    Dim lSerialIndex    As Long
    Dim lSwitchesIndex  As Long
    Dim sError          As String
    Dim vSplit          As Variant
    
    On Error GoTo EH
    DebugLog FUNC_NAME, "sDeviceString=" & sDeviceString
    pvSetLastError vbNullString
    Set oOptions = ParseDeviceString(sDeviceString)
    m_lTimeout = C_Lng(JsonItem(oOptions, "Timeout"))
    If m_lTimeout <= 0 Then
        m_lTimeout = DEF_TIMEOUT
    End If
    m_sDevice = ToConnectorDevice(oOptions, DEF_SERIAL_PORT, DEF_SERIAL_SPEED, DEF_SOCKET_PORT)
    If Not Connect() Then
        GoTo QH
    End If
    pvGetCommandConfig CONF_NAME, "InfoDiagnostics", eCmd, ucsIslCmdInfoDiagnostics, sParam, "0", _
        "DaisyPrefix", sDaisyPrefix, "DY", _
        "IncotexPrefix", sIncotexPrefix, "IN", _
        "SerialIndex", lSerialIndex, 4, _
        "SwitchesIndex", lSwitchesIndex, 3
    m_vInfoDiagnostics = Split(pvSendCommand(m_oConnector, m_lTimeout, eCmd, sParam, sError), DEF_SEP)
    If LenB(At(m_vInfoDiagnostics, 0)) = 0 Then
        sParam = vbNullString
        m_vInfoDiagnostics = Split(pvSendCommand(m_oConnector, m_lTimeout, eCmd, sParam, sError), DEF_SEP)
    End If
    If LenB(sError) <> 0 Then
        pvSetLastError sError
        GoTo QH
    End If
    '--- load config values
    sSerial = At(m_vInfoDiagnostics, lSerialIndex)
    With m_uPrintData.Config
        .RowChars = GetConfigNumber(sSerial, "RowChars", C_Lng(JsonItem(oOptions, "RowChars")))
        .CommentChars = GetConfigNumber(sSerial, "CommentChars", C_Lng(JsonItem(oOptions, "CommentChars")))
        .ItemChars = GetConfigNumber(sSerial, "ItemChars", C_Lng(JsonItem(oOptions, "ItemChars")))
        .AbsoluteDiscount = C_Bool(GetConfigValue(sSerial, "AbsoluteDiscount", True))
        .NegativePrices = C_Bool(GetConfigValue(sSerial, "NegativePrices", True))
        .MinDiscount = GetConfigNumber(sSerial, "MinDiscount", Zndbl(C_Dbl(JsonItem(oOptions, "MinDiscount")), DEF_MIN_DISCOUNT))
        .MaxDiscount = GetConfigNumber(sSerial, "MaxDiscount", Zndbl(C_Dbl(JsonItem(oOptions, "MaxDiscount")), DEF_MAX_DISCOUNT))
        .MaxReceiptRows = GetConfigNumber(sSerial, "MaxReceiptRows", Znl(C_Lng(JsonItem(oOptions, "MaxReceiptRows")), DEF_MAX_RECEIPT_ROWS))
    End With
    With m_uConfig
        .Charset = ucsCrsAuto
        .IsEcr = vbUseDefault
        .IsDaisy = (Left$(GetDeviceSerialNo(), Len(sDaisyPrefix)) = sDaisyPrefix) And LenB(sDaisyPrefix) <> 0
        .IsIncotex = (Left$(GetDeviceSerialNo(), Len(sIncotexPrefix)) = sIncotexPrefix) And LenB(sIncotexPrefix) <> 0
        .MaxPaymentLen = GetConfigNumber(sSerial, "MaxPaymentLen", Znl(C_Lng(JsonItem(oOptions, "MaxPaymentLen")), DEF_MAX_PAYMENT_LEN))
        vSplit = Split(DEF_MAX_CGINFO, "|")
        For lIdx = LBound(.MaxCgInfo) To UBound(.MaxCgInfo)
            .MaxCgInfo(lIdx) = GetConfigNumber(sSerial, "MaxCgInfo" & lIdx, C_Lng(vSplit(lIdx)))
        Next
        .MaxOperatorName = GetConfigNumber(sSerial, "MaxOperatorName", Znl(C_Lng(JsonItem(oOptions, "MaxOperatorName")), DEF_MAX_OPERATOR_NAME))
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
    Set m_oConnector = InitDeviceConnector(m_sDevice, m_lTimeout, LocalizedText(ucsFscLciConnectorErrors), m_uPrintData.LastError)
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
    If eCmd = ucsIslCmdNoop Then
        Exit Function
    End If
    '--- auto-sense encoding
    If m_uConfig.Charset = ucsCrsAuto Then
        pvGetCommandConfig CONF_NAME, "InitDecimals", eDecimalsCmd, ucsIslCmdInitDecimals, sDecimalsParam, vbNullString, _
            "Index", lDecimalsIndex, 2
        vResult = Split(pvSendCommand(m_oConnector, m_lTimeout, eDecimalsCmd, sDecimalsParam, sError), DEF_SEP)
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
    If Status(ucsIslStbPrintingError) Then
        pvSetLastError Zn(ErrorText, Printf(pvInternal(ucsErrUnknownError), "0x" & Hex$(m_eStatus)))
        GoTo QH
    End If
    If Memory(ucsIslFmeMemoryError) Then
        pvSetLastError Zn(MemoryText, Printf(pvInternal(ucsErrUnknownMemory), "0x" & Hex$(m_eMemory)))
        GoTo QH
    End If
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function CopyLastReceipt(Optional ReceiptNumber As String) As Boolean
    Const FUNC_NAME     As String = "CopyLastReceipt"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsIslCommandsEnum
    Dim sParam          As String
    Dim sFormat         As String
    
    On Error GoTo EH
    If LenB(ReceiptNumber) = 0 Then
        pvGetCommandConfig CONF_NAME, "PrintDuplicate", eCmd, ucsIslCmdPrintDuplicate, sParam, "1"
        SendCommand eCmd, sParam
        If LenB(m_uPrintData.LastError) <> 0 Then
            GoTo QH
        End If
    Else
        sParam = "P"
        pvGetCommandConfig CONF_NAME, "InitEcTape" & sParam, eCmd, ucsIslCmdInitEcTape, sParam, sParam, _
            "Format", sFormat, "%1" & DEF_SEP & "%2"
        SendCommand eCmd, Printf(sFormat, sParam, ReceiptNumber)
        If LenB(m_uPrintData.LastError) <> 0 Then
            GoTo QH
        End If
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
    Dim eCmd            As UcsIslCommandsEnum
    Dim sParam          As String
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoDateTime", eCmd, ucsIslCmdInfoDateTime, sParam, vbNullString
    GetClock = C_Date(SendCommand(eCmd, sParam))
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
    SendCommand eCmd, Printf(sParam, Format$(NewDate, sFormat))
    If LenB(m_uPrintData.LastError) <> 0 Then
        GoTo QH
    End If
    '--- success
    SetClock = True
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function RunXReport(ByVal ReportType As UcsFiscalReportsTypeEnum) As Boolean
    Const FUNC_NAME     As String = "RunXReport"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsIslCommandsEnum
    Dim sParam          As String
    
    On Error GoTo EH
    If Not CancelReceipt() Then
        pvSetLastError pvInternal(ucsErrCannotCancelReceipt)
        GoTo QH
    End If
    Select Case ReportType
    Case ucsFscRstDaily
        pvGetCommandConfig CONF_NAME, "PrintDailyReport", eCmd, ucsIslCmdPrintDailyReport, sParam, "2"
    Case ucsFscRstDailyByItems
        pvGetCommandConfig CONF_NAME, "PrintDailyReportItems", eCmd, ucsIslCmdPrintDailyReportItems, sParam, "2"
    Case ucsFscRstDailyByDepartment
        If IsDaisy Then
            pvGetCommandConfig CONF_NAME, "ExtendedPrintDailyReport", eCmd, ucsIslCmdPrintDailyReport, sParam, "9"
        Else
            pvGetCommandConfig CONF_NAME, "PrintDailyReportDepartments", eCmd, ucsIslCmdPrintDailyReportDepartments, sParam, "2"
        End If
    Case ucsFscRstDailyByItemsAndDepartment
        If IsDaisy Then
            pvGetCommandConfig CONF_NAME, "ExtendedPrintDailyReportItems", eCmd, ucsIslCmdPrintDailyReportItems, sParam, "9"
        Else
            pvGetCommandConfig CONF_NAME, "PrintDailyReportItemsDepartments", eCmd, ucsIslCmdPrintDailyReportItemsDepartments, sParam, "2"
        End If
    Case Else
        pvSetLastError pvInternal(ucsErrInvalidReportType)
        GoTo QH
    End Select
    SendCommand eCmd, sParam
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
    Dim eCmd            As UcsIslCommandsEnum
    Dim sParam          As String
    
    On Error GoTo EH
    If Not CancelReceipt() Then
        pvSetLastError pvInternal(ucsErrCannotCancelReceipt)
        GoTo QH
    End If
    Select Case ReportType
    Case ucsFscRstDaily
        pvGetCommandConfig CONF_NAME, "PrintDailyReport", eCmd, ucsIslCmdPrintDailyReport, sParam, "0"
    Case ucsFscRstDailyByItems
        pvGetCommandConfig CONF_NAME, "PrintDailyReportItems", eCmd, ucsIslCmdPrintDailyReportItems, sParam, "0"
    Case ucsFscRstDailyByDepartment
        If IsDaisy Then
            pvGetCommandConfig CONF_NAME, "ExtendedPrintDailyReport", eCmd, ucsIslCmdPrintDailyReport, sParam, "8"
        Else
            pvGetCommandConfig CONF_NAME, "PrintDailyReportDepartments", eCmd, ucsIslCmdPrintDailyReportDepartments, sParam, "0"
        End If
    Case ucsFscRstDailyByItemsAndDepartment
        If IsDaisy Then
            pvGetCommandConfig CONF_NAME, "ExtendedPrintDailyReportItems", eCmd, ucsIslCmdPrintDailyReportItems, sParam, "8"
        Else
            pvGetCommandConfig CONF_NAME, "PrintDailyReportItemsDepartments", eCmd, ucsIslCmdPrintDailyReportItemsDepartments, sParam, "0"
        End If
    Case Else
        pvSetLastError Printf(pvInternal(ucsErrInvalidReportType), ReportType)
        GoTo QH
    End Select
    SendCommand eCmd, sParam
    If IsDaisy Then
        pvGetCommandConfig CONF_NAME, "ExtendedOperatorsReset", eCmd, ucsIslCmdInitOperatorReset, sParam, "0,9999"
        SendCommand eCmd, sParam
        If LenB(m_uPrintData.LastError) <> 0 Then
            GoTo QH
        End If
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
        pvGetCommandConfig CONF_NAME, "PrintReportByDateShort", eCmd, ucsIslCmdPrintReportByDateShort, sParam, "%1" & DEF_SEP & "%2", _
            "Format", sFormat, FORMAT_DATE_REPORT
    Case ucsFscRstPeriodDetailed
        pvGetCommandConfig CONF_NAME, "PrintReportByDateDetailed", eCmd, ucsIslCmdPrintReportByDateDetailed, sParam, "%1" & DEF_SEP & "%2", _
            "Format", sFormat, FORMAT_DATE_REPORT
    Case Else
        pvSetLastError Printf(pvInternal(ucsErrInvalidReportType), ReportType)
        GoTo QH
    End Select
    sParam = Printf(sParam, Format$(StartDate, sFormat), Format$(EndDate, sFormat))
    SendCommand eCmd, sParam
    If LenB(m_uPrintData.LastError) <> 0 Then
        GoTo QH
    End If
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
    Dim eCmd            As UcsIslCommandsEnum
    Dim sParam          As String
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoLastReceipt", eCmd, ucsIslCmdInfoLastReceipt, sParam, vbNullString
    GetLastReceiptNo = SendCommand(eCmd, sParam)
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
    Dim lOffset         As Long
    
    On Error GoTo EH
    If IsDaisy Then
        pvGetCommandConfig CONF_NAME, "ExtendedInitSetting", eCmd, ucsIslCmdExtendedInitSetting, sParam, "R17", _
            "Index", lIndex, 1
    ElseIf IsIncotex Then
        pvGetCommandConfig CONF_NAME, "InfoSums", eCmd, ucsIslCmdInfoSums, sParam, vbNullString, _
            "Index", lIndex, 10, _
            "Offset", lOffset, -1
    Else
        pvGetCommandConfig CONF_NAME, "InitInvoiceNo", eCmd, ucsIslCmdInitInvoiceNo, sParam, vbNullString, _
            "Index", lIndex, 0
    End If
    GetLastInvoiceNo = At(Split(SendCommand(eCmd, sParam), DEF_SEP), lIndex)
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
        vSplit = Split(SendCommand(eCmd, sParam), sDelimiter)
        If LenB(m_uPrintData.LastError) = 0 Then
            LastReceiptDate = C_Date(At(vSplit, lDateIndex))
            GetLastQRCodeInfo = RTrim$(At(vSplit, lIndex))
        End If
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
    Dim eCmd            As UcsIslCommandsEnum
    Dim sParam          As String
    Dim lIndex          As Long
    Dim vSplit          As Variant
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoBulstat", eCmd, ucsIslCmdInfoBulstat, sParam, vbNullString, _
        "Index", lIndex, IIf(IsIncotex, -1, 0)
    vSplit = Split(SendCommand(eCmd, sParam), DEF_SEP)
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
    GetTaxCaption = RTrim$(At(Split(SendCommand(eCmd, sParam), DEF_SEP), lIndex))
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
    If Index >= 1 And Index <= 6 Then
        pvGetCommandConfig CONF_NAME, "InitHeaderFooter" & Index, eCmd, ucsIslCmdInitHeaderFooter, sParam, "I" & (Index - 1)
        GetHeaderText = RTrim$(pvStripBold(SendCommand(eCmd, sParam)))
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
    If Index >= 1 And Index <= 2 Then
        pvGetCommandConfig CONF_NAME, "InitHeaderFooter" & Index, eCmd, ucsIslCmdInitHeaderFooter, sParam, "I" & (Index + 5)
        GetFooterText = RTrim$(pvStripBold(SendCommand(eCmd, sParam)))
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
    SendCommand eCmd, sParam
    If LenB(m_uPrintData.LastError) <> 0 Then
        GoTo QH
    End If
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
    If C_Lng(At(Split(SendCommand(eTransCmd, sTransParam), DEF_SEP), lTransIndex)) <> lTransStatus Then
        If Status(ucsIslStbFiscalPrinting) Then
            If IsDaisy Or IsIncotex Then
                pvGetCommandConfig CONF_NAME, "ExtendedFiscalCancel", eCmd, ucsIslCmdExtendedFiscalCancel, sParam, vbNullString
            Else
                '--- note: FP-3530 can cancel always, FP-550 sometimes can not
                pvGetCommandConfig CONF_NAME, "FiscalCancel", eCmd, ucsIslCmdFiscalCancel, sParam, vbNullString
            End If
            SendCommand eCmd, sParam
            If Status(ucsIslStbFiscalPrinting) Then
                '--- note: this helps if sum payed is less than total receipt sum
                pvGetCommandConfig CONF_NAME, "FiscalPayment", eCmd, ucsIslCmdFiscalPayment, sParam, vbTab
                SendCommand eCmd, sParam
                '--- note: when printing invoice, if no contragent info set then cancel fails!
                pvGetCommandConfig CONF_NAME, "FiscalCgInfo", eCmd, ucsIslCmdFiscalCgInfo, sParam, "0000000000" & IIf(IsEcr, vbTab & "0", vbNullString)
                SendCommand eCmd, sParam
                '--- for FP-550
                pvGetCommandConfig CONF_NAME, "FiscalClose", eCmd, ucsIslCmdFiscalClose, sParam, vbNullString
                SendCommand eCmd, sParam
            End If
        Else
            pvGetCommandConfig CONF_NAME, "NonFiscalClose", eCmd, ucsIslCmdNonFiscalClose, sParam, vbNullString
            SendCommand eCmd, sParam
        End If
        '--- success (or failure)
        CancelReceipt = C_Lng(At(Split(SendCommand(eTransCmd, sTransParam), DEF_SEP), lTransIndex)) = lTransStatus
    Else
        CancelReceipt = True
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetDeviceProtocol() As String
    Const FUNC_NAME     As String = "GetDeviceProtocol"
    
    On Error GoTo EH
    pvGetDeviceInfo m_oConnector, m_lTimeout, m_vInfoDiagnostics, GetDeviceProtocol, vbNullString, vbNullString
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetDeviceModel() As String
    Const FUNC_NAME     As String = "GetDeviceModel"
    
    On Error GoTo EH
    pvGetDeviceInfo m_oConnector, m_lTimeout, m_vInfoDiagnostics, vbNullString, GetDeviceModel, vbNullString
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetDeviceFirmware() As String
    Const FUNC_NAME     As String = "GetDeviceFirmware"
    
    On Error GoTo EH
    pvGetDeviceInfo m_oConnector, m_lTimeout, m_vInfoDiagnostics, vbNullString, vbNullString, GetDeviceFirmware
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
    pvGetCommandConfig CONF_NAME, "InfoStatus", eCmd, ucsIslCmdInfoStatus, sParam, "X"
    pvSendCommand m_oConnector, m_lTimeout, eCmd, sParam, m_uPrintData.LastError
    If LenB(m_uPrintData.LastError) <> 0 Then
        '--- do nothing
    ElseIf Status(ucsIslStbPrintingError) Then
        pvSetLastError Zn(ErrorText, Printf(pvInternal(ucsErrUnknownError), "0x" & Hex$(m_eStatus)))
    ElseIf Memory(ucsIslFmeMemoryError) Then
        pvSetLastError Zn(MemoryText, Printf(pvInternal(ucsErrUnknownMemory), "0x" & Hex$(m_eMemory)))
    End If
    CurrentStatus = m_uPrintData.LastError
    '--- success
    GetDeviceStatus = True
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetTotalsByPaymentTypes() As Variant
    Const FUNC_NAME     As String = "GetTotalsByPaymentTypes"
    Const CONF_NAME     As String = FUNC_NAME
    Dim vPmtSum         As Variant
    Dim vRevSum         As Variant
    Dim vRetVal         As Variant
    Dim lIdx            As Long
    Dim eCmd            As UcsIslCommandsEnum
    Dim sParam          As String
    Dim lPmtIndex       As Long
    Dim lOffset         As Long
    Dim lRevOffset      As Long
    
    On Error GoTo EH
    If Not IsEcr Then
        pvGetCommandConfig CONF_NAME, "InfoTotalsByPayments", eCmd, ucsIslCmdInfoTotalsByPayments, sParam, IIf(IsDaisy, "A", vbNullString), _
            "Offset", lOffset, 0
        vPmtSum = Split(SendCommand(eCmd, sParam), DEF_SEP)
    Else
        pvGetCommandConfig CONF_NAME, "InfoTotalsByPayments0", eCmd, ucsIslCmdInfoTotalsByPayments, sParam, "0", _
            "Offset", lOffset, -10
        vPmtSum = Split(SendCommand(eCmd, sParam), DEF_SEP)
        pvGetCommandConfig CONF_NAME, "InfoTotalsByPayments1", eCmd, ucsIslCmdInfoTotalsByPayments, sParam, "1", _
            "RevOffset", lRevOffset, -10
        vRevSum = Split(SendCommand(eCmd, sParam), DEF_SEP)
    End If
    If UBound(vPmtSum) > 0 Then
        ReDim vRetVal(0 To MAX_PMT_TYPE - 1) As Variant
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
        GetTotalsByPaymentTypes = vRetVal
    Else
        GetTotalsByPaymentTypes = Array()
    End If
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
    Dim vRetVal         As Variant
    Dim lIdx            As Long
    Dim eCmd            As UcsIslCommandsEnum
    Dim sParam          As String
    Dim lOffset         As Long
    Dim lRevOffset      As Long
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoTaxRates", eCmd, ucsIslCmdInfoTaxRates, sParam, vbNullString
    vTaxRates = Split(Replace(SendCommand(eCmd, sParam), "%", vbNullString), DEF_SEP)
    If IsDaisy Or IsIncotex Then
        pvGetCommandConfig CONF_NAME, "InfoTotalsByTaxGroups0", eCmd, ucsIslCmdInfoTotalsByTaxGroups, sParam, "T", _
            "Offset", lOffset, 0
        vPmtSum = Split(SendCommand(eCmd, sParam), DEF_SEP)
        vRevSum = vPmtSum
        lRevOffset = 8
    Else
        pvGetCommandConfig CONF_NAME, "InfoTotalsByTaxGroups0", eCmd, ucsIslCmdInfoTotalsByTaxGroups, sParam, "0", _
            "Offset", lOffset, 0
        vPmtSum = Split(SendCommand(eCmd, sParam), DEF_SEP)
        pvGetCommandConfig CONF_NAME, "InfoTotalsByTaxGroups5", eCmd, ucsIslCmdInfoTotalsByTaxGroups, sParam, "5", _
            "RevOffset", lRevOffset, IIf(IsIncotex Or IsDaisy, 1, 0)
        vRevSum = Split(SendCommand(eCmd, sParam), DEF_SEP)
    End If
    If UBound(vPmtSum) > 0 Or UBound(vRevSum) > 0 Then
        ReDim vRetVal(0 To MAX_TAX_GROUP - 1) As Variant
        For lIdx = 0 To UBound(vRetVal)
            vRetVal(lIdx) = Array(C_Dbl(At(vTaxRates, lIdx)), ParseSum(At(vPmtSum, lIdx + lOffset)), ParseSum(At(vRevSum, lIdx + lRevOffset)))
        Next
        GetTotalsByTaxGroups = vRetVal
    Else
        GetTotalsByTaxGroups = Array()
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
    Dim lIdx            As Long
    
    On Error GoTo EH
    If Not IsEcr Then
        pvGetCommandConfig CONF_NAME, "InfoNap", eCmd, ucsIslCmdInfoNap, sParam, "1", _
            "Index", lIndex, 2
        vSplit = Split(SendCommand(eCmd, sParam), DEF_SEP)
        If UBound(vSplit) >= lIndex Then
            lReceipts = C_Lng(At(vSplit, lIndex))
        Else
            vSplit = Empty
        End If
    End If
    If Not IsArray(vSplit) Then
        pvGetCommandConfig CONF_NAME, "InfoOperator", eCmd, ucsIslCmdInfoOperator, sParam, "%1", _
            "Index", lIndex, 0
        lReceipts = 0
        For lIdx = 1 To 100
            vSplit = Split(SendCommand(eCmd, Printf(sParam, lIdx)), DEF_SEP)
            If LenB(m_uPrintData.LastError) <> 0 Then
                Exit For
            End If
            lReceipts = lReceipts + C_Lng(At(vSplit, lIndex))
        Next
    End If
    If IsIncotex Or IsEcr Then
        pvGetCommandConfig CONF_NAME, "InfoLastFiscalBlock", eCmd, ucsIslCmdInfoLastFiscalBlock, sParam, IIf(IsEcr, "*0", vbNullString), _
            "Index", lIndex, -1
        vSplit = Split(SendCommand(eCmd, sParam), DEF_SEP)
        dLastDate = pvToDate(At(vSplit, lIndex))
    Else
        pvGetCommandConfig CONF_NAME, "InfoLastFiscalDateTime", eCmd, ucsIslCmdInfoLastFiscalDateTime, sParam, "T"
        dLastDate = C_Date(SendCommand(eCmd, sParam))
    End If
    GetTotalsStats = Array(lReceipts, IIf(dLastDate = 0, Empty, dLastDate))
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function PrintServiceDeposit(ByVal Amount As Double) As Variant
    Const FUNC_NAME     As String = "PrintServiceDeposit"
    Const CONF_NAME     As String = FUNC_NAME
    Dim vSplit          As Variant
    Dim vRetVal         As Variant
    Dim lIdx            As Long
    Dim eCmd            As UcsIslCommandsEnum
    Dim sParam          As String
    Dim sFormat         As String
    
    On Error GoTo EH
    sFormat = IIf(Abs(Amount) > DBL_EPSILON, FORMAT_BASE_2, "\")
    pvGetCommandConfig CONF_NAME, "FiscalServiceDeposit" & -(Abs(Amount) > DBL_EPSILON), eCmd, ucsIslCmdFiscalServiceDeposit, sParam, "%1", _
        "Format", sFormat, sFormat
    sParam = Printf(sParam, SafeFormat(Amount, sFormat))
    vSplit = Split(SendCommand(eCmd, sParam), DEF_SEP)
    If UBound(vSplit) > 0 Then
        ReDim vRetVal(0 To 2) As Double
        For lIdx = 0 To UBound(vRetVal)
            vRetVal(lIdx) = ParseSum(At(vSplit, lIdx + 1))
        Next
        PrintServiceDeposit = vRetVal
    Else
        PrintServiceDeposit = EmptyDoubleArray
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
    GetOperatorName = RTrim$(At(Split(Replace(SendCommand(eCmd, sParam), vbTab, DEF_SEP), DEF_SEP), lIndex))
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetDefaultPassword(OperatorCode As String) As String
    If IsEcr Or IsDaisy Or IsIncotex Then
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
        vSplit = Split(Replace(SendCommand(eCmd, sParam), vbLf, DEF_SEP), DEF_SEP)
        If LenB(LastError) = 0 Then
            GetDepartmentName = RTrim$(At(vSplit, lIndex))
            TaxGroup = Limit(Asc(At(vSplit, lTaxIndex) & " ") - 191, 0, 8)
        End If
    Else
        pvGetCommandConfig CONF_NAME, "InfoOperator", eCmd, ucsIslCmdInfoDepartment, sParam, "%1", _
            "Index", lIndex, 5, _
            "TaxIndex", lTaxIndex, 0
        sParam = Printf(sParam, DepartmentCode)
        vSplit = Split(Mid$(SendCommand(eCmd, sParam), 2), DEF_SEP)
        If LenB(LastError) = 0 Then
            GetDepartmentName = RTrim$(At(vSplit, lIndex))
            TaxGroup = Limit(Asc(At(vSplit, lTaxIndex) & " ") - 191, 0, 8)
        End If
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
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
    Dim lSize           As Long
    Dim sBuffer         As String
    
    On Error GoTo EH
    DebugLog FUNC_NAME, "eCmd=" & eCmd & IIf(LenB(Data) <> 0, ", Data=" & Data, vbNullString)
    '--- prepare command and transcode cyrillic
    sSend = pvTranscode(Data, True)
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
    Do While Right$(sText, 1) <> Chr$(SYM_TERMINATOR)
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
        sText = sText & Replace(sBuffer, Chr$(SYM_SYN), vbNullString)
    Loop
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
        pvSendCommand = pvTranscode(Mid$(sText, 5, lSize - 4), False)
        DebugLog FUNC_NAME, "RetVal=" & pvSendCommand & _
            IIf(Status(ucsIslStbPrintingError), ", ErrorText=" & Zn(ErrorText, Printf(pvInternal(ucsErrUnknownError), "0x" & Hex$(m_eStatus))), vbNullString) & _
            IIf(Memory(ucsIslFmeMemoryError), ", MemoryText=" & Zn(MemoryText, Printf(pvInternal(ucsErrUnknownMemory), "0x" & Hex$(m_eMemory))), vbNullString)
    Else
        '--- check for GSM tax terminal interferences: "AT+CPIN?\r\n" or "AT+CIMI\r\n"
        lIdx = InStr(sText, Chr$(SYM_PREAMBLE))
        If lIdx > 1 Then
            DebugLog FUNC_NAME, "Trim bogus symbols, Prefix=0x" & ToHexDump(Left$(sText, lIdx - 1))
            sText = Mid$(sText, lIdx)
            GoTo RetryParse
        End If
        Error = Printf(pvInternal(ucsErrResponseMissingPreamble), "0x" & ToHexDump(sText))
        pvSetStatus vbNullString '--- clear
    End If
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function pvCheckSum(sText As String) As String
    Dim baText()        As Byte
    Dim lIdx            As Long
    Dim lSum            As Long
    
    baText = ToAscii(sText)
    For lIdx = 1 To UBound(baText)
        lSum = lSum + baText(lIdx)
    Next
    pvCheckSum = pvToWireLong(lSum)
End Function

Private Function pvToWireLong(ByVal lValue As Long) As String
    pvToWireLong = Chr$((lValue \ &H1000 Mod &H10) + &H30) & Chr$((lValue \ &H100 Mod &H10) + &H30) & Chr$((lValue \ &H10 Mod &H10) + &H30) & Chr$((lValue \ &H1 Mod &H10) + &H30)
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
            SendCommand eCmd, Zn(Printf(sParam, C_Str(vElem)), " ")
            If LenB(m_uPrintData.LastError) <> 0 Then
                GoTo QH
            End If
        Next
    Else
        SendCommand eCmd, Zn(Printf(sParam, C_Str(vText)), " ")
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
    Dim sRevType        As String
    Dim sTaxNoPrefix    As String
    Dim sTaxNoSuffix    As String
    Dim sCodeType       As String
    Dim sCodeData       As String
    Dim lMaxSize        As Long
    
    On Error GoTo EH
    '--- ring sale
    Do While uCtx.Row < lRowCount
        With uRow(uCtx.Row)
        Select Case .RowType
        Case ucsRowInit
            pvGetCommandConfig CONF_NAME, "InfoOperator", eCmd, ucsIslCmdInfoOperator, sParam, "%1", _
                "Index", lIndex, 5
            sParam = Printf(sParam, .InitOperatorCode)
            sData = At(Split(SendCommand(eCmd, sParam), DEF_SEP), lIndex)
            If LenB(m_uPrintData.LastError) = 0 Then
                If LenB(.InitOperatorName) <> 0 Then
                    If .InitOperatorName <> sData Then
                        sParam = "%1" & DEF_SEP & "%2" & DEF_SEP & "%3" & IIf(IsDaisy, vbTab & "1,1,1", vbNullString)
                        pvGetCommandConfig CONF_NAME, "InitOperatorName", eCmd, ucsIslCmdInitOperatorName, sParam, sParam, _
                            "MinName", lMinName, 10
                        sOperParamMax = Printf(sParam, .InitOperatorCode, .InitOperatorPassword, Left$(.InitOperatorName, m_uConfig.MaxOperatorName))
                        sOperParamMin = Printf(sParam, .InitOperatorCode, .InitOperatorPassword, Left$(.InitOperatorName, LimitLong(Len(sData), lMinName)))
                        SendCommand eCmd, sOperParamMax
                        If LenB(m_uPrintData.LastError) <> 0 And sOperParamMin <> sOperParamMax Then
                            SendCommand eCmd, sOperParamMin
                            '--- skip last error check
                        End If
                    End If
                Else
                    .InitOperatorName = RTrim$(sData)
                End If
            End If
            If LenB(.InitInvData(ucsInvDocNo)) <> 0 Then
                Select Case .InitReceiptType
                Case ucsFscRcpInvoice, ucsFscRcpCreditNote
                    If IsDaisy Then
                        pvGetCommandConfig CONF_NAME, "ExtendedInitSetting", eCmd, ucsIslCmdExtendedInitSetting, sParam, "P%1" & DEF_SEP & "%2", _
                            "StartIndex", lIndex, 19, _
                            "EndIndex", lEndIndex, 20
                        If lIndex > 0 Then
                            SendCommand eCmd, Printf(sParam, lIndex, .InitInvData(ucsInvDocNo))
                        End If
                        If lEndIndex > 0 Then
                            SendCommand eCmd, Printf(sParam, lEndIndex, .InitInvData(ucsInvDocNo))
                        End If
                    ElseIf IsIncotex Then
                        pvGetCommandConfig CONF_NAME, "ExtendedInitInvoiceNo", eCmd, ucsIslCmdExtendedInitInvoiceNo, sParam, "%1" & vbLf & "%1"
                        sParam = Printf(sParam, .InitInvData(ucsInvDocNo))
                        SendCommand eCmd, sParam
                    Else
                        pvGetCommandConfig CONF_NAME, "InitInvoiceNo", eCmd, ucsIslCmdInitInvoiceNo, sParam, "%1" & DEF_SEP & "%1"
                        sParam = Printf(sParam, .InitInvData(ucsInvDocNo))
                        SendCommand eCmd, sParam
                    End If
                    If LenB(m_uPrintData.LastError) <> 0 Then
                        GoTo QH
                    End If
                End Select
            End If
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
                        pvGetCommandConfig CONF_NAME, "FiscalNoPrintReversal" & .InitReceiptType, eCmd, ucsIslCmdFiscalNoPrintOpen, sParam, sParam, _
                            "FormatDate", sFormatDate, FORMAT_DATE_EXTENDED
                    Else
                        pvGetCommandConfig CONF_NAME, "FiscalReversal" & .InitReceiptType, eCmd, ucsIslCmdFiscalOpen, sParam, sParam, _
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
                        pvGetCommandConfig CONF_NAME, "FiscalNoPrintOpen" & .InitReceiptType, eCmd, ucsIslCmdFiscalNoPrintOpen, sParam, sParam
                    Else
                        pvGetCommandConfig CONF_NAME, "FiscalOpen" & .InitReceiptType, eCmd, ucsIslCmdFiscalOpen, sParam, sParam
                    End If
                    sParam = Printf(sParam, .InitOperatorCode, .InitOperatorPassword, .InitUniqueSaleNo)
                End Select
            ElseIf IsIncotex Then
                '--- %1=<OpNum>,%2=<UNPstring>,%3=<RevDocNo>[,%4=<F1>[%5=(<F2><RevInvoiceNo>,<dd-mm-yy hh:mm:ss>,origDevFMNo)]]
                Select Case .InitReceiptType
                Case ucsFscRcpReversal, ucsFscRcpCreditNote
                    sRevType = Switch( _
                        .InitRevData(ucsRevType) = ucsFscRevTaxBaseReduction, "V", _
                        .InitRevData(ucsRevType) = ucsFscRevOperatorError, "R", _
                        True, "S")
                    sParam = "%1" & DEF_SEP & "%2" & DEF_SEP & "%3" & DEF_SEP & "%4"
                Case Else
                    sParam = "%1" & DEF_SEP & "%2" & DEF_SEP & "0" & IIf(.InitReceiptType = ucsFscRcpInvoice, DEF_SEP & "I", vbNullString)
                End Select
                pvGetCommandConfig CONF_NAME, "FiscalOpen" & .InitReceiptType, eCmd, ucsIslCmdFiscalOpen, sParam, sParam, _
                    "FormatDate", sFormatDate, FORMAT_DATE_EXTENDED
                sParam = Printf(sParam, .InitOperatorCode, _
                    Zn(.InitUniqueSaleNo, "0"), _
                    .InitRevData(ucsRevReceiptNo), _
                    sRevType, _
                    sRevType & Zn(At(.InitRevData, ucsRevInvoiceNo), "0") & _
                        DEF_SEP & Format$(.InitRevData(ucsRevReceiptDate), sFormatDate) & _
                        DEF_SEP & .InitRevData(ucsRevFiscalMemoryNo))
            ElseIf IsEcr Then
                Select Case .InitReceiptType
                Case ucsFscRcpReversal, ucsFscRcpCreditNote
                    '--- %1=<OpCode>,%2=<OpPwd>,%3=<NSale>,<TillNmb>,%4=(<DocType>,<DocNumber>,<DocDateTime>,<FMNumber>)[,<Invoice>,%5=(<InvNumber>,<Reason>)]
                    sParam = "%1" & DEF_SEP & "%2" & DEF_SEP & "%3" & DEF_SEP & "1" & DEF_SEP & "%4" & IIf(.InitReceiptType = ucsFscRcpCreditNote, DEF_SEP & "I" & DEF_SEP & "%5", vbNullString)
                    pvGetCommandConfig CONF_NAME, "FiscalReversal" & .InitReceiptType, eCmd, ucsIslCmdFiscalReversal, sParam, sParam, _
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
                    pvGetCommandConfig CONF_NAME, "FiscalOpen" & .InitReceiptType, eCmd, ucsIslCmdFiscalOpen, sParam, sParam
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
                    '--- %1=<OpNum>,%2=<Password>,<TillNum>[,<Invoice>%3=<InvNum>][,<UNP>],%4=(<StTipe><DocNo>)%5=([,<StUNP>,<StDT>,<StFMIN>])[#<RevType>]
                    sParam = "%1" & DEF_SEP & "%2" & DEF_SEP & "1" & IIf(.InitReceiptType = ucsFscRcpCreditNote, DEF_SEP & "I" & DEF_SEP & "%3", vbNullString) & DEF_SEP & "%4%5"
                    pvGetCommandConfig CONF_NAME, "FiscalReversal" & .InitReceiptType, eCmd, ucsIslCmdFiscalReversal, sParam, sParam, _
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
                    pvGetCommandConfig CONF_NAME, "FiscalOpen" & .InitReceiptType, eCmd, ucsIslCmdFiscalOpen, sParam, sParam
                    sParam = Printf(sParam, .InitOperatorCode, .InitOperatorPassword, _
                        IIf(LenB(.InitUniqueSaleNo) <> 0, DEF_SEP & .InitUniqueSaleNo, vbNullString))
                End Select
            End If
            bInLogin = True
            SendCommand eCmd, sParam
            If LenB(m_uPrintData.LastError) <> 0 Then
                GoTo QH
            End If
            bInLogin = False
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
            '--- [<L1>][<Lf><L2>]<Tab><TaxCd><[Sign]Price>[*<Qwan>][,Perc|;Abs]
            '--- [<L1>][<Lf><L2>]<Tab><Dept><Tab><[Sign]Price>[*<Qwan>][,Perc|;Abs]
            pvGetCommandConfig CONF_NAME, "FiscalSell" & .PrintRowType, eCmd, ucsIslCmdFiscalSell, sParam, "%1", _
                "FormatPrice", sFormatPrice, FORMAT_BASE_2, _
                "FormatQty", sFormatQty, FORMAT_BASE_3, _
                "FormatDisc", sFormatDisc, FORMAT_BASE_2
            vSplit = WrapText(.PluItemName & IIf(LenB(.PluUnitOfMeasure) <> 0, " (" & .PluUnitOfMeasure & ")", vbNullString), _
                LimitLong(pvCommentChars, , pvItemChars))
            For lIdx = 0 To UBound(vSplit) - 2
                If Not pvPrintLine(C_Str(vSplit(lIdx))) Then
                    GoTo QH
                End If
            Next
            sData = At(vSplit, UBound(vSplit))
            If UBound(vSplit) > 0 Then
                sData = At(vSplit, UBound(vSplit) - 1) & vbLf & sData
            End If
            If .PluDepartmentNo <> 0 Then
                sData = sData & vbTab & .PluDepartmentNo & vbTab
            Else
                sData = sData & vbTab & Chr$(191 + .PluTaxGroup)
            End If
            sData = sData & SafeFormat(.PluPrice, sFormatPrice)
            If Abs(.PluQuantity - 1) > DBL_EPSILON Then
                sData = sData & "*" & SafeFormat(.PluQuantity, sFormatQty)
            End If
            Select Case .DiscType
            Case ucsFscDscPlu
                If Abs(.PluPrice) > DBL_EPSILON Then
                    sData = sData & "," & SafeFormat(.DiscValue, sFormatDisc)
                End If
            Case ucsFscDscPluAbs
                sData = sData & ";" & SafeFormat(.DiscValue, sFormatDisc)
            End Select
            sParam = Printf(sParam, sData)
            SendCommand eCmd, sParam
            If LenB(m_uPrintData.LastError) <> 0 Then
                GoTo QH
            End If
        Case ucsRowDiscount
            Select Case .PrintRowType
            Case ucsFscRcpNonfiscal
                If .DiscType = ucsFscDscSubtotal Then
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
                End If
            Case Else
                Select Case .DiscType
                Case ucsFscDscSubtotal
                    pvGetCommandConfig CONF_NAME, "FiscalSubtotal" & .DiscType, eCmd, ucsIslCmdFiscalSubtotal, sParam, "10" & DEF_SEP & "%1", _
                        "FormatDisc", sFormatDisc, FORMAT_BASE_2
                Case ucsFscDscSubtotalAbs
                    pvGetCommandConfig CONF_NAME, "FiscalSubtotal" & .DiscType, eCmd, ucsIslCmdFiscalSubtotal, sParam, "10" & ";" & "%1", _
                        "FormatDisc", sFormatDisc, FORMAT_BASE_2
                End Select
                sParam = Printf(sParam, SafeFormat(.DiscValue, sFormatDisc))
                SendCommand eCmd, sParam
                If LenB(m_uPrintData.LastError) <> 0 Then
                    GoTo QH
                End If
            End Select
            If .DiscType = ucsFscDscSubtotal Then
                For lIdx = 1 To UBound(uCtx.GrpTotal)
                    If Abs(uCtx.GrpTotal(lIdx)) > DBL_EPSILON Then
                        dblTotal = Round(uCtx.GrpTotal(lIdx) * .DiscValue / 100#, DEF_PRICE_SCALE)
                        uCtx.GrpTotal(lIdx) = Round(uCtx.GrpTotal(lIdx) + dblTotal, DEF_PRICE_SCALE)
                    End If
                Next
            End If
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
                sParam = Printf(sParam, LimitLong(.BarcodeHeight * 6, 24, 240))
                SendCommand eCmd, sParam
                '--- skip last error check
            End If
            sCodeData = .BarcodeText
            Select Case .BarcodeType
            Case ucsFscBrcEan8
                sCodeType = "1"
                sCodeData = Pad(sCodeData, 7, "0")
            Case ucsFscBrcEan13
                sCodeType = "2"
                sCodeData = Pad(sCodeData, 12, "0")
            Case ucsFscBrcCode128
                sCodeType = "3"
            Case ucsFscBrcQRcode
                If IsDaisy Or IsEcr Then
                    pvSetLastError Printf("Barcode type %1 not supported", .BarcodeType)
                    GoTo QH
                End If
                sCodeType = "Q"
            End Select
            If IsDaisy Then
                sParam = "%1" & DEF_SEP & "%2" & vbTab & "C,0,%3,0"
            ElseIf IsEcr Then
                sParam = "%1" & ";" & "%2"
            Else
                sParam = "%1" & "C;" & "%2"
            End If
            pvGetCommandConfig CONF_NAME, "PrintBarcode", eCmd, ucsIslCmdPrintBarcode, sParam, sParam
            sParam = Printf(sParam, sCodeType, sCodeData, .BarcodeHeight)
            SendCommand eCmd, sParam
            If LenB(m_uPrintData.LastError) <> 0 Then
                GoTo QH
            End If
        Case ucsRowPayment
            If .PmtType <> 0 Then
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
                SendCommand eCmd, sParam
                If LenB(m_uPrintData.LastError) <> 0 Then
                    GoTo QH
                End If
                uCtx.PmtPrinted = True
            Else
                '--- check if anything left for payment
                pvGetCommandConfig CONF_NAME, "InfoTransaction", eCmd, ucsIslCmdInfoTransaction, sParam, "T", _
                    "Index", lIndex, 0, _
                    "Status", lStatus, 0, _
                    "AmountIndex", lAmountIndex, 2, _
                    "PaidIndex", lPaidIndex, 3
                vSplit = Split(SendCommand(eCmd, sParam), DEF_SEP)
                If LenB(m_uPrintData.LastError) <> 0 Then
                    GoTo QH
                End If
                If C_Lng(At(vSplit, lIndex)) <> lStatus Then
                    If C_Dbl(At(vSplit, lAmountIndex)) > C_Dbl(At(vSplit, lPaidIndex)) _
                            Or LenB(At(vSplit, lAmountIndex)) = 0 _
                            Or Not uCtx.PmtPrinted Then
                        pvGetCommandConfig CONF_NAME, "FiscalPayment" & .PmtType, eCmd, ucsIslCmdFiscalPayment, sParam, vbTab
                        SendCommand eCmd, sParam
                        If LenB(m_uPrintData.LastError) <> 0 Then
                            GoTo QH
                        End If
                        uCtx.PmtPrinted = True
                    End If
                    Select Case .PrintRowType
                    Case ucsFscRcpInvoice, ucsFscRcpCreditNote
                        pvGetCommandConfig CONF_NAME, "FiscalCgInfo", eCmd, ucsIslCmdFiscalCgInfo, sParam, Replace("%1#%2#%3#%4#%5#%6", "#", vbTab), _
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
                            lMaxSize = m_uConfig.MaxCgInfo(0) - Len(sTaxNoPrefix) - Len(sTaxNoSuffix)
                            sParam = Printf(sParam, sTaxNoPrefix & Left$(.InitInvData(ucsInvCgTaxNo), lMaxSize) & sTaxNoSuffix, _
                                Left$(.InitOperatorName, m_uConfig.MaxCgInfo(1)), _
                                Left$(.InitInvData(ucsInvCgPrsReceive), m_uConfig.MaxCgInfo(2)), _
                                Left$(.InitInvData(ucsInvCgName), m_uConfig.MaxCgInfo(3)), _
                                Left$(.InitInvData(ucsInvCgVatNo), m_uConfig.MaxCgInfo(4)), _
                                sFormatAddr)
                        End With
                        SendCommand eCmd, sParam
                        If LenB(m_uPrintData.LastError) <> 0 Then
                            GoTo QH
                        End If
                    End Select
                    pvGetCommandConfig CONF_NAME, "FiscalClose" & .PrintRowType, eCmd, ucsIslCmdFiscalClose, sParam, vbNullString
                    SendCommand eCmd, sParam
                    If LenB(m_uPrintData.LastError) <> 0 Then
                        GoTo QH
                    End If
                End If
                pvGetCommandConfig CONF_NAME, "InfoStatus", eCmd, ucsIslCmdInfoStatus, sParam, "W"
                SendCommand eCmd, sParam
                '--- skip last error check
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

Private Function pvGetDeviceInfo(oConnector As IDeviceConnector, ByVal lTimeout As Long, vResult As Variant, sProtocol As String, sModel As String, sFirmware As String, Optional Error As String) As Boolean
    Const FUNC_NAME     As String = "pvGetDeviceInfo"
    Const CONF_NAME     As String = "GetDeviceInfo"
    Dim eCmd            As UcsIslCommandsEnum
    Dim sParam          As String
    Dim sDaisyPrefix    As String
    Dim sIncotexPrefix  As String
    Dim lNameIndex      As Long
    Dim lSerialIndex    As Long
    Dim vSplit          As Variant
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoDiagnostics", eCmd, ucsIslCmdInfoDiagnostics, sParam, "0", _
        "DaisyPrefix", sDaisyPrefix, "DY", _
        "IncotexPrefix", sIncotexPrefix, "IN", _
        "NameIndex", lNameIndex, 0, _
        "SerialIndex", lSerialIndex, 4
    If Not IsArray(vResult) Then
        vResult = Split(pvSendCommand(oConnector, lTimeout, eCmd, sParam, Error), DEF_SEP)
        If LenB(At(vResult, 0)) = 0 Then
            sParam = vbNullString
            vResult = Split(pvSendCommand(oConnector, lTimeout, eCmd, sParam, Error), DEF_SEP)
        End If
    End If
    If UBound(vResult) >= lSerialIndex And LenB(Error) = 0 Then
        '--- check serial
        If Left$(At(vResult, lSerialIndex), Len(sDaisyPrefix)) = sDaisyPrefix And LenB(sDaisyPrefix) <> 0 Then
            sProtocol = STR_PROTOCOL_DAISY
        ElseIf Left$(At(vResult, lSerialIndex), Len(sIncotexPrefix)) = sIncotexPrefix And LenB(sIncotexPrefix) <> 0 Then
            sProtocol = STR_PROTOCOL_INCOTEX
        ElseIf LenB(At(vResult, lSerialIndex)) <> 0 Then
            sProtocol = STR_PROTOCOL_DATECS_FP
        End If
        vSplit = Split2(RTrim$(At(vResult, lNameIndex)) & RTrim$(" " & At(vResult, lNameIndex + 1)), " ")
        sModel = RTrim$(vSplit(0))
        sFirmware = RTrim$(vSplit(1))
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

Private Function pvGetPmtMapping() As Object
    Const CONF_NAME     As String = "GetPmtMapping"
    Dim eCmd            As UcsIslCommandsEnum
    Dim sParam          As String
    Dim lIdx            As Long
    Dim vConfig         As Variant
    Dim lOffset         As Long
    Dim lTagNo          As Long
    Dim vSplit          As Variant
    Dim oRetVal         As Object
    
    If IsDaisy Or IsIncotex Then
        pvGetCommandConfig CONF_NAME, "ExtendedInitPayment", eCmd, ucsIslCmdExtendedInitPayment, sParam, "R%1", _
            "Offset", lOffset, 3
        For lIdx = 0 To 4
            vSplit = Split(Replace(SendCommand(eCmd, Printf(sParam, lIdx)), vbTab, DEF_SEP), DEF_SEP)
            If lIdx > 0 Then
                lTagNo = C_Lng(At(vSplit, lOffset))
            End If
            If IsEmpty(JsonItem(oRetVal, "#" & lTagNo)) Then
                JsonItem(oRetVal, "#" & lTagNo) = Array(lIdx, At(vSplit, 1))
            End If
        Next
    ElseIf IsEcr Then
        pvGetCommandConfig CONF_NAME, "EcrReadRow", eCmd, ucsIslCmdEcrReadRow, sParam, "21;1;", _
            "Offset", lOffset, 9
        vConfig = Split(SendCommand(eCmd, sParam), ";")
        pvGetCommandConfig CONF_NAME, "EcrReadRow", eCmd, ucsIslCmdEcrReadRow, sParam, "3;%1;"
        For lIdx = 0 To 5
            If lIdx > 0 Then
                lTagNo = At(vConfig, lOffset + lIdx)
            End If
            If IsEmpty(JsonItem(oRetVal, "#" & lTagNo)) Then
                vSplit = Split(SendCommand(eCmd, Printf(sParam, lIdx + 1)), ";")
                '--- 1 - disabled, 0 - not disabled
                If C_Lng(At(vSplit, 3)) = 0 Then
                    JsonItem(oRetVal, "#" & lTagNo) = Array(lIdx, At(vSplit, 2))
                End If
            End If
        Next
    End If
    Set pvGetPmtMapping = oRetVal
End Function

Private Function pvToDevicePmtIndex(ByVal PmtType As UcsFiscalPaymentTypeEnum) As Long
    Dim vInfo           As Variant
    
    pvToDevicePmtIndex = -1             ' unsupported
    If Not IsObject(m_uConfig.PmtMapping) Then
        Set m_uConfig.PmtMapping = pvGetPmtMapping()
    End If
    If Not m_uConfig.PmtMapping Is Nothing Then
        AssignVariant vInfo, JsonItem(C_Obj(m_uConfig.PmtMapping), "#" & (PmtType - ucsFscPmtCash))
        If IsObject(vInfo) Then
            pvToDevicePmtIndex = C_Lng(JsonItem(C_Obj(vInfo), 0))
        End If
    ElseIf IsDaisy Or IsIncotex Then
        Select Case PmtType
        Case ucsFscPmtCash
            pvToDevicePmtIndex = 0      ' P=SCash (В БРОЙ)
        Case ucsFscPmtCheque
            pvToDevicePmtIndex = 1      ' N=SCheck (С ЧЕК)
        Case ucsFscPmtCard
            pvToDevicePmtIndex = 2      ' C=SCards (КРЕДИТ)
        Case ucsFscPmtBank
'            pvToDevicePmtIndex = 3      ' D=SCash (USD)
            pvToDevicePmtIndex = -1     ' unsuported
        Case ucsFscPmtEUR
            pvToDevicePmtIndex = 4      ' B=SCash (EURO)
        End Select
    ElseIf IsEcr Then
        Select Case PmtType
        Case ucsFscPmtCash
            pvToDevicePmtIndex = 0      ' P=SCash (В БРОЙ)
        Case ucsFscPmtCoupon
            pvToDevicePmtIndex = 5      ' I=ST (КУПОН)
        Case ucsFscPmtVoucher
            pvToDevicePmtIndex = 4      ' J=SOT (ВАУЧЕР)
        Case ucsFscPmtCard
            pvToDevicePmtIndex = 2      ' C=SCards (ДЕБ.КАРТА)
        Case ucsFscPmtBank
'            pvToDevicePmtIndex = 1      ' N=SCards (КРЕДИТ)
            pvToDevicePmtIndex = -1     ' unsuported
        Case ucsFscPmtCustom1
            pvToDevicePmtIndex = 3      ' D=SR1 (НЗОК)
        End Select
    Else
        Select Case PmtType
        Case ucsFscPmtCash
            pvToDevicePmtIndex = 0      ' P=SCash (В брой)
        Case ucsFscPmtCheque
            pvToDevicePmtIndex = 2      ' C=SChecks (С чек)
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
        Case ucsFscPmtCard
            pvToDevicePmtIndex = 3      ' D=SCards (С дебитна карта)
        Case ucsFscPmtBank
            pvToDevicePmtIndex = 15     ' r=SW (Банкови трансфери)
        Case ucsFscPmtCustom1
            pvToDevicePmtIndex = 4      ' I=SR1 (Програмируем тип плащане 1)
        Case ucsFscPmtCustom2
            pvToDevicePmtIndex = 7      ' L=SR2 (Програмируем тип плащане 2)
        End Select
    End If
End Function

Private Function pvToDevicePmtName(ByVal PmtType As UcsFiscalPaymentTypeEnum) As String
    Const CONF_NAME     As String = "GetPaymentName"
    Dim eCmd            As UcsIslCommandsEnum
    Dim sParam          As String
    Dim lIndex          As Long
    
    lIndex = pvToDevicePmtIndex(PmtType)
    If lIndex >= 0 Then
        If IsDaisy Then
            pvGetCommandConfig CONF_NAME, "ExtendedInitText" & PmtType, eCmd, ucsIslCmdExtendedInitText, sParam, "R" & (60 + lIndex)
        ElseIf IsIncotex Then
            If PmtType >= 7 Then
                GoTo QH
            End If
            pvGetCommandConfig CONF_NAME, "ExtendedInitText" & PmtType, eCmd, ucsIslCmdExtendedInitText, sParam, "R" & (7 + lIndex)
        Else
            pvGetCommandConfig CONF_NAME, "InitPaymentType" & PmtType, eCmd, ucsIslCmdInitPaymentType, sParam, Mid$(DefPmtTypes, lIndex + 1, 1)
        End If
        pvToDevicePmtName = RTrim$(pvStripBold(SendCommand(eCmd, sParam)))
        '--- check for failure (returns 'F')
        If pvToDevicePmtName = "F" Then
            If lIndex < 4 Then
                pvToDevicePmtName = pvText(ucsTxtPaymentCash + lIndex)
            Else
                pvToDevicePmtName = vbNullString
            End If
        End If
    End If
QH:
End Function

Private Sub pvGetCommandConfig( _
            sFunction As String, _
            sKey As String, _
            eCmd As UcsIslCommandsEnum, _
            ByVal eCmdDefault As UcsIslCommandsEnum, _
            sParam As String, _
            sParamDefault As String, _
            ParamArray A() As Variant)
    Dim lIdx            As Long
    
    eCmd = GetConfigForCommand(m_uConfig.ConfigCommands, m_uConfig.LocalizedCommands, sFunction, sKey, eCmdDefault)
    sParam = GetConfigForCommand(m_uConfig.ConfigCommands, m_uConfig.LocalizedCommands, sFunction, sKey & "Param", sParamDefault)
    For lIdx = 0 To UBound(A) Step 3
        A(lIdx + 1) = GetConfigForCommand(m_uConfig.ConfigCommands, m_uConfig.LocalizedCommands, sFunction, sKey & A(lIdx + 0), A(lIdx + 2))
    Next
End Sub

Private Sub pvSetLastError(sError As String, Optional ByVal ErrNum As UcsFiscalErrorsEnum = -1)
    If ErrNum < 0 Then
        m_uPrintData.LastErrNo = IIf(LenB(sError) = 0, ucsFerNone, ucsFerGeneralError)
    Else
        m_uPrintData.LastErrNo = ErrNum
    End If
    m_uPrintData.LastError = sError
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
    Dim eIdx            As UcsIslConfigIndexEnum
    
    m_lTimeout = DEF_TIMEOUT
    m_lSequence = DEF_SEQUENCE
    LocalizedText(ucsFscLciInternalErrors) = STR_INTERNAL
    LocalizedText(ucsFscLciDipSwitchesTexts) = STR_DIP_SWITCHES
    LocalizedText(ucsFscLciReceiptTexts) = STR_RECEIPT_TEXTS
    LocalizedText(ucsFscLciExtendedErrors) = STR_EXT_ERRORS
    eIdx = ucsIslCfgDatecsModelA
    LocalizedText(ucsFscLciPrinterStatuses, eIdx) = STR_STATUSES_A
    LocalizedText(ucsFscLciPrinterErrors, eIdx) = STR_ERRORS_A
    LocalizedText(ucsFscLciMemoryTexts, eIdx) = STR_MEMORY_A
    eIdx = ucsIslCfgDatecsModelB
    LocalizedText(ucsFscLciPrinterStatuses, eIdx) = STR_STATUSES_B
    LocalizedText(ucsFscLciPrinterErrors, eIdx) = STR_ERRORS_B
    LocalizedText(ucsFscLciMemoryTexts, eIdx) = STR_MEMORY_B
    m_uConfig.CommEventTimeout = DEF_COMM_EVENT_TIMEOUT * 2
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

Private Function IDeviceProtocol_AddPayment(ByVal PmtType As UcsFiscalPaymentTypeEnum, Name As String, ByVal Amount As Double, Optional ByVal Rate As Double) As Boolean
    IDeviceProtocol_AddPayment = PpdAddPayment(m_uPrintData, PmtType, Name, Amount)
End Function

Private Function IDeviceProtocol_AddPLU(Name As String, ByVal Price As Double, ByVal Quantity As Double, ByVal TaxGroup As Long, UnitOfMeasure As String, ByVal DepartmentNo As Long) As Boolean
    IDeviceProtocol_AddPLU = PpdAddPLU(m_uPrintData, Name, Price, Quantity, TaxGroup, UnitOfMeasure, DepartmentNo)
End Function

Private Function IDeviceProtocol_AutodetectDevices(Ports As Variant) As Variant
    IDeviceProtocol_AutodetectDevices = AutodetectDevices(Ports)
End Function

Private Function IDeviceProtocol_CancelReceipt() As Boolean
    IDeviceProtocol_CancelReceipt = CancelReceipt()
End Function

Private Function IDeviceProtocol_PrintServiceDeposit(OperatorCode As String, OperatorPassword As String, ByVal Amount As Double) As Variant
    IDeviceProtocol_PrintServiceDeposit = PrintServiceDeposit(Amount)
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

Private Function IDeviceProtocol_EndReceipt(Optional ResumeToken As String) As Boolean
    If PpdEndReceipt(m_uPrintData, ResumeToken) Then
        IDeviceProtocol_EndReceipt = pvPrintReceipt(m_uPrintData.ExecCtx, m_uPrintData.Row, m_uPrintData.RowCount)
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
    IDeviceProtocol_GetCommandLog = m_sCommandLog
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

