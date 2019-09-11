VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDatecsProtocol"
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
' Implementation of Datecs X models protocol
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "cDatecsProtocol"
Implements IDeviceProtocol

'=========================================================================
' Public events
'=========================================================================

Event CommandComplete(ByVal lCmd As Long, sData As String, sResult As String)

'=========================================================================
' Public enums
'=========================================================================

Public Enum UcsDatecsStatusBitsEnum
    ucsDcsStbSyntaxError = 2 ^ 0
    ucsDcsStbInvalidCommand = 2 ^ 1
    ucsDcsStbClockNotSet = 2 ^ 2
    ucsDcsStbFailurePrintingMechanism = 2 ^ 4
    ucsDcsStbGeneralError = 2 ^ 5
    ucsDcsStbCoverOpen = 2 ^ 6
    ucsDcsStbSumOverflow = 2 ^ (8 + 0)
    ucsDcsStbCommandNotPermitted = 2 ^ (8 + 1)
    ucsDcsStbOutOfPaper = 2 ^ (16 + 0)
    ucsDcsStbNearOutOfPaper = 2 ^ (16 + 1)
    ucsDcsStbEJFull = 2 ^ (16 + 2)
    ucsDcsStbFiscalReceiptOpen = 2 ^ (16 + 3)
    ucsDcsStbNearEJFull = 2 ^ (16 + 4)
    ucsDcsStbNonFiscalReceiptOpen = 2 ^ (16 + 5)
End Enum

Public Enum UcsDatecsFiscalMemoryEnum
    ucsDcsFmeAccessError = 2 ^ 0
    ucsDcsFmeTaxNumberNotSet = 2 ^ 1
    ucsDcsFmeSerialNoSet = 2 ^ 2
    ucsDcsFmeNearOutOnMemory = 2 ^ 3
    ucsDcsFmeOutOfMemory = 2 ^ 4
    ucsDcsFmeMemoryError = 2 ^ 5
    ucsDcsFmeFormatted = 2 ^ (8 + 1)
    ucsDcsFmeDeviceFiscalized = 2 ^ (8 + 3)
    ucsDcsFmeVatRatesAreSetOnce = 2 ^ (8 + 4)
End Enum

Public Enum UcsDatecsCommandsEnum
    ucsDcsCmdNoop = -1
    '--- init
    ucsDcsCmdInitDateTime = 61
    ucsDcsCmdInitInvoiceNo = 66
    ucsDcsCmdInitBulstat = 98
    ucsDcsCmdInitOperatorPassword = 101
    ucsDcsCmdInitItem = 107
    ucsDcsCmdInitOpenDrawer = 106
    '--- transaction
    ucsDcsCmdNonFiscalOpen = 38
    ucsDcsCmdNonFiscalClose = 39
    ucsDcsCmdNonFiscalText = 42
    ucsDcsCmdFiscalReversal = 43
    ucsDcsCmdLineFeed = 44
    ucsDcsCmdFiscalOpen = 48
    ucsDcsCmdFiscalSell = 49
    ucsDcsCmdFiscalSubtotal = 51
    ucsDcsCmdFiscalPayment = 53
    ucsDcsCmdFiscalText = 54
    ucsDcsCmdFiscalClose = 56
    ucsDcsCmdFiscalCgInfo = 57
    ucsDcsCmdFiscalCancel = 60
    ucsDcsCmdFiscalServiceDeposit = 70
    '--- info
    ucsDcsCmdInfoDateTime = 62
    ucsDcsCmdInfoLastFiscalBlock = 64
    ucsDcsCmdInfoTotalsByTaxGroups = 65
    ucsDcsCmdInfoFreeMemoryRecords = 68
    ucsDcsCmdInfoStatus = 74
    ucsDcsCmdInfoTransaction = 76
    ucsDcsCmdInfoLastFiscalDateTime = 86
    ucsDcsCmdInfoDepartment = 88
    ucsDcsCmdInfoDiagnostics = 90
    ucsDcsCmdInfoBulstat = 99
    ucsDcsCmdInfoTotalsByPayments = 110
    ucsDcsCmdInfoOperator = 112
    ucsDcsCmdInfoMemoryDump = 116
    ucsDcsCmdInfoDevice = 123
    ucsDcsCmdInfoProgramming = 255                 '-- command above &H7F
    '--- print
    ucsDcsCmdPrintReportTaxRates = 50
    ucsDcsCmdPrintDailyReport = 69
    ucsDcsCmdPrintDiagnostics = 71
    ucsDcsCmdPrintBarcode = 84
    ucsDcsCmdPrintSeparator = 92
    ucsDcsCmdPrintReportByDateDetailed = 94
    ucsDcsCmdPrintReportByNumberShort = 95
    ucsDcsCmdPrintReportByOperators = 105
    ucsDcsCmdPrintDuplicate = 109
    ucsDcsCmdPrintReportByItems = 111
    '--- paper
    ucsDcsCmdPaperLinefeed = 44
    ucsDcsCmdPaperCut = 46
    '--- display
    ucsDcsCmdDisplayClear = 33
    ucsDcsCmdDisplaySecondLine = 35
    ucsDcsCmdDisplayFirstLine = 47
    ucsDcsCmdDisplayClock = 63
    '--- misc
    ucsDcsCmdMiscRingSound = 80
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

Private Const STR_INTERNAL              As String = "Timeout waiting for response|Invalid message format or checksum (NACK)|Invalid response length: %1|Response missing postamble: %1|Response missing preamble: %1|No receipt started| or invalid operator password|Error in line %1: %2|Cannot cancel previous receipt|Invalid report type|Error %1|Memory %1|Unsupported payment type %1"
Private Const STR_STATUSES              As String = "Syntax error|Command code is invalid|2|3|Failure in printing mechanism|5|Cover is open|7|Overflow during command execution|Command is not permitted|2|3|4|5|6|7|Out of paper|Near paper end|EJ is full|Fiscal receipt is open|EJ nearly full|Nonfiscal receipt is open|6|7"
Private Const STR_ERRORS_ONLY           As String = "Syntax error|Command code is invalid|2|3|Failure in printing mechanism|5|6|7|Overflow during command execution|Command is not permitted|2|3|4|5|6|7|Out of paper|1|2|3|4|5|6|7"
Private Const STR_MEMORY                As String = "Error when trying to access data stored in the FM|Tax number is set|Serial number and number of FM are set|There is space for less then 60 reports in FM|FM is full|FM error|FM is not found or damaged|7|0|FM is formatted|2|Device is fiscalized|VAT are set at least once|5|6|7"
Private Const STR_EXT_ERRORS            As String = "1|Overflow after operation|3|No more sales in this bill|4|No more payments in this bill|5|Attempt to perform zero transaction|6|Attempt to perform a sale, after payment is initiated|7|Operation not allowed|8|TAX group with disabled sales|11|More than one decimal point|12|More than one character '+' or '-'|13|Symbol '+' or '-' is not in the first position|14|Illegal symbol, e.g. barcode contains not only numbers|15|Extra numbers after decimal point|16|Extra characters|20|With current settings cannot process this command from PC|21|Value outside limits|22|Review system parameter 8|23|Attempt to ""deep"" void after discount/surcharge on subtotal|24|Attempt to ""deep"" void of non-existent transaction|25|Attempt to make a payment without having sales|26|Attempt to sell an item with quantity exceeding its stock|41|Barcode is incorrect (wrong checksum)|42|Attempt to sell without barcode|" & _
                                                    "43|Attempt to program with weight barcode|44|Attempt to sell without programmed barcode|45|Attempt to program already existing barcode|66|Incorrect password|71|!!! Incorrect data in fiscal memory !!!!|72|!!! Error saving in fiscal memory !!!!|76|Information needed from NRA servers|90|Needs a reset with monthly report|91|Needs a reset with daily z-report|92|Needs a reset with report by operators|93|Needs a reset with report by articles|94|Can not reprogram this field|81|Daily z-report overflow|83|Report by operators overflow|84|Report by articles overflow|84|Monthly report overflow|88|Electronic journal overflow|102|No communication between fiscal unit and tax terminal|104|Incorrect communication between fiscal unit and tax terminal|110|SIM card replaced|111|Error in communication between tax terminal and NRA servers|113|NRA server rejects data sent|" & _
                                                    "117|Failed to connect tax terminal with mobile operator network|118|The operation is denied|119|Incorrect value|120|Missing value"
Private Const STR_RECEIPT_TEXTS         As String = "CASH|CARD|CHEQUE|CREDIT|SURCHARGE %1|DISCOUNT %1|SUBTOTAL|Sales %1|EIC"
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
Private Const DEF_ITEM_CHARS            As Long = 72
Private Const DEF_MIN_DISCOUNT          As Double = -99
Private Const DEF_MAX_DISCOUNT          As Double = 99
Private Const DEF_MAX_RECEIPT_ROWS      As Long = 1000
Private Const DEF_MAX_PAYMENT_LEN       As Long = 16
Private Const DEF_MAX_CGINFO            As String = "36|36|36|36|36|16|16"
Private Const DEF_MAX_OPERATOR_NAME     As Long = 32
Private Const DEF_COMM_EVENT_TIMEOUT    As Long = 250
Private Const DEF_SEP                   As String = vbTab
Private Const ucsFscDscPluAbs           As Long = ucsFscDscPlu + 100
Private Const ucsFscDscSubtotalAbs      As Long = ucsFscDscSubtotal + 100
Private Const ucsFscRcpNonfiscal        As Long = ucsFscRcpSale + 100
Private Const FORMAT_DATE_EXTENDED      As String = "dd\-mm\-yy hh\:nn\:ss"
Private Const FORMAT_DATE_REPORT        As String = "dd\-mm\-yy"

Private m_sDevice                   As String
Private m_oConnector                As IDeviceConnector
Private m_lTimeout                  As Long
Private m_lSequence                 As Long
Private m_eStatus                   As UcsDatecsStatusBitsEnum
Private m_eMemory                   As UcsDatecsFiscalMemoryEnum
Private m_sCommandLog               As String
Private m_vInfoDiagnostics          As Variant
Private m_uPrintData                As UcsProtocolPrintData
Private m_uConfig                   As UcsConfigValues

Private Enum UcsInternalErrors
    ucsErrTimeoutWaitingForResponse
    [__ucsErrInvalidMessageFormatOrChecksum]
    ucsErrInvalidResponseLength
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
    ucsTxtPaymentCard
    ucsTxtPaymentBank
    ucsTxtPaymentCheque
    ucsTxtSurcharge
    ucsTxtDiscount
    ucsTxtSubTotal
    ucsTxtPluSales
    ucsTxtTaxNo
End Enum

Private Enum UcsProtocolMarkers
    SYM_PREAMBLE = &H1
    SYM_TERMINATOR = &H3
    SYM_SEPARATOR = &H4
    SYM_POSTAMBLE = &H5
    SYM_NAK = &H15
    SYM_SYN = &H16
End Enum

Private Type UcsAutodetect
    Connector           As cSerialPortConnector
    WriteOver           As OVERLAPPED
    ReadOver            As OVERLAPPED
    Buffer              As String
End Type

Private Type UcsConfigValues
    IsEcr               As VbTriState
    MaxPaymentLen       As Long
    MaxCgInfo(0 To 6)   As Long
    MaxOperatorName     As Long
    CommEventTimeout    As Long
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

Property Get Status(ByVal eValue As UcsDatecsStatusBitsEnum) As Boolean
    Status = (m_eStatus And eValue) <> 0
End Property

Property Get StatusText() As String
    StatusText = pvGetFlagsText(m_eStatus, m_uConfig.LocalizedText(ucsFscLciPrinterStatuses))
End Property

Property Get ErrorText() As String
    ErrorText = pvGetFlagsText(m_eStatus, m_uConfig.LocalizedText(ucsFscLciPrinterErrors))
End Property

Property Get Memory(ByVal eValue As UcsDatecsFiscalMemoryEnum) As Boolean
    Memory = (m_eMemory And eValue) <> 0
End Property

Property Get MemoryText() As String
    MemoryText = pvGetFlagsText(m_eMemory, m_uConfig.LocalizedText(ucsFscLciMemoryTexts))
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

Property Get LocalizedText(ByVal eIdx As UcsFiscalLocalizedIndexesEnum) As String
    If IsArray(m_uConfig.LocalizedText(eIdx)) Then
        LocalizedText = Join(m_uConfig.LocalizedText(eIdx), "|")
    End If
End Property

Property Let LocalizedText(ByVal eIdx As UcsFiscalLocalizedIndexesEnum, sValue As String)
    m_uConfig.LocalizedText(eIdx) = SplitOrReindex(sValue, "|")
    Select Case eIdx
    Case ucsFscLciInternalErrors
        m_uPrintData.LocalizedText.ErrNoReceiptStarted = pvInternal(ucsErrNoReceiptStarted)
    Case ucsFscLciReceiptTexts
        m_uPrintData.LocalizedText.TxtSurcharge = pvText(ucsTxtSurcharge)
        m_uPrintData.LocalizedText.TxtDiscount = pvText(ucsTxtDiscount)
        m_uPrintData.LocalizedText.TxtPluSales = pvText(ucsTxtPluSales)
    End Select
End Property

Property Get IsEcr() As Boolean
    If m_uConfig.IsEcr = vbUseDefault Then
        m_uConfig.IsEcr = IIf(Left$(GetDeviceModel(), 1) <> "F", vbTrue, vbFalse)
    End If
    IsEcr = (m_uConfig.IsEcr = vbTrue)
End Property

'= private ===============================================================

Private Property Get pvRowChars() As Long
    Const FUNC_NAME     As String = "pvRowChars"
    
    On Error GoTo EH
    If m_uPrintData.Config.RowChars = 0 Then
        m_uPrintData.Config.RowChars = C_Lng(pvSettings("PrintColumns"))
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
    pvCommentChars = pvRowChars - 2
End Property

Private Property Get pvItemChars() As Long
    pvItemChars = m_uPrintData.Config.ItemChars
    If pvItemChars = 0 Then
        pvItemChars = DEF_ITEM_CHARS
    End If
End Property

Private Property Get pvSettings(sName As String, Optional ByVal Index As Long = -1) As String
    Const CONF_NAME     As String = "Settings"
    Dim eCmd            As UcsDatecsCommandsEnum
    Dim sParam          As String
    Dim lIndex          As Long
    
    pvGetCommandConfig CONF_NAME, "InfoProgramming", eCmd, ucsDcsCmdInfoProgramming, sParam, "%1" & DEF_SEP & "%2" & DEF_SEP & DEF_SEP, _
        "Index", lIndex, 1
    sParam = Printf(sParam, sName, IIf(Index >= 0, Index, vbNullString))
    pvSettings = At(Split(SendCommand(eCmd, sParam), DEF_SEP), lIndex)
End Property

Private Property Let pvSettings(sName As String, Optional ByVal Index As Long = -1, sValue As String)
    Const CONF_NAME     As String = "Settings"
    Dim eCmd            As UcsDatecsCommandsEnum
    Dim sParam          As String
    Dim lIndex          As Long
    
    pvGetCommandConfig CONF_NAME, "InitProgramming", eCmd, ucsDcsCmdInfoProgramming, sParam, "%1" & DEF_SEP & "%2" & DEF_SEP & "%3" & DEF_SEP, _
        "Index", lIndex, 1
    sParam = Printf(sParam, sName, IIf(Index >= 0, Index, vbNullString), sValue)
    SendCommand eCmd, sParam
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
    Dim sFirmware       As String
    Dim sError          As String
    Dim vSpeeds         As Variant
    Dim baText()        As Byte
    
    On Error GoTo EH
    sProbeText = GetConfigValue(STR_PROTOCOL_DATECS_X, "ProbeText", Chr$(SYM_PREAMBLE) & Chr$(&H30) & Chr$(&H30) & Chr$(&H32) & Chr$(&H3B) & Chr$(&H20) & Chr$(&H30) & Chr$(&H30) & Chr$(&H35) & Chr$(&H3A) & Chr$(&H9) & Chr$(SYM_POSTAMBLE) & Chr$(&H30) & Chr$(&H31) & Chr$(&H3C) & Chr$(&H3A) & Chr$(SYM_TERMINATOR))
    sReplyText = GetConfigValue(STR_PROTOCOL_DATECS_X, "ReplyText", Chr$(SYM_SYN)) ' SYN (&H16)
    lTimeout = GetConfigNumber(STR_PROTOCOL_DATECS_X, "AutodetectTimeout", DEF_AUTODETECTTIMEOUT)
    lDeviceInfoTimeout = GetConfigNumber(STR_PROTOCOL_DATECS_X, "DeviceInfoTimeout", DEF_DEVICEINFOTIMEOUT)
    AssignVariant vSpeeds, GetConfigValue(STR_PROTOCOL_DATECS_X, "AutodetectSpeeds", Split(DEF_AUTODETECTSPEEDS, "|"))
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
                    End If
                    If LenB(.Buffer) <> 0 Then
'                        Debug.Print lIdx, ToHexDump(.Buffer), Timer
                        If Left$(.Buffer, Len(sReplyText)) = sReplyText Then
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
    Dim eCmd            As UcsDatecsCommandsEnum
    Dim sParam          As String
    Dim lSerialIndex    As Long
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
    If LenB(JsonItem(oOptions, "IP")) <> 0 Then
        m_sDevice = Trim$(JsonItem(oOptions, "IP")) & _
            ":" & Znl(C_Lng(JsonItem(oOptions, "Port")), DEF_SOCKET_PORT)
    Else
        m_sDevice = Zn(Trim$(JsonItem(oOptions, "Port")), DEF_SERIAL_PORT) & _
            "," & Znl(C_Lng(JsonItem(oOptions, "Speed")), DEF_SERIAL_SPEED) & _
            "," & JsonItem(oOptions, "Persistent") & _
            "," & Znl(C_Lng(JsonItem(oOptions, "BaudRate")), 8) & _
            "," & IIf(UCase$(Zn(JsonItem(oOptions, "Parity"), "N")) = "Y", "Y", "N") & _
            "," & Znl(C_Lng(JsonItem(oOptions, "StopBits")), 1)
    End If
    If Not Connect() Then
        GoTo QH
    End If
    pvGetCommandConfig CONF_NAME, "InfoDiagnostics", eCmd, ucsDcsCmdInfoDiagnostics, sParam, DEF_SEP, _
        "SerialIndex", lSerialIndex, 7
    m_vInfoDiagnostics = Split(pvSendCommand(m_oConnector, m_lTimeout, eCmd, sParam, sError), DEF_SEP)
    If LenB(sError) <> 0 Then
        pvSetLastError sError
        GoTo QH
    End If
    '--- load config values
    sSerial = At(m_vInfoDiagnostics, lSerialIndex)
    With m_uPrintData.Config
        .RowChars = GetConfigNumber(sSerial, "RowChars", C_Lng(JsonItem(oOptions, "RowChars")))
        .ItemChars = GetConfigNumber(sSerial, "ItemChars", C_Lng(JsonItem(oOptions, "ItemChars")))
        .AbsoluteDiscount = C_Bool(GetConfigValue(sSerial, "AbsoluteDiscount", True))
        .NegativePrices = C_Bool(GetConfigValue(sSerial, "NegativePrices", False))
        .MinDiscount = GetConfigNumber(sSerial, "MinDiscount", Zndbl(C_Dbl(JsonItem(oOptions, "MinDiscount")), DEF_MIN_DISCOUNT))
        .MaxDiscount = GetConfigNumber(sSerial, "MaxDiscount", Zndbl(C_Dbl(JsonItem(oOptions, "MaxDiscount")), DEF_MAX_DISCOUNT))
        .MaxReceiptRows = GetConfigNumber(sSerial, "MaxReceiptRows", Znl(C_Lng(JsonItem(oOptions, "MaxReceiptRows")), DEF_MAX_RECEIPT_ROWS))
    End With
    With m_uConfig
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

Public Function SendCommand(ByVal eCmd As UcsDatecsCommandsEnum, Optional Data As String) As String
    Const FUNC_NAME     As String = "SendCommand"
    Dim sError          As String
    Dim lPos            As Long
    
    On Error GoTo EH
    pvSetLastError vbNullString
    If eCmd = ucsDcsCmdNoop Then
        Exit Function
    End If
    '--- auto-sense encoding
    SendCommand = pvSendCommand(m_oConnector, m_lTimeout, eCmd, Data, sError)
    RaiseEvent CommandComplete(eCmd, Data, SendCommand)
    If LenB(sError) <> 0 Then
        pvSetLastError sError
        GoTo QH
    End If
    If Status(ucsDcsStbGeneralError) Then
        pvSetLastError Zn(ErrorText, Printf(pvInternal(ucsErrUnknownError), "0x" & Hex$(m_eStatus)))
        GoTo QH
    End If
    If Memory(ucsDcsFmeMemoryError) Then
        pvSetLastError Zn(MemoryText, Printf(pvInternal(ucsErrUnknownMemory), "0x" & Hex$(m_eMemory)))
        GoTo QH
    End If
    lPos = InStr(SendCommand, DEF_SEP)
    If lPos > 1 Then
        If Left$(SendCommand, lPos - 1) <> "0" Then
            pvSetLastError "Error " & Left$(SendCommand, lPos - 1)
        End If
    End If
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function CopyLastReceipt(Optional ReceiptNumber As String) As Boolean
    Const CONF_NAME     As String = "CopyLastReceipt"
    Dim eCmd            As UcsDatecsCommandsEnum
    Dim sParam          As String
'    Dim sFormat         As String
    
    On Error GoTo EH_Cancel
    If LenB(ReceiptNumber) = 0 Then
        pvGetCommandConfig CONF_NAME, "PrintDuplicate", eCmd, ucsDcsCmdPrintDuplicate, sParam, vbNullString
        SendCommand eCmd, sParam
    Else
'        sParam = "P"
'        pvGetCommandConfig CONF_NAME, "InitEcTape" & sParam, eCmd, ucsDcsCmdInitEcTape, sParam, sParam, _
'            "Format", sFormat, "%1,%2"
'        SendCommand eCmd, Printf(sFormat, sParam, ReceiptNumber)
    End If
    '--- success
    CopyLastReceipt = True
EH_Cancel:
End Function

Public Function GetClock() As Date
    Const FUNC_NAME     As String = "GetClock"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsDatecsCommandsEnum
    Dim sParam          As String
    Dim lIndex          As Long
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoDateTime", eCmd, ucsDcsCmdInfoDateTime, sParam, vbNullString, _
        "Index", lIndex, 1
    GetClock = C_Date(Replace(At(Split(SendCommand(eCmd, sParam), DEF_SEP), lIndex), "DST", vbNullString))
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function SetClock(ByVal NewDate As Date) As Boolean
    Const FUNC_NAME     As String = "SetClock"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsDatecsCommandsEnum
    Dim sParam          As String
    Dim sFormat         As String
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InitDateTime", eCmd, ucsDcsCmdInitDateTime, sParam, "%1" & DEF_SEP, _
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
    Dim eCmd            As UcsDatecsCommandsEnum
    Dim sParam          As String
    
    On Error GoTo EH
    If Not CancelReceipt() Then
        pvSetLastError pvInternal(ucsErrCannotCancelReceipt)
        GoTo QH
    End If
    Select Case ReportType
    Case ucsFscRstDaily
        pvGetCommandConfig CONF_NAME, "PrintDailyReport", eCmd, ucsDcsCmdPrintDailyReport, sParam, "X" & DEF_SEP
    Case ucsFscRstDailyByItems
        pvGetCommandConfig CONF_NAME, "PrintDailyReportItems", eCmd, ucsDcsCmdPrintDailyReport, sParam, "G" & DEF_SEP
    Case ucsFscRstDailyByDepartment
        pvGetCommandConfig CONF_NAME, "PrintDailyReportDepartments", eCmd, ucsDcsCmdPrintDailyReport, sParam, "D" & DEF_SEP
    Case Else
        pvSetLastError Printf(pvInternal(ucsErrInvalidReportType), ReportType)
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
    Dim eCmd            As UcsDatecsCommandsEnum
    Dim sParam          As String
    
    On Error GoTo EH
    If Not CancelReceipt() Then
        pvSetLastError pvInternal(ucsErrCannotCancelReceipt)
        GoTo QH
    End If
    Select Case ReportType
    Case ucsFscRstDaily
        pvGetCommandConfig CONF_NAME, "PrintDailyReport", eCmd, ucsDcsCmdPrintDailyReport, sParam, "Z" & DEF_SEP
    Case Else
        pvSetLastError Printf(pvInternal(ucsErrInvalidReportType), ReportType)
        GoTo QH
    End Select
    SendCommand eCmd, sParam
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
    Dim eCmd            As UcsDatecsCommandsEnum
    Dim sFormat         As String
    Dim sParam          As String
    
    On Error GoTo EH
    If Not CancelReceipt() Then
        pvSetLastError pvInternal(ucsErrCannotCancelReceipt)
        GoTo QH
    End If
    Select Case ReportType
    Case ucsFscRstPeriodShort
        '--- {Type}<SEP>{Start}<SEP>{End}<SEP>
        sParam = "0" & DEF_SEP & "%1" & DEF_SEP & "%2" & DEF_SEP
        pvGetCommandConfig CONF_NAME, "PrintReportByDateShort", eCmd, ucsDcsCmdPrintReportByDateDetailed, sParam, sParam, _
            "Format", sFormat, FORMAT_DATE_REPORT
    Case ucsFscRstPeriodDetailed
        sParam = "1" & DEF_SEP & "%1" & DEF_SEP & "%2" & DEF_SEP
        pvGetCommandConfig CONF_NAME, "PrintReportByDateDetailed", eCmd, ucsDcsCmdPrintReportByDateDetailed, sParam, sParam, _
            "Format", sFormat, FORMAT_DATE_REPORT
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
    Dim eCmd            As UcsDatecsCommandsEnum
    Dim sParam          As String
    Dim lIndex          As Long
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoDevice", eCmd, ucsDcsCmdInfoDevice, sParam, "3" & DEF_SEP, _
        "Index", lIndex, 1
    GetLastReceiptNo = At(Split(SendCommand(eCmd, sParam), DEF_SEP), lIndex)
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetLastInvoiceNo() As String
    Const FUNC_NAME     As String = "GetLastInvoiceNo"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsDatecsCommandsEnum
    Dim sParam          As String
    Dim lIndex          As Long
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InitInvoiceNo", eCmd, ucsDcsCmdInitInvoiceNo, sParam, vbNullString, _
        "Index", lIndex, 3
    GetLastInvoiceNo = At(Split(SendCommand(eCmd, sParam), DEF_SEP), lIndex)
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
    Dim vSplit          As Variant
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoStatus", eCmd, ucsDcsCmdInfoStatus, sParam, "0" & DEF_SEP, _
        "Index", lIndex, 5, _
        "DateIndex", lDateIndex, 6
    vSplit = Split(SendCommand(eCmd, sParam), DEF_SEP)
    If LenB(m_uPrintData.LastError) = 0 Then
        LastReceiptDate = C_Date(At(vSplit, lDateIndex))
        GetLastQRCodeInfo = RTrim$(At(vSplit, lIndex))
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
        "Index", lIndex, 7
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
        "Index", lIndex, 8
    GetFiscalMemoryNo = At(m_vInfoDiagnostics, lIndex)
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetTaxNo() As String
    Const FUNC_NAME     As String = "GetTaxNo"
    Const CONF_NAME     As String = FUNC_NAME
    Dim sParam          As String
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoSettings", 0, 0, sParam, "TAXnumber"
    GetTaxNo = RTrim$(pvSettings(sParam))
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetTaxCaption() As String
    Const FUNC_NAME     As String = "GetTaxCaption"
    Const CONF_NAME     As String = FUNC_NAME
    Dim sParam          As String
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoSettings", 0, 0, sParam, "TAXlabel"
    GetTaxCaption = RTrim$(pvSettings(sParam))
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetHeaderText(ByVal Index As Long) As String
    Const FUNC_NAME     As String = "GetHeaderText"
    Const CONF_NAME     As String = FUNC_NAME
    Dim sParam          As String
    
    On Error GoTo EH
    If Index >= 1 And Index <= 10 Then
        pvGetCommandConfig CONF_NAME, "InfoSettings", 0, 0, sParam, "Header"
        GetHeaderText = RTrim$(pvSettings(sParam, Index - 1))
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetFooterText(ByVal Index As Long) As String
    Const FUNC_NAME     As String = "GetFooterText"
    Const CONF_NAME     As String = FUNC_NAME
    Dim sParam          As String
    
    On Error GoTo EH
    If Index >= 1 And Index <= 10 Then
        pvGetCommandConfig CONF_NAME, "InfoSettings", 0, 0, sParam, "Footer"
        GetFooterText = RTrim$(pvSettings(sParam, Index - 1))
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetPaymentName(ByVal Index As Long) As String
    Const FUNC_NAME     As String = "GetPaymentName"
    Const CONF_NAME     As String = FUNC_NAME
    Dim sParam          As String
    Dim lIndex          As Long
    
    On Error GoTo EH
    If Index < 0 Then
        '--- custom payment types: 5, 6, 7 & 8
        Index = 4 - Index
    End If
    If Index > 3 Then
        lIndex = pvToDevicePmtIndex(Index)
        pvGetCommandConfig CONF_NAME, "InfoSettings" & Index, 0, 0, sParam, "PayName", _
            "Index", lIndex, lIndex
        If lIndex >= 0 Then
            GetPaymentName = RTrim$(pvStripBold(pvSettings(sParam, lIndex)))
        End If
'        Debug.Print Index, lIndex, sParam, GetPaymentName
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
    Dim eCmd            As UcsDatecsCommandsEnum
    Dim sParam          As String
    
    On Error GoTo EH_Cancel
    pvGetCommandConfig CONF_NAME, "InitOpenDrawer", eCmd, ucsDcsCmdInitOpenDrawer, sParam, vbNullString
    SendCommand eCmd, sParam
    '--- success
    OpenDrawer = True
EH_Cancel:
End Function

Public Function CancelReceipt() As Boolean
    Const FUNC_NAME     As String = "CancelReceipt"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eTransCmd       As UcsDatecsCommandsEnum
    Dim sTransParam     As String
    Dim lTransIndex     As Long
    Dim lStatusClosed   As Long
    Dim eCmd            As UcsDatecsCommandsEnum
    Dim sParam          As String
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoTransaction", eTransCmd, ucsDcsCmdInfoTransaction, sTransParam, vbNullString, _
        "Index", lTransIndex, 1, _
        "StatusClosed", lStatusClosed, 0
    If C_Lng(At(Split(SendCommand(eTransCmd, sTransParam), DEF_SEP), lTransIndex)) <> lStatusClosed Then
        If Status(ucsDcsStbFiscalReceiptOpen) Then
            pvGetCommandConfig CONF_NAME, "FiscalCancel", eCmd, ucsDcsCmdFiscalCancel, sParam, vbNullString
            SendCommand eCmd, sParam
            If Status(ucsDcsStbFiscalReceiptOpen) Then
                pvGetCommandConfig CONF_NAME, "FiscalPayment", eCmd, ucsDcsCmdFiscalPayment, sParam, DEF_SEP & DEF_SEP & DEF_SEP
                SendCommand eCmd, sParam
                '--- {Seller}<SEP>{Receiver}<SEP>{Buyer}<SEP>{Address1}<SEP>{Address2}<SEP>{TypeTAXN}<SEP>{TAXN}<SEP>{VATN}<SEP>
                pvGetCommandConfig CONF_NAME, "FiscalCgInfo", eCmd, ucsDcsCmdFiscalCgInfo, sParam, DEF_SEP & DEF_SEP & DEF_SEP & DEF_SEP & DEF_SEP & "0" & DEF_SEP & "0000000000" & DEF_SEP & DEF_SEP
                SendCommand eCmd, sParam
                pvGetCommandConfig CONF_NAME, "FiscalClose", eCmd, ucsDcsCmdFiscalClose, sParam, vbNullString
                SendCommand eCmd, sParam
            End If
        Else
            pvGetCommandConfig CONF_NAME, "NonFiscalClose", eCmd, ucsDcsCmdNonFiscalClose, sParam, vbNullString
            SendCommand eCmd, sParam
        End If
        '--- success (or failure)
        CancelReceipt = C_Lng(At(Split(SendCommand(eTransCmd, sTransParam), DEF_SEP), lTransIndex)) = lStatusClosed
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
    Dim eCmd            As UcsDatecsCommandsEnum
    Dim sParam          As String
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoStatus", eCmd, ucsDcsCmdInfoStatus, sParam, vbNullString
    pvSendCommand m_oConnector, m_lTimeout, eCmd, sParam, m_uPrintData.LastError
    If LenB(m_uPrintData.LastError) <> 0 Then
        '--- do nothing
    ElseIf Status(ucsDcsStbGeneralError) Then
        pvSetLastError Zn(ErrorText, Printf(pvInternal(ucsErrUnknownError), "0x" & Hex$(m_eStatus)))
    ElseIf Memory(ucsDcsFmeMemoryError) Then
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

Public Function GetTotalsByPayments() As Variant
    Const FUNC_NAME     As String = "GetTotalsByPayments"
    Const CONF_NAME     As String = FUNC_NAME
    Dim vPmtSum         As Variant
    Dim vRevSum         As Variant
    Dim vRetVal         As Variant
    Dim lIdx            As Long
    Dim lPmtIndex       As Long
    Dim eCmd            As UcsDatecsCommandsEnum
    Dim sParam          As String
    Dim lOffset         As Long
    Dim lRevOffset      As Long
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoTotalsByPayments0", eCmd, ucsDcsCmdInfoTotalsByPayments, sParam, "0" & DEF_SEP, _
        "Offset", lOffset, 1
    vPmtSum = Split(SendCommand(eCmd, sParam), DEF_SEP)
    pvGetCommandConfig CONF_NAME, "InfoTotalsByPayments1", eCmd, ucsDcsCmdInfoTotalsByPayments, sParam, "1" & DEF_SEP, _
        "RevOffset", lRevOffset, 1
    vRevSum = Split(SendCommand(eCmd, sParam), DEF_SEP)
    If UBound(vPmtSum) > 0 Then
        ReDim vRetVal(0 To 7) As Variant
        For lIdx = 0 To UBound(vRetVal)
            lPmtIndex = pvToDevicePmtIndex(lIdx + 1)
            If lPmtIndex >= 0 Then
                vRetVal(lIdx) = Array(GetPaymentName(lIdx + 1), ParseSum(At(vPmtSum, lPmtIndex + lOffset)), ParseSum(At(vRevSum, lPmtIndex + lRevOffset)))
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

Public Function GetTotalsByTaxGroups() As Variant
    Const FUNC_NAME     As String = "GetTotalsByTaxGroups"
    Const CONF_NAME     As String = FUNC_NAME
    Dim vTaxRates       As Variant
    Dim vPmtSum         As Variant
    Dim vRevSum         As Variant
    Dim vRetVal         As Variant
    Dim lIdx            As Long
    Dim eCmd            As UcsDatecsCommandsEnum
    Dim sParam          As String
    Dim lOffset         As Long
    Dim lRevOffset      As Long
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoTaxRates", 0, 0, sParam, "valVat", _
        "Offset", lOffset, 0
    ReDim vTaxRates(0 To 7) As String
    For lIdx = 0 To UBound(vTaxRates)
        vTaxRates(lIdx) = pvSettings(sParam, lIdx + lOffset)
    Next
    pvGetCommandConfig CONF_NAME, "InfoTotalsByTaxGroups0", eCmd, ucsDcsCmdInfoTotalsByTaxGroups, sParam, "0" & DEF_SEP, _
        "Offset", lOffset, 2
    vPmtSum = Split(SendCommand(eCmd, sParam), DEF_SEP)
    pvGetCommandConfig CONF_NAME, "InfoTotalsByTaxGroups2", eCmd, ucsDcsCmdInfoTotalsByTaxGroups, sParam, "2" & DEF_SEP, _
        "RevOffset", lRevOffset, 2
    vRevSum = Split(SendCommand(eCmd, sParam), DEF_SEP)
    If UBound(vPmtSum) > 0 And UBound(vRevSum) > 0 Then
        ReDim vRetVal(0 To 7) As Variant
        For lIdx = 0 To UBound(vRetVal)
            vRetVal(lIdx) = Array(ParseSum(At(vTaxRates, lIdx)), ParseSum(At(vPmtSum, lIdx + lOffset)), ParseSum(At(vRevSum, lIdx + lRevOffset)))
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
    Dim eCmd            As UcsDatecsCommandsEnum
    Dim sParam          As String
    Dim lIndex          As Long
    Dim lReceipts       As Long
    Dim dLastDate       As Date
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoDailyStats", 0, 0, sParam, "nFBon" ' "nFBonDailyCount"
    lReceipts = C_Lng(pvSettings(sParam))
    pvGetCommandConfig CONF_NAME, "InfoDevice", eCmd, ucsDcsCmdInfoDevice, sParam, "3" & DEF_SEP, _
        "Index", lIndex, 4
    dLastDate = C_Date(At(Split(SendCommand(eCmd, sParam), DEF_SEP), lIndex))
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
    Dim eCmd            As UcsDatecsCommandsEnum
    Dim sParam          As String
    Dim sFormat         As String

    On Error GoTo EH
    '--- %1={Type}<SEP>%2={Amount}<SEP>
    sParam = "%1" & DEF_SEP & "%2" & DEF_SEP
    pvGetCommandConfig CONF_NAME, "FiscalServiceDeposit" & -(Abs(Amount) > DBL_EPSILON), eCmd, ucsDcsCmdFiscalServiceDeposit, sParam, sParam, _
        "Format", sFormat, "0.00"
    sParam = Printf(sParam, IIf(Amount >= 0, "0", "1"), SafeFormat(Amount, sFormat))
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
            Optional ByVal cmd As UcsDatecsCommandsEnum, _
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
    Dim sParam          As String
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoSettings", 0, 0, sParam, "OperName"
    GetOperatorName = pvSettings(sParam, C_Lng(OperatorCode))
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetDefaultPassword(OperatorCode As String) As String
    If IsEcr Then
        GetDefaultPassword = OperatorCode
    Else
        GetDefaultPassword = DEF_PASSWORD
    End If
End Function

'= private ===============================================================

Private Function pvSendCommand( _
            oConnector As IDeviceConnector, _
            ByVal lTimeout As Long, _
            ByVal eCmd As UcsDatecsCommandsEnum, _
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
    sSend = STR_CHR1 & pvToWireLong(&H20 + 10 + Len(sSend)) & Chr$(m_lSequence) & pvToWireLong(eCmd) & sSend & Chr$(SYM_POSTAMBLE)
    '--- increment sequence
    m_lSequence = IIf(m_lSequence >= &H7F, &H20, m_lSequence + 1)
    '--- calc checksum
    baText = ToAscii(sSend)
    For lIdx = 1 To UBound(baText)
        lSum = lSum + baText(lIdx)
    Next
    '--- append checksum and terminator
    sSend = sSend & pvToWireLong(lSum) & Chr$(SYM_TERMINATOR)
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
    Do While Right$(sText, 1) <> Chr$(SYM_TERMINATOR)
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
        If InStr(sBuffer, Chr$(SYM_NAK)) > 0 Then
            DebugLog FUNC_NAME, "Will retry send on NAK received, sBuffer=0x" & ToHexDump(sBuffer)
            GoTo RetrySend
        End If
        '--- remove SYN (&H16)
        sText = sText & Replace(sBuffer, Chr$(SYM_SYN), vbNullString)
    Loop
RetryParse:
    '--- response complete w/ TERMINATOR (&H3)
    Debug.Assert Right$(sText, 1) = Chr$(SYM_TERMINATOR)
    If Left$(sText, 1) = STR_CHR1 Then
        lIdx = InStr(sText, Chr$(SYM_POSTAMBLE))
        If lIdx > 5 Then
            lSize = pvFromWireLong(Mid$(sText, 2, 4)) - &H20
            '--- check message length
            If lSize >= 10 And lSize < lIdx Then
                '--- check for status info
                lIdx = InStrRev(sText, Chr$(SYM_SEPARATOR), lIdx)
                If lIdx > 10 Then
                    pvSetStatus Mid$(sText, lIdx + 1, 8)
                    lSize = lIdx - 1
                Else
                    pvSetStatus vbNullString '--- clear
                End If
                pvSendCommand = pvTranscode(Mid$(sText, 11, lSize - 10), False)
                DebugLog FUNC_NAME, "RetVal=" & pvSendCommand & _
                    IIf(Status(ucsDcsStbGeneralError), ", ErrorText=" & Zn(ErrorText, Printf(pvInternal(ucsErrUnknownError), "0x" & Hex$(m_eStatus))), vbNullString) & _
                    IIf(Memory(ucsDcsFmeMemoryError), ", MemoryText=" & Zn(MemoryText, Printf(pvInternal(ucsErrUnknownMemory), "0x" & Hex$(m_eMemory))), vbNullString)
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

Private Function pvToWireLong(ByVal lValue As Long) As String
    pvToWireLong = Chr$((lValue \ &H1000 Mod &H10) + &H30) & Chr$((lValue \ &H100 Mod &H10) + &H30) & Chr$((lValue \ &H10 Mod &H10) + &H30) & Chr$((lValue \ &H1 Mod &H10) + &H30)
End Function

Private Function pvFromWireLong(baValue() As Byte) As Long
    pvFromWireLong = (baValue(0) - &H30) * &H1000 + (baValue(2) - &H30) * &H100 + (baValue(4) - &H30) * &H10 + (baValue(6) - &H30)
End Function

Private Function pvTranscode(sText As String, ByVal bToPrinter As Boolean) As String
    #If bToPrinter Then '--- touch arg
    #End If
    pvTranscode = sText
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

Private Function pvPrintLine(vText As Variant) As Boolean
    Const FUNC_NAME     As String = "pvPrintLine"
    Const CONF_NAME     As String = "PrintLine"
    Dim vElem               As Variant
    Dim eCmd                As UcsDatecsCommandsEnum
    Dim sParam              As String
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "FiscalText", eCmd, ucsDcsCmdFiscalText, sParam, "%1" & DEF_SEP & DEF_SEP & DEF_SEP & DEF_SEP & DEF_SEP & DEF_SEP
    If IsArray(vText) Then
        For Each vElem In vText
            SendCommand eCmd, Printf(sParam, C_Str(vElem))
            If LenB(m_uPrintData.LastError) <> 0 Then
                GoTo QH
            End If
        Next
    Else
        SendCommand eCmd, Printf(sParam, C_Str(vText))
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
    Dim bInLogin        As Boolean
    Dim sData           As String
    Dim dblTotal        As Double
    Dim dblDisc         As Double
    Dim eCmd            As UcsDatecsCommandsEnum
    Dim sParam          As String
    Dim lIndex          As Long
    Dim lStatusClosed   As Long
    Dim sFormatPrice    As String
    Dim sFormatQty      As String
    Dim sFormatDisc     As String
    Dim sFormatAmount   As String
    Dim sFormatZero     As String
    Dim sFormatDate     As String
    Dim lAmountIndex    As Long
    Dim lPaidIndex      As Long
    Dim sDiscType       As String
    
    On Error GoTo EH
    '--- ring sale
    Do While uCtx.Row < lRowCount
        With uRow(uCtx.Row)
        Select Case .RowType
        Case ucsRowInit
            pvGetCommandConfig CONF_NAME, "InfoOperator", 0, 0, sParam, "OperName"
            sData = pvSettings(sParam, C_Lng(.InitOperatorCode) - 1)
            If LenB(.InitOperatorName) <> 0 Then
                If .InitOperatorName <> sData Then
                    pvSettings(sParam, C_Lng(.InitOperatorCode) - 1) = Left$(.InitOperatorName, m_uConfig.MaxOperatorName)
                End If
            Else
                .InitOperatorName = RTrim$(sData)
            End If
            If LenB(.InitInvData(ucsInvDocNo)) <> 0 Then
                Select Case .InitReceiptType
                Case ucsFscRcpInvoice, ucsFscRcpCreditNote
                    pvGetCommandConfig CONF_NAME, "InitInvoiceNo", eCmd, ucsDcsCmdInitInvoiceNo, sParam, "%1" & DEF_SEP & "%1" & DEF_SEP
                    sParam = Printf(sParam, .InitInvData(ucsInvDocNo))
                    SendCommand eCmd, sParam
                End Select
            End If
            Select Case .InitReceiptType
            Case ucsFscRcpReversal, ucsFscRcpCreditNote
                '--- %1={OpCode}<SEP>%2={OpPwd}<SEP>{TillNmb}<SEP>%3={Storno}<SEP>%4={DocNum}<SEP>%5={DateTime}<SEP>%6={FMNumber}
                '---    <SEP>{Invoice}<SEP>%7={ToInvoice}<SEP>%8={Reason}<SEP>%9={NSale}<SEP>
                sParam = "%1" & DEF_SEP & "%2" & DEF_SEP & "1" & DEF_SEP & "%3" & DEF_SEP & "%4" & DEF_SEP & "%5" & DEF_SEP & "%6" & _
                    IIf(.InitReceiptType = ucsFscRcpCreditNote, DEF_SEP & "I" & DEF_SEP & "%7" & DEF_SEP & "%8", DEF_SEP & DEF_SEP & DEF_SEP) & _
                    DEF_SEP & "%9" & DEF_SEP
                pvGetCommandConfig CONF_NAME, "FiscalReversal" & .InitReceiptType, eCmd, ucsDcsCmdFiscalReversal, sParam, sParam, _
                    "FormatDate", sFormatDate, FORMAT_DATE_EXTENDED
                sParam = Printf(sParam, .InitOperatorCode, .InitOperatorPassword, .InitRevData(ucsRevType), _
                    .InitRevData(ucsRevReceiptNo), Format$(.InitRevData(ucsRevReceiptDate), sFormatDate), .InitRevData(ucsRevFiscalMemoryNo), _
                    .InitRevData(ucsRevInvoiceNo), .InitRevData(ucsRevReason), .InitUniqueSaleNo)
            Case Else
                '--- %1={OpCode}<SEP>%2={OpPwd}%3=(<SEP>{NSale})<SEP>{TillNmb}<SEP>{Invoice}<SEP>
                sParam = "%1" & DEF_SEP & "%2%3" & DEF_SEP & "1" & DEF_SEP & IIf(.InitReceiptType = ucsFscRcpInvoice, "I", vbNullString) & DEF_SEP
                pvGetCommandConfig CONF_NAME, "FiscalOpen" & .InitReceiptType, eCmd, ucsDcsCmdFiscalOpen, sParam, sParam
                sParam = Printf(sParam, .InitOperatorCode, .InitOperatorPassword, _
                    IIf(LenB(.InitUniqueSaleNo) <> 0, DEF_SEP & .InitUniqueSaleNo, vbNullString))
            End Select
            bInLogin = True
            SendCommand eCmd, sParam
            If LenB(m_uPrintData.LastError) <> 0 Then
                GoTo QH
            End If
            bInLogin = False
        Case ucsRowPlu
            uCtx.PluCount = uCtx.PluCount + 1
            dblTotal = Round(.PluPrice * .PluQuantity, 2)
            If .DiscType = ucsFscDscPlu Then
                dblDisc = Round(dblTotal * .DiscValue / 100#, 2)
            ElseIf .DiscType = ucsFscDscPluAbs Then
                dblDisc = Round(.DiscValue, 2)
            Else
                dblDisc = 0
            End If
            uCtx.GrpTotal(.PluTaxGroup) = Round(uCtx.GrpTotal(.PluTaxGroup) + dblTotal + dblDisc, 2)
            vSplit = WrapText(.PluName, pvItemChars)
            For lIdx = 0 To UBound(vSplit) - 2
                pvPrintLine C_Str(vSplit(lIdx))
            Next
            sData = At(vSplit, UBound(vSplit))
            If UBound(vSplit) > 0 Then
                sData = At(vSplit, UBound(vSplit) - 1) & vbLf & sData
            End If
            Select Case .DiscType
            Case ucsFscDscPlu
                sDiscType = Switch(.DiscValue > 0, "1", .DiscValue < 0, "2", True, vbNullString)
            Case ucsFscDscPluAbs
                sDiscType = Switch(.DiscValue > 0, "3", .DiscValue < 0, "4", True, vbNullString)
            End Select
            '--- %1={PluName}<SEP>%2={TaxCd}<SEP>%3={Price}<SEP>%4={Quantity}<SEP>%5={DiscountType}<SEP>%6={DiscountValue}<SEP>{Department}<SEP>[ {Unit}<SEP> ]
            sParam = "%1" & DEF_SEP & "%2" & DEF_SEP & "%3" & DEF_SEP & "%4" & DEF_SEP & "%5" & DEF_SEP & "%6" & DEF_SEP & "0" & DEF_SEP
            pvGetCommandConfig CONF_NAME, "FiscalSell" & .PrintRowType, eCmd, ucsDcsCmdFiscalSell, sParam, sParam, _
                "FormatPrice", sFormatPrice, "0.00", _
                "FormatQty", sFormatQty, "0.000", _
                "FormatDisc", sFormatDisc, "0.00"
            sParam = Printf(sParam, sData, .PluTaxGroup, SafeFormat(.PluPrice, sFormatPrice), _
                IIf(.PluQuantity <> 1, SafeFormat(.PluQuantity, sFormatQty), vbNullString), _
                sDiscType, _
                IIf(LenB(sDiscType) <> 0, SafeFormat(.DiscValue, sFormatDisc), vbNullString))
            SendCommand eCmd, sParam
        Case ucsRowDiscount
            Select Case .PrintRowType
            Case ucsFscRcpNonfiscal
                If .DiscType = ucsFscDscSubtotal Then
                    pvGetCommandConfig CONF_NAME, "NonFiscalSubtotal" & .DiscType, 0, 0, vbNullString, vbNullString, _
                        "FormatDisc", sFormatDisc, "0.00", _
                        "FormatAmount", sFormatAmount, "0.00"
                    pvPrintLine AlignText(pvText(ucsTxtSubTotal), SafeFormat(SumArray(uCtx.GrpTotal), sFormatAmount), pvCommentChars)
                    For lIdx = 1 To UBound(uCtx.GrpTotal)
                        If Abs(uCtx.GrpTotal(lIdx)) > DBL_EPSILON Then
                            dblTotal = Round(uCtx.GrpTotal(lIdx) * .DiscValue / 100#, 2)
                            pvPrintLine AlignText(Printf(IIf(dblTotal > DBL_EPSILON, pvText(ucsTxtSurcharge), pvText(ucsTxtDiscount)), "*" & Chr$(191 + lIdx) & "  " & SafeFormat(Abs(.DiscValue), sFormatDisc) & "%"), _
                                SafeFormat(dblTotal, sFormatAmount), pvCommentChars)
                        End If
                    Next
                End If
            Case Else
                Select Case .DiscType
                Case ucsFscDscSubtotal
                    sDiscType = Switch(.DiscValue > 0, "1", .DiscValue < 0, "2", True, vbNullString)
                Case ucsFscDscSubtotalAbs
                    sDiscType = Switch(.DiscValue > 0, "3", .DiscValue < 0, "4", True, vbNullString)
                End Select
                '--- %1={Print}<SEP>{Display}<SEP>%2={DiscountType}<SEP>%3={DiscountValue}<SEP>
                sParam = "%1" & DEF_SEP & DEF_SEP & "%2" & DEF_SEP & "%3" & DEF_SEP
                pvGetCommandConfig CONF_NAME, "FiscalSubtotal" & .DiscType, eCmd, ucsDcsCmdFiscalSubtotal, sParam, sParam, _
                    "FormatDisc", sFormatDisc, "0.00"
                sParam = Printf(sParam, IIf(LenB(sDiscType) = 0, "1", vbNullString), _
                    sDiscType, _
                    IIf(LenB(sDiscType) <> 0, SafeFormat(.DiscValue, sFormatDisc), vbNullString))
                SendCommand eCmd, sParam
            End Select
            If .DiscType = ucsFscDscSubtotal Then
                For lIdx = 1 To UBound(uCtx.GrpTotal)
                    If Abs(uCtx.GrpTotal(lIdx)) > DBL_EPSILON Then
                        dblTotal = Round(uCtx.GrpTotal(lIdx) * .DiscValue / 100#, 2)
                        uCtx.GrpTotal(lIdx) = Round(uCtx.GrpTotal(lIdx) + dblTotal, 2)
                    End If
                Next
            End If
        Case ucsRowLine
            If .LineWordWrap Then
                vSplit = WrapText(.LineText, pvCommentChars)
            Else
                vSplit = Array(Left$(.LineText, pvCommentChars))
            End If
            pvPrintLine vSplit
        Case ucsRowPayment
            If .PmtType <> 0 Then
                '--- %1={PaidMode}<SEP>%2={Amount}<SEP>{Type}<SEP>
                sParam = "%1" & DEF_SEP & "%2" & DEF_SEP & DEF_SEP
                pvGetCommandConfig CONF_NAME, "FiscalPayment" & .PmtType, eCmd, ucsDcsCmdFiscalPayment, sParam, sParam, _
                    "FormatAmount", sFormatAmount, "0.00", _
                    "FormatZero", sFormatZero, "\"
                lIndex = pvToDevicePmtIndex(.PmtType)
                If lIndex < 0 Then
                    pvSetLastError Printf(pvInternal(ucsErrUnsupportedPaymentType), .PmtType)
                End If
                sParam = Printf(sParam, pvToDevicePmtIndex(.PmtType), SafeFormat(.PmtAmount, IIf(Abs(.PmtAmount) < DBL_EPSILON, sFormatZero, sFormatAmount)))
                SendCommand eCmd, sParam
                uCtx.PmtPrinted = True
            Else
                '--- check if anything left for payment
                pvGetCommandConfig CONF_NAME, "InfoTransaction", eCmd, ucsDcsCmdInfoTransaction, sParam, vbNullString, _
                    "Index", lIndex, 1, _
                    "StatusClosed", lStatusClosed, 0, _
                    "AmountIndex", lAmountIndex, 4, _
                    "PaidIndex", lPaidIndex, 5
                vSplit = Split(SendCommand(eCmd, sParam), DEF_SEP)
                If C_Lng(At(vSplit, lIndex)) <> lStatusClosed Then
                    If C_Dbl(At(vSplit, lAmountIndex)) > C_Dbl(At(vSplit, lPaidIndex)) _
                            Or LenB(At(vSplit, lAmountIndex)) = 0 _
                            Or Not uCtx.PmtPrinted Then
                        sParam = "%1" & DEF_SEP & DEF_SEP & DEF_SEP
                        pvGetCommandConfig CONF_NAME, "FiscalPayment" & .PmtType, eCmd, ucsDcsCmdFiscalPayment, sParam, sParam
                        sParam = Printf(sParam, pvToDevicePmtIndex(.PmtType))
                        SendCommand eCmd, sParam
                        If Not Status(ucsDcsStbGeneralError) Then
                            uCtx.PmtPrinted = True
                        End If
                    End If
                    Select Case .PrintRowType
                    Case ucsFscRcpInvoice, ucsFscRcpCreditNote
                        '--- %1={Seller}<SEP>%2={Receiver}<SEP>%3={Buyer}<SEP>%4={Address1}<SEP>%5={Address2}<SEP>%6={TypeTAXN}<SEP>%7={TAXN}<SEP>%8={VATN}<SEP>
                        sParam = "%1" & DEF_SEP & "%2" & DEF_SEP & "%3" & DEF_SEP & "%4" & DEF_SEP & "%5" & DEF_SEP & "%6" & DEF_SEP & "%7" & DEF_SEP & "%8" & DEF_SEP
                        pvGetCommandConfig CONF_NAME, "FiscalCgInfo", eCmd, ucsDcsCmdFiscalCgInfo, sParam, sParam
                        With uRow(0)
                            sParam = Printf(sParam, Left$(.InitOperatorName, m_uConfig.MaxCgInfo(0)), _
                                Left$(.InitInvData(ucsInvCgPrsReceive), m_uConfig.MaxCgInfo(1)), _
                                Left$(.InitInvData(ucsInvCgName), m_uConfig.MaxCgInfo(2)), _
                                Left$(.InitInvData(ucsInvCgCity), m_uConfig.MaxCgInfo(3)), _
                                Left$(.InitInvData(ucsInvCgAddress), m_uConfig.MaxCgInfo(4)), _
                                .InitInvData(ucsInvCgTaxNoType), _
                                Left$(.InitInvData(ucsInvCgTaxNo), m_uConfig.MaxCgInfo(5)), _
                                Left$(.InitInvData(ucsInvCgVatNo), m_uConfig.MaxCgInfo(6)))
                        End With
                        SendCommand eCmd, sParam
                    End Select
                    pvGetCommandConfig CONF_NAME, "FiscalClose" & .PrintRowType, eCmd, ucsDcsCmdFiscalClose, sParam, vbNullString
                    SendCommand eCmd, sParam
                End If
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
    Dim eCmd            As UcsDatecsCommandsEnum
    Dim sParam          As String
    Dim lNameIndex      As Long
    Dim lSerialIndex    As Long
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoDiagnostics", eCmd, ucsDcsCmdInfoDiagnostics, sParam, DEF_SEP, _
        "NameIndex", lNameIndex, 1, _
        "SerialIndex", lSerialIndex, 7
    If Not IsArray(vResult) Then
        vResult = Split(pvSendCommand(oConnector, lTimeout, eCmd, sParam, Error), DEF_SEP)
    End If
    If UBound(vResult) >= lSerialIndex And LenB(Error) = 0 Then
        '--- check serial
        sProtocol = STR_PROTOCOL_DATECS_X
        sModel = Trim$(At(vResult, lNameIndex))
        sFirmware = Trim$(At(vResult, lNameIndex + 1) & " " & At(vResult, lNameIndex + 2) & " " & At(vResult, lNameIndex + 3))
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

Private Function pvToDevicePmtIndex(ByVal PmtType As UcsFiscalPaymentTypeEnum) As Long
    Select Case PmtType
    Case ucsFscPmtCash
        pvToDevicePmtIndex = 0   ' SCash (б апни)
    Case ucsFscPmtCard
        pvToDevicePmtIndex = 2   ' SCards (деа.йюпрю)
    Case ucsFscPmtBank
        pvToDevicePmtIndex = 1   ' ??? (йпедхр)
    Case 4
        pvToDevicePmtIndex = 3   ' SChecks (вей)
    Case ucsFscPmtCustom1, 5
        pvToDevicePmtIndex = 4   ' ST (йсонм)
    Case ucsFscPmtCustom2, 6
        pvToDevicePmtIndex = 5   ' SOT (бюсвеп)
    Case Else
        pvToDevicePmtIndex = -1  ' unsupported
    End Select
End Function

Private Sub pvGetCommandConfig( _
            sFunc As String, _
            sKey As String, _
            eCmd As UcsDatecsCommandsEnum, _
            ByVal eCmdDefault As UcsDatecsCommandsEnum, _
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
        m_uPrintData.LastErrNo = IIf(LenB(sError) = 0, ucsFerNone, ucsFerGeneralError)
    Else
        m_uPrintData.LastErrNo = ErrNum
    End If
    m_uPrintData.LastError = sError
End Sub

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Initialize()
    m_lTimeout = DEF_TIMEOUT
    m_lSequence = &H20
    LocalizedText(ucsFscLciInternalErrors) = STR_INTERNAL
    LocalizedText(ucsFscLciPrinterStatuses) = STR_STATUSES
    LocalizedText(ucsFscLciPrinterErrors) = STR_ERRORS_ONLY
    LocalizedText(ucsFscLciMemoryTexts) = STR_MEMORY
    LocalizedText(ucsFscLciReceiptTexts) = STR_RECEIPT_TEXTS
    LocalizedText(ucsFscLciExtendedErrors) = STR_EXT_ERRORS
    m_uConfig.CommEventTimeout = DEF_COMM_EVENT_TIMEOUT * 2
    Set m_uConfig.ConfigCommands = GetConfigCollection(STR_PROTOCOL_DATECS_X, "Commands")
End Sub

'=========================================================================
' IDeviceProtocol interface
'=========================================================================

Private Function IDeviceProtocol_AddDiscount(ByVal DiscType As UcsFiscalDiscountTypeEnum, ByVal Value As Double) As Boolean
    IDeviceProtocol_AddDiscount = PpdAddDiscount(m_uPrintData, DiscType, Value)
End Function

Private Function IDeviceProtocol_AddLine(Line As String, Optional Command As String, Optional ByVal WordWrap As Boolean = True) As Boolean
    IDeviceProtocol_AddLine = PpdAddLine(m_uPrintData, Line, WordWrap)
End Function

Private Function IDeviceProtocol_AddPayment(ByVal Number As UcsFiscalPaymentTypeEnum, Name As String, ByVal Amount As Double, Optional ByVal Rate As Double) As Boolean
    IDeviceProtocol_AddPayment = PpdAddPayment(m_uPrintData, Number, Name, Amount)
End Function

Private Function IDeviceProtocol_AddPLU(Name As String, ByVal Price As Double, ByVal Quantity As Double, ByVal TaxGroup As Long) As Boolean
    IDeviceProtocol_AddPLU = PpdAddPLU(m_uPrintData, Name, Price, Quantity, TaxGroup)
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

Private Function IDeviceProtocol_EndReceipt(Optional ResumeTicket As String) As Boolean
    m_uPrintData.Config.RowChars = pvRowChars
    m_uPrintData.Config.ItemChars = pvItemChars
    If PpdEndReceipt(m_uPrintData, ResumeTicket) Then
        IDeviceProtocol_EndReceipt = pvPrintReceipt(m_uPrintData.ExecCtx, m_uPrintData.Row, m_uPrintData.RowCount)
    End If
End Function

Private Function IDeviceProtocol_GetCharsPerLine() As Long
    IDeviceProtocol_GetCharsPerLine = pvCommentChars
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

Private Function IDeviceProtocol_GetResumeTicket() As String
    IDeviceProtocol_GetResumeTicket = PpdGetResumeTicket(m_uPrintData)
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
    LocalizedText(Index) = Text
End Sub

Private Function IDeviceProtocol_StartReceipt(ByVal ReceiptType As UcsFiscalReceiptTypeEnum, OperatorCode As String, OperatorName As String, OperatorPassword As String, Optional UniqueSaleNo As String, Optional TableNo As String, Optional InvDocNo As String, Optional InvCgTaxNo As String, Optional ByVal InvCgTaxNoType As UcsFiscalTaxNoTypeEnum, Optional InvCgVatNo As String, Optional InvCgName As String, Optional InvCgCity As String, Optional InvCgAddress As String, Optional InvCgPrsReceive As String, Optional OwnData As String, Optional ByVal RevType As UcsFiscalReversalTypeEnum, Optional RevReceiptNo As String, Optional ByVal RevReceiptDate As Date, Optional RevFiscalMemoryNo As String, Optional RevInvoiceNo As String, Optional RevReason As String) As Boolean
    Dim sOperCode       As String
    Dim sOperPass       As String
    
    sOperCode = Zn(OperatorCode, DEF_OPERNO)
    sOperPass = Zn(OperatorPassword, GetDefaultPassword(sOperCode))
    IDeviceProtocol_StartReceipt = PpdStartReceipt(m_uPrintData, ReceiptType, sOperCode, OperatorName, sOperPass, UniqueSaleNo, InvDocNo, InvCgTaxNo, InvCgTaxNoType, InvCgVatNo, InvCgName, InvCgCity, InvCgAddress, InvCgPrsReceive, RevType, RevReceiptNo, RevReceiptDate, RevFiscalMemoryNo, RevInvoiceNo, RevReason)
End Function
