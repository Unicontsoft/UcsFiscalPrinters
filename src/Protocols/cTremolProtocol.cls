VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTremolProtocol"
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
' Implementation of Tremol protocol (previously Zeka)
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "cTremolProtocol"
Implements IDeviceProtocol

'=========================================================================
' Public events
'=========================================================================

Event CommandComplete(ByVal lCmd As Long, sData As String, sResult As String)

'=========================================================================
' Public enums
'=========================================================================

Public Enum UcsTremolCommandsEnum
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
    ucsZekCmdFiscalDepartmentSell = &H34
    ucsZekCmdFiscalPayment = &H35
    ucsZekCmdFiscalPayAndClose = &H36
    ucsZekCmdPrintText = &H37
    ucsZekCmdFiscalClose = &H38
    ucsZekCmdFiscalCancel = &H39
    ucsZekCmdPrintDuplicate = &H3A
    ucsZekCmdFiscalServiceDeposit = &H3B
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
    ucsZekCmdInfoTotalsByTaxGroups = &H6D
    ucsZekCmdInfoRegisters = &H6E
    ucsZekCmdInfoInvoiceNo = &H70
    ucsZekCmdInfoLastReceipt = &H71
    ucsZekCmdInfoTransaction = &H72
    ucsZekCmdInfoReportLastDate = &H73
    '--- print
    ucsZekCmdPrintDiagnostics = &H22
    ucsZekCmdPrintBarcode = &H51
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

'--- error codes
Private Const ERROR_IO_PENDING              As Long = 997

Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function CreateEvent Lib "kernel32" Alias "CreateEventA" (ByVal lpEventAttributes As Long, ByVal bManualReset As Long, ByVal bInitialState As Long, ByVal lpName As Long) As Long
Private Declare Function WaitForMultipleObjects Lib "kernel32" (ByVal nCount As Long, lpHandles As Long, ByVal bWaitAll As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
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

Private Const STR_INTERNAL              As String = "Timeout waiting for response|Invalid message format or checksum (NACK)|Invalid response length: %1|Invalid response checksum: %1|Response missing preamble: %1|No receipt started|Error in line %1: %2|Invalid report type|Device not online|Device not ready| or cannot set operator password|Unsupported payment type %1|Barcode type %1 not supported|Cannot print"
Private Const STR_STATUSES              As String = "Unknown error|Invalid command|Illegal command|Command denied because of uncommited report|Syntax error|Input register overflow|Input register is zero|Missing transaction for void|Insufficient subtotal|Tax groups conflict"
Private Const STR_ERRORS                As String = "Unknown error|Out of paper|Daily registers overflow|Clock not adjusted|Fiscal receipt open|Receipt not fully paid|Non-fiscal receipt open|Receipt already paid out|Read-only fiscal memory|Bad password or command not allowed|Display missing|24 hours without daily report|Printer overheated|Power down|Electronic journal overflow|Not enough conditions met"
Private Const STR_RECEIPT_TEXTS         As String = "Cash|Card|Bank|Surcharge%1|Discount%1|Subtotal|Sales %1|EIC;Citizen #;Foreigner #;Official #"
Private Const DEF_SEP                   As String = ";"
Private Const DEF_AUTODETECTTIMEOUT     As Long = 50
Private Const DEF_DEVICEINFOTIMEOUT     As Long = 1000
Private Const DEF_AUTODETECTSPEEDS      As String = "115200|38400|9600|19200|57600"
Private Const DEF_TIMEOUT               As Long = 3000
Private Const DEF_PRINT_TIMEOUT         As Long = 30000
Private Const DEF_OPERNO                As String = "1"
Private Const DEF_ECR_PASSWORD          As String = "0"
Private Const DEF_FP_PASSWORD           As String = "0000"
Private Const DEF_SOCKET_PORT           As Long = 8000
Private Const DEF_SERIAL_PORT           As String = "COM1"
Private Const DEF_SERIAL_SPEED          As Long = 115200
Private Const DEF_MIN_DISCOUNT          As Double = -100
Private Const DEF_MAX_DISCOUNT          As Double = 100
Private Const DEF_MAX_RECEIPT_ROWS      As Long = 1000
Private Const DEF_CODEPAGE              As Long = 1251
Private Const DEF_PING_TIMEOUT          As Long = 500
Private Const DEF_ROW_CHARS             As Long = 32
Private Const DEF_MAX_OPERCODE          As Long = 2
Private Const DEF_MAX_OPERNAME          As Long = 20
Private Const DEF_MAX_OPERPASSWORD      As Long = 4
Private Const DEF_MAX_COMPANY           As Long = 26
Private Const DEF_MAX_PRSRECEIVE        As Long = 16
Private Const DEF_MAX_VATNO             As Long = 13
Private Const DEF_MAX_TAXNO             As Long = 13
Private Const DEF_MAX_CITY              As Long = 30
Private Const DEF_MAX_INVOICENO         As Long = 10
Private Const DEF_MAX_PLUNAME           As Long = 36
Private Const DEF_MAX_COMMENT           As Long = 34
Private Const DEF_MAX_FISCALMEMORYNO    As Long = 8
Private Const DEF_MAX_UNIQUESALENO      As Long = 24
Private Const DEF_MAX_DEPARTMENTCODE    As Long = 2
Private Const DEF_MAX_RECEIPTNO         As Long = 6
Private Const DEF_SEQUENCE              As Long = &H20
Private Const LNG_MAX_SEQUENCE          As Long = &H9F
Private Const LNG_STATUS_READY          As Long = &H40
Private Const LNG_STATUS_IS_BUSY        As Long = &H41
Private Const LNG_STATUS_OUT_OF_PAPER   As Long = &H42
Private Const LNG_STATUS_OUT_OF_PAPER_LEGACY As Long = &H7
Private Const LNG_STATUS_OUT_OF_PAPER_AND_IS_BUSY As Long = &H43
Private Const LNG_STATUS_OVERHEATED     As Long = &H44
Private Const ucsFscRcpNonfiscal        As Long = ucsFscRcpSale + 100
Private Const FORMAT_DATE_EXTENDED      As String = "dd\-mm\-yy hh\:nn\:ss"
Private Const FORMAT_DATE_REPORT        As String = "ddmmyy"
Private Const MAX_HEADER_LINES          As Long = 6
Private Const MAX_FOOTER_LINES          As Long = 1

Private m_sDevice                   As String
Private m_oConnector                As IDeviceConnector
Private m_lTimeout                  As Long
Private m_lSequence                 As Long
Private m_vInfoDiagnostics          As Variant
Private m_eError                    As Long
Private m_eStatus                   As Long
Private m_sPaymentTypes             As String
Private m_uPrintData                As UcsProtocolPrintData
Private m_uConfig                   As UcsConfigValues

Private Enum UcsInternalErrors
    [__ucsErrTimeoutWaitingForResponse]
    ucsErrInvalidMessageFormatOrChecksum
    ucsErrInvalidResponseLength
    ucsErrInvalidResponseChecksum
    ucsErrResponseMissingPreamble
    ucsErrNoReceiptStarted
    ucsErrErrorInLine
    ucsErrInvalidReportType
    ucsErrDeviceNotOnline
    ucsErrDeviceNotReady
    ucsErrPasswordNotSet
    ucsErrUnsupportedPaymentType
    ucsErrBarcodeNotSupported
    ucsErrCannotPrint
End Enum

Private Enum UcsDeviceErrorsAndStatuses
    '--- device errors
    ucsErrOutOfPaper = 1
    ucsErrBadPassword = 9
    ucsErrPrinterOverheated = 13
    '--- device statuses
    ucsErrIllegalCommand = 2
End Enum

Private Enum UcsReceiptTextsEnum
    ucsTxtPaymentCash
    [__ucsTxtPaymentCard]
    [__ucsTxtPaymentBank]
    ucsTxtSurcharge
    ucsTxtDiscount
    ucsTxtSubTotal
    ucsTxtPluSales
    ucsTxtTaxNoCaption
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

Private Type UcsAutodetect
    Connector           As cSerialPortConnector
    WriteOver           As OVERLAPPED
    ReadOver            As OVERLAPPED
    Buffer              As String
End Type

Private Type UcsConfigValues
    DefaultPassword     As String
    CodePage            As Long
    IsEcr               As VbTriState
    PingTimeout         As Long
    DetailedReceipt     As Boolean
    PrintVat            As Boolean
    DelayPrint          As Boolean
    BufferPrint         As Boolean
    MaxOperCode         As Long
    MaxOperName         As Long
    MaxOperPassword     As Long
    MaxCompany          As Long
    MaxPrsReceive       As Long
    MaxVatNo            As Long
    MaxTaxNo            As Long
    MaxCity             As Long
    MaxInvoiceNo        As Long
    MaxPluName          As Long
    MaxComment          As Long
    MaxFiscalMemoryNo   As Long
    MaxUniqueSaleNo     As Long
    MaxDepartmentCode   As Long
    MaxReceiptNo        As Long
    ExtendedPmtTypes    As Boolean
    ConfigCommands      As Collection
    LocalizedCommands   As Collection
    LocalizedText(0 To [_ucsFscLciMax] - 1) As Variant
End Type

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunction As String)
    Debug.Print "Critical error: " & Err.Description & " [" & MODULE_NAME & "." & sFunction & "]"
    Logger.Log vbLogEventTypeInformation, MODULE_NAME, sFunction & "(" & Erl & ")", "Run-time error: " & Err.Description
End Sub

Private Sub RaiseError(sFunction As String)
    Debug.Print "Critical error: " & Err.Description & " [" & MODULE_NAME & "." & sFunction & "]"
    Logger.Log vbLogEventTypeInformation, MODULE_NAME, sFunction & "(" & Erl & ")", "Run-time error: " & Err.Description
    Err.Raise Err.Number, MODULE_NAME & "." & sFunction & "(" & Erl & ")" & vbCrLf & Err.Source, Err.Description
End Sub

Private Sub DebugLog(sFunction As String, sText As String)
    Logger.Log vbLogEventTypeInformation, MODULE_NAME, sFunction, sText
End Sub

Private Sub DebugDataDump(sFunction As String, sPrefix As String, sData As String)
    Logger.DataDump MODULE_NAME, sFunction, sPrefix, sData
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
    If m_eStatus <> 0 Then
        StatusText = At(m_uConfig.LocalizedText(ucsFscLciPrinterStatuses), m_eStatus, At(m_uConfig.LocalizedText(ucsFscLciPrinterStatuses), 0))
    End If
End Property

Property Get ErrorNo() As Long
    ErrorNo = m_eError
End Property

Property Get ErrorText() As String
    If m_eError <> 0 Then
        ErrorText = At(m_uConfig.LocalizedText(ucsFscLciPrinterErrors), m_eError, At(m_uConfig.LocalizedText(ucsFscLciPrinterErrors), 0))
    End If
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
    m_uConfig.LocalizedText(eIdx) = Split(sValue, "|")
End Property

Property Get IsEcr() As Boolean
    Const FUNC_NAME     As String = "IsEcr [get]"
    Const CONF_NAME     As String = "IsEcr"
    Dim eCmd            As UcsIslCommandsEnum
    Dim sParam          As String
    Dim lTimeout        As Long
    Dim lTypeIndex      As Long
    Dim vInfoVersion    As Variant
    
    On Error GoTo EH
    If m_uConfig.IsEcr = vbUseDefault Then
        pvGetCommandConfig CONF_NAME, "InfoVersion", eCmd, ucsZekCmdInfoVersion, sParam, vbNullString, _
            "Timeout", lTimeout, 100, _
            "TypeIndex", lTypeIndex, 0
        vInfoVersion = Split(pvTransmitCommand(m_oConnector, lTimeout, eCmd, sParam), DEF_SEP)
        m_uConfig.IsEcr = IIf(At(vInfoVersion, lTypeIndex) = "1", vbTrue, vbFalse)
    End If
    IsEcr = (m_uConfig.IsEcr = vbTrue)
    Exit Property
EH:
    RaiseError FUNC_NAME
End Property

Private Property Get pvRowChars() As Long
    pvRowChars = pvZfplibValue(m_uPrintData.Config.RowChars, "FPLineWidth", DEF_ROW_CHARS)
End Property

Private Property Get pvCommentChars() As Long
    pvCommentChars = m_uPrintData.Config.CommentChars
    If pvCommentChars = 0 Then
        pvCommentChars = pvRowChars - 2
    End If
End Property

Private Property Get pvItemChars() As Long
    pvItemChars = m_uPrintData.Config.ItemChars
    If pvItemChars = 0 Then
        pvItemChars = pvRowChars - 6
    End If
End Property

Private Property Get pvPingTimeout() As Long
    pvPingTimeout = pvZfplibValue(m_uConfig.PingTimeout, "PingTimeout", DEF_PING_TIMEOUT)
End Property

Private Property Get DEF_EXTENDED_PMTTYPES() As Boolean
    DEF_EXTENDED_PMTTYPES = (Val(At(Split(m_sPaymentTypes, DEF_SEP), 12)) > 0)
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function AutodetectDevices(vPorts As Variant) As Variant
    Const FUNC_NAME     As String = "AutodetectDevices"
    Const BUFFER_SIZE   As Long = 1000
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
    Dim vSpeeds         As Variant
    
    On Error GoTo EH
    sProbeText = GetConfigValue(STR_PROTOCOL_TREMOL, "ProbeText", Chr$(ucsZekCmdFastIsReady))
    sReplyText = GetConfigValue(STR_PROTOCOL_TREMOL, "ReplyText", Chr$(ucsZekCmdFastIsReady))
    lTimeout = GetConfigNumber(STR_PROTOCOL_TREMOL, "AutodetectTimeout", DEF_AUTODETECTTIMEOUT)
    lDeviceInfoTimeout = GetConfigNumber(STR_PROTOCOL_TREMOL, "DeviceInfoTimeout", DEF_DEVICEINFOTIMEOUT)
    AssignVariant vSpeeds, GetConfigValue(STR_PROTOCOL_TREMOL, "AutodetectSpeeds", Split(DEF_AUTODETECTSPEEDS, "|"))
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
'                    Call WaitForMultipleObjects(1, .WriteOver.hEvent, 0, 10)
                    lResult = WriteFile(.Connector.hComm, ByVal sProbeText, Len(sProbeText), lActual, .WriteOver)
                End If
            End With
        Next
        For lIdx = 0 To UBound(aPorts)
            With aPorts(lIdx)
                If Not .Connector Is Nothing Then
                    If GetOverlappedResult(.Connector.hComm, .WriteOver, lActual, 1) <> 0 Then
'                        Debug.Print "Write GetOverlappedResult, vSpeed=" & vSpeed & ", lIdx=" & lIdx & ", lActual=" & lActual, TimerEx, MODULE_NAME
                    End If
                    .Buffer = String$(BUFFER_SIZE, 0)
                    Call WaitForMultipleObjects(1, .ReadOver.hEvent, 0, 10)
                    lResult = ReadFile(.Connector.hComm, ByVal StrPtr(.Buffer), LenB(.Buffer), lActual, .ReadOver)
                    If lResult = 0 And Err.LastDllError = ERROR_IO_PENDING Or lActual = 0 Then
                        hEvents(lEventsCount) = .ReadOver.hEvent
                        lEventsCount = lEventsCount + 1
                    Else
                        .Buffer = Left$(StrConv(.Buffer, vbUnicode), lActual)
                    End If
                End If
            End With
        Next
        If lEventsCount > 0 Then
            Call Sleep(10)
            lResult = WaitForMultipleObjects(lEventsCount, hEvents(0), 1, lTimeout)
            lEventsCount = 0
        End If
        For lIdx = 0 To UBound(aPorts)
            With aPorts(lIdx)
                If Not .Connector Is Nothing Then
                    If Len(.Buffer) = BUFFER_SIZE Then
                        If GetOverlappedResult(.Connector.hComm, .ReadOver, lActual, 1) <> 0 Then
'                            Debug.Print "Read GetOverlappedResult, vSpeed = " & vSpeed & ", lIdx = " & lIdx & ", lActual=" & lActual, TimerEx
                        End If
                        .Buffer = Left$(StrConv(.Buffer, vbUnicode), lActual)
                    End If
                    If LenB(.Buffer) <> 0 Then
                        If Right$(.Buffer, Len(sReplyText)) = sReplyText Then
                            DebugLog FUNC_NAME, "Auto-detected " & .Connector.Device & " reply 0x" & ToHexDump(.Buffer)
                            If pvGetDeviceInfo(.Connector, lDeviceInfoTimeout, Empty, sProtocol, sModel, sFirmware) Then
                                vRetVal(lIdx) = Array(At(vPorts, lIdx), vSpeed, sProtocol, sModel, sFirmware)
'                                Debug.Print "Auto-detected vSpeed=" & vSpeed & ", lIdx=" & lIdx & ", sProtocol=" & sProtocol & ", sModel=" & sModel & ", sFirmware=" & sFirmware
                            End If
                        Else
                            DebugLog FUNC_NAME, "Discarding " & .Connector.Device & " reply 0x" & ToHexDump(.Buffer)
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
    Dim eCmd            As UcsTremolCommandsEnum
    Dim sParam          As String
    Dim lSerialIndex    As Long
    
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
    If Not WaitDevice(-1) Then
        GoTo QH
    End If
    If Not WaitDevice(-1, CheckFastStatus:=True) Then
        Init = True
        GoTo QH
    End If
    pvGetCommandConfig CONF_NAME, "InfoPaymentTypes", eCmd, ucsZekCmdInfoPaymentTypes, sParam, vbNullString
    m_sPaymentTypes = SendCommand(eCmd, sParam)
    If LenB(m_uPrintData.LastError) <> 0 Then
        GoTo QH
    End If
    pvGetCommandConfig CONF_NAME, "InfoDiagnostics", eCmd, ucsZekCmdInfoDiagnostics, sParam, vbNullString, _
        "SerialIndex", lSerialIndex, 0
    m_vInfoDiagnostics = Split(pvTransmitCommand(m_oConnector, m_lTimeout, eCmd, sParam, m_uPrintData.LastError), DEF_SEP)
    If LenB(m_uPrintData.LastError) <> 0 Then
        GoTo QH
    End If
    '--- load config values
    sSerial = At(m_vInfoDiagnostics, lSerialIndex)
    With m_uPrintData.Config
        .RowChars = GetConfigNumber(sSerial, "RowChars", C_Lng(JsonItem(oOptions, "RowChars")))
        .CommentChars = GetConfigNumber(sSerial, "CommentChars", C_Lng(JsonItem(oOptions, "CommentChars")))
        .ItemChars = GetConfigNumber(sSerial, "ItemChars", C_Lng(JsonItem(oOptions, "ItemChars")))
        .AbsoluteDiscount = C_Bool(GetConfigValue(sSerial, "AbsoluteDiscount", JsonBoolItem(oOptions, "AbsoluteDiscount", True)))
        .NegativePrices = C_Bool(GetConfigValue(sSerial, "NegativePrices", JsonBoolItem(oOptions, "NegativePrices", True)))
        .MinDiscount = GetConfigNumber(sSerial, "MinDiscount", Zndbl(C_Dbl(JsonItem(oOptions, "MinDiscount")), DEF_MIN_DISCOUNT))
        .MaxDiscount = GetConfigNumber(sSerial, "MaxDiscount", Zndbl(C_Dbl(JsonItem(oOptions, "MaxDiscount")), DEF_MAX_DISCOUNT))
        .MaxReceiptRows = GetConfigNumber(sSerial, "MaxReceiptRows", Znl(C_Lng(JsonItem(oOptions, "MaxReceiptRows")), DEF_MAX_RECEIPT_ROWS))
    End With
    With m_uConfig
        .DefaultPassword = GetConfigValue(sSerial, "DefaultPassword", JsonItem(oOptions, "DefaultPassword"))
        .CodePage = GetConfigNumber(sSerial, "CodePage", Znl(C_Lng(JsonItem(oOptions, "CodePage")), DEF_CODEPAGE))
        .IsEcr = GetConfigNumber(sSerial, "IsEcr", Znl(C_Lng(JsonItem(oOptions, "IsEcr")), vbUseDefault))
        .PingTimeout = GetConfigNumber(sSerial, "PingTimeout", C_Lng(JsonItem(oOptions, "PingTimeout")))
        .DetailedReceipt = C_Bool(GetConfigValue(sSerial, "DetailedReceipt", JsonBoolItem(oOptions, "DetailedReceipt")))
        .PrintVat = C_Bool(GetConfigValue(sSerial, "PrintVat", JsonBoolItem(oOptions, "PrintVat")))
        .DelayPrint = C_Bool(GetConfigValue(sSerial, "DelayPrint", JsonBoolItem(oOptions, "DelayPrint", Default:=True)))
        .BufferPrint = C_Bool(GetConfigValue(sSerial, "BufferPrint", JsonBoolItem(oOptions, "BufferPrint")))
        .MaxOperCode = GetConfigNumber(sSerial, "MaxOperCode", Znl(C_Lng(JsonItem(oOptions, "MaxOperCode")), DEF_MAX_OPERCODE))
        .MaxOperName = GetConfigNumber(sSerial, "MaxOperName", Znl(C_Lng(JsonItem(oOptions, "MaxOperName")), DEF_MAX_OPERNAME))
        .MaxOperPassword = GetConfigNumber(sSerial, "MaxOperPassword", Znl(C_Lng(JsonItem(oOptions, "MaxOperPassword")), DEF_MAX_OPERPASSWORD))
        .MaxCompany = GetConfigNumber(sSerial, "MaxCompany", Znl(C_Lng(JsonItem(oOptions, "MaxCompany")), DEF_MAX_COMPANY))
        .MaxPrsReceive = GetConfigNumber(sSerial, "MaxPrsReceive", Znl(C_Lng(JsonItem(oOptions, "MaxPrsReceive")), DEF_MAX_PRSRECEIVE))
        .MaxVatNo = GetConfigNumber(sSerial, "MaxVatNo", Znl(C_Lng(JsonItem(oOptions, "MaxVatNo")), DEF_MAX_VATNO))
        .MaxTaxNo = GetConfigNumber(sSerial, "MaxTaxNo", Znl(C_Lng(JsonItem(oOptions, "MaxTaxNo")), DEF_MAX_TAXNO))
        .MaxCity = GetConfigNumber(sSerial, "MaxCity", Znl(C_Lng(JsonItem(oOptions, "MaxCity")), DEF_MAX_CITY))
        .MaxInvoiceNo = GetConfigNumber(sSerial, "MaxInvoiceNo", Znl(C_Lng(JsonItem(oOptions, "MaxInvoiceNo")), DEF_MAX_INVOICENO))
        .MaxPluName = GetConfigNumber(sSerial, "MaxPluName", Znl(C_Lng(JsonItem(oOptions, "MaxPluName")), DEF_MAX_PLUNAME))
        .MaxComment = GetConfigNumber(sSerial, "MaxComment", Znl(C_Lng(JsonItem(oOptions, "MaxComment")), DEF_MAX_COMMENT))
        .MaxFiscalMemoryNo = GetConfigNumber(sSerial, "MaxFiscalMemoryNo", Znl(C_Lng(JsonItem(oOptions, "MaxFiscalMemoryNo")), DEF_MAX_FISCALMEMORYNO))
        .MaxUniqueSaleNo = GetConfigNumber(sSerial, "MaxUniqueSaleNo", Znl(C_Lng(JsonItem(oOptions, "MaxUniqueSaleNo")), DEF_MAX_UNIQUESALENO))
        .MaxDepartmentCode = GetConfigNumber(sSerial, "MaxDepartmentCode", Znl(C_Lng(JsonItem(oOptions, "MaxDepartmentCode")), DEF_MAX_DEPARTMENTCODE))
        .MaxReceiptNo = GetConfigNumber(sSerial, "MaxReceiptNo", Znl(C_Lng(JsonItem(oOptions, "MaxReceiptNo")), DEF_MAX_RECEIPTNO))
        If Not IsEmpty(JsonItem(oOptions, "ExtendedPmtTypes")) Then
            .ExtendedPmtTypes = JsonBoolItem(oOptions, "ExtendedPmtTypes")
        Else
            .ExtendedPmtTypes = GetConfigNumber(sSerial, "ExtendedPmtTypes", DEF_EXTENDED_PMTTYPES)
        End If
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
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function SendCommand(ByVal eCmd As UcsTremolCommandsEnum, Optional Data As String) As String
    Const FUNC_NAME     As String = "SendCommand"
    
    On Error GoTo EH
    pvSetLastError vbNullString
    If eCmd = ucsZekCmdNoop Then
        Exit Function
    End If
    If WaitDevice() Then
        SendCommand = pvTransmitCommand(m_oConnector, m_lTimeout, eCmd, Data, m_uPrintData.LastError)
    End If
    If LenB(m_uPrintData.LastError) <> 0 Then
        m_uPrintData.LastErrNo = IIf(m_eError = ucsErrBadPassword, ucsFerInvalidPassword, ucsFerGeneralError)
    End If
    '--- raise event
    RaiseEvent CommandComplete(eCmd, Data, SendCommand)
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function WaitDevice(Optional ByVal Timeout As Long, Optional CheckFastStatus As Boolean) As Boolean
    Const FUNC_NAME     As String = "WaitDevice"
    Const CONF_NAME     As String = FUNC_NAME
    Dim sError          As String
    Dim eCmd            As UcsTremolCommandsEnum
    Dim sParam          As String
    Dim sStatus         As String
    Dim dblTimer        As Double
    Dim sCurrent        As String
    Dim lIter           As Long
    Dim sDelay          As String
    
    On Error GoTo EH
    pvSetLastError vbNullString
    If Timeout < 0 Then
        pvGetCommandConfig CONF_NAME, "FastIsOnline", eCmd, ucsZekCmdFastIsOnline, sParam, vbNullString, _
            "Status", sStatus, Chr$(ucsZekCmdFastIsOnline)
        If pvTransmitCommand(m_oConnector, pvPingTimeout, eCmd, sParam, sError) <> sStatus Then
            pvSetLastError pvInternal(ucsErrDeviceNotOnline)
            GoTo QH
        End If
        If CheckFastStatus Then
            pvGetCommandConfig CONF_NAME, "FastStatus", eCmd, ucsZekCmdFastStatus, sParam, vbNullString, _
                "ReadyStatus", sStatus, LNG_STATUS_READY
            sCurrent = pvTransmitCommand(m_oConnector, pvPingTimeout, eCmd, sParam, sError)
            If LenB(sCurrent) <> 0 Then
                Select Case Asc(sCurrent)
                Case C_Lng(sStatus) '--- LNG_STATUS_READY
                    '--- success
                    WaitDevice = True
                    If lIter > 2 Then
                        DebugLog FUNC_NAME, "Ready after " & lIter
                    End If
                Case LNG_STATUS_IS_BUSY
                    pvSetLastError pvInternal(ucsErrDeviceNotOnline)
                Case LNG_STATUS_OUT_OF_PAPER, LNG_STATUS_OUT_OF_PAPER_LEGACY, LNG_STATUS_OUT_OF_PAPER_AND_IS_BUSY
                    pvSetLastError At(m_uConfig.LocalizedText(ucsFscLciPrinterErrors), ucsErrOutOfPaper)
                Case LNG_STATUS_OVERHEATED
                    pvSetLastError At(m_uConfig.LocalizedText(ucsFscLciPrinterErrors), ucsErrPrinterOverheated)
                End Select
            End If
        Else
            '--- success
            WaitDevice = True
        End If
    Else
        Timeout = Znl(Timeout, m_lTimeout)
        dblTimer = TimerEx + Timeout / 1000#
        Do While TimerEx < dblTimer
            lIter = lIter + 1
            pvGetCommandConfig CONF_NAME, "FastIsOnline", eCmd, ucsZekCmdFastIsOnline, sParam, vbNullString, _
                "Status", sStatus, Chr$(ucsZekCmdFastIsOnline)
            If pvTransmitCommand(m_oConnector, pvPingTimeout, eCmd, sParam, sError) = sStatus Then
                lIter = lIter + 1
                pvGetCommandConfig CONF_NAME, "FastStatus", eCmd, ucsZekCmdFastStatus, sParam, vbNullString, _
                    "ReadyStatus", sStatus, LNG_STATUS_READY, _
                    "Delay", sDelay, 40
                sCurrent = pvTransmitCommand(m_oConnector, pvPingTimeout, eCmd, sParam, sError)
                If LenB(sCurrent) <> 0 Then
                    Select Case Asc(sCurrent)
                    Case C_Lng(sStatus) '--- LNG_STATUS_READY
                        '--- success
                        WaitDevice = True
                        If lIter > 2 Then
                            DebugLog FUNC_NAME, "Ready after " & lIter
                        End If
                        GoTo QH
                    End Select
                End If
                Call Sleep(C_Lng(sDelay))
            End If
        Loop
        m_uPrintData.LastError = pvInternal(ucsErrDeviceNotReady)
    End If
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function CopyLastReceipt(Optional ReceiptNo As String) As Boolean
    Const FUNC_NAME     As String = "CopyLastReceipt"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsTremolCommandsEnum
    Dim sParam          As String
    Dim sPrevReceiptNo  As String
    Dim sReceiptNo      As String
    Dim lPrintTimeout   As Long
    
    On Error GoTo EH
    sPrevReceiptNo = GetLastReceiptNo()
    If LenB(ReceiptNo) = 0 Then
        pvGetCommandConfig CONF_NAME, "PrintDuplicate", eCmd, ucsZekCmdPrintDuplicate, sParam, vbNullString, _
            "PrintTimeout", lPrintTimeout, 0
        SendCommand eCmd, sParam
    Else
        pvGetCommandConfig CONF_NAME, "PrintReportDaily", eCmd, ucsZekCmdPrintReportDaily, sParam, "J1" & DEF_SEP & "N" & DEF_SEP & "%1" & DEF_SEP & "%1", _
            "PrintTimeout", lPrintTimeout, 0
        SendCommand eCmd, Printf(sParam, Pad(ReceiptNo, -m_uConfig.MaxReceiptNo))
    End If
    If Not WaitDevice(Znl(lPrintTimeout, DEF_PRINT_TIMEOUT)) Then
        GoTo QH
    End If
    sReceiptNo = GetLastReceiptNo()
    If LenB(sReceiptNo) = 0 Or sReceiptNo = sPrevReceiptNo Then
        If LenB(m_uPrintData.LastError) = 0 Then
            pvSetLastError pvInternal(ucsErrCannotPrint)
        End If
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
    Dim eCmd            As UcsTremolCommandsEnum
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
    Dim eCmd            As UcsTremolCommandsEnum
    Dim sParam          As String
    Dim sFormat         As String
    Dim lPrintTimeout   As Long
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InitDateTime", eCmd, ucsZekCmdInitDateTime, sParam, "%1", _
        "Format", sFormat, FORMAT_DATE_EXTENDED, _
        "PrintTimeout", lPrintTimeout, 0
    sParam = Printf(sParam, Format$(NewDate, sFormat))
    SendCommand eCmd, sParam
    '--- success (or failure)
    SetClock = WaitDevice(Znl(lPrintTimeout, DEF_PRINT_TIMEOUT))
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function RunXReport(ByVal ReportType As UcsFiscalReportsTypeEnum) As Boolean
    Const FUNC_NAME     As String = "RunXReport"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsTremolCommandsEnum
    Dim sParam          As String
    Dim sPrevReceiptNo  As String
    Dim sReceiptNo      As String
    Dim lPrintTimeout   As Long
    
    If Not CancelReceipt() Then
        GoTo QH
    End If
    On Error GoTo EH
    sPrevReceiptNo = GetLastReceiptNo()
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
    If Not WaitDevice(Znl(lPrintTimeout, DEF_PRINT_TIMEOUT)) Then
        GoTo QH
    End If
    sReceiptNo = GetLastReceiptNo()
    If LenB(sReceiptNo) = 0 Or sReceiptNo = sPrevReceiptNo Then
        If LenB(m_uPrintData.LastError) = 0 Then
            pvSetLastError pvInternal(ucsErrCannotPrint)
        End If
        GoTo QH
    End If
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
    Dim eCmd            As UcsTremolCommandsEnum
    Dim sParam          As String
    Dim sPrevReceiptNo  As String
    Dim sReceiptNo      As String
    Dim lPrintTimeout   As Long
    
    If Not CancelReceipt() Then
        GoTo QH
    End If
    On Error GoTo EH
    sPrevReceiptNo = GetLastReceiptNo()
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
    If Not WaitDevice(Znl(lPrintTimeout, DEF_PRINT_TIMEOUT)) Then
        GoTo QH
    End If
    sReceiptNo = GetLastReceiptNo()
    If LenB(sReceiptNo) = 0 Or sReceiptNo = sPrevReceiptNo Then
        If LenB(m_uPrintData.LastError) = 0 Then
            pvSetLastError pvInternal(ucsErrCannotPrint)
        End If
        GoTo QH
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
    Dim eCmd            As UcsTremolCommandsEnum
    Dim sParam          As String
    Dim sFormat         As String
    Dim sPrevReceiptNo  As String
    Dim sReceiptNo      As String
    Dim lPrintTimeout   As Long
    
    If Not CancelReceipt() Then
        GoTo QH
    End If
    On Error GoTo EH
    sPrevReceiptNo = GetLastReceiptNo()
    Select Case ReportType
    Case ucsFscRstPeriodShort
        pvGetCommandConfig CONF_NAME, "PrintReportByDateShort", eCmd, ucsZekCmdPrintReportByDateShort, sParam, "%1" & DEF_SEP & "%2", _
            "Format", sFormat, FORMAT_DATE_REPORT, _
            "PrintTimeout", lPrintTimeout, 0
    Case ucsFscRstPeriodDetailed
        pvGetCommandConfig CONF_NAME, "PrintReportByDateDetailed", eCmd, ucsZekCmdPrintReportByDateDetailed, sParam, "%1" & DEF_SEP & "%2", _
            "Format", sFormat, FORMAT_DATE_REPORT, _
            "PrintTimeout", lPrintTimeout, 0
    Case Else
        pvSetLastError Printf(pvInternal(ucsErrInvalidReportType), ReportType)
        GoTo QH
    End Select
    sParam = Printf(sParam, Format$(StartDate, sFormat), Format$(EndDate, sFormat))
    SendCommand eCmd, sParam
    If Not WaitDevice(Znl(lPrintTimeout, DEF_PRINT_TIMEOUT)) Then
        GoTo QH
    End If
    sReceiptNo = GetLastReceiptNo()
    If LenB(sReceiptNo) = 0 Or sReceiptNo = sPrevReceiptNo Then
        If LenB(m_uPrintData.LastError) = 0 Then
            pvSetLastError pvInternal(ucsErrCannotPrint)
        End If
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
    Dim eCmd            As UcsTremolCommandsEnum
    Dim sParam          As String
    Dim lIndex          As Long
    Dim vSplit          As Variant
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoLastReceipt", eCmd, ucsZekCmdInfoLastReceipt, sParam, vbNullString, _
        "Index", lIndex, 0
    vSplit = Split(SendCommand(eCmd, sParam), DEF_SEP)
    '--- success
    GetLastReceiptNo = RTrim$(At(vSplit, lIndex))
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetLastInvoiceNo() As String
    Const FUNC_NAME     As String = "GetLastInvoiceNo"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsTremolCommandsEnum
    Dim sParam          As String
    Dim lIndex          As Long
    Dim vSplit          As Variant
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoInvoiceNo", eCmd, ucsZekCmdInfoInvoiceNo, sParam, vbNullString, _
        "Index", lIndex, 0
    vSplit = Split(SendCommand(eCmd, sParam), DEF_SEP)
    If Not IsNumeric(At(vSplit, lIndex)) Then
        GoTo QH
    End If
    '--- success
    GetLastInvoiceNo = C_Dbl(At(vSplit, lIndex)) - 1
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetLastQRCodeInfo(Optional LastReceiptDate As Date) As String
    Const FUNC_NAME     As String = "GetLastQRCodeInfo"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsTremolCommandsEnum
    Dim sParam          As String
    Dim lIndex          As Long
    Dim lDateIndex      As Long
    Dim sDelimiter      As String
    Dim vSplit          As Variant
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoTransaction", eCmd, ucsZekCmdInfoTransaction, sParam, "B", _
        "Index", lIndex, 1, _
        "DateIndex", lDateIndex, 2, _
        "Delimiter", sDelimiter, "*"
    vSplit = Split(SendCommand(eCmd, sParam), sDelimiter)
    LastReceiptDate = C_Date(At(vSplit, lDateIndex))
    '--- success
    GetLastQRCodeInfo = RTrim$(At(vSplit, lIndex))
QH:
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
        "Index", lIndex, 0
    GetDeviceSerialNo = RTrim$(At(m_vInfoDiagnostics, lIndex))
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
        "Index", lIndex, 1
    GetFiscalMemoryNo = RTrim$(At(m_vInfoDiagnostics, lIndex))
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetTaxNo() As String
    Const FUNC_NAME     As String = "GetTaxNo"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsTremolCommandsEnum
    Dim sParam          As String
    Dim lIndex          As Long
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoBulstat", eCmd, ucsZekCmdInfoBulstat, sParam, vbNullString, _
        "Index", lIndex, 0
    GetTaxNo = RTrim$(At(Split(SendCommand(eCmd, sParam), DEF_SEP), lIndex))
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetTaxCaption() As String
    Const FUNC_NAME     As String = "GetTaxCaption"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsTremolCommandsEnum
    Dim sParam          As String
    Dim lIndex          As Long
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoBulstat", eCmd, ucsZekCmdInfoBulstat, sParam, vbNullString, _
        "Index", lIndex, 1
    lIndex = C_Lng(At(Split(SendCommand(eCmd, sParam), DEF_SEP), lIndex))
    GetTaxCaption = RTrim$(At(Split(pvText(ucsTxtTaxNoCaption), DEF_SEP), lIndex))
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetHeaderText(ByVal Index As Long) As String
    Const FUNC_NAME     As String = "GetHeaderText"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsTremolCommandsEnum
    Dim sParam          As String
    Dim lIndex          As Long
    
    On Error GoTo EH
    If Index >= 1 And Index <= MAX_HEADER_LINES Then
        pvGetCommandConfig CONF_NAME, "InfoHeaderFooter" & Index, eCmd, ucsZekCmdInfoHeaderFooter, sParam, C_Str(Index), _
            "Index", lIndex, 1
        GetHeaderText = RTrim$(At(Split(SendCommand(eCmd, sParam), DEF_SEP), lIndex))
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetFooterText(ByVal Index As Long) As String
    Const FUNC_NAME     As String = "GetFooterText"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsTremolCommandsEnum
    Dim sParam          As String
    Dim lIndex          As Long
    
    On Error GoTo EH
    If Index >= 1 And Index <= MAX_FOOTER_LINES Then
        pvGetCommandConfig CONF_NAME, "InfoHeaderFooter" & Index, eCmd, ucsZekCmdInfoHeaderFooter, sParam, C_Str(7 + Index), _
            "Index", lIndex, 1
        GetFooterText = RTrim$(At(Split(SendCommand(eCmd, sParam), DEF_SEP), lIndex))
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetPaymentName(ByVal PmtType As UcsFiscalPaymentTypeEnum) As String
    Const FUNC_NAME     As String = "GetPaymentName"
    
    On Error GoTo EH
    If PmtType >= ucsFscPmtCash And PmtType <= ucsFscPmtCash Then
        GetPaymentName = pvText(ucsTxtPaymentCash + PmtType - ucsFscPmtCash)
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
    Dim eCmd            As UcsTremolCommandsEnum
    Dim sParam          As String
    Dim lPrintTimeout   As Long
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InitOpenDrawer", eCmd, ucsZekCmdInitOpenDrawer, sParam, vbNullString, _
        "PrintTimeout", lPrintTimeout, 0
    SendCommand eCmd, sParam
    '--- success (or failure)
    OpenDrawer = WaitDevice(Znl(lPrintTimeout, DEF_PRINT_TIMEOUT))
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function CancelReceipt() As Boolean
    Const FUNC_NAME     As String = "CancelReceipt"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eTransCmd       As UcsTremolCommandsEnum
    Dim sTransParam     As String
    Dim lTransIndex     As Long
    Dim lTransStatus    As Long
    Dim eCmd            As UcsTremolCommandsEnum
    Dim sParam          As String
    Dim lPrintTimeout   As Long
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoTransaction", eTransCmd, ucsZekCmdInfoTransaction, sTransParam, vbNullString, _
        "Index", lTransIndex, 0, _
        "Status", lTransStatus, 0
    If C_Lng(At(Split(SendCommand(eTransCmd, sTransParam), DEF_SEP), lTransIndex)) = lTransStatus Then
        '--- nothing to do
        CancelReceipt = True
        GoTo QH
    End If
    pvGetCommandConfig CONF_NAME, "NonFiscalClose", eCmd, ucsZekCmdNonFiscalClose, sParam, vbNullString, _
        "PrintTimeout", lPrintTimeout, 0
    SendCommand eCmd, sParam
    If LenB(m_uPrintData.LastError) <> 0 Then
        pvGetCommandConfig CONF_NAME, "FiscalCancel", eCmd, ucsZekCmdFiscalCancel, sParam, vbNullString
        SendCommand eCmd, sParam
        If LenB(m_uPrintData.LastError) <> 0 Then
            pvGetCommandConfig CONF_NAME, "FiscalPayment", eCmd, ucsZekCmdFiscalPayment, sParam, "0" & DEF_SEP & "0" & DEF_SEP & "-"
            SendCommand eCmd, sParam
            pvGetCommandConfig CONF_NAME, "FiscalCancel", eCmd, ucsZekCmdFiscalCancel, sParam, vbNullString
            SendCommand eCmd, sParam
        End If
        If LenB(m_uPrintData.LastError) <> 0 Then
            pvGetCommandConfig CONF_NAME, "FiscalPayAndClose", eCmd, ucsZekCmdFiscalPayAndClose, sParam, vbNullString
            SendCommand eCmd, sParam
        End If
    End If
    If Not WaitDevice(Znl(lPrintTimeout, DEF_PRINT_TIMEOUT)) Then
        GoTo QH
    End If
    '--- success (or failure)
    CancelReceipt = C_Lng(At(Split(SendCommand(eTransCmd, sTransParam), DEF_SEP), lTransIndex)) = lTransStatus
QH:
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
    
    On Error GoTo EH
    WaitDevice -1, CheckFastStatus:=True
    CurrentStatus = m_uPrintData.LastError
    '--- success
    GetDeviceStatus = True
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetTotalsByPaymentTypes() As Variant
    Const FUNC_NAME     As String = "GetTotalsByPaymentTypes"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsTremolCommandsEnum
    Dim sParam          As String
    Dim vPayments       As Variant
    Dim vChanges        As Variant
    Dim vRetVal(0 To MAX_PMT_TYPE - 1) As Variant
    Dim lIdx            As Long
    Dim lPmtIndex       As Long
    Dim lOffset         As Long
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoRegisters4", eCmd, ucsZekCmdInfoRegisters, sParam, "4", _
        "Offset", lOffset, 1
    vPayments = Split(SendCommand(eCmd, sParam), DEF_SEP)
    pvGetCommandConfig CONF_NAME, "InfoRegisters6", eCmd, ucsZekCmdInfoRegisters, sParam, "6"
    vChanges = Split(SendCommand(eCmd, sParam), DEF_SEP)
    If UBound(vPayments) <= 0 Then
        GoTo QH
    End If
    For lIdx = 0 To UBound(vRetVal)
        lPmtIndex = pvToDevicePmtIndex(lIdx + 1)
        If lPmtIndex >= 0 Then
            vRetVal(lIdx) = Array(pvToDevicePmtName(lIdx + 1), ParseSum(At(vPayments, lPmtIndex + lOffset)) - ParseSum(At(vChanges, lPmtIndex + lOffset)))
        End If
    Next
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
    Dim vRetVal(0 To MAX_TAX_GROUP - 1) As Variant
    Dim lIdx            As Long
    Dim eCmd            As UcsTremolCommandsEnum
    Dim sParam          As String
    Dim lRevOffset      As Long
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoTaxRates", eCmd, ucsZekCmdInfoTaxRates, sParam, vbNullString
    vTaxRates = Split(Replace(SendCommand(eCmd, sParam), "%", vbNullString), DEF_SEP)
    pvGetCommandConfig CONF_NAME, "InfoTotalsByTaxGroups", eCmd, ucsZekCmdInfoTotalsByTaxGroups, sParam, vbNullString, _
        "RevOffset", lRevOffset, 9
    vPmtSum = Split(SendCommand(eCmd, sParam), DEF_SEP)
    If UBound(vPmtSum) <= 0 Then
        GetTotalsByTaxGroups = Array()
        GoTo QH
    End If
    For lIdx = 0 To UBound(vRetVal)
        vRetVal(lIdx) = Array(C_Dbl(At(vTaxRates, lIdx)), ParseSum(At(vPmtSum, lIdx)), ParseSum(At(vPmtSum, lIdx + lRevOffset)))
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
    Dim eCmd            As UcsTremolCommandsEnum
    Dim sParam          As String
    Dim lIndex          As Long
    Dim vSplit          As Variant
    Dim lReceipts       As Long
    Dim dLastDate       As Date
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoRegisters1", eCmd, ucsZekCmdInfoRegisters, sParam, "1", _
        "Index", lIndex, 1
    vSplit = Split(SendCommand(eCmd, sParam), DEF_SEP)
    lReceipts = C_Lng(At(vSplit, lIndex))
    pvGetCommandConfig CONF_NAME, "InfoRegisters4", eCmd, ucsZekCmdInfoRegisters, sParam, "5", _
        "Index", lIndex, 4
    vSplit = Split(SendCommand(eCmd, sParam), DEF_SEP)
    dLastDate = C_Date(At(vSplit, lIndex))
    GetTotalsStats = Array(lReceipts, IIf(dLastDate = 0 Or dLastDate = #1/1/2008#, Empty, dLastDate))
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function PrintServiceDeposit( _
        Optional ByVal OperatorCode As String, _
        Optional ByVal OperatorPassword As String, _
        Optional ByVal PmtType As UcsFiscalPaymentTypeEnum, _
        Optional ByVal Amount As Double, _
        Optional Comment As String) As Variant
    Const FUNC_NAME     As String = "PrintServiceDeposit"
    Const CONF_NAME     As String = FUNC_NAME
    Dim vSplit          As Variant
    Dim vRetVal(0 To 2) As Double
    Dim eCmd            As UcsTremolCommandsEnum
    Dim sParam          As String
    Dim sFormat         As String
    Dim lIndex          As Long
    Dim lPrintTimeout   As Long
    Dim lPmtIndex       As Long
    
    On Error GoTo EH
    lPmtIndex = pvToDevicePmtIndex(PmtType)
    If Abs(Amount) > DBL_EPSILON Then
        pvGetCommandConfig CONF_NAME, "InfoParameters", eCmd, ucsZekCmdInfoParameters, sParam, vbNullString, _
            "Index", lIndex, 9
        vSplit = Split(SendCommand(eCmd, sParam), DEF_SEP)
        If C_Bool(At(vSplit, lIndex)) Then
            '--- use single operator
            OperatorCode = DEF_OPERNO
            OperatorPassword = GetDefaultPassword(DEF_OPERNO)
        End If
        pvGetCommandConfig CONF_NAME, "FiscalServiceDeposit", eCmd, ucsZekCmdFiscalServiceDeposit, sParam, "%1" & DEF_SEP & "%2" & DEF_SEP & "%3" & DEF_SEP & "%4@%5", _
            "Format", sFormat, FORMAT_BASE_2, _
            "PrintTimeout", lPrintTimeout, 0
        sParam = Printf(sParam, Pad(Zn(OperatorCode, DEF_OPERNO), -m_uConfig.MaxOperCode), _
            Pad(Zn(OperatorPassword, GetDefaultPassword(Zn(OperatorCode, DEF_OPERNO))), m_uConfig.MaxOperPassword), _
            lPmtIndex, _
            SafeFormat(Amount, sFormat), _
            Left$(Comment, m_uConfig.MaxComment))
        SendCommand eCmd, sParam
        If LenB(m_uPrintData.LastError) <> 0 Then
            GoTo QH
        End If
        If Not WaitDevice(Znl(lPrintTimeout, DEF_PRINT_TIMEOUT)) Then
            GoTo QH
        End If
    End If
    pvGetCommandConfig CONF_NAME, "InfoRegisters0", eCmd, ucsZekCmdInfoRegisters, sParam, "0"
    vSplit = Split(SendCommand(eCmd, sParam), DEF_SEP)
    If LenB(m_uPrintData.LastError) <> 0 Then
        GoTo QH
    End If
    vRetVal(0) = C_Dbl(At(vSplit, lPmtIndex + 1))
    pvGetCommandConfig CONF_NAME, "InfoRegisters2", eCmd, ucsZekCmdInfoRegisters, sParam, "2"
    vSplit = Split(SendCommand(eCmd, sParam), DEF_SEP)
    vRetVal(1) = C_Dbl(At(vSplit, lPmtIndex + 1))
    pvGetCommandConfig CONF_NAME, "InfoRegisters3", eCmd, ucsZekCmdInfoRegisters, sParam, "3"
    vSplit = Split(SendCommand(eCmd, sParam), DEF_SEP)
    vRetVal(2) = C_Dbl(At(vSplit, lPmtIndex + 1))
    '--- success
    PrintServiceDeposit = vRetVal
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function PrintBarcode(ByVal eType As UcsFiscalBarcodeTypeEnum, ByVal sCodeData As String, Optional ByVal Height As Long) As Boolean
    Const FUNC_NAME     As String = "PrintBarcode"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsTremolCommandsEnum
    Dim sParam          As String
    Dim sCodeType       As String
    
    On Error GoTo EH
    #If Height Then '--- touch args
    #End If
    Select Case eType
    Case ucsFscBrcEan8
        sCodeType = "D"
        sCodeData = Pad(sCodeData, 7, "0")
    Case ucsFscBrcEan13
        sCodeType = "C"
        sCodeData = Pad(sCodeData, 12, "0")
    Case ucsFscBrcCode128
        sCodeType = "I"
    Case Else
        pvSetLastError Printf(pvInternal(ucsErrBarcodeNotSupported), eType)
        GoTo QH
    End Select
    pvGetCommandConfig CONF_NAME, "PrintBarcode", eCmd, ucsZekCmdPrintBarcode, sParam, "P" & DEF_SEP & "%1" & DEF_SEP & "%2" & DEF_SEP & "%3"
    sParam = Printf(sParam, sCodeType, Len(sCodeData), sCodeData)
    SendCommand eCmd, sParam
    If LenB(m_uPrintData.LastError) <> 0 Then
        GoTo QH
    End If
    '--- success
    PrintBarcode = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Public Sub SetLocalizedCommand( _
            sFunction As String, _
            sKey As String, _
            Optional ByVal cmd As UcsTremolCommandsEnum, _
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
    Dim eCmd            As UcsTremolCommandsEnum
    Dim sParam          As String
    Dim lIndex          As Long
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoOperator", eCmd, ucsZekCmdInfoOperator, sParam, "%1", _
        "Index", lIndex, 1
    sParam = Printf(sParam, Pad(OperatorCode, -m_uConfig.MaxOperCode))
    GetOperatorName = RTrim$(At(Split(SendCommand(eCmd, sParam), DEF_SEP), lIndex))
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetDefaultPassword(OperatorCode As String) As String
    If LenB(m_uConfig.DefaultPassword) <> 0 Then
        GetDefaultPassword = m_uConfig.DefaultPassword
    ElseIf IsEcr Then
        If C_Lng(OperatorCode) > 1 Then
            GetDefaultPassword = OperatorCode
        Else
            GetDefaultPassword = DEF_ECR_PASSWORD
        End If
    Else
        GetDefaultPassword = DEF_FP_PASSWORD
    End If
End Function

Public Function GetDepartmentName(DepartmentCode As String, Optional TaxGroup As Long) As String
    Const FUNC_NAME     As String = "GetOperatorName"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsTremolCommandsEnum
    Dim sParam          As String
    Dim lIndex          As Long
    Dim lTaxIndex       As Long
    Dim vSplit          As Variant
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoDepartment", eCmd, ucsZekCmdInfoDepartment, sParam, "%1", _
        "Index", lIndex, 1, _
        "TaxIndex", lTaxIndex, 2
    sParam = Printf(sParam, Pad(DepartmentCode, -m_uConfig.MaxDepartmentCode))
    vSplit = Split(SendCommand(eCmd, sParam), DEF_SEP)
    If UBound(vSplit) <= 0 Then
        GoTo QH
    End If
    '--- success
    GetDepartmentName = RTrim$(At(vSplit, lIndex))
    TaxGroup = Limit(Asc(At(vSplit, lTaxIndex) & " ") - 191, 0, 8)
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetTaxRates() As Variant
    Const FUNC_NAME     As String = "GetTaxRates"
    Const CONF_NAME     As String = FUNC_NAME
    Dim eCmd            As UcsTremolCommandsEnum
    Dim sParam          As String
    Dim vTaxRates       As Variant
    Dim vRetVal(0 To MAX_TAX_GROUP - 1) As Double
    Dim lIdx            As Long
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoTaxRates", eCmd, ucsZekCmdInfoTaxRates, sParam, vbNullString
    vTaxRates = Split(Replace(SendCommand(eCmd, sParam), "%", vbNullString), DEF_SEP)
    If UBound(vTaxRates) <= 0 Then
        GoTo QH
    End If
    For lIdx = 0 To UBound(vRetVal)
        vRetVal(lIdx) = C_Dbl(At(vTaxRates, lIdx))
    Next
    '--- success
    GetTaxRates = vRetVal
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

'= private ===============================================================

Private Function pvTransmitCommand( _
            oConnector As IDeviceConnector, _
            ByVal lTimeout As Long, _
            ByVal eCmd As UcsTremolCommandsEnum, _
            Optional Data As String, _
            Optional Error As String) As String
    Const FUNC_NAME     As String = "pvTransmitCommand"
    Dim sText           As String
    Dim baText()        As Byte
    Dim lIdx            As Long
    Dim lSize           As Long
    Dim bLogging        As Boolean
    Dim sBuffer         As String
    Dim dblTimerEx      As Double
    
    On Error GoTo EH
    bLogging = (eCmd > [_ucsZekCmdFastMax])
    If bLogging Then
        DebugLog FUNC_NAME, "eCmd=0x" & Hex$(eCmd) & IIf(LenB(Data) <> 0, ", Data=" & Data, vbNullString)
    End If
    dblTimerEx = TimerEx
RetrySend:
    If TimerEx > dblTimerEx + lTimeout / 1000# Then
        GoTo QH
    End If
    '--- prepare command
    If eCmd <= [_ucsZekCmdFastMax] Then
        sText = Chr$(eCmd)
    ElseIf eCmd = ucsZekCmdInitLogo Or eCmd = ucsZekCmdInitLogoByNum Then
        sText = Chr$(ZFP_STX) & Chr$(&H39) & Chr$(&H37) & Chr$(eCmd) & Data
        lTimeout = 60000
    Else
        sText = Chr$(ZFP_STX) & Chr$(ZFP_BASEDATA + 3 + Len(Data)) & Chr$(m_lSequence) & Chr$(eCmd) & Data
        '--- increment sequence
        m_lSequence = IIf(m_lSequence >= LNG_MAX_SEQUENCE, DEF_SEQUENCE, m_lSequence + 1)
        '--- append checksum and terminator
        sText = sText & pvCheckSum(sText) & Chr$(ZFP_ETX)
    End If
    '--- send request
    DebugDataDump FUNC_NAME, "[SEND ] ", sText
    baText = ToAscii(sText, m_uConfig.CodePage)
    If Not oConnector.WriteData(baText, lTimeout) Then
        Error = oConnector.GetLastError()
        DebugLog FUNC_NAME, "WriteData failed, Error=" & Error
        GoTo QH
    End If
    '--- read response
    sText = vbNullString
    Do
        If TimerEx > dblTimerEx + lTimeout / 1000# Then
            GoTo QH
        End If
        If Not oConnector.ReadData(baText, lTimeout) Then
            Error = oConnector.GetLastError()
            DebugLog FUNC_NAME, "ReadData failed, Error=" & Error
            GoTo QH
        End If
        sBuffer = FromAscii(baText, m_uConfig.CodePage)
        DebugDataDump FUNC_NAME, "[RECV ] ", sBuffer
        sText = sText & sBuffer
        '--- restart timeout timer
        If LenB(sBuffer) <> 0 Then
            dblTimerEx = TimerEx
        End If
RetryParse:
        If LenB(sText) <> 0 Then
            Select Case Asc(sText)
            Case ZFP_STX            ' &H2
                '--- check if response complete
                If Len(sText) > 6 And Asc(Right$(sText, 1)) = ZFP_ETX Then      ' &HA
                    lSize = Asc(Mid$(sText, 2, 1)) - ZFP_BASEDATA - 3
                    '--- check message length
                    If Len(sText) - 7 <> lSize Then
                        Error = Printf(pvInternal(ucsErrInvalidResponseLength), "0x" & ToHexDump(sText))
                        Exit Do
                    End If
                    '--- check sum
                    If pvCheckSum(Mid$(sText, 1, Len(sText) - 3)) <> Mid$(sText, lSize + 5, 2) Then
                        Error = Printf(pvInternal(ucsErrInvalidResponseChecksum), "0x" & ToHexDump(sText))
                        Exit Do
                    End If
                    pvTransmitCommand = Mid$(sText, 5, lSize)
                    If bLogging Then
                        DebugLog FUNC_NAME, "RetVal=" & pvTransmitCommand
                    End If
                    Exit Do
                End If
            Case ZFP_ACK            ' &H6
                If Len(sText) > 6 And Asc(Right$(sText, 1)) = ZFP_ETX Then      ' &HA
                    lSize = Len(sText) - 5
                    '--- parse respone
                    sText = Mid$(sText, 3, lSize)
                    '--- need at least 2 bytes
                    If Len(sText) < 2 Then
                        Error = Printf(pvInternal(ucsErrInvalidResponseLength), "0x" & ToHexDump(sText))
                        Exit Do
                    End If
                    m_eError = Asc(Mid$(sText, 1)) - &H30
                    m_eStatus = Asc(Mid$(sText, 2)) - &H30
                    If m_eError <> 0 Or m_eStatus <> 0 Then
                        DebugLog FUNC_NAME, "eCmd=0x" & Hex$(eCmd) & ", StatusText=" & StatusText & ", ErrorText=" & ErrorText & ", Received=0x" & ToHexDump(sText)
                        Error = IIf(LenB(StatusText) <> 0, StatusText & ". ", vbNullString) & ErrorText
                    Else
                        Error = vbNullString
                    End If
                    Exit Do
                End If
            Case ZFP_RETRY          ' &HE
                DebugLog FUNC_NAME, "eCmd=0x" & Hex$(eCmd) & ", Retry requested, sText=" & sText
                GoTo RetrySend
            Case ZFP_NACK           ' &H15
                Error = Printf(pvInternal(ucsErrInvalidMessageFormatOrChecksum), "0x" & ToHexDump(sText))
                Exit Do
            Case Else
                If eCmd <= [_ucsZekCmdFastMax] Then
                    pvTransmitCommand = Right$(sText, 1)
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

Private Function pvPrintLine(vText As Variant) As Boolean
    Const FUNC_NAME     As String = "pvPrintLine"
    Const CONF_NAME     As String = "PrintLine"
    Dim vElem           As Variant
    Dim eCmd            As UcsTremolCommandsEnum
    Dim sParam          As String
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "PrintText", eCmd, ucsZekCmdPrintText, sParam, "%1"
    If IsArray(vText) Then
        For Each vElem In vText
            SendCommand eCmd, Printf(sParam, Zn(C_Str(vElem), " "))
            If LenB(m_uPrintData.LastError) <> 0 Then
                GoTo QH
            End If
        Next
    Else
        SendCommand eCmd, Printf(sParam, Zn(C_Str(vText), " "))
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
    Dim sData           As String
    Dim dblTotal        As Double
    Dim dblDisc         As Double
    Dim vSplit          As Variant
    Dim lIdx            As Long
    Dim eCmd            As UcsTremolCommandsEnum
    Dim sParam          As String
    Dim lIndex          As Long
    Dim lPassIndex      As Long
    Dim sFormatPrice    As String
    Dim sFormatQty      As String
    Dim sFormatDisc     As String
    Dim sFormatAmount   As String
    Dim sFormatZero     As String
    Dim lPrintType      As Long
    
    On Error GoTo EH
    '--- ring sale
    Do While uCtx.Row < lRowCount
        With uRow(uCtx.Row)
        Select Case .RowType
        Case ucsRowInit
            '--- note: cannot print barcodes in fiscal receipt, so print all barcodes in advance
            For lIdx = uCtx.Row + 1 To lRowCount - 1
                With uRow(lIdx)
                If .RowType = ucsRowBarcode Then
                    If Not PrintBarcode(.BarcodeType, .BarcodeText, .BarcodeHeight) Then
                        GoTo QH
                    End If
                End If
                End With
            Next
            pvGetCommandConfig CONF_NAME, "InfoParameters", eCmd, ucsZekCmdInfoParameters, sParam, vbNullString, _
                "Index", lIndex, 9
            vSplit = Split(SendCommand(eCmd, sParam), DEF_SEP)
            If C_Bool(At(vSplit, lIndex)) Then
                '--- use single operator
                .InitOperatorCode = DEF_OPERNO
                .InitOperatorPassword = GetDefaultPassword(DEF_OPERNO)
            End If
            pvGetCommandConfig CONF_NAME, "InfoOperator", eCmd, ucsZekCmdInfoOperator, sParam, "%1", _
                "Index", lIndex, 1, _
                "PassIndex", lPassIndex, 2
            sParam = Printf(sParam, Pad(.InitOperatorCode, -m_uConfig.MaxOperCode))
            vSplit = Split(SendCommand(eCmd, sParam), DEF_SEP)
            If LenB(m_uPrintData.LastError) <> 0 Then
                GoTo QH
            End If
            '--- init operator name (if needed)
            If LenB(.InitOperatorName) = 0 Then
                .InitOperatorName = Trim$(At(vSplit, lIndex))
            End If
            If Trim$(At(vSplit, lIndex)) <> Left$(Trim$(.InitOperatorName), m_uConfig.MaxOperName) Or LenB(Trim$(At(vSplit, lPassIndex))) = 0 Then
                pvGetCommandConfig CONF_NAME, "InitOperator", eCmd, ucsZekCmdInitOperator, sParam, "%1" & DEF_SEP & "%2" & DEF_SEP & "%3"
                sParam = Printf(sParam, Pad(.InitOperatorCode, -m_uConfig.MaxOperCode), _
                    Pad(.InitOperatorName, m_uConfig.MaxOperName), _
                    Pad(.InitOperatorPassword, m_uConfig.MaxOperPassword))
                SendCommand eCmd, sParam
                '--- can return "Command denied because of uncommited report"
                If LenB(Trim$(At(vSplit, lPassIndex))) = 0 And LenB(m_uPrintData.LastError) <> 0 Then
                    pvSetLastError m_uPrintData.LastError & pvInternal(ucsErrPasswordNotSet), ucsFerPasswordNotSet
                    GoTo QH
                End If
            End If
            '--- setup invoice no
            If LenB(.InitInvData(ucsInvDocNo)) <> 0 Then
                Select Case .InitReceiptType
                Case ucsFscRcpInvoice, ucsFscRcpCreditNote
                    pvGetCommandConfig CONF_NAME, "InitInvoiceNo", eCmd, ucsZekCmdInitInvoiceNo, sParam, "%1" & DEF_SEP & "%1"
                    sParam = Printf(sParam, Pad(.InitInvData(ucsInvDocNo), -m_uConfig.MaxInvoiceNo))
                    SendCommand eCmd, sParam
                    If LenB(m_uPrintData.LastError) <> 0 Then
                        GoTo QH
                    End If
                End Select
            End If
            '--- setup fiscal receipt
            Select Case .InitReceiptType
            Case ucsFscRcpSale
                sParam = "%1" & DEF_SEP & "%2" & DEF_SEP & "%3" & DEF_SEP & "%4" & DEF_SEP & "%5%7"
                lPrintType = 0
            Case ucsFscRcpInvoice
                sParam = "%1" & DEF_SEP & "%2" & DEF_SEP & "%3" & DEF_SEP & "%4" & DEF_SEP & "%5" & DEF_SEP & "%6%7"
                lPrintType = 1
            Case ucsFscRcpReversal
                sParam = "%1" & DEF_SEP & "%2" & DEF_SEP & "%3" & DEF_SEP & "%4" & DEF_SEP & "%5" & DEF_SEP & "%8"
                lPrintType = 64
            Case ucsFscRcpCreditNote
                sParam = "%1" & DEF_SEP & "%2" & DEF_SEP & "%3" & DEF_SEP & "%4" & DEF_SEP & "%5" & DEF_SEP & "%6" & DEF_SEP & "%9"
                lPrintType = 65
            End Select
            pvGetCommandConfig CONF_NAME, "FiscalOpen" & .InitReceiptType, eCmd, ucsZekCmdFiscalOpen, sParam, sParam, _
                "PrintType", lPrintType, lPrintType
            sParam = Printf(sParam, Pad(.InitOperatorCode, -m_uConfig.MaxOperCode), _
                Pad(.InitOperatorPassword, m_uConfig.MaxOperPassword), _
                IIf(m_uConfig.DetailedReceipt, 1, 0), _
                IIf(m_uConfig.PrintVat, 1, 0), _
                Chr$(IIf(m_uConfig.DelayPrint, 2, 0) Or IIf(m_uConfig.BufferPrint, 4, 0) Or lPrintType), _
                Pad(.InitInvData(ucsInvCgName), m_uConfig.MaxCompany) & _
                    DEF_SEP & Pad(.InitInvData(ucsInvCgPrsReceive), m_uConfig.MaxPrsReceive) & _
                    DEF_SEP & Pad(.InitInvData(ucsInvCgVatNo), m_uConfig.MaxVatNo) & _
                    DEF_SEP & Pad(.InitInvData(ucsInvCgTaxNo), m_uConfig.MaxTaxNo) & _
                    DEF_SEP & Pad(.InitInvData(ucsInvCgCity), m_uConfig.MaxCity) & _
                    DEF_SEP & .InitInvData(ucsInvCgTaxNoType), _
                IIf(LenB(.InitUniqueSaleNo) <> 0, "$" & Left$(.InitUniqueSaleNo, m_uConfig.MaxUniqueSaleNo), vbNullString), _
                .InitRevData(ucsRevType) & _
                    DEF_SEP & .InitRevData(ucsRevReceiptNo) & _
                    DEF_SEP & Format$(.InitRevData(ucsRevReceiptDate), FORMAT_DATE_EXTENDED) & _
                    DEF_SEP & Pad(.InitRevData(ucsRevFiscalMemoryNo), m_uConfig.MaxFiscalMemoryNo) & _
                    IIf(LenB(.InitUniqueSaleNo) <> 0, DEF_SEP & Left$(.InitUniqueSaleNo, m_uConfig.MaxUniqueSaleNo), vbNullString), _
                .InitRevData(ucsRevType) & _
                    DEF_SEP & Pad(.InitRevData(ucsRevInvoiceNo), -m_uConfig.MaxInvoiceNo) & _
                    DEF_SEP & Format$(.InitRevData(ucsRevReceiptDate), FORMAT_DATE_EXTENDED) & _
                    DEF_SEP & .InitRevData(ucsRevReceiptNo) & _
                    DEF_SEP & Pad(.InitRevData(ucsRevFiscalMemoryNo), m_uConfig.MaxFiscalMemoryNo) & _
                    IIf(LenB(.InitUniqueSaleNo) <> 0, DEF_SEP & Left$(.InitUniqueSaleNo, m_uConfig.MaxUniqueSaleNo), vbNullString))
            SendCommand eCmd, sParam
            If LenB(m_uPrintData.LastError) <> 0 Then
                GoTo QH
            End If
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
            If .PluDepartmentNo <> 0 Then
                pvGetCommandConfig CONF_NAME, "FiscalDepartmentSell" & .InitReceiptType, eCmd, ucsZekCmdFiscalDepartmentSell, sParam, "%1", _
                    "FormatPrice", sFormatPrice, FORMAT_BASE_2, _
                    "FormatQty", sFormatQty, FORMAT_BASE_3, _
                    "FormatDisc", sFormatDisc, FORMAT_BASE_2
            Else
                pvGetCommandConfig CONF_NAME, "FiscalSell" & .InitReceiptType, eCmd, ucsZekCmdFiscalSell, sParam, "%1", _
                    "FormatPrice", sFormatPrice, FORMAT_BASE_2, _
                    "FormatQty", sFormatQty, FORMAT_BASE_3, _
                    "FormatDisc", sFormatDisc, FORMAT_BASE_2
            End If
            vSplit = WrapText(.PluItemName & IIf(LenB(.PluUnitOfMeasure) <> 0, " (" & .PluUnitOfMeasure & ")", vbNullString), _
                LimitLong(pvCommentChars, , pvItemChars))
            For lIdx = 0 To UBound(vSplit) - 1
                If Not pvPrintLine(C_Str(vSplit(lIdx))) Then
                    GoTo QH
                End If
            Next
            sData = Pad(At(vSplit, UBound(vSplit)), m_uConfig.MaxPluName)
            If .PluDepartmentNo <> 0 Then
                sData = sData & DEF_SEP & Chr$(&H80 + .PluDepartmentNo) & DEF_SEP
            Else
                sData = sData & DEF_SEP & Chr$(191 + .PluTaxGroup) & DEF_SEP
            End If
            If Abs(.PluQuantity) > DBL_EPSILON Then
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
                    sData = sData & ":" & SafeFormat(.DiscValue, sFormatDisc)
                End Select
            Else
                sData = sData & SafeFormat(0, sFormatPrice)
            End If
            sParam = Printf(sParam, sData)
            SendCommand eCmd, sParam
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
                If .DiscType = ucsFscDscSubtotal Then
                    pvGetCommandConfig CONF_NAME, "FiscalSubtotal" & .DiscType, eCmd, ucsZekCmdFiscalSubtotal, sParam, "1" & DEF_SEP & "0,%1", _
                        "FormatDisc", sFormatDisc, FORMAT_BASE_2
                    sParam = Printf(sParam, SafeFormat(.DiscValue, sFormatDisc))
                    SendCommand eCmd, sParam
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
            '--- do nothing
        Case ucsRowPayment
            pvGetCommandConfig CONF_NAME, "InfoTransaction", eCmd, ucsZekCmdInfoTransaction, sParam, vbNullString, _
                "Index", lIndex, 0
            vSplit = Split(SendCommand(eCmd, sParam), DEF_SEP)
            If LenB(m_uPrintData.LastError) <> 0 Then
                GoTo QH
            End If
            If .PmtType <> 0 Then
                If pvToDevicePmtIndex(.PmtType) < 0 Then
                    pvSetLastError Printf(pvInternal(ucsErrUnsupportedPaymentType), .PmtType)
                    GoTo QH
                End If
                pvGetCommandConfig CONF_NAME, "FiscalPayment" & .PmtType, eCmd, ucsZekCmdFiscalPayment, sParam, pvToDevicePmtIndex(.PmtType) & DEF_SEP & "0" & DEF_SEP & "%1", _
                    "FormatAmount", sFormatAmount, FORMAT_BASE_2, _
                    "FormatZero", sFormatZero, "\"""
                sParam = Printf(sParam, SafeFormat(.PmtAmount, IIf(Abs(.PmtAmount) < DBL_EPSILON, sFormatZero, sFormatAmount)))
            ElseIf C_Lng(At(vSplit, lIndex)) <> 0 Then
                pvGetCommandConfig CONF_NAME, "FiscalPayAndClose", eCmd, ucsZekCmdFiscalPayAndClose, sParam, vbNullString
            End If
            SendCommand eCmd, sParam
            If LenB(m_uPrintData.LastError) <> 0 Then
                GoTo QH
            End If
            uCtx.PmtPrinted = True
            If .PmtType = 0 And m_uConfig.DelayPrint Then
                If Not WaitDevice(lRowCount * m_lTimeout) Then
                    GoTo QH
                End If
            End If
        End Select
        End With
        If LenB(m_uPrintData.LastError) <> 0 Then
            GoTo QH
        End If
        uCtx.Row = uCtx.Row + 1
    Loop
    '--- success
    pvPrintReceipt = True
QH:
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
    Dim vSplit          As Variant
    Dim eCmd            As UcsTremolCommandsEnum
    Dim sParam          As String
    Dim lSerialIndex    As Long
    Dim lModelIndex     As Long
    Dim lFirmwareIndex  As Long
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoDiagnostics", eCmd, ucsZekCmdInfoDiagnostics, sParam, vbNullString, _
        "SerialIndex", lSerialIndex, 0
    If Not IsArray(vResult) Then
        vResult = Split(pvTransmitCommand(oConnector, lTimeout, eCmd, sParam, Error), DEF_SEP)
    End If
    If LenB(At(vResult, lSerialIndex)) <> 0 Then
        sProtocol = STR_PROTOCOL_TREMOL
        pvGetCommandConfig CONF_NAME, "InfoVersion", eCmd, ucsZekCmdInfoVersion, sParam, vbNullString, _
            "ModelIndex", lModelIndex, 3, _
            "FirmwareIndex", lFirmwareIndex, 4
        vSplit = Split(pvTransmitCommand(oConnector, lTimeout, eCmd, sParam), DEF_SEP)
        sModel = RTrim$(At(vSplit, lModelIndex))
        sFirmware = RTrim$(At(vSplit, lFirmwareIndex))
        '--- success
        pvGetDeviceInfo = True
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function pvToDevicePmtIndex(ByVal PmtType As UcsFiscalPaymentTypeEnum) As Long
    Const FUNC_NAME     As String = "pvToDevicePmtIndex"
    
    On Error GoTo EH
    pvToDevicePmtIndex = -1             ' unsupported
    If m_uConfig.ExtendedPmtTypes Then
        Select Case PmtType
        Case ucsFscPmtCash
            pvToDevicePmtIndex = 0      ' 0=SCash (Лева)
        Case ucsFscPmtCheque
            pvToDevicePmtIndex = 1      ' 1=SChecks (Чек)
        Case ucsFscPmtCoupon
            pvToDevicePmtIndex = 2      ' 2=ST (Талон)
        Case ucsFscPmtVoucher
            pvToDevicePmtIndex = 3      ' 3=SOT (В.Талон)
        Case ucsFscPmtPackaging
            pvToDevicePmtIndex = 4      ' 4=SP (Амбалаж)
        Case ucsFscPmtMaintenance
            pvToDevicePmtIndex = 5      ' 5=SSelf (Обслужване)
        Case ucsFscPmtDamage
            pvToDevicePmtIndex = 6      ' 6=SDmg (Повреди)
        Case ucsFscPmtCard
            pvToDevicePmtIndex = 7      ' 7=SCards (Карта)
        Case ucsFscPmtBank
            pvToDevicePmtIndex = 8      ' 8=SW (Банка)
        Case ucsFscPmtCustom1
            pvToDevicePmtIndex = 9      ' 9=SR1 (Резерв 1)
        Case ucsFscPmtCustom2
            pvToDevicePmtIndex = 10     ' 10=SR2 (Резерв 2)
        Case ucsFscPmtEUR
            pvToDevicePmtIndex = 11     ' 0=SCash (Евро)
        End Select
    Else
        Select Case PmtType
        Case ucsFscPmtCash
            pvToDevicePmtIndex = 0      ' 0=SCash (Лева)
        Case ucsFscPmtCard
            pvToDevicePmtIndex = 1      ' 7=SCards (Карта)
        Case ucsFscPmtCheque
            pvToDevicePmtIndex = 2      ' 1=SChecks (Чек)
        Case ucsFscPmtCoupon
            pvToDevicePmtIndex = 3      ' 2=ST (Талон)
        Case ucsFscPmtEUR
            pvToDevicePmtIndex = 4      ' 0=SCash (Евро)
        End Select
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Function pvToDevicePmtName(ByVal PmtType As UcsFiscalPaymentTypeEnum) As String
    Const FUNC_NAME     As String = "pvToDevicePmtName"
    Dim lPmtIndex       As Long
    
    On Error GoTo EH
    lPmtIndex = pvToDevicePmtIndex(PmtType)
    If lPmtIndex >= 0 Then
        pvToDevicePmtName = RTrim$(At(Split(m_sPaymentTypes, DEF_SEP), lPmtIndex))
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Sub pvGetCommandConfig( _
            sFunction As String, _
            sKey As String, _
            eCmd As UcsTremolCommandsEnum, _
            ByVal eCmdDefault As UcsTremolCommandsEnum, _
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

Private Function pvZfplibValue(lValue As Long, sRegValue As String, ByVal lDefValue As Long) As Long
    Const FUNC_NAME     As String = "pvZfplibValue"
    
    On Error GoTo EH
    If lValue = 0 Then
        lValue = C_Lng(RegReadString(HKEY_LOCAL_MACHINE, "Software\Tremol\ZFPLib", sRegValue))
        If lValue = 0 Then
            lValue = lDefValue
        End If
    End If
    pvZfplibValue = lValue
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

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

Private Function pvCheckSum(sText As String) As String
    Const FUNC_NAME     As String = "pvCheckSum"
    Dim baText()        As Byte
    Dim lIdx            As Long
    Dim lSum            As Long
    
    On Error GoTo EH
    baText = ToAscii(sText, m_uConfig.CodePage)
    For lIdx = 1 To Len(sText) - 1
        lSum = lSum Xor baText(lIdx)
    Next
    pvCheckSum = Chr$((lSum \ &H10 Mod &H10) + &H30) & Chr$((lSum \ &H1 Mod &H10) + &H30)
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Initialize()
    m_lTimeout = DEF_TIMEOUT
    m_lSequence = DEF_SEQUENCE
    LocalizedText(ucsFscLciInternalErrors) = STR_INTERNAL
    LocalizedText(ucsFscLciPrinterStatuses) = STR_STATUSES
    LocalizedText(ucsFscLciPrinterErrors) = STR_ERRORS
    LocalizedText(ucsFscLciReceiptTexts) = STR_RECEIPT_TEXTS
    Set m_uConfig.ConfigCommands = GetConfigCollection(STR_PROTOCOL_TREMOL, "Commands")
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
    IDeviceProtocol_PrintServiceDeposit = PrintServiceDeposit(OperatorCode, OperatorPassword, ucsFscPmtCash, Amount)
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
    
End Function

Private Function IDeviceProtocol_GetOperatorName(OperatorCode As String) As String
    IDeviceProtocol_GetOperatorName = GetOperatorName(OperatorCode)
End Function

Private Function IDeviceProtocol_GetDefaultPassword(OperatorCode As String) As String
    IDeviceProtocol_GetDefaultPassword = GetDefaultPassword(OperatorCode)
End Function

Private Function IDeviceProtocol_GetDeviceProtocol() As String
    IDeviceProtocol_GetDeviceProtocol = GetDeviceProtocol()
End Function

Private Function IDeviceProtocol_GetDeviceModel() As String
    IDeviceProtocol_GetDeviceModel = GetDeviceModel()
End Function

Private Function IDeviceProtocol_GetDeviceFirmware() As String
    IDeviceProtocol_GetDeviceFirmware = GetDeviceFirmware()
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
    LocalizedText(Index) = Text
End Sub

Private Function IDeviceProtocol_StartReceipt(ByVal ReceiptType As UcsFiscalReceiptTypeEnum, OperatorCode As String, OperatorName As String, OperatorPassword As String, Optional UniqueSaleNo As String, Optional TableNo As String, Optional ByVal DisablePrinting As Boolean, Optional InvDocNo As String, Optional InvCgTaxNo As String, Optional ByVal InvCgTaxNoType As UcsFiscalTaxNoTypeEnum, Optional InvCgVatNo As String, Optional InvCgName As String, Optional InvCgCity As String, Optional InvCgAddress As String, Optional InvCgPrsReceive As String, Optional OwnData As String, Optional ByVal RevType As UcsFiscalReversalTypeEnum, Optional RevReceiptNo As String, Optional ByVal RevReceiptDate As Date, Optional RevFiscalMemoryNo As String, Optional RevInvoiceNo As String, Optional RevReason As String) As Boolean
    Dim sOperCode       As String
    Dim sOperPass       As String
    
    m_uPrintData.Config.RowChars = pvRowChars
    m_uPrintData.Config.CommentChars = pvCommentChars
    m_uPrintData.Config.ItemChars = pvItemChars
    sOperCode = Zn(OperatorCode, DEF_OPERNO)
    sOperPass = Zn(OperatorPassword, GetDefaultPassword(sOperCode))
    IDeviceProtocol_StartReceipt = PpdStartReceipt(m_uPrintData, ReceiptType, sOperCode, OperatorName, sOperPass, TableNo, UniqueSaleNo, DisablePrinting, InvDocNo, InvCgTaxNo, InvCgTaxNoType, InvCgVatNo, InvCgName, InvCgCity, InvCgAddress, InvCgPrsReceive, RevType, RevReceiptNo, RevReceiptDate, RevFiscalMemoryNo, RevInvoiceNo, RevReason)
End Function

Private Function IDeviceProtocol_GetDepartmentName(DepartmentCode As String, Optional TaxGroup As Long) As String
    IDeviceProtocol_GetDepartmentName = GetDepartmentName(DepartmentCode, TaxGroup)
End Function

Private Function IDeviceProtocol_GetTaxRates() As Variant
    IDeviceProtocol_GetTaxRates = GetTaxRates()
End Function
