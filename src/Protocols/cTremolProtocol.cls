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
' Protocol handler for Tremol/Zeka cash registers & FP
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

'--- error codes
Private Const ERROR_IO_PENDING              As Long = 997
'--- for PurgeComm
Private Const PURGE_TXCLEAR                 As Long = &H4     '  Kill the transmit queue if there.
Private Const PURGE_RXCLEAR                 As Long = &H8     '  Kill the typeahead buffer if there.

Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function CreateEvent Lib "kernel32" Alias "CreateEventA" (ByVal lpEventAttributes As Long, ByVal bManualReset As Long, ByVal bInitialState As Long, ByVal lpName As Long) As Long
Private Declare Function WaitForMultipleObjects Lib "kernel32" (ByVal nCount As Long, lpHandles As Long, ByVal bWaitAll As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function PurgeComm Lib "kernel32" (ByVal hFile As Long, ByVal dwFlags As Long) As Long

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
Private Const DEF_TIMEOUT               As Long = 3000
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
Private Const DEF_MAX_PRSRECEIVE        As Long = 16
Private Const DEF_MAX_VATNO             As Long = 13
Private Const DEF_MAX_TAXNO             As Long = 13
Private Const DEF_MAX_CITY              As Long = 30
Private Const DEF_MAX_INVOICENO         As Long = 10
Private Const DEF_MAX_PLUNAME           As Long = 36
Private Const DEF_MAX_COMMENT           As Long = 34
Private Const DEF_MAX_FISCALMEMORYNO    As Long = 8
Private Const DEF_MAX_UNIQUESALENO      As Long = 24
Private Const DEF_DELIM                 As String = ";"
Private Const LNG_STATUS_READY          As Long = &H40
Private Const LNG_STATUS_IS_BUSY        As Long = &H41
Private Const LNG_STATUS_OUT_OF_PAPER   As Long = &H42
Private Const LNG_STATUS_OUT_OF_PAPER_LEGACY As Long = &H7
Private Const LNG_STATUS_OUT_OF_PAPER_AND_IS_BUSY As Long = &H43
Private Const LNG_STATUS_OVERHEATED     As Long = &H44
Private Const ucsFscRetNonfiscal        As Long = ucsFscRetSale + 100
Private Const FORMAT_DATE_EXTENDED      As String = "dd\-mm\-yy hh\:nn\:ss"

Private m_sDevice                   As String
Private m_oConnector                As IDeviceConnector
Private m_lTimeout                  As Long
Private m_lSequence                 As Long
Private m_eLastErrNo                As UcsFiscalErrorsEnum
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
    [__ucsErrWriteFileFailure]
    [__ucsErrTimeoutWaitingForResponse]
    [__ucsErrReadFileFailed]
    ucsErrInvalidMessageFormatOrChecksum
    [__ucsErrAlreadyConnected]
    [__ucsErrNotConnected]
    ucsErrInvalidResponseLength
    ucsErrResponseMissingPreamble
    [__ucsErrWaitCommEventFailed]
    ucsErrNoReceiptStarted
    ucsErrErrorInLine
    [__ucsErrCannotCancelReceipt]
    ucsErrInvalidReportType
    ucsErrDeviceNotOnline
    ucsErrDeviceNotReady
    ucsErrPasswordNotSet
    [__ucsErrInvalidPassword]
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
    ucsTxtPaymentCard
    ucsTxtPaymentCheque
    ucsTxtSurcharge
    ucsTxtDiscount
    [__ucsTxtTotal]
    ucsTxtSubTotal
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
    ucsTxtBulstatCaption
    [__ucsTxtPluVoid]
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
    ucsStoReason
    ucsStoReceiptNo
    ucsStoReceiptDate
    ucsStoFiscalMemoryNo
    ucsStoInvoiceNo
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
    Connector           As cSerialPortConnector
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
    MaxPrsReceive       As Long
    MaxVatNo        As Long
    MaxTaxNo        As Long
    MaxCity             As Long
    MaxInvoiceNo        As Long
    MaxPluName          As Long
    MaxComment          As Long
    MaxFiscalMemoryNo   As Long
    MaxUniqueSaleNo     As Long
    ExtendedPmtTypes    As Boolean
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
    LastErrNo = m_eLastErrNo
End Property

Property Get LastError() As String
    LastError = m_sLastError
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
        vInfoVersion = Split(pvSendCommand(m_oConnector, lTimeout, eCmd, sParam), DEF_DELIM)
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

Private Property Get DEF_EXTENDED_PMTTYPES() As Boolean
    DEF_EXTENDED_PMTTYPES = IsNumeric(At(Split(m_sPaymentTypes, DEF_DELIM), 12))
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
    Dim vSpeeds         As Variant
    
    On Error GoTo EH
    sProbeText = GetConfigValue(STR_PROTOCOL_TREMOL_FP, "ProbeText", Chr$(ucsZekCmdFastIsReady))
    sReplyText = GetConfigValue(STR_PROTOCOL_TREMOL_FP, "ReplyText", Chr$(ucsZekCmdFastIsReady))
    lTimeout = GetConfigNumber(STR_PROTOCOL_TREMOL_FP, "AutodetectTimeout", DEF_AUTODETECTTIMEOUT)
    lDeviceInfoTimeout = GetConfigNumber(STR_PROTOCOL_TREMOL_FP, "DeviceInfoTimeout", DEF_DEVICEINFOTIMEOUT)
    AssignVariant vSpeeds, GetConfigValue(STR_PROTOCOL_TREMOL_FP, "AutodetectSpeeds", Split(DEF_AUTODETECTSPEEDS, "|"))
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
                    lResult = WriteFile(.Connector.hComm, ByVal sProbeText, Len(sProbeText), lActual, .WriteOver)
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
                    lResult = ReadFile(.Connector.hComm, ByVal StrPtr(.Buffer), LenB(.Buffer), lActual, .ReadOver)
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
                If Not .Connector Is Nothing Then
                    If InStr(.Buffer, vbNullChar) > 0 Then
                        .Buffer = StrConv(.Buffer, vbUnicode)
                        .Buffer = Left$(.Buffer, InStr(.Buffer, vbNullChar) - 1)
                    End If
                    If LenB(.Buffer) <> 0 Then
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
    Dim eCmd            As UcsTremolCommandsEnum
    Dim sParam          As String
    Dim lSerialIndex    As Long
    
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
    If Not WaitDevice(-1) Then
        GoTo QH
    End If
    If Not WaitDevice(-1, CheckFastStatus:=True) Then
        Init = True
        GoTo QH
    End If
    pvGetCommandConfig CONF_NAME, "InfoPaymentTypes", eCmd, ucsZekCmdInfoPaymentTypes, sParam, vbNullString
    m_sPaymentTypes = SendCommand(eCmd, sParam)
    If LenB(m_sLastError) <> 0 Then
        GoTo QH
    End If
    pvGetCommandConfig CONF_NAME, "InfoDiagnostics", eCmd, ucsZekCmdInfoDiagnostics, sParam, vbNullString, _
        "SerialIndex", lSerialIndex, 0
    m_vInfoDiagnostics = Split(pvSendCommand(m_oConnector, m_lTimeout, eCmd, sParam, m_sLastError), DEF_DELIM)
    If LenB(m_sLastError) <> 0 Then
        GoTo QH
    End If
    '--- load config values
    sSerial = At(m_vInfoDiagnostics, lSerialIndex)
    With m_uConfig
        .IsEcr = vbUseDefault
        .RowChars = GetConfigNumber(sSerial, "RowChars", RowChars)
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
        .MaxPrsReceive = GetConfigNumber(sSerial, "MaxPrsReceive", DEF_MAX_PRSRECEIVE)
        .MaxVatNo = GetConfigNumber(sSerial, "MaxVatNo", DEF_MAX_VATNO)
        .MaxTaxNo = GetConfigNumber(sSerial, "MaxTaxNo", DEF_MAX_TAXNO)
        .MaxCity = GetConfigNumber(sSerial, "MaxCity", DEF_MAX_CITY)
        .MaxInvoiceNo = GetConfigNumber(sSerial, "MaxInvoiceNo", DEF_MAX_INVOICENO)
        .MaxPluName = GetConfigNumber(sSerial, "MaxPluName", DEF_MAX_PLUNAME)
        .MaxComment = GetConfigNumber(sSerial, "MaxComment", DEF_MAX_COMMENT)
        .MaxFiscalMemoryNo = GetConfigNumber(sSerial, "MaxFiscalMemoryNo", DEF_MAX_FISCALMEMORYNO)
        .MaxUniqueSaleNo = GetConfigNumber(sSerial, "MaxUniqueSaleNo", DEF_MAX_UNIQUESALENO)
        .ExtendedPmtTypes = GetConfigNumber(sSerial, "ExtendedPmtTypes", DEF_EXTENDED_PMTTYPES)
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

Public Function SendCommand(ByVal eCmd As UcsTremolCommandsEnum, Optional Data As String) As String
    Const FUNC_NAME     As String = "SendCommand"
    
    On Error GoTo EH
    pvSetLastError vbNullString
    If eCmd = ucsZekCmdNoop Then
        Exit Function
    End If
    If WaitDevice() Then
        SendCommand = pvSendCommand(m_oConnector, m_lTimeout, eCmd, Data, m_sLastError)
    End If
    If LenB(m_sLastError) <> 0 Then
        m_eLastErrNo = IIf(m_eError = ucsErrBadPassword, ucsFerInvalidPassword, ucsFerGeneralError)
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
        If pvSendCommand(m_oConnector, pvPingTimeout, eCmd, sParam, sError) <> sStatus Then
            pvSetLastError pvInternal(ucsErrDeviceNotOnline)
            GoTo QH
        End If
        If CheckFastStatus Then
            pvGetCommandConfig CONF_NAME, "FastStatus", eCmd, ucsZekCmdFastStatus, sParam, vbNullString, _
                "ReadyStatus", sStatus, LNG_STATUS_READY
            sCurrent = pvSendCommand(m_oConnector, pvPingTimeout, eCmd, sParam, sError)
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
            If pvSendCommand(m_oConnector, pvPingTimeout, eCmd, sParam, sError) = sStatus Then
                lIter = lIter + 1
                pvGetCommandConfig CONF_NAME, "FastStatus", eCmd, ucsZekCmdFastStatus, sParam, vbNullString, _
                    "ReadyStatus", sStatus, LNG_STATUS_READY, _
                    "Delay", sDelay, 40
                sCurrent = pvSendCommand(m_oConnector, pvPingTimeout, eCmd, sParam, sError)
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
        m_sLastError = pvInternal(ucsErrDeviceNotReady)
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
            Optional InvCgVatNo As String, _
            Optional InvCgName As String, _
            Optional InvCgCity As String, _
            Optional InvCgAddress As String, _
            Optional InvCgPrsReceive As String, _
            Optional ByVal StornoReason As UcsFiscalStornoReasonEnum, _
            Optional StornoReceiptNo As String, _
            Optional StornoReceiptDate As Date, _
            Optional StornoFiscalMemoryNo As String, _
            Optional StornoInvoiceNo As String) As Boolean
    Const FUNC_NAME     As String = "StartReceipt"
    Const InvCgTaxNoType As Long = 0
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
        SplitCgAddress Trim$(SafeText(InvCgCity)) & vbCrLf & Trim$(SafeText(InvCgAddress)), sCity, sAddress, 30
        .InitInvData = Array(SafeText(InvDocNo), SafeText(InvCgTaxNo), SafeText(InvCgVatNo), _
            SafeText(InvCgName), sCity, sAddress, SafeText(InvCgPrsReceive), InvCgTaxNoType)
        .InitStornoData = Array(Switch(.InitReceiptType = ucsFscRetCreditNote, 2, StornoReason = ucsFscRsnOperatorError, 0, True, 1), _
            SafeText(StornoReceiptNo), StornoReceiptDate, SafeText(StornoFiscalMemoryNo), SafeText(StornoInvoiceNo))
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
            ByVal PmtType As UcsFiscalPaymentTypeEnum, _
            Name As String, _
            ByVal Amount As Double) As Boolean
    Const FUNC_NAME     As String = "AddPayment"
    
    On Error GoTo EH
    '--- sanity check
    If m_lRowCount = 0 Then
        pvSetLastError pvInternal(ucsErrNoReceiptStarted)
        GoTo QH
    End If
    With m_uRow(pvAddRow())
        .RowType = ucsRowPayment
        .PmtType = PmtType
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
    pvConvertExtraRows
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
    '--- need resume ticket only if payment processed
    With m_uCtxPrint
        If .PmtPrinted Then
            For lIdx = LBound(.GrpTotal) To UBound(.GrpTotal)
                GetResumeTicket = GetResumeTicket & .GrpTotal(lIdx) & STR_CHR1
            Next
            GetResumeTicket = GetResumeTicket & .Paid & STR_CHR1 & .PluCount & STR_CHR1 & -.PmtPrinted & STR_CHR1 & -.ChangePrinted & STR_CHR1 & .Row
        End If
    End With
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
    If LenB(ReceiptNo) = 0 Then
        sPrevReceiptNo = GetLastReceiptNo()
        pvGetCommandConfig CONF_NAME, "PrintDuplicate", eCmd, ucsZekCmdPrintDuplicate, sParam, vbNullString, _
            "PrintTimeout", lPrintTimeout, 0
        SendCommand eCmd, sParam
        If Not LenB(m_sLastError) = 0 Then
            GoTo QH
        End If
        If Not WaitDevice(Znl(lPrintTimeout, DEF_PRINT_TIMEOUT)) Then
            GoTo QH
        End If
        sReceiptNo = GetLastReceiptNo()
        If LenB(sReceiptNo) = 0 Or sReceiptNo = sPrevReceiptNo Then
            pvSetLastError At(m_uConfig.LocalizedText(ucsFscLciPrinterStatuses), ucsErrIllegalCommand)
            GoTo QH
        End If
    Else
        pvGetCommandConfig CONF_NAME, "PrintReportDaily", eCmd, ucsZekCmdPrintReportDaily, sParam, "J1;N;%1;%1", _
            "PrintTimeout", lPrintTimeout, 0
        SendCommand eCmd, Printf(sParam, Pad(ReceiptNo, 5))
        If Not LenB(m_sLastError) = 0 Then
            GoTo QH
        End If
        If Not WaitDevice(Znl(lPrintTimeout, DEF_PRINT_TIMEOUT)) Then
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
    Dim eCmd            As UcsTremolCommandsEnum
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
    Dim eCmd            As UcsTremolCommandsEnum
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
    Dim eCmd            As UcsTremolCommandsEnum
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
    vSplit = Split(SendCommand(eCmd, sParam), DEF_DELIM)
    If LenB(m_sLastError) = 0 Then
        GetLastReceiptNo = RTrim$(At(vSplit, lIndex))
    End If
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
    vSplit = Split(SendCommand(eCmd, sParam), DEF_DELIM)
    If LenB(m_sLastError) = 0 And IsNumeric(At(vSplit, lIndex)) Then
        GetLastInvoiceNo = C_Dbl(At(vSplit, lIndex)) - 1
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
    GetTaxNo = RTrim$(At(Split(SendCommand(eCmd, sParam), DEF_DELIM), lIndex))
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
    lIndex = C_Lng(At(Split(SendCommand(eCmd, sParam), DEF_DELIM), lIndex))
    GetTaxCaption = RTrim$(At(Split(pvText(ucsTxtBulstatCaption), DEF_DELIM), lIndex))
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
    If Index >= 1 And Index <= 6 Then
        pvGetCommandConfig CONF_NAME, "InfoHeaderFooter" & Index, eCmd, ucsZekCmdInfoHeaderFooter, sParam, C_Str(Index), _
            "Index", lIndex, 1
        GetHeaderText = RTrim$(At(Split(SendCommand(eCmd, sParam), DEF_DELIM), lIndex))
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
    If Index >= 1 And Index <= 1 Then
        pvGetCommandConfig CONF_NAME, "InfoHeaderFooter" & Index, eCmd, ucsZekCmdInfoHeaderFooter, sParam, C_Str(7 + Index), _
            "Index", lIndex, 1
        GetFooterText = RTrim$(At(Split(SendCommand(eCmd, sParam), DEF_DELIM), lIndex))
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetPaymentName(ByVal PmtType As UcsFiscalPaymentTypeEnum) As String
    Const FUNC_NAME     As String = "GetPaymentName"
    Dim lPmtIndex       As Long
    
    On Error GoTo EH
    If PmtType >= ucsFscPmtCash And PmtType <= ucsFscPmtCheque Then
        GetPaymentName = pvText(ucsTxtPaymentCash + PmtType - ucsFscPmtCash)
    Else
        lPmtIndex = pvToDevicePmtIndex(PmtType)
        If lPmtIndex >= 0 Then
            GetPaymentName = RTrim$(At(Split(m_sPaymentTypes, DEF_DELIM), lPmtIndex))
        End If
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
    pvGetDeviceInfo m_oConnector, m_lTimeout, m_vInfoDiagnostics, GetDeviceProtocol, vbNullString
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

Public Function GetDeviceStatus(CurrentStatus As String) As Boolean
    Const FUNC_NAME     As String = "GetDeviceStatus"
    
    On Error GoTo EH
    WaitDevice -1, CheckFastStatus:=True
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
    Dim eCmd            As UcsTremolCommandsEnum
    Dim sParam          As String
    Dim vPayments       As Variant
    Dim vChanges        As Variant
    Dim vRetVal         As Variant
    Dim lIdx            As Long
    Dim lPmtIndex       As Long
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoRegisters4", eCmd, ucsZekCmdInfoRegisters, sParam, "4"
    vPayments = Split(SendCommand(eCmd, sParam), DEF_DELIM)
    pvGetCommandConfig CONF_NAME, "InfoRegisters6", eCmd, ucsZekCmdInfoRegisters, sParam, "6"
    vChanges = Split(SendCommand(eCmd, sParam), DEF_DELIM)
    If UBound(vPayments) > 0 Then
        ReDim vRetVal(0 To 7) As Double
        For lIdx = 0 To UBound(vRetVal)
            lPmtIndex = pvToDevicePmtIndex(lIdx + 1)
            If lPmtIndex >= 0 Then
                vRetVal(lIdx) = ParseSum(At(vPayments, lPmtIndex + 1)) - ParseSum(At(vChanges, lPmtIndex + 1))
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
    Dim vVatGroups      As Variant
    Dim vTaxRates       As Variant
    Dim vRetVal         As Variant
    Dim lIdx            As Long
    Dim eCmd            As UcsTremolCommandsEnum
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
        GetTotalsByVatGroups = Array()
    End If
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
        Optional ByVal PmtType As UcsFiscalPaymentTypeEnum, _
        Optional ByVal Value As Double, _
        Optional Comment As String) As Variant
    Const FUNC_NAME     As String = "CashDebitCredit"
    Const CONF_NAME     As String = FUNC_NAME
    Dim vSplit          As Variant
    Dim vRetVal         As Variant
    Dim eCmd            As UcsTremolCommandsEnum
    Dim sParam          As String
    Dim sFormat         As String
    Dim lIndex          As Long
    Dim lPrintTimeout   As Long
    Dim lPmtIndex       As Long
    
    On Error GoTo EH
    lPmtIndex = pvToDevicePmtIndex(PmtType)
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
            lPmtIndex, _
            SafeFormat(Value, sFormat), _
            Left$(Comment, m_uConfig.MaxComment))
        SendCommand eCmd, sParam
        If LenB(m_sLastError) <> 0 Then
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
    If LenB(m_sLastError) <> 0 Then
        vRetVal = EmptyDoubleArray
        GoTo QH
    End If
    ReDim vRetVal(0 To 2) As Double
    vRetVal(0) = C_Dbl(At(vSplit, lPmtIndex + 1))
    pvGetCommandConfig CONF_NAME, "InfoRegisters2", eCmd, ucsZekCmdInfoRegisters, sParam, "2"
    vSplit = Split(SendCommand(eCmd, sParam), DEF_DELIM)
    vRetVal(1) = C_Dbl(At(vSplit, lPmtIndex + 1))
    pvGetCommandConfig CONF_NAME, "InfoRegisters3", eCmd, ucsZekCmdInfoRegisters, sParam, "3"
    vSplit = Split(SendCommand(eCmd, sParam), DEF_DELIM)
    vRetVal(2) = C_Dbl(At(vSplit, lPmtIndex + 1))
QH:
    CashDebitCredit = vRetVal
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Sub SetLocalizedCommand( _
            sFunction As String, _
            sKey As String, _
            Optional ByVal cmd As UcsTremolCommandsEnum, _
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

'= private ===============================================================

Private Function pvSendCommand( _
            oConnector As IDeviceConnector, _
            ByVal lTimeout As Long, _
            ByVal eCmd As UcsTremolCommandsEnum, _
            Optional Data As String, _
            Optional Error As String) As String
    Const FUNC_NAME     As String = "pvSendCommand"
    Dim sText           As String
    Dim baText()        As Byte
    Dim lIdx            As Long
    Dim lSum            As Long
    Dim lSize           As Long
    Dim bLogging        As Boolean
    
    On Error GoTo EH
    bLogging = (eCmd > [_ucsZekCmdFastMax])
    If bLogging Then
        DebugLog FUNC_NAME, "eCmd=0x" & Hex$(eCmd) & IIf(LenB(Data) <> 0, ", Data=" & Data, vbNullString)
    End If
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
    '--- send request
    DebugDataDump FUNC_NAME, "send ", sText
    baText = ToAscii(sText)
    If Not oConnector.WriteData(baText, lTimeout) Then
        Error = oConnector.GetLastError()
        DebugLog FUNC_NAME, "WriteData failed, Error=" & Error
        GoTo QH
    End If
    '--- read response
    sText = vbNullString
    Do
        If Not oConnector.ReadData(baText, lTimeout) Then
            Error = oConnector.GetLastError()
            DebugLog FUNC_NAME, "ReadData failed, Error=" & Error
            GoTo QH
        End If
        DebugDataDump FUNC_NAME, "recv ", FromAscii(baText)
        sText = sText & FromAscii(baText)
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
                            DebugLog FUNC_NAME, "eCmd=0x" & Hex$(eCmd) & ", ErrorText=" & ErrorText & ", Received=0x" & ToHexDump(sText)
                            Error = ErrorText
                        ElseIf m_eStatus <> 0 Then
                            DebugLog FUNC_NAME, "eCmd=0x" & Hex$(eCmd) & ", StatusText=" & StatusText & ", Received=0x" & ToHexDump(sText)
                            Error = StatusText
                        Else
                            Error = vbNullString
                        End If
                    Else
                        Error = Printf(pvInternal(ucsErrInvalidResponseLength), "0x" & ToHexDump(sText))
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
    Dim eCmd            As UcsTremolCommandsEnum
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
    Dim eCmd            As UcsTremolCommandsEnum
    Dim sParam          As String
    Dim lIndex          As Long
    Dim lPassIndex      As Long
    Dim sFormatPrice    As String
    Dim sFormatQty      As String
    Dim sFormatDisc     As String
    Dim sFormatAmount   As String
    Dim lPrintType      As Long
    
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
            If LenB(.InitInvData(ucsInvDocNo)) <> 0 Then
                Select Case .InitReceiptType
                Case ucsFscRetInvoice, ucsFscRetCreditNote
                    pvGetCommandConfig CONF_NAME, "InitInvoiceNo", eCmd, ucsZekCmdInitInvoiceNo, sParam, "%1;%1"
                    sParam = Printf(sParam, Pad(.InitInvData(ucsInvDocNo), -m_uConfig.MaxInvoiceNo))
                    SendCommand eCmd, sParam
                    If LenB(m_sLastError) <> 0 Then
                        GoTo QH
                    End If
                End Select
            End If
            '--- setup fiscal receipt
            Select Case .InitReceiptType
            Case ucsFscRetSale
                sParam = "%1;%2;%3;0;%4%6"
                lPrintType = 0
            Case ucsFscRetInvoice
                sParam = "%1;%2;%3;0;%4;%5%6"
                lPrintType = 1
            Case ucsFscRetStorno
                sParam = "%1;%2;%3;0;%4;%7"
                lPrintType = 64
            Case ucsFscRetCreditNote
                sParam = "%1;%2;%3;0;%4;%5;%8"
                lPrintType = 65
            End Select
            pvGetCommandConfig CONF_NAME, "FiscalOpen" & .InitReceiptType, eCmd, ucsZekCmdFiscalOpen, sParam, sParam, _
                "PrintType", lPrintType, lPrintType
            sParam = Printf(sParam, Pad(.InitOperatorCode, -m_uConfig.MaxOperCode), _
                Pad(.InitOperatorPassword, m_uConfig.MaxOperPassword), _
                IIf(m_uConfig.DetailedReceipt, 1, 0), _
                Chr$(IIf(m_uConfig.DelayPrint, 2, 0) Or lPrintType), _
                Pad(.InitInvData(ucsInvCgName), m_uConfig.MaxCompany) & _
                    ";" & Pad(.InitInvData(ucsInvCgPrsReceive), m_uConfig.MaxPrsReceive) & _
                    ";" & Pad(.InitInvData(ucsInvCgVatNo), m_uConfig.MaxVatNo) & _
                    ";" & Pad(.InitInvData(ucsInvCgTaxNo), m_uConfig.MaxTaxNo) & _
                    ";" & Pad(.InitInvData(ucsInvCgCity), m_uConfig.MaxCity) & IIf(m_uConfig.ExtendedPmtTypes, _
                    ";" & .InitInvData(ucsInvCgTaxNoType), vbNullString), _
                IIf(LenB(.InitUniqueSaleNo) <> 0, "$" & Left$(.InitUniqueSaleNo, m_uConfig.MaxUniqueSaleNo), vbNullString), _
                .InitStornoData(ucsStoReason) & _
                    ";" & .InitStornoData(ucsStoReceiptNo) & _
                    ";" & Format$(.InitStornoData(ucsStoReceiptDate), FORMAT_DATE_EXTENDED) & _
                    ";" & Pad(.InitStornoData(ucsStoFiscalMemoryNo), m_uConfig.MaxFiscalMemoryNo) & _
                    IIf(LenB(.InitUniqueSaleNo) <> 0, ";" & Left$(.InitUniqueSaleNo, m_uConfig.MaxUniqueSaleNo), vbNullString), _
                .InitStornoData(ucsStoReason) & _
                    ";" & Pad(.InitStornoData(ucsStoInvoiceNo), -m_uConfig.MaxInvoiceNo) & _
                    ";" & Format$(.InitStornoData(ucsStoReceiptDate), FORMAT_DATE_EXTENDED) & _
                    ";" & .InitStornoData(ucsStoReceiptNo) & _
                    ";" & Pad(.InitStornoData(ucsStoFiscalMemoryNo), m_uConfig.MaxFiscalMemoryNo) & _
                    IIf(LenB(.InitUniqueSaleNo) <> 0, ";" & Left$(.InitUniqueSaleNo, m_uConfig.MaxUniqueSaleNo), vbNullString))
            SendCommand eCmd, sParam
            If LenB(m_sLastError) <> 0 Then
                GoTo QH
            End If
        Case ucsRowPlu
            uCtx.PluCount = uCtx.PluCount + 1
            dblTotal = Round(.PluPrice * .PluQuantity, 2)
            dblDisc = Round(dblTotal * IIf(.DiscType = ucsFscDstPlu, .DiscValue, 0) / 100#, 2)
            uCtx.GrpTotal(.PluVatGroup) = Round(uCtx.GrpTotal(.PluVatGroup) + dblTotal + dblDisc, 2)
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
        Case ucsRowDiscount
            Select Case .PrintRowType
            Case ucsFscRetNonfiscal
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
            Case Else
                If .DiscType = ucsFscDstSubtotal Then
                    pvGetCommandConfig CONF_NAME, "FiscalSubtotal" & .DiscType, eCmd, ucsZekCmdFiscalSubtotal, sParam, "1;0,%1", _
                        "FormatDisc", sFormatDisc, "0.00"
                    sParam = Printf(sParam, SafeFormat(.DiscValue, sFormatDisc))
                    SendCommand eCmd, sParam
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
            pvPrintLine vSplit
        Case ucsRowPayment
            pvGetCommandConfig CONF_NAME, "InfoTransaction", eCmd, ucsZekCmdInfoTransaction, sParam, vbNullString, _
                "Index", lIndex, 0
            vSplit = Split(SendCommand(eCmd, sParam), DEF_DELIM)
            If LenB(m_sLastError) <> 0 Then
                GoTo QH
            End If
            If .PmtType <> 0 Then
                pvGetCommandConfig CONF_NAME, "FiscalPayment" & .PmtType, eCmd, ucsZekCmdFiscalPayment, sParam, pvToDevicePmtIndex(.PmtType) & ";0;%1", _
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
        pvSetLastError Printf(pvInternal(ucsErrErrorInLine), uCtx.PluCount, m_sLastError), m_eLastErrNo
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
                        .PrintRowType = ucsFscRetNonfiscal
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

Private Function pvGetDeviceInfo(oConnector As IDeviceConnector, ByVal lTimeout As Long, vResult As Variant, sProtocol As String, sModel As String, Optional Error As String) As Boolean
    Const FUNC_NAME     As String = "pvGetDeviceInfo"
    Const CONF_NAME     As String = "GetDeviceInfo"
    Dim vSplit          As Variant
    Dim eCmd            As UcsTremolCommandsEnum
    Dim sParam          As String
    Dim lSerialIndex    As Long
    Dim lRegDateIndex   As Long
    Dim lModelIndex     As Long
    Dim sSuffix         As String
    
    On Error GoTo EH
    pvGetCommandConfig CONF_NAME, "InfoDiagnostics", eCmd, ucsZekCmdInfoDiagnostics, sParam, vbNullString, _
        "SerialIndex", lSerialIndex, 0
    If Not IsArray(vResult) Then
        vResult = Split(pvSendCommand(oConnector, lTimeout, eCmd, sParam, Error), DEF_DELIM)
    End If
    If LenB(At(vResult, lSerialIndex)) <> 0 Then
        sProtocol = STR_PROTOCOL_TREMOL_FP
        pvGetCommandConfig CONF_NAME, "InfoVersion", eCmd, ucsZekCmdInfoVersion, sParam, vbNullString, _
            "RegDateIndex", lRegDateIndex, 3, _
            "ModelIndex", lModelIndex, 4, _
            "Suffix", sSuffix, "Ê.Ñ."
        vSplit = Split(pvSendCommand(oConnector, lTimeout, eCmd, sParam), DEF_DELIM)
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

Private Function pvToDevicePmtIndex(ByVal PmtType As UcsFiscalPaymentTypeEnum) As Long
    Const FUNC_NAME     As String = "pvToDevicePmtIndex"
    
    On Error GoTo EH
    If PmtType > 4 Then
        PmtType = 4 - PmtType
    End If
    If m_uConfig.ExtendedPmtTypes Then
        Select Case PmtType
        Case ucsFscPmtCash
            pvToDevicePmtIndex = 0   ' Cash
        Case ucsFscPmtCard
            pvToDevicePmtIndex = 7   ' Card
        Case ucsFscPmtCheque
            pvToDevicePmtIndex = 8   ' Bank
        Case ucsFscPmtCustom1
            pvToDevicePmtIndex = 2   ' Talon
        Case ucsFscPmtCustom2
            pvToDevicePmtIndex = 3   ' V.Talon
        Case ucsFscPmtCustom3
            pvToDevicePmtIndex = 9   ' Prg Name1
        Case ucsFscPmtCustom4
            pvToDevicePmtIndex = 10  ' Prg Name2
        Case Else
            pvToDevicePmtIndex = -1  ' unsupported
        End Select
    Else
        Select Case PmtType
        Case ucsFscPmtCash, ucsFscPmtCard, ucsFscPmtCheque
            pvToDevicePmtIndex = PmtType - 1           ' 0, 1, 2
        Case ucsFscPmtCustom1, ucsFscPmtCustom2
            pvToDevicePmtIndex = 2 - PmtType           ' 3, 4
        Case Else
            pvToDevicePmtIndex = -1  ' unsupported
        End Select
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Sub pvGetCommandConfig( _
            sFunc As String, _
            sKey As String, _
            eCmd As UcsTremolCommandsEnum, _
            ByVal eCmdDefault As UcsTremolCommandsEnum, _
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
        m_eLastErrNo = IIf(LenB(sError) = 0, ucsFerNone, ucsFerGeneralError)
    Else
        m_eLastErrNo = ErrNum
    End If
    m_sLastError = sError
End Sub

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Initialize()
    m_lTimeout = DEF_TIMEOUT
    m_lSequence = ZFP_BASEDATA
    LocalizedText(ucsFscLciInternalErrors) = STR_INTERNAL
    LocalizedText(ucsFscLciPrinterStatuses) = STR_STATUSES
    LocalizedText(ucsFscLciPrinterErrors) = STR_ERRORS
    LocalizedText(ucsFscLciReceiptTexts) = STR_RECEIPT_TEXTS
    Set m_uConfig.ConfigCommands = GetConfigCollection(STR_PROTOCOL_TREMOL_FP, "Commands")
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

Private Function IDeviceProtocol_StartReceipt(ByVal ReceiptType As UcsFiscalReceiptTypeEnum, OperatorCode As String, OperatorName As String, OperatorPassword As String, Optional UniqueSaleNo As String, Optional TableNo As String, Optional InvDocNo As String, Optional InvCgTaxNo As String, Optional InvCgVatNo As String, Optional InvCgName As String, Optional InvCgCity As String, Optional InvCgAddress As String, Optional InvCgPrsReceive As String, Optional OwnData As String, Optional ByVal StornoReason As UcsFiscalStornoReasonEnum, Optional StornoReceiptNo As String, Optional ByVal StornoReceiptDate As Date, Optional StornoFiscalMemoryNo As String, Optional StornoInvoiceNo As String) As Boolean
    IDeviceProtocol_StartReceipt = StartReceipt(ReceiptType, OperatorCode, OperatorName, OperatorPassword, UniqueSaleNo, InvDocNo, InvCgTaxNo, InvCgVatNo, InvCgName, InvCgCity, InvCgAddress, InvCgPrsReceive, StornoReason, StornoReceiptNo, StornoReceiptDate, StornoFiscalMemoryNo, StornoInvoiceNo)
End Function

