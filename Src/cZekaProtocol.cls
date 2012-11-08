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
' $Header: /UcsFiscalPrinter/Src/cZekaProtocol.cls 14    8.11.12 17:07 Wqw $
'
'   Unicontsoft Fiscal Printers Project
'   Copyright (c) 2008-2012 Unicontsoft
'
'   Protocol for Zeka cash registers & FP
'
' $Log: /UcsFiscalPrinter/Src/cZekaProtocol.cls $
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
    ucsZekCmdIsOnline = &H4
    ucsZekCmdIsReady = &H5
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

Private Const STR_INTERNAL              As String = "No device info set|CreateFile failed: |SetCommTimeouts failed: |BuildCommDCB failed: |SetCommState failed: |WriteFile failure: |Timeout waiting for response|ReadFile failed: |Invalid message format or checksum (NACK)|Already connected|Not connected|Invalid response length|Response missing preamble|WaitCommEvent failed: |No receipt started|Error in line %1: %2|Cannot cancel previous receipt"
Private Const STR_STATUSES              As String = "Unknown error|Invalid command|Illegal command|Command denied because of uncommited report|Syntax error|Input register overflow|Input register is zero|Missing transaction for void|Insufficient subtotal|Tax groups conflict"
Private Const STR_ERRORS                As String = "Unknown error|Out of paper|Daily registers overflow|Clock not adjusted|Fiscal receipt open|Receipt not fully paid|Non-fiscal receipt open|Receipt already paid out|Read-only fiscal memory|Bad password or command not allowed|Display missing|24 hours without daily report|Printer overheated|Power down|Electronic journal overflow|Not enough conditions met"
Private Const STR_RECEIPT_TEXTS         As String = "Cash|Card|Cheque|Surcharge%1|Discount%1|Total|Subtotal|Invoice No %1|Sales %1|Unit Price|Quantity|Value*%1|Sum Group %1|VAT*%1=%2|Netto Sum|Total sum|%1 items|1 item|Change|Seller: |Receiver: |Buyer: |Tax No: |VAT No: |EIC;Citizen #;Foreigner #;Official #|VOID"
Private Const DEF_TIMEOUT               As Long = 2000
Private Const DEF_OPERNO                As String = "1"
Private Const DEF_PASSWORD              As String = "0"
Private Const DEF_MAX_RECEIPT_ROWS      As Long = 120
Private Const DEF_MIN_DISCOUNT          As Double = -100
Private Const DEF_MAX_DISCOUNT          As Double = 100
Private Const DEF_PING_TIMEOUT          As Long = 200
Private Const DEF_LINE_WIDTH            As Long = 32

Private m_sDevice                   As String
Private m_hComm                     As Long
Private m_lTimeout                  As Long
Private m_lCashDeskNo               As Long
Private m_lSequence                 As Long
Private m_sLastError                As String
Private m_uRow()                    As UcsRowData
Private m_lRowCount                 As Long
Private m_bFiscalPaymentExecuted    As Boolean
Private m_sCommandLog               As String
Private m_vInfoDiagnostics          As Variant
Private m_eError                    As Long
Private m_eStatus                   As Long
Private m_sPaymentTypes             As String
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
    ucsErrResponseMissingPreamble
    ucsErrWaitCommEventFailed
    ucsErrNoReceiptStarted
    ucsErrErrorInLine
    ucsErrCannotCancelReceipt
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

Private Type UcsAutodetect
    hComm               As Long
    WriteOver           As OVERLAPPED
    ReadOver            As OVERLAPPED
    Buffer              As String
End Type

Private Type UcsConfigValues
    RowChars            As Long
    MaxReceiptRows      As Long
    MinDiscount         As Double
    MaxDiscount         As Double
    PingTimeout         As Long
    DetailedReceipt     As Boolean
    DelayPrint          As Boolean
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
    sText = Chr$(ucsZekCmdIsOnline)
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
            Call Sleep(30)
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
                    If .Buffer = Chr$(ucsZekCmdIsOnline) Then
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
    m_vInfoDiagnostics = Split(pvSendCommand(m_hComm, m_lTimeout, ucsZekCmdInfoDiagnostics, sError:=m_sLastError), ";")
    If LenB(m_sLastError) <> 0 Then
        GoTo QH
    End If
    '--- load config values
    sSerial = At(m_vInfoDiagnostics, 0)
    With m_uConfig
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
        .PingTimeout = C_Lng(GetConfigValue(sSerial, "PingTimeout", 0))
        .DetailedReceipt = C_Bool(GetConfigValue(sSerial, "DetailedReceipt", False))
        .DelayPrint = C_Bool(GetConfigValue(sSerial, "DelayPrint", True))
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

Public Function SendCommand(ByVal eCmd As UcsZekaCommandsEnum, Optional sData As String) As String
    Const FUNC_NAME     As String = "SendCommand"
    
    On Error GoTo EH
    If Not WaitDevice() Then
        Exit Function
    End If
    m_sLastError = vbNullString
    SendCommand = pvSendCommand(m_hComm, m_lTimeout, eCmd, sData, m_sLastError)
    '--- raise event
    RaiseEvent CommandComplete(eCmd, sData, SendCommand)
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function WaitDevice() As Boolean
    Const FUNC_NAME     As String = "WaitDevice"
    Dim sError          As String
    
    On Error GoTo EH
    Do
        If pvSendCommand(m_hComm, pvPingTimeout, ucsZekCmdIsOnline, sError:=sError) <> Chr$(ucsZekCmdIsOnline) Then
            If pvSendCommand(m_hComm, m_lTimeout, ucsZekCmdIsOnline, sError:=sError) <> Chr$(ucsZekCmdIsOnline) Then
                m_sLastError = sError
                Exit Function
            End If
        End If
        If pvSendCommand(m_hComm, pvPingTimeout, ucsZekCmdIsReady) = Chr$(ucsZekCmdIsReady) Then
            WaitDevice = True
            Exit Function
        End If
    Loop
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
        .InitOperatorCode = Zn(OperatorCode, DEF_OPERNO)
        .InitOperatorName = OperatorName
        .InitOperatorPassword = Zn(OperatorPassword, DEF_PASSWORD)
        SplitCgAddress Trim$(InvCity) & vbCrLf & Trim$(InvAddress), sCity, sAddress, 30
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
            Value As Double) As Boolean
    Const FUNC_NAME     As String = "AddDiscount"
    Dim lIdx            As Long
    
    On Error GoTo EH
    '--- sanity check
    If m_lRowCount = 0 Then
        Err.Raise vbObjectError, , pvInternal(ucsErrNoReceiptStarted)
    End If
    Select Case DiscType
    Case ucsFscDstTotal
        AddPLU Printf(IIf(Value > 0, pvText(ucsTxtSurcharge), pvText(ucsTxtDiscount)), vbNullString), Value
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
    Dim lRow            As Long
    Dim sData           As String
    Dim dblTotal        As Double
    Dim dblDisc         As Double
    Dim uCtx            As UcsContext
    Dim vSplit          As Variant
    Dim lIdx            As Long
    
    On Error GoTo EH
    '--- sanity check
    If m_lRowCount = 0 Then
        Err.Raise vbObjectError, , pvInternal(ucsErrNoReceiptStarted)
    End If
    If Not CancelReceipt() Then
        Err.Raise vbObjectError, , pvInternal(ucsErrCannotCancelReceipt)
    End If
    '--- fix fiscal receipts with for more than 120 PLUs
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
                '--- init operator name (if needed)
                vSplit = Split(SendCommand(ucsZekCmdInfoOperator, Pad(.InitOperatorCode, -2)), ";")
                If LenB(m_sLastError) <> 0 Then
                    GoTo QH
                End If
                If LenB(.InitOperatorName) <> 0 Then
                    If Trim$(At(vSplit, 1)) <> Trim$(.InitOperatorName) Or Trim$(At(vSplit, 2)) <> Trim$(.InitOperatorPassword) Then
                        sData = Pad(.InitOperatorCode, -2) & ";" & Pad(.InitOperatorName, 20) & ";" & Pad(.InitOperatorPassword, 4)
                        SendCommand ucsZekCmdInitOperator, sData
                        '--- can return "Command denied because of uncommited report"
'                        If LenB(m_sLastError) <> 0 Then
'                            GoTo QH
'                        End If
                    End If
                Else
                    .InitOperatorName = Trim$(At(vSplit, 1))
                End If
                '--- setup fiscal receipt
                sData = Pad(.InitOperatorCode, -2) & ";" & Pad(.InitOperatorPassword, 4)
                If .InitReceiptType = ucsFscRetInvoiceFiscal Then
                    sData = sData & ";" & IIf(m_uConfig.DetailedReceipt, "1", "0") & ";0" & ";" & IIf(m_uConfig.DelayPrint, "3", "1") & ";" & _
                        Pad(.InitInvData(ucsInvCompany), 26) & ";" & _
                        Pad(.InitInvData(ucsInvMol), 16) & ";" & _
                        Pad(.InitInvData(ucsInvVatNumber), 13) & ";" & _
                        Pad(.InitInvData(ucsInvTaxNumber), 13) & ";" & _
                        Pad(.InitInvData(ucsInvCity), 30)
                    If C_Lng(.InitInvData(ucsInvNumber)) <> 0 Then
                        SendCommand ucsZekCmdInitInvoiceNo, Pad(C_Lng(.InitInvData(ucsInvNumber)), -10) & ";" & Pad(C_Lng(.InitInvData(ucsInvNumber)), -10)
                        If LenB(m_sLastError) <> 0 Then
                            GoTo QH
                        End If
                    End If
                ElseIf m_uConfig.DelayPrint Or m_uConfig.DetailedReceipt Then
                    sData = sData & ";" & IIf(m_uConfig.DetailedReceipt, "1", "0") & ";0" & ";" & IIf(m_uConfig.DelayPrint, "2", "0")
                End If
                SendCommand ucsZekCmdFiscalOpen, sData
            Case ucsFscRetNonfiscal, ucsFscRetInvoiceNonfiscal
                vSplit = Split(SendCommand(ucsZekCmdInfoOperator, Pad(.InitOperatorCode, -2)), ";")
                If LenB(m_sLastError) <> 0 Then
                    GoTo QH
                End If
                If LenB(.InitOperatorName) = 0 Then
                    .InitOperatorName = Trim$(At(vSplit, 1))
                End If
                '--- setup non-fiscal receipt
                sData = Pad(.InitOperatorCode, -2) & ";" & Pad(.InitOperatorPassword, 4)
                SendCommand ucsZekCmdNonFiscalOpen, sData
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
                vSplit = WrapText(.PluName, pvRowChars)
                For lIdx = 0 To UBound(vSplit) - 1
                    If Not pvPrintLine(C_Str(vSplit(lIdx))) Then
                        GoTo QH
                    End If
                Next
                sData = Pad(At(vSplit, UBound(vSplit)), 36)
                sData = sData & ";" & Chr$(191 + .PluVatGroup) & ";"
                If .PluQuantity <> 0 Then
                    sData = sData & pvSafeFormat(.PluPrice, "0.00")
                    If .PluQuantity <> 1 Then
                        sData = sData & "*" & pvSafeFormat(.PluQuantity, "0.000")
                    End If
                    If .DiscType = ucsFscDstPlu And .PluPrice >= 0 Then
                        sData = sData & "," & pvSafeFormat(.DiscValue, "0.00")
                    End If
                Else
                    sData = sData & "0.00"
                End If
                SendCommand ucsZekCmdFiscalSell, sData
            Case ucsFscRetNonfiscal, ucsFscRetInvoiceNonfiscal
                If .PluPrice < 0 Then
                    pvPrintLine CenterText(pvText(ucsTxtPluVoid), pvRowChars)
                End If
                If .PrintRowType = ucsFscRetNonfiscal Then
                    pvPrintLine WrapText(.PluName, pvRowChars)
                    If .PluQuantity <> 1 Then
                        pvPrintLine AlignText(vbNullString, pvSafeFormat(.PluQuantity, "0.000") & " x " & pvSafeFormat(Abs(.PluPrice), "0.00"), pvRowChars)
                    End If
                    pvPrintLine AlignText(Printf(pvText(ucsTxtPluTotal), Chr$(191 + .PluVatGroup)), pvSafeFormat(Abs(dblTotal), "0.00"), pvRowChars)
                Else
                    vSplit = WrapText(.PluName, pvRowChars)
                    For lIdx = 0 To UBound(vSplit)
                        pvPrintLine C_Str(vSplit(lIdx))
                    Next
                    pvPrintLine AlignText(pvText(ucsTxtPluQuantity), pvSafeFormat(.PluQuantity, "0.000"), pvRowChars)
                    pvPrintLine AlignText(pvText(ucsTxtPluPrice), pvSafeFormat(Abs(.PluPrice), "0.00"), pvRowChars)
                    pvPrintLine AlignText(Printf(pvText(ucsTxtPluTotal), Chr$(191 + .PluVatGroup)), pvSafeFormat(Abs(dblTotal), "0.00"), pvRowChars)
                End If
                If .DiscType = ucsFscDstPlu Then
                    pvPrintLine AlignText(Printf(IIf(IIf(.PluPrice < 0, -1, 1) * dblDisc > 0, pvText(ucsTxtSurcharge), pvText(ucsTxtDiscount)), "*" & Chr$(191 + .PluVatGroup) & "  " & pvSafeFormat(Abs(.DiscValue), "0.00") & "%"), _
                        pvSafeFormat(IIf(.PluPrice < 0, -1, 1) * dblDisc, "0.00"), pvRowChars)
                End If
                If .PluPrice < 0 Then
                    pvPrintLine CenterText(String(pvRowChars - 10, "-"), pvRowChars)
                End If
            End Select
        Case ucsRowDiscount
            Select Case .PrintRowType
            Case ucsFscRetFiscal, ucsFscRetInvoiceFiscal
                If .DiscType = ucsFscDstSubtotal Then
                    SendCommand ucsZekCmdFiscalSubtotal, "1;0," & pvSafeFormat(.DiscValue, "0.00")
                End If
            Case ucsFscRetNonfiscal, ucsFscRetInvoiceNonfiscal
                If .DiscType = ucsFscDstSubtotal Then
                    pvPrintLine AlignText(pvText(ucsTxtSubTotal), pvSafeFormat(SumArray(uCtx.GrpTotal), "0.00"), pvRowChars)
                    For lIdx = 1 To UBound(uCtx.GrpTotal)
                        If uCtx.GrpTotal(lIdx) <> 0 Then
                            dblTotal = Round(uCtx.GrpTotal(lIdx) * .DiscValue / 100#, 2)
                            pvPrintLine AlignText(Printf(IIf(dblTotal > 0, pvText(ucsTxtSurcharge), pvText(ucsTxtDiscount)), "*" & Chr$(191 + lIdx) & "  " & pvSafeFormat(Abs(.DiscValue), "0.00") & "%"), _
                                pvSafeFormat(dblTotal, "0.00"), pvRowChars)
                        End If
                    Next
                End If
            End Select
            If .DiscType = ucsFscDstSubtotal Then
                For lIdx = 1 To UBound(uCtx.GrpTotal)
                    If uCtx.GrpTotal(lIdx) <> 0 Then
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
            Case Else ' ucsFscRetNonfiscal, ucsFscRetInvoiceNonfiscal
                If uCtx.PmtPrinted And Not uCtx.ChangePrinted Then
                    dblTotal = Round(SumArray(uCtx.GrpTotal) - uCtx.Paid, 2)
                    If dblTotal > 0 Then
                        pvPrintLine AlignText(GetPaymentName(1), pvSafeFormat(dblTotal, "0.00"), pvRowChars)
                    ElseIf dblTotal < 0 Then
                        pvPrintLine AlignText(pvText(ucsTxtFooterChange), pvSafeFormat(-dblTotal, "0.00"), pvRowChars)
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
                    SendCommand ucsZekCmdFiscalPayment, (.PmtType - 1) & ";0;" & pvSafeFormat(.PmtAmount, "0.00")
                Else
                    SendCommand ucsZekCmdFiscalPayAndClose
                End If
            Case ucsFscRetNonfiscal, ucsFscRetInvoiceNonfiscal
                If Not uCtx.PmtPrinted Then
                    If uCtx.PluCount > 0 Or .PmtType <> 0 Then
                        pvPrintLine vbNullString
                        pvPrintLine AlignText(pvText(ucsTxtFooterTotal), pvSafeFormat(SumArray(uCtx.GrpTotal), "0.00"), pvRowChars)
                        If .PrintRowType = ucsFscRetInvoiceNonfiscal Then
                            '--- print invoice totals by VAT group
                            vSplit = Split(Replace(SendCommand(ucsZekCmdInfoTaxRates), "%", vbNullString), ";")
                            For lIdx = 1 To UBound(uCtx.GrpTotal)
                                If uCtx.GrpTotal(lIdx) <> 0 Then
                                    dblTotal = Round(uCtx.GrpTotal(lIdx) / (1 + C_Dbl(At(vSplit, lIdx - 1)) / 100#), 2)
                                    pvPrintLine AlignText(Printf(pvText(ucsTxtGroupTotal), Chr$(191 + lIdx)), pvSafeFormat(uCtx.GrpTotal(lIdx), "0.00"), pvRowChars)
                                    pvPrintLine AlignText(Printf(pvText(ucsTxtGroupVat), Chr$(191 + lIdx), pvSafeFormat(C_Dbl(At(vSplit, lIdx - 1)), "0.00") & "%"), _
                                        pvSafeFormat(Round(uCtx.GrpTotal(lIdx) - dblTotal, 2), "0.00"), pvRowChars)
                                    pvPrintLine AlignText(pvText(ucsTxtGroupNetto), pvSafeFormat(dblTotal, "0.00"), pvRowChars)
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
                    pvPrintLine AlignText(.PmtName, pvSafeFormat(.PmtAmount, "0.00"), pvRowChars)
                    uCtx.Paid = Round(uCtx.Paid + .PmtAmount, 2)
                Else
                    If uCtx.PmtPrinted And Not uCtx.ChangePrinted Then
                        dblTotal = Round(SumArray(uCtx.GrpTotal) - uCtx.Paid, 2)
                        If dblTotal > 0 Then
                            pvPrintLine AlignText(GetPaymentName(1), pvSafeFormat(dblTotal, "0.00"), pvRowChars)
                        ElseIf dblTotal < 0 Then
                            pvPrintLine AlignText(pvText(ucsTxtFooterChange), pvSafeFormat(-dblTotal, "0.00"), pvRowChars)
                        End If
                        uCtx.ChangePrinted = True
                    End If
                    If .PrintRowType = ucsFscRetInvoiceNonfiscal Then
                        With m_uRow(0)
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
                    SendCommand ucsZekCmdNonFiscalClose
                End If
            End Select
        End Select
        End With
        If LenB(m_sLastError) <> 0 Then
            GoTo QH
        End If
    Next
    '--- success
    EndReceipt = True
QH:
    If LenB(m_sLastError) <> 0 And uCtx.PluCount > 0 Then
        m_sLastError = Printf(pvInternal(ucsErrErrorInLine), uCtx.PluCount, m_sLastError)
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function CopyLastReceipt() As Boolean
    Const FUNC_NAME     As String = "CopyLastReceipt"
    
    On Error GoTo EH
    SendCommand ucsZekCmdPrintDuplicate
    If LenB(m_sLastError) = 0 Then
        '--- success
        CopyLastReceipt = True
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetClock() As Date
    Const FUNC_NAME     As String = "GetClock"
    Dim vSplit          As Variant
    
    On Error GoTo EH
    vSplit = Split(SendCommand(ucsZekCmdInfoDateTime), " ")
    GetClock = C_Date(At(vSplit, 0)) + C_Date(At(vSplit, 1))
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Sub SetClock(ByVal NewDate As Date)
    Const FUNC_NAME     As String = "SetClock"
    
    On Error GoTo EH
    SendCommand ucsZekCmdInitDateTime, Format$(NewDate, "dd-MM-yy") & " " & Format$(NewDate, "hh:mm:ss")
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Public Function RunXReport(ByVal ReportType As UcsFiscalReportsTypeEnum) As Boolean
    Const FUNC_NAME     As String = "RunXReport"
    Dim eCmd            As UcsZekaCommandsEnum
    
    If Not CancelReceipt() Then
        Err.Raise vbObjectError, , pvInternal(ucsErrCannotCancelReceipt)
    End If
    On Error GoTo EH
    Select Case ReportType
    Case ucsFscRstDaily
        eCmd = ucsZekCmdPrintReportDaily
    Case ucsFscRstDailyByItems
        eCmd = ucsZekCmdPrintReportDailyItems
    Case ucsFscRstDailyByDepartment
        eCmd = ucsZekCmdPrintReportDailyDepartments
    End Select
    If eCmd <> 0 Then
        SendCommand eCmd, "X"
        If LenB(m_sLastError) = 0 Then
            '--- success
            RunXReport = True
        End If
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function RunZReport(ByVal ReportType As UcsFiscalReportsTypeEnum) As Boolean
    Const FUNC_NAME     As String = "RunZReport"
    Dim eCmd            As UcsZekaCommandsEnum
    
    If Not CancelReceipt() Then
        Err.Raise vbObjectError, , pvInternal(ucsErrCannotCancelReceipt)
    End If
    On Error GoTo EH
    Select Case ReportType
    Case ucsFscRstDaily
        eCmd = ucsZekCmdPrintReportDaily
    Case ucsFscRstDailyByItems
        eCmd = ucsZekCmdPrintReportDailyItems
    Case ucsFscRstDailyByDepartment
        eCmd = ucsZekCmdPrintReportDailyDepartments
    End Select
    If eCmd <> 0 Then
        SendCommand eCmd, "Z"
        If LenB(m_sLastError) = 0 Then
            '--- success
            RunZReport = True
        End If
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function RunPeriodReport(ByVal ReportType As UcsFiscalReportsTypeEnum, ByVal StartDate As Date, ByVal EndDate As Date) As Boolean
    Const FUNC_NAME     As String = "RunPeriodReport"
    Dim eCmd            As UcsZekaCommandsEnum
    
    If Not CancelReceipt() Then
        Err.Raise vbObjectError, , pvInternal(ucsErrCannotCancelReceipt)
    End If
    On Error GoTo EH
    Select Case ReportType
    Case ucsFscRstPeriodShort
        eCmd = ucsZekCmdPrintReportByDateShort
    Case ucsFscRstPeriodDetailed
        eCmd = ucsZekCmdPrintReportByDateDetailed
    End Select
    If eCmd <> 0 Then
        SendCommand eCmd, Format$(StartDate, "ddmmyy") & ";" & Format$(EndDate, "ddmmyy")
        If LenB(m_sLastError) = 0 Then
            '--- success
            RunPeriodReport = True
        End If
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetLastReceiptNumber() As Long
    Const FUNC_NAME     As String = "GetLastReceiptNumber"
    
    On Error GoTo EH
    GetLastReceiptNumber = C_Lng(At(Split(SendCommand(ucsZekCmdInfoLastReceipt), ";"), 0))
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetLastInvoiceNumber() As String
    Const FUNC_NAME     As String = "GetLastInvoiceNumber"
    
    On Error GoTo EH
    GetLastInvoiceNumber = C_Lng(At(Split(SendCommand(ucsZekCmdInfoInvoiceNo), ";"), 0))
    If GetLastInvoiceNumber <> 0 Then
        GetLastInvoiceNumber = GetLastInvoiceNumber - 1
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetSerialNumber() As String
    Const FUNC_NAME     As String = "GetSerialNumber"
    
    On Error GoTo EH
    GetSerialNumber = At(m_vInfoDiagnostics, 0)
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetFiscalNumber() As String
    Const FUNC_NAME     As String = "GetFiscalNumber"
    
    On Error GoTo EH
    GetFiscalNumber = At(m_vInfoDiagnostics, 1)
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetTaxNumber() As String
    Const FUNC_NAME     As String = "GetTaxNumber"
    
    On Error GoTo EH
    GetTaxNumber = At(Split(SendCommand(ucsZekCmdInfoBulstat), ";"), 0)
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetTaxCaption() As String
    Const FUNC_NAME     As String = "GetTaxCaption"
    
    On Error GoTo EH
    GetTaxCaption = At(Split(pvText(ucsTxtBulstatCaption), ";"), C_Lng(At(Split(SendCommand(ucsZekCmdInfoBulstat), ";"), 1)))
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetHeaderText(ByVal Index As Long) As String
    Const FUNC_NAME     As String = "GetHeaderText"
    
    On Error GoTo EH
    If Index >= 1 And Index <= 6 Then
        GetHeaderText = At(Split(SendCommand(ucsZekCmdInfoHeaderFooter, C_Str(Index)), ";"), 1)
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetFooterText(ByVal Index As Long) As String
    Const FUNC_NAME     As String = "GetFooterText"
    
    On Error GoTo EH
    If Index >= 1 And Index <= 1 Then
        GetFooterText = At(Split(SendCommand(ucsZekCmdInfoHeaderFooter, C_Str(7 + Index)), ";"), 1)
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetPaymentName(ByVal Index As Long) As String
    Const FUNC_NAME     As String = "GetPaymentName"
    
    On Error GoTo EH
    If Index >= 1 And Index <= 3 Then
        GetPaymentName = pvText(ucsTxtPaymentCash + Index - 1)
    ElseIf Index >= 4 And Index <= 5 Then
        If LenB(m_sPaymentTypes) = 0 Then
            m_sPaymentTypes = SendCommand(ucsZekCmdInfoPaymentTypes)
        End If
        GetPaymentName = Trim$(At(Split(m_sPaymentTypes, ";"), Index - 1))
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function OpenDrawer() As Boolean
    Const FUNC_NAME     As String = "OpenDrawer"
    
    On Error GoTo EH
    SendCommand ucsZekCmdInitOpenDrawer
    If LenB(m_sLastError) = 0 Then
        '--- success
        OpenDrawer = True
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function CancelReceipt() As Boolean
    Const FUNC_NAME     As String = "CancelReceipt"
    
    On Error GoTo EH
    If C_Lng(At(Split(SendCommand(ucsZekCmdInfoTransaction), ";"), 0)) = 1 Then
        SendCommand ucsZekCmdFiscalCancel
        If LenB(m_sLastError) <> 0 Then
            SendCommand ucsZekCmdFiscalPayAndClose
            If LenB(m_sLastError) <> 0 Then
                GoTo QH
            End If
        End If
    End If
    '--- success
    CancelReceipt = True
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

Public Function GetTotalsByVatGroups() As Variant
    Const FUNC_NAME     As String = "GetTotalsByVatGroups"
    Dim vSplit          As Variant
    Dim vRetVal         As Variant
    Dim lIdx            As Long
    
    On Error GoTo EH
    vSplit = Split(SendCommand(ucsZekCmdInfoTotalsByVatGroups), ";")
    If UBound(vSplit) > 0 Then
        ReDim vRetVal(0 To 7) As Double
        For lIdx = 0 To UBound(vRetVal)
            If InStr(At(vSplit, lIdx), ".") > 0 Then
                vRetVal(lIdx) = C_Dbl(Replace(At(vSplit, lIdx), ".", g_sDecimalSeparator))
            Else
                vRetVal(lIdx) = C_Dbl(At(vSplit, lIdx)) / 100#
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

Public Function CashDebitCredit(OperatorCode As String, OperatorPassword As String, ByVal PmtNumber As UcsFiscalPaymentTypeEnum, ByVal dblValue As Double, Optional Comment As String) As Variant
    Const FUNC_NAME     As String = "CashDebitCredit"
    Dim sData           As String
    Dim vSplit          As Variant
    Dim vRetVal         As Variant
    
    On Error GoTo EH
    If dblValue <> 0 Then
        sData = Pad(Zn(OperatorCode, DEF_OPERNO), -2) & ";" & Pad(Zn(OperatorPassword, DEF_PASSWORD), 4) & ";" & LimitLong(PmtNumber - 1, 0, 7) & ";" & pvSafeFormat(dblValue, "0.00") & "@" & Left$(Comment, 34)
        SendCommand ucsZekCmdAdminCashDebitCredit, sData
        If LenB(LastError) <> 0 Then
            vRetVal = EmptyDoubleArray
            GoTo QH
        End If
    End If
    vSplit = Split(SendCommand(ucsZekCmdInfoRegisters, "0"), ";")
    If LenB(LastError) <> 0 Then
        vRetVal = EmptyDoubleArray
        GoTo QH
    End If
    ReDim vRetVal(0 To 2) As Double
    vRetVal(0) = C_Dbl(At(vSplit, PmtNumber))
    vSplit = Split(SendCommand(ucsZekCmdInfoRegisters, "2"), ";")
    vRetVal(1) = C_Dbl(At(vSplit, PmtNumber))
    vSplit = Split(SendCommand(ucsZekCmdInfoRegisters, "3"), ";")
    vRetVal(2) = C_Dbl(At(vSplit, PmtNumber))
QH:
    CashDebitCredit = vRetVal
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
    If BuildCommDCB("baud=" & At(vSplit, 1, "9600") & " data=" & At(vSplit, 2, "8") & " parity=" & At(vSplit, 3, "N") & " stop=" & At(vSplit, 4, "1"), uDCB) = 0 Then
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
    RaiseError FUNC_NAME
End Function

Private Function pvSendCommand( _
            ByVal hComm As Long, _
            ByVal lTimeout As Long, _
            ByVal eCmd As UcsZekaCommandsEnum, _
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
    DebugLog FUNC_NAME, "eCmd=0x" & Hex(eCmd) & ", sData=" & sData
    uOver.hEvent = CreateEvent(0, 1, 0, 0)
RetrySend:
    '--- prepare command
    If eCmd = ucsZekCmdIsOnline Or eCmd = ucsZekCmdIsReady Then
        sText = Chr$(eCmd)
    ElseIf eCmd = ucsZekCmdInitLogo Or eCmd = ucsZekCmdInitLogoByNum Then
        sText = Chr$(ZFP_STX) & Chr$(&H39) & Chr$(&H37) & Chr$(eCmd) & sData
        lTimeout = 10000
    Else
        sText = Chr$(ZFP_STX) & Chr$(ZFP_BASEDATA + 3 + Len(sData)) & Chr$(m_lSequence) & Chr$(eCmd) & sData
        '--- increment sequence
        m_lSequence = IIf(m_lSequence >= &H9F, ZFP_BASEDATA, m_lSequence + 1)
        '--- calc checksum
        baText = StrConv(sText, vbFromUnicode)
        For lIdx = 1 To UBound(baText)
            lSum = lSum Xor baText(lIdx)
        Next
        '--- append checksum and terminator
        sText = sText & Chr$((lSum \ &H10 Mod &H10) + &H30) & Chr$((lSum \ &H1 Mod &H10) + &H30) & Chr$(ZFP_ETX)
    End If
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
        DebugLog FUNC_NAME, "CancelIo after WriteFile, lResult=" & lResult & ", lTimeout=" & lTimeout
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
                sText = sText & Left$(sBuffer, lActual)
            End If
        End If
        If lResult <> WAIT_OBJECT_0 Then
            DebugLog FUNC_NAME, "CancelIo after ReadFile, lResult=" & lResult & ", lTimeout=" & lTimeout
            If lResult = WAIT_TIMEOUT Then
                sError = pvInternal(ucsErrTimeoutWaitingForResponse)
            End If
            Call CancelIo(hComm)
            Call Sleep(100)
            Call FlushFileBuffers(hComm)
            Exit Do
        End If
        If LenB(sText) <> 0 Then
            Select Case Asc(sText)
            Case ZFP_STX            ' &H2
                '--- check if response complete
                If Len(sText) > 6 And Asc(Right$(sText, 1)) = ZFP_ETX Then
                    lSize = Asc(Mid$(sText, 2, 1)) - ZFP_BASEDATA - 3
                    '--- check message length
                    If Len(sText) - 7 = lSize Then
                        pvSendCommand = Mid$(sText, 5, lSize)
                        DebugLog FUNC_NAME, "RetVal=" & pvSendCommand
                    Else
                        sError = pvInternal(ucsErrInvalidResponseLength)
                    End If
                    Exit Do
                End If
            Case ZFP_PING, ZFP_BUSY ' &H4, &H5
                If Len(sText) = 1 Then
                    pvSendCommand = sText
                Else
                    sError = pvInternal(ucsErrInvalidResponseLength)
                End If
                Exit Do
            Case ZFP_ACK            ' &H6
                If Len(sText) > 6 And Asc(Right$(sText, 1)) = ZFP_ETX Then
                    lSize = Len(sText) - 5
                    '--- parse respone
                    sText = Mid$(sText, 3, lSize)
                    If Len(sText) > 1 Then
                        m_eError = Asc(Mid$(sText, 1)) - &H30
                        m_eStatus = Asc(Mid$(sText, 2)) - &H30
                        If m_eError <> 0 Then
                            sError = ErrorText & IIf(m_eStatus <> 0, ". ", vbNullString)
                        Else
                            sError = vbNullString
                        End If
                        If m_eStatus <> 0 Then
                            sError = sError & StatusText
                        End If
                    Else
                        sError = pvInternal(ucsErrInvalidResponseLength)
                    End If
                    Exit Do
                End If
            Case ZFP_RETRY          ' &HE
                DebugLog FUNC_NAME, "Retry requested, sText=" & sText
                Call CancelIo(hComm)
                Call Sleep(100)
                Call FlushFileBuffers(hComm)
                GoTo RetrySend
            Case ZFP_NACK           ' &H15
                sError = pvInternal(ucsErrInvalidMessageFormatOrChecksum)
                Exit Do
            Case Else
                sError = pvInternal(ucsErrResponseMissingPreamble)
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

Private Function pvPrintLine(vText As Variant) As Boolean
    Const FUNC_NAME     As String = "pvPrintLine"
    Dim vElem           As Variant
    
    On Error GoTo EH
    If IsArray(vText) Then
        For Each vElem In vText
            If LenB(C_Str(vElem)) = 0 Then
                SendCommand ucsZekCmdPrintText, " "
            Else
                SendCommand ucsZekCmdPrintText, C_Str(vElem)
            End If
            If LenB(m_sLastError) <> 0 Then
                Exit Function
            End If
        Next
    Else
        If LenB(C_Str(vText)) = 0 Then
            SendCommand ucsZekCmdPrintText, " "
        Else
            SendCommand ucsZekCmdPrintText, C_Str(vText)
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
                        AddPLU Printf(IIf(dblTotal * dblDiscount > 0, pvText(ucsTxtSurcharge), pvText(ucsTxtDiscount)), "*" & Chr$(191 + m_uRow(lRow).PluVatGroup) & "  " & pvSafeFormat(Abs(dblDiscount), "0.00") & "%"), _
                            Round(dblTotal * dblDiscount / 100#, 2), 1, m_uRow(lRow).PluVatGroup, lRow + 1
                    End If
                ElseIf m_uRow(lRow).DiscType = ucsFscDstPlu And m_uRow(lRow).PluPrice < 0 Then
                    '--- convert PLU discount on void rows
                    dblTotal = Round(m_uRow(lRow).PluQuantity * m_uRow(lRow).PluPrice, 2)
                    dblDiscount = m_uRow(lRow).DiscValue
                    m_uRow(lRow).DiscType = 0
                    m_uRow(lRow).DiscValue = 0
                    AddPLU Printf(IIf(dblTotal * dblDiscount > 0, pvText(ucsTxtSurcharge), pvText(ucsTxtDiscount)), "*" & Chr$(191 + m_uRow(lRow).PluVatGroup) & "  " & pvSafeFormat(Abs(dblDiscount), "0.00") & "%"), _
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
                            If uSum.GrpTotal(lIdx) <> 0 Then
                                AddPLU Printf(IIf(uSum.GrpTotal(lIdx) * dblDiscount > 0, pvText(ucsTxtSurcharge), pvText(ucsTxtDiscount)), "*" & Chr$(191 + lIdx) & "  " & pvSafeFormat(Abs(dblDiscount), "0.00") & "%"), _
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

Private Function pvGetDeviceInfo(ByVal hComm As Long, ByVal lTimeout As Long, vResult As Variant, sProtocol As String, sModel As String, Optional sError As String) As Boolean
    Const FUNC_NAME     As String = "pvGetDeviceInfo"
    Dim vSplit          As Variant
    
    On Error GoTo EH
    If Not IsArray(vResult) Then
        vResult = Split(pvSendCommand(hComm, lTimeout, ucsZekCmdInfoDiagnostics, sError:=sError), ";")
    End If
    If LenB(At(vResult, 0)) <> 0 Then
        sProtocol = STR_PROTOCOL_ZEKA_FP
        vSplit = Split(pvSendCommand(hComm, lTimeout, ucsZekCmdInfoVersion), ";")
        sModel = Trim$(At(vSplit, 3) & " " & At(vSplit, 4))
        If InStr(sModel, "Ê.Ñ.") > 0 Then
            sModel = Trim$(Left$(sModel, InStr(sModel, "Ê.Ñ.") - 1))
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
    m_lSequence = ZFP_BASEDATA
    LocalizedText(ucsFscLciInternalErrors) = STR_INTERNAL
    LocalizedText(ucsFscLciPrinterStatuses) = STR_STATUSES
    LocalizedText(ucsFscLciPrinterErrors) = STR_ERRORS
    LocalizedText(ucsFscLciReceiptTexts) = STR_RECEIPT_TEXTS
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
    IDeviceProtocol_GetDefaultPassword = DEF_PASSWORD
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

