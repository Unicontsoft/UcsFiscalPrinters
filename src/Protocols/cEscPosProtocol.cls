VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEscPosProtocol"
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
' ESP/POS printer protocol
'
'=========================================================================
Option Explicit
DefObj A-Z
Const MODULE_NAME As String = "cEscPosProtocol"
Implements IDeviceProtocol

'=========================================================================
' API
'=========================================================================

'--- error codes
Private Const ERROR_IO_PENDING              As Long = 997
'--- for WaitCommEvent
Private Const EV_RXCHAR                     As Long = &H1
Private Const EV_ERR                        As Long = &H80

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

Private Const STR_INTERNAL              As String = "No receipt started|Error in line %1: %2|Unsupported discount type %1"
    
Private Const STR_RECEIPT_TEXTS         As String = "CASH|CHEQUE|COUPON|EXT.COUPON|PACKAGING|INT.USAGE|DAMAGE|CARD|BANK|CUSTOM 1|CUSTOM 2|EUR|SURCHARGE %1|DISCOUNT %1|TOTAL|SUBTOTAL|INVOICE No %1|Sales %1|UNIT PRICE|QUANTITY|VALUE|SUM GROUP %1|VAT %1=%2|NETTO SUM|TOTAL|%1 ITEMS|1 ITEM|CHANGE|SELLER: |RECEIVER: |BUYER: |TAX No: |VAT No: |SYSTEM RECEIPT|ORDER LIST|Date:     %1|Operator: %1|Table:    %1|EIC: %1"
Private Const DEF_AUTODETECTTIMEOUT     As Long = 53
Private Const DEF_CONNECTIONTIMEOUT     As Long = 3000
Private Const DEF_TIMEOUT               As Long = 15000
Private Const DEF_OPERNO                As String = "1"
Private Const DEF_DELIM                 As String = ","
Private Const DEF_AUTODETECTSPEEDS      As String = "115200|38400|9600|19200|57600|4800|2400"
Private Const DEF_SOCKET_PORT           As Long = 9100
Private Const DEF_CODEPAGE              As Long = 866
Private Const DEF_ROW_CHARS             As Long = 48
Private Const DEF_MIN_DISCOUNT          As Double = -100
Private Const DEF_MAX_DISCOUNT          As Double = 100
Private Const DEF_MAX_RECEIPT_ROWS      As Long = 1000
Private Const DEF_MAX_ITEM_LINES        As Long = 2
Private Const STR_GS                    As String = "" ' Chr$(29)
Private Const STR_ESC                   As String = "" ' Chr$(27)

Private m_sDevice                   As String
Private m_oConnector                As IDeviceConnector
Private m_lConnectionTimeout        As Long
Private m_lCodePage                 As Long
Private m_lCommandTimeout           As Long
Private m_uPrintData                As UcsProtocolPrintData
Private m_uConfig                   As UcsConfigValues

Private Enum UcsInternalErrors
    ucsErrNoReceiptStarted
    ucsErrErrorInLine
    ucsErrInvalidDiscType
End Enum

Private Enum UcsReceiptTextsEnum
    ucsTxtPaymentCash
    ucsTxtPaymentCheque
    ucsTxtPaymentCoupon
    ucsTxtPaymentVoucher
    ucsTxtPaymentPackaging
    ucsTxtPaymentMaintenance
    ucsTxtPaymentDamage
    ucsTxtPaymentCard
    ucsTxtPaymentBank
    ucsTxtPaymentCustom1
    ucsTxtPaymentCustom2
    ucsTxtPaymentEUR
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
    ucsTxtNonfiscalReceipt
    ucsTxtOrderList
    ucsTxtHeaderDate
    ucsTxtHeaderOper
    ucsTxtHeaderTable
    ucsTxtHeaderEic
End Enum

Private Type UcsAutodetect
    Connector           As cSerialPortConnector
    WaitOver            As OVERLAPPED
    EventMask           As Long
End Type

Private Type UcsConfigValues
    ConfigCommands      As Collection
    LocalizedCommands   As Collection
    LocalizedText(0 To [_ucsFscLciMax] - 1) As Variant
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

End Property

Property Get LocalizedText(ByVal eIdx As UcsFiscalLocalizedIndexesEnum) As String
    If IsArray(m_uConfig.LocalizedText(eIdx)) Then
        LocalizedText = Join(m_uConfig.LocalizedText(eIdx), "|")
    End If
End Property

Property Let LocalizedText(ByVal eIdx As UcsFiscalLocalizedIndexesEnum, sValue As String)
    m_uConfig.LocalizedText(eIdx) = SplitOrReindex(sValue, "|")
End Property

Private Property Get pvRowChars() As Long
    Const CONF_NAME     As String = "RowChars"
    Dim sValue          As String
    
    If m_uPrintData.Config.RowChars = 0 Then
        pvGetCommandConfig CONF_NAME, "RowChars", 0, 0, vbNullString, vbNullString, _
                "Value", sValue, vbNullString
        m_uPrintData.Config.RowChars = C_Lng(sValue)
        If m_uPrintData.Config.RowChars <= 0 Then
            m_uPrintData.Config.RowChars = DEF_ROW_CHARS
        End If
    End If
    pvRowChars = m_uPrintData.Config.RowChars
End Property

Private Property Get pvCommentChars() As Long
    pvCommentChars = m_uPrintData.Config.CommentChars
    If pvCommentChars = 0 Then
        pvCommentChars = pvRowChars
    End If
End Property

Private Property Get pvItemChars() As Long
    pvItemChars = m_uPrintData.Config.ItemChars
    If pvItemChars = 0 Then
        pvItemChars = pvRowChars - 7
    End If
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function AutodetectDevices(vPorts As Variant) As Variant
    Const FUNC_NAME     As String = "AutodetectDevices"
    Dim vRetVal         As Variant
    Dim vSpeed          As Variant
    Dim aPorts()        As UcsAutodetect
    Dim lIdx            As Long
    Dim baBuffer()      As Byte
    Dim sModel          As String
    Dim sCmd            As String
    Dim lJdx            As Long
    Dim lResult         As Long
    Dim hEvents()       As Long
    Dim lEventsCount    As Long
    Dim lTimeout        As Long
    
    On Error GoTo EH
    vRetVal = vPorts
    If Not IsArray(vPorts) Then
        GoTo QH
    End If
    If UBound(vPorts) < 0 Then
        GoTo QH
    End If
    lTimeout = GetConfigNumber(STR_PROTOCOL_TREMOL, "AutodetectTimeout", DEF_AUTODETECTTIMEOUT)
    For Each vSpeed In Split(DEF_AUTODETECTSPEEDS, "|")
        ReDim aPorts(0 To UBound(vPorts)) As UcsAutodetect
        ReDim hEvents(0 To UBound(vPorts)) As Long
        For lIdx = 0 To UBound(aPorts)
            If Not IsArray(vRetVal(lIdx)) Then
                With aPorts(lIdx)
                    Set .Connector = InitDeviceConnector(vRetVal(lIdx) & "," & vSpeed, lTimeout, LocalizedText(ucsFscLciConnectorErrors))
                    If Not .Connector Is Nothing Then
                        sCmd = STR_ESC & "@"            '--- Initialize printer
                        '--- STR_GS & "I1"
                        sCmd = STR_GS & "r" & Chr$(1)   '--- Transmits paper sensor status
                        .Connector.WriteData ToAscii(sCmd), 1
                        .WaitOver.hEvent = CreateEvent(0, 1, 0, 0)
                        Call SetCommMask(.Connector.hComm, EV_RXCHAR Or EV_ERR)
                        lResult = WaitCommEvent(.Connector.hComm, .EventMask, .WaitOver)
                        If lResult = 0 And Err.LastDllError = ERROR_IO_PENDING Then
                            hEvents(lEventsCount) = .WaitOver.hEvent
                            lEventsCount = lEventsCount + 1
                        End If
                    End If
                End With
            End If
        Next
        If lEventsCount > 0 Then
            lResult = WaitForMultipleObjects(lEventsCount, hEvents(0), 1, lTimeout)
            lEventsCount = 0
        End If
        For lIdx = 0 To UBound(aPorts)
            With aPorts(lIdx)
                If Not .Connector Is Nothing Then
                    If .Connector.ReadData(baBuffer, lTimeout) And UBound(baBuffer) >= 0 Then
                        For lJdx = 0 To UBound(baBuffer)
                            If (baBuffer(lJdx) And &H90) = 0 And baBuffer(lJdx) <> 0 Then
                                DebugLog FUNC_NAME, "Auto-detecting " & .Connector.Device & " reply 0x" & ToHexDump(StrConv(baBuffer, vbUnicode))
                                sModel = pvRawGetAutodetectInfo(.Connector, 67) & pvRawGetAutodetectInfo(.Connector, 66) & _
                                         pvRawGetAutodetectInfo(.Connector, 68) & pvRawGetAutodetectInfo(.Connector, 65)
                                If Len(sModel) > 1 Then
                                    vRetVal(lIdx) = Array(At(vPorts, lIdx), vSpeed, STR_PROTOCOL_ESCPOS, Trim$(Mid$(sModel, 2)))
                                End If
                                Exit For
                            End If
                        Next
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

Private Function pvRawGetAutodetectInfo(oConnector As IDeviceConnector, ByVal lIndex As Long) As String
    Dim sCmd            As String
    Dim sRetVal         As String
    Dim baBuffer()      As Byte
    Dim lRepeat         As Long
    
Repeat:
    sCmd = STR_GS & "I" & Chr$(lIndex)
    If oConnector.WriteData(ToAscii(sCmd), 100) Then
        If oConnector.ReadData(baBuffer, 100) Then
            If baBuffer(0) = 95 Then
                sRetVal = Mid$(StrConv(baBuffer, vbUnicode), 2)
                If InStr(sRetVal, vbNullChar) > 0 Then
                    sRetVal = Left$(sRetVal, InStr(sRetVal, vbNullChar) - 1)
                End If
                sRetVal = Trim$(sRetVal)
            End If
            If LenB(sRetVal) = 0 And lRepeat < 10 Then
                lRepeat = lRepeat + 1
                GoTo Repeat
            End If
        End If
    End If
    If LenB(sRetVal) <> 0 Then
        pvRawGetAutodetectInfo = ", " & sRetVal
    End If
End Function

Public Function Init(sDeviceString As String) As Boolean
    Const FUNC_NAME     As String = "Init"
    Dim oOptions        As Object
    
    On Error GoTo EH
    DebugLog FUNC_NAME, "sDeviceString=" & sDeviceString
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
    sDeviceString = ToDeviceString(oOptions)
    m_lCodePage = Znl(C_Lng(JsonItem(oOptions, "CodePage")), DEF_CODEPAGE)
    With m_uPrintData.Config
        .RowChars = C_Lng(JsonItem(oOptions, "RowChars"))
        .CommentChars = C_Lng(JsonItem(oOptions, "CommentChars"))
        .ItemChars = C_Lng(JsonItem(oOptions, "ItemChars"))
        .AbsoluteDiscount = True
        .NegativePrices = True
        .MinDiscount = Zndbl(C_Dbl(JsonItem(oOptions, "MinDiscount")), DEF_MIN_DISCOUNT)
        .MaxDiscount = Zndbl(C_Dbl(JsonItem(oOptions, "MaxDiscount")), DEF_MAX_DISCOUNT)
        .MaxReceiptRows = Znl(C_Lng(JsonItem(oOptions, "MaxReceiptRows")), DEF_MAX_RECEIPT_ROWS)
        .MaxItemLines = Znl(C_Lng(JsonItem(oOptions, "MaxItemLines")), DEF_MAX_ITEM_LINES)
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
    Dim sCmd            As String
    
    On Error GoTo EH
    DebugLog FUNC_NAME, "m_sDevice=" & m_sDevice & ", m_lConnectionTimeout=" & m_lConnectionTimeout & ", m_lCodePage=" & m_lCodePage
    Set m_oConnector = InitDeviceConnector(m_sDevice, m_lConnectionTimeout, LocalizedText(ucsFscLciConnectorErrors), m_uPrintData.LastError)
    If m_oConnector Is Nothing Then
        DebugLog FUNC_NAME, "InitDeviceConnector failed, m_uPrintData.LastError=" & m_uPrintData.LastError
        GoTo QH
    End If
    sCmd = STR_ESC & "@"                        '--- Initialize printer
    If m_lCodePage = 866 Then
        sCmd = sCmd & STR_ESC & "t" & Chr$(17)  '--- Select character code table: Page 17 [PC866 (Cyrillic #2)]
    End If
    sCmd = sCmd & STR_ESC & "M" & vbNullChar       '--- Select character font: Font A (12 x 24), Font B (9 x 17)
    If Not m_oConnector.WriteData(ToAscii(sCmd), m_lCommandTimeout) Then
        m_uPrintData.LastError = m_oConnector.GetLastError()
        DebugLog FUNC_NAME, "WriteData failed, m_uPrintData.LastError=" & m_uPrintData.LastError
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

Public Function GetClock() As Date
    GetClock = GetCurrentNow
End Function

Public Function GetPaymentName(ByVal Index As Long) As String
    GetPaymentName = pvText(ucsTxtPaymentCash + Index - 1)
End Function

Public Function OpenDrawer() As Boolean
    Dim sCmd            As String
    
    '--- first drawer (Pin2): pulse ON for 50ms then OFF for 500ms
    sCmd = STR_ESC & "p" & vbNullChar & Chr$(25) & Chr$(250)
    If Not m_oConnector.WriteData(ToAscii(sCmd), m_lCommandTimeout) Then
        m_uPrintData.LastError = m_oConnector.GetLastError()
        GoTo QH
    End If
    '--- success
    OpenDrawer = True
QH:
End Function

Public Function CancelReceipt() As Boolean
    CancelReceipt = True
End Function

Public Function GetDeviceProtocol() As String
    GetDeviceProtocol = STR_PROTOCOL_ESCPOS
End Function

Public Function GetDeviceModel() As String
    Dim sModel          As String
    
    sModel = pvRawGetAutodetectInfo(m_oConnector, 67) & pvRawGetAutodetectInfo(m_oConnector, 66) & _
             pvRawGetAutodetectInfo(m_oConnector, 68) & pvRawGetAutodetectInfo(m_oConnector, 65)
    GetDeviceModel = Trim$(Mid$(sModel, 2))
End Function

Public Function GetDeviceStatus(StatusText As String) As UcsFiscalStatusCodeEnum
    Dim sCmd            As String
    Dim baBuffer()      As Byte
    
    sCmd = STR_GS & "r" & Chr$(1)
    If Not m_oConnector.WriteData(ToAscii(sCmd), m_lCommandTimeout) Then
        StatusText = m_oConnector.GetLastError()
        GetDeviceStatus = ucsFscStaFailed
        GoTo QH
    End If
    If Not m_oConnector.ReadData(baBuffer, m_lCommandTimeout) Then
        StatusText = m_oConnector.GetLastError()
        GetDeviceStatus = ucsFscStaFailed
        GoTo QH
    End If
    StatusText = vbNullString
    GetDeviceStatus = ucsFscStaReady
QH:
    pvSetLastError vbNullString
End Function

'= private ===============================================================

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
    Dim vElem           As Variant
    
    If IsArray(vText) Then
        DebugLog FUNC_NAME, "vText=" & Join(vText, vbCrLf)
        For Each vElem In vText
            If Not m_oConnector.WriteData(ToAscii(C_Str(vElem) & vbLf, m_lCodePage), m_lCommandTimeout) Then
                m_uPrintData.LastError = m_oConnector.GetLastError()
                GoTo QH
            End If
        Next
    Else
        DebugLog FUNC_NAME, "vText=" & vText
        If Not m_oConnector.WriteData(ToAscii(C_Str(vText) & vbLf, m_lCodePage), m_lCommandTimeout) Then
            m_uPrintData.LastError = m_oConnector.GetLastError()
            GoTo QH
        End If
    End If
    '--- success
    pvPrintLine = True
QH:
End Function

Private Function pvSendPaperFeed(Optional ByVal Lines As Long = 1) As Boolean
    Const FUNC_NAME     As String = "pvSendPaperFeed"
    Dim sCmd            As String
    
    DebugLog FUNC_NAME, "Lines=" & Lines
    sCmd = STR_ESC & "d" & Chr$(Lines)
    If Not m_oConnector.WriteData(ToAscii(sCmd), m_lCommandTimeout) Then
        m_uPrintData.LastError = m_oConnector.GetLastError()
        GoTo QH
    End If
    '--- success
    pvSendPaperFeed = True
QH:
End Function

Private Function pvSendPaperCut(Optional ByVal Full As Boolean = True, Optional ByVal Lines As Long = 2) As Boolean
    Const FUNC_NAME     As String = "pvSendPaperCut"
    Dim sCmd            As String
    
    DebugLog FUNC_NAME, "Full=" & Full & ", Lines=" & Lines
    sCmd = STR_GS & "V" & IIf(Full, "A" & vbNullChar, "0")
    If Lines > 0 Then
        sCmd = sCmd & STR_ESC & "d" & Chr$(Lines)
    End If
    If Not m_oConnector.WriteData(ToAscii(sCmd), m_lCommandTimeout) Then
        m_uPrintData.LastError = m_oConnector.GetLastError()
        GoTo QH
    End If
    '--- success
    pvSendPaperCut = True
QH:
End Function

Private Function pvPrintReceipt(uCtx As UcsPpdExecuteContext, uRow() As UcsPpdRowData, ByVal lRowCount As Long) As Boolean
    Const FUNC_NAME     As String = "pvPrintReceipt"
    Const CONF_NAME     As String = "PrintReceipt"
    Dim lIdx            As Long
    Dim vSplit          As Variant
    Dim vTemp           As Variant
    Dim sData           As String
    Dim dblTotal        As Double
    Dim dblDisc         As Double
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
            Select Case .InitReceiptType
            Case ucsFscRcpOrderList
                pvPrintLine CenterText(pvText(ucsTxtOrderList), pvCommentChars)
                pvPrintLine vbNullString
                pvPrintLine Printf(pvText(ucsTxtHeaderDate), GetCurrentNow())
                pvPrintLine Printf(pvText(ucsTxtHeaderOper), .InitOperatorName)
                pvPrintLine Printf(pvText(ucsTxtHeaderTable), .InitTableNo)
                pvPrintLine String(pvRowChars, "-")
            Case Else
                pvPrintLine CenterText(At(.InitOwnData, ucsOwnName), pvCommentChars)
                pvPrintLine CenterText(At(.InitOwnData, ucsOwnAddress), pvCommentChars)
                pvPrintLine CenterText(Printf(pvText(ucsTxtHeaderEic), At(.InitOwnData, ucsOwnBulstat)), pvCommentChars)
                If LenB(At(.InitOwnData, ucsOwnDepName)) <> 0 Then
                    pvPrintLine CenterText(At(.InitOwnData, ucsOwnDepName), pvCommentChars)
                End If
                If LenB(At(.InitOwnData, ucsOwnDepAddress)) <> 0 Then
                    pvPrintLine CenterText(At(.InitOwnData, ucsOwnDepAddress), pvCommentChars)
                End If
                If .InitReceiptType = ucsFscRcpInvoice Or .InitReceiptType = ucsFscRcpCreditNote Then
                    sData = .InitInvData(ucsInvDocNo)
                    If Len(sData) <= 10 Then
                        sData = Printf(pvText(ucsTxtHeaderInvoiceNo), sData)
                    End If
                    pvPrintLine CenterText(sData, pvCommentChars)
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
            Select Case .PrintRowType
            Case ucsFscRcpOrderList
                pvGetCommandConfig CONF_NAME, "OrderList" & .PrintRowType, 0, 0, vbNullString, vbNullString, _
                    "FormatQty", sFormatQty, FORMAT_BASE_3
                vSplit = WrapMultiline(.PluItemName, pvItemChars)
                If UBound(vSplit) = 0 Then
                    ReDim Preserve vSplit(0 To 1) As Variant
                End If
                sData = " " & IIf(.PluPrice < -DBL_EPSILON, "-", vbNullString) & SafeFormat(.PluQuantity, sFormatQty)
                lIdx = UBound(vSplit) - 1
                If Len(vSplit(lIdx)) + Len(sData) > pvCommentChars - 2 Then
                    vTemp = WrapMultiline(vSplit(lIdx), pvCommentChars - Len(sData) - 3)
                    vSplit(lIdx) = Trim$(At(vTemp, 0))
                    vSplit(UBound(vSplit)) = Trim$(At(vTemp, 1) & " " & vSplit(UBound(vSplit)))
                End If
                vSplit(lIdx) = AlignText(vSplit(lIdx), sData, pvCommentChars)
                pvPrintLine vSplit
            Case Else
                pvGetCommandConfig CONF_NAME, "NonFiscalSell" & .PrintRowType, 0, 0, vbNullString, vbNullString, _
                    "FormatPrice", sFormatPrice, FORMAT_BASE_2, _
                    "FormatQty", sFormatQty, FORMAT_BASE_3, _
                    "FormatDisc", sFormatDisc, FORMAT_BASE_2, _
                    "FormatAmount", sFormatAmount, FORMAT_BASE_2
                If uCtx.PluCount = 1 Then
                    pvPrintLine AlignText(Format$(uRow(0).InitOperatorCode, "0000") & " " & uRow(0).InitOperatorName, Format$(1, "00"), pvCommentChars)
                End If
                If .PrintRowType = ucsFscRcpSale Then
                    vSplit = WrapMultiline(.PluItemName, pvItemChars)
                    If m_uPrintData.Config.MaxItemLines > 0 Then
                        If UBound(vSplit) >= m_uPrintData.Config.MaxItemLines Then
                            ReDim Preserve vSplit(0 To m_uPrintData.Config.MaxItemLines - 1) As Variant
                        End If
                    End If
                    If Abs(.PluQuantity - 1) > DBL_EPSILON Then
                        If UBound(vSplit) = 0 Then
                            ReDim Preserve vSplit(0 To 1) As Variant
                            vSplit(1) = vSplit(0)
                            vSplit(0) = vbNullString
                        End If
                        sData = " " & IIf(.PluPrice < -DBL_EPSILON, "-", vbNullString) & SafeFormat(.PluQuantity, sFormatQty) & " x " & SafeFormat(Abs(.PluPrice), sFormatPrice)
                        lIdx = UBound(vSplit) - 1
                        If Len(vSplit(lIdx)) + Len(sData) > pvCommentChars - 2 Then
                            vTemp = WrapMultiline(vSplit(lIdx), pvCommentChars - Len(sData) - 3)
                            vSplit(lIdx) = Trim$(At(vTemp, 0))
                            vSplit(UBound(vSplit)) = Trim$(At(vTemp, 1) & " " & vSplit(UBound(vSplit)))
                        End If
                        vSplit(lIdx) = AlignText(vSplit(lIdx), sData, pvCommentChars - 2)
                    End If
                    vSplit(UBound(vSplit)) = AlignText(vSplit(UBound(vSplit)), " " & SafeFormat(dblTotal, sFormatAmount) & " " & Chr$(191 + .PluTaxGroup), pvCommentChars)
                    pvPrintLine vSplit
                Else
                    vSplit = WrapMultiline(.PluItemName, pvCommentChars - 2)
                    If m_uPrintData.Config.MaxItemLines > 0 Then
                        If UBound(vSplit) >= m_uPrintData.Config.MaxItemLines Then
                            ReDim Preserve vSplit(0 To m_uPrintData.Config.MaxItemLines - 1) As Variant
                        End If
                    End If
                    For lIdx = 0 To UBound(vSplit)
                        pvPrintLine C_Str(vSplit(lIdx))
                    Next
                    pvPrintLine AlignText(pvText(ucsTxtPluPrice), SafeFormat(.PluPrice, sFormatPrice), pvCommentChars - 2)
                    pvPrintLine AlignText(pvText(ucsTxtPluQuantity), SafeFormat(.PluQuantity, sFormatQty), pvCommentChars - 2)
                    pvPrintLine AlignText(pvText(ucsTxtPluTotal), SafeFormat(dblTotal, sFormatAmount) & " " & Chr$(191 + .PluTaxGroup), pvCommentChars)
                End If
                Select Case .DiscType
                Case ucsFscDscPlu, ucsFscDscPluAbs
                    sData = AlignText(Printf(IIf(dblDisc > DBL_EPSILON, pvText(ucsTxtSurcharge), pvText(ucsTxtDiscount)), IIf(.DiscType = ucsFscDscPlu, SafeFormat(Abs(.DiscValue), sFormatDisc) & " %", vbNullString)), _
                        SafeFormat(dblDisc, sFormatAmount) & " " & Chr$(191 + .PluTaxGroup), pvCommentChars)
                    pvPrintLine sData
                End Select
            End Select
        Case ucsRowDiscount
            pvGetCommandConfig CONF_NAME, "NonFiscalSubtotal", 0, 0, vbNullString, vbNullString, _
                "FormatDisc", sFormatDisc, FORMAT_BASE_2, _
                "FormatAmount", sFormatAmount, FORMAT_BASE_2
            Select Case .DiscType
            Case ucsFscDscSubtotal
                pvPrintLine AlignText(pvText(ucsTxtSubTotal), SafeFormat(SumArray(uCtx.GrpTotal), sFormatAmount), pvCommentChars - 2)
                For lIdx = 1 To UBound(uCtx.GrpTotal)
                    If Abs(uCtx.GrpTotal(lIdx)) > DBL_EPSILON Then
                        dblTotal = Round(uCtx.GrpTotal(lIdx) * .DiscValue / 100#, DEF_PRICE_SCALE)
                        pvPrintLine AlignText(Printf(IIf(dblTotal > DBL_EPSILON, pvText(ucsTxtSurcharge), pvText(ucsTxtDiscount)), SafeFormat(Abs(.DiscValue), sFormatDisc) & " %"), _
                            SafeFormat(dblTotal, sFormatAmount) & " " & Chr$(191 + lIdx), pvCommentChars)
                    End If
                Next
            Case Else
                pvSetLastError Printf(pvInternal(ucsErrInvalidDiscType), .DiscType)
                GoTo QH
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
            End Select
        Case ucsRowLine
            If .LineWordWrap Then
                vSplit = WrapMultiline(.LineText, pvCommentChars)
            Else
                vSplit = Array(Left$(.LineText, pvCommentChars))
            End If
            Select Case .PrintRowType
            Case Else
                If uCtx.PmtPrinted And Not uCtx.ChangePrinted Then
                    pvGetCommandConfig CONF_NAME, "NonFiscalFooter", 0, 0, vbNullString, vbNullString, _
                        "FormatAmount", sFormatAmount, FORMAT_BASE_2
                    dblTotal = Round(SumArray(uCtx.GrpTotal) - uCtx.Paid, DEF_PRICE_SCALE)
                    If dblTotal > DBL_EPSILON Then
                        pvPrintLine AlignText(GetPaymentName(1), SafeFormat(dblTotal, sFormatAmount), pvCommentChars - 2)
                    ElseIf dblTotal < -DBL_EPSILON Then
                        pvPrintLine AlignText(pvText(ucsTxtFooterChange), SafeFormat(-dblTotal, sFormatAmount), pvCommentChars - 2)
                    End If
                    uCtx.ChangePrinted = True
                End If
                If uCtx.PluCount = 0 Then
                    For lIdx = 0 To UBound(vSplit)
                        vSplit(lIdx) = CenterText(Trim$(vSplit(lIdx)), pvCommentChars)
                    Next
                End If
                pvPrintLine vSplit
            End Select
        Case ucsRowPayment
            Select Case .PrintRowType
            Case ucsFscRcpOrderList
                If .PmtType = 0 Then
                    pvPrintLine String(pvCommentChars, "-")
                    If Not pvSendPaperFeed(Lines:=3) Then
                        GoTo QH
                    End If
                    If Not pvSendPaperCut() Then
                        GoTo QH
                    End If
                End If
            Case Else
                pvGetCommandConfig CONF_NAME, "NonFiscalFooter", 0, 0, vbNullString, vbNullString, _
                    "FormatAmount", sFormatAmount, FORMAT_BASE_2
                If Not uCtx.PmtPrinted Then
                    If uCtx.PluCount > 0 Or .PmtType <> 0 Then
                        pvPrintLine String$(pvCommentChars, "-")
                        pvPrintLine AlignText(pvText(ucsTxtFooterTotal), SafeFormat(SumArray(uCtx.GrpTotal), sFormatAmount), pvCommentChars - 2)
                        If .PrintRowType = ucsFscRcpInvoice Or .PrintRowType = ucsFscRcpCreditNote Then
                            '--- print invoice totals by VAT groups
                            vSplit = Split("0,20,20,9", DEF_DELIM)
                            For lIdx = 1 To UBound(uCtx.GrpTotal)
                                If Abs(uCtx.GrpTotal(lIdx)) > DBL_EPSILON Then
                                    '--- first, round VAT
                                    dblTotal = Round(uCtx.GrpTotal(lIdx) - uCtx.GrpTotal(lIdx) / (1 + C_Dbl(At(vSplit, lIdx - 1)) / 100#), DEF_PRICE_SCALE)
                                    '--- then, calc TaxBase = Total - VAT
                                    dblTotal = Round(uCtx.GrpTotal(lIdx) - dblTotal, DEF_PRICE_SCALE)
                                    pvPrintLine AlignText(Printf(pvText(ucsTxtGroupTotal), Chr$(191 + lIdx)), SafeFormat(uCtx.GrpTotal(lIdx), sFormatAmount), pvCommentChars - 2)
                                    pvPrintLine AlignText(Printf(pvText(ucsTxtGroupVat), Chr$(191 + lIdx), SafeFormat(C_Dbl(At(vSplit, lIdx - 1)), sFormatAmount) & "%"), _
                                        SafeFormat(Round(uCtx.GrpTotal(lIdx) - dblTotal, DEF_PRICE_SCALE), sFormatAmount), pvCommentChars - 2)
                                    pvPrintLine AlignText(pvText(ucsTxtGroupNetto), SafeFormat(dblTotal, sFormatAmount), pvCommentChars - 2)
                                End If
                            Next
                            pvPrintLine String$(pvCommentChars, "-")
                        End If
                        uCtx.PmtPrinted = True
                    End If
                End If
                If .PmtType <> 0 Then
                    If LenB(.PmtName) = 0 Then
                        .PmtName = GetPaymentName(.PmtType)
                    End If
                    pvPrintLine AlignText(.PmtName, SafeFormat(.PmtAmount, sFormatAmount), pvCommentChars - 2)
                    uCtx.Paid = Round(uCtx.Paid + .PmtAmount, DEF_PRICE_SCALE)
                Else
                    If uCtx.PmtPrinted And Not uCtx.ChangePrinted Then
                        dblTotal = Round(SumArray(uCtx.GrpTotal) - uCtx.Paid, DEF_PRICE_SCALE)
                        If dblTotal > DBL_EPSILON Then
                            pvPrintLine AlignText(GetPaymentName(1), SafeFormat(dblTotal, sFormatAmount), pvCommentChars - 2)
                        ElseIf dblTotal < -DBL_EPSILON Then
                            pvPrintLine AlignText(pvText(ucsTxtFooterChange), SafeFormat(-dblTotal, sFormatAmount), pvCommentChars - 2)
                        End If
                        uCtx.ChangePrinted = True
                    End If
                    If .PrintRowType = ucsFscRcpInvoice Or .PrintRowType = ucsFscRcpCreditNote Then
                        With uRow(0)
                            pvPrintLine AlignText(pvText(ucsTxtFooterSeller), STR_CHR1 & ".", pvCommentChars)
                            pvPrintLine AlignText(vbNullString, "/" & .InitOperatorName & "/", pvCommentChars)
                            pvPrintLine String$(pvCommentChars, "-")
                            pvPrintLine AlignText(pvText(ucsTxtFooterReceiver), STR_CHR1 & ".", pvCommentChars)
                            pvPrintLine AlignText(vbNullString, "/" & Zn(C_Str(.InitInvData(ucsInvCgPrsReceive)), .InitInvData(ucsInvCgPrsReceive)) & "/", pvCommentChars)
                            pvPrintLine pvText(ucsTxtFooterBuyer) & .InitInvData(ucsInvCgName)
                            pvPrintLine pvText(ucsTxtFooterTaxNo) & .InitInvData(ucsInvCgTaxNo)
                            If LenB(.InitInvData(ucsInvCgVatNo)) <> 0 Then
                                pvPrintLine pvText(ucsTxtFooterVatNo) & .InitInvData(ucsInvCgVatNo)
                            End If
                            pvPrintLine .InitInvData(ucsInvCgCity)
                            pvPrintLine .InitInvData(ucsInvCgAddress)
                        End With
                    End If
                    sData = At(uRow(0).InitOwnData, ucsOwnFooter1)
                    If LenB(sData) <> 0 Then
                        pvPrintLine CenterText(sData, pvCommentChars)
                    End If
                    sData = At(uRow(0).InitOwnData, ucsOwnFooter2)
                    If LenB(sData) <> 0 Then
                        pvPrintLine CenterText(sData, pvCommentChars)
                    End If
                    If uCtx.PluCount > 0 Then
                        sData = IIf(uCtx.PluCount = 1, pvText(ucsTxtFooterArticleSingle), Printf(pvText(ucsTxtFooterArticles), uCtx.PluCount))
                        pvPrintLine CenterText(sData, pvCommentChars)
                    End If
                    vSplit = Split(GetCurrentNow, " ")
                    sData = AlignText(At(vSplit, 0), At(vSplit, 1), pvCommentChars)
                    pvPrintLine sData
                    sData = CenterText(pvText(ucsTxtNonfiscalReceipt), pvCommentChars)
                    pvPrintLine sData
                    If Not pvSendPaperFeed(Lines:=3) Then
                        GoTo QH
                    End If
                    If Not pvSendPaperCut() Then
                        GoTo QH
                    End If
                End If
            End Select
        End Select
        End With
        uCtx.Row = uCtx.Row + 1
    Loop
    '--- success
    pvPrintReceipt = True
    Exit Function
QH:
    pvSetLastError m_oConnector.GetLastError()
    If uCtx.PluCount > 0 And Not uCtx.PmtPrinted Then
        pvSetLastError Printf(pvInternal(ucsErrErrorInLine), uCtx.PluCount, m_uPrintData.LastError)
    End If
    Exit Function
EH:
    pvSetLastError Err.Description
    PrintError FUNC_NAME
    If uCtx.PluCount > 0 And Not uCtx.PmtPrinted Then
        pvSetLastError Printf(pvInternal(ucsErrErrorInLine), uCtx.PluCount, m_uPrintData.LastError)
    End If
    Err.Raise vbObjectError, , m_uPrintData.LastError
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

Private Sub pvSetLastError(sError As String)
    m_uPrintData.LastError = sError
End Sub

'Private Function SafeText(sText As String) As String
'    Dim lIdx            As Long
'
'    SafeText = sText
'    For lIdx = 0 To 31
'        Select Case lIdx
'        Case 10, 13
'        Case Else
'            SafeText = Replace(SafeText, Chr$(lIdx), vbNullString)
'        End Select
'    Next
'End Function

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Initialize()
    m_lConnectionTimeout = DEF_CONNECTIONTIMEOUT
    m_lCommandTimeout = DEF_TIMEOUT
    LocalizedText(ucsFscLciInternalErrors) = STR_INTERNAL
    Set m_uConfig.ConfigCommands = GetConfigCollection(STR_PROTOCOL_ESCPOS, "Commands")
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
    
End Function

Private Function IDeviceProtocol_Connect() As Boolean
    IDeviceProtocol_Connect = Connect()
End Function

Private Function IDeviceProtocol_PrintDuplicate(Optional DuplicateNo As String, Optional ReceiptNo As String, Optional ReceiptDate As Date) As Boolean
    
End Function

Private Function IDeviceProtocol_Disconnect() As Boolean
    IDeviceProtocol_Disconnect = Disconnect()
End Function

Private Function IDeviceProtocol_PrintReceipt(Optional ResumeToken As String, Optional ReceiptNo As String, Optional ReceiptDate As Date, Optional ReceiptAmount As Double, Optional InvoiceNo As String) As Boolean
    If PpdEndReceipt(m_uPrintData, ResumeToken) Then
        IDeviceProtocol_PrintReceipt = pvPrintReceipt(m_uPrintData.ExecCtx, m_uPrintData.Row, m_uPrintData.RowCount)
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

End Function

Private Function IDeviceProtocol_GetDefaultPassword(OperatorCode As String) As String
    
End Function

Private Function IDeviceProtocol_GetDeviceProtocol() As String
    IDeviceProtocol_GetDeviceProtocol = GetDeviceProtocol()
End Function

Private Function IDeviceProtocol_GetDeviceModel() As String
    IDeviceProtocol_GetDeviceModel = GetDeviceModel()
End Function

Private Function IDeviceProtocol_GetDeviceFirmware() As String
    
End Function

Private Function IDeviceProtocol_GetDeviceStatus(StatusText As String) As UcsFiscalStatusCodeEnum
    IDeviceProtocol_GetDeviceStatus = GetDeviceStatus(StatusText)
End Function

Private Function IDeviceProtocol_GetFiscalMemoryNo() As String

End Function

Private Function IDeviceProtocol_GetFooterText(ByVal Index As Long) As String

End Function

Private Function IDeviceProtocol_GetHeaderText(ByVal Index As Long) As String
    If Index = 1 Then
        If Not GetDeviceStatus(vbNullString) Then
            Err.Raise vbObjectError, , LastError
        End If
    End If
End Function

Private Function IDeviceProtocol_GetLastReceiptNo() As String

End Function

Private Function IDeviceProtocol_GetLastError(Optional ErrNo As UcsFiscalErrorsEnum) As String
    ErrNo = LastErrNo
    IDeviceProtocol_GetLastError = LastError
End Function

Private Function IDeviceProtocol_GetLastInvoiceNo() As String

End Function

Private Function IDeviceProtocol_GetLastQRCodeInfo(Optional LastReceiptDate As Date) As String

End Function

Private Function IDeviceProtocol_GetPaymentName(ByVal Index As Long) As String
    IDeviceProtocol_GetPaymentName = GetPaymentName(Index)
End Function

Private Function IDeviceProtocol_GetResumeToken() As String
    
End Function

Private Function IDeviceProtocol_GetDeviceSerialNo() As String
    IDeviceProtocol_GetDeviceSerialNo = "ESC_" & At(Split(m_sDevice, ","), 0)
End Function

Private Function IDeviceProtocol_GetTaxCaption() As String

End Function

Private Function IDeviceProtocol_GetTaxNo() As String

End Function

Private Function IDeviceProtocol_GetTotalsByPaymentTypes() As Variant
    
End Function

Private Function IDeviceProtocol_GetTotalsByTaxGroups() As Variant
    
End Function

Private Function IDeviceProtocol_GetTotalsStats() As Variant
    
End Function

Private Function IDeviceProtocol_Init(DeviceString As String, Optional EventSink As Object) As Boolean
    IDeviceProtocol_Init = Init(DeviceString)
End Function

Private Property Get IDeviceProtocol_IsConnected() As Boolean
    IDeviceProtocol_IsConnected = IsConnected
End Property

Private Function IDeviceProtocol_OpenDrawer() As Boolean
    IDeviceProtocol_OpenDrawer = OpenDrawer()
End Function

Private Function IDeviceProtocol_PrintMonthlyReport(ByVal ReportType As UcsFiscalReportsTypeEnum, ByVal StartDate As Date, ByVal EndDate As Date, Optional ReceiptNo As String, Optional ReceiptDate As Date) As Boolean
    
End Function

Private Function IDeviceProtocol_PrintDailyXReport(ByVal ReportType As UcsFiscalReportsTypeEnum, Optional ReceiptNo As String, Optional ReceiptDate As Date) As Boolean
    
End Function

Private Function IDeviceProtocol_PrintDailyZReport(ByVal ReportType As UcsFiscalReportsTypeEnum, Optional ReceiptNo As String, Optional ReceiptDate As Date) As Boolean
    
End Function

Private Function IDeviceProtocol_SetClock(ByVal NewDate As Date) As Boolean
    IDeviceProtocol_SetClock = True
End Function

Private Sub IDeviceProtocol_SetLocalizedText(ByVal Index As UcsFiscalLocalizedIndexesEnum, Text As String, Optional ByVal ConfigIndex As Long)
    LocalizedText(Index) = Text
End Sub

Private Function IDeviceProtocol_StartReceipt(ByVal ReceiptType As UcsFiscalReceiptTypeEnum, OperatorCode As String, OperatorName As String, OperatorPassword As String, Optional UniqueSaleNo As String, Optional TableNo As String, Optional ByVal DisablePrinting As Boolean, Optional InvDocNo As String, Optional InvCgTaxNo As String, Optional ByVal InvCgTaxNoType As UcsFiscalTaxNoTypeEnum, Optional InvCgVatNo As String, Optional InvCgName As String, Optional InvCgCity As String, Optional InvCgAddress As String, Optional InvCgPrsReceive As String, Optional OwnData As String, Optional ByVal RevType As UcsFiscalReversalTypeEnum, Optional RevReceiptNo As String, Optional ByVal RevReceiptDate As Date, Optional RevFiscalMemoryNo As String, Optional RevInvoiceNo As String, Optional RevReason As String) As Boolean
    m_uPrintData.Config.RowChars = pvRowChars
    m_uPrintData.Config.CommentChars = pvCommentChars
    m_uPrintData.Config.ItemChars = pvItemChars
    IDeviceProtocol_StartReceipt = PpdStartReceipt(m_uPrintData, ReceiptType, Zn(OperatorCode, DEF_OPERNO), OperatorName, OperatorPassword, TableNo, UniqueSaleNo, DisablePrinting, InvDocNo, InvCgTaxNo, InvCgTaxNoType, InvCgVatNo, InvCgName, InvCgCity, InvCgAddress, InvCgPrsReceive, RevType, RevReceiptNo, RevReceiptDate, RevFiscalMemoryNo, RevInvoiceNo, RevReason, OwnData)
End Function

Private Function IDeviceProtocol_GetDepartmentName(DepartmentCode As String, Optional TaxGroup As Long) As String

End Function

Private Function IDeviceProtocol_GetTaxRates() As Variant

End Function
