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
' $Header: /UcsFiscalPrinter/Src/cEscPosProtocol.cls 5     26.06.18 16:30 Wqw $
'
'   Unicontsoft Fiscal Printers Project
'   Copyright (c) 2008-2018 Unicontsoft
'
'   ESP/POS printer protocol
'
' $Log: /UcsFiscalPrinter/Src/cEscPosProtocol.cls $
' 
' 5     26.06.18 16:30 Wqw
' REF: copy last receipt params
'
' 4     13.06.18 12:21 Wqw
' REF: impl configurable row chars
'
' 3     15.05.18 12:29 Wqw
' REF: handle open drawer and get device model
'
' 2     14.05.18 12:38 Wqw
' REF: impl AutodetectDevices
'
' 1     25.04.18 10:38 Wqw
' Initial implementation
'
'=========================================================================
Option Explicit
DefObj A-Z
Const MODULE_NAME As String = "cEscPosProtocol"
Implements IDeviceProtocol

'=========================================================================
' API
'=========================================================================

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_INTERNAL              As String = "No receipt started|Error in line %1: %2"
Private Const STR_RECEIPT_TEXTS         As String = "CASH|CARD|CHEQUE|CREDIT|SURCHARGE %1|DISCOUNT %1|TOTAL|SUBTOTAL|INVOICE No %1|Sales %1|UNIT PRICE|QUANTITY|VALUE|SUM GROUP %1|VAT %1=%2|NETTO SUM|TOTAL|%1 ITEMS|1 ITEM|CHANGE|SELLER: |RECEIVER: |BUYER: |TAX No: |VAT No: |SYSTEM RECEIPT|ORDER LIST|Date:     %1|Operator: %1|Table:    %1|EIC: %1"
Private Const DEF_TIMEOUT               As Long = 3000
Private Const DEF_OPERNO                As String = "1"
Private Const DEF_DELIM                 As String = ","
Private Const DEF_AUTODETECTSPEEDS      As String = "115200|9600"
Private Const DEF_ROW_CHARS             As Long = 48
Private Const ucsFscDstPluAbs           As Long = ucsFscDstPlu + 100
Private Const ucsFscDstSubtotalAbs      As Long = ucsFscDstSubtotal + 100
Private Const STR_GS                    As String = "" ' Chr$(29)
Private Const STR_ESC                   As String = "" ' Chr$(27)

Private m_sDevice                   As String
Private m_oConnector                As IDeviceConnector
Private m_lTimeout                  As Long
Private m_lCashDeskNo               As Long
Private m_eLastErrNumber            As UcsFiscalErrorsEnum
Private m_sLastError                As String
Private m_uRow()                    As UcsRowData
Private m_lRowCount                 As Long
Private m_uConfig                   As UcsConfigValues
Private m_uCtxPrint                 As UcsContext

Private Enum UcsInternalErrors
    ucsErrNoReceiptStarted
    ucsErrErrorInLine
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
    ucsTxtNonfiscalReceipt
    ucsTxtOrderList
    ucsTxtHeaderDate
    ucsTxtHeaderOper
    ucsTxtHeaderTable
    ucsTxtHeaderEic
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
    InitTableNo         As String
    InitInvData         As Variant
    InitOwnData         As Variant
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

Private Enum UcsOwnDataIndex
    ucsOwnName
    ucsOwnAddress
    ucsOwnBulstat
    ucsOwnDepName
    ucsOwnDepAddress
    ucsOwnFooter
End Enum

Private Type UcsContext
    GrpTotal(1 To 8)    As Double
    Paid                As Double
    PluCount            As Long
    PmtPrinted          As Boolean
    ChangePrinted       As Boolean
    Row                 As Long
End Type

Private Type UcsConfigValues
    RowChars            As Long
    ItemChars           As Long
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
    
    If m_uConfig.RowChars = 0 Then
        pvGetCommandConfig CONF_NAME, "RowChars", 0, 0, vbNullString, vbNullString, _
                "Value", sValue, vbNullString
        m_uConfig.RowChars = C_Lng(sValue)
        If m_uConfig.RowChars <= 0 Then
            m_uConfig.RowChars = DEF_ROW_CHARS
        End If
    End If
    pvRowChars = m_uConfig.RowChars
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function AutodetectDevices(vPorts As Variant) As Variant
    Const FUNC_NAME     As String = "AutodetectDevices"
    Dim vRetVal         As Variant
    Dim vSpeed          As Variant
    Dim aPorts          As Variant
    Dim oConnector      As IDeviceConnector
    Dim lIdx            As Long
    Dim baBuffer()      As Byte
    Dim dblTimer        As Double
    Dim bHasMore        As Boolean
    Dim sModel          As String
    Dim sCmd            As String
    
    On Error GoTo EH
    vRetVal = vPorts
    For Each vSpeed In Split(DEF_AUTODETECTSPEEDS, "|")
        ReDim aPorts(0 To UBound(vPorts)) As Variant
        For lIdx = 0 To UBound(aPorts)
            If Not IsArray(vRetVal(lIdx)) Then
                Set oConnector = InitDeviceConnector(vRetVal(lIdx) & "," & vSpeed, 1000, LocalizedText(ucsFscLciConnectorErrors), m_sLastError)
                If Not oConnector Is Nothing Then
                    sCmd = STR_ESC & "@"            '--- Initialize printer
                    '--- STR_GS & "I1"
                    sCmd = STR_GS & "r" & Chr$(1)   '--- Transmits paper sensor status
                    oConnector.WriteData ToAscii(sCmd), 100
                    Set aPorts(lIdx) = oConnector
                End If
            End If
        Next
        dblTimer = DateTimer
        Do While DateTimer < dblTimer + 0.2
            bHasMore = False
            For lIdx = 0 To UBound(aPorts)
                If IsObject(aPorts(lIdx)) Then
                    bHasMore = True
                    Set oConnector = aPorts(lIdx)
                    If oConnector.ReadData(baBuffer, 0) Then
                        sModel = pvGetAutodetectInfo(oConnector, 67) & pvGetAutodetectInfo(oConnector, 66) & _
                                 pvGetAutodetectInfo(oConnector, 68) & pvGetAutodetectInfo(oConnector, 65)
                        vRetVal(lIdx) = Array(At(vPorts, lIdx), vSpeed, STR_PROTOCOL_ESP_POS, Mid$(sModel, 2))
                        aPorts(lIdx) = Empty
                    End If
                End If
            Next
            If Not bHasMore Then
                Exit For
            End If
            Call Sleep(10)
        Loop
    Next
    AutodetectDevices = vRetVal
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Function pvGetAutodetectInfo(oConnector As IDeviceConnector, ByVal lIndex As Long) As String
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
                If InStr(sRetVal, Chr$(0)) > 0 Then
                    sRetVal = Left$(sRetVal, InStr(sRetVal, Chr$(0)) - 1)
                End If
            End If
            If LenB(sRetVal) = 0 And lRepeat < 10 Then
                lRepeat = lRepeat + 1
                GoTo Repeat
            End If
        End If
    End If
    If LenB(sRetVal) <> 0 Then
        pvGetAutodetectInfo = ", " & sRetVal
    End If
End Function

Public Function Init( _
            sDevice As String, _
            Optional ByVal Timeout As Long, _
            Optional ByVal CashDeskNo As Long, _
            Optional ByVal RowChars As Long) As Boolean
    Const FUNC_NAME     As String = "Init"
    Dim vSplit          As Variant
    
    On Error GoTo EH
    DebugLog FUNC_NAME, "sDevice=" & sDevice & ", Timeout=" & Timeout
    vSplit = Split(sDevice, ",")
    If LCase$(Left$(At(vSplit, 0), 3)) = "com" Then
        m_sDevice = sDevice
    Else
        m_sDevice = vSplit(0) & IIf(InStr(At(vSplit, 0), ":") = 0, ":9100", vbNullString)
    End If
    m_lTimeout = IIf(Timeout <> 0, Timeout, DEF_TIMEOUT)
    m_lCashDeskNo = CashDeskNo
    m_uConfig.RowChars = RowChars
    '--- success
    Init = Connect()
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function Connect() As Boolean
    Const FUNC_NAME     As String = "Connect"
    Dim sCmd            As String
    
    On Error GoTo EH
    Set m_oConnector = InitDeviceConnector(m_sDevice, m_lTimeout, LocalizedText(ucsFscLciConnectorErrors), m_sLastError)
    If m_oConnector Is Nothing Then
        GoTo QH
    End If
    sCmd = STR_ESC & "@"                        '--- Initialize printer
    sCmd = sCmd & STR_ESC & "t" & Chr$(17)      '--- Select character code table: Page 17 [PC866 (Cyrillic #2)]
    sCmd = sCmd & STR_ESC & "M" & Chr$(0)       '--- Select character font: Font A (12 x 24), Font B (9 x 17)
    If Not m_oConnector.WriteData(ToAscii(sCmd), m_lTimeout) Then
        m_sLastError = m_oConnector.GetLastError()
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
    Set m_oConnector = Nothing
    '--- success
    Disconnect = True
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function StartReceipt( _
            ByVal ReceiptType As UcsFiscalReceiptTypeEnum, _
            Optional OperatorCode As String, _
            Optional OperatorName As String, _
            Optional OperatorPassword As String, _
            Optional TableNo As String, _
            Optional InvNumber As String, _
            Optional InvTaxNumber As String, _
            Optional InvVatNumber As String, _
            Optional InvCompany As String, _
            Optional InvCity As String, _
            Optional InvAddress As String, _
            Optional InvMol As String, _
            Optional InvReceiver As String, _
            Optional OwnData As String) As Boolean
    Const FUNC_NAME     As String = "StartReceipt"
    Dim sCity           As String
    Dim sAddress        As String
    
    On Error GoTo EH
    ReDim m_uRow(0 To 10) As UcsRowData
    m_lRowCount = 0
    With m_uRow(pvAddRow())
        .RowType = ucsRowInit
        .InitReceiptType = LimitLong(ReceiptType, 1, [_ucsFscRetMax] - 1)
        Select Case .InitReceiptType
        Case ucsFscRetFiscal
            .InitReceiptType = ucsFscRetNonfiscal
        Case ucsFscRetInvoiceFiscal
            .InitReceiptType = ucsFscRetInvoiceNonfiscal
        End Select
        .InitReceiptType = ucsFscRetOrderList
        .InitOperatorCode = Zn(SafeText(OperatorCode), DEF_OPERNO)
        .InitOperatorName = SafeText(OperatorName)
        .InitOperatorPassword = SafeText(OperatorPassword)
        .InitTableNo = TableNo
        If LenB(.InitOperatorPassword) = 0 Then
            .InitOperatorPassword = IDeviceProtocol_GetDefaultPassword(.InitOperatorCode)
        End If
        .InitCashDeskNo = m_lCashDeskNo
        SplitCgAddress Trim$(SafeText(InvCity)) & vbCrLf & Trim$(SafeText(InvAddress)), sCity, sAddress, pvRowChars - 2
        .InitInvData = Array(SafeText(InvNumber), SafeText(InvTaxNumber), SafeText(InvVatNumber), _
            SafeText(InvCompany), sCity, sAddress, SafeText(InvMol), SafeText(InvReceiver))
        .InitOwnData = Split(OwnData, STR_CHR1)
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
        .PluQuantity = Round(IIf(bNegative, Abs(Quantity), Quantity), 3)
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
            .DiscType = DiscType
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
    
    On Error GoTo EH
    #If sResumeTicket Then '--- touch args
    #End If
    '--- sanity check
    If m_lRowCount = 0 Then
        pvSetLastError pvInternal(ucsErrNoReceiptStarted)
        GoTo QH
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

Public Function GetClock() As Date
    GetClock = Now
End Function

Public Function GetPaymentName(ByVal Index As Long) As String
    GetPaymentName = pvText(ucsTxtPaymentCash + Index - 1)
End Function

Public Function OpenDrawer() As Boolean
    Dim sCmd            As String
    
    '--- first drawer (Pin2): pulse ON for 50ms then OFF for 500ms
    sCmd = STR_ESC & "p" & Chr$(0) & Chr$(25) & Chr$(250)
    If Not m_oConnector.WriteData(ToAscii(sCmd), m_lTimeout) Then
        m_sLastError = m_oConnector.GetLastError()
        GoTo QH
    End If
    '--- success
    OpenDrawer = True
QH:
End Function

Public Function CancelReceipt() As Boolean
    CancelReceipt = True
End Function

Public Function GetDeviceModel() As String
    GetDeviceModel = pvGetAutodetectInfo(m_oConnector, 67) & pvGetAutodetectInfo(m_oConnector, 66) & _
                     pvGetAutodetectInfo(m_oConnector, 68) & pvGetAutodetectInfo(m_oConnector, 65)
End Function

Public Function GetDeviceProtocol() As String
    GetDeviceProtocol = STR_PROTOCOL_ESP_POS
End Function

Public Function GetDeviceStatus(CurrentStatus As String) As Boolean
    Dim sCmd            As String
    Dim baBuffer()      As Byte
    
    sCmd = STR_GS & "r" & Chr$(1)
    If Not m_oConnector.WriteData(ToAscii(sCmd), m_lTimeout) Then
        m_sLastError = m_oConnector.GetLastError()
        GoTo QH
    End If
    If Not m_oConnector.ReadData(baBuffer, m_lTimeout) Then
        m_sLastError = m_oConnector.GetLastError()
        GoTo QH
    End If
    CurrentStatus = vbNullString
    '--- success
    GetDeviceStatus = True
QH:
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
    Dim vElem           As Variant
    
    If IsArray(vText) Then
        For Each vElem In vText
            If Not m_oConnector.WriteData(ToAscii(C_Str(vElem) & vbLf, 866), m_lTimeout) Then
                m_sLastError = m_oConnector.GetLastError()
                GoTo QH
            End If
        Next
    Else
        If Not m_oConnector.WriteData(ToAscii(C_Str(vText) & vbLf, 866), m_lTimeout) Then
            m_sLastError = m_oConnector.GetLastError()
            GoTo QH
        End If
    End If
    '--- success
    pvPrintLine = True
QH:
End Function

Private Function pvSendPaperFeed(Optional ByVal Lines As Long = 1) As Boolean
    Dim sCmd            As String
    
    sCmd = STR_ESC & "d" & Chr$(Lines)
    If Not m_oConnector.WriteData(ToAscii(sCmd), m_lTimeout) Then
        m_sLastError = m_oConnector.GetLastError()
        GoTo QH
    End If
    '--- success
    pvSendPaperFeed = True
QH:
End Function

Private Function pvSendPaperCut(Optional ByVal Full As Boolean = True, Optional ByVal Lines As Long = 2) As Boolean
    Dim sCmd            As String
    
    sCmd = STR_GS & "V" & IIf(Full, "A" & Chr$(0), "0")
    If Lines > 0 Then
        sCmd = sCmd & STR_ESC & "d" & Chr$(Lines)
    End If
    If Not m_oConnector.WriteData(ToAscii(sCmd), m_lTimeout) Then
        m_sLastError = m_oConnector.GetLastError()
        GoTo QH
    End If
    '--- success
    pvSendPaperCut = True
QH:
End Function

Private Function pvPrintReceipt(uCtx As UcsContext, uRow() As UcsRowData, ByVal lRowCount As Long) As Boolean
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
            Case ucsFscRetNonfiscal, ucsFscRetInvoiceNonfiscal
                pvPrintLine CenterText(At(.InitOwnData, ucsOwnName), pvRowChars)
                pvPrintLine CenterText(At(.InitOwnData, ucsOwnAddress), pvRowChars)
                pvPrintLine CenterText(Printf(pvText(ucsTxtHeaderEic), At(.InitOwnData, ucsOwnBulstat)), pvRowChars)
                If LenB(At(.InitOwnData, ucsOwnDepName)) <> 0 Then
                    pvPrintLine CenterText(At(.InitOwnData, ucsOwnDepName), pvRowChars)
                End If
                If LenB(At(.InitOwnData, ucsOwnDepAddress)) <> 0 Then
                    pvPrintLine CenterText(At(.InitOwnData, ucsOwnDepAddress), pvRowChars)
                End If
                If .InitReceiptType = ucsFscRetInvoiceNonfiscal Then
                    sData = .InitInvData(ucsInvNumber)
                    If Len(sData) <= 10 Then
                        sData = Printf(pvText(ucsTxtHeaderInvoiceNo), sData)
                    End If
                    pvPrintLine CenterText(sData, pvRowChars)
                End If
            Case ucsFscRetOrderList
                pvPrintLine CenterText(pvText(ucsTxtOrderList), pvRowChars)
                pvPrintLine vbNullString
                pvPrintLine Printf(pvText(ucsTxtHeaderDate), Now())
                pvPrintLine Printf(pvText(ucsTxtHeaderOper), .InitOperatorName)
                pvPrintLine Printf(pvText(ucsTxtHeaderTable), .InitTableNo)
                pvPrintLine String(pvRowChars, "-")
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
            Case ucsFscRetNonfiscal, ucsFscRetInvoiceNonfiscal
                pvGetCommandConfig CONF_NAME, "NonFiscalSell" & .PrintRowType, 0, 0, vbNullString, vbNullString, _
                    "FormatPrice", sFormatPrice, "0.00", _
                    "FormatQty", sFormatQty, "0.000", _
                    "FormatDisc", sFormatDisc, "0.00", _
                    "FormatAmount", sFormatAmount, "0.00"
                If uCtx.PluCount = 1 Then
                    pvPrintLine AlignText(Format$(uRow(0).InitOperatorCode, "0000") & " " & uRow(0).InitOperatorName, Format$(uRow(0).InitCashDeskNo, "00"), pvRowChars)
                End If
                If .PrintRowType = ucsFscRetNonfiscal Then
                    vSplit = WrapMultiline(.PluName, pvRowChars - 7)
                    If .PluQuantity <> 1 Then
                        If UBound(vSplit) = 0 Then
                            ReDim Preserve vSplit(0 To 1) As Variant
                            vSplit(1) = vSplit(0)
                            vSplit(0) = vbNullString
                        End If
                        sData = " " & IIf(.PluPrice < -DBL_EPSILON, "-", vbNullString) & SafeFormat(.PluQuantity, sFormatQty) & " x " & SafeFormat(Abs(.PluPrice), sFormatPrice)
                        lIdx = UBound(vSplit) - 1
                        If Len(vSplit(lIdx)) + Len(sData) > pvRowChars - 2 Then
                            vTemp = WrapMultiline(vSplit(lIdx), pvRowChars - Len(sData) - 3)
                            vSplit(lIdx) = Trim$(At(vTemp, 0))
                            vSplit(UBound(vSplit)) = Trim$(At(vTemp, 1) & " " & vSplit(UBound(vSplit)))
                        End If
                        vSplit(lIdx) = AlignText(vSplit(lIdx), sData, pvRowChars - 2)
                    End If
                    vSplit(UBound(vSplit)) = AlignText(vSplit(UBound(vSplit)), " " & SafeFormat(dblTotal, sFormatAmount) & " " & Chr$(191 + .PluVatGroup), pvRowChars)
                    pvPrintLine vSplit
                Else
                    vSplit = WrapMultiline(.PluName, pvRowChars - 2)
                    For lIdx = 0 To UBound(vSplit)
                        pvPrintLine C_Str(vSplit(lIdx))
                    Next
                    pvPrintLine AlignText(pvText(ucsTxtPluPrice), SafeFormat(.PluPrice, sFormatPrice), pvRowChars - 2)
                    pvPrintLine AlignText(pvText(ucsTxtPluQuantity), SafeFormat(.PluQuantity, sFormatQty), pvRowChars - 2)
                    pvPrintLine AlignText(pvText(ucsTxtPluTotal), SafeFormat(dblTotal, sFormatAmount) & " " & Chr$(191 + .PluVatGroup), pvRowChars)
                End If
                Select Case .DiscType
                Case ucsFscDstPlu, ucsFscDstPluAbs
                    sData = AlignText(Printf(IIf(dblDisc > DBL_EPSILON, pvText(ucsTxtSurcharge), pvText(ucsTxtDiscount)), IIf(.DiscType = ucsFscDstPlu, SafeFormat(Abs(.DiscValue), sFormatDisc) & " %", vbNullString)), _
                        SafeFormat(dblDisc, sFormatAmount) & " " & Chr$(191 + .PluVatGroup), pvRowChars)
                    pvPrintLine sData
                End Select
            Case ucsFscRetOrderList
                pvGetCommandConfig CONF_NAME, "OrderList" & .PrintRowType, 0, 0, vbNullString, vbNullString, _
                    "FormatQty", sFormatQty, "0.000"
                vSplit = WrapMultiline(.PluName, pvRowChars - 7)
                If UBound(vSplit) = 0 Then
                    ReDim Preserve vSplit(0 To 1) As Variant
                End If
                sData = " " & IIf(.PluPrice < -DBL_EPSILON, "-", vbNullString) & SafeFormat(.PluQuantity, sFormatQty)
                lIdx = UBound(vSplit) - 1
                If Len(vSplit(lIdx)) + Len(sData) > pvRowChars - 2 Then
                    vTemp = WrapMultiline(vSplit(lIdx), pvRowChars - Len(sData) - 3)
                    vSplit(lIdx) = Trim$(At(vTemp, 0))
                    vSplit(UBound(vSplit)) = Trim$(At(vTemp, 1) & " " & vSplit(UBound(vSplit)))
                End If
                vSplit(lIdx) = AlignText(vSplit(lIdx), sData, pvRowChars)
                pvPrintLine vSplit
            End Select
        Case ucsRowDiscount
            Select Case .PrintRowType
            Case ucsFscRetNonfiscal, ucsFscRetInvoiceNonfiscal
                pvGetCommandConfig CONF_NAME, "NonFiscalSubtotal", 0, 0, vbNullString, vbNullString, _
                    "FormatDisc", sFormatDisc, "0.00", _
                    "FormatAmount", sFormatAmount, "0.00"
                Select Case .DiscType
                Case ucsFscDstSubtotal
                    pvPrintLine AlignText(pvText(ucsTxtSubTotal), SafeFormat(SumArray(uCtx.GrpTotal), sFormatAmount), pvRowChars - 2)
                    For lIdx = 1 To UBound(uCtx.GrpTotal)
                        If Abs(uCtx.GrpTotal(lIdx)) > DBL_EPSILON Then
                            dblTotal = Round(uCtx.GrpTotal(lIdx) * .DiscValue / 100#, 2)
                            pvPrintLine AlignText(Printf(IIf(dblTotal > DBL_EPSILON, pvText(ucsTxtSurcharge), pvText(ucsTxtDiscount)), SafeFormat(Abs(.DiscValue), sFormatDisc) & " %"), _
                                SafeFormat(dblTotal, sFormatAmount) & " " & Chr$(191 + lIdx), pvRowChars)
                        End If
                    Next
                Case ucsFscDstSubtotalAbs
                    '--- ToDo: impl absolute subtotal discount
                End Select
            End Select
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
                vSplit = WrapMultiline(.LineText, pvRowChars)
            Else
                vSplit = Array(Left$(.LineText, pvRowChars))
            End If
            Select Case .PrintRowType
            Case Else
                If uCtx.PmtPrinted And Not uCtx.ChangePrinted Then
                    pvGetCommandConfig CONF_NAME, "NonFiscalFooter", 0, 0, vbNullString, vbNullString, _
                        "FormatAmount", sFormatAmount, "0.00"
                    dblTotal = Round(SumArray(uCtx.GrpTotal) - uCtx.Paid, 2)
                    If dblTotal > DBL_EPSILON Then
                        pvPrintLine AlignText(GetPaymentName(1), SafeFormat(dblTotal, sFormatAmount), pvRowChars - 2)
                    ElseIf dblTotal < -DBL_EPSILON Then
                        pvPrintLine AlignText(pvText(ucsTxtFooterChange), SafeFormat(-dblTotal, sFormatAmount), pvRowChars - 2)
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
            Case ucsFscRetNonfiscal, ucsFscRetInvoiceNonfiscal
                pvGetCommandConfig CONF_NAME, "NonFiscalFooter", 0, 0, vbNullString, vbNullString, _
                    "FormatAmount", sFormatAmount, "0.00"
                If Not uCtx.PmtPrinted Then
                    If uCtx.PluCount > 0 Or .PmtType <> 0 Then
                        pvPrintLine String$(pvRowChars, "-")
                        pvPrintLine AlignText(pvText(ucsTxtFooterTotal), SafeFormat(SumArray(uCtx.GrpTotal), sFormatAmount), pvRowChars - 2)
                        If .PrintRowType = ucsFscRetInvoiceNonfiscal Then
                            '--- print invoice totals by VAT groups
                            vSplit = Split("0,20,20,9", DEF_DELIM)
                            For lIdx = 1 To UBound(uCtx.GrpTotal)
                                If Abs(uCtx.GrpTotal(lIdx)) > DBL_EPSILON Then
                                    '--- first, round VAT
                                    dblTotal = Round(uCtx.GrpTotal(lIdx) - uCtx.GrpTotal(lIdx) / (1 + C_Dbl(At(vSplit, lIdx - 1)) / 100#), 2)
                                    '--- then, calc TaxBase = Total - VAT
                                    dblTotal = Round(uCtx.GrpTotal(lIdx) - dblTotal, 2)
                                    pvPrintLine AlignText(Printf(pvText(ucsTxtGroupTotal), Chr$(191 + lIdx)), SafeFormat(uCtx.GrpTotal(lIdx), sFormatAmount), pvRowChars - 2)
                                    pvPrintLine AlignText(Printf(pvText(ucsTxtGroupVat), Chr$(191 + lIdx), SafeFormat(C_Dbl(At(vSplit, lIdx - 1)), sFormatAmount) & "%"), _
                                        SafeFormat(Round(uCtx.GrpTotal(lIdx) - dblTotal, 2), sFormatAmount), pvRowChars - 2)
                                    pvPrintLine AlignText(pvText(ucsTxtGroupNetto), SafeFormat(dblTotal, sFormatAmount), pvRowChars - 2)
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
                    pvPrintLine AlignText(.PmtName, SafeFormat(.PmtAmount, sFormatAmount), pvRowChars - 2)
                    uCtx.Paid = Round(uCtx.Paid + .PmtAmount, 2)
                Else
                    If uCtx.PmtPrinted And Not uCtx.ChangePrinted Then
                        dblTotal = Round(SumArray(uCtx.GrpTotal) - uCtx.Paid, 2)
                        If dblTotal > DBL_EPSILON Then
                            pvPrintLine AlignText(GetPaymentName(1), SafeFormat(dblTotal, sFormatAmount), pvRowChars - 2)
                        ElseIf dblTotal < -DBL_EPSILON Then
                            pvPrintLine AlignText(pvText(ucsTxtFooterChange), SafeFormat(-dblTotal, sFormatAmount), pvRowChars - 2)
                        End If
                        uCtx.ChangePrinted = True
                    End If
                    If .PrintRowType = ucsFscRetInvoiceNonfiscal Then
                        With uRow(0)
                            pvPrintLine AlignText(pvText(ucsTxtFooterSeller), STR_CHR1 & ".", pvRowChars)
                            pvPrintLine AlignText(vbNullString, "/" & .InitOperatorName & "/", pvRowChars)
                            pvPrintLine String$(pvRowChars, "-")
                            pvPrintLine AlignText(pvText(ucsTxtFooterReceiver), STR_CHR1 & ".", pvRowChars)
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
                    sData = At(uRow(0).InitOwnData, 5)
                    If LenB(sData) <> 0 Then
                        pvPrintLine CenterText(sData, pvRowChars)
                    End If
                    sData = At(uRow(0).InitOwnData, 6)
                    If LenB(sData) <> 0 Then
                        pvPrintLine CenterText(sData, pvRowChars)
                    End If
                    If uCtx.PluCount > 0 Then
                        sData = IIf(uCtx.PluCount = 1, pvText(ucsTxtFooterArticleSingle), Printf(pvText(ucsTxtFooterArticles), uCtx.PluCount))
                        pvPrintLine CenterText(sData, pvRowChars)
                    End If
                    vSplit = Split(Now, " ")
                    sData = AlignText(At(vSplit, 0), At(vSplit, 1), pvRowChars)
                    pvPrintLine sData
                    sData = CenterText(pvText(ucsTxtNonfiscalReceipt), pvRowChars)
                    pvPrintLine sData
                    If Not pvSendPaperFeed(Lines:=3) Then
                        GoTo QH
                    End If
                    If Not pvSendPaperCut() Then
                        GoTo QH
                    End If
                End If
            Case ucsFscRetOrderList
                If .PmtType = 0 Then
                    pvPrintLine String(pvRowChars, "-")
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
        pvSetLastError Printf(pvInternal(ucsErrErrorInLine), uCtx.PluCount, m_sLastError)
    End If
    Exit Function
EH:
    pvSetLastError Err.Description
    PrintError FUNC_NAME
    If uCtx.PluCount > 0 And Not uCtx.PmtPrinted Then
        pvSetLastError Printf(pvInternal(ucsErrErrorInLine), uCtx.PluCount, m_sLastError)
    End If
    Err.Raise vbObjectError, , m_sLastError
End Function

Private Sub pvGetCommandConfig( _
            sFunc As String, _
            sKey As String, _
            eCmd As UcsIclCommandsEnum, _
            ByVal eCmdDefault As UcsIclCommandsEnum, _
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

Private Sub pvSetLastError(sError As String)
    m_sLastError = sError
End Sub

Private Function SafeText(sText As String) As String
    Dim lIdx            As Long
    
    SafeText = sText
    For lIdx = 0 To 31
        Select Case lIdx
        Case 10, 13
        Case Else
            SafeText = Replace(SafeText, Chr$(lIdx), vbNullString)
        End Select
    Next
End Function

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Initialize()
    LocalizedText(ucsFscLciInternalErrors) = STR_INTERNAL
    Set m_uConfig.ConfigCommands = GetConfigCollection(STR_PROTOCOL_ESP_POS, "Commands")
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
    
End Function

Private Function IDeviceProtocol_Connect() As Boolean
    IDeviceProtocol_Connect = Connect()
End Function

Private Function IDeviceProtocol_CopyLastReceipt(Optional ReceiptNumber As String) As Boolean
    
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
    
End Function

Private Function IDeviceProtocol_GetDefaultPassword(OperatorCode As String) As String
    
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

Private Function IDeviceProtocol_GetLastReceiptNumber() As String

End Function

Private Function IDeviceProtocol_GetLastError(Optional ErrCategory As UcsFiscalErrorsEnum) As String
    ErrCategory = LastErrNumber
    IDeviceProtocol_GetLastError = LastError
End Function

Private Function IDeviceProtocol_GetLastInvoiceNumber() As String

End Function

Private Function IDeviceProtocol_GetPaymentName(ByVal Index As Long) As String
    IDeviceProtocol_GetPaymentName = GetPaymentName(Index)
End Function

Private Function IDeviceProtocol_GetResumeTicket() As String
    
End Function

Private Function IDeviceProtocol_GetSerialNumber() As String

End Function

Private Function IDeviceProtocol_GetTaxCaption() As String

End Function

Private Function IDeviceProtocol_GetTaxNumber() As String

End Function

Private Function IDeviceProtocol_GetTotalsByPayments() As Variant
    
End Function

Private Function IDeviceProtocol_GetTotalsByVatGroups() As Variant
    
End Function

Private Function IDeviceProtocol_GetTotalsStats() As Variant
    
End Function

Private Function IDeviceProtocol_Init(Device As String, Optional ByVal Timeout As Long, Optional ByVal CashDeskNo As Long, Optional ByVal RowChars As Long) As Boolean
    IDeviceProtocol_Init = Init(Device, Timeout, CashDeskNo, RowChars)
End Function

Private Property Get IDeviceProtocol_IsConnected() As Boolean
    IDeviceProtocol_IsConnected = IsConnected
End Property

Private Function IDeviceProtocol_OpenDrawer() As Boolean
    IDeviceProtocol_OpenDrawer = OpenDrawer()
End Function

Private Function IDeviceProtocol_RunPeriodReport(ByVal ReportType As UcsFiscalReportsTypeEnum, ByVal StartDate As Date, ByVal EndDate As Date) As Boolean
    
End Function

Private Function IDeviceProtocol_RunXReport(ByVal ReportType As UcsFiscalReportsTypeEnum) As Boolean
    
End Function

Private Function IDeviceProtocol_RunZReport(ByVal ReportType As UcsFiscalReportsTypeEnum) As Boolean
    
End Function

Private Function IDeviceProtocol_SetClock(ByVal NewDate As Date) As Boolean
    IDeviceProtocol_SetClock = True
End Function

Private Sub IDeviceProtocol_SetLocalizedText(ByVal Index As UcsFiscalLocalizedIndexesEnum, Text As String)
    LocalizedText(Index) = Text
End Sub

Private Function IDeviceProtocol_StartReceipt(ByVal ReceiptType As UcsFiscalReceiptTypeEnum, OperatorCode As String, OperatorName As String, OperatorPassword As String, Optional TableNo As String, Optional InvNumber As String, Optional InvTaxNumber As String, Optional InvVatNumber As String, Optional InvCompany As String, Optional InvCity As String, Optional InvAddress As String, Optional InvMol As String, Optional InvReceiver As String, Optional OwnData As String) As Boolean
    IDeviceProtocol_StartReceipt = StartReceipt(ReceiptType, OperatorCode, OperatorName, OperatorPassword, TableNo, InvNumber, InvTaxNumber, InvVatNumber, InvCompany, InvCity, InvAddress, InvMol, InvReceiver, OwnData)
End Function

