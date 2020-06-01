VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cProxyProtocol"
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
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "cProxyProtocol"
Implements IDeviceProtocol

'=========================================================================
' Constant and member variables
'=========================================================================

Private Const STR_INTERNAL              As String = "Device not found|Invalid ready state %1|Invalid JSON: %1|Unsupported discount type %1"
Private Const DEF_ROW_CHARS             As Long = 32
Private Const DEF_ITEM_CHARS            As Long = 72
Private Const DEF_MIN_DISCOUNT          As Double = -100
Private Const DEF_MAX_DISCOUNT          As Double = 100
Private Const DEF_MAX_RECEIPT_ROWS      As Long = 1000
Private Const DEF_MAX_ITEM_LINES        As Long = 2
Private Const DEF_HEADERS               As String = "Наименование на задълженото лице|Адрес на задълженото лице|Наименование на обект|Адрес на обект"
Private Const DEF_FOOTERS               As String = "Заповядайте отново!"
Private Const DEF_PMTNAMES              As String = "В брой|С чек|Талони|В.талони|Амбалаж|В.обслужване|Повреди|С карта|По банкa|НЗОК|Прог.тип 2"
Private Const DEF_TAXCAPTION            As String = "ЕИК"
Private Const DEF_TAXNO                 As String = "Незададен"
Private Const DEF_TAXRATES              As String = "0|20|20|9"
'--- numeric
Private Const READYSTATE_COMPLETE       As Long = 4

Private m_uPrintData                As UcsProtocolPrintData
Private m_uConfig                   As UcsConfigValues
Private m_uRemote                   As UcsParsedUrl
Private m_sLastReceiptNo            As String
Private m_dLastReceiptDate          As Date
Private m_dblLastReceiptAmount      As Double
Private m_sLastInvoiceNo            As String
Private m_sResumeToken              As String

Private Enum UcsInternalErrors
    ucsErrDeviceNotFound
    ucsErrInvalidReadyState
    ucsErrInvalidJSON
    ucsErrUnsupportedDiscType
End Enum

Private Type UcsConfigValues
    DefaultPassword     As String
    DeviceSerialNo      As String
    FiscalMemoryNo      As String
    DeviceModel         As String
    FirmwareVersion     As String
    TaxNo               As String
    TaxCaption          As String
    Headers             As Variant
    Footers             As Variant
    PaymentNames        As Variant
    TaxRates            As Variant
    LocalizedText(0 To [_ucsFscLciMax] - 1) As Variant
End Type

'=========================================================================
' Error handling
'=========================================================================

'Private Sub PrintError(sFunction As String)
'    Debug.Print "Critical error: " & Err.Description & " [" & MODULE_NAME & "." & sFunction & "]"
'    Logger.Log vbLogEventTypeError, MODULE_NAME, sFunction & "(" & Erl & ")", Err.Description
'End Sub

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

Property Get LastError() As String
    LastError = m_uPrintData.LastError
End Property

Property Get LocalizedText(ByVal eIdx As UcsFiscalLocalizedIndexesEnum) As String
    If IsArray(m_uConfig.LocalizedText(eIdx)) Then
        LocalizedText = Join(m_uConfig.LocalizedText(eIdx), "|")
    End If
End Property

Property Let LocalizedText(ByVal eIdx As UcsFiscalLocalizedIndexesEnum, sValue As String)
    m_uConfig.LocalizedText(eIdx) = Split(sValue, "|")
End Property

'= private ===============================================================

Private Property Get pvRowChars() As Long
    If m_uPrintData.Config.RowChars = 0 Then
        m_uPrintData.Config.RowChars = DEF_ROW_CHARS
    End If
    pvRowChars = m_uPrintData.Config.RowChars
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
        pvItemChars = DEF_ITEM_CHARS
    End If
End Property

Private Property Get pvBaseAddress() As String
    pvBaseAddress = m_uRemote.Protocol & "://" & m_uRemote.Host & ":" & m_uRemote.Port & m_uRemote.Path
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function Init(sDeviceString As String) As Boolean
    Const FUNC_NAME     As String = "Init"
    Dim oOptions        As Object
    Dim sDeviceSerialNo As String
    Dim oResponse       As Object
    
    On Error GoTo EH
    DebugLog FUNC_NAME, "sDeviceString=" & sDeviceString
    Set oOptions = ParseDeviceString(sDeviceString)
    If ParseUrl(C_Str(JsonItem(oOptions, "Url")), m_uRemote) Then
        pvSendRequest "GET", vbNullString, Nothing, oResponse
    End If
    sDeviceSerialNo = Zn(C_Str(JsonItem(oResponse, "DeviceSerialNo")), C_Str(JsonItem(oOptions, "DeviceSerialNo")))
    If LenB(sDeviceSerialNo) = 0 Then
        pvSetLastError pvInternal(ucsErrDeviceNotFound)
        GoTo QH
    End If
    With m_uPrintData.Config
        .RowChars = GetConfigNumber(sDeviceSerialNo, "RowChars", C_Lng(JsonItem(oOptions, "RowChars")))
        .CommentChars = GetConfigNumber(sDeviceSerialNo, "CommentChars", C_Lng(JsonItem(oOptions, "CommentChars")))
        .ItemChars = GetConfigNumber(sDeviceSerialNo, "ItemChars", C_Lng(JsonItem(oOptions, "ItemChars")))
        .AbsoluteDiscount = C_Bool(GetConfigValue(sDeviceSerialNo, "AbsoluteDiscount", True))
        .NegativePrices = C_Bool(GetConfigValue(sDeviceSerialNo, "NegativePrices", True))
        .MinDiscount = GetConfigNumber(sDeviceSerialNo, "MinDiscount", Zndbl(C_Dbl(JsonItem(oOptions, "MinDiscount")), DEF_MIN_DISCOUNT))
        .MaxDiscount = GetConfigNumber(sDeviceSerialNo, "MaxDiscount", Zndbl(C_Dbl(JsonItem(oOptions, "MaxDiscount")), DEF_MAX_DISCOUNT))
        .MaxReceiptRows = GetConfigNumber(sDeviceSerialNo, "MaxReceiptRows", Znl(C_Lng(JsonItem(oOptions, "MaxReceiptRows")), DEF_MAX_RECEIPT_ROWS))
        .MaxItemLines = GetConfigNumber(sDeviceSerialNo, "MaxItemLines", Znl(C_Lng(JsonItem(oOptions, "MaxItemLines")), DEF_MAX_ITEM_LINES))
        '--- from response
        .RowChars = Znl(C_Lng(JsonItem(oResponse, "CharsPerLine")), .RowChars)
        .CommentChars = Znl(C_Lng(JsonItem(oResponse, "CommentTextMaxLength")), .CommentChars)
        .ItemChars = Znl(C_Lng(JsonItem(oResponse, "ItemNameMaxLength")), .ItemChars)
    End With
    With m_uConfig
        .DefaultPassword = GetConfigValue(sDeviceSerialNo, "DefaultPassword", JsonItem(oOptions, "DefaultPassword"))
        .DeviceSerialNo = Zn(C_Str(JsonItem(oResponse, "DeviceSerialNo")), C_Str(JsonItem(oOptions, "DeviceSerialNo")))
        .FiscalMemoryNo = Zn(C_Str(JsonItem(oResponse, "FiscalMemoryNo")), C_Str(JsonItem(oOptions, "FiscalMemoryNo")))
        .DeviceModel = Zn(C_Str(JsonItem(oResponse, "DeviceModel")), "Proxy")
        .FirmwareVersion = Zn(C_Str(JsonItem(oResponse, "FirmwareVersion")), "N/A")
        .TaxNo = Zn(C_Str(JsonItem(oResponse, "TaxNo")), Zn(C_Str(JsonItem(oOptions, "TaxNo")), DEF_TAXNO))
        .TaxCaption = Zn(C_Str(JsonItem(oResponse, "TaxCaption")), Zn(C_Str(JsonItem(oOptions, "TaxCaption")), DEF_TAXCAPTION))
        .Headers = JsonItem(oResponse, "Headers/*")
        If IsMissing(.Headers) Or Not IsArray(.Headers) Then
            .Headers = Split(Zn(C_Str(JsonItem(oOptions, "Headers")), DEF_HEADERS), "|")
        End If
        .Footers = JsonItem(oResponse, "Footers/*")
        If IsMissing(.Footers) Or Not IsArray(.Footers) Then
            .Footers = Split(Zn(C_Str(JsonItem(oOptions, "Footers")), DEF_FOOTERS), "|")
        End If
        .PaymentNames = JsonItem(oResponse, "PaymentNames/*")
        If IsMissing(.PaymentNames) Or Not IsArray(.PaymentNames) Then
            .PaymentNames = Split(Zn(C_Str(JsonItem(oOptions, "PaymentNames")), DEF_PMTNAMES), "|")
        End If
        .TaxRates = JsonItem(oResponse, "TaxRates/*")
        If IsMissing(.TaxRates) Or Not IsArray(.TaxRates) Then
            .TaxRates = Split(Zn(C_Str(JsonItem(oOptions, "TaxRates")), DEF_TAXRATES), "|")
        End If
    End With
    '--- success
    Init = True
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetTaxRates() As Variant
    Const FUNC_NAME     As String = "GetTaxRates"
    Dim vRetVal(0 To MAX_TAX_GROUP - 1) As Double
    Dim lIdx            As Long
    
    On Error GoTo EH
    For lIdx = 0 To MAX_TAX_GROUP - 1
        vRetVal(lIdx) = C_Dbl(At(m_uConfig.TaxRates, lIdx, -1))
    Next
    GetTaxRates = vRetVal
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function CancelReceipt() As Boolean
    Const FUNC_NAME     As String = "CancelReceipt"
    Dim oRequest        As Object
    
    On Error GoTo EH
    If LenB(m_uRemote.Host) = 0 Then
        CancelReceipt = True
        GoTo QH
    End If
    JsonItem(oRequest, vbNullString) = Empty
    If Not pvSendRequest("POST", "/receipt", oRequest, Nothing) Then
         GoTo QH
    End If
    '--- success
    CancelReceipt = True
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function PrintServiceDeposit( _
            OperatorCode As String, _
            OperatorPassword As String, _
            ByVal Amount As Double, _
            Optional ReceiptNo As String, _
            Optional ReceiptDate As Date, _
            Optional ReceiptAmount As Double) As Variant
    Const FUNC_NAME     As String = "PrintServiceDeposit"
    Dim vRetVal(0 To 2) As Double
    Dim oRequest        As Object
    Dim oResponse       As Object
    
    On Error GoTo EH
    If LenB(m_uRemote.Host) = 0 Then
        m_dLastReceiptDate = GetCurrentNow
        PrintServiceDeposit = vRetVal
        GoTo QH
    End If
    If LenB(OperatorCode) <> 0 Then
        JsonItem(oRequest, "Operator/Code") = OperatorCode
    End If
    If LenB(OperatorPassword) <> 0 Then
        JsonItem(oRequest, "Operator/Password") = OperatorPassword
    End If
    If Abs(Amount) > DBL_EPSILON Then
        JsonItem(oRequest, "Amount") = Amount
    End If
    If Not pvSendRequest(IIf(oRequest Is Nothing, "GET", "POST"), "/deposit", oRequest, oResponse) Then
        GoTo QH
    End If
    If Not IsEmpty(JsonItem(oResponse, "ReceiptNo")) Then
        m_sLastReceiptNo = C_Str(JsonItem(oResponse, "ReceiptNo"))
        m_dLastReceiptDate = C_Date(JsonItem(oResponse, "ReceiptDateTime"))
        m_dblLastReceiptAmount = C_Dbl(JsonItem(oResponse, "ReceiptAmount"))
        ReceiptNo = m_sLastReceiptNo
        ReceiptDate = m_dLastReceiptDate
        ReceiptAmount = m_dblLastReceiptAmount
    End If
    vRetVal(0) = C_Dbl(JsonItem(oResponse, "TotalAvailable"))
    vRetVal(1) = C_Dbl(JsonItem(oResponse, "TotalDeposits"))
    vRetVal(2) = C_Dbl(JsonItem(oResponse, "TotalWithdraws"))
    '--- success
    PrintServiceDeposit = vRetVal
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function PrintDuplicate(Optional DuplicateNo As String) As Boolean
    Const FUNC_NAME     As String = "PrintDuplicate"
    Dim oRequest        As Object
    Dim oResponse       As Object
    
    On Error GoTo EH
    JsonItem(oRequest, "PrintDuplicate") = True
    If LenB(DuplicateNo) <> 0 Then
        JsonItem(oRequest, "Invoice/DocNo") = DuplicateNo
    End If
    If Not pvSendRequest("POST", "/receipt", oRequest, oResponse) Then
        GoTo QH
    End If
    m_sLastReceiptNo = C_Str(JsonItem(oResponse, "ReceiptNo"))
    m_dLastReceiptDate = C_Date(JsonItem(oResponse, "ReceiptDateTime"))
    m_dblLastReceiptAmount = C_Dbl(JsonItem(oResponse, "ReceiptAmount"))
    '--- success
    PrintDuplicate = True
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function PrintMonthlyReport(ByVal ReportType As UcsFiscalReportsTypeEnum, ByVal StartDate As Date, ByVal EndDate As Date) As Boolean
    Const FUNC_NAME     As String = "PrintMonthlyReport"
    Dim oRequest        As Object
    Dim oResponse       As Object
    
    On Error GoTo EH
    JsonItem(oRequest, "ReportType") = ucsFscRptMonthlyByDate
    JsonItem(oRequest, "StartDate") = StartDate
    JsonItem(oRequest, "EndDate") = EndDate
    If ReportType = ucsFscRstMonthlyDetailed Then
        JsonItem(oRequest, "IsDetailed") = True
    End If
    If Not pvSendRequest("POST", "/report", oRequest, oResponse) Then
        GoTo QH
    End If
    m_sLastReceiptNo = C_Str(JsonItem(oResponse, "ReceiptNo"))
    m_dLastReceiptDate = C_Date(JsonItem(oResponse, "ReceiptDateTime"))
    m_dblLastReceiptAmount = C_Dbl(JsonItem(oResponse, "ReceiptAmount"))
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function PrintDailyReport(ByVal IsClear As Boolean, ByVal ReportType As UcsFiscalReportsTypeEnum) As Boolean
    Const FUNC_NAME     As String = "PrintDailyReport"
    Dim oRequest        As Object
    Dim oResponse       As Object
    
    On Error GoTo EH
    JsonItem(oRequest, "ReportType") = IIf(ReportType = ucsFscRstDailyByOperators, ucsFscRptDailyByOperators, ucsFscRptDaily)
    JsonItem(oRequest, "IsClear") = IsClear
    If ReportType = ucsFscRstDailyByItems Or ReportType = ucsFscRstDailyByItemsAndDepartments Then
        JsonItem(oRequest, "IsItems") = True
    End If
    If ReportType = ucsFscRstDailyByDepartments Or ReportType = ucsFscRstDailyByItemsAndDepartments Then
        JsonItem(oRequest, "IsDepartments") = True
    End If
    If Not pvSendRequest("POST", "/report", oRequest, oResponse) Then
        GoTo QH
    End If
    m_sLastReceiptNo = C_Str(JsonItem(oResponse, "ReceiptNo"))
    m_dLastReceiptDate = C_Date(JsonItem(oResponse, "ReceiptDateTime"))
    m_dblLastReceiptAmount = C_Dbl(JsonItem(oResponse, "ReceiptAmount"))
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetDefaultPassword(OperatorCode As String) As String
    Const FUNC_NAME     As String = "GetDefaultPassword"
    Dim oRequest        As Object
    Dim oResponse       As Object
    
    On Error GoTo EH
    If LenB(m_uRemote.Host) = 0 Then
        If LenB(m_uConfig.DefaultPassword) <> 0 Then
            GetDefaultPassword = m_uConfig.DefaultPassword
        Else
            GetDefaultPassword = OperatorCode
        End If
        GoTo QH
    End If
    JsonItem(oRequest, "Operator/Code") = OperatorCode
    If Not pvSendRequest("POST", vbNullString, oRequest, oResponse) Then
        GoTo QH
    End If
    GetDefaultPassword = C_Str(JsonItem(oResponse, "Operator/DefaultPassword"))
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetClock() As Date
    Const FUNC_NAME     As String = "GetClock"
    Dim oResponse       As Object
    
    On Error GoTo EH
    If LenB(m_uRemote.Host) = 0 Then
        GetClock = GetCurrentNow
        GoTo QH
    End If
    If Not pvSendRequest("GET", "/datetime", Nothing, oResponse) Then
        GoTo QH
    End If
    GetClock = C_Date(JsonItem(oResponse, "DeviceDateTime"))
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function SetClock(ByVal NewDate As Date) As Boolean
    Const FUNC_NAME     As String = "SetClock"
    Dim oRequest        As Object
    Dim oResponse       As Object
    
    On Error GoTo EH
    If LenB(m_uRemote.Host) = 0 Then
        SetClock = True
        GoTo QH
    End If
    JsonItem(oRequest, "DeviceDateTime") = NewDate
    If Not pvSendRequest("POST", "/datetime", oRequest, oResponse) Then
        GoTo QH
    End If
    '--- success
    SetClock = True
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetDeviceStatus(StatusText As String) As UcsFiscalStatusCodeEnum
    Const FUNC_NAME     As String = "GetDeviceStatus"
    Dim oResponse       As Object
    
    On Error GoTo EH
    If LenB(m_uRemote.Host) = 0 Then
        GoTo QH
    End If
    If Not pvSendRequest("GET", "/status", Nothing, oResponse) Then
        StatusText = LastError
        GetDeviceStatus = ucsFscStaFailed
        GoTo QH
    End If
    StatusText = C_Str(JsonItem(oResponse, "DeviceStatus"))
    GetDeviceStatus = JsonEnumValue(JsonItem(oResponse, "DeviceStatusCode"), STR_ENUM_STATUS_CODE)
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetTotalsByPaymentTypes() As Variant
    Const FUNC_NAME     As String = "GetTotalsByPaymentTypes"
    Dim oRequest        As Object
    Dim oResponse       As Object
    Dim vRetVal         As Variant
    Dim lIdx            As Long
    
    On Error GoTo EH
    If LenB(m_uRemote.Host) = 0 Then
        GoTo QH
    End If
    JsonItem(oRequest, "IncludePaymentTypes") = True
    If Not pvSendRequest("POST", "/totals", oRequest, oResponse) Then
        GoTo QH
    End If
    ReDim vRetVal(0 To JsonItem(oResponse, -1) - 1) As Variant
    For lIdx = 0 To UBound(vRetVal)
        If IsEmpty(JsonItem(oResponse, lIdx & "/Reversal")) Then
            vRetVal(lIdx) = Array(JsonItem(oResponse, lIdx & "/PaymentName"), C_Dbl(JsonItem(oResponse, lIdx & "/Amount")))
        Else
            vRetVal(lIdx) = Array(JsonItem(oResponse, lIdx & "/PaymentName"), C_Dbl(JsonItem(oResponse, lIdx & "/Amount")), C_Dbl(JsonItem(oResponse, lIdx & "/Reversal")))
        End If
    Next
    GetTotalsByPaymentTypes = vRetVal
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetTotalsByTaxGroups() As Variant
    Const FUNC_NAME     As String = "GetTotalsByTaxGroups"
    Dim oRequest        As Object
    Dim oResponse       As Object
    Dim vRetVal         As Variant
    Dim lIdx            As Long
    
    On Error GoTo EH
    If LenB(m_uRemote.Host) = 0 Then
        GoTo QH
    End If
    JsonItem(oRequest, "IncludeTaxGroups") = True
    If Not pvSendRequest("POST", "/totals", oRequest, oResponse) Then
        GoTo QH
    End If
    ReDim vRetVal(0 To JsonItem(oResponse, -1) - 1) As Variant
    For lIdx = 0 To UBound(vRetVal)
        If IsEmpty(JsonItem(oResponse, lIdx & "/Reversal")) Then
            vRetVal(lIdx) = Array(C_Dbl(JsonItem(oResponse, lIdx & "/VatRate")), C_Dbl(JsonItem(oResponse, lIdx & "/Amount")))
        Else
            vRetVal(lIdx) = Array(C_Dbl(JsonItem(oResponse, lIdx & "/VatRate")), C_Dbl(JsonItem(oResponse, lIdx & "/Amount")), C_Dbl(JsonItem(oResponse, lIdx & "/Reversal")))
        End If
    Next
    GetTotalsByTaxGroups = vRetVal
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetTotalsStats() As Variant
    Const FUNC_NAME     As String = "GetTotalsByPaymentTypes"
    Dim oRequest        As Object
    Dim oResponse       As Object
    
    On Error GoTo EH
    If LenB(m_uRemote.Host) = 0 Then
        GoTo QH
    End If
    JsonItem(oRequest, "IncludeStats") = True
    If Not pvSendRequest("POST", "/totals", oRequest, oResponse) Then
        GoTo QH
    End If
    GetTotalsStats = Array(C_Lng(JsonItem(oResponse, "NumReceipts")), C_Date(JsonItem(oResponse, "LastZReportDateTime")))
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

'= private ===============================================================

Private Function pvPrintReceipt(uCtx As UcsPpdExecuteContext, uRow() As UcsPpdRowData, ByVal lRowCount As Long) As Boolean
    Const FUNC_NAME     As String = "pvPrintReceipt"
    Dim oRequest        As Object
    Dim oResponse       As Object
    Dim oItem           As Object
    
    On Error GoTo EH
    If LenB(m_uRemote.Host) = 0 Then
        m_dLastReceiptDate = GetCurrentNow
        pvPrintReceipt = True
        GoTo QH
    End If
    Do While uCtx.Row < lRowCount
        With uRow(uCtx.Row)
        Select Case .RowType
        Case ucsRowInit
            JsonItem(oRequest, "ReceiptType") = .InitReceiptType
            JsonItem(oRequest, "Operator/Code") = .InitOperatorCode
            JsonItem(oRequest, "Operator/Name") = .InitOperatorName
            JsonItem(oRequest, "Operator/Password") = .InitOperatorPassword
            Select Case .InitReceiptType
            Case ucsFscRcpInvoice
                JsonItem(oRequest, "Invoice/DocNo") = .InitInvData(ucsInvDocNo)
                JsonItem(oRequest, "Invoice/CgTaxNo") = .InitInvData(ucsInvCgTaxNo)
                JsonItem(oRequest, "Invoice/CgVatNo") = .InitInvData(ucsInvCgVatNo)
                JsonItem(oRequest, "Invoice/CgName") = .InitInvData(ucsInvCgName)
                JsonItem(oRequest, "Invoice/CgAddress") = .InitInvData(ucsInvCgAddress)
                JsonItem(oRequest, "Invoice/CgPrsReceive") = .InitInvData(ucsInvCgPrsReceive)
            Case ucsFscRcpReversal
                JsonItem(oRequest, "Reversal/ReversalType") = .InitRevData(ucsRevType)
                JsonItem(oRequest, "Reversal/ReceiptNo") = .InitRevData(ucsRevReceiptNo)
                JsonItem(oRequest, "Reversal/ReceiptDateTime") = .InitRevData(ucsRevReceiptDate)
                JsonItem(oRequest, "Reversal/FiscalMemoryNo") = .InitRevData(ucsRevFiscalMemoryNo)
            Case ucsFscRcpCreditNote
                JsonItem(oRequest, "Invoice/DocNo") = .InitInvData(ucsInvDocNo)
                JsonItem(oRequest, "Invoice/CgTaxNo") = .InitInvData(ucsInvCgTaxNo)
                JsonItem(oRequest, "Invoice/CgVatNo") = .InitInvData(ucsInvCgVatNo)
                JsonItem(oRequest, "Invoice/CgName") = .InitInvData(ucsInvCgName)
                JsonItem(oRequest, "Invoice/CgAddress") = .InitInvData(ucsInvCgAddress)
                JsonItem(oRequest, "Invoice/CgPrsReceive") = .InitInvData(ucsInvCgPrsReceive)
                JsonItem(oRequest, "Reversal/ReversalType") = .InitRevData(ucsRevType)
                JsonItem(oRequest, "Reversal/ReceiptNo") = .InitRevData(ucsRevReceiptNo)
                JsonItem(oRequest, "Reversal/ReceiptDateTime") = .InitRevData(ucsRevReceiptDate)
                JsonItem(oRequest, "Reversal/FiscalMemoryNo") = .InitRevData(ucsRevFiscalMemoryNo)
                JsonItem(oRequest, "Reversal/InvoiceNo") = .InitRevData(ucsRevInvoiceNo)
                JsonItem(oRequest, "Reversal/Reason") = .InitRevData(ucsRevReason)
            End Select
            JsonItem(oRequest, "UniqueSaleNo") = .InitUniqueSaleNo
        Case ucsRowPlu
            JsonItem(oRequest, "Rows/-1") = Array(.PluItemName, .PluPrice, .PluTaxGroup, .PluQuantity, .DiscValue, .PluUnitOfMeasure, .PluDepartmentNo)
        Case ucsRowDiscount
            Set oItem = Nothing
            If .DiscType = ucsFscDscSubtotal Then
                JsonItem(oItem, "DiscountPercent") = .DiscValue
            ElseIf .DiscType = ucsFscDscSubtotalAbs Then
                JsonItem(oItem, "DiscountAmount") = .DiscValue
            Else
                pvSetLastError Printf(pvInternal(ucsErrUnsupportedDiscType), .DiscType)
                GoTo QH
            End If
            JsonItem(oRequest, "Rows/-1") = oItem
        Case ucsRowLine
            JsonItem(oRequest, "Rows/-1") = Array(.LineText)
        Case ucsRowBarcode
            Set oItem = Nothing
            JsonItem(oItem, "Text") = .BarcodeText
            JsonItem(oItem, "BarcodeType") = .BarcodeType
            JsonItem(oItem, "Height") = .BarcodeHeight
            JsonItem(oRequest, "Rows/-1") = oItem
        Case ucsRowPayment
            If .PmtType > 0 Then
                JsonItem(oRequest, "Rows/-1") = Array(.PmtName, .PmtType, .PmtAmount)
            End If
        End Select
        End With
        uCtx.Row = uCtx.Row + 1
    Loop
    If Not pvSendRequest("POST", "/receipt", oRequest, oResponse) Then
        If JsonItem(oResponse, "ErrorCode") = "InvalidPassword" Then
            m_uPrintData.LastErrNo = ucsFerInvalidPassword
        End If
        GoTo QH
    End If
    m_sLastReceiptNo = C_Str(JsonItem(oResponse, "ReceiptNo"))
    m_dLastReceiptDate = C_Date(JsonItem(oResponse, "ReceiptDateTime"))
    m_dblLastReceiptAmount = C_Dbl(JsonItem(oResponse, "ReceiptAmount"))
    m_sLastInvoiceNo = C_Str(JsonItem(oResponse, "LastInvoiceNo"))
    m_sResumeToken = C_Str(JsonItem(oResponse, "ResumeToken"))
    '--- success
    pvPrintReceipt = True
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Function pvSendRequest( _
            sVerb As String, _
            sPath As String, _
            oRequest As Object, _
            oResponse As Object, _
            Optional Error As String) As Boolean
    Const FUNC_NAME     As String = "pvSendRequest"
    Const DBL_TIMEOUT   As Double = 30
    Const Async         As Boolean = False
    Const USER_AGENT    As String = "Proxy Protocol"
    Dim dblTimerEx      As Double
    Dim sPayload        As String
    Dim vJson           As Object
    
    On Error GoTo EH
    dblTimerEx = TimerEx
    DebugLog FUNC_NAME, "sVerb=" & sVerb & ", sPath=" & sPath
    If LenB(m_uRemote.Host) = 0 Then
        GoTo QH
    End If
    '--- note: Timeout:=30 forces MSXML2.XMLHTTP usage
    With InitRequest(sVerb, pvBaseAddress & sPath & IIf(sVerb = "GET", "?ts=" & TimerEx, vbNullString), Timeout:=DBL_TIMEOUT, Async:=Async)
        .SetRequestHeader "User-Agent", USER_AGENT & " " & LIB_NAME & "/" & App.Major & "." & App.Minor & "." & App.Revision
        .SetRequestHeader "Content-Type", "application/json"
        If LenB(m_uRemote.User) <> 0 Then
            .SetRequestHeader m_uRemote.User, m_uRemote.Pass
        End If
        If Not oRequest Is Nothing Then
            sPayload = JsonDump(oRequest, Minimize:=True)
            DebugDataDump FUNC_NAME, "[SEND ] ", sPayload
            .Send sPayload
        Else
            .Send
        End If
        If .ReadyState <> READYSTATE_COMPLETE Then
            Error = Printf(pvInternal(ucsErrInvalidReadyState), .ReadyState)
            GoTo QH
        ElseIf .Status < 200 Or .Status > 299 Then
            Error = .Status & " " & .StatusText
            GoTo QH
        End If
        DebugDataDump FUNC_NAME, "[RECV ] ", .ResponseText
        If Not JsonParse(.ResponseText, vJson, Error:=Error) Then
            DebugLog FUNC_NAME, "Error parsing ResponseText=" & .ResponseText & ", Error=" & Error
            Error = Printf(pvInternal(ucsErrInvalidJSON), Error)
            GoTo QH
        End If
    End With
    Set oResponse = C_Obj(vJson)
    If Not C_Bool(JsonItem(oResponse, "Ok")) Then
        DebugLog FUNC_NAME, "Request failed, vJson=" & JsonDump(vJson, Minimize:=True)
        Error = C_Str(JsonItem(oResponse, "ErrorText"))
        GoTo QH
    End If
    Error = vbNullString
    '--- success
    pvSendRequest = True
QH:
    If LenB(Error) <> 0 Or LenB(JsonItem(oResponse, "ErrorCode")) <> 0 Then
        pvSetLastError Error, JsonItem(oResponse, "ErrorCode")
    End If
    DebugLog FUNC_NAME, "RetVal=" & pvSendRequest & ", Error=" & Error & ", Elapsed=" & Format$(TimerEx - dblTimerEx, "0.000")
    Exit Function
EH:
    Error = Err.Description
    Resume QH
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

Private Sub pvSetLastError(sError As String, Optional ErrorCode As String)
    Const FUNC_NAME     As String = "pvSetLastError"
    
    On Error GoTo EH
    If LenB(ErrorCode) = 0 Then
        m_uPrintData.LastErrNo = IIf(LenB(sError) = 0, ucsFerNone, ucsFerGeneralError)
    ElseIf ErrorCode = "InvalidPassword" Then
        m_uPrintData.LastErrNo = ucsFerInvalidPassword
    ElseIf ErrorCode = "PasswordNotSet" Then
        m_uPrintData.LastErrNo = ucsFerPasswordNotSet
    End If
    m_uPrintData.LastError = Zn(sError, ErrorCode)
    Exit Sub
EH:
    RaiseError FUNC_NAME
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
    
End Function

Private Function IDeviceProtocol_CancelReceipt() As Boolean
    IDeviceProtocol_CancelReceipt = CancelReceipt()
End Function

Private Function IDeviceProtocol_PrintServiceDeposit(OperatorCode As String, OperatorPassword As String, ByVal Amount As Double, Optional ReceiptNo As String, Optional ReceiptDate As Date, Optional ReceiptAmount As Double) As Variant
    IDeviceProtocol_PrintServiceDeposit = PrintServiceDeposit(OperatorCode, OperatorPassword, Amount, ReceiptNo, ReceiptDate, ReceiptAmount)
End Function

Private Function IDeviceProtocol_Connect() As Boolean
    IDeviceProtocol_Connect = True
End Function

Private Function IDeviceProtocol_PrintDuplicate(Optional DuplicateNo As String, Optional ReceiptNo As String, Optional ReceiptDate As Date) As Boolean
    IDeviceProtocol_PrintDuplicate = PrintDuplicate(DuplicateNo)
    ReceiptNo = m_sLastReceiptNo
    ReceiptDate = m_dLastReceiptDate
End Function

Private Function IDeviceProtocol_Disconnect() As Boolean
    IDeviceProtocol_Disconnect = True
End Function

Private Function IDeviceProtocol_PrintReceipt(Optional ResumeToken As String, Optional ReceiptNo As String, Optional ReceiptDate As Date, Optional ReceiptAmount As Double, Optional InvoiceNo As String) As Boolean
    If PpdEndReceipt(m_uPrintData, ResumeToken) Then
        IDeviceProtocol_PrintReceipt = pvPrintReceipt(m_uPrintData.ExecCtx, m_uPrintData.Row, m_uPrintData.RowCount)
        With m_uPrintData.ExecCtx
            ReceiptNo = m_sLastReceiptNo
            ReceiptDate = m_dLastReceiptDate
            ReceiptAmount = m_dblLastReceiptAmount
            InvoiceNo = m_sLastInvoiceNo
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
    IDeviceProtocol_GetClock = GetClock
End Function

Private Function IDeviceProtocol_GetCommandLog() As String
    
End Function

Private Function IDeviceProtocol_GetOperatorName(OperatorCode As String) As String
    
End Function

Private Function IDeviceProtocol_GetDefaultPassword(OperatorCode As String) As String
    IDeviceProtocol_GetDefaultPassword = GetDefaultPassword(OperatorCode)
End Function

Private Function IDeviceProtocol_GetDeviceModel() As String
    IDeviceProtocol_GetDeviceModel = m_uConfig.DeviceModel
End Function

Private Function IDeviceProtocol_GetDeviceFirmware() As String
    IDeviceProtocol_GetDeviceFirmware = m_uConfig.FirmwareVersion
End Function

Private Function IDeviceProtocol_GetDeviceProtocol() As String
    IDeviceProtocol_GetDeviceProtocol = STR_PROTOCOL_PROXY
End Function

Private Function IDeviceProtocol_GetDeviceStatus(StatusText As String) As UcsFiscalStatusCodeEnum
    IDeviceProtocol_GetDeviceStatus = GetDeviceStatus(StatusText)
End Function

Private Function IDeviceProtocol_GetFiscalMemoryNo() As String
    IDeviceProtocol_GetFiscalMemoryNo = m_uConfig.FiscalMemoryNo
End Function

Private Function IDeviceProtocol_GetFooterText(ByVal Index As Long) As String
    IDeviceProtocol_GetFooterText = At(m_uConfig.Footers, Index - 1)
End Function

Private Function IDeviceProtocol_GetHeaderText(ByVal Index As Long) As String
    IDeviceProtocol_GetHeaderText = At(m_uConfig.Headers, Index - 1)
End Function

Private Function IDeviceProtocol_GetLastError(Optional ErrNo As UcsFiscalErrorsEnum) As String
    ErrNo = m_uPrintData.LastErrNo
    IDeviceProtocol_GetLastError = m_uPrintData.LastError
End Function

Private Function IDeviceProtocol_GetLastReceiptNo() As String
    IDeviceProtocol_GetLastReceiptNo = m_sLastReceiptNo
End Function

Private Function IDeviceProtocol_GetLastInvoiceNo() As String
    IDeviceProtocol_GetLastInvoiceNo = m_sLastInvoiceNo
End Function

Private Function IDeviceProtocol_GetLastQRCodeInfo(Optional LastReceiptDate As Date) As String
    LastReceiptDate = m_dLastReceiptDate
    IDeviceProtocol_GetLastQRCodeInfo = m_sLastReceiptNo
End Function

Private Function IDeviceProtocol_GetPaymentName(ByVal Index As Long) As String
    IDeviceProtocol_GetPaymentName = At(m_uConfig.PaymentNames, Index - 1)
End Function

Private Function IDeviceProtocol_GetResumeToken() As String
    IDeviceProtocol_GetResumeToken = m_sResumeToken
End Function

Private Function IDeviceProtocol_GetDeviceSerialNo() As String
    IDeviceProtocol_GetDeviceSerialNo = m_uConfig.DeviceSerialNo
End Function

Private Function IDeviceProtocol_GetTaxCaption() As String
    IDeviceProtocol_GetTaxCaption = m_uConfig.TaxCaption
End Function

Private Function IDeviceProtocol_GetTaxNo() As String
    IDeviceProtocol_GetTaxNo = m_uConfig.TaxNo
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
    IDeviceProtocol_IsConnected = True
End Property

Private Function IDeviceProtocol_OpenDrawer() As Boolean
    IDeviceProtocol_OpenDrawer = True
End Function

Private Function IDeviceProtocol_PrintMonthlyReport(ByVal ReportType As UcsFiscalReportsTypeEnum, ByVal StartDate As Date, ByVal EndDate As Date, Optional ReceiptNo As String, Optional ReceiptDate As Date) As Boolean
    IDeviceProtocol_PrintMonthlyReport = PrintMonthlyReport(ReportType, StartDate, EndDate)
    ReceiptNo = m_sLastReceiptNo
    ReceiptDate = m_dLastReceiptDate
End Function

Private Function IDeviceProtocol_PrintDailyXReport(ByVal ReportType As UcsFiscalReportsTypeEnum, Optional ReceiptNo As String, Optional ReceiptDate As Date) As Boolean
    IDeviceProtocol_PrintDailyXReport = PrintDailyReport(False, ReportType)
    ReceiptNo = m_sLastReceiptNo
    ReceiptDate = m_dLastReceiptDate
End Function

Private Function IDeviceProtocol_PrintDailyZReport(ByVal ReportType As UcsFiscalReportsTypeEnum, Optional ReceiptNo As String, Optional ReceiptDate As Date) As Boolean
    IDeviceProtocol_PrintDailyZReport = PrintDailyReport(True, ReportType)
    ReceiptNo = m_sLastReceiptNo
    ReceiptDate = m_dLastReceiptDate
End Function

Private Function IDeviceProtocol_SetClock(ByVal NewDate As Date) As Boolean
    IDeviceProtocol_SetClock = SetClock(NewDate)
End Function

Private Sub IDeviceProtocol_SetLocalizedText(ByVal Index As UcsFiscalLocalizedIndexesEnum, Text As String, Optional ByVal ConfigIndex As Long)
    LocalizedText(Index) = Text
End Sub

Private Function IDeviceProtocol_StartReceipt(ByVal ReceiptType As UcsFiscalReceiptTypeEnum, OperatorCode As String, OperatorName As String, OperatorPassword As String, Optional UniqueSaleNo As String, Optional TableNo As String, Optional ByVal DisablePrinting As Boolean, Optional InvDocNo As String, Optional InvCgTaxNo As String, Optional ByVal InvCgTaxNoType As UcsFiscalTaxNoTypeEnum, Optional InvCgVatNo As String, Optional InvCgName As String, Optional InvCgCity As String, Optional InvCgAddress As String, Optional InvCgPrsReceive As String, Optional OwnData As String, Optional ByVal RevType As UcsFiscalReversalTypeEnum, Optional RevReceiptNo As String, Optional ByVal RevReceiptDate As Date, Optional RevFiscalMemoryNo As String, Optional RevInvoiceNo As String, Optional RevReason As String) As Boolean
    m_uPrintData.Config.RowChars = pvRowChars
    m_uPrintData.Config.CommentChars = pvCommentChars
    m_uPrintData.Config.ItemChars = pvItemChars
    IDeviceProtocol_StartReceipt = PpdStartReceipt(m_uPrintData, ReceiptType, OperatorCode, OperatorName, OperatorPassword, TableNo, UniqueSaleNo, DisablePrinting, InvDocNo, InvCgTaxNo, InvCgTaxNoType, InvCgVatNo, InvCgName, InvCgCity, InvCgAddress, InvCgPrsReceive, RevType, RevReceiptNo, RevReceiptDate, RevFiscalMemoryNo, RevInvoiceNo, RevReason, OwnData)
End Function

Private Function IDeviceProtocol_GetDepartmentName(DepartmentCode As String, Optional TaxGroup As Long) As String
    
End Function

Private Function IDeviceProtocol_GetTaxRates() As Variant
    IDeviceProtocol_GetTaxRates = GetTaxRates()
End Function
