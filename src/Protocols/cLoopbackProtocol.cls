VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLoopbackProtocol"
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
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "cLoopbackProtocol"
Implements IDeviceProtocol

'=========================================================================
' Constant and member variables
'=========================================================================

Private Const DEF_ROW_CHARS             As Long = 32
Private Const DEF_ITEM_CHARS            As Long = 72
Private Const DEF_MIN_DISCOUNT          As Double = -100
Private Const DEF_MAX_DISCOUNT          As Double = 100
Private Const DEF_MAX_RECEIPT_ROWS      As Long = 1000
Private Const DEF_HEADER                As String = "Наименование на задълженото лице|Адрес на задълженото лице|Наименование на обект|Адрес на обект"
Private Const DEF_FOOTER                As String = "Заповядайте отново!"
Private Const DEF_PMTNAMES              As String = "В брой|С чек|Талони|В.талони|Амбалаж|В.обслужване|Повреди|С карта|По банкa|НЗОК|Прог.тип 2"
Private Const DEF_TAXCAPTION            As String = "ЕИК"
Private Const DEF_TAXNO                 As String = "Незададен"

Private m_uPrintData                As UcsProtocolPrintData
Private m_uConfig                   As UcsConfigValues
Private m_dLastReceiptDate          As Date

Private Type UcsConfigValues
    DeviceSerialNo      As String
    FiscalMemoryNo      As String
    HeaderText          As Variant
    FooterText          As Variant
    PaymentNames        As Variant
    TaxCaption          As String
    TaxNo               As String
    LocalizedText(0 To [_ucsFscLciMax] - 1) As Variant
End Type

'=========================================================================
' Error handling
'=========================================================================

'Private Sub PrintError(sFunction As String)
'    Debug.Print "Critical error: " & Err.Description & " [" & MODULE_NAME & "." & sFunction & "]"
'    OutputDebugLog MODULE_NAME, sFunction & "(" & Erl & ")", "Run-time error: " & Err.Description
'End Sub

Private Sub RaiseError(sFunction As String)
    Debug.Print "Critical error: " & Err.Description & " [" & MODULE_NAME & "." & sFunction & "]"
    OutputDebugLog MODULE_NAME, sFunction & "(" & Erl & ")", "Run-time error: " & Err.Description
    Err.Raise Err.Number, MODULE_NAME & "." & sFunction & "(" & Erl & ")" & vbCrLf & Err.Source, Err.Description
End Sub

Private Sub DebugLog(sFunction As String, sText As String)
    OutputDebugLog MODULE_NAME, sFunction, sText
End Sub

'=========================================================================
' Properties
'=========================================================================

Private Property Get pvRowChars() As Long
    pvRowChars = DEF_ROW_CHARS
End Property

Private Property Get pvCommentChars() As Long
    pvCommentChars = pvRowChars - 2
End Property

Private Property Get pvItemChars() As Long
    pvItemChars = DEF_ITEM_CHARS
End Property

Property Get LocalizedText(ByVal eIdx As UcsFiscalLocalizedIndexesEnum) As String
    If IsArray(m_uConfig.LocalizedText(eIdx)) Then
        LocalizedText = Join(m_uConfig.LocalizedText(eIdx), "|")
    End If
End Property

Property Let LocalizedText(ByVal eIdx As UcsFiscalLocalizedIndexesEnum, sValue As String)
    m_uConfig.LocalizedText(eIdx) = Split(sValue, "|")
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function Init(sDeviceString As String) As Boolean
    Const FUNC_NAME     As String = "Init"
    Dim oOptions        As Object
    Dim sSerial         As String
    
    On Error GoTo EH
    DebugLog FUNC_NAME, "sDeviceString=" & sDeviceString
    Set oOptions = ParseDeviceString(sDeviceString)
    sSerial = JsonItem(oOptions, "DeviceSerialNo")
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
        .DeviceSerialNo = C_Str(JsonItem(oOptions, "DeviceSerialNo"))
        .FiscalMemoryNo = C_Str(JsonItem(oOptions, "FiscalMemoryNo"))
        .HeaderText = Split(Zn(C_Str(JsonItem(oOptions, "HeaderText")), DEF_HEADER), "|")
        .FooterText = Split(Zn(C_Str(JsonItem(oOptions, "FooterText")), DEF_FOOTER), "|")
        .PaymentNames = Split(Zn(C_Str(JsonItem(oOptions, "PaymentName")), DEF_PMTNAMES), "|")
        .TaxCaption = Zn(C_Str(JsonItem(oOptions, "TaxCaption")), DEF_TAXCAPTION)
        .TaxNo = Zn(C_Str(JsonItem(oOptions, "TaxNo")), DEF_TAXNO)
    End With
    '--- success
    Init = True
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetTaxRates() As Variant
    Dim vTaxRates       As Variant
    Dim vRetVal(0 To MAX_TAX_GROUP - 1) As Double
    Dim lIdx            As Long
    Const FUNC_NAME     As String = "GetTaxRates"
    
    On Error GoTo EH
    vTaxRates = Split("0 20 20 9 -1 -1 -1 -1")
    For lIdx = 0 To UBound(vRetVal)
        vRetVal(lIdx) = C_Dbl(At(vTaxRates, lIdx))
    Next
    IDeviceProtocol_GetTaxRates = vRetVal
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

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
    
End Function

Private Function IDeviceProtocol_CancelReceipt() As Boolean
    IDeviceProtocol_CancelReceipt = True
End Function

Private Function IDeviceProtocol_PrintServiceDeposit(OperatorCode As String, OperatorPassword As String, ByVal Amount As Double) As Variant
    m_dLastReceiptDate = Now
    IDeviceProtocol_PrintServiceDeposit = True
End Function

Private Function IDeviceProtocol_Connect() As Boolean
    IDeviceProtocol_Connect = True
End Function

Private Function IDeviceProtocol_CopyLastReceipt(Optional ReceiptNumber As String) As Boolean
    IDeviceProtocol_CopyLastReceipt = False
End Function

Private Function IDeviceProtocol_Disconnect() As Boolean
    IDeviceProtocol_Disconnect = True
End Function

Private Function IDeviceProtocol_EndReceipt(Optional ResumeToken As String) As Boolean
    If PpdEndReceipt(m_uPrintData, ResumeToken) Then
        m_dLastReceiptDate = Now
        IDeviceProtocol_EndReceipt = True
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
    IDeviceProtocol_GetClock = Now
End Function

Private Function IDeviceProtocol_GetCommandLog() As String
    
End Function

Private Function IDeviceProtocol_GetOperatorName(OperatorCode As String) As String
    
End Function

Private Function IDeviceProtocol_GetDefaultPassword(OperatorCode As String) As String
    IDeviceProtocol_GetDefaultPassword = OperatorCode
End Function

Private Function IDeviceProtocol_GetDeviceModel() As String
    IDeviceProtocol_GetDeviceModel = STR_PROTOCOL_LOOPBACK
End Function

Private Function IDeviceProtocol_GetDeviceFirmware() As String
    IDeviceProtocol_GetDeviceFirmware = "N/A"
End Function

Private Function IDeviceProtocol_GetDeviceProtocol() As String
    IDeviceProtocol_GetDeviceProtocol = STR_PROTOCOL_LOOPBACK
End Function

Private Function IDeviceProtocol_GetDeviceStatus(CurrentStatus As String) As Boolean
    
End Function

Private Function IDeviceProtocol_GetFiscalMemoryNo() As String
    IDeviceProtocol_GetFiscalMemoryNo = m_uConfig.FiscalMemoryNo
End Function

Private Function IDeviceProtocol_GetFooterText(ByVal Index As Long) As String
    IDeviceProtocol_GetFooterText = At(m_uConfig.FooterText, Index - 1)
End Function

Private Function IDeviceProtocol_GetHeaderText(ByVal Index As Long) As String
    IDeviceProtocol_GetHeaderText = At(m_uConfig.HeaderText, Index - 1)
End Function

Private Function IDeviceProtocol_GetLastError(Optional ErrNo As UcsFiscalErrorsEnum) As String
    
End Function

Private Function IDeviceProtocol_GetLastReceiptNo() As String
    
End Function

Private Function IDeviceProtocol_GetLastInvoiceNo() As String
    
End Function

Private Function IDeviceProtocol_GetLastQRCodeInfo(Optional LastReceiptDate As Date) As String
    LastReceiptDate = m_dLastReceiptDate
End Function

Private Function IDeviceProtocol_GetPaymentName(ByVal Index As Long) As String
    IDeviceProtocol_GetPaymentName = At(m_uConfig.PaymentNames, Index - 1)
End Function

Private Function IDeviceProtocol_GetResumeToken() As String
    
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
    
End Function

Private Function IDeviceProtocol_GetTotalsByTaxGroups() As Variant
    
End Function

Private Function IDeviceProtocol_GetTotalsStats() As Variant
    
End Function

Private Function IDeviceProtocol_Init(DeviceString As String) As Boolean
    IDeviceProtocol_Init = Init(DeviceString)
End Function

Private Property Get IDeviceProtocol_IsConnected() As Boolean
    IDeviceProtocol_IsConnected = True
End Property

Private Function IDeviceProtocol_OpenDrawer() As Boolean
    
End Function

Private Function IDeviceProtocol_RunPeriodReport(ByVal ReportType As UcsFiscalReportsTypeEnum, ByVal StartDate As Date, ByVal EndDate As Date) As Boolean
    m_dLastReceiptDate = Now
    IDeviceProtocol_RunPeriodReport = True
End Function

Private Function IDeviceProtocol_RunXReport(ByVal ReportType As UcsFiscalReportsTypeEnum) As Boolean
    m_dLastReceiptDate = Now
    IDeviceProtocol_RunXReport = True
End Function

Private Function IDeviceProtocol_RunZReport(ByVal ReportType As UcsFiscalReportsTypeEnum) As Boolean
    m_dLastReceiptDate = Now
    IDeviceProtocol_RunZReport = True
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
    IDeviceProtocol_StartReceipt = PpdStartReceipt(m_uPrintData, ReceiptType, OperatorCode, OperatorName, OperatorPassword, TableNo, UniqueSaleNo, DisablePrinting, InvDocNo, InvCgTaxNo, InvCgTaxNoType, InvCgVatNo, InvCgName, InvCgCity, InvCgAddress, InvCgPrsReceive, RevType, RevReceiptNo, RevReceiptDate, RevFiscalMemoryNo, RevInvoiceNo, RevReason, OwnData)
End Function

Private Function IDeviceProtocol_GetDepartmentName(DepartmentCode As String, Optional TaxGroup As Long) As String
    
End Function

Private Function IDeviceProtocol_GetTaxRates() As Variant
    IDeviceProtocol_GetTaxRates = GetTaxRates()
End Function
