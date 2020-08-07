Attribute VB_Name = "mdProtocolPrintData"
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
' Protocol's print data functions
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "mdProtocolPrintData"

'=========================================================================
' Public Enums
'=========================================================================

Public Enum UcsPpdInvDataIndex
    ucsInvDocNo
    ucsInvCgTaxNo
    ucsInvCgVatNo
    ucsInvCgName
    ucsInvCgCity
    ucsInvCgAddress
    ucsInvCgPrsReceive
    ucsInvCgTaxNoType
End Enum

Public Enum UcsPpdRevDataIndex
    ucsRevType
    ucsRevReceiptNo
    ucsRevReceiptDate
    ucsRevFiscalMemoryNo
    ucsRevInvoiceNo
    ucsRevReason
End Enum

Public Enum UcsPpdOwnDataIndex
    ucsOwnName
    ucsOwnAddress
    ucsOwnBulstat
    ucsOwnDepName
    ucsOwnDepAddress
    ucsOwnFooter1
    ucsOwnFooter2
End Enum

Public Enum UcsPpdRowTypeEnum
    ucsRowInit = 1
    ucsRowPlu
    ucsRowLine
    ucsRowDiscount
    ucsRowBarcode
    ucsRowPayment
End Enum

'=========================================================================
' Public Types
'=========================================================================

Private Const ERR_NO_RECEIPT_STARTED    As String = "No receipt started"
Private Const ERR_INVALID_DISCTYPE      As String = "Invalid discount type: %1"
Private Const TXT_SURCHARGE             As String = "Surcharge %1"
Private Const TXT_DISCOUNT              As String = "Discount %1"
Private Const TXT_PLUSALES              As String = "Sales %1"
Private Const STR_SEP                   As String = "|"
Public Const ucsFscDscPluAbs            As Long = ucsFscDscPlu + 100
Public Const ucsFscRcpNonfiscal         As Long = ucsFscRcpSale + 100
Public Const MIN_TAX_GROUP              As Long = 1
Public Const MAX_TAX_GROUP              As Long = 8
Public Const DEF_TAX_GROUP              As Long = 2
Public Const MIN_PMT_TYPE               As Long = 1
Public Const MAX_PMT_TYPE               As Long = [_ucsFscPmtMax] - 1
Public Const DEF_PMT_TYPE               As Long = 1
Public Const DEF_PRICE_SCALE            As Long = 2
Public Const DEF_QUANTITY_SCALE         As Long = 3

Public Type UcsPpdRowData
    RowType             As UcsPpdRowTypeEnum
    InitReceiptType     As UcsFiscalReceiptTypeEnum
    InitOperatorCode    As String
    InitOperatorName    As String
    InitOperatorPassword As String
    InitTableNo         As String
    InitUniqueSaleNo    As String
    InitDisablePrinting  As Boolean
    InitInvData         As Variant
    InitRevData         As Variant
    InitOwnData         As Variant
    PluItemName         As String
    PluPrice            As Double
    PluQuantity         As Double
    PluTaxGroup         As Long
    PluUnitOfMeasure    As String
    PluDepartmentNo     As Long
    LineText            As String
    LineCommand         As String
    LineWordWrap        As Boolean
    DiscType            As UcsFiscalDiscountTypeEnum
    DiscValue           As Double
    BarcodeType         As UcsFiscalBarcodeTypeEnum
    BarcodeText         As String
    BarcodeHeight       As Long
    PmtType             As UcsFiscalPaymentTypeEnum
    PmtName             As String
    PmtAmount           As Double
    PrintRowType        As UcsFiscalReceiptTypeEnum
End Type

Public Type UcsPpdExecuteContext
    GrpTotal(MIN_TAX_GROUP To MAX_TAX_GROUP) As Double
    Paid                As Double
    PluCount            As Long
    PmtPrinted          As Boolean
    ChangePrinted       As Boolean
    Row                 As Long
    ReceiptNo           As String
    ReceiptDate         As Date
    ReceiptAmount       As Double
    InvoiceNo           As String
    CommitSent          As Boolean
End Type

Public Type UcsPpdConfigValues
    RowChars            As Long
    CommentChars        As Long
    ItemChars           As Long
    AbsoluteDiscount    As Boolean
    NegativePrices      As Boolean
    MinDiscount         As Double
    MaxDiscount         As Double
    MaxReceiptRows      As Long
    MaxItemLines        As Long
End Type

Public Type UcsPpdLocalizedTexts
    ErrNoReceiptStarted As String
    ErrInvalidDiscType  As String
    TxtSurcharge        As String
    TxtDiscount         As String
    TxtPluSales         As String
End Type

Public Type UcsProtocolPrintData
    Row()               As UcsPpdRowData
    RowCount            As Long
    ExecCtx             As UcsPpdExecuteContext
    LastError           As String
    LastErrNo           As UcsFiscalErrorsEnum
    Config              As UcsPpdConfigValues
    LocalizedText       As UcsPpdLocalizedTexts
End Type

'=========================================================================
' Error handling
'=========================================================================

Private Sub RaiseError(sFunction As String)
    Debug.Print "Critical error: " & Err.Description & " [" & MODULE_NAME & "." & sFunction & "]"
    Logger.Log vbLogEventTypeError, MODULE_NAME, sFunction & "(" & Erl & ")", Err.Description
    Err.Raise Err.Number, MODULE_NAME & "." & sFunction & "(" & Erl & ")" & vbCrLf & Err.Source, Err.Description
End Sub

'=========================================================================
' Functions
'=========================================================================

Public Function PpdStartReceipt( _
            uData As UcsProtocolPrintData, _
            ByVal ReceiptType As UcsFiscalReceiptTypeEnum, _
            Optional OperatorCode As String, _
            Optional OperatorName As String, _
            Optional OperatorPassword As String, _
            Optional TableNo As String, _
            Optional UniqueSaleNo As String, _
            Optional ByVal DisablePrinting As Boolean, _
            Optional InvDocNo As String, _
            Optional InvCgTaxNo As String, _
            Optional ByVal InvCgTaxNoType As UcsFiscalTaxNoTypeEnum, _
            Optional InvCgVatNo As String, _
            Optional InvCgName As String, _
            Optional InvCgCity As String, _
            Optional InvCgAddress As String, _
            Optional InvCgPrsReceive As String, _
            Optional ByVal RevType As UcsFiscalReversalTypeEnum, _
            Optional RevReceiptNo As String, _
            Optional RevReceiptDate As Date, _
            Optional RevFiscalMemoryNo As String, _
            Optional RevInvoiceNo As String, _
            Optional RevReason As String, _
            Optional OwnData As String) As Boolean
    Const FUNC_NAME     As String = "PpdStartReceipt"
    Dim uCtxEmpty       As UcsPpdExecuteContext
    Dim sCity           As String
    Dim sAddress        As String

    On Error GoTo EH
    uData.ExecCtx = uCtxEmpty
    ReDim uData.Row(0 To 10) As UcsPpdRowData
    uData.RowCount = 0
    With uData.Row(pvAddRow(uData))
        .RowType = ucsRowInit
        .InitReceiptType = Clamp(ReceiptType, 1, [_ucsFscRcpMax] - 1)
        .InitOperatorCode = SafeText(OperatorCode)
        .InitOperatorName = SafeText(OperatorName)
        .InitOperatorPassword = SafeText(OperatorPassword)
        .InitTableNo = TableNo
        .InitUniqueSaleNo = SafeText(UniqueSaleNo)
        .InitDisablePrinting = DisablePrinting
        SplitCgAddress Trim$(SafeText(InvCgCity)) & vbCrLf & Trim$(SafeText(InvCgAddress)), sCity, sAddress, uData.Config.CommentChars
        .InitInvData = Array(SafeText(InvDocNo), SafeText(InvCgTaxNo), SafeText(InvCgVatNo), _
            SafeText(InvCgName), sCity, sAddress, SafeText(InvCgPrsReceive), InvCgTaxNoType)
        .InitRevData = Array(RevType, SafeText(RevReceiptNo), RevReceiptDate, SafeText(RevFiscalMemoryNo), _
            SafeText(RevInvoiceNo), SafeText(RevReason))
        .InitOwnData = Split(OwnData, STR_SEP)
        .PrintRowType = uData.Row(0).InitReceiptType
    End With
    '--- success
    PpdStartReceipt = True
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function PpdAddPLU( _
            uData As UcsProtocolPrintData, _
            Name As String, _
            ByVal Price As Double, _
            Optional ByVal Quantity As Double = 1, _
            Optional ByVal TaxGroup As Long = 2, _
            Optional UnitOfMeasure As String, _
            Optional ByVal DepartmentNo As Long, _
            Optional ByVal BeforeIndex As Long) As Boolean
    Const FUNC_NAME     As String = "PpdAddPLU"
    Dim uRow            As UcsPpdRowData
    Dim bNegative       As Boolean

    On Error GoTo EH
    '--- sanity check
    If uData.RowCount = 0 Then
        pvSetLastError uData, Zn(uData.LocalizedText.ErrNoReceiptStarted, ERR_NO_RECEIPT_STARTED)
        GoTo QH
    End If
    With uRow
        .RowType = ucsRowPlu
        .PluItemName = RTrim$(SafeText(Name))
        bNegative = (Round(Price, DEF_PRICE_SCALE) * Round(Quantity, DEF_QUANTITY_SCALE) < -DBL_EPSILON)
        .PluPrice = IIf(bNegative, -1, 1) * Round(Abs(Price), DEF_PRICE_SCALE)
        .PluQuantity = Round(Abs(Quantity), DEF_QUANTITY_SCALE)
        .PluTaxGroup = Clamp(TaxGroup, MIN_TAX_GROUP, MAX_TAX_GROUP)
        .PluUnitOfMeasure = UnitOfMeasure
        .PluDepartmentNo = DepartmentNo
        .PrintRowType = uData.Row(0).InitReceiptType
    End With
    pvInsertRow uData, BeforeIndex, uRow
    '--- success
    PpdAddPLU = True
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function PpdAddLine( _
            uData As UcsProtocolPrintData, _
            Line As String, _
            Optional ByVal WordWrap As Boolean = True, _
            Optional ByVal BeforeIndex As Long) As Boolean
    Const FUNC_NAME     As String = "PpdAddLine"
    Dim uRow            As UcsPpdRowData

    On Error GoTo EH
    '--- sanity check
    If uData.RowCount = 0 Then
        pvSetLastError uData, Zn(uData.LocalizedText.ErrNoReceiptStarted, ERR_NO_RECEIPT_STARTED)
        GoTo QH
    End If
    With uRow
        .RowType = ucsRowLine
        .LineText = RTrim$(SafeText(Line))
        .LineWordWrap = WordWrap
        .PrintRowType = uData.Row(0).InitReceiptType
    End With
    pvInsertRow uData, BeforeIndex, uRow
    '--- success
    PpdAddLine = True
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function PpdAddDiscount( _
            uData As UcsProtocolPrintData, _
            ByVal DiscType As UcsFiscalDiscountTypeEnum, _
            ByVal Value As Double, _
            Optional ByVal BeforeIndex As Long) As Boolean
    Const FUNC_NAME     As String = "PpdAddDiscount"
    Dim uRow            As UcsPpdRowData
    Dim lIdx            As Long
    Dim sText           As String

    On Error GoTo EH
    '--- sanity check
    If uData.RowCount = 0 Then
        pvSetLastError uData, Zn(uData.LocalizedText.ErrNoReceiptStarted, ERR_NO_RECEIPT_STARTED)
        GoTo QH
    End If
    Select Case DiscType
    Case ucsFscDscTotal
        PpdAddPLU uData, Printf(IIf(Value > DBL_EPSILON, Zn(uData.LocalizedText.TxtSurcharge, TXT_SURCHARGE), _
            Zn(uData.LocalizedText.TxtDiscount, TXT_DISCOUNT)), vbNullString), Value, BeforeIndex:=BeforeIndex
    Case ucsFscDscPlu
        For lIdx = IIf(BeforeIndex <> 0, BeforeIndex, uData.RowCount) - 1 To 0 Step -1
            With uData.Row(lIdx)
                If .RowType = ucsRowPlu Then
                    .DiscType = DiscType
                    .DiscValue = Round(Value, DEF_PRICE_SCALE)
                    Exit For
                End If
            End With
        Next
    Case ucsFscDscSubtotal, ucsFscDscSubtotalAbs
        If DiscType = ucsFscDscSubtotalAbs And Not uData.Config.AbsoluteDiscount Then
            sText = IIf(Value > DBL_EPSILON, Zn(uData.LocalizedText.TxtSurcharge, TXT_SURCHARGE), Zn(uData.LocalizedText.TxtDiscount, TXT_DISCOUNT))
            PpdAddPLU uData, Trim$(Printf(sText, vbNullString)), Value, TaxGroup:=pvGetLastTaxGroup(uData), BeforeIndex:=BeforeIndex
        Else
            With uRow
                .RowType = ucsRowDiscount
                .DiscType = DiscType
                .DiscValue = Round(Value, DEF_PRICE_SCALE)
                .PrintRowType = uData.Row(0).InitReceiptType
            End With
        End If
        pvInsertRow uData, BeforeIndex, uRow
    Case Else
        pvSetLastError uData, Printf(Zn(uData.LocalizedText.ErrInvalidDiscType, ERR_INVALID_DISCTYPE), DiscType)
        GoTo QH
    End Select
    '--- success
    PpdAddDiscount = True
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function PpdAddBarcode( _
            uData As UcsProtocolPrintData, _
            ByVal BarcodeType As UcsFiscalBarcodeTypeEnum, _
            Text As String, _
            Optional ByVal Height As Long) As Boolean
    Const FUNC_NAME     As String = "PpdAddBarcode"
    
    On Error GoTo EH
    '--- sanity check
    If uData.RowCount = 0 Then
        pvSetLastError uData, Zn(uData.LocalizedText.ErrNoReceiptStarted, ERR_NO_RECEIPT_STARTED)
        GoTo QH
    End If
    With uData.Row(pvAddRow(uData))
        .RowType = ucsRowBarcode
        .BarcodeType = Clamp(BarcodeType, 1, [_ucsFscBrcMax] - 1)
        .BarcodeText = Text
        .BarcodeHeight = Height
        .PrintRowType = uData.Row(0).InitReceiptType
    End With
    '--- success
    PpdAddBarcode = True
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function PpdAddPayment( _
            uData As UcsProtocolPrintData, _
            ByVal PmtType As UcsFiscalPaymentTypeEnum, _
            Optional PmtName As String, _
            Optional ByVal Amount As Double) As Boolean
    Const FUNC_NAME     As String = "PpdAddPayment"

    On Error GoTo EH
    '--- sanity check
    If uData.RowCount = 0 Then
        pvSetLastError uData, Zn(uData.LocalizedText.ErrNoReceiptStarted, ERR_NO_RECEIPT_STARTED)
        GoTo QH
    End If
    With uData.Row(pvAddRow(uData))
        .RowType = ucsRowPayment
        .PmtType = Clamp(PmtType, MIN_PMT_TYPE, MAX_PMT_TYPE)
        .PmtName = SafeText(PmtName)
        .PmtAmount = Round(Amount, DEF_PRICE_SCALE)
        .PrintRowType = uData.Row(0).InitReceiptType
    End With
    '--- success
    PpdAddPayment = True
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function PpdEndReceipt( _
            uData As UcsProtocolPrintData, _
            sResumeToken As String) As Boolean
    Const FUNC_NAME     As String = "PpdEndReceipt"
    Dim oToken          As Object
    Dim lIdx            As Long

    On Error GoTo EH
    '--- sanity check
    If uData.RowCount = 0 Then
        pvSetLastError uData, Zn(uData.LocalizedText.ErrNoReceiptStarted, ERR_NO_RECEIPT_STARTED)
        GoTo QH
    End If
    '--- restore context
    Set oToken = JsonParseObject(sResumeToken)
    With uData.ExecCtx
        For lIdx = LBound(.GrpTotal) To UBound(.GrpTotal)
            .GrpTotal(lIdx) = C_Dbl(JsonItem(oToken, "GrpTotal/" & lIdx - LBound(.GrpTotal)))
        Next
        .Paid = C_Dbl(JsonItem(oToken, "Paid"))
        .PluCount = C_Lng(JsonItem(oToken, "Paid"))
        .PmtPrinted = C_Bool(JsonItem(oToken, "PmtPrinted"))
        .ChangePrinted = C_Bool(JsonItem(oToken, "ChangePrinted"))
        .Row = C_Lng(JsonItem(oToken, "Row"))
        .ReceiptNo = C_Str(JsonItem(oToken, "ReceiptNo"))
        .ReceiptDate = C_Date(JsonItem(oToken, "ReceiptDate"))
        .ReceiptAmount = C_Dbl(JsonItem(oToken, "ReceiptAmount"))
        .InvoiceNo = C_Str(JsonItem(oToken, "InvoiceNo"))
    End With
    '--- fix fiscal receipts with for more than uData.MaxReceiptRows PLUs
    pvConvertExtraRows uData
    '--- append final payment (total)
    With uData.Row(pvAddRow(uData))
        .RowType = ucsRowPayment
        .PrintRowType = uData.Row(0).InitReceiptType
    End With
    '--- success
    PpdEndReceipt = True
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function PpdGetResumeToken(uData As UcsProtocolPrintData) As String
    Const FUNC_NAME     As String = "PpdGetResumeToken"
    Dim oToken          As Object

    On Error GoTo EH
    '--- sanity check
    If uData.RowCount = 0 Then
        pvSetLastError uData, Zn(uData.LocalizedText.ErrNoReceiptStarted, ERR_NO_RECEIPT_STARTED)
        GoTo QH
    End If
    '--- need resume token only if payment processed
    With uData.ExecCtx
        If .PmtPrinted Then
            JsonItem(oToken, "GrpTotal/*") = .GrpTotal
            JsonItem(oToken, "Paid") = .Paid
            JsonItem(oToken, "PluCount") = .PluCount
            If .PmtPrinted Then
                JsonItem(oToken, "PmtPrinted") = True
            End If
            If .ChangePrinted Then
                JsonItem(oToken, "ChangePrinted") = True
            End If
            JsonItem(oToken, "Row") = .Row
            JsonItem(oToken, "ReceiptNo") = Zn(.ReceiptNo, Empty)
            JsonItem(oToken, "ReceiptDate") = IIf(.ReceiptDate <> 0, .ReceiptDate, Empty)
            JsonItem(oToken, "ReceiptAmount") = IIf(Abs(.ReceiptAmount) > DBL_EPSILON, .ReceiptAmount, Empty)
            JsonItem(oToken, "InvoiceNo") = Zn(.InvoiceNo, Empty)
            PpdGetResumeToken = JsonDump(oToken, Minimize:=True)
        End If
    End With
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function PpdGetTotalsByPaymentTypes(uRow() As UcsPpdRowData, ByVal lRowCount As Long) As Double()
    Const FUNC_NAME     As String = "PpdGetTotalsByPaymentTypes"
    Dim vRetVal(0 To MAX_PMT_TYPE) As Double
    Dim lIdx            As Long
    Dim uSum            As UcsPpdExecuteContext
    
    On Error GoTo EH
    '--- calc payments
    For lIdx = 0 To lRowCount - 1
        With uRow(lIdx)
            If .RowType = ucsRowPayment Then
                vRetVal(.PmtType) = vRetVal(.PmtType) + .PmtAmount
            End If
        End With
    Next
    '--- calc receipt total
    pvGetSubtotals uRow, lRowCount, uSum
    vRetVal(0) = SumArray(uSum.GrpTotal)
    '--- success
    PpdGetTotalsByPaymentTypes = vRetVal
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

'= private ===============================================================

Private Function pvAddRow(uData As UcsProtocolPrintData) As Long
    Const FUNC_NAME     As String = "pvAddRow"

    On Error GoTo EH
    If uData.RowCount > UBound(uData.Row) Then
        ReDim Preserve uData.Row(0 To 2 * UBound(uData.Row)) As UcsPpdRowData
    End If
    pvAddRow = uData.RowCount
    uData.RowCount = uData.RowCount + 1
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Sub pvInsertRow(uData As UcsProtocolPrintData, ByVal lRow As Long, uRow As UcsPpdRowData)
    Const FUNC_NAME     As String = "pvInsertRow"
    Dim lIdx            As Long

    On Error GoTo EH
    If lRow = 0 Or lRow >= uData.RowCount Then
        uData.Row(pvAddRow(uData)) = uRow
    Else
        '--- shift rows down and insert new row
        For lIdx = pvAddRow(uData) To lRow + 1 Step -1
            uData.Row(lIdx) = uData.Row(lIdx - 1)
        Next
        uData.Row(lRow) = uRow
    End If
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Private Sub pvSetLastError(uData As UcsProtocolPrintData, sError As String, Optional ByVal ErrNum As UcsFiscalErrorsEnum = -1)
    If ErrNum < 0 Then
        uData.LastErrNo = IIf(LenB(sError) = 0, ucsFerNone, ucsFerGeneralError)
    Else
        uData.LastErrNo = ErrNum
    End If
    uData.LastError = sError
End Sub

Private Sub pvConvertExtraRows(uData As UcsProtocolPrintData)
    Const FUNC_NAME     As String = "pvConvertExtraRows"
    Dim uCtx            As UcsPpdExecuteContext
    Dim lIdx            As Long
    Dim lRow            As Long
    Dim lCount          As Long
    Dim lTotal          As Long
    Dim dblTotal        As Double
    Dim uSum            As UcsPpdExecuteContext
    Dim dblDiscount     As Double
    Dim dblDiscTotal    As Double
    Dim dblPrice        As Double
    Dim vSplit          As Variant

    On Error GoTo EH
    '--- convert out-of-range discounts to PLU rows
    '--- note: uData.RowCount may change in loop on PpdAddPLU
    Do While lRow < uData.RowCount
        '--- note: 'With' locks uData.Row array and fails if auto-grow needed in PpdAddPLU
'        With uData.Row(lRow)
        If uData.Row(lRow).RowType = ucsRowPlu Then
            dblPrice = uData.Row(lRow).PluPrice
            dblTotal = Round(uData.Row(lRow).PluQuantity * dblPrice, DEF_PRICE_SCALE)
            dblDiscTotal = Round(dblTotal * uData.Row(lRow).DiscValue / 100#, DEF_PRICE_SCALE)
            If Not uData.Config.NegativePrices And dblPrice < DBL_EPSILON Then '--- less than or *equal to* 0 (dblPrice <= 0)
                vSplit = WrapText(uData.Row(lRow).PluItemName, uData.Config.ItemChars)
                lIdx = Clamp(UBound(vSplit), , 1)
                vSplit(lIdx) = AlignText(vSplit(lIdx), SafeFormat(dblTotal + dblDiscTotal, FORMAT_BASE_2) & " " & Chr$(191 + uData.Row(lRow).PluTaxGroup), uData.Config.CommentChars)
                uData.Row(lRow).RowType = ucsRowLine
                uData.Row(lRow).LineText = vSplit(0)
                If lIdx > 0 Then
                    PpdAddLine uData, At(vSplit, 1), False, lRow + 1
                    lRow = lRow + 1
                ElseIf lIdx = 0 And Abs(uData.Row(lRow).PluQuantity - 1) > DBL_EPSILON Then
                    PpdAddLine uData, AlignText(vbNullString, SafeFormat(uData.Row(lRow).PluQuantity, FORMAT_BASE_3) & " x " & SafeFormat(uData.Row(lRow).PluPrice, FORMAT_BASE_2), uData.Config.CommentChars - 2), False, lRow
                End If
                If dblPrice < -DBL_EPSILON Then
                    PpdAddDiscount uData, ucsFscDscSubtotalAbs, dblTotal + dblDiscTotal, lRow + 1
                End If
            ElseIf (uData.Row(lRow).DiscValue < uData.Config.MinDiscount Or uData.Row(lRow).DiscValue > uData.Config.MaxDiscount) Then
                dblDiscount = Limit(uData.Row(lRow).DiscValue, uData.Config.MinDiscount, uData.Config.MaxDiscount)
                If uData.Config.AbsoluteDiscount Then
                    uData.Row(lRow).DiscType = ucsFscDscPluAbs
                    uData.Row(lRow).DiscValue = dblDiscTotal
                ElseIf dblDiscTotal = Round(dblTotal * dblDiscount / 100#, DEF_PRICE_SCALE) Then
                    uData.Row(lRow).DiscValue = dblDiscount
                Else
                    dblDiscount = uData.Row(lRow).DiscValue
                    uData.Row(lRow).DiscType = 0
                    uData.Row(lRow).DiscValue = 0
                    PpdAddPLU uData, Printf(IIf(dblDiscTotal > DBL_EPSILON, Zn(uData.LocalizedText.TxtSurcharge, TXT_SURCHARGE), Zn(uData.LocalizedText.TxtDiscount, TXT_DISCOUNT)), SafeFormat(Abs(dblDiscount), FORMAT_BASE_2) & " %"), _
                        dblDiscTotal, TaxGroup:=uData.Row(lRow).PluTaxGroup, BeforeIndex:=lRow + 1
                End If
            ElseIf uData.Row(lRow).DiscType = ucsFscDscPlu And dblPrice < -DBL_EPSILON Then '--- less than 0 (dblPrice < 0)
                '--- convert PLU discount on void rows
'                If uData.Config.AbsoluteDiscount Then
'                    uData.Row(lRow).DiscType = ucsFscDscPluAbs
'                    uData.Row(lRow).DiscValue = dblDiscTotal
'                Else
                    dblDiscount = uData.Row(lRow).DiscValue
                    uData.Row(lRow).DiscType = 0
                    uData.Row(lRow).DiscValue = 0
                    PpdAddPLU uData, Printf(IIf(dblTotal * dblDiscount > DBL_EPSILON, Zn(uData.LocalizedText.TxtSurcharge, TXT_SURCHARGE), Zn(uData.LocalizedText.TxtDiscount, TXT_DISCOUNT)), SafeFormat(Abs(dblDiscount), FORMAT_BASE_2) & " %"), _
                        dblDiscTotal, TaxGroup:=uData.Row(lRow).PluTaxGroup, BeforeIndex:=lRow + 1
'                End If
            End If
        ElseIf uData.Row(lRow).RowType = ucsRowDiscount Then
            If (uData.Row(lRow).DiscValue < uData.Config.MinDiscount Or uData.Row(lRow).DiscValue > uData.Config.MaxDiscount) And uData.Row(lRow).DiscType = ucsFscDscSubtotal Then
                pvGetSubtotals uData.Row, lRow, uSum
                dblDiscount = Limit(uData.Row(lRow).DiscValue, uData.Config.MinDiscount, uData.Config.MaxDiscount)
                lCount = 0
                For lIdx = LBound(uSum.GrpTotal) To UBound(uSum.GrpTotal)
                    If Round(uSum.GrpTotal(lIdx) * uData.Row(lRow).DiscValue / 100#, DEF_PRICE_SCALE) <> Round(uSum.GrpTotal(lIdx) * dblDiscount / 100#, DEF_PRICE_SCALE) Then
                        lCount = lCount + 1
                    End If
                Next
                If lCount = 0 Then
                    uData.Row(lRow).DiscValue = dblDiscount
                Else
                    dblDiscount = uData.Row(lRow).DiscValue
                    uData.Row(lRow).DiscValue = 0
                    For lIdx = UBound(uSum.GrpTotal) To LBound(uSum.GrpTotal) Step -1
                        If Abs(uSum.GrpTotal(lIdx)) > DBL_EPSILON Then
                            PpdAddPLU uData, Printf(IIf(uSum.GrpTotal(lIdx) * dblDiscount > DBL_EPSILON, Zn(uData.LocalizedText.TxtSurcharge, TXT_SURCHARGE), Zn(uData.LocalizedText.TxtDiscount, TXT_DISCOUNT)), SafeFormat(Abs(dblDiscount), FORMAT_BASE_2) & " %"), _
                                Round(uSum.GrpTotal(lIdx) * dblDiscount / 100#, DEF_PRICE_SCALE), TaxGroup:=lIdx, BeforeIndex:=lRow + 1
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
    For lRow = 0 To uData.RowCount - 1
        With uData.Row(lRow)
            If .RowType = ucsRowPlu Then
                lCount = lCount + 1
                uCtx.GrpTotal(.PluTaxGroup) = 1
            End If
        End With
    Next
    If lCount > uData.Config.MaxReceiptRows Then
        '--- count different VAT groups in PLUs
        For lRow = LBound(uCtx.GrpTotal) To UBound(uCtx.GrpTotal)
            If Abs(uCtx.GrpTotal(lRow)) > DBL_EPSILON Then
                lTotal = lTotal + 1
                uCtx.GrpTotal(lRow) = 0
            End If
        Next
        '--- set extra rows to nonfiscal printing and calc GrpTotal by VAT groups
        lCount = 0
        For lRow = 0 To uData.RowCount - 1
            With uData.Row(lRow)
                If .RowType = ucsRowPlu Then
                    lCount = lCount + 1
                    If lCount > uData.Config.MaxReceiptRows - lTotal Then
                        .PrintRowType = ucsFscRcpNonfiscal
                        dblTotal = Round(.PluQuantity * .PluPrice, DEF_PRICE_SCALE)
                        If .DiscType = ucsFscDscPlu Then
                            dblTotal = Round(dblTotal + Round(dblTotal * .DiscValue / 100#, DEF_PRICE_SCALE), DEF_PRICE_SCALE)
                        ElseIf .DiscType = ucsFscDscPluAbs Then
                            dblTotal = Round(dblTotal + .DiscValue, DEF_PRICE_SCALE)
                        End If
                        If .PluTaxGroup > 0 Then
                            uCtx.GrpTotal(.PluTaxGroup) = Round(uCtx.GrpTotal(.PluTaxGroup) + dblTotal, DEF_PRICE_SCALE)
                        End If
                    End If
                ElseIf .RowType = ucsRowDiscount And .DiscType = ucsFscDscSubtotal Then
                    If lCount > uData.Config.MaxReceiptRows - lTotal Then
                        .PrintRowType = ucsFscRcpNonfiscal
                        pvGetSubtotals uData.Row, lRow, uSum
                        For lIdx = LBound(uCtx.GrpTotal) To UBound(uCtx.GrpTotal)
                            uCtx.GrpTotal(lIdx) = Round(uCtx.GrpTotal(lIdx) + Round(uSum.GrpTotal(lIdx) * .DiscValue / 100#, DEF_PRICE_SCALE), DEF_PRICE_SCALE)
                        Next
                    End If
                End If
            End With
        Next
        '--- find first payment row
        For lRow = 0 To uData.RowCount - 1
            If uData.Row(lRow).RowType = ucsRowPayment Then
                Exit For
            End If
        Next
        '--- append fiscal rows for GrpTotal by VAT groups
        For lIdx = LBound(uCtx.GrpTotal) To UBound(uCtx.GrpTotal)
            If Abs(uCtx.GrpTotal(lIdx)) > DBL_EPSILON Then
                PpdAddPLU uData, Printf(Zn(uData.LocalizedText.TxtPluSales, TXT_PLUSALES), Chr$(191 + lIdx)), _
                    uCtx.GrpTotal(lIdx), TaxGroup:=lIdx, BeforeIndex:=lRow
                lRow = lRow + 1
            End If
        Next
    End If
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Private Sub pvGetSubtotals(uRow() As UcsPpdRowData, ByVal lRowCount As Long, uSum As UcsPpdExecuteContext)
    Const FUNC_NAME     As String = "pvGetSubtotals"
    Dim lIdx            As Long
    Dim lJdx            As Long
    Dim dblTotal        As Double
    Dim uEmpty          As UcsPpdExecuteContext

    On Error GoTo EH
    uSum = uEmpty
    For lIdx = 0 To lRowCount - 1
        With uRow(lIdx)
        If .RowType = ucsRowPlu Then
            dblTotal = Round(.PluQuantity * .PluPrice, DEF_PRICE_SCALE)
            Select Case .DiscType
            Case ucsFscDscPlu
                dblTotal = Round(dblTotal + Round(dblTotal * .DiscValue / 100#, DEF_PRICE_SCALE), DEF_PRICE_SCALE)
            Case ucsFscDscPluAbs
                dblTotal = Round(dblTotal + .DiscValue, DEF_PRICE_SCALE)
            End Select
            If .PluTaxGroup > 0 Then
                uSum.GrpTotal(.PluTaxGroup) = Round(uSum.GrpTotal(.PluTaxGroup) + dblTotal, DEF_PRICE_SCALE)
            End If
        ElseIf .RowType = ucsRowDiscount Then
            Select Case .DiscType
            Case ucsFscDscSubtotal
                For lJdx = LBound(uSum.GrpTotal) To UBound(uSum.GrpTotal)
                    dblTotal = Round(uSum.GrpTotal(lJdx) * .DiscValue / 100#, DEF_PRICE_SCALE)
                    uSum.GrpTotal(lJdx) = Round(uSum.GrpTotal(lJdx) + dblTotal, DEF_PRICE_SCALE)
                Next
            Case ucsFscDscSubtotalAbs
                '--- ToDo: fix for multiple tax groups
                For lJdx = LBound(uSum.GrpTotal) To UBound(uSum.GrpTotal)
                    If Abs(uSum.GrpTotal(lJdx)) > DBL_EPSILON Then
                        uSum.GrpTotal(lJdx) = Round(uSum.GrpTotal(lJdx) - .DiscValue, DEF_PRICE_SCALE)
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

Private Function pvGetLastTaxGroup(uData As UcsProtocolPrintData) As Long
    Const FUNC_NAME     As String = "pvGetLastTaxGroup"
    Dim lIdx            As Long
    
    On Error GoTo EH
    pvGetLastTaxGroup = DEF_TAX_GROUP
    For lIdx = uData.RowCount - 1 To 0 Step -1
        If uData.Row(lIdx).RowType = ucsRowPlu Then
            pvGetLastTaxGroup = uData.Row(lIdx).PluTaxGroup
            Exit Function
        End If
    Next
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

