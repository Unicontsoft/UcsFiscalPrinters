VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IDeviceProtocol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=========================================================================
' $Header: /UcsFiscalPrinter/Src/IDeviceProtocol.cls 16    20.06.13 12:48 Wqw $
'
'   Unicontsoft Fiscal Printers Project
'   Copyright (c) 2008-2013 Unicontsoft
'
'   Protocol za uprawlenie na fiscalno ustrojstwo
'
' $Log: /UcsFiscalPrinter/Src/IDeviceProtocol.cls $
' 
' 16    20.06.13 12:48 Wqw
' ADD: custom payment types
'
' 15    18.06.13 17:16 Wqw
' ADD: Function GetTotalsByPayments, GetTotalsStats
'
' 14    14.03.13 16:26 Wqw
' ADD: UcsFiscalErrorsEnum for GetLastError. REF: ByVal numeric params
'
' 13    4.01.13 12:21 Wqw
' REF: set clock returns bool, get last receipt number returns string
'
' 12    3.01.13 16:39 Wqw
' ADD: Function GetResumeTicket, GetDeviceStatus, Property IsConnected
'
' 11    9.10.12 15:19 Wqw
' REF: params of CashDebitCredit
'
' 10    5.10.12 14:19 Wqw
' REF: enum casing
'
' 9     5.10.12 14:15 Wqw
' ADD: Function CashDebitCredit
'
' 8     6.08.12 18:36 Wqw
' ADD: Function AutodetectDevices
'
' 7     23.03.12 15:26 Wqw
' ADD: GetTotalsByVatGroups, CancelReceipt
'
' 6     8.12.11 15:48 Wqw
' REF: params na GetDefaultPassword
'
' 5     17.06.11 11:35 Wqw
' REF: enums
'
' 4     22.02.11 13:53 Wqw
' ADD: Function GetDeviceProtocol
'
' 3     22.02.11 10:05 Wqw
' ADD: Function GetDeviceModel
'
' 2     21.02.11 13:44 Wqw
' REF: rename bon to receipt. ADD: Function RunPeriodReport
'
' 1     14.02.11 18:13 Wqw
' Initial implementation
'
'=========================================================================
Option Explicit

Public Enum UcsFiscalReceiptTypeEnum
    ucsFscRetNonfiscal = 1
    ucsFscRetFiscal
    ucsFscRetInvoiceFiscal
    ucsFscRetInvoiceNonfiscal
    [_ucsFscRetMax]
End Enum

Public Enum UcsFiscalDiscountTypeEnum
    ucsFscDstPlu = 1
    ucsFscDstSubtotal
    ucsFscDstTotal
    [_ucsFscDstMax]
End Enum

Public Enum UcsFiscalPaymentTypeEnum
    ucsFscPmtCash = 1
    ucsFscPmtCard = 2
    ucsFscPmtCheque = 3
    ucsFscPmtCustom1 = -1
    ucsFscPmtCustom2 = -2
    ucsFscPmtCustom3 = -3
End Enum

Public Enum UcsFiscalLocalizedIndexesEnum
    ucsFscLciInternalErrors
    ucsFscLciPrinterStatuses
    ucsFscLciPrinterErrors
    ucsFscLciDipSwitchesTexts
    ucsFscLciMemoryTexts
    ucsFscLciReceiptTexts
    [_ucsFscLciMax]
End Enum

Public Enum UcsFiscalReportsTypeEnum
    ucsFscRstDaily
    ucsFscRstDailyByItems
    ucsFscRstDailyByDepartment
    ucsFscRstDailyByItemsAndDepartment
    ucsFscRstPeriodShort
    ucsFscRstPeriodDetailed
End Enum

Public Enum UcsFiscalErrorsEnum
    ucsFerNone
    ucsFerGeneralError
    ucsFerInvalidPassword
    ucsFerPasswordNotSet
End Enum

Public Function Init(Device As String, Optional ByVal Timeout As Long, Optional ByVal CashDeskNo As Long) As Boolean

End Function

Public Function Connect() As Boolean

End Function

Public Function Disconnect() As Boolean

End Function

Property Get IsConnected() As Boolean

End Property

Public Function GetLastError(Optional ErrNumber As UcsFiscalErrorsEnum) As String

End Function

Public Function StartReceipt( _
            ByVal ReceiptType As UcsFiscalReceiptTypeEnum, _
            OperatorCode As String, _
            OperatorName As String, _
            OperatorPassword As String, _
            Optional InvNumber As String, _
            Optional InvTaxNumber As String, _
            Optional InvVatNumber As String, _
            Optional InvCompany As String, _
            Optional InvCity As String, _
            Optional InvAddress As String, _
            Optional InvMol As String, _
            Optional InvReceiver As String) As Boolean
            
End Function

Public Function AddPLU( _
            Name As String, _
            ByVal Price As Double, _
            ByVal Quantity As Double, _
            ByVal VatGroup As Long) As Boolean
            
End Function

Public Function AddLine( _
            Line As String, _
            Optional Command As String, _
            Optional ByVal WordWrap As Boolean = True) As Boolean
            
End Function

Public Function AddDiscount( _
            ByVal DiscType As UcsFiscalDiscountTypeEnum, _
            ByVal Value As Double) As Boolean
            
End Function

Public Function AddPayment( _
            ByVal Number As UcsFiscalPaymentTypeEnum, _
            Name As String, _
            ByVal Amount As Double, _
            Optional ByVal Rate As Double) As Boolean
            
End Function

Public Function EndReceipt(sResumeTicket As String) As Boolean

End Function

Public Function CopyLastReceipt() As Boolean

End Function

Public Function GetClock() As Date

End Function

Public Function SetClock(ByVal NewDate As Date) As Boolean

End Function

Public Function RunXReport(ByVal ReportType As UcsFiscalReportsTypeEnum) As Boolean

End Function

Public Function RunZReport(ByVal ReportType As UcsFiscalReportsTypeEnum) As Boolean

End Function

Public Function RunPeriodReport(ByVal ReportType As UcsFiscalReportsTypeEnum, ByVal StartDate As Date, ByVal EndDate As Date) As Boolean

End Function

Public Function GetLastReceiptNumber() As String

End Function

Public Function GetLastInvoiceNumber() As String

End Function

Public Function GetSerialNumber() As String

End Function

Public Function GetFiscalNumber() As String

End Function

Public Sub SetLocalizedText(ByVal Index As UcsFiscalLocalizedIndexesEnum, Text As String)

End Sub

Public Function GetTaxNumber() As String

End Function

Public Function GetTaxCaption() As String

End Function

Public Function GetHeaderText(ByVal Index As Long) As String

End Function

Public Function GetFooterText(ByVal Index As Long) As String

End Function

Public Function OpenDrawer() As Boolean

End Function

Public Function GetPaymentName(ByVal Index As Long) As String

End Function

Public Function GetCharsPerLine() As Long

End Function

Public Function GetCommandLog() As String

End Function

Public Function GetDefaultPassword(OperatorCode As String) As String

End Function

Public Function GetDeviceProtocol() As String

End Function

Public Function GetDeviceModel() As String

End Function

Public Function GetTotalsByVatGroups() As Variant

End Function

Public Function GetTotalsByPayments() As Variant

End Function

Public Function GetTotalsStats() As Variant

End Function

Public Function CancelReceipt() As Boolean

End Function

Public Function AutodetectDevices(Ports As Variant) As Variant

End Function

Public Function CashDebitCredit( _
            OperatorCode As String, _
            OperatorPassword As String, _
            ByVal Value As Double) As Variant
    
End Function

Public Function GetResumeTicket() As String

End Function

Public Function GetDeviceStatus(CurrentStatus As String) As Boolean

End Function

