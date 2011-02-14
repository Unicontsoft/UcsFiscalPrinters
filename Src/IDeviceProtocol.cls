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
' $Header: /UcsFiscalPrinter/Src/IDeviceProtocol.cls 1     14.02.11 18:13 Wqw $
'
'   Unicontsoft Fiscal Printers Project
'   Copyright (c) 2008-2011 Unicontsoft
'
'   Protocol za uprawlenie na fiscalno ustrojstwo
'
' $Log: /UcsFiscalPrinter/Src/IDeviceProtocol.cls $
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
    ucsFscRetInvoiceNonFiscal
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
    ucsFscPmtCard
    ucsFscPmtCheque
    ucsFscPmtCredit
End Enum

Public Enum UcsFiscalLocalizedIndexesEnum
    ucsFscLciInternalErrors
    ucsFscLciPrinterStatuses
    ucsFscLciPrinterErrors
    ucsFscLciDipSwitchesTexts
    ucsFscLciMemoryTexts
End Enum

Public Function Init(Device As String, Optional ByVal Timeout As Long, Optional ByVal CashDeskNo As Long) As Boolean

End Function

Public Function Connect() As Boolean

End Function

Public Function Disconnect() As Boolean

End Function

Public Function GetLastError() As String

End Function

Public Function StartBon( _
            ByVal BonType As UcsFiscalReceiptTypeEnum, _
            ByVal PLUNumbering As Boolean, _
            CashierName As String, _
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
            ByVal Number As Long, _
            Name As String, _
            Price As Double, _
            Quantity As Double, _
            VatGroup As Long) As Boolean
            
End Function

Public Function AddLine( _
            Line As String, _
            Optional Command As String, _
            Optional ByVal WordWrap As Boolean = True) As Boolean
            
End Function

Public Function AddDiscount( _
            ByVal DiscType As UcsFiscalDiscountTypeEnum, _
            Value As Double) As Boolean
            
End Function

Public Function AddPayment( _
            ByVal Number As UcsFiscalPaymentTypeEnum, _
            Name As String, _
            Amount As Double, _
            Rate As Double) As Boolean
            
End Function

Public Function EndBon() As Boolean

End Function

Public Function CopyLastBon() As Boolean

End Function

Public Function GetClock() As Date

End Function

Public Sub SetClock(ByVal NewDate As Date)

End Sub

Public Function RunXReport() As Boolean

End Function

Public Function RunZReport() As Boolean

End Function

Public Function GetLastBonNumber() As Long

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

Public Function GetPaymentType(ByVal Index As Long) As String

End Function

Public Function GetCharsPerLine() As Long

End Function

Public Function GetPaymentExecuted() As Boolean

End Function

Public Function GetCommandLog() As String

End Function

Public Function GetDefaultPassword() As String

End Function
