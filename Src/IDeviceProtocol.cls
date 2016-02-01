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
Attribute VB_Description = "IDeviceProtocol is a common interface implemented by all device protocol handlers. It can be used to access a fiscal device independently of the implemented protocol so that the particular protocol selection can be left to be chosen by the end-user or as an application configuration option"
''
' IDeviceProtocol is a common interface implemented by all device protocol handlers. It can be used to access a fiscal
' device independently of the implemented protocol so that the particular protocol selection can be left to be chosen
' by the end-user or as an application configuration option.
'
' If any method return `False` or an exception is raised then `GetLastError` method can be used to retrieve current
' error text and error category as number.
'
' See `PrintReceipt` method of `cFiscalPrinter` class for a complete workflow for printing a receipt or invoice.
'
' Basicly one can start by calling `Init` to open the serial COM port, then `StartReceipt` to specify the type of
' receipt to be printed, followed by multiple calls to `AddPLU` and `AddDiscount` for each product to be printed on the
' receipt, followed by one or multiple calls to `AddPayment` to specify paid amount by payment types and finally
' complete and print the receipt by calling `EndReceipt`.
'
' Also see `PrintReport` method of `cFiscalPrinter` class for a complete workflow for printing a daily or monthly
' fiscal reports.
'
'=========================================================================
' $Header: /UcsFiscalPrinter/Src/IDeviceProtocol.cls 18    1.02.16 17:17 Wqw $
'
'   Unicontsoft Fiscal Printers Project
'   Copyright (c) 2008-2016 Unicontsoft
'
'   Abstract interface for fiscal device protocol handling
'
' $Log: /UcsFiscalPrinter/Src/IDeviceProtocol.cls $
' 
' 18    1.02.16 17:17 Wqw
' ADD: entries in UcsFiscalLocalizedIndexesEnum
'
' 17    20.11.15 16:40 Wqw
' REF: add VBDOX documentation and typelib descriptions
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
    ucsFscLciExtendedErrors
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

''
' Used to initialize the protocol handler. Here the COM port is opened and initial device configuration is retrieved
'
' @param `Device`       [input] string formatted as "ComPort[,BaudRate[,Persistent[,DataBits[,Parity[,StopBits]]]]]"
' @param `Timeout`      [in, optional] COM port open timeout in milliseconds, usually defaults to 3000
' @param `ChashDeskNo`  [in, optional] cash desk number if supported by the device, defaults to 1
' @return `True` on success, `False` on error. `GetLastError` method can be use to retrieve error details
'
Public Function Init(Device As String, Optional ByVal Timeout As Long, Optional ByVal CashDeskNo As Long) As Boolean
Attribute Init.VB_Description = "Used to initialize the protocol handler. Here the COM port is opened and initial device configuration is retrieved"

End Function

''
' Used to explicitly open the COM port of the device
'
' @return `True` on success, `False` on error. `GetLastError` method can be use to retrieve error details
'
Public Function Connect() As Boolean
Attribute Connect.VB_Description = "Used to explicitly open the COM port of the device"

End Function

''
' Used to explicitly close the COM port of the device
'
' @return `True` on success, `False` on error. `GetLastError` method can be use to retrieve error details
'
Public Function Disconnect() As Boolean
Attribute Disconnect.VB_Description = "Used to explicitly close the COM port of the device"

End Function

''
' Returns if the device COM port is connected
'
Property Get IsConnected() As Boolean
Attribute IsConnected.VB_Description = "Returns if the device COM port is connected"

End Property

''
' Used to retrieve details about last error that occurred during protocol operations
'
' @param `ErrNumber` [out, optional] Error category of the last error
' @return Text description of the last error
'
Public Function GetLastError(Optional ErrCategory As UcsFiscalErrorsEnum) As String
Attribute GetLastError.VB_Description = "Used to retrieve details about last error that occurred during protocol operations"

End Function

''
' Used to begin printing of a new receipt. The actual printing of the receipt is usually postponed until `EndReceipt`
' is called
'
' @param `ReceiptType`      [input] designates if receipt is fiscal or non-fiscal and if the receipt is an invoice or
'                                   regular sales receipt
' @param `OperatorCode`     [input] operator number to be authenticated (usually 1-8)
' @param `OperatorName`     [input] operator name to be printed on receipt
' @param `OperatorPassword` [input] operator password, can be `Empty` if default for device
' @param `InvNumber`        [in, optional] the number of the invoice/credit note to be printed
' @param `InvTaxNumber`     [in, optional] local identification number of the company (country specific)
' @param `InvVatNumber`     [in, optional] VAT identification number of the company (EU specific)
' @param `InvCompany`       [in, optional] name of the company that is receiving the document
' @param `InvCity`          [in, optional] city of the company that is receiving the document, can be empty string
' @param `InvAddress`       [in, optional] address of the company that is receiving the document
' @param `InvMol`           [in, optional] name of the person that is fiscally responsible for the company (not used)
' @param `InvReceiver`      [in, optional] name of the person that is receiving the document
' @return `True` on success, `False` on error. `GetLastError` method can be use to retrieve error details
'
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
Attribute StartReceipt.VB_Description = "Used to begin printing of a new receipt. The actual printing of the receipt is usually postponed until `EndReceipt` is called"
            
End Function

''
' Used to print a row on the final receipt with product description, price and quantity
'
' @param `Name`     [input] full product name or description
' @param `Price`    [input] final sale price (incl. VAT) rounded to 0.00, can be negative for voids
' @param `Quantity` [input] quantity rounded to 0.000
' @param `VatGroup` [input] VAT group the row total is accumulated to. VAT total will be printed at the end of receipt
' @return `True` on success, `False` on error. `GetLastError` method can be use to retrieve error details
' @remark When printing negative rows currently reached receipt subtotal cannot become negative or the protocol handler
' will reorder receipt rows to prevent fiscal device from rejecting the whole receipt.
'
Public Function AddPLU( _
            Name As String, _
            ByVal Price As Double, _
            ByVal Quantity As Double, _
            ByVal VatGroup As Long) As Boolean
Attribute AddPLU.VB_Description = "Used to print a row on the final receipt with product description, price and quantity"
            
End Function

''
' Used to print a row with free text that does not get stored in fiscal memory
'
' @param `Line`     [input] row text
' @param `Command`  [input] text format. One of "DBWIDTH" or "CONDENSED". Not all devices support this param
' @param `WordWrap` [input] if `True` text will be wrapped and can be printed on multiple rows
' @return `True` on success, `False` on error. `GetLastError` method can be use to retrieve error details
'
Public Function AddLine( _
            Line As String, _
            Optional Command As String, _
            Optional ByVal WordWrap As Boolean = True) As Boolean
Attribute AddLine.VB_Description = "Used to print a row with free text that does not get stored in fiscal memory"
            
End Function

''
' Used to discount/surcharge previous PLU row, or to print discount/surcharge on current subtotal or receipt total
'
' @param `DiscType` [input] type of discount/surcharge. See `UcsFiscalDiscountTypeEnum`
' @param `Value`    [input] discount/srcharge percent in range -100.00 to 100.00 (positive for discount)
' @return `True` on success, `False` on error. `GetLastError` method can be use to retrieve error details
'
Public Function AddDiscount( _
            ByVal DiscType As UcsFiscalDiscountTypeEnum, _
            ByVal Value As Double) As Boolean
Attribute AddDiscount.VB_Description = "Used to discount/surcharge previous PLU row, or to print discount/surcharge on current subtotal or receipt total"
            
End Function

''
' Used to finalize PLU rows and print payment information
'
' @param `Number`   [input] payment type. See `UcsFiscalPaymentTypeEnum`
' @param `Name      [input] payment name if `Number` is one of custom payment types
' @param `Amount`   [input] payment amount rounded to 0.00. Cannot be negative
' @return `True` on success, `False` on error. `GetLastError` method can be use to retrieve error details
' @remark After adding a payment no more PLUs can be printed. Several payments can be accumulated on a single receipt
' e.g. 10.00 by card and 50.00 in cash, etc. are able to pay out a 35.00 total receipt
'
Public Function AddPayment( _
            ByVal Number As UcsFiscalPaymentTypeEnum, _
            Name As String, _
            ByVal Amount As Double, _
            Optional ByVal Rate As Double) As Boolean
Attribute AddPayment.VB_Description = "Used to finalize PLU rows and print payment information"
            
End Function

''
' Used to finalize receipt payments and batch print the complete receipt
'
' @param `ResumeTicket` [in, optional] if previous invocation failed can be used to resume printing on the same line
' @return `True` on success, `False` on error. `GetLastError` method can be use to retrieve error details
' @remark All previously added payments must add up to enough total amount so to be able to pay out all the PLU rows
' incl. discounted/surcharged rows
'
Public Function EndReceipt(Optional ResumeTicket As String) As Boolean
Attribute EndReceipt.VB_Description = "Used to finalize receipt payments and batch print the complete receipt"

End Function

''
' Used to print a duplicate of the last fiscal receipt
'
' @return `True` on success, `False` on error. `GetLastError` method can be use to retrieve error details
' @remark Most fiscal devices are able to print a single duplicate of last fiscal receipt only. On second copy request
' an error is returned.
'
Public Function CopyLastReceipt() As Boolean
Attribute CopyLastReceipt.VB_Description = "Used to print a duplicate of the last fiscal receipt"

End Function

''
' Used to retrieve current date/time from the clock on the fiscal device
'
' @return Clock's current date/time
'
Public Function GetClock() As Date
Attribute GetClock.VB_Description = "Used to retrieve current date/time from the clock on the fiscal device"

End Function

''
' Used to set current date/time of the clock on the fiscal device
'
' @param `NewDate` [input] new date/time to set to the clock
' @return `True` on success, `False` on error. `GetLastError` method can be use to retrieve error details
'
Public Function SetClock(ByVal NewDate As Date) As Boolean
Attribute SetClock.VB_Description = "Used to set current date/time of the clock on the fiscal device"

End Function

''
' Used to print fiscal x-report
'
' @param `ReportType` [input] type of detailed report to print: by items, by departments, by both
' @return `True` on success, `False` on error. `GetLastError` method can be use to retrieve error details
'
Public Function RunXReport(ByVal ReportType As UcsFiscalReportsTypeEnum) As Boolean
Attribute RunXReport.VB_Description = "Used to print fiscal x-report"

End Function

''
' Used to print fiscal z-report and clear fiscal memory for the day
'
' @param `ReportType` [input] type of detailed report to print: by items, by departments, by both
' @return `True` on success, `False` on error. `GetLastError` method can be use to retrieve error details
'
Public Function RunZReport(ByVal ReportType As UcsFiscalReportsTypeEnum) As Boolean
Attribute RunZReport.VB_Description = "Used to print fiscal z-report and clear fiscal memory for the day"

End Function

''
' Used to print monthly fiscal report by date range
'
' @param `ReportType` [input] type of report to print: short or details
' @param `StartDate`  [input] start of date range
' @param `EndDate`    [input] end of date range
' @return `True` on success, `False` on error. `GetLastError` method can be use to retrieve error details
'
Public Function RunPeriodReport( _
            ByVal ReportType As UcsFiscalReportsTypeEnum, _
            ByVal StartDate As Date, _
            ByVal EndDate As Date) As Boolean
Attribute RunPeriodReport.VB_Description = "Used to print monthly fiscal report by date range"

End Function

''
' Used to retrieve the fiscal number of the last fiscal slip printed. This includes sales receipts, invoices, fiscal
' reports and credit/debit slips
'
' @return Last receipt number as string
'
Public Function GetLastReceiptNumber() As String
Attribute GetLastReceiptNumber.VB_Description = "Used to retrieve the fiscal number of the last fiscal slip printed. This includes sales receipts, invoices, fiscal reports and credit/debit slips"

End Function

''
' Used to retrieve last invoice number printed which has come from current fiscal device invoices range
'
' @return Last invoice number as string
'
Public Function GetLastInvoiceNumber() As String
Attribute GetLastInvoiceNumber.VB_Description = "Used to retrieve last invoice number printed which has come from current fiscal device invoices range"

End Function

''
' Used to retrieve fiscal device serial number as reported by firmware
'
' @return Serial number as string
'
Public Function GetSerialNumber() As String
Attribute GetSerialNumber.VB_Description = "Used to retrieve fiscal device serial number as reported by firmware"

End Function

''
' Used to retrieve fiscal number of the registered fiscal memory
'
' @return Fiscal number as string
'
Public Function GetFiscalNumber() As String
Attribute GetFiscalNumber.VB_Description = "Used to retrieve fiscal number of the registered fiscal memory"

End Function

''
' Used to set localized strings used for non-fiscal receipt texts and component internal errors
'
' @param `Index` [input] type of localized strings. See `UcsFiscalLocalizedIndexesEnum`
'
Public Sub SetLocalizedText(ByVal Index As UcsFiscalLocalizedIndexesEnum, Text As String)
Attribute SetLocalizedText.VB_Description = "Used to set localized strings used for non-fiscal receipt texts and component internal errors"

End Sub

''
' Used to retrieve tax number of the company that the fiscal device belongs to
'
' @return Company's tax number as string
'
Public Function GetTaxNumber() As String
Attribute GetTaxNumber.VB_Description = "Used to retrieve tax number of the company that the fiscal device belongs to"

End Function

''
' Used to retrieve the kind of tax number the owning company is using
'
' @return Tax number king as string
'
Public Function GetTaxCaption() As String
Attribute GetTaxCaption.VB_Description = "Used to retrieve the kind of tax number the owning company is using"

End Function

''
' Used to retrieve template text for the header rows printed on receipts
'
' @param `Index` [input] index of the header row to be retrieved
' @return Template text
'
Public Function GetHeaderText(ByVal Index As Long) As String
Attribute GetHeaderText.VB_Description = "Used to retrieve template text for the header rows printed on receipts"

End Function

''
' Used to retrieve template text for the footer rows printed on receipts
'
' @param `Index` [input] index of the footer row to be retrieved
' @return Template text
'
Public Function GetFooterText(ByVal Index As Long) As String
Attribute GetFooterText.VB_Description = "Used to retrieve template text for the footer rows printed on receipts"

End Function

''
' Used to send an impulse to cash drawer opener if connected to the fiscal device
'
' @return `True` on success, `False` on error. `GetLastError` method can be use to retrieve error details
'
Public Function OpenDrawer() As Boolean
Attribute OpenDrawer.VB_Description = "Used to send an impulse to cash drawer opener if connected to the fiscal device"

End Function

''
' Used to retrieve standard and custom payment names
'
' @param `Index` [input] type of payment to be retrieved
' @return Localized payment name
'
Public Function GetPaymentName(ByVal Index As Long) As String
Attribute GetPaymentName.VB_Description = "Used to retrieve standard and custom payment names"

End Function

''
' Used to retrieve width of the paper roll in device characters for free text rows
'
' @return Number of characters available on a single row
' @remark This is related to the maximum size of PLU names to be printed without wrapping to multiple rows. Usually
' max size of PLU name is 7 characters less.
'
Public Function GetCharsPerLine() As Long
Attribute GetCharsPerLine.VB_Description = "Used to retrieve width of the paper roll in device characters for free text rows"

End Function

''
' Used to retrieve debug log from device protocol handler
'
' @return Communication log, including dent commands
'
Public Function GetCommandLog() As String
Attribute GetCommandLog.VB_Description = "Used to retrieve debug log from device protocol handler"

End Function

''
' Used to retrieve default operator password as specified in fiscal device documentation
'
' @param `OperatorCode` [input] operator number to be authenticated (usually 1-8)
' @return Default operator password in plain-text
' @remark Might not match current operator password if device is manually configured.
'
Public Function GetDefaultPassword(OperatorCode As String) As String
Attribute GetDefaultPassword.VB_Description = "Used to retrieve default operator password as specified in fiscal device documentation"

End Function

''
' Used to retrieve current protocol handler descriptive name
'
' @return Device protocol name
'
Public Function GetDeviceProtocol() As String
Attribute GetDeviceProtocol.VB_Description = "Used to retrieve current protocol handler descriptive name"

End Function

''
' Used to retrieve device model name as reported by current fiscal device firmware
'
' @return Device model name
'
Public Function GetDeviceModel() As String
Attribute GetDeviceModel.VB_Description = "Used to retrieve device model name as reported by current fiscal device firmware"

End Function

''
' Used to retrieve totals by VAT groups since last z-report
'
' @return Array of totals for each VAT group (1-8)
'
Public Function GetTotalsByVatGroups() As Variant
Attribute GetTotalsByVatGroups.VB_Description = "Used to retrieve totals by VAT groups since last z-report"

End Function

''
' Used to retrieve totals by payment types since last z-report
'
' @return Array of totals for each payment type (1-8)
'
Public Function GetTotalsByPayments() As Variant
Attribute GetTotalsByPayments.VB_Description = "Used to retrieve totals by payment types since last z-report"

End Function

''
' Used to retrieve total number of receipt printed and date/time of last receipt printed
'
' @return Array of (ReceiptsCount, LastReceiptTime)
'
Public Function GetTotalsStats() As Variant
Attribute GetTotalsStats.VB_Description = "Used to retrieve total number of receipt printed and date/time of last receipt printed"

End Function

''
' Used to cancel previous pending receipt if `EndReceipt` returned error before finalizing (closing) the fiscal
' receipt
'
' @return `True` on success, `False` on error. `GetLastError` method can be use to retrieve error details
' @remark It might by imposible to cancel the fiscal receipt if payment rows are sent to fiscal device. The only
' possible way to get back to normal operations is to pay the receipt in full, thus making it accountable in the
' daily fiscal report.
'
Public Function CancelReceipt() As Boolean
Attribute CancelReceipt.VB_Description = "Used to cancel previous pending receipt if `EndReceipt` returned error before finalizing (closing) the fiscal receipt"

End Function

''
' Used to auto-detect devices that implemented current protocol on the selected COM ports
'
' @param `Ports` [intput] Array of strings with COM ports to check
' @return Array of arrays with COM ports and auto-detected devices info.
'
Public Function AutodetectDevices(Ports As Variant) As Variant
Attribute AutodetectDevices.VB_Description = "Used to auto-detect devices that implemented current protocol on the selected COM ports"

End Function

''
' Used to credit/debit cash desk outside of regular sales and daily z-report with memory clear
'
' @param `OperatorCode`     [input] operator number to be authenticated (usually 1-8)
' @param `OperatorPassword` [input] operator password, can be `Empty` if default for device
' @param `Value`            [input] positive for debig, negative for creadit, 0.00 for info only
' @return Array of doubles with (TotalAvailable, TotalDebit, TotalCredit)
'
Public Function CashDebitCredit( _
            OperatorCode As String, _
            OperatorPassword As String, _
            ByVal Value As Double) As Variant
Attribute CashDebitCredit.VB_Description = "Used to credit/debit cash desk outside of regular sales and daily z-report with memory clear"
    
End Function

''
' Used to retrieve resume-ticket after failed `EndReceipt`
'
' @return String encoded resume-ticket
' @remark Result can be passed to `EndReceipt` to resume printing exactly from the previsouly failed row e.g. if the
' device ran out-of-paper the operator can change the roll and resume printing and later print a brand new duplicate
' receipt for the end client.
'
Public Function GetResumeTicket() As String
Attribute GetResumeTicket.VB_Description = "Used to retrieve resume-ticket after failed `EndReceipt`"

End Function

''
' Used to retrieve current device status without performing any other operation
'
' @param `CurrentStatus` [output] device error or empty string for ready
' @return `True` on success, `False` on error. `GetLastError` method can be use to retrieve error details
' @remark Can be used to poll device while displaying "Out of paper" error message.
'
Public Function GetDeviceStatus(CurrentStatus As String) As Boolean
Attribute GetDeviceStatus.VB_Description = "Used to retrieve current device status without performing any other operation"

End Function

