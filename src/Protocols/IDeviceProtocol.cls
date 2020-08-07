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
' Abstract interface for fiscal device protocol handling
'
'=========================================================================
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
Option Explicit

Public Enum UcsFiscalReceiptTypeEnum
    ucsFscRcpSale = 1
    ucsFscRcpReversal
    ucsFscRcpInvoice
    ucsFscRcpCreditNote
    ucsFscRcpOrderList
    [_ucsFscRcpMax]
End Enum

Public Enum UcsFiscalDiscountTypeEnum
    ucsFscDscPlu = 1
    ucsFscDscSubtotal
    ucsFscDscTotal
    ucsFscDscSubtotalAbs
    [_ucsFscDscMax]
End Enum

Public Enum UcsFiscalBarcodeTypeEnum
    ucsFscBrcEan8 = 1
    ucsFscBrcEan13
    ucsFscBrcCode128
    ucsFscBrcQRcode
    [_ucsFscBrcMax]
End Enum

Public Enum UcsFiscalPaymentTypeEnum
    ucsFscPmtCash = 1           ' SCash
    ucsFscPmtCheque = 2         ' SChecks
    ucsFscPmtCoupon = 3         ' ST
    ucsFscPmtVoucher = 4        ' SOT
    ucsFscPmtPackaging = 5      ' SP
    ucsFscPmtMaintenance = 6    ' SSelf
    ucsFscPmtDamage = 7         ' SDmg
    ucsFscPmtCard = 8           ' SCards
    ucsFscPmtBank = 9           ' SW
    ucsFscPmtCustom1 = 10       ' SR1 ("Резерв 1"/ отложено плащане)
    ucsFscPmtCustom2 = 11       ' SR2 ("Резерв 2" / вътрешно потребление)
    ucsFscPmtEUR = 12
    [_ucsFscPmtMax]
End Enum

Public Enum UcsFiscalLocalizedIndexesEnum
    ucsFscLciInternalErrors
    ucsFscLciPrinterStatuses
    ucsFscLciPrinterErrors
    ucsFscLciMemoryTexts
    ucsFscLciDipSwitchesTexts
    ucsFscLciReceiptTexts
    ucsFscLciExtendedErrors
    ucsFscLciConnectorErrors
    [_ucsFscLciMax]
End Enum

Public Enum UcsFiscalReportsTypeEnum
    ucsFscRstDaily
    ucsFscRstDailyByItems
    ucsFscRstDailyByDepartments
    ucsFscRstDailyByItemsAndDepartments
    ucsFscRstDailyByOperators
    ucsFscRstMonthlyShort
    ucsFscRstMonthlyDetailed
End Enum

Public Enum UcsFiscalErrorsEnum
    ucsFerNone
    ucsFerGeneralError
    ucsFerInvalidPassword
    ucsFerPasswordNotSet
End Enum

Public Enum UcsFiscalReversalTypeEnum
    ucsFscRevOperatorError
    ucsFscRevRefund
    ucsFscRevTaxBaseReduction
End Enum

Public Enum UcsFiscalTaxNoTypeEnum
    ucsFscTxnEIC                        ' (default) a.k.a. Bulstat
    ucsFscTxnCitizenNo
    ucsFscTxnForeignerNo
    ucsFscTxnOfficialNo
End Enum

Public Enum UcsFiscalStatusCodeEnum
    ucsFscStaReady
    ucsFscStaBusy
    ucsFscStaFailed
End Enum

''
' Used to initialize the protocol handler. Here the COM port is opened and initial device configuration is retrieved
'
' @param `DeviceString`     [input] semi-colon delimited key=value pairs as "Port=COM1;Speed=9600[;BaudRate=num][;Persistent=bool]
'                                   [;DataBits=num] [;Parity=bool][;StopBits=num][;Timeout=num][;RowChars=num]"
' @param `EventSink`        [in, optional] reference to sink object for protocol events
' @return `True` on success, `False` on error. `GetLastError` method can be use to retrieve error details
'
Public Function Init(DeviceString As String, Optional EventSink As Object) As Boolean
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
' @param `ErrNumber`        [out, optional] Error category of the last error
' @return Text description of the last error
'
Public Function GetLastError(Optional ErrNo As UcsFiscalErrorsEnum) As String
Attribute GetLastError.VB_Description = "Used to retrieve details about last error that occurred during protocol operations"

End Function

''
' Used to begin printing of a new receipt. The actual printing of the receipt is usually postponed until `EndReceipt`
' is called
'
' @param `ReceiptType`      [input] designates if receipt if the receipt is an invoice or regular sales receipt
' @param `OperatorCode`     [input] operator number to be authenticated (usually 1-8)
' @param `OperatorName`     [input] operator name to be printed on receipt
' @param `OperatorPassword` [input] operator password, can be `Empty` if default for device
' @param `UniqueSaleNo`     [input] unique sale number
' @param `TableNo`          [input] table no (for order-list printing only)
' @param `DisablePrinting`  [input] disable printing (for FDType=11 or 21 only)
' @param `InvDocNo`         [in, optional] number of the invoice/credit note to be printed
' @param `InvCgTaxNo`       [in, optional] identification number of the company (country specific)
' @param `InvCgTaxNoType`   [in, optional] type of identification number supplied: bulstat, EGN, foreigner or custom
' @param `InvCgVatNo`       [in, optional] VAT identification number of the company (EU specific)
' @param `InvCgName`        [in, optional] name of the company that is receiving the document
' @param `InvCgCity`        [in, optional] city of the company that is receiving the document, can be empty string
' @param `InvCgAddress`     [in, optional] address of the company that is receiving the document
' @param `InvCgPrsReceive`  [in, optional] name of the person that is receiving the document
' @param `OwnData`          [in, optional] owner info: name, address, eik, dep name, dep address, footer1, footer2
' @param `RevType`          [in, optional] operator error, refund or tax-base deduction
' @param `RevReceiptNo`     [in, optional] number of reversed receipt
' @param `RevReceiptDate`   [in, optional] date/time of reversed receipt
' @param `RevFiscalMemoryNo` [in, optional] fiscal memory of reversed receipt
' @param `RevInvoiceNo`     [in, optional] number of credited invoice
' @param `RevReason`        [in, optional] credit note reason as plain-text (quantity or price adjustment)
' @return `True` on success, `False` on error. `GetLastError` method can be use to retrieve error details
'
Public Function StartReceipt( _
            ByVal ReceiptType As UcsFiscalReceiptTypeEnum, _
            OperatorCode As String, _
            OperatorName As String, _
            OperatorPassword As String, _
            Optional UniqueSaleNo As String, _
            Optional TableNo As String, _
            Optional ByVal DisablePrinting As Boolean, _
            Optional InvDocNo As String, _
            Optional InvCgTaxNo As String, _
            Optional ByVal InvCgTaxNoType As UcsFiscalTaxNoTypeEnum, _
            Optional InvCgVatNo As String, _
            Optional InvCgName As String, _
            Optional InvCgCity As String, _
            Optional InvCgAddress As String, _
            Optional InvCgPrsReceive As String, _
            Optional OwnData As String, _
            Optional ByVal RevType As UcsFiscalReversalTypeEnum, _
            Optional RevReceiptNo As String, _
            Optional ByVal RevReceiptDate As Date, _
            Optional RevFiscalMemoryNo As String, _
            Optional RevInvoiceNo As String, _
            Optional RevReason As String) As Boolean
Attribute StartReceipt.VB_Description = "Used to begin printing of a new receipt. The actual printing of the receipt is usually postponed until `EndReceipt` is called"
            
End Function

''
' Used to print a row on the final receipt with product description, price and quantity
'
' @param `Name`             [input] full product name or description
' @param `Price`            [input] final sale price (incl. VAT) rounded to 0.00. Can be negative for voids
' @param `Quantity`         [input] quantity rounded to 0.000
' @param `TaxGroup`         [in, optional] VAT group the row total is accumulated to. VAT total will be printed at the end of receipt
' @param `UnitOfMeasure`    [in, optional] unit of measure for `Quantity` value
' @param `DepartmentNo`     [in, optional] department number
' @return `True` on success, `False` on error. `GetLastError` method can be use to retrieve error details
' @remark When printing negative rows currently reached receipt subtotal cannot become negative or the protocol handler
' will reorder receipt rows to prevent fiscal device from rejecting the whole receipt.
'
Public Function AddPLU( _
            Name As String, _
            ByVal Price As Double, _
            Optional ByVal Quantity As Double = 1, _
            Optional ByVal TaxGroup As Long = 2, _
            Optional UnitOfMeasure As String, _
            Optional ByVal DepartmentNo As Long) As Boolean
Attribute AddPLU.VB_Description = "Used to print a row on the final receipt with product description, price and quantity"
            
End Function

''
' Used to print a row with plain-text
'
' @param `Line`             [input] row text
' @param `Command`          [input] text format. One of "DBWIDTH" or "CONDENSED". Not all devices support this param
' @param `WordWrap`         [input] if `True` text will be wrapped and can be printed on multiple rows
' @return `True` on success, `False` on error. `GetLastError` method can be use to retrieve error details
'
Public Function AddLine( _
            Line As String, _
            Optional Command As String, _
            Optional ByVal WordWrap As Boolean = True) As Boolean
Attribute AddLine.VB_Description = "Used to print a row with plain-text"
            
End Function

''
' Used to discount/surcharge previous PLU row, or to print discount/surcharge on current subtotal or receipt total
'
' @param `DiscType`         [input] type of discount/surcharge. See `UcsFiscalDiscountTypeEnum`
' @param `Value`            [input] discount/srcharge percent in range -100.00 to 100.00 (positive for discount)
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
' @param `PmtType`          [input] payment type. See `UcsFiscalPaymentTypeEnum`
' @param `PmtName`          [in, optional] payment name if `PmtType` is one of custom payment types
' @param `Amount`           [in, optional] positive payment amount. If skipped (or 0.00) pays the remaining receipt
' @return `True` on success, `False` on error. `GetLastError` method can be use to retrieve error details
' @remark After adding a payment no more PLUs can be printed. Several payments can be accumulated on a single receipt
' e.g. 10.00 by card and 50.00 in cash, etc. are able to pay out a 35.00 total receipt
'
Public Function AddPayment( _
            ByVal PmtType As UcsFiscalPaymentTypeEnum, _
            Optional PmtName As String, _
            Optional ByVal Amount As Double, _
            Optional ByVal Rate As Double) As Boolean
Attribute AddPayment.VB_Description = "Used to finalize PLU rows and print payment information"
            
End Function

''
' Used to print barcode
'
' @param `BarcodeType`      [input] type of barcode. See `UcsFiscalBarcodeTypeEnum`
' @param `Text`             [input] barcode contents. Allowed symbols depend on type of barcode
' @param `Height`           [input] barcode height in mm
' @return `True` on success, `False` on error. `GetLastError` method can be use to retrieve error details
'
Public Function AddBarcode( _
            ByVal BarcodeType As UcsFiscalBarcodeTypeEnum, _
            Text As String, _
            Optional ByVal Height As Long) As Boolean
Attribute AddBarcode.VB_Description = "Used to print barcode"
            
End Function

''
' Used to finalize receipt payments and batch print the complete receipt
'
' @param `ResumeToken`      [in, optional]  if previous invocation failed can be used to resume printing on the same line
' @param `ReceiptNo`        [out, optional] receipt number of the receipt just printed
' @param `ReceiptDate`      [out, optional] date/time of the receipt just printed
' @param `ReceiptAmount`    [out, optional] total amount of the receipt just printed
' @param `InvoiceNo`        [out, optional] invoice number of the extended receipt printed
' @return `True` on success, `False` on error. `GetLastError` method can be use to retrieve error details
' @remark All previously added payments must add up to enough total amount so to be able to pay out all the PLU rows
' incl. discounted/surcharged rows
'
Public Function PrintReceipt( _
            Optional ResumeToken As String, _
            Optional ReceiptNo As String, _
            Optional ReceiptDate As Date, _
            Optional ReceiptAmount As Double, _
            Optional InvoiceNo As String) As Boolean

End Function

''
' Used to print a duplicate of the last fiscal receipt
'
' @param `DuplicateNo`      [in, optional]  uses EJ to copy receipt by number if set
' @param `ReceiptNo`        [out, optional] receipt number of the receipt just printed
' @param `ReceiptDate`      [out, optional] date/time of the receipt just printed
' @return `True` on success, `False` on error. `GetLastError` method can be use to retrieve error details
' @remark Most fiscal devices are able to print a single duplicate of last fiscal receipt only. On second copy request
' an error is returned.
'
Public Function PrintDuplicate( _
            Optional DuplicateNo As String, _
            Optional ReceiptNo As String, _
            Optional ReceiptDate As Date) As Boolean

End Function

''
' Used to print fiscal x-report
'
' @param `ReportType`       [input] type of detailed report to print: by products, by departments, by both
' @param `ReceiptNo`        [out, optional] receipt number of the receipt just printed
' @param `ReceiptDate`      [out, optional] date/time of the receipt just printed
' @return `True` on success, `False` on error. `GetLastError` method can be use to retrieve error details
'
Public Function PrintDailyXReport( _
            ByVal ReportType As UcsFiscalReportsTypeEnum, _
            Optional ReceiptNo As String, _
            Optional ReceiptDate As Date) As Boolean

End Function

''
' Used to print fiscal z-report and clear fiscal memory for the day
'
' @param `ReportType`       [input] type of detailed report to print: by products, by departments, by both
' @param `ReceiptNo`        [out, optional] receipt number of the receipt just printed
' @param `ReceiptDate`      [out, optional] date/time of the receipt just printed
' @return `True` on success, `False` on error. `GetLastError` method can be use to retrieve error details
'
Public Function PrintDailyZReport( _
            ByVal ReportType As UcsFiscalReportsTypeEnum, _
            Optional ReceiptNo As String, _
            Optional ReceiptDate As Date) As Boolean

End Function

''
' Used to print monthly fiscal report by date range
'
' @param `ReportType`       [input] type of report to print: short or details
' @param `StartDate`        [input] start of date range
' @param `EndDate`          [input] end of date range
' @param `ReceiptNo`        [out, optional] receipt number of the receipt just printed
' @param `ReceiptDate`      [out, optional] date/time of the receipt just printed
' @return `True` on success, `False` on error. `GetLastError` method can be use to retrieve error details
'
Public Function PrintMonthlyReport( _
            ByVal ReportType As UcsFiscalReportsTypeEnum, _
            ByVal StartDate As Date, _
            ByVal EndDate As Date, _
            Optional ReceiptNo As String, _
            Optional ReceiptDate As Date) As Boolean

End Function

''
' Used to debit (and credit) cashdesk outside of regular sales and daily z-report with memory clear
'
' @param `OperatorCode`     [input] operator number to be authenticated (usually 1-8)
' @param `OperatorPassword` [input] operator password, can be `Empty` if default for device
' @param `Amount`           [input] positive for debit, negative for credit, 0.00 for info only
' @param `ReceiptNo`        [out, optional] receipt number of the receipt just printed
' @param `ReceiptDate`      [out, optional] date/time of the receipt just printed
' @param `ReceiptAmount`    [out, optional] total amount of the receipt just printed
' @return Array of doubles with (TotalAvailable, TotalDebit, TotalCredit)
'
Public Function PrintServiceDeposit( _
            OperatorCode As String, _
            OperatorPassword As String, _
            ByVal Amount As Double, _
            Optional ReceiptNo As String, _
            Optional ReceiptDate As Date, _
            Optional ReceiptAmount As Double) As Variant
    
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
' @param `NewDate`          [input] new date/time to set to the clock
' @return `True` on success, `False` on error. `GetLastError` method can be use to retrieve error details
'
Public Function SetClock(ByVal NewDate As Date) As Boolean
Attribute SetClock.VB_Description = "Used to set current date/time of the clock on the fiscal device"

End Function

''
' Used to retrieve the receipt number of the last receipt printed. This includes fiscal receipts, invoices, fiscal
' reports and service deposits/withdraw
'
' @return Last receipt number as string
'
Public Function GetLastReceiptNo() As String
Attribute GetLastReceiptNo.VB_Description = "Used to retrieve the receipt number of the last fiscal slip printed. This includes sales receipts, invoices, fiscal reports and service deposits/withdraw"

End Function

''
' Used to retrieve last invoice number printed which has come from current fiscal device invoices range
'
' @return Last invoice number as string
'
Public Function GetLastInvoiceNo() As String
Attribute GetLastInvoiceNo.VB_Description = "Used to retrieve last invoice number printed which has come from current fiscal device invoices range"

End Function

''
' Used to retrieve the receipt number and date/time of the last sale/invoice printed.
'
' @return Last receipt number as string
'
Public Function GetLastQRCodeInfo(Optional LastReceiptDate As Date) As String
Attribute GetLastQRCodeInfo.VB_Description = "Used to retrieve the receipt  number and date/time of the last sale/invoice printed."

End Function

''
' Used to retrieve serial number of the fiscal device as reported by firmware
'
' @return Serial number as string
'
Public Function GetDeviceSerialNo() As String
Attribute GetDeviceSerialNo.VB_Description = "Used to retrieve serial number of the fiscal device as reported by firmware"

End Function

''
' Used to retrieve number of the registered fiscal memory
'
' @return Fiscal number as string
'
Public Function GetFiscalMemoryNo() As String
Attribute GetFiscalMemoryNo.VB_Description = "Used to retrieve number of the registered fiscal memory"

End Function

''
' Used to set localized strings used for non-fiscal receipt texts and component internal errors
'
' @param `Index`            [input] type of localized strings. See `UcsFiscalLocalizedIndexesEnum`
'
Public Sub SetLocalizedText(ByVal Index As UcsFiscalLocalizedIndexesEnum, Text As String, Optional ByVal ConfigIndex As Long)
Attribute SetLocalizedText.VB_Description = "Used to set localized strings used for non-fiscal receipt texts and component internal errors"

End Sub

''
' Used to retrieve tax number of the company that the fiscal device belongs to
'
' @return Company's tax number as string
'
Public Function GetTaxNo() As String
Attribute GetTaxNo.VB_Description = "Used to retrieve tax number of the company that the fiscal device belongs to"

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
' @param `Index`            [input] index of the header row to be retrieved
' @return Template text
'
Public Function GetHeaderText(ByVal Index As Long) As String
Attribute GetHeaderText.VB_Description = "Used to retrieve template text for the header rows printed on receipts"

End Function

''
' Used to retrieve template text for the footer rows printed on receipts
'
' @param `Index`            [input] index of the footer row to be retrieved
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
' @param `Index`            [input] type of payment to be retrieved
' @return Localized payment name
'
Public Function GetPaymentName(ByVal Index As Long) As String
Attribute GetPaymentName.VB_Description = "Used to retrieve standard and custom payment names"

End Function

''
' Used to retrieve width of the paper roll in device characters
'
' @return Total number of characters printable on paper roll
'
Public Function GetCharsPerLine() As Long
Attribute GetCharsPerLine.VB_Description = "Used to retrieve width of the paper roll in device characters"

End Function

''
' Used to retrieve maximum width of plain-text rows in device characters
'
' @return Maximum number of characters available for a comment row
'
Public Function GetCommentTextMaxLength() As Long
Attribute GetCommentTextMaxLength.VB_Description = "Used to retrieve maximum width of plain-text rows in device characters"

End Function

''
' Used to retrieve maximum width of product names. The device might wrap names on two rows when printing
'
' @return Maximum number of characters available for product name
'
Public Function GetItemNameMaxLength() As Long
Attribute GetItemNameMaxLength.VB_Description = "Used to retrieve maximum width of product names. The device might wrap names on two rows when printing"

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
' Used to retrieve operator name as currently configured
'
' @param `OperatorCode`     [input] operator number to be authenticated (usually 1-8)
' @return Operator name
'
Public Function GetOperatorName(OperatorCode As String) As String
Attribute GetOperatorName.VB_Description = "Used to retrieve operator name as currently configured"

End Function

''
' Used to retrieve default operator password as specified in fiscal device documentation
'
' @param `OperatorCode`     [input] operator number to be authenticated (usually 1-8)
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
' Used to retrieve device model name as reported by fiscal device firmware
'
' @return Device model name
'
Public Function GetDeviceModel() As String
Attribute GetDeviceModel.VB_Description = "Used to retrieve device model name as reported by fiscal device firmware"

End Function

''
' Used to retrieve device firmware version and revision
'
' @return Device firmware version
'
Public Function GetDeviceFirmware() As String
Attribute GetDeviceFirmware.VB_Description = "Used to retrieve device firmware version and revision"

End Function

''
' Used to retrieve totals by VAT groups since last z-report
'
' @return Array of totals for each VAT group (1-8)
'
Public Function GetTotalsByTaxGroups() As Variant
Attribute GetTotalsByTaxGroups.VB_Description = "Used to retrieve totals by VAT groups since last z-report"

End Function

''
' Used to retrieve totals by payment types since last z-report
'
' @return Array of totals for each payment type (1-8): ucsFscPmtCash, ucsFscPmtCard, ucsFscPmtBank,
'         reserved, ucsFscPmtCustom1, ucsFscPmtCustom2, ucsFscPmtCustom3 and ucsFscPmtCustom4
' @remark Index 4 in return array is unused and should be 0
'
Public Function GetTotalsByPaymentTypes() As Variant
Attribute GetTotalsByPaymentTypes.VB_Description = "Used to retrieve totals by payment types since last z-report"

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
' @param `Ports`            [input] Array of strings with COM ports to check
' @return Array of arrays with COM ports and auto-detected devices info.
'
Public Function AutodetectDevices(Ports As Variant) As Variant
Attribute AutodetectDevices.VB_Description = "Used to auto-detect devices that implemented current protocol on the selected COM ports"

End Function

''
' Used to retrieve resume-token after failed `EndReceipt`
'
' @return String encoded resume-token
' @remark Result can be passed to `EndReceipt` to resume printing exactly from the previsouly failed row e.g. if the
' device ran out-of-paper the operator can change the roll and resume printing and later print a brand new duplicate
' receipt for the end client.
'
Public Function GetResumeToken() As String
Attribute GetResumeToken.VB_Description = "Used to retrieve resume-token after failed `EndReceipt`"

End Function

''
' Used to retrieve current device status without performing any other operation
'
' @param `CurrentStatus`    [output] device error or empty string for ready
' @return `True` on success, `False` on error. `GetLastError` method can be use to retrieve error details
' @remark Can be used to poll device while displaying "Out of paper" error message.
'
Public Function GetDeviceStatus(StatusText As String) As UcsFiscalStatusCodeEnum
Attribute GetDeviceStatus.VB_Description = "Used to retrieve current device status without performing any other operation"

End Function

''
' Used to retrieve department name as currently configured
'
' @param `DepartmentCode`     [input] department number
' @return Department name
'
Public Function GetDepartmentName(DepartmentCode As String, Optional TaxGroup As Long) As String
Attribute GetDepartmentName.VB_Description = "Used to retrieve department name as currently configured"

End Function
''
' Used to retrieve current device VAT rates
'
' @return Array of VAT rates for each VAT group (1-8)
'
Public Function GetTaxRates() As Variant

End Function
