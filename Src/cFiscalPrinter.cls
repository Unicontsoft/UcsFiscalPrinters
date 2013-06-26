VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFiscalPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=========================================================================
' $Header: /UcsFiscalPrinter/Src/cFiscalPrinter.cls 29    26.06.13 10:13 Wqw $
'
'   Unicontsoft Fiscal Printers Project
'   Copyright (c) 2008-2013 Unicontsoft
'
'   Fiscal devices management using IDeviceProtocol protocols
'
' $Log: /UcsFiscalPrinter/Src/cFiscalPrinter.cls $
' 
' 29    26.06.13 10:13 Wqw
' REF: zeka internal err msgs
'
' 28    20.06.13 12:49 Wqw
' REF: handle custom payment types
'
' 27    18.06.13 17:17 Wqw
' ADD: Function GetTotalsByPayments, GetTotalsStats
'
' 26    14.03.13 16:28 Wqw
' REF: err handling on PrintReceipt
'
' 25    26.02.13 19:39 Wqw
' REF: language consts
'
' 24    8.01.13 15:43 Wqw
' REF: err handling, returns receipt no as string
'
' 23    4.01.13 12:23 Wqw
' REF: err handling
'
' 22    3.01.13 16:41 Wqw
' ADD: Function GetStatus. REF: handles resume tickets on print receipt
'
' 21    16.11.12 18:53 Wqw
' REF: PrintReceipt err handler returns warning i fiscal payment has been
' executed
'
' 20    19.10.12 0:36 Wqw
' REF: rename report type enum
'
' 19    9.10.12 15:21 Wqw
' REF: params of CashDebitCredit
'
' 18    5.10.12 14:20 Wqw
' ADD: Function CashDebitCredit
'
' 17    29.08.12 15:06 Wqw
' REF: localized texts can by customized in config
'
' 16    6.08.12 18:41 Wqw
' REF: impl EnumPorts with AutodetectDevices
'
' 15    23.03.12 15:30 Wqw
' ADD: GetTotalsByVatGroups. REF: PrintReceipt checks if receipt already
' printed
'
' 14    8.12.11 15:49 Wqw
' REF: multi-speed check in EnumPorts
'
' 13    9.08.11 23:25 Wqw
' REF: supports zeka protocol
'
' 12    17.06.11 13:23 Wqw
' REF: enum casing
'
' 11    10.05.11 15:15 Wqw
' REF: guard GetCommandLog in err handlers
'
' 10    8.03.11 13:04 Wqw
' REF: vat group idwa numeric weche
'
' 9     23.02.11 17:10 Wqw
' REF: text const
'
' 8     22.02.11 17:38 Wqw
' REF: text consts
'
' 7     22.02.11 13:52 Wqw
' REF: EnumPorts wryshta doplnitelno speed i protocol
'
' 6     22.02.11 10:50 Wqw
' REF: GetHeaderFooter wryshta array s empty ako ne moje da dostypi FP
'
' 5     22.02.11 10:26 Wqw
' ADD: EnumPorts
'
' 4     22.02.11 10:07 Wqw
' REF: izneseni metodi w otdelen admin class
'
' 3     21.02.11 16:28 Wqw
' ADD: Property IsShared
'
' 2     21.02.11 13:46 Wqw
' REF: impl datecs fp and daisy ecr support
'
' 1     14.02.11 18:13 Wqw
' Initial implementation
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "cFiscalPrinter"

'=========================================================================
' Public enums
'=========================================================================

Public Enum UcsFiscalPrintReportTypeEnum
    ucsFscRptDaily = 1
    ucsFscRptNumber
    ucsFscRptDate
    ucsFscRptOperator
End Enum

Public Enum UcsFiscalParamIndexesEnum
    ucsFscPixItem = 0
    ucsFscPixPrice
    ucsFscPixGroup
    ucsFscPixQuantity
    ucsFscPixDiscount
    ucsFscPixWayOfPmt = ucsFscPixPrice
    ucsFscPixAmount
    ucsFscPixInvoiceNo = 0
    ucsFscPixCgName
    ucsFscPixCgBulstat
    ucsFscPixCgVatNo
    ucsFscPixCgAddress
    ucsFscPixCgPersonReceiver
    ucsFscPixOwnPersonSeller
    ucsFscPixLnkDoc
    ucsFscPixInvoiceNoPrefix
    ucsFscPixOperatorCode = 0
    ucsFscPixOperatorName
    ucsFscPixOperatorPass
    ucsFscPixCashDeskNo
    ucsFscPixLastReceipt
    ucsFscPixPrevReceipt
    ucsFscPixPrevCount
    ucsFscPixResumeTicket
End Enum

Public Enum UcsFiscalHeaderIndexesEnum
    ucsFscHdrHeader1 = 0
    ucsFscHdrHeader2
    ucsFscHdrBulstat
    ucsFscHdrHeader3
    ucsFscHdrHeader4
    ucsFscHdrHeader5
    ucsFscHdrHeader6
    ucsFscHdrFooter1
    ucsFscHdrFooter2
    ucsFscHdrOperatorInfo
    ucsFscHdrDateTime
    ucsFscHdrLastReceipt
    ucsFscHdrWop1
    ucsFscHdrWop2
    ucsFscHdrWop3
    ucsFscHdrWopCustom1
    ucsFscHdrWopCustom2
    ucsFscHdrWopCustom3
    ucsFscHdrCharsPerLine
    ucsFscHdrDefOperPassword
    [_ucsFscHdrLast]
End Enum

'=========================================================================
' Constants and member variables
'=========================================================================

'--- strings
Private Const STR_ELTRADE_ECR_INTERNAL  As String = "Не е указано устройство|Грешка при отваряне: |Грешка при SetCommTimeouts: |Грешка при BuildCommDCB: |Грешка при SetCommState: |Грешка при WriteFile: |Време за достъп изтече в очакване на отговор|Грешка при ReadFile: |Вече има отворена връзка|Няма отворена връзка|Невалидно заглавие на отговор от устройството|Грешка при WaitCommEvent: |Няма започната бележка|Липсва хартия|Невалиден параметър 'Command'|Непозволена команда за печат на повече от един дубликат|Грешка при печат|Фатална грешка на фискалното устройство|Невъзможно отказване на предишна транзакция|Грешка на ред %1: %2"
Private Const STR_ELTRADE_ECR_TEXTS     As String = "СУМА|МЕЖДИННА СУМА|Продажби %1|ед. цена|количество|сума*%1|РЕСТО|всичко|ДДС*%1|нето ст-ст|Фактура #|Продавач:|   * * * * * * * *|Получател:|ИДЕНТ # |ЗДДС # |%1 артикули|МОЛ:|Получил:|НАДБАВКА|ОТСТЪПКА"
Private Const STR_DATECS_FP_INTERNAL    As String = "Не е указано устройство|Грешка при отваряне: |Грешка при SetCommTimeouts: |Грешка при BuildCommDCB: |Грешка при SetCommState: |Грешка при WriteFile: |Време за достъп изтече в очакване на отговор|Грешка при ReadFile: |Невалиден формат на съобщение или сума за проверка (NAK)|Вече има отворена връзка|Няма отворена връзка|Невалидна дължина на отговора: %1|Липсва символ за край на отговора: %1|Липсва символ за начало на отговора: %1|Грешка при WaitCommEvent: %1|Няма започната бележка| или невалидна парола на оператор|Грешка на ред %1: %2|Невъзможно отказване на предишна транзакция|Невалиден тип справка|Грешка %1|Памет %1"
Private Const STR_DATECS_FP_STATUSES    As String = "Синтактична грешка|Невалидна команда|Неустановени дата и час|3|Неизправност в механизма на печатащото устройство|5|6|7|Аритметично препълване|Непозволена команда|Зануляване на опертивна памет|3|Разрушено съдържание на опертивна памет|5|6|7|Няма хартия|1|2|Отворен фискален бон|4|Отворен служебен (нефискален) бон|6|7"
Private Const STR_DATECS_FP_ERRORS      As String = "Синтактична грешка|Невалидна команда|2|3|Неизправност в механицва на печатащото устройство|5|6|7|1|Непозволена команда|Зануляване на опертивна памет|3|Разрушено съдържание на опертивна памет|5|6|7|Няма хартия|1|2|3|4|5|6|7"
Private Const STR_DATECS_FP_DIP_SWITCHES As String = "Автоматично центриране на header и footer|Предварителен header|Sw1.3|Sw1.4|Нулиране на паметта|Прозрачен дисплей|Без данни на дисплея|7"
Private Const STR_DATECS_FP_MEMORY      As String = "Грешка при запис|1|Няма модул фискална памет|Малко свободно място във фискалната памет|Пълна фискална памет|5|6|7|Фискалната памет забранена за запис|Форматирана фискална памет|2|Принтерът е във фискален режим|Зададени данъчни ставки|Програмирани индивидуален номер и номер на фискалната памет|6|7"
Private Const STR_DATECS_FP_TEXTS       As String = "В БРОЙ|С ДЕБИТНА КАРТА|С ЧЕК|С КРЕДИТНА КАРТА|НАДБАВКА %1|ОТСТЪПКА %1|СУМА|МЕЖДИННА СУМА|ФАКТУРА No %1|Продажби %1|ЕДИНИЧНА ЦЕНА|КОЛИЧЕСТВО|СУМА|ВСИЧКО ГРУПА %1|ДДС %1=%2|НЕТО СТОЙНОСТ|ОБЩО|%1 АРТИКУЛА|1 АРТИКУЛ|РЕСТО|ПРОДАВАЧ: |ПОЛУЧАТЕЛ: |КУПУВАЧ: |ИДЕНТ. No: |ЗДДС No: "
Private Const STR_ZEKA_FP_INTERNAL      As String = "Не е указано устройство|Грешка при отваряне: %1|Грешка при SetCommTimeouts: %1|Грешка при BuildCommDCB: %1|Грешка при SetCommState: %1|Грешка при WriteFile: %1|Време за достъп изтече в очакване на отговор|Грешка при ReadFile: %1|Невалиден формат на съобщение или сума за проверка (NAK)|Вече има отворена връзка|Няма отворена връзка|Невалидна дължина на отговора: %1|Липсва символ за начало на отговора: %1|Грешка при WaitCommEvent: %1|Няма започната бележка|Грешка на ред %1: %2|Невъзможно отказване на предишна транзакция|Невалиден тип справка|Устройството е недостъпно|Устройството е заето| или невъзможно установяване парола на оператор| или невалидна парола на оператор"
Private Const STR_ZEKA_FP_STATUSES      As String = "Неизвестна грешка|Невалидна команда|Непозволена команда|Непозволена поради ненулев отчет|Синтактична грешка|Препълване на входните регистри|Нулев входен регистър|Липсва транзакция която да се войдира|Недостатъчна налична сума|Конфликт в данъчните групи"
Private Const STR_ZEKA_FP_ERRORS        As String = "Неизвестна грешка|Няма хартия|Препълване на дневните регистри|Несверен/грешен часовник|Отворен фискален бон|Сметка с остатък за плащане|Отворен нефискален бон|Сметка с приключено плащане|Фискална памет само за четене|Грешна парола или непозволена команда|Липсващ външен дисплей|24 часа без дневен отчет|Прегрят принтер|Спад на напрежение във фискален бон|Препълване в електронната контролна лента|Недостатъчни условия"
Private Const STR_ZEKA_FP_TEXTS         As String = "В брой|С карта|С чек|Надбавка%1|Отстъпка%1|Сума|Междинна сума|Фактура No %1|Продажби %1|Единична цена|Количество|Сума*%1|Всичко група %1|ДДС*%1=%2|Нето стойност|Обща сума|%1 Продажби|1 Продажба|Ресто|Продавач: |Получател: |Купувач: |Идент. No: |ЗДДС No: |ЕИК;ЕГН;ЛНЧ;Служебен No|КОРЕКЦИЯ"
Private Const STR_DUPLICATE             As String = "ДУБЛИКАТ"
Private Const STR_ORIGINAL              As String = "ОРИГИНАЛ"
Private Const STR_HEAD_CREDIT_NOTE2     As String = "към ф-ра "
Private Const STR_CASHIER               As String = "Касиер"
'--- errors
Private Const ERR_INIT_PROTOCOL         As String = "Грешка при инициализация на протокол %1"
Private Const ERR_MISSING_ROWS          As String = "Липсват редове за печат"
Private Const ERR_NOT_IMPLEMENTED       As String = "Командата не се поддържа от протокола"
'--- config entries
Private Const CFG_INTERNAL_ERRORS       As String = "InternalErrors"
Private Const CFG_PRINTER_STATUSES      As String = "PrinterStatuses"
Private Const CFG_PRINTER_ERRORS        As String = "PrinterErrors"
Private Const CFG_DIP_SWITCHES_TEXTS    As String = "DipSwitchesTexts"
Private Const CFG_MEMORY_TEXTS          As String = "MemoryTexts"
Private Const CFG_RECEIPT_TEXTS         As String = "ReceiptTexts"

Private Enum UcsRowTypeEnum
    ucsRwtSell = 1
    ucsRwtPayment
    ucsRwtText
End Enum

Private Type UcsRow
    Type            As UcsRowTypeEnum
    Item            As String
    Price           As String
    Discount        As String
    Group           As String
    Quantity        As String
    WayOfPayment    As String
    Amount          As String
End Type

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunc As String)
    Debug.Print MODULE_NAME & "." & sFunc & ": " & Err.Description
    OutputDebugLog MODULE_NAME, sFunc & "(" & Erl & ")", "Run-time error: " & Err.Description
End Sub

Private Sub DebugLog(sFunc As String, sText As String)
    OutputDebugLog MODULE_NAME, sFunc, sText
End Sub

'=========================================================================
' Methods
'=========================================================================

Public Function EnumPorts() As Variant
    Const FUNC_NAME     As String = "EnumPorts"
    Dim vPorts          As Variant
    Dim vElem           As Variant
    Dim oFP             As IDeviceProtocol
    Dim lIdx            As Long
    
    On Error GoTo EH
    vPorts = EnumSerialPorts
    For Each vElem In Array(New cZekaProtocol, New cICLProtocol, New cEltradeProtocol)
        Set oFP = vElem
        vPorts = oFP.AutodetectDevices(vPorts)
    Next
    For lIdx = 0 To UBound(vPorts)
        If Not IsArray(vPorts(lIdx)) Then
            vPorts(lIdx) = Array(vPorts(lIdx), vbNullString, vbNullString, vbNullString)
        End If
    Next
    EnumPorts = vPorts
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Public Function GetHeaderFooter( _
            DeviceString As String, _
            Optional OperatorData As Variant, _
            Optional RetVal As Variant, _
            Optional LastError As String, _
            Optional CommandLog As String) As Boolean
    Const FUNC_NAME     As String = "GetHeaderFooter"
    Dim oFP             As IDeviceProtocol
    Dim vInput          As Variant
    
    On Error GoTo EH
    LastError = vbNullString
    vInput = RetVal
    ReDim RetVal(0 To [_ucsFscHdrLast] - 1) As Variant
    Set oFP = pvInit(DeviceString)
    If LenB(At(vInput, ucsFscHdrHeader1, "1")) <> 0 Then
        RetVal(ucsFscHdrHeader1) = Trim$(oFP.GetHeaderText(1))
    End If
    If LenB(At(vInput, ucsFscHdrHeader2, "1")) <> 0 Then
        RetVal(ucsFscHdrHeader2) = Trim$(oFP.GetHeaderText(2))
    End If
    If LenB(At(vInput, ucsFscHdrBulstat, "1")) <> 0 Then
        RetVal(ucsFscHdrBulstat) = oFP.GetTaxNumber() & "," & Replace(Trim$(oFP.GetTaxCaption()), ":", vbNullString)
    End If
    If LenB(At(vInput, ucsFscHdrHeader3, "1")) <> 0 Then
        RetVal(ucsFscHdrHeader3) = Trim$(oFP.GetHeaderText(3))
    End If
    If LenB(At(vInput, ucsFscHdrHeader4, "1")) <> 0 Then
        RetVal(ucsFscHdrHeader4) = Trim$(oFP.GetHeaderText(4))
    End If
    If LenB(At(vInput, ucsFscHdrHeader5, "1")) <> 0 Then
        RetVal(ucsFscHdrHeader5) = Trim$(oFP.GetHeaderText(5))
    End If
    If LenB(At(vInput, ucsFscHdrHeader6, "1")) <> 0 Then
        RetVal(ucsFscHdrHeader6) = Trim$(oFP.GetHeaderText(6))
    End If
    If LenB(At(vInput, ucsFscHdrFooter1, "1")) <> 0 Then
        RetVal(ucsFscHdrFooter1) = Trim$(oFP.GetFooterText(1))
    End If
    If LenB(At(vInput, ucsFscHdrFooter2, "1")) <> 0 Then
        RetVal(ucsFscHdrFooter2) = Trim$(oFP.GetFooterText(2))
    End If
    If LenB(At(vInput, ucsFscHdrOperatorInfo, "1")) <> 0 Then
        RetVal(ucsFscHdrOperatorInfo) = At(OperatorData, 0)
    End If
    If LenB(At(vInput, ucsFscHdrDateTime, "1")) <> 0 Then
        RetVal(ucsFscHdrDateTime) = oFP.GetClock()
    End If
    If LenB(At(vInput, ucsFscHdrLastReceipt, "1")) <> 0 Then
        RetVal(ucsFscHdrLastReceipt) = oFP.GetLastReceiptNumber()
    End If
    If LenB(At(vInput, ucsFscHdrWop1, "1")) <> 0 Then
        RetVal(ucsFscHdrWop1) = oFP.GetPaymentName(1)
    End If
    If LenB(At(vInput, ucsFscHdrWop2, "1")) <> 0 Then
        RetVal(ucsFscHdrWop2) = oFP.GetPaymentName(2)
    End If
    If LenB(At(vInput, ucsFscHdrWop3, "1")) <> 0 Then
        RetVal(ucsFscHdrWop3) = oFP.GetPaymentName(3)
    End If
    If LenB(At(vInput, ucsFscHdrWopCustom1, "1")) <> 0 Then
        RetVal(ucsFscHdrWopCustom1) = oFP.GetPaymentName(-1)
    End If
    If LenB(At(vInput, ucsFscHdrWopCustom2, "1")) <> 0 Then
        RetVal(ucsFscHdrWopCustom2) = oFP.GetPaymentName(-2)
    End If
    If LenB(At(vInput, ucsFscHdrWopCustom3, "1")) <> 0 Then
        RetVal(ucsFscHdrWopCustom3) = oFP.GetPaymentName(-3)
    End If
    If LenB(At(vInput, ucsFscHdrCharsPerLine, "1")) <> 0 Then
        RetVal(ucsFscHdrCharsPerLine) = oFP.GetCharsPerLine()
    End If
    If LenB(At(vInput, ucsFscHdrDefOperPassword, "1")) <> 0 Then
        RetVal(ucsFscHdrDefOperPassword) = oFP.GetDefaultPassword(At(OperatorData, 0))
    End If
    '--- success
    GetHeaderFooter = True
QH:
    If oFP.IsConnected Then
        oFP.Disconnect
    End If
    CommandLog = oFP.GetCommandLog()
    Exit Function
EH:
    If LenB(LastError) = 0 Then
        LastError = Err.Description
        DebugLog FUNC_NAME, "LastError=" & LastError
        If Not oFP Is Nothing Then
            DebugLog FUNC_NAME, "oFP.GetLastError=" & oFP.GetLastError()
            Resume QH
        End If
    End If
End Function

Public Function PrintReceipt( _
            DeviceString As String, _
            ByVal ReceiptType As UcsFiscalReceiptTypeEnum, _
            Rows As Variant, _
            Optional OperatorData As Variant, _
            Optional InvoiceData As Variant, _
            Optional PrintDuplicate As Boolean, _
            Optional LastError As String, _
            Optional CommandLog As String) As Boolean
    Const FUNC_NAME     As String = "PrintReceipt"
    Dim oFP             As IDeviceProtocol
    Dim vElem           As Variant
    Dim lIdx            As Long
    Dim sPrevReceipt    As String
    Dim sResumeTicket   As String
    Dim eLastNum        As UcsFiscalErrorsEnum
    Dim sLastErr        As String
    
    On Error GoTo EH
    LastError = vbNullString
    sPrevReceipt = At(OperatorData, ucsFscPixPrevReceipt)
    sResumeTicket = At(OperatorData, ucsFscPixResumeTicket)
    Set oFP = pvInit(DeviceString)
    If LenB(sResumeTicket) = 0 Then
        If Not oFP.CancelReceipt() Then
            LastError = oFP.GetLastError()
            GoTo QH
        End If
    End If
    If LenB(sPrevReceipt) <> 0 Then
        ValueAt(OperatorData, ucsFscPixPrevReceipt) = oFP.GetLastReceiptNumber()
        '--- check if receipt already printed
        If LenB(Trim$(sPrevReceipt)) <> 0 Then
            If Pad(sPrevReceipt, -Len(At(OperatorData, ucsFscPixPrevReceipt))) <> At(OperatorData, ucsFscPixPrevReceipt) And C_Lng(At(OperatorData, ucsFscPixPrevCount)) <> 0 Then
                If C_Str(At(OperatorData, ucsFscPixPrevCount)) <> At(oFP.GetTotalsStats(), 0) Then
                    ValueAt(OperatorData, ucsFscPixLastReceipt) = At(OperatorData, ucsFscPixPrevReceipt)
                    ValueAt(OperatorData, ucsFscPixPrevReceipt) = Empty
                    GoTo QH
                End If
            End If
        End If
    End If
    If IsArray(Rows) Then
        oFP.StartReceipt ReceiptType, At(OperatorData, ucsFscPixOperatorCode), _
            At(OperatorData, ucsFscPixOperatorName, STR_CASHIER), _
            At(OperatorData, ucsFscPixOperatorPass), _
            IIf(ReceiptType = ucsFscRetInvoiceNonfiscal, At(InvoiceData, ucsFscPixInvoiceNoPrefix), vbNullString) & At(InvoiceData, ucsFscPixInvoiceNo), _
            At(InvoiceData, ucsFscPixCgBulstat), _
            At(InvoiceData, ucsFscPixCgVatNo), _
            At(InvoiceData, ucsFscPixCgName), _
            vbNullString, _
            At(InvoiceData, ucsFscPixCgAddress), _
            At(InvoiceData, ucsFscPixCgPersonReceiver), _
            vbNullString
        If LenB(At(InvoiceData, ucsFscPixLnkDoc)) <> 0 Then
            oFP.AddLine STR_HEAD_CREDIT_NOTE2 & At(InvoiceData, ucsFscPixLnkDoc)
        End If
        If PrintDuplicate Then
            oFP.AddLine STR_DUPLICATE
        ElseIf ReceiptType = ucsFscRetInvoiceNonfiscal Then
            oFP.AddLine STR_ORIGINAL
        End If
        For Each vElem In Rows
            With pvGetRowData(vElem)
                Select Case .Type
                Case ucsRwtSell
                    oFP.AddPLU .Item, C_Dbl(.Price), IIf(LenB(.Quantity) <> 0, C_Dbl(.Quantity), 1), pvGetVatGroup(.Group)
                    If C_Dbl(.Discount) <> 0 Then
                        oFP.AddDiscount ucsFscDstPlu, -C_Dbl(.Discount)
                    End If
                Case ucsRwtText
                    oFP.AddLine .Item
                Case ucsRwtPayment
                    lIdx = C_Lng(.WayOfPayment)
                    If lIdx < 0 Then
                        lIdx = LimitLong(lIdx, -3, -1)
                    Else
                        lIdx = LimitLong(lIdx, 1, 8)
                    End If
                    oFP.AddPayment lIdx, Zn(.Item, oFP.GetPaymentName(lIdx)), C_Dbl(.Amount)
                End Select
            End With
        Next
        If Not oFP.EndReceipt(sResumeTicket) Then
            If LenB(oFP.GetLastError()) <> 0 Then
                Err.Raise vbObjectError, , oFP.GetLastError()
            End If
        End If
        oFP.OpenDrawer
        '--- success
        PrintReceipt = True
    End If
    If PrintDuplicate Then
        Select Case ReceiptType
        Case ucsFscRetNonfiscal, ucsFscRetInvoiceNonfiscal
            If IsEmpty(Rows) Then
                LastError = ERR_MISSING_ROWS
                GoTo QH
            End If
        Case Else
            If Not oFP.CopyLastReceipt() Then
                LastError = oFP.GetLastError()
                GoTo QH
            End If
            '--- success
            PrintReceipt = True
        End Select
    End If
    ValueAt(OperatorData, ucsFscPixLastReceipt) = oFP.GetLastReceiptNumber()
QH:
    If oFP.IsConnected Then
        oFP.Disconnect
    End If
    CommandLog = oFP.GetCommandLog()
    Exit Function
EH:
    If LenB(LastError) = 0 Then
        LastError = Err.Description
        DebugLog FUNC_NAME, "LastError=" & LastError
        If Not oFP Is Nothing Then
            sLastErr = oFP.GetLastError(eLastNum)
            DebugLog FUNC_NAME, "sLastErr=" & sLastErr & ", eLastNum=" & eLastNum
            If eLastNum = ucsFerInvalidPassword Then
                ValueAt(OperatorData, ucsFscPixOperatorPass) = Empty
            End If
            ValueAt(OperatorData, ucsFscPixResumeTicket) = oFP.GetResumeTicket()
            Resume QH
        End If
    End If
End Function

Public Function PrintReport( _
            DeviceString As String, _
            ByVal ReportType As UcsFiscalPrintReportTypeEnum, _
            Optional ReportData As Variant, _
            Optional OperatorData As Variant, _
            Optional LastError As String, _
            Optional CommandLog As String) As Boolean
    Const FUNC_NAME     As String = "PrintReport"
    Dim oFP             As IDeviceProtocol
    Dim bResult         As Boolean
    
    On Error GoTo EH
    LastError = vbNullString
    Set oFP = pvInit(DeviceString)
    Select Case ReportType
    Case ucsFscRptDaily
        '-- ReportData = { IsClear, IsItems, IsDepartments }
        If C_Bool(At(ReportData, 1)) And C_Bool(At(ReportData, 2)) Then
            If C_Bool(At(ReportData, 0)) Then
                bResult = oFP.RunZReport(ucsFscRstDailyByItemsAndDepartment)
            Else
                bResult = oFP.RunXReport(ucsFscRstDailyByItemsAndDepartment)
            End If
        ElseIf C_Bool(At(ReportData, 1)) Then
            If C_Bool(At(ReportData, 0)) Then
                bResult = oFP.RunZReport(ucsFscRstDailyByItems)
            Else
                bResult = oFP.RunXReport(ucsFscRstDailyByItems)
            End If
        ElseIf C_Bool(At(ReportData, 2)) Then
            If C_Bool(At(ReportData, 0)) Then
                bResult = oFP.RunZReport(ucsFscRstDailyByDepartment)
            Else
                bResult = oFP.RunXReport(ucsFscRstDailyByDepartment)
            End If
        Else
            If C_Bool(At(ReportData, 0)) Then
                bResult = oFP.RunZReport(ucsFscRstDaily)
            Else
                bResult = oFP.RunXReport(ucsFscRstDaily)
            End If
        End If
    Case ucsFscRptNumber
        '-- ReportData = { FromNum, ToNum, IsDetailed }
    Case ucsFscRptDate
        '-- ReportData = { FromDate, ToDate, IsDetailed }
        If C_Date(At(ReportData, 0)) <> 0 And C_Date(At(ReportData, 1)) <> 0 Then
            bResult = oFP.RunPeriodReport(IIf(C_Bool(At(ReportData, 2)), ucsFscRstPeriodDetailed, ucsFscRstPeriodShort), C_Date(At(ReportData, 0)), C_Date(At(ReportData, 1)))
        End If
    Case ucsFscRptOperator
        '-- ReportData = Empty
    End Select
    If Not bResult Then
        LastError = oFP.GetLastError()
        GoTo QH
    End If
    ValueAt(OperatorData, ucsFscPixLastReceipt) = oFP.GetLastReceiptNumber()
    '--- success
    PrintReport = True
QH:
    If oFP.IsConnected Then
        oFP.Disconnect
    End If
    CommandLog = oFP.GetCommandLog()
    Exit Function
EH:
    If LenB(LastError) = 0 Then
        LastError = Err.Description
        DebugLog FUNC_NAME, "LastError=" & LastError
        If Not oFP Is Nothing Then
            DebugLog FUNC_NAME, "oFP.GetLastError=" & oFP.GetLastError()
            Resume QH
        End If
    End If
End Function

Public Function OpenDrawer( _
            DeviceString As String, _
            Optional LastError As String, _
            Optional CommandLog As String) As Boolean
    Const FUNC_NAME     As String = "OpenDrawer"
    Dim oFP             As IDeviceProtocol
    
    On Error GoTo EH
    LastError = vbNullString
    Set oFP = pvInit(DeviceString)
    If Not oFP.OpenDrawer() Then
        LastError = oFP.GetLastError()
        GoTo QH
    End If
    '--- success
    OpenDrawer = True
QH:
    If oFP.IsConnected Then
        oFP.Disconnect
    End If
    CommandLog = oFP.GetCommandLog()
    Exit Function
EH:
    If LenB(LastError) = 0 Then
        LastError = Err.Description
        DebugLog FUNC_NAME, "LastError=" & LastError
        If Not oFP Is Nothing Then
            DebugLog FUNC_NAME, "oFP.GetLastError=" & oFP.GetLastError()
            Resume QH
        End If
    End If
End Function

Public Function InitDateTime( _
            DeviceString As String, _
            sDateTime As String, _
            Optional LastError As String, _
            Optional CommandLog As String) As Boolean
    Const FUNC_NAME     As String = "InitDateTime"
    Dim oFP             As IDeviceProtocol
    Dim bResult         As Boolean
    
    On Error GoTo EH
    LastError = vbNullString
    Set oFP = pvInit(DeviceString)
    If C_Date(sDateTime) = 0 Then
        bResult = oFP.SetClock(Now)
    Else
        bResult = oFP.SetClock(C_Date(sDateTime))
    End If
    If Not bResult Then
        LastError = oFP.GetLastError()
        GoTo QH
    End If
    sDateTime = oFP.GetClock()
    '--- success
    InitDateTime = True
QH:
    If oFP.IsConnected Then
        oFP.Disconnect
    End If
    CommandLog = oFP.GetCommandLog()
    Exit Function
EH:
    If LenB(LastError) = 0 Then
        LastError = Err.Description
        DebugLog FUNC_NAME, "LastError=" & LastError
        If Not oFP Is Nothing Then
            DebugLog FUNC_NAME, "oFP.GetLastError=" & oFP.GetLastError()
            Resume QH
        End If
    End If
End Function

Public Function GetTotalsByPayments( _
            DeviceString As String, _
            vPayments As Variant, _
            Optional LastError As String, _
            Optional CommandLog As String) As Boolean
    Const FUNC_NAME     As String = "GetTotalsByPayments"
    Dim oFP             As IDeviceProtocol
    Dim lIdx            As Long
    Dim vTotals         As Variant

    On Error GoTo EH
    LastError = vbNullString
    Set oFP = pvInit(DeviceString)
    vTotals = oFP.GetTotalsByPayments()
    If Not IsArray(vTotals) Then
        LastError = ERR_NOT_IMPLEMENTED
        GoTo QH
    ElseIf UBound(vTotals) < 0 Then
        LastError = oFP.GetLastError()
        GoTo QH
    End If
    ReDim vPayments(0 To UBound(vTotals)) As Variant
    For lIdx = 0 To UBound(vTotals)
        vPayments(lIdx) = Array(vTotals(lIdx), oFP.GetPaymentName(lIdx + 1))
    Next
    '--- success
    GetTotalsByPayments = True
QH:
    If oFP.IsConnected Then
        oFP.Disconnect
    End If
    CommandLog = oFP.GetCommandLog()
    Exit Function
EH:
    If LenB(LastError) = 0 Then
        LastError = Err.Description
        DebugLog FUNC_NAME, "LastError=" & LastError
        If Not oFP Is Nothing Then
            DebugLog FUNC_NAME, "oFP.GetLastError=" & oFP.GetLastError()
            Resume QH
        End If
    End If
End Function

Public Function GetTotalsByVatGroups( _
            DeviceString As String, _
            vTotals As Variant, _
            Optional LastError As String, _
            Optional CommandLog As String) As Boolean
    Const FUNC_NAME     As String = "GetTotalsByVatGroups"
    Dim oFP             As IDeviceProtocol
    
    On Error GoTo EH
    LastError = vbNullString
    Set oFP = pvInit(DeviceString)
    vTotals = oFP.GetTotalsByVatGroups()
    If Not IsArray(vTotals) Then
        LastError = ERR_NOT_IMPLEMENTED
        GoTo QH
    ElseIf UBound(vTotals) < 0 Then
        LastError = oFP.GetLastError()
        GoTo QH
    End If
    '--- success
    GetTotalsByVatGroups = True
QH:
    If oFP.IsConnected Then
        oFP.Disconnect
    End If
    CommandLog = oFP.GetCommandLog()
    Exit Function
EH:
    If LenB(LastError) = 0 Then
        LastError = Err.Description
        DebugLog FUNC_NAME, "LastError=" & LastError
        If Not oFP Is Nothing Then
            DebugLog FUNC_NAME, "oFP.GetLastError=" & oFP.GetLastError()
            Resume QH
        End If
    End If
End Function

Public Function GetTotalsStats( _
            DeviceString As String, _
            vStats As Variant, _
            Optional LastError As String, _
            Optional CommandLog As String) As Boolean
    Const FUNC_NAME     As String = "GetTotalsStats"
    Dim oFP             As IDeviceProtocol

    On Error GoTo EH
    LastError = vbNullString
    Set oFP = pvInit(DeviceString)
    vStats = oFP.GetTotalsStats()
    If Not IsArray(vStats) Then
        LastError = ERR_NOT_IMPLEMENTED
        GoTo QH
    ElseIf UBound(vStats) < 0 Then
        LastError = oFP.GetLastError()
        GoTo QH
    End If
    '--- success
    GetTotalsStats = True
QH:
    If oFP.IsConnected Then
        oFP.Disconnect
    End If
    CommandLog = oFP.GetCommandLog()
    Exit Function
EH:
    If LenB(LastError) = 0 Then
        LastError = Err.Description
        DebugLog FUNC_NAME, "LastError=" & LastError
        If Not oFP Is Nothing Then
            DebugLog FUNC_NAME, "oFP.GetLastError=" & oFP.GetLastError()
            Resume QH
        End If
    End If
End Function

Public Function CashDebitCredit( _
            DeviceString As String, _
            dblValue As Double, _
            vTotals As Variant, _
            Optional OperatorData As Variant, _
            Optional LastError As String, _
            Optional CommandLog As String) As Boolean
    Const FUNC_NAME     As String = "CashDebitCredit"
    Dim oFP             As IDeviceProtocol
    
    On Error GoTo EH
    LastError = vbNullString
    Set oFP = pvInit(DeviceString)
    vTotals = oFP.CashDebitCredit(At(OperatorData, ucsFscPixOperatorCode), At(OperatorData, ucsFscPixOperatorPass), dblValue)
    If Not IsArray(vTotals) Then
        LastError = ERR_NOT_IMPLEMENTED
        GoTo QH
    ElseIf UBound(vTotals) < 0 Then
        LastError = oFP.GetLastError()
        GoTo QH
    End If
    ValueAt(OperatorData, ucsFscPixLastReceipt) = oFP.GetLastReceiptNumber()
    '--- success
    CashDebitCredit = True
QH:
    If oFP.IsConnected Then
        oFP.Disconnect
    End If
    CommandLog = oFP.GetCommandLog()
    Exit Function
EH:
    If LenB(LastError) = 0 Then
        LastError = Err.Description
        DebugLog FUNC_NAME, "LastError=" & LastError
        If Not oFP Is Nothing Then
            DebugLog FUNC_NAME, "oFP.GetLastError=" & oFP.GetLastError()
            Resume QH
        End If
    End If
End Function

Public Function GetStatus( _
            DeviceString As String, _
            sStatus As String, _
            Optional LastError As String, _
            Optional CommandLog As String) As Boolean
    Const FUNC_NAME     As String = "GetStatus"
    Dim oFP             As IDeviceProtocol
    
    On Error GoTo EH
    LastError = vbNullString
    Set oFP = pvInit(DeviceString)
    oFP.GetDeviceStatus sStatus
    '--- success
    GetStatus = True
QH:
    If oFP.IsConnected Then
        oFP.Disconnect
    End If
    CommandLog = oFP.GetCommandLog()
    Exit Function
EH:
    If LenB(LastError) = 0 Then
        LastError = Err.Description
        DebugLog FUNC_NAME, "LastError=" & LastError
        If Not oFP Is Nothing Then
            DebugLog FUNC_NAME, "oFP.GetLastError=" & oFP.GetLastError()
            Resume QH
        End If
    End If
End Function

'= private ===============================================================

Private Function pvInit(DeviceString As String) As IDeviceProtocol
    Dim vSplit          As Variant
    Dim sProtocol       As String
    
    vSplit = Split(DeviceString, ";")
    sProtocol = UCase$(At(vSplit, 0))
    '--- figure out model
    Select Case sProtocol
    Case STR_PROTOCOL_ELTRADE_ECR
        Set pvInit = New cEltradeProtocol
        pvInit.SetLocalizedText ucsFscLciInternalErrors, GetConfigValue(sProtocol, CFG_INTERNAL_ERRORS, STR_ELTRADE_ECR_INTERNAL)
        pvInit.SetLocalizedText ucsFscLciReceiptTexts, GetConfigValue(sProtocol, CFG_RECEIPT_TEXTS, STR_ELTRADE_ECR_TEXTS)
    Case STR_PROTOCOL_DATECS_FP, STR_PROTOCOL_DAISY_ECR
        Set pvInit = New cICLProtocol
        pvInit.SetLocalizedText ucsFscLciInternalErrors, GetConfigValue(sProtocol, CFG_INTERNAL_ERRORS, STR_DATECS_FP_INTERNAL)
        pvInit.SetLocalizedText ucsFscLciPrinterStatuses, GetConfigValue(sProtocol, CFG_PRINTER_STATUSES, STR_DATECS_FP_STATUSES)
        pvInit.SetLocalizedText ucsFscLciPrinterErrors, GetConfigValue(sProtocol, CFG_PRINTER_ERRORS, STR_DATECS_FP_ERRORS)
        pvInit.SetLocalizedText ucsFscLciDipSwitchesTexts, GetConfigValue(sProtocol, CFG_DIP_SWITCHES_TEXTS, STR_DATECS_FP_DIP_SWITCHES)
        pvInit.SetLocalizedText ucsFscLciMemoryTexts, GetConfigValue(sProtocol, CFG_MEMORY_TEXTS, STR_DATECS_FP_MEMORY)
        pvInit.SetLocalizedText ucsFscLciReceiptTexts, GetConfigValue(sProtocol, CFG_RECEIPT_TEXTS, STR_DATECS_FP_TEXTS)
    Case STR_PROTOCOL_ZEKA_FP
        Set pvInit = New cZekaProtocol
        pvInit.SetLocalizedText ucsFscLciInternalErrors, GetConfigValue(sProtocol, CFG_INTERNAL_ERRORS, STR_ZEKA_FP_INTERNAL)
        pvInit.SetLocalizedText ucsFscLciPrinterStatuses, GetConfigValue(sProtocol, CFG_PRINTER_STATUSES, STR_ZEKA_FP_STATUSES)
        pvInit.SetLocalizedText ucsFscLciPrinterErrors, GetConfigValue(sProtocol, CFG_PRINTER_ERRORS, STR_ZEKA_FP_ERRORS)
        pvInit.SetLocalizedText ucsFscLciReceiptTexts, GetConfigValue(sProtocol, CFG_RECEIPT_TEXTS, STR_ZEKA_FP_TEXTS)
    Case Else
        Set pvInit = New cEltradeProtocol
        pvInit.SetLocalizedText ucsFscLciInternalErrors, GetConfigValue(sProtocol, CFG_INTERNAL_ERRORS, STR_ELTRADE_ECR_INTERNAL)
        pvInit.SetLocalizedText ucsFscLciPrinterErrors, GetConfigValue(sProtocol, CFG_PRINTER_ERRORS, STR_DATECS_FP_ERRORS)
    End Select
    If Not pvInit.Init(At(vSplit, 1, "COM1"), C_Lng(At(vSplit, 2)), C_Lng(At(vSplit, 3))) Then
        Err.Raise vbObjectError, , Zn(pvInit.GetLastError(), Printf(ERR_INIT_PROTOCOL, At(vSplit, 0, STR_NONE)))
    End If
End Function

Private Function pvGetRowData(vRow As Variant) As UcsRow
    Const FUNC_NAME     As String = "pvGetRowData"
    
    On Error GoTo EH
    If UBound(vRow) = ucsFscPixItem Then
        pvGetRowData.Type = ucsRwtText
        pvGetRowData.Item = C_Str(vRow(ucsFscPixItem))
    ElseIf UBound(vRow) = ucsFscPixAmount Then
        pvGetRowData.Type = ucsRwtPayment
        pvGetRowData.Item = C_Str(vRow(ucsFscPixItem))
        pvGetRowData.WayOfPayment = C_Str(vRow(ucsFscPixWayOfPmt))
        pvGetRowData.Amount = C_Str(vRow(ucsFscPixAmount))
    Else
        pvGetRowData.Type = ucsRwtSell
        pvGetRowData.Item = C_Str(vRow(ucsFscPixItem))
        pvGetRowData.Price = C_Str(vRow(ucsFscPixPrice))
        pvGetRowData.Group = C_Str(vRow(ucsFscPixGroup))
        pvGetRowData.Quantity = C_Str(vRow(ucsFscPixQuantity))
        pvGetRowData.Discount = At(vRow, ucsFscPixDiscount)
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function pvGetVatGroup(sVatGroup As String) As Long
    pvGetVatGroup = C_Lng(sVatGroup)
    If pvGetVatGroup < 1 Or pvGetVatGroup > 8 Then
        pvGetVatGroup = 2
    End If
End Function

