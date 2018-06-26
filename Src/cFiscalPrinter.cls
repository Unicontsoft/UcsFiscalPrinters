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
Attribute VB_Description = "cFiscalPrinter class is used as a high-level device management component that can instantiate specific device drivers and can be used to call methods of the chatty IDeviceProtocol protocol in batches"
''
' cFiscalPrinter class is used as a high-level device management component that can instantiate specific device drivers
' and can be used to call methods of the chatty IDeviceProtocol protocol in batches.
'
' Each public function returns `True` on success and `False` on error. If present, output parameter `LastError` is
' populated with the error reported by the native protocol driver that was used to perform the operation and
' `CommandLog` is populated with debug info from the driver too.
'
' @param `DeviceString` Most functions accept `DeviceString` as first input parameter which is formatted as
'                       "Protocol[;Device[;TimeoutMs [;CashDesk]]]" where:
'   `Protocol`          - is one of `STR_PROTOCOL_Xxx` string constants in `mdGlobals`
'   `Device`            - is in "ComPort[,BaudRate[,Persistent[,DataBits[,Parity[,StopBits]]]]]" format, defaults to
'                         "COM1", e.g. "COM1,9600,P,8,N,1"
'   `TimeoutMs`         - is timeout in milliseconds, defaults to protocol's DEF_TIMEOUT constant (usually 3000)
'   `CashDesk`          - is the cash desk number if multiple cash desks are supported by the device, defaults to 1
'
' @param `OperatorData` Some functions accept `OperatorData` as an array of (OperatorCode, OperatorName, OperatorPass
'                       [, CashDeskNo[, LastReceipt[, PrevReceipt[, PrevCount[, ResumeTicket]]]]]) where:
'   `OperatorCode`      - is operator number to be authenticated (usually 1-8)
'   `OperatorName`      - is operator name to be printed on receipt
'   `OperatorPass`      - is operator password, can be `Empty` if default for device
'   `CashDeskNo`        - is cash desk number, defaults to 1
'   `LastReceipt`       - [output] gets the number of the last printed receipt
'   `PrevReceipt`       - [output] gets the number of the receipt printed before the current one
'   `PrevCount`         - not used
'   `ResumeTicket`      - [in/out] gets resume ticket on print error, can be supplied to resume receipt printing
'
' @param `InvoiceData`  Function `PrintReceipt` accepts `InvoiceData` as an array of (InvoiceNo, CgName, CgTaxNo,
'                       CgVatNo, CgAddress, CgPersonReceiver, OwnPersonSeller, LnkDoc, InvoiceNoPrefix) where:
'   `InvoiceNo`         - the number of the invoice/credit note to be printed
'   `CgName`            - name of the company that is receiving the document
'   `CgTaxNo`           - local identification number of the company (country specific)
'   `CgVatNo`           - VAT identification number of the company (EU specific)
'   `CgAddress`         - address of the company that is receiving the document
'   `CgPersonReceiver`  - name of the person that is receiving the document
'   `OwnPersonSeller`   - name of the person that is issueing the document (not used)
'   `LnkDoc`            - for credit notes: document number of the invoice that is being corrected
'   `InvoiceNoPrefix`   - for non-fiscal receipts: prefix to be printed before `InvoiceNo`
'
' @param `Rows`         Function `PrintReceipt` function accept `Rows` as an arrays of rows to be printed on the
'                       receipt, each element can be one of three types of rows for PLUs, payments and free text
'                       printing respectively (determined by array size):
'                       1. array of (ItemName, Price, Group, Quantity, Discount) for PLU row
'                       2. array of (WayOfPmtName, WayOfPmt, Amount) for payment row, cannot by followed by PLUs rows
'                       3. array of (Text) for free text row, where:
'   `ItemName`          - is the name of the PLU to be printed
'   `Price`             - is a double with PLU price, rounded to 0.00
'   `Group`             - is the VAT group the row will be accumulated to (1-8)
'   `Quantity`          - is a double with row quantity, rounded to 0.000
'   `Discount`          - is optional positive percent for discount and negative for surcharge (-100.00 to 100.00)
'   `PaymentName`       - is the name of the payment if custom type of payment is used (officially registered vouchers,
'                         etc.)
'   `WayOfPmt`          - is the type of payment for the amount. Standard types are 1-3 (cash, card, cheque), custom
'                         types are 5-8
'   `Amount`            - is a double with the amount that is paid using `WayOfPmt`, rounded to 0.00
'
' @remarks Protocol strings can be localized using `UcsFP01.conf` JSON file. Private `pvInit` function reads protocol
' specific strings under a section with protocol name (`STR_PROTOCOL_Xxx` constants) which include (by protocol):
' `ELTRADE ECR`      - reads `InternalErrors` and `ReceiptTexts`
' `DATECS FP/ECR`    - reads `InternalErrors`, `PrinterStatuses`, `PrinterErrors`, `DipSwitchesTexts`, `MemoryTexts`,
'                      `ReceiptTexts` and `ExtendedErrors`
' `DAISY FP/ECR`     - reads `InternalErrors`, `PrinterStatuses`, `PrinterErrors`, `DipSwitchesTexts`, `MemoryTexts`,
'                      `ReceiptTexts` and `ExtendedErrors`,
' `TREMOL ECR`       - reads `InternalErrors`, `PrinterStatuses`, `PrinterErrors` and `ReceiptTexts`
'
'=========================================================================
' $Header: /UcsFiscalPrinter/Src/cFiscalPrinter.cls 43    26.06.18 16:31 Wqw $
'
'   Unicontsoft Fiscal Printers Project
'   Copyright (c) 2008-2018 Unicontsoft
'
'   Fiscal devices management using IDeviceProtocol protocol handlers
'
' $Log: /UcsFiscalPrinter/Src/cFiscalPrinter.cls $
' 
' 43    26.06.18 16:31 Wqw
' REF: copy last receipt params
'
' 42    26.06.18 13:21 Wqw
' REF: debug log last error
'
' 41    13.06.18 12:20 Wqw
' REF: Init params
'
' 40    28.05.18 16:37 Wqw
' ADD: entries in UcsFiscalHeaderIndexesEnum
'
' 39    15.05.18 12:29 Wqw
' REF: esp/pos consts
'
' 38    14.05.18 12:38 Wqw
' REF: auto-detect esp/pos printers
'
' 37    25.04.18 10:40 Wqw
' REF: impl ESC/POS driver support
'
' 36    1.02.16 17:21 Wqw
' REF: impl extended errors for daisy fp/ecr
'
' 35    28.01.16 16:01 Wqw
' REF: daisy/tremol driver name
'
' 34    20.11.15 16:40 Wqw
' REF: add VBDOX documentation and typelib descriptions
'
' 33    30.01.15 15:07 Wqw
' REF: calls flush debug log on terminate
'
' 32    28.01.15 15:34 Wqw
' REF: error messages
'
' 31    26.11.14 19:19 Wqw
' REF: spelling
'
' 30    16.09.14 16:10 Wqw
' REF: error msgs spelling
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
    ucsFscPixCgTaxNo
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
    ucsFscPixTableNo
    ucsFscPixLastReceipt
    ucsFscPixPrevReceipt
    ucsFscPixPrevCount
    ucsFscPixResumeTicket
    ucsFscPixOwnData
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
    ucsFscHdrWopCustom4
    ucsFscHdrCharsPerLine
    ucsFscHdrDefOperPassword
    [_ucsFscHdrLast]
End Enum

'=========================================================================
' Constants and member variables
'=========================================================================

'--- strings
Private Const STR_TEXTS                 As String = "ДУБЛИКАТ|ОРИГИНАЛ|към ф-ра |Касиер"
Private Const STR_INTERNAL              As String = "Грешка при инициализация на протокол %1|Липсват редове за печат|Командата не се поддържа от протокола"
Private Const STR_ELTRADE_ECR_INTERNAL  As String = "Не е указано устройство|Грешка при отваряне: %1|Грешка при SetCommTimeouts: %1|Грешка при BuildCommDCB: %1|Грешка при SetCommState: %1|Грешка при WriteFile: %1|Време за достъп изтече в очакване на отговор|Грешка при ReadFile: %1|Вече има отворена връзка|Няма отворена връзка|Невалидно заглавие на отговор от устройството|Грешка при WaitCommEvent: %1|Няма започната бележка|Липсва хартия|Невалиден параметър 'Command'|Непозволена команда за печат на повече от един дубликат|Грешка при печат|Фатална грешка на фискалното устройство|Невъзможно отказване на предишна транзакция|Грешка на ред %1: %2"
Private Const STR_ELTRADE_ECR_TEXTS     As String = "СУМА|МЕЖДИННА СУМА|Продажби %1|ед. цена|количество|сума*%1|РЕСТО|всичко|ДДС*%1|нето ст-ст|Фактура #|Продавач:|   * * * * * * * *|Получател:|ИДЕНТ # |ЗДДС # |%1 артикули|МОЛ:|Получил:|НАДБАВКА|ОТСТЪПКА"
Private Const STR_DATECS_FP_INTERNAL    As String = "Не е указано устройство|Грешка при отваряне: %1|Грешка при SetCommTimeouts: %1|Грешка при BuildCommDCB: %1|Грешка при SetCommState: %1|Грешка при WriteFile: %1|Време за достъп изтече в очакване на отговор|Грешка при ReadFile: %1|Невалиден формат на съобщение или сума за проверка (NAK)|Вече има отворена връзка|Няма отворена връзка|Невалидна дължина на отговора: %1|Липсва символ за край на отговора: %1|Липсва символ за начало на отговора: %1|Грешка при WaitCommEvent: %1|Няма започната бележка| или невалидна парола на оператор|Грешка на ред %1: %2|Невъзможно отказване на предишна транзакция|Невалиден тип справка|Грешка %1|Памет %1"
Private Const STR_DATECS_FP_STATUSES    As String = "Синтактична грешка|Невалидна команда|Неустановени дата и час|3|Неизправност в механизма на печатащото устройство|5|6|7|Аритметично препълване|Непозволена команда|Зануляване на оперативна памет|3|Разрушено съдържание на оперативна памет|5|6|7|Няма хартия|1|2|Отворен фискален бон|4|Отворен служебен (нефискален) бон|6|7"
Private Const STR_DATECS_FP_ERRORS      As String = "Синтактична грешка|Невалидна команда|2|3|Неизправност в механизма на печатащото устройство|5|6|7|1|Непозволена команда|Зануляване на оперативна памет|3|Разрушено съдържание на оперативна памет|5|6|7|Няма хартия|1|2|3|4|5|6|7"
Private Const STR_DATECS_FP_DIP_SWITCHES As String = "Автоматично центриране на header и footer|Предварителен header|Sw1.3|Sw1.4|Нулиране на паметта|Прозрачен дисплей|Без данни на дисплея|7"
Private Const STR_DATECS_FP_MEMORY      As String = "Грешка при запис|1|Няма модул фискална памет|Малко свободно място във фискалната памет|Пълна фискална памет|5|6|7|Фискалната памет забранена за запис|Форматирана фискална памет|2|Принтерът е във фискален режим|Зададени данъчни ставки|Програмирани индивидуален номер и номер на фискалната памет|6|7"
Private Const STR_DATECS_FP_TEXTS       As String = "В БРОЙ|С ДЕБИТНА КАРТА|С ЧЕК|С КРЕДИТНА КАРТА|НАДБАВКА %1|ОТСТЪПКА %1|СУМА|МЕЖДИННА СУМА|ФАКТУРА No %1|Продажби %1|ЕДИНИЧНА ЦЕНА|КОЛИЧЕСТВО|СУМА|ВСИЧКО ГРУПА %1|ДДС %1=%2|НЕТО СТОЙНОСТ|ОБЩО|%1 АРТИКУЛА|1 АРТИКУЛ|РЕСТО|ПРОДАВАЧ: |ПОЛУЧАТЕЛ: |КУПУВАЧ: |ИДЕНТ. No: |ЗДДС No: |СИСТЕМЕН БОН"
Private Const STR_ZEKA_FP_INTERNAL      As String = "Не е указано устройство|Грешка при отваряне: %1|Грешка при SetCommTimeouts: %1|Грешка при BuildCommDCB: %1|Грешка при SetCommState: %1|Грешка при WriteFile: %1|Време за достъп изтече в очакване на отговор|Грешка при ReadFile: %1|Невалиден формат на съобщение или сума за проверка (NAK)|Вече има отворена връзка|Няма отворена връзка|Невалидна дължина на отговора: %1|Липсва символ за начало на отговора: %1|Грешка при WaitCommEvent: %1|Няма започната бележка|Грешка на ред %1: %2|Невъзможно отказване на предишна транзакция|Невалиден тип справка|Устройството е недостъпно|Устройството е заето| или невъзможно установяване парола на оператор| или невалидна парола на оператор"
Private Const STR_ZEKA_FP_STATUSES      As String = "Неизвестна грешка|Невалидна команда|Непозволена команда|Непозволена поради ненулев отчет|Синтактична грешка|Препълване на входните регистри|Нулев входен регистър|Липсва транзакция която да се войдира|Недостатъчна налична сума|Конфликт в данъчните групи"
Private Const STR_ZEKA_FP_ERRORS        As String = "Неизвестна грешка|Няма хартия|Препълване на дневните регистри|Несверен/грешен часовник|Отворен фискален бон|Сметка с остатък за плащане|Отворен нефискален бон|Сметка с приключено плащане|Фискална памет само за четене|Грешна парола или непозволена команда|Липсващ външен дисплей|24 часа без дневен отчет|Прегрят принтер|Спад на напрежение във фискален бон|Препълване в електронната контролна лента|Недостатъчни условия"
Private Const STR_ZEKA_FP_TEXTS         As String = "В брой|С карта|С чек|Надбавка%1|Отстъпка%1|Сума|Междинна сума|Фактура No %1|Продажби %1|Единична цена|Количество|Сума*%1|Всичко група %1|ДДС*%1=%2|Нето стойност|Обща сума|%1 Продажби|1 Продажба|Ресто|Продавач: |Получател: |Купувач: |Идент. No: |ЗДДС No: |ЕИК;ЕГН;ЛНЧ;Служебен No|КОРЕКЦИЯ"
Private Const STR_DAISY_FP_EXT_ERRORS   As String = "1|Въпросната операция ще доведе до препълване|3|Нямате право на повече продажби в този бон|4|Нямате право на повече плащания в този бон|5|Опит за извършване на нулева транзакция|6|Опит за извършване на продажба, след като е започнато плащане|7|Нямате право на такава операция|8|Забранена за продажби дан. група|11|Въвеждане на повече от една десетична точка|12|Въвеждане на повече от един символ '+' или '-'|13|Символът '+' или '-' не е в първа позиция|14|Недопустим символ, напр. баркод, който съдържа не само цифри|15|Повече от допустимия брой знаци след десетичната точка|16|Въведени са повече от разрешения брой символи|20|В дадената ситуация не се обслужва избраната от Вас команда от PC|21|Стойността е извън допустимите граници|22|Виж системен параметър 8|23|Опит за ""дълбок"" войд след отстъпка/надбавка в/у междинна сума|24|Опит за ""дълбок"" войд на несъществуваща транзакция|25|Опит за извършване на плащане, без да има продажби|" & _
                                                    "26|Опит за продажба на артикул с количество, надвишаващо запаса му|41|Некоректен баркод (грешна контролна сума)|42|Опит за продажба с нулев баркод|43|Опит за програмиране с тегловен баркод|44|Опит за продажба с непрограмиран баркод|45|Опит за програмиране на вече съществуващ баркод|66|Некоректна парола|71|!!! Некоректни данни във ФП !!!!|72|!!! Грешка при запис във ФП !!!!|76|Необходима е информация от сървъра на НАП|90|Не е нулиран периодичният отчет|91|Не е нулиран дневният финансов отчет|92|Не е нулиран отчетът по оператори|93|Не е нулиран отчетът по артикули|94|Не може да се препрограмира това поле|81|Дневният финансов отчет е препълнен|83|Отчетът по оператори е препълнен|84|Отчетът по артикули е препълнен|84|Периодичният отчет е препълнен|88|КЛЕН е препълнен|102|Няма комуникация между ФУ и ДТ|104|Некоректна комуникация между ФУ и ДТ|110|Подменена SIM карта|111|Грешка при комуникация между ДТ и сървъра на НАП|" & _
                                                    "113|Сървърът на НАП не приема подадените му данни|117|Неуспешен опит за свързване на ДТ с мрежата на мобилния оператор|118|Операцията е забранена|119|Грешна въведена стойност|120|Невъведена стойност"
Private Const STR_ESP_POS_INTERNAL      As String = "Няма започната бележка|Грешка на ред %1: %2"
Private Const STR_ESP_POS_TEXTS         As String = "В БРОЙ|С ДЕБИТНА КАРТА|С ЧЕК|С КРЕДИТНА КАРТА|НАДБАВКА %1|ОТСТЪПКА %1|СУМА|МЕЖДИННА СУМА|ФАКТУРА No %1|Продажби %1|ЕДИНИЧНА ЦЕНА|КОЛИЧЕСТВО|СУМА|ВСИЧКО ГРУПА %1|ДДС %1=%2|НЕТО СТОЙНОСТ|ОБЩО|%1 АРТИКУЛА|1 АРТИКУЛ|РЕСТО|ПРОДАВАЧ: |ПОЛУЧАТЕЛ: |КУПУВАЧ: |ИДЕНТ. No: |ЗДДС No: |СИСТЕМЕН БОН|ПОРЪЧКА КЪМ КУХНЯ|Дата:     %1|Оператор: %1|Маса:     %1|ЕИК №: %1"
Private Const STR_ESP_POS_CONNECTOR     As String = "Не е указано устройство|Грешка при отваряне: %1|Грешка при SetCommTimeouts: %1|Грешка при BuildCommDCB: %1|Грешка при SetCommState: %1|Грешка при WriteFile: %1|Време за достъп изтече в очакване на отговор|Грешка при ReadFile: %1|Грешка при WaitCommEvent: %1"
'--- config entries
Private Const CFG_INTERNAL_ERRORS       As String = "InternalErrors"
Private Const CFG_PRINTER_STATUSES      As String = "PrinterStatuses"
Private Const CFG_PRINTER_ERRORS        As String = "PrinterErrors"
Private Const CFG_DIP_SWITCHES_TEXTS    As String = "DipSwitchesTexts"
Private Const CFG_MEMORY_TEXTS          As String = "MemoryTexts"
Private Const CFG_RECEIPT_TEXTS         As String = "ReceiptTexts"
Private Const CFG_EXTENDED_ERRORS       As String = "ExtendedErrors"
Private Const CFG_CONNECTOR_ERRORS      As String = "ConnectorErrors"

Private m_uConfig                   As UcsConfigValues

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

Private Type UcsConfigValues
    LocalizedText(0 To [_ucsFscLciMax] - 1) As Variant
End Type

Private Enum UcsInternalErrors
    ucsErrInitProtocol
    ucsErrMissingRows
    ucsErrNotImplemented
End Enum

Private Enum UcsReceiptTextsEnum
    ucsTxtDuplicate
    ucsTxtOriginal
    ucsTxtHeadCreditNote
    ucsTxtCashier
End Enum

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

Private Property Get LocalizedText(ByVal eIdx As UcsFiscalLocalizedIndexesEnum) As String
    LocalizedText = Join(m_uConfig.LocalizedText(eIdx), "|")
End Property

Private Property Let LocalizedText(ByVal eIdx As UcsFiscalLocalizedIndexesEnum, sValue As String)
    m_uConfig.LocalizedText(eIdx) = Split(sValue, "|")
End Property

'=========================================================================
' Methods
'=========================================================================

''
' Used for enumeration of available system serial COM ports and auto-detection of connected fiscal devices
'
' @return An array with available COM ports including info for auto-detected devices. Each element of the returned
'         array is an array of (port, auto-detected baud rate, auto-detected FP protocol, auto-detected FP model).
'         Auto-detected elements can be `Empty` if no device is found on the specified COM port.
'
Public Function EnumPorts() As Variant
Attribute EnumPorts.VB_Description = "Used for enumeration of available system serial COM ports and auto-detection of connected fiscal devices"
    Const FUNC_NAME     As String = "EnumPorts"
    Dim vPorts          As Variant
    Dim vElem           As Variant
    Dim oFP             As IDeviceProtocol
    Dim lIdx            As Long
    
    On Error GoTo EH
    vPorts = EnumSerialPorts
    For Each vElem In Array(pvInit(STR_PROTOCOL_ZEKA_FP), pvInit(STR_PROTOCOL_ESP_POS), _
                            pvInit(STR_PROTOCOL_DATECS_FP), pvInit(STR_PROTOCOL_ELTRADE_ECR))
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

''
' Used to retrieve an array with info from the fiscal device including header/footer texts, operator name, current
' date/time, last receipt number, names of payment types, max characters on a row and operator's default password
'
' @param `DeviceString` [input]  see `DeviceString` description in class header
' @param `OperatorData` [input]  see `OperatorData` description in class header
' @param `RetVal`       [in/out] param with the array with requested strings which get populated on return. Look at
'                                `UcsFiscalHeaderIndexesEnum` enum for proper indexes meaning
' @param `LastError`    [output] see `LastError` description in class header
' @param `CommandLog`   [output] see `CommandLog` description in class header
'
Public Function GetHeaderFooter( _
            DeviceString As String, _
            Optional OperatorData As Variant, _
            Optional RetVal As Variant, _
            Optional LastError As String, _
            Optional CommandLog As String) As Boolean
Attribute GetHeaderFooter.VB_Description = "Used to retrieve an array with info from the fiscal device including header/footer texts, operator name, current date/time, last receipt number, names of payment types, max characters on a row and operator's default password"
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
    If LenB(At(vInput, ucsFscHdrWopCustom4, "1")) <> 0 Then
        RetVal(ucsFscHdrWopCustom3) = oFP.GetPaymentName(-4)
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

''
' Used to batch print a fiscal or non-fiscal receipt or invoice
'
' @param `DeviceString`   [input]  see `DeviceString` description in class header
' @param `ReceiptType`    [input]  designates if receipt is fiscal or non-fiscal and if the receipt is an invoice or
'                                  regular sales receipt
' @param `Rows`           [input]  see `Rows` description in class header
' @param `OperatorData`   [in/out] see `OperatorData` description in class header
' @param `InvoiceData`    [input]  see `InvoiceData` description in class header
' @param `PrintDuplicate` [input]  a flag to request receipt duplicate to be printed if the fiscal device supports
'                                  this operation
' @param `LastError`      [output] see `LastError` description in class header
' @param `CommandLog`     [output] see `CommandLog` description in class header
' @return `True` on success, `False` on error and populates `LastError` and `CommandLog`
'
Public Function PrintReceipt( _
            DeviceString As String, _
            ByVal ReceiptType As UcsFiscalReceiptTypeEnum, _
            Rows As Variant, _
            Optional OperatorData As Variant, _
            Optional InvoiceData As Variant, _
            Optional PrintDuplicate As Boolean, _
            Optional LastError As String, _
            Optional CommandLog As String) As Boolean
Attribute PrintReceipt.VB_Description = "Used to batch print a fiscal or non-fiscal receipt or invoice"
    Const FUNC_NAME     As String = "PrintReceipt"
    Dim oFP             As IDeviceProtocol
    Dim vElem           As Variant
    Dim lIdx            As Long
    Dim sPrevReceipt    As String
    Dim sResumeTicket   As String
    Dim sOwnData        As String
    Dim eLastNum        As UcsFiscalErrorsEnum
    Dim sLastErr        As String
    
    On Error GoTo EH
    LastError = vbNullString
    sPrevReceipt = At(OperatorData, ucsFscPixPrevReceipt)
    sResumeTicket = At(OperatorData, ucsFscPixResumeTicket)
    sOwnData = At(OperatorData, ucsFscPixOwnData)
    Set oFP = pvInit(DeviceString)
    If LenB(sResumeTicket) = 0 Then
        If Not oFP.CancelReceipt() Then
            LastError = oFP.GetLastError()
            DebugLog FUNC_NAME, "Not oFP.CancelReceipt(), LastError=" & LastError
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
            At(OperatorData, ucsFscPixOperatorName, pvText(ucsTxtCashier)), _
            At(OperatorData, ucsFscPixOperatorPass), _
            At(OperatorData, ucsFscPixTableNo), _
            IIf(ReceiptType = ucsFscRetInvoiceNonfiscal, At(InvoiceData, ucsFscPixInvoiceNoPrefix), vbNullString) & At(InvoiceData, ucsFscPixInvoiceNo), _
            At(InvoiceData, ucsFscPixCgTaxNo), _
            At(InvoiceData, ucsFscPixCgVatNo), _
            At(InvoiceData, ucsFscPixCgName), _
            vbNullString, _
            At(InvoiceData, ucsFscPixCgAddress), _
            At(InvoiceData, ucsFscPixCgPersonReceiver), _
            vbNullString, _
            sOwnData
        If LenB(At(InvoiceData, ucsFscPixLnkDoc)) <> 0 Then
            oFP.AddLine pvText(ucsTxtHeadCreditNote) & At(InvoiceData, ucsFscPixLnkDoc)
        End If
        If PrintDuplicate Then
            oFP.AddLine pvText(ucsTxtDuplicate)
        ElseIf ReceiptType = ucsFscRetInvoiceNonfiscal Then
            oFP.AddLine pvText(ucsTxtOriginal)
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
                LastError = pvInternal(ucsErrMissingRows)
                GoTo QH
            End If
        Case Else
            If Not oFP.CopyLastReceipt(At(InvoiceData, ucsFscPixInvoiceNo)) Then
                LastError = oFP.GetLastError()
                DebugLog FUNC_NAME, "Not oFP.CopyLastReceipt(), LastError=" & LastError
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

''
' Used to print a fiscal report as supported by the fiscal device
'
' @param `DeviceString` [input]  see `DeviceString` description in class header
' @param `ReportType`   [input]  designates the type of report to be printed. Only daily x-report or z-report and
'                                report by date range are supported
' @param `ReportData`   [input]  an array of (IsClear, IsItems, IsDepartments) for daily report and array of
'                                (FromDate, ToDate, IsDetailed) for date range report, where:
'   `IsClear`           - a boolean, when `True` clears fiscal memory (z-report)
'   `IsItems`           - a boolean, when `True` prints detailed report by items
'   `IsDepartments`     - a boolean, when `True` prints detailed report by department
'   `FromDate`-`ToDate` - specify the date range to print
'   `IsDetailed`        - a boolean, when `True` date range report is detailed by date
' @param `OperatorData` [input]  see `OperatorData` description in class header
' @param `LastError`    [output] see `LastError` description in class header
' @param `CommandLog`   [output] see `CommandLog` description in class header
' @return `True` on success, `False` on error and populates `LastError` and `CommandLog`
'
Public Function PrintReport( _
            DeviceString As String, _
            ByVal ReportType As UcsFiscalPrintReportTypeEnum, _
            Optional ReportData As Variant, _
            Optional OperatorData As Variant, _
            Optional LastError As String, _
            Optional CommandLog As String) As Boolean
Attribute PrintReport.VB_Description = "Used to print a fiscal report as supported by the fiscal device"
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
        DebugLog FUNC_NAME, "Not bResult, LastError=" & LastError
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

''
' Used to send an impulse to the drawer opener. The fiscal device has to be connected with a compatible cash
' drawer for this command to work
'
' @param `DeviceString` [input]  see `DeviceString` description in class header
' @param `LastError`    [output] see `LastError` description in class header
' @param `CommandLog`   [output] see `CommandLog` description in class header
' @return `True` on success, `False` on error and populates `LastError` and `CommandLog`
'
Public Function OpenDrawer( _
            DeviceString As String, _
            Optional LastError As String, _
            Optional CommandLog As String) As Boolean
Attribute OpenDrawer.VB_Description = "Used to send an impulse to the drawer opener. The fiscal device has to be connected with a compatible cash drawer for this command to work"
    Const FUNC_NAME     As String = "OpenDrawer"
    Dim oFP             As IDeviceProtocol
    
    On Error GoTo EH
    LastError = vbNullString
    Set oFP = pvInit(DeviceString)
    If Not oFP.OpenDrawer() Then
        LastError = oFP.GetLastError()
        DebugLog FUNC_NAME, "Not oFP.OpenDrawer(), LastError=" & LastError
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

''
' Used to adjust the date and time of the fiscal device clock
'
' @param `DeviceString` [input]  see `DeviceString` description in class header
' @param `sDateTime`    [input]  a date and time to be set or `Empty` for current date and time
' @param `LastError`    [output] see `LastError` description in class header
' @param `CommandLog`   [output] see `CommandLog` description in class header
' @return `True` on success, `False` on error and populates `LastError` and `CommandLog`
'
Public Function InitDateTime( _
            DeviceString As String, _
            sDateTime As String, _
            Optional LastError As String, _
            Optional CommandLog As String) As Boolean
Attribute InitDateTime.VB_Description = "Used to adjust the date and time of the fiscal device clock"
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
        DebugLog FUNC_NAME, "Not bResult, LastError=" & LastError
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

''
' Used to retrieve accumulated totals by payment types since last z-report
'
' @param `DeviceString` [input]  see `DeviceString` description in class header
' @param `vPayments`    [output] array indexed by 1-3 for standard payment types and 5-8 for custom. Each element is
'                                an array of (Amount, PaymentName) where `Amount` is a double total sum and
'                                `PaymentName` is a localised string for standard or custom payment name as returned
'                                by the fiscal device
' @param `LastError`    [output] see `LastError` description in class header
' @param `CommandLog`   [output] see `CommandLog` description in class header
' @return `True` on success, `False` on error and populates `LastError` and `CommandLog`
'
Public Function GetTotalsByPayments( _
            DeviceString As String, _
            vPayments As Variant, _
            Optional LastError As String, _
            Optional CommandLog As String) As Boolean
Attribute GetTotalsByPayments.VB_Description = "Used to retrieve accumulated totals by payment types since last z-report"
    Const FUNC_NAME     As String = "GetTotalsByPayments"
    Dim oFP             As IDeviceProtocol
    Dim lIdx            As Long
    Dim vTotals         As Variant

    On Error GoTo EH
    LastError = vbNullString
    Set oFP = pvInit(DeviceString)
    vTotals = oFP.GetTotalsByPayments()
    If Not IsArray(vTotals) Then
        LastError = pvInternal(ucsErrNotImplemented)
        GoTo QH
    ElseIf UBound(vTotals) < 0 Then
        LastError = oFP.GetLastError()
        DebugLog FUNC_NAME, "UBound(vTotals) < 0, LastError=" & LastError
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

''
' Used to retrieve accumulated totals by VAT groups (1-8) since last z-report
'
' @param `DeviceString` [input]  see `DeviceString` description in class header
' @param `vTotals`      [output] array of doubles (TotalByVat1, TotalByVat2, ..., TotalByVat8) where each total is
'                                indexed by VAT group
' @param `LastError`    [output] see `LastError` description in class header
' @param `CommandLog`   [output] see `CommandLog` description in class header
' @return `True` on success, `False` on error and populates `LastError` and `CommandLog`
'
Public Function GetTotalsByVatGroups( _
            DeviceString As String, _
            vTotals As Variant, _
            Optional LastError As String, _
            Optional CommandLog As String) As Boolean
Attribute GetTotalsByVatGroups.VB_Description = "Used to retrieve accumulated totals by VAT groups (1-8) since last z-report"
    Const FUNC_NAME     As String = "GetTotalsByVatGroups"
    Dim oFP             As IDeviceProtocol
    
    On Error GoTo EH
    LastError = vbNullString
    Set oFP = pvInit(DeviceString)
    vTotals = oFP.GetTotalsByVatGroups()
    If Not IsArray(vTotals) Then
        LastError = pvInternal(ucsErrNotImplemented)
        GoTo QH
    ElseIf UBound(vTotals) < 0 Then
        LastError = oFP.GetLastError()
        DebugLog FUNC_NAME, "UBound(vTotals) < 0, LastError=" & LastError
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

''
' Used to retrieve number of receipts printed since last z-report and the date/time of the last receipt printed
'
' @param `DeviceString` [input]  see `DeviceString` description in class header
' @param `vStats`       [output] array of (ReceiptsCount, LastReceiptTime) where `LastReceiptTime` can be `Empty`
'                                if no receipt is printed after last z-report
' @param `LastError`    [output] see `LastError` description in class header
' @param `CommandLog`   [output] see `CommandLog` description in class header
' @return `True` on success, `False` on error and populates `LastError` and `CommandLog`
'
Public Function GetTotalsStats( _
            DeviceString As String, _
            vStats As Variant, _
            Optional LastError As String, _
            Optional CommandLog As String) As Boolean
Attribute GetTotalsStats.VB_Description = "Used to retrieve number of receipts printed since last z-report and the date/time of the last receipt printed"
    Const FUNC_NAME     As String = "GetTotalsStats"
    Dim oFP             As IDeviceProtocol

    On Error GoTo EH
    LastError = vbNullString
    Set oFP = pvInit(DeviceString)
    vStats = oFP.GetTotalsStats()
    If Not IsArray(vStats) Then
        LastError = pvInternal(ucsErrNotImplemented)
        GoTo QH
    ElseIf UBound(vStats) < 0 Then
        LastError = oFP.GetLastError()
        DebugLog FUNC_NAME, "UBound(vStats) < 0, LastError=" & LastError
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

''
' Used to debit/credit the fiscal device with cash total that is collected or loaded outside of regular client sales,
' e.g. debitting with initial daily change. Can be used to retrieve debit/credit totals too
'
' @param `DeviceString` [input]  see `DeviceString` description in class header
' @param `dblValue`     [input]  positive amount to be debitted and negative for credit. When 0.00 is passed no
'                                operation is performed, only `vTotals` is populated
' @param `vTotals`      [output] total debit/credit operations since last z-report as an array of doubles with
'                                (TotalAvailable, TotalDebit, TotalCredit) where:
'   `TotalAvailable`    - total available cash, including regular sales
'   `TotalDebit`        - total inbound debit, outside regular sales
'   `TotalCredit`       - total outbound credit, outside regular sales
' @param `LastError`    [output] see `LastError` description in class header
' @param `CommandLog`   [output] see `CommandLog` description in class header
' @return `True` on success, `False` on error and populates `LastError` and `CommandLog`
'
Public Function CashDebitCredit( _
            DeviceString As String, _
            dblValue As Double, _
            vTotals As Variant, _
            Optional OperatorData As Variant, _
            Optional LastError As String, _
            Optional CommandLog As String) As Boolean
Attribute CashDebitCredit.VB_Description = "Used to debit/credit the fiscal device with cash total that is collected or loaded outside of regular client sales, e.g. debitting with initial daily change. Can be used to retrieve debit/credit totals too"
    Const FUNC_NAME     As String = "CashDebitCredit"
    Dim oFP             As IDeviceProtocol
    
    On Error GoTo EH
    LastError = vbNullString
    Set oFP = pvInit(DeviceString)
    vTotals = oFP.CashDebitCredit(At(OperatorData, ucsFscPixOperatorCode), At(OperatorData, ucsFscPixOperatorPass), dblValue)
    If Not IsArray(vTotals) Then
        LastError = pvInternal(ucsErrNotImplemented)
        GoTo QH
    ElseIf UBound(vTotals) < 0 Then
        LastError = oFP.GetLastError()
        DebugLog FUNC_NAME, "UBound(vTotals) < 0, LastError=" & LastError
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

''
' Used to retrieve additional device status without performing any other operation, e.g. can be used to poll the
' device if the operator is ready with loading a new paper roll after receiving out-of-paper error on receipt printing
'
' @param `DeviceString` [input]  see `DeviceString` description in class header
' @param `sStatus`      [output] current device status. Empty string when the device is ready to print
' @param `LastError`    [output] see `LastError` description in class header
' @param `CommandLog`   [output] see `CommandLog` description in class header
' @return `True` on success, `False` on error and populates `LastError` and `CommandLog`
'
Public Function GetStatus( _
            DeviceString As String, _
            sStatus As String, _
            Optional LastError As String, _
            Optional CommandLog As String) As Boolean
Attribute GetStatus.VB_Description = "Used to retrieve additional device status without performing any other operation, e.g. can be used to poll the device if the operator is ready with loading a new paper roll after receiving out-of-paper error on receipt printing"
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

''
' Used to set localized strings used for non-fiscal receipt texts and component internal errors
'
Public Sub SetLocalizedText(ByVal Index As UcsFiscalLocalizedIndexesEnum, Text As String)
Attribute SetLocalizedText.VB_Description = "Used to set localized strings used for non-fiscal receipt texts and component internal errors"
    LocalizedText(Index) = Text
End Sub

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
        pvInit.SetLocalizedText ucsFscLciExtendedErrors, GetConfigValue(sProtocol, CFG_EXTENDED_ERRORS, STR_DAISY_FP_EXT_ERRORS)
    Case STR_PROTOCOL_ZEKA_FP
        Set pvInit = New cZekaProtocol
        pvInit.SetLocalizedText ucsFscLciInternalErrors, GetConfigValue(sProtocol, CFG_INTERNAL_ERRORS, STR_ZEKA_FP_INTERNAL)
        pvInit.SetLocalizedText ucsFscLciPrinterStatuses, GetConfigValue(sProtocol, CFG_PRINTER_STATUSES, STR_ZEKA_FP_STATUSES)
        pvInit.SetLocalizedText ucsFscLciPrinterErrors, GetConfigValue(sProtocol, CFG_PRINTER_ERRORS, STR_ZEKA_FP_ERRORS)
        pvInit.SetLocalizedText ucsFscLciReceiptTexts, GetConfigValue(sProtocol, CFG_RECEIPT_TEXTS, STR_ZEKA_FP_TEXTS)
    Case STR_PROTOCOL_ESP_POS
        Set pvInit = New cEscPosProtocol
        pvInit.SetLocalizedText ucsFscLciInternalErrors, GetConfigValue(sProtocol, CFG_INTERNAL_ERRORS, STR_ESP_POS_INTERNAL)
        pvInit.SetLocalizedText ucsFscLciReceiptTexts, GetConfigValue(sProtocol, CFG_RECEIPT_TEXTS, STR_ESP_POS_TEXTS)
        pvInit.SetLocalizedText ucsFscLciConnectorErrors, GetConfigValue(sProtocol, CFG_CONNECTOR_ERRORS, STR_ESP_POS_CONNECTOR)
    Case Else
        Set pvInit = New cEltradeProtocol
        pvInit.SetLocalizedText ucsFscLciInternalErrors, GetConfigValue(sProtocol, CFG_INTERNAL_ERRORS, STR_ELTRADE_ECR_INTERNAL)
        pvInit.SetLocalizedText ucsFscLciPrinterErrors, GetConfigValue(sProtocol, CFG_PRINTER_ERRORS, STR_DATECS_FP_ERRORS)
    End Select
    If LenB(At(vSplit, 1)) <> 0 Then
        If Not pvInit.Init(At(vSplit, 1, "COM1"), C_Lng(At(vSplit, 2)), C_Lng(At(vSplit, 3)), C_Lng(At(vSplit, 4))) Then
            Err.Raise vbObjectError, , Zn(pvInit.GetLastError(), Printf(pvInternal(ucsErrInitProtocol), At(vSplit, 0, STR_NONE)))
        End If
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
        pvText = At(Split(STR_TEXTS, "|"), lIdx)
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Initialize()
    LocalizedText(ucsFscLciInternalErrors) = GetConfigValue(MODULE_NAME, CFG_INTERNAL_ERRORS, STR_INTERNAL)
    LocalizedText(ucsFscLciReceiptTexts) = GetConfigValue(MODULE_NAME, CFG_RECEIPT_TEXTS, STR_TEXTS)
End Sub

Private Sub Class_Terminate()
    FlushDebugLog
End Sub
