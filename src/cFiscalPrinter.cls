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
' Fiscal device operations using IDeviceProtocol handlers
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "cFiscalPrinter"

'=========================================================================
' Public events
'=========================================================================

Event BeforeCommitReceipt(oCtx As Object, Cancel As Boolean)

'=========================================================================
' Public enums
'=========================================================================

Public Enum UcsFiscalPrintReportTypeEnum
    ucsFscRptDaily = 1
    ucsFscRptMonthlyByReceiptNo
    ucsFscRptMonthlyByDate
    ucsFscRptDailyByOperators
End Enum

'=========================================================================
' Constants and member variables
'=========================================================================

'--- strings
Private Const STR_TEXTS                 As String = "Касиер"
Private Const STR_INTERNAL              As String = "Грешка при инициализация на протокол %1|Командата не се поддържа от протокола|Невалидна JSON заявка|Невалиден тип справка: %1"
Private Const STR_DATECS_INTERNAL       As String = "Време за достъп изтече в очакване на отговор|Невалидна дължина на отговора: %1|Невалидна сума за проверка: %1|Липсва символ за край на отговора: %1|Липсва символ за начало на отговора: %1|Няма започната бележка| или невалидна парола на оператор|Грешка на ред %1: %2|Невъзможно отказване на предишна транзакция|Невалиден тип справка|Грешка %1|Памет %1|Неподдържан вид плащане %1|Невъзможно отпечатване|Грешен номер на фактура %1 върнат от диапазон на ФУ|Неподдържан вид отстъпка %1|Прекъсване на фискализация на бон"
Private Const STR_DATECS_STATUSES       As String = "Получените данни имат синктактична грешка|Кодът на получената команда е невалиден|Не е сверен часовника|3|Механизмът на печатащото устройство има неизправност|5|Отворен е капакът на принтера|7|При изпълнение на командата се е получило препълване на някои полета от сумите|Изпълнението на командата не е позволено в текущия фискален режим|2|3|4|5|6|7|Свършила е хартията|Останала е малко хартия|Край на КЛЕН|Отворен е фискален бон|Близък край на КЛЕН|Отворен е служебен бон"
Private Const STR_DATECS_ERRORS         As String = "Получените данни имат синктактична грешка|Кодът на получената команда е невалиден|2|3|Механизмът на печатащото устройство има неизправност|5|6|7|0|Изпълнението на командата не е позволено в текущия фискален режим|2|3|4|5|6|7|Свършила е хартията"
Private Const STR_DATECS_MEMORY         As String = "Грешка при запис във фискалната памет|Зададен е ЕИК|Зададени са индивидуален номер на ФУ и номер на ФП|Има място за по-малко от 50 записа във ФП|ФП е пълна|5|ФП не е намерена или е повредена|7|0|ФП е форматирана|2|ФУ е във фискален режим|Зададени са поне веднъж данъчните ставки"
Private Const STR_DATECS_TEXTS          As String = "В БРОЙ|НАДБАВКА %1|ОТСТЪПКА %1|МЕЖДИННА СУМА|Продажби %1"
Private Const STR_ISL_INTERNAL          As String = "Време за достъп изтече в очакване на отговор|Невалидна дължина на отговора: %1|Невалидна сума за проверка: %1|Липсва символ за край на отговора: %1|Липсва символ за начало на отговора: %1|Няма започната бележка| или невалидна парола на оператор|Грешка на ред %1: %2|Невъзможно отказване на предишна транзакция|Невалиден тип справка|Грешка %1|Памет %1|Неподдържан вид плащане %1|Баркод от тип  %1 не се поддържа|Невъзможно отпечатване|Грешен номер на фактура %1 върнат от диапазон на ФУ|Неподдържан вид отстъпка %1|Прекъсване на фискализация на бон"
Private Const STR_ISL_DIP_SWITCHES      As String = "Автоматично центриране на header и footer|Предварителен header|Sw1.3|Sw1.4|Нулиране на паметта|Прозрачен дисплей|Без данни на дисплея|7"
Private Const STR_ISL_TEXTS             As String = "В БРОЙ|С КАРТА|С ЧЕК|С КАРТА|НАДБАВКА %1|ОТСТЪПКА %1|МЕЖДИННА СУМА|Продажби %1|ЕИК"
Private Const STR_ISL_STATUSES_A        As String = "Получените данни имат синтактична грешка|Кодът на получената команда е невалиден|Часовникът не е установен|Не е свързан клиентски дисплей|Механизмът на печатащото устройство има неизправност|5|Отворен е капакът на принтера|7|При изпълнение на командата се е получило препълване на някои полета от сумите|Изпълнението на командата не е позволено в текущия фискален режим|Извършено е зануляване на оперативната памет|Слаба батерия (Часовникът за реално време е в състояние RESET)|Отворен сторно бон|Отворен е служебен бон за печат на завъртян на 90 градуса текст|Вграденият данъчен терминал не отговаря|7|Свършила е хартията|Останала е малко хартия|Край на КЛЕН (по-малко от 1 MB от КЛЕН свободни)|Отворен е фискален бон|Близък край на КЛЕН (по-малко от 10 MB от КЛЕН свободни)|Отворен е служебен бон|Много близък край на КЛЕН (допускат се само определени бонове)"
Private Const STR_ISL_ERRORS_A          As String = "Получените данни имат синтактична грешка|Кодът на получената команда е невалиден|2|3|Механизмът на печатащото устройство има неизправност|5|6|7|0|Изпълнението на командата не е позволено в текущия фискален режим|Извършено е зануляване на оперативната памет|Слаба батерия (Часовникът за реално време е в състояние RESET)|4|5|6|7|Свършила е хартията"
Private Const STR_ISL_MEMORY_A          As String = "Има грешка при запис във фискалната памет|Зададен е ЕИК по БУЛСТАТ|Зададени са индивидуален номер на принтера и номер на фискалната памет|Има място за по-малко от 50 записа във ФП|Фискалната памет е пълна|5|Печатащата глава е прегряла|7|Фискалната памет е установена в режим READONLY (заключена)|Фискалната памет е форматирана|Последният запис във фискалната памет не е успешен|Принтерът е във фискален режим|Зададени са поне веднъж данъчните ставки|Грешка при четене от фискалната памет"
Private Const STR_ISL_STATUSES_B        As String = "Получените данни имат синктактична грешка|Кодът на получената команда е невалиден|Не е сверен часовника|Не е свързан клиентски дисплей|4|5|6|7|При изпълнение на командата се е получило препълване на някои полета от сумите|Изпълнението на командата не е позволено в текущия фискален режим|2|3|4|Има неизпратени документи за повече от настроеното време за предупреждение|Вграденият данъчен терминал не отговаря|7|Свършила е хартията|1|Край на КЛЕН (по-малко от 1 MB от КЛЕН свободни)|Отворен е фискален бон|Близък край на КЛЕН (по-малко от 10 MB от КЛЕН свободни)|Отворен е служебен бон"
Private Const STR_ISL_ERRORS_B          As String = "Получените данни имат синктактична грешка|Кодът на получената команда е невалиден|2|3|4|5|6|7|0|Изпълнението на командата не е позволено в текущия фискален режим|2|3|4|5|6|7|Свършила е хартията"
Private Const STR_ISL_MEMORY_B          As String = "Грешка при запис във ФП|Зададен е ЕИК|Зададени са индивидуален номер на ФУ и номер на ФП|Има място за по-малко от 50 записа във ФП|ФП е пълна|5|6|7|0|ФП е форматирана|2|ФУ е във фискален режим|Зададени са поне веднъж данъчните ставки"
Private Const STR_ISL_STATUSES_DAISY    As String = "Синтактична грешка|Невалидна команда|Не са въведени дата и час|Няма външен дисплей|Грешка на печатащото устройство|5|6|7|Препълване на някои полета в сумите|Неразрешена команда в текущия режим (не е Връзка с PC)|Нулиран RAM|3|4|Грешка в cutter|Грешна парола|7|Няма хартия|Малко хартия|Няма хартия (контролна лента)|Отворен е фискален бон|Малко хартия(контролна лента)|Отворен е нефискален бон|Разрешен печат на документ"
Private Const STR_ISL_ERRORS_DAISY      As String = "Синтактична грешка|Невалидна команда|2|3|Грешка на печатащото устройство|5|6|7|0|Неразрешена команда в текущия режим (не е Връзка с PC)|Нулиран RAM|3|4|5|6|7|Няма хартия"
Private Const STR_ISL_MEMORY_DAISY      As String = "Грешка при запис във ФП|Проблем в данъчния терминал|Грешен запис във ФП|Място за по–малко от 50 записа във ФП|Пълна ФП|5|6|7|Препълнена ФП|1|2|ФУ е въведено в експлоатация|Зададени данъчни ставки|Програмирани са инд. номер на ФУ и номер на ФП|Готова ФП"
Private Const STR_ISL_EXTENDED_DAISY    As String = "1|Въпросната операция ще доведе до препълване|3|Нямате право на повече продажби в този бон|4|Нямате право на повече плащания в този бон|5|Опит за извършване на нулева транзакция|6|Опит за извършване на продажба, след като е започнато плащане|7|Нямате право на такава операция|8|Забранена за продажби дан. група|9|Грешен диапазон за номера на фактури|11|Въвеждане на повече от една десетична точка|10|Некоректна дата|11|Въвеждане на повече от една десетична точка|12|Въвеждане на повече от един символ '+' или '-'|13|Символът '+' или '-' не е в първа позиция|14|Недопустим символ, напр. баркод, който съдържа не само цифри|15|Повече от допустимия брой знаци след десетичната точка|16|Въведени са повече от разрешения брой символи|19|Има отворен бон |20|В дадената ситуация не се обслужва избраната от Вас команда от PC|21|Стойността е извън допустимите граници|22|Виж системен параметър 8|" & _
                                                    "23|Опит за ""дълбок"" войд след отстъпка/надбавка в/у междинна сума|24|Опит за ""дълбок"" войд на несъществуваща транзакция|25|Опит за извършване на плащане, без да има продажби|26|Опит за продажба на артикул с количество, надвишаващо запаса му|27|Некоректна комуникация с ел. везна|28|Опит за продажба без наличност|30|ФП е пълна|41|Некоректен баркод (грешна контролна сума)|42|Опит за продажба с нулев баркод|43|Опит за програмиране с тегловен баркод|44|Опит за продажба с непрограмиран баркод|45|Опит за програмиране на вече съществуващ баркод|61|Опит за работа с КЛЕН, неотговарящ на изискванията|66|Некоректна парола|10|!!! Не е открита ФП !!!!|71|!!! Некоректни данни във ФП !!!!|72|!!! Грешка при запис във ФП !!!!|76|Необходима е информация от сървъра на НАП|81|Дневният финансов отчет с нулиране и запис във фискалната памет е препълнен|82|Не е пуснат дневен отчет|" & _
                                                    "83|Отчетът по оператори е препълнен |84|Отчетът по артикули е препълнен|85|Периодичен отчет е препълнен |88|КЛЕН е препълнен|90|Не е нулиран периодичният отчет|91|Не е нулиран дневният финансов отчет|92|Не е нулиран отчетът по оператори|93|Не е нулиран отчетът по артикули|94|Не може да се препрограмира това поле|97|Не може да се препрограмира това поле|98|Не може да изпрати данни към, сървъра на НАП|102|Няма комуникация между ФУ и ДТ|104|Некоректна комуникация между ФУ и ДТ|108|Опит за въвеждане на грешна парола|110|Подменена SIM карта|111|Грешка при комуникация между ДТ и сървъра на НАП|113|Сървърът на НАП не приема подадените му данни|117|Неуспешен опит за свързване на ДТ с мрежата на мобилния оператор|118|Операцията е забранена|119|Грешна въведена стойност|120|Невъведена стойност|123|Повреден клен|124|ФУ не разпознава поставения КЛЕН (КЛЕН-ът е активиран на друго ФУ)|" & _
                                                    "125|Ако ФУ е регистрирано, то не може да работи без КЛЕН"
Private Const STR_ISL_STATUSES_INCO     As String = "В получените данни има синтактична грешка|Кодът на получената команда е невалиден|Часовникът не е установен|3|4|5|Отворен е капакът на принтера|7|0|Изпълнението на командата не е позволено|Оперативната памет е нулирана|Не е зададен диапазон на броене на брояча на фактурите|4|Поредно въвеждане на 3 грешни пароли|6|7|Край на хартията|1|2|Oтворен фискален бон|4|Отворен служебен бон"
Private Const STR_ISL_ERRORS_INCO       As String = "В получените данни има синтактична грешка|Кодът на получената команда е невалиден|2|3|4|5|Отворен е капакът на принтера|7|0|Изпълнението на командата не е позволено|Оперативната памет е нулирана|Не е зададен диапазон на броене на брояча на фактурите|4|5|6|7|Край на хартията"
Private Const STR_ISL_MEMORY_INCO       As String = "Грешка при запис във фискалната памет|1|Грешен запис във ФП|Има място за по-малко от 50 записа във ФП|Фискалната памет е пълна|Грешка във ФП|6|7|Фискалната памет е установена в режим READONLY|1|2|Фискалното устройство е фискализирано|Записани са данъчни ставки във ФП|Програмирани са индивидуалните номера на ФП и на ФУ"
Private Const STR_ISL_EXTENDED_INCO     As String = "1|Препълване от умножение|2|Препълване на дневен отчет|3|Неправилна операция|4|Отстъпката надхвърля стойността на продажбата или междинната сума|5|Нулева сума|6|Препълване на фискален бон|7|Препълване на количество|8|Препълване на цена или сума|9|Нулева стойност|10|Свободни цени не са разрешени за този департамент или артикул (стока)|11|Дробни количества не са разрешени|12|Достигнат е максималният брой операции във фискален бон. Фискалният бон трябва да се приключи|13|Процентната стойност е извън обхвата 0.00 – 99.99 Описание на програмния интерфейс 6|14|Операцията е забранена, защото би довела до отрицателен резултат|15|Грешка при достъп до артикул|16|Наличното количество е по-малко от въведеното|18|Няма регистрации във фискалния бон|19|Препълване на периодичен отчет|21|Наличната сума е недостатъчна|22|Препълване на отчет по оператори|24|Препълване на фискалната памет|25|Грешни граници за отчет на ФП|" & _
                                                    "26|Препълване на отчет по стоки/групи стоки|28|Не може да се отпечата дубликат|29|Фактури се отпечатват след фискализация на ФУ|32|Този бар-код е програмиран за друга стока|33|Необходим е отчет на тази стока/група стоки|34|Препълване на отчет по стокови групи|36|Недостатъчна наличност|37|Невалидна (деактивирана) данъчна група|42|Съответното плащане е забранено или е забранено връщането на ресто|43|Изразходвани са номерата на фактури от диапазона на броене на брояча на фактурите|73|Грешка - литиева батерия|75|Сумата/цената надхвърля програмирания лимит (параметър 2- поле 9; параметър 4 – поле 6 или 7 за ЕКАФП)|76|Грешка на часовника за реално време|77|Грешка в комуникацията с носителя на КЛЕН|78|Запълване на носителя на КЛЕН|79|Не е намерен валиден носител на КЛЕН"
Private Const STR_ISL_STATUSES_ELTRADE  As String = "Получените данни имат синтактична грешка|Кодът на получената команда е невалиден|Часовникът не е установен|Не е свързан клиентски дисплей|Механизмът на печатащото устройство има неизправност|5|Отворен е капакът на принтера|7|При изпълнение на командата се е получило препълване на някои полета от сумите|Изпълнението на командата не е позволено в текущия фискален режим|Извършено е зануляване на оперативната памет|Слаба батерия (Часовникът за реално време е в състояние RESET)|Установено е разрушаване на съдържанието на оперативната памет (RAM) след включване|Отворен е служебен бон за печат на завъртян на 90 градуса текст|Вграденият данъчен терминал не отговаря|7|Свършила е хартията|Останала е малко хартия|Край на КЛЕН (по-малко от 1 MB от КЛЕН свободни)|Отворен е фискален бон|Близък край на КЛЕН (по-малко от 10 MB от КЛЕН свободни)|Отворен е служебен бон|"
Private Const STR_ISL_ERRORS_ELTRADE    As String = "Получените данни имат синтактична грешка|Кодът на получената команда е невалиден|2|3|Механизмът на печатащото устройство има неизправност|5|6|7|0|Изпълнението на командата не е позволено в текущия фискален режим|Извършено е зануляване на оперативната памет|Слаба батерия (Часовникът за реално време е в състояние RESET)|Установено е разрушаване на съдържанието на оперативната памет (RAM) след включване|5|6|7|Свършила е хартията"
Private Const STR_ISL_MEMORY_ELTRADE    As String = "Има грешка при запис във фискалната памет|Зададен е ЕИК|Зададени са индивидуален номер на принтера и номер на фискалната памет|Има място за по-малко от 50 записа във ФП|Фискалната памет е пълна|Грешка във ФП|6|7|Фискалната памет е установена в режим READONLY (заключена)|Фискалната памет е форматирана|Последният запис във фискалната памет не е успешен|Принтерът е във фискален режим|Зададени са поне веднъж данъчните ставки|Грешка при четене от фискалната памет"
Private Const STR_TREMOL_INTERNAL       As String = "Време за достъп изтече в очакване на отговор|Невалиден формат на съобщение или сума за проверка (NAK)|Невалидна дължина на отговора: %1|Невалидна сума за проверка: %1|Липсва символ за начало на отговора: %1|Няма започната бележка|Грешка на ред %1: %2|Невалиден тип справка|Устройството е недостъпно|Устройството е заето| или невъзможно установяване парола на оператор|Неподдържан вид плащане %1|Баркод от тип %1 не се поддържа|Невъзможно отпечатване|Грешен номер на фактура %1 върнат от диапазон на ФУ|Неподдържан вид отстъпка %1|Прекъсване на фискализация на бон"
Private Const STR_TREMOL_STATUSES       As String = "Неизвестна грешка|Невалидна команда|Непозволена команда|Непозволена поради ненулев отчет|Синтактична грешка|Препълване на входните регистри|Нулев входен регистър|Липсва транзакция която да се войдира|Недостатъчна налична сума|Конфликт в данъчните групи"
Private Const STR_TREMOL_ERRORS         As String = "Неизвестна грешка|Няма хартия|Препълване на дневните регистри|Несверен/грешен часовник|Отворен фискален бон|Сметка с остатък за плащане|Отворен нефискален бон|Сметка с приключено плащане|Фискална памет само за четене|Грешна парола или непозволена команда|Липсващ външен дисплей|24 часа без дневен отчет|Прегрят принтер|Спад на напрежение във фискален бон|Препълване в електронната контролна лента|Недостатъчни условия"
Private Const STR_TREMOL_TEXTS          As String = "В брой|С карта|По банка|Надбавка%1|Отстъпка%1|Междинна сума|Продажби %1|ЕИК;ЕГН;ЛНЧ;Служебен No"
Private Const STR_ESP_POS_INTERNAL      As String = "Няма започната бележка|Грешка на ред %1: %2|Неподдържан вид отстъпка %1"
Private Const STR_ESP_POS_TEXTS         As String = "В БРОЙ|С ЧЕК|ТАЛОНИ|В.ТАЛОНИ|АМБАЛАЖ|В.ОБСЛУЖВАНЕ|ПОВРЕДИ|С КАРТА|ПО БАНКA|НЗОК|ПРОГ.ТИП 2|ЕВРО|НАДБАВКА %1|ОТСТЪПКА %1|СУМА|МЕЖДИННА СУМА|ФАКТУРА No %1|Продажби %1|ЕДИНИЧНА ЦЕНА|КОЛИЧЕСТВО|СУМА|ВСИЧКО ГРУПА %1|ДДС %1=%2|НЕТО СТОЙНОСТ|ОБЩО|%1 АРТИКУЛА|1 АРТИКУЛ|РЕСТО|ПРОДАВАЧ: |ПОЛУЧАТЕЛ: |КУПУВАЧ: |ИДЕНТ. No: |ЗДДС No: |СИСТЕМЕН БОН|ПОРЪЧКА КЪМ КУХНЯ|Дата:     %1|Оператор: %1|Маса:     %1|ЕИК №: %1"
Private Const STR_PROXY_INTERNAL        As String = "Устройството не е намерено|Невалидно състояние на готовност %1|Невалидeн JSON: %1|Неподдържан вид отстъпка %1"
Private Const STR_CONNECTOR_ERRORS      As String = "Не е указано устройство|Грешка от %1: %2|Време за достъп изтече в очакване на отговор"
Private Const ERR_UNKNOWN_PROTOCOL      As String = "Неизвестен протокол: %1"
'--- enums texts
Private Const STR_ENUM_PAYMENT_TYPE     As String = "|Cash|Cheque|Coupon|Voucher|Packing|Maintenance|Damage|Card|Bank|Custom1|Custom2|EUR"
Private Const STR_ENUM_RECEIPT_TYPE     As String = "|Sale|Reversal|Invoice|CreditNote|OrderList"
Private Const STR_ENUM_CGTAXNO_TYPE     As String = "EIC|CitizenNo|ForeignerNo|OfficialNo"
Private Const STR_ENUM_REVERSAL_TYPE    As String = "OperatorError|Refund|TaxBaseReduction"
Private Const STR_ENUM_BARCODE_TYPE     As String = "|Ean8|Ean13|Code128|QRcode"
Private Const STR_ENUM_REPORT_TYPE      As String = "|Daily|MonthlyByReceiptNo|MonthlyByDate|DailyByOperators"
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
Private m_sCommandLog               As String

Private Enum UcsRowTypeEnum
    ucsRwtSell = 1
    ucsRwtPayment
    ucsRwtText
    ucsRwtBarcode
    ucsRwtDiscount
End Enum

Private Type UcsRowData
    RowType         As UcsRowTypeEnum
    ItemName        As String
    Price           As String
    Discount        As String
    TaxGroup        As String
    Quantity        As String
    UnitOfMeasure   As String
    DepartmentNo    As String
    PaymentName     As String
    PaymentType     As String
    Amount          As String
    Text            As String
    BarcodeType     As String
    Height          As String
End Type

Private Type UcsConfigValues
    LocalizedText(0 To [_ucsFscLciMax] - 1) As Variant
End Type

Private Enum UcsInternalErrors
    ucsErrInitProtocol
    ucsErrNotImplemented
    ucsErrInvalidJsonRequest
    ucsErrInvalidReportType
End Enum

Private Enum UcsReceiptTextsEnum
    ucsTxtCashier
End Enum

Private Enum UcsFiscalParamIndexesEnum
    ucsFscPixItemName = 0
    ucsFscPixPrice
    ucsFscPixTaxGroup
    ucsFscPixQuantity
    ucsFscPixDiscount
    ucsFscPixUnitOfMeasure
    ucsFscPixDepartmentNo
    ucsFscPixPaymentName = 0
    ucsFscPixPaymentType
    ucsFscPixAmount
    ucsFscPixText = 0
End Enum

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunction As String)
    Debug.Print "Critical error: " & Err.Description & " [" & MODULE_NAME & "." & sFunction & "]"
    Logger.Log vbLogEventTypeError, MODULE_NAME, sFunction & "(" & Erl & ")", Err.Description
End Sub

Private Sub RaiseError(sFunction As String)
    Debug.Print "Critical error: " & Err.Description & " [" & MODULE_NAME & "." & sFunction & "]"
    Logger.Log vbLogEventTypeError, MODULE_NAME, sFunction & "(" & Erl & ")", Err.Description
    Err.Raise Err.Number, MODULE_NAME & "." & sFunction & "(" & Erl & ")" & vbCrLf & Err.Source, Err.Description
End Sub

Private Sub DebugLog(sFunction As String, sText As String)
    Logger.Log vbLogEventTypeDebug, MODULE_NAME, sFunction, sText
End Sub

'=========================================================================
' Properties
'=========================================================================

Public Property Get LastCommandLog() As String
    LastCommandLog = m_sCommandLog
End Property

Private Property Get LocalizedText(ByVal eIdx As UcsFiscalLocalizedIndexesEnum) As String
    If IsArray(m_uConfig.LocalizedText(eIdx)) Then
        LocalizedText = Join(m_uConfig.LocalizedText(eIdx), "|")
    End If
End Property

Private Property Let LocalizedText(ByVal eIdx As UcsFiscalLocalizedIndexesEnum, sValue As String)
    m_uConfig.LocalizedText(eIdx) = Split(sValue, "|")
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function EnumPorts(sResponse As String) As Boolean
    Const FUNC_NAME     As String = "EnumPorts"
    Dim vPorts          As Variant
    Dim vElem           As Variant
    Dim oFP             As IDeviceProtocol
    Dim oItem           As Object
    Dim oRetVal         As Object
    
    On Error GoTo EH
    JsonItem(oRetVal, "Ok") = True
    vPorts = EnumSerialPorts
    For Each vElem In Array(pvInit("Protocol=" & STR_PROTOCOL_TREMOL, SkipInit:=True), _
                            pvInit("Protocol=" & STR_PROTOCOL_DATECS_FP, SkipInit:=True), _
                            pvInit("Protocol=" & STR_PROTOCOL_DATECS_X, SkipInit:=True), _
                            pvInit("Protocol=" & STR_PROTOCOL_ESCPOS, SkipInit:=True))
        Set oFP = vElem
        vPorts = oFP.AutodetectDevices(vPorts)
    Next
    For Each vElem In vPorts
        Set oItem = Nothing
        If IsArray(vElem) Then
            JsonItem(oItem, "Port") = Zn(Trim$(At(vElem, 0)), Empty)
            JsonItem(oItem, "Speed") = Zn(Trim$(At(vElem, 1)), Empty)
            JsonItem(oItem, "Protocol") = Zn(Trim$(At(vElem, 2)), Empty)
            JsonItem(oItem, "Model") = Zn(Trim$(At(vElem, 3)), Empty)
            JsonItem(oItem, "Firmware") = Zn(Trim$(At(vElem, 4)), Empty)
            JsonItem(oItem, "DeviceSerialNo") = Zn(Trim$(At(vElem, 5)), Empty)
            JsonItem(oItem, "FiscalMemoryNo") = Zn(Trim$(At(vElem, 6)), Empty)
            JsonItem(oItem, "ModelGroup") = Zn(Trim$(At(vElem, 7)), Empty)
        Else
            JsonItem(oItem, "Port") = Zn(Trim$(C_Str(vElem)), Empty)
        End If
        If Not oItem Is Nothing Then
            JsonItem(oRetVal, "SerialPorts/-1") = oItem
        End If
    Next
QH:
    sResponse = JsonDump(oRetVal, Minimize:=True)
    '--- success
    EnumPorts = True
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

''
' Used to retrieve an array with info from the fiscal device including header/footer texts, operator name, current
' date/time, last receipt number, names of payment types, max characters on a row and operator's default password
'
Public Function GetDeviceInfo(sRequest As String, sResponse As String) As Boolean
    Const FUNC_NAME     As String = "GetDeviceInfo"
    Dim oFP             As IDeviceProtocol
    Dim oJson           As Object
    Dim oRetVal         As Object
    Dim sLastErr        As String
    Dim eLastNo         As UcsFiscalErrorsEnum
    
    On Error GoTo EH
    If Not pvCommonProlog(sRequest, oJson, sLastErr, oFP, oRetVal) Then
        GoTo QH
    End If
    JsonItem(oRetVal, "DeviceSerialNo") = Zn(Trim$(oFP.GetDeviceSerialNo()), Empty)
    JsonItem(oRetVal, "FiscalMemoryNo") = Zn(Trim$(oFP.GetFiscalMemoryNo()), Empty)
    JsonItem(oRetVal, "DeviceProtocol") = Zn(Trim$(oFP.GetDeviceProtocol()), Empty)
    JsonItem(oRetVal, "DeviceModel") = Zn(Trim$(oFP.GetDeviceModel()), Empty)
    JsonItem(oRetVal, "FirmwareVersion") = Zn(Trim$(oFP.GetDeviceFirmware()), Empty)
    JsonItem(oRetVal, "DeviceString") = Zn(JsonItem(oJson, "DeviceString"), Empty)
    JsonItem(oRetVal, "CharsPerLine") = Znl(oFP.GetCharsPerLine(), Empty)
    JsonItem(oRetVal, "CommentTextMaxLength") = Znl(oFP.GetCommentTextMaxLength(), Empty)
    JsonItem(oRetVal, "ItemNameMaxLength") = Znl(oFP.GetItemNameMaxLength(), Empty)
    If LenB(C_Str(JsonItem(oJson, "Operator/Code"))) <> 0 Then
        JsonItem(oRetVal, "Operator/Code") = Zn(Trim$(JsonItem(oJson, "Operator/Code")), Empty)
        JsonItem(oRetVal, "Operator/Name") = Zn(Trim$(oFP.GetOperatorName(JsonItem(oJson, "Operator/Code"))), Empty)
        JsonItem(oRetVal, "Operator/DefaultPassword") = Zn(oFP.GetDefaultPassword(JsonItem(oJson, "Operator/Code")), Empty)
    End If
    pvCommonIncludesMiddleware oFP, oJson, oRetVal
QH:
    GetDeviceInfo = pvCommonEpilog(oFP, sLastErr, eLastNo, oRetVal, sResponse)
    Exit Function
EH:
    If pvCommonErrHandler(FUNC_NAME, oFP, sLastErr, eLastNo) Then
        Resume QH
    End If
End Function

''
' Used to batch print a fiscal receipt or invoice
'
Public Function PrintReceipt(sRequest As String, sResponse As String) As Boolean
    Const FUNC_NAME     As String = "PrintReceipt"
    Dim oFP             As IDeviceProtocol
    Dim oJson           As Object
    Dim oRetVal         As Object
    Dim sLastErr        As String
    Dim eLastNo         As UcsFiscalErrorsEnum
    Dim vElem           As Variant
    Dim lIdx            As Long
    Dim sResumeToken    As String
    Dim sPrevReceiptNo  As String
    Dim eReceiptType    As UcsFiscalReceiptTypeEnum
    Dim sReceiptNo      As String
    Dim dReceiptDate    As Date
    Dim dblReceiptAmount As Double
    Dim sInvoiceNo      As String
    
    On Error GoTo EH
    If Not pvCommonProlog(sRequest, oJson, sLastErr, oFP, oRetVal) Then
        GoTo QH
    End If
    pvCommonIncludesMiddleware oFP, oJson, oRetVal
    sResumeToken = JsonItem(oJson, "ResumeToken")
    If LenB(sResumeToken) = 0 Then
        If Not oFP.CancelReceipt() Then
            sLastErr = Zn(oFP.GetLastError(eLastNo), pvInternal(ucsErrNotImplemented))
            DebugLog FUNC_NAME, "Not oFP.CancelReceipt(), sLastErr=" & sLastErr & ", eLastNo=" & eLastNo
            GoTo QH
        End If
    End If
    sPrevReceiptNo = JsonItem(oJson, "PrevReceiptNo")
    If LenB(sPrevReceiptNo) <> 0 Then
        JsonItem(oRetVal, "PrevReceiptNo") = StripZeros(oFP.GetLastQRCodeInfo())
        If LenB(JsonItem(oRetVal, "PrevReceiptNo")) = 0 Then
            JsonItem(oRetVal, "PrevReceiptNo") = StripZeros(oFP.GetLastReceiptNo())
        End If
        '--- check if receipt already printed
        If LenB(Trim$(sPrevReceiptNo)) <> 0 Then
            If Pad(sPrevReceiptNo, -Len(JsonItem(oRetVal, "PrevReceiptNo"))) <> JsonItem(oRetVal, "PrevReceiptNo") And C_Lng(JsonItem(oJson, "PrevReceiptCount")) <> 0 Then
                If C_Str(JsonItem(oJson, "PrevReceiptCount")) <> At(oFP.GetTotalsStats(), 0) Then
                    JsonItem(oRetVal, "PrevReceiptNo") = Empty
                    GoTo ReturnReceiptInfo
                End If
            End If
        End If
    End If
    If IsObject(JsonItem(oJson, "Rows")) Then
        eReceiptType = JsonEnumValue(JsonItem(oJson, "ReceiptType"), STR_ENUM_RECEIPT_TYPE)
        If Not oFP.StartReceipt(eReceiptType, _
                JsonItem(oJson, "Operator/Code"), _
                Zn(JsonItem(oJson, "Operator/Name"), pvText(ucsTxtCashier)), _
                JsonItem(oJson, "Operator/Password"), _
                UniqueSaleNo:=JsonItem(oJson, "UniqueSaleNo"), _
                TableNo:=JsonItem(oJson, "TableNo"), _
                DisablePrinting:=JsonBoolItem(oJson, "DisablePrinting"), _
                InvDocNo:=JsonItem(oJson, "Invoice/DocNo"), _
                InvCgTaxNo:=JsonItem(oJson, "Invoice/CgTaxNo"), _
                InvCgTaxNoType:=JsonEnumValue(JsonItem(oJson, "Invoice/CgTaxNoType"), STR_ENUM_CGTAXNO_TYPE), _
                InvCgVatNo:=JsonItem(oJson, "Invoice/CgVatNo"), _
                InvCgName:=JsonItem(oJson, "Invoice/CgName"), _
                InvCgAddress:=JsonItem(oJson, "Invoice/CgAddress"), _
                InvCgPrsReceive:=JsonItem(oJson, "Invoice/CgPrsReceive"), _
                OwnData:=JsonItem(oJson, "OwnData"), _
                RevType:=JsonEnumValue(JsonItem(oJson, "Reversal/ReversalType"), STR_ENUM_REVERSAL_TYPE), _
                RevReceiptNo:=JsonItem(oJson, "Reversal/ReceiptNo"), _
                RevReceiptDate:=C_Date(JsonItem(oJson, "Reversal/ReceiptDateTime")), _
                RevFiscalMemoryNo:=JsonItem(oJson, "Reversal/FiscalMemoryNo"), _
                RevInvoiceNo:=JsonItem(oJson, "Reversal/InvoiceNo"), _
                RevReason:=JsonItem(oJson, "Reversal/Reason")) Then
            sLastErr = Zn(oFP.GetLastError(eLastNo), pvInternal(ucsErrNotImplemented))
            DebugLog FUNC_NAME, "Not oFP.StartReceipt(), sLastErr=" & sLastErr & ", eLastNo=" & eLastNo
            GoTo QH
        End If
        For Each vElem In JsonKeys(oJson, "Rows")
            With pvGetRowData(C_Obj(JsonItem(oJson, "Rows/" & vElem)))
                Select Case .RowType
                Case ucsRwtSell
                    oFP.AddPLU .ItemName, C_Dbl(.Price), IIf(LenB(.Quantity) <> 0, C_Dbl(.Quantity), 1), pvGetTaxGroup(.TaxGroup), .UnitOfMeasure, C_Dbl(.DepartmentNo)
                    If Abs(C_Dbl(.Discount)) > DBL_EPSILON Then
                        oFP.AddDiscount ucsFscDscPlu, -C_Dbl(.Discount)
                    End If
                Case ucsRwtPayment
                    lIdx = JsonEnumValue(.PaymentType, STR_ENUM_PAYMENT_TYPE)
                    If lIdx < MIN_PMT_TYPE Or lIdx > MAX_PMT_TYPE Then
                        lIdx = DEF_PMT_TYPE
                    End If
                    oFP.AddPayment lIdx, Zn(.PaymentName, Trim$(oFP.GetPaymentName(lIdx))), C_Dbl(.Amount)
                Case ucsRwtText
                    oFP.AddLine .Text
                Case ucsRwtBarcode
                    oFP.AddBarcode JsonEnumValue(.BarcodeType, STR_ENUM_BARCODE_TYPE), .Text, C_Dbl(.Height)
                Case ucsRwtDiscount
                    oFP.AddDiscount C_Lng(.Discount), -C_Dbl(.Amount)
                End Select
            End With
        Next
        If Not oFP.PrintReceipt(sResumeToken, sReceiptNo, dReceiptDate, dblReceiptAmount, sInvoiceNo) Then
            sLastErr = Zn(oFP.GetLastError(eLastNo), pvInternal(ucsErrNotImplemented))
            DebugLog FUNC_NAME, "Not oFP.EndReceipt(), sLastErr=" & sLastErr & ", eLastNo=" & eLastNo
            GoTo QH
        End If
        If JsonBoolItem(oJson, "OpenDrawer") Then
            oFP.OpenDrawer
        End If
    End If
    If JsonBoolItem(oJson, "PrintDuplicate") Then
        If Not oFP.PrintDuplicate(JsonItem(oJson, "Invoice/DocNo"), sReceiptNo, dReceiptDate) Then
            sLastErr = Zn(oFP.GetLastError(eLastNo), pvInternal(ucsErrNotImplemented))
            DebugLog FUNC_NAME, "Not oFP.PrintDuplicate(), sLastErr=" & sLastErr & ", eLastNo=" & eLastNo
            GoTo QH
        End If
    End If
ReturnReceiptInfo:
    If LenB(sReceiptNo) = 0 And dReceiptDate = 0 Then
        sReceiptNo = oFP.GetLastQRCodeInfo(dReceiptDate)
        If LenB(sReceiptNo) = 0 And dReceiptDate = 0 Then
            sReceiptNo = oFP.GetLastReceiptNo()
            dReceiptDate = oFP.GetClock()
        End If
    End If
    JsonItem(oRetVal, "ReceiptNo") = StripZeros(sReceiptNo)
    JsonItem(oRetVal, "ReceiptDateTime") = dReceiptDate
    JsonItem(oRetVal, "DeviceSerialNo") = Trim$(oFP.GetDeviceSerialNo())
    JsonItem(oRetVal, "FiscalMemoryNo") = Trim$(oFP.GetFiscalMemoryNo())
    If Abs(dblReceiptAmount) > DBL_EPSILON Then
        JsonItem(oRetVal, "ReceiptAmount") = dblReceiptAmount
    End If
    Select Case eReceiptType
    Case ucsFscRcpInvoice, ucsFscRcpCreditNote
        If LenB(sInvoiceNo) = 0 Then
            sInvoiceNo = oFP.GetLastInvoiceNo()
        End If
    End Select
    JsonItem(oRetVal, "LastInvoiceNo") = Zn(StripZeros(sInvoiceNo), Empty)
QH:
    If LenB(sLastErr) <> 0 And Not oFP Is Nothing Then
        JsonItem(oRetVal, "ResumeToken") = Zn(oFP.GetResumeToken(), Empty)
    Else
        JsonItem(oRetVal, "ResumeToken") = Empty
    End If
    PrintReceipt = pvCommonEpilog(oFP, sLastErr, eLastNo, oRetVal, sResponse)
    Exit Function
EH:
    If pvCommonErrHandler(FUNC_NAME, oFP, sLastErr, eLastNo) Then
        Resume QH
    End If
End Function

''
' Used to print a fiscal report as supported by the fiscal device
'
Public Function PrintReport(sRequest As String, sResponse As String) As Boolean
    Const FUNC_NAME     As String = "PrintReport"
    Dim oFP             As IDeviceProtocol
    Dim oJson           As Object
    Dim oRetVal         As Object
    Dim sLastErr        As String
    Dim eLastNo         As UcsFiscalErrorsEnum
    Dim bResult         As Boolean
    Dim ePrintType      As UcsFiscalPrintReportTypeEnum
    Dim eReportType     As UcsFiscalReportsTypeEnum
    Dim sReceiptNo      As String
    Dim dReceiptDate    As Date
    
    On Error GoTo EH
    If Not pvCommonProlog(sRequest, oJson, sLastErr, oFP, oRetVal) Then
        GoTo QH
    End If
    pvCommonIncludesMiddleware oFP, oJson, oRetVal
    ePrintType = JsonEnumValue(JsonItem(oJson, "ReportType"), STR_ENUM_REPORT_TYPE)
    Select Case ePrintType
    Case ucsFscRptDaily
        If JsonBoolItem(oJson, "IsItems") And JsonBoolItem(oJson, "IsDepartments") Then
            eReportType = ucsFscRstDailyByItemsAndDepartments
        ElseIf JsonBoolItem(oJson, "IsItems") Then
            eReportType = ucsFscRstDailyByItems
        ElseIf JsonBoolItem(oJson, "IsDepartments") Then
            eReportType = ucsFscRstDailyByDepartments
        Else
            eReportType = ucsFscRstDaily
        End If
        If JsonBoolItem(oJson, "IsClear") Then
            bResult = oFP.PrintDailyZReport(eReportType, sReceiptNo, dReceiptDate)
        Else
            bResult = oFP.PrintDailyXReport(eReportType, sReceiptNo, dReceiptDate)
        End If
    Case ucsFscRptMonthlyByReceiptNo
        '--- ToDo: impl
    Case ucsFscRptMonthlyByDate
        eReportType = IIf(JsonBoolItem(oJson, "IsDetailed"), ucsFscRstMonthlyDetailed, ucsFscRstMonthlyShort)
        If C_Date(JsonItem(oJson, "FromDate")) <> 0 And C_Date(JsonItem(oJson, "ToDate")) <> 0 Then
            bResult = oFP.PrintMonthlyReport(eReportType, C_Date(JsonItem(oJson, "FromDate")), C_Date(JsonItem(oJson, "ToDate")), sReceiptNo, dReceiptDate)
        End If
    Case ucsFscRptDailyByOperators
        eReportType = ucsFscRstDailyByOperators
        If JsonBoolItem(oJson, "IsClear") Then
            bResult = oFP.PrintDailyZReport(eReportType, sReceiptNo, dReceiptDate)
        Else
            bResult = oFP.PrintDailyXReport(eReportType, sReceiptNo, dReceiptDate)
        End If
    Case Else
        sLastErr = Printf(pvInternal(ucsErrInvalidReportType), ePrintType)
        DebugLog FUNC_NAME, "Case Else, sLastErr=" & sLastErr & ", eLastNo=" & eLastNo
        GoTo QH
    End Select
    If Not bResult Then
        sLastErr = Zn(oFP.GetLastError(eLastNo), pvInternal(ucsErrNotImplemented))
        DebugLog FUNC_NAME, "Not bResult, sLastErr=" & sLastErr & ", eLastNo=" & eLastNo
        GoTo QH
    End If
    If LenB(sReceiptNo) And dReceiptDate = 0 Then
        sReceiptNo = oFP.GetLastReceiptNo()
        dReceiptDate = oFP.GetClock()
    End If
    JsonItem(oRetVal, "ReceiptNo") = StripZeros(sReceiptNo)
    JsonItem(oRetVal, "ReceiptDateTime") = dReceiptDate
    JsonItem(oRetVal, "DeviceSerialNo") = Trim$(oFP.GetDeviceSerialNo())
    JsonItem(oRetVal, "FiscalMemoryNo") = Trim$(oFP.GetFiscalMemoryNo())
QH:
    PrintReport = pvCommonEpilog(oFP, sLastErr, eLastNo, oRetVal, sResponse)
    Exit Function
EH:
    If pvCommonErrHandler(FUNC_NAME, oFP, sLastErr, eLastNo) Then
        Resume QH
    End If
End Function

''
' Used to send an impulse to the drawer opener. The fiscal device has to be connected with a compatible cash
' drawer for this command to work
'
Public Function OpenDrawer(sRequest As String, sResponse As String) As Boolean
    Const FUNC_NAME     As String = "OpenDrawer"
    Dim oFP             As IDeviceProtocol
    Dim oJson           As Object
    Dim oRetVal         As Object
    Dim sLastErr        As String
    Dim eLastNo         As UcsFiscalErrorsEnum
    
    On Error GoTo EH
    If Not pvCommonProlog(sRequest, oJson, sLastErr, oFP, oRetVal) Then
        GoTo QH
    End If
    pvCommonIncludesMiddleware oFP, oJson, oRetVal
    If Not oFP.OpenDrawer() Then
        sLastErr = Zn(oFP.GetLastError(eLastNo), pvInternal(ucsErrNotImplemented))
        DebugLog FUNC_NAME, "Not oFP.OpenDrawer(), sLastErr=" & sLastErr & ", eLastNo=" & eLastNo
        GoTo QH
    End If
QH:
    OpenDrawer = pvCommonEpilog(oFP, sLastErr, eLastNo, oRetVal, sResponse)
    Exit Function
EH:
    If pvCommonErrHandler(FUNC_NAME, oFP, sLastErr, eLastNo) Then
        Resume QH
    End If
End Function

''
' Used to adjust the date and time of the fiscal device clock
'
Public Function InitDateTime(sRequest As String, sResponse As String) As Boolean
    Const FUNC_NAME     As String = "InitDateTime"
    Dim oFP             As IDeviceProtocol
    Dim oJson           As Object
    Dim oRetVal         As Object
    Dim sLastErr        As String
    Dim eLastNo         As UcsFiscalErrorsEnum
    Dim dDateTime       As Date
    Dim bResult         As Boolean
    Dim sStatus         As String
    Dim lTolerance      As Long
    Dim dCurrent        As Date
    
    On Error GoTo EH
    If Not pvCommonProlog(sRequest, oJson, sLastErr, oFP, oRetVal) Then
        GoTo QH
    End If
    pvCommonIncludesMiddleware oFP, oJson, oRetVal
    dDateTime = C_Date(JsonItem(oJson, "DeviceDateTime"))
    lTolerance = C_Lng(JsonItem(oJson, "AdjustTolerance"))
    dCurrent = oFP.GetClock()
    If dCurrent = 0 Or (dDateTime = 0 And lTolerance = 0) Then
        GoTo SkipSetClock
    End If
    If dDateTime = 0 Then
        dDateTime = GetCurrentNow
    End If
    If lTolerance = 0 Then
        bResult = oFP.SetClock(dDateTime)
    ElseIf Abs(DateDiff("s", dCurrent, dDateTime)) >= lTolerance Then
        bResult = oFP.SetClock(dDateTime)
    Else
        GoTo SkipSetClock
    End If
    If Not bResult Then
        sLastErr = Zn(oFP.GetLastError(eLastNo), pvInternal(ucsErrNotImplemented))
        DebugLog FUNC_NAME, "Not bResult, sLastErr=" & sLastErr & ", eLastNo=" & eLastNo
        GoTo QH
    Else
        JsonItem(oRetVal, "PreviousDateTime") = dCurrent
        dCurrent = oFP.GetClock()
    End If
SkipSetClock:
    If oFP.GetDeviceStatus(sStatus) Then
        JsonItem(oRetVal, "DeviceStatus") = sStatus
    End If
    JsonItem(oRetVal, "DeviceDateTime") = dCurrent
QH:
    InitDateTime = pvCommonEpilog(oFP, sLastErr, eLastNo, oRetVal, sResponse)
    Exit Function
EH:
    If pvCommonErrHandler(FUNC_NAME, oFP, sLastErr, eLastNo) Then
        Resume QH
    End If
End Function

''
' Used to retrieve number of receipts printed since last z-report and the date/time of the last receipt printed
' Used to retrieve accumulated totals by VAT groups (1-8) since last z-report
' Used to retrieve accumulated totals by payment types since last z-report
'
Public Function GetDailyTotals(sRequest As String, sResponse As String) As Boolean
    Const FUNC_NAME     As String = "GetDailyTotals"
    Dim oFP             As IDeviceProtocol
    Dim oJson           As Object
    Dim oRetVal         As Object
    Dim bIncludeAll     As Variant
    Dim sLastErr        As String
    Dim eLastNo         As UcsFiscalErrorsEnum
    Dim vTotals         As Variant
    Dim lIdx            As Long
    Dim oItem           As Object
    Dim dblTotalAmount  As Double
    Dim dblTotalReversal As Double
    
    On Error GoTo EH
    If Not pvCommonProlog(sRequest, oJson, sLastErr, oFP, oRetVal) Then
        GoTo QH
    End If
    pvCommonIncludesMiddleware oFP, oJson, oRetVal
    bIncludeAll = JsonBoolItem(oJson, "IncludeAllTotals")
    If JsonBoolItem(oJson, "IncludeStats") Or bIncludeAll Then
        vTotals = oFP.GetTotalsStats()
        If IsArray(vTotals) Then
            JsonItem(oRetVal, "NumReceipts") = vTotals(0)
            JsonItem(oRetVal, "LastZReportDateTime") = vTotals(1)
        End If
        JsonItem(oRetVal, "TotalAmount") = 0#
        JsonItem(oRetVal, "TotalReversal") = 0#
    End If
    If JsonBoolItem(oJson, "IncludeServiceDeposit") Or bIncludeAll Then
        vTotals = oFP.PrintServiceDeposit(JsonItem(oJson, "Operator/Code"), JsonItem(oJson, "Operator/Password"), 0)
        If IsArray(vTotals) Then
            JsonItem(oRetVal, "TotalAvailable") = vTotals(0)
            JsonItem(oRetVal, "TotalDeposits") = vTotals(1)
            JsonItem(oRetVal, "TotalWithdraws") = vTotals(2)
        End If
    End If
    If JsonBoolItem(oJson, "IncludeStats") Or JsonBoolItem(oJson, "IncludeTaxGroups") Or bIncludeAll Then
        vTotals = oFP.GetTotalsByTaxGroups()
        If IsArray(vTotals) Then
            For lIdx = 0 To UBound(vTotals)
                If IsArray(vTotals(lIdx)) Then
                    If JsonBoolItem(oJson, "IncludeTaxGroups") Or bIncludeAll Then
                        Set oItem = Nothing
                        JsonItem(oItem, "TaxGroup") = lIdx + 1
                        If vTotals(lIdx)(0) >= 0 And vTotals(lIdx)(0) < 100 Then
                            JsonItem(oItem, "VatRate") = vTotals(lIdx)(0)
                        End If
                        JsonItem(oItem, "Amount") = vTotals(lIdx)(1)
                        If UBound(vTotals(lIdx)) >= 2 Then
                            JsonItem(oItem, "Reversal") = vTotals(lIdx)(2)
                        End If
                        JsonItem(oRetVal, "TotalsByTaxGroups/-1") = oItem
                    End If
                    dblTotalAmount = dblTotalAmount + vTotals(lIdx)(1)
                    If UBound(vTotals(lIdx)) >= 2 Then
                        dblTotalReversal = dblTotalReversal + vTotals(lIdx)(2)
                    End If
                End If
            Next
        End If
        JsonItem(oRetVal, "TotalAmount") = dblTotalAmount
        JsonItem(oRetVal, "TotalReversal") = dblTotalReversal
    End If
    If JsonBoolItem(oJson, "IncludePaymentTypes") Or bIncludeAll Then
        vTotals = oFP.GetTotalsByPaymentTypes()
        If IsArray(vTotals) Then
            For lIdx = 0 To UBound(vTotals)
                If IsArray(vTotals(lIdx)) Then
                    Set oItem = Nothing
                    JsonItem(oItem, "PaymentType") = lIdx + 1
                    JsonItem(oItem, "PaymentName") = vTotals(lIdx)(0)
                    JsonItem(oItem, "Amount") = vTotals(lIdx)(1)
                    If UBound(vTotals(lIdx)) >= 2 Then
                        JsonItem(oItem, "Reversal") = vTotals(lIdx)(2)
                    End If
                    JsonItem(oRetVal, "TotalsByPaymentTypes/-1") = oItem
                End If
            Next
        End If
    End If
QH:
    GetDailyTotals = pvCommonEpilog(oFP, sLastErr, eLastNo, oRetVal, sResponse)
    Exit Function
EH:
    If pvCommonErrHandler(FUNC_NAME, oFP, sLastErr, eLastNo) Then
        Resume QH
    End If
End Function

''
' Used to debit (and credit) the fiscal device with cash total that is deposited (or withdrawn) outside of regular client
' sales, e.g. depositing initial daily change. Can be used to retrieve total daily deposits/withdraws too
'
Public Function PrintServiceDeposit(sRequest As String, sResponse As String) As Boolean
    Const FUNC_NAME     As String = "PrintServiceDeposit"
    Dim oFP             As IDeviceProtocol
    Dim oJson           As Object
    Dim oRetVal         As Object
    Dim sLastErr        As String
    Dim eLastNo         As UcsFiscalErrorsEnum
    Dim vTotals         As Variant
    Dim sReceiptNo      As String
    Dim dReceiptDate    As Date
    Dim dblReceiptAmount As Double
    
    On Error GoTo EH
    If Not pvCommonProlog(sRequest, oJson, sLastErr, oFP, oRetVal) Then
        GoTo QH
    End If
    pvCommonIncludesMiddleware oFP, oJson, oRetVal
    vTotals = oFP.PrintServiceDeposit(C_Str(JsonItem(oJson, "Operator/Code")), C_Str(JsonItem(oJson, "Operator/Password")), _
        C_Dbl(JsonItem(oJson, "Amount")), sReceiptNo, dReceiptDate, dblReceiptAmount)
    If Not IsArray(vTotals) Then
        sLastErr = Zn(oFP.GetLastError(eLastNo), pvInternal(ucsErrNotImplemented))
        DebugLog FUNC_NAME, "Not IsArray(vTotals), sLastErr=" & sLastErr & ", eLastNo=" & eLastNo
        GoTo QH
    ElseIf UBound(vTotals) < 0 Then
        sLastErr = Zn(oFP.GetLastError(eLastNo), pvInternal(ucsErrNotImplemented))
        DebugLog FUNC_NAME, "UBound(vTotals) < 0, sLastErr=" & sLastErr & ", eLastNo=" & eLastNo
        GoTo QH
    End If
    If Abs(C_Dbl(JsonItem(oJson, "Amount"))) > DBL_EPSILON Then
        If LenB(sReceiptNo) And dReceiptDate = 0 Then
            sReceiptNo = oFP.GetLastReceiptNo()
            dReceiptDate = oFP.GetClock()
        End If
        JsonItem(oRetVal, "ReceiptNo") = StripZeros(sReceiptNo)
        JsonItem(oRetVal, "ReceiptDateTime") = dReceiptDate
        JsonItem(oRetVal, "DeviceSerialNo") = Trim$(oFP.GetDeviceSerialNo())
        JsonItem(oRetVal, "FiscalMemoryNo") = Trim$(oFP.GetFiscalMemoryNo())
        If Abs(dblReceiptAmount) > DBL_EPSILON Then
            JsonItem(oRetVal, "ReceiptAmount") = dblReceiptAmount
        End If
    End If
    JsonItem(oRetVal, "TotalAvailable") = vTotals(0)
    JsonItem(oRetVal, "TotalDeposits") = vTotals(1)
    JsonItem(oRetVal, "TotalWithdraws") = vTotals(2)
QH:
    PrintServiceDeposit = pvCommonEpilog(oFP, sLastErr, eLastNo, oRetVal, sResponse)
    Exit Function
EH:
    If pvCommonErrHandler(FUNC_NAME, oFP, sLastErr, eLastNo) Then
        Resume QH
    End If
End Function

''
' Used to retrieve additional device status without performing any other operation, e.g. can be used to poll the
' device if the operator is ready with loading a new paper roll after receiving out-of-paper error on receipt printing
'
Public Function GetDeviceStatus(sRequest As String, sResponse As String) As Boolean
    Const FUNC_NAME     As String = "GetDeviceStatus"
    Dim oFP             As IDeviceProtocol
    Dim oJson           As Object
    Dim oRetVal         As Object
    Dim sLastErr        As String
    Dim eLastNo         As UcsFiscalErrorsEnum
    Dim eStatusCode     As UcsFiscalStatusCodeEnum
    Dim sStatusText     As String
    
    On Error GoTo EH
    If Not pvCommonProlog(sRequest, oJson, sLastErr, oFP, oRetVal) Then
        GoTo QH
    End If
    pvCommonIncludesMiddleware oFP, oJson, oRetVal
    eStatusCode = oFP.GetDeviceStatus(sStatusText)
    JsonItem(oRetVal, "DeviceStatusCode") = At(Split(STR_ENUM_STATUS_CODE, "|"), eStatusCode, "unknown " & eStatusCode)
    JsonItem(oRetVal, "DeviceStatus") = Zn(sStatusText, Empty)
QH:
    GetDeviceStatus = pvCommonEpilog(oFP, sLastErr, eLastNo, oRetVal, sResponse)
    Exit Function
EH:
    If pvCommonErrHandler(FUNC_NAME, oFP, sLastErr, eLastNo) Then
        Resume QH
    End If
End Function

''
' Used to set localized strings used for non-fiscal receipt texts and component internal errors
'
Public Sub SetLocalizedText(ByVal Index As UcsFiscalLocalizedIndexesEnum, Text As String)
    LocalizedText(Index) = Text
End Sub

'= callbacks =============================================================

Public Sub OnBeforeCommitReceipt(oCtx As Object, Cancel As Boolean)
Attribute OnBeforeCommitReceipt.VB_MemberFlags = "40"
    RaiseEvent BeforeCommitReceipt(oCtx, Cancel)
End Sub

'= private ===============================================================

Private Function pvInit(DeviceString As String, Optional ByVal SkipInit As Boolean) As IDeviceProtocol
    Dim oOptions        As Object
    Dim sProtocol       As String
    Dim lConfigIdx      As Long
    
    Set oOptions = ParseDeviceString(DeviceString)
    sProtocol = UCase$(JsonItem(oOptions, "Protocol"))
    '--- figure out model
    Select Case sProtocol
    Case STR_PROTOCOL_DATECS_X
        Set pvInit = New cDatecsProtocol
        pvInit.SetLocalizedText ucsFscLciInternalErrors, GetConfigValue(sProtocol, CFG_INTERNAL_ERRORS, STR_DATECS_INTERNAL)
        pvInit.SetLocalizedText ucsFscLciPrinterStatuses, GetConfigValue(sProtocol, CFG_PRINTER_STATUSES, STR_DATECS_STATUSES)
        pvInit.SetLocalizedText ucsFscLciPrinterErrors, GetConfigValue(sProtocol, CFG_PRINTER_ERRORS, STR_DATECS_ERRORS)
        pvInit.SetLocalizedText ucsFscLciMemoryTexts, GetConfigValue(sProtocol, CFG_MEMORY_TEXTS, STR_DATECS_MEMORY)
        pvInit.SetLocalizedText ucsFscLciReceiptTexts, GetConfigValue(sProtocol, CFG_RECEIPT_TEXTS, STR_DATECS_TEXTS)
        pvInit.SetLocalizedText ucsFscLciConnectorErrors, GetConfigValue(sProtocol, CFG_CONNECTOR_ERRORS, STR_CONNECTOR_ERRORS)
    Case STR_PROTOCOL_DATECS_FP, STR_PROTOCOL_DAISY, STR_PROTOCOL_INCOTEX, STR_PROTOCOL_ELTRADE
        Set pvInit = New cIslProtocol
        pvInit.SetLocalizedText ucsFscLciInternalErrors, GetConfigValue(sProtocol, CFG_INTERNAL_ERRORS, STR_ISL_INTERNAL)
        pvInit.SetLocalizedText ucsFscLciDipSwitchesTexts, GetConfigValue(sProtocol, CFG_DIP_SWITCHES_TEXTS, STR_ISL_DIP_SWITCHES)
        pvInit.SetLocalizedText ucsFscLciReceiptTexts, GetConfigValue(sProtocol, CFG_RECEIPT_TEXTS, STR_ISL_TEXTS)
        pvInit.SetLocalizedText ucsFscLciConnectorErrors, GetConfigValue(sProtocol, CFG_CONNECTOR_ERRORS, STR_CONNECTOR_ERRORS)
        lConfigIdx = ucsIslCfgDatecsModelA '--- FPs
        pvInit.SetLocalizedText ucsFscLciPrinterStatuses, GetConfigValue(sProtocol & lConfigIdx, CFG_PRINTER_STATUSES, STR_ISL_STATUSES_A), lConfigIdx
        pvInit.SetLocalizedText ucsFscLciPrinterErrors, GetConfigValue(sProtocol & lConfigIdx, CFG_PRINTER_ERRORS, STR_ISL_ERRORS_A), lConfigIdx
        pvInit.SetLocalizedText ucsFscLciMemoryTexts, GetConfigValue(sProtocol & lConfigIdx, CFG_MEMORY_TEXTS, STR_ISL_MEMORY_A), lConfigIdx
        lConfigIdx = ucsIslCfgDatecsModelB '--- ECRs
        pvInit.SetLocalizedText ucsFscLciPrinterStatuses, GetConfigValue(sProtocol & lConfigIdx, CFG_PRINTER_STATUSES, STR_ISL_STATUSES_B), lConfigIdx
        pvInit.SetLocalizedText ucsFscLciPrinterErrors, GetConfigValue(sProtocol & lConfigIdx, CFG_PRINTER_ERRORS, STR_ISL_ERRORS_B), lConfigIdx
        pvInit.SetLocalizedText ucsFscLciMemoryTexts, GetConfigValue(sProtocol & lConfigIdx, CFG_MEMORY_TEXTS, STR_ISL_MEMORY_B), lConfigIdx
        lConfigIdx = ucsIslCfgDaisy
        pvInit.SetLocalizedText ucsFscLciPrinterStatuses, GetConfigValue(sProtocol & lConfigIdx, CFG_PRINTER_STATUSES, STR_ISL_STATUSES_DAISY), lConfigIdx
        pvInit.SetLocalizedText ucsFscLciPrinterErrors, GetConfigValue(sProtocol & lConfigIdx, CFG_PRINTER_ERRORS, STR_ISL_ERRORS_DAISY), lConfigIdx
        pvInit.SetLocalizedText ucsFscLciMemoryTexts, GetConfigValue(sProtocol & lConfigIdx, CFG_MEMORY_TEXTS, STR_ISL_MEMORY_DAISY), lConfigIdx
        pvInit.SetLocalizedText ucsFscLciExtendedErrors, GetConfigValue(sProtocol & lConfigIdx, CFG_EXTENDED_ERRORS, STR_ISL_EXTENDED_DAISY), lConfigIdx
        lConfigIdx = ucsIslCfgIncotex
        pvInit.SetLocalizedText ucsFscLciPrinterStatuses, GetConfigValue(sProtocol & lConfigIdx, CFG_PRINTER_STATUSES, STR_ISL_STATUSES_INCO), lConfigIdx
        pvInit.SetLocalizedText ucsFscLciPrinterErrors, GetConfigValue(sProtocol & lConfigIdx, CFG_PRINTER_ERRORS, STR_ISL_ERRORS_INCO), lConfigIdx
        pvInit.SetLocalizedText ucsFscLciMemoryTexts, GetConfigValue(sProtocol & lConfigIdx, CFG_MEMORY_TEXTS, STR_ISL_MEMORY_INCO), lConfigIdx
        pvInit.SetLocalizedText ucsFscLciExtendedErrors, GetConfigValue(sProtocol & lConfigIdx, CFG_EXTENDED_ERRORS, STR_ISL_EXTENDED_INCO), lConfigIdx
        lConfigIdx = ucsIslCfgEltrade
        pvInit.SetLocalizedText ucsFscLciPrinterStatuses, GetConfigValue(sProtocol & lConfigIdx, CFG_PRINTER_STATUSES, STR_ISL_STATUSES_ELTRADE), lConfigIdx
        pvInit.SetLocalizedText ucsFscLciPrinterErrors, GetConfigValue(sProtocol & lConfigIdx, CFG_PRINTER_ERRORS, STR_ISL_ERRORS_ELTRADE), lConfigIdx
        pvInit.SetLocalizedText ucsFscLciMemoryTexts, GetConfigValue(sProtocol & lConfigIdx, CFG_MEMORY_TEXTS, STR_ISL_MEMORY_ELTRADE), lConfigIdx
    Case STR_PROTOCOL_TREMOL
        Set pvInit = New cTremolProtocol
        pvInit.SetLocalizedText ucsFscLciInternalErrors, GetConfigValue(sProtocol, CFG_INTERNAL_ERRORS, STR_TREMOL_INTERNAL)
        pvInit.SetLocalizedText ucsFscLciPrinterStatuses, GetConfigValue(sProtocol, CFG_PRINTER_STATUSES, STR_TREMOL_STATUSES)
        pvInit.SetLocalizedText ucsFscLciPrinterErrors, GetConfigValue(sProtocol, CFG_PRINTER_ERRORS, STR_TREMOL_ERRORS)
        pvInit.SetLocalizedText ucsFscLciReceiptTexts, GetConfigValue(sProtocol, CFG_RECEIPT_TEXTS, STR_TREMOL_TEXTS)
        pvInit.SetLocalizedText ucsFscLciConnectorErrors, GetConfigValue(sProtocol, CFG_CONNECTOR_ERRORS, STR_CONNECTOR_ERRORS)
    Case STR_PROTOCOL_ESCPOS
        Set pvInit = New cEscPosProtocol
        pvInit.SetLocalizedText ucsFscLciInternalErrors, GetConfigValue(sProtocol, CFG_INTERNAL_ERRORS, STR_ESP_POS_INTERNAL)
        pvInit.SetLocalizedText ucsFscLciReceiptTexts, GetConfigValue(sProtocol, CFG_RECEIPT_TEXTS, STR_ESP_POS_TEXTS)
        pvInit.SetLocalizedText ucsFscLciConnectorErrors, GetConfigValue(sProtocol, CFG_CONNECTOR_ERRORS, STR_CONNECTOR_ERRORS)
    Case STR_PROTOCOL_PROXY
        Set pvInit = New cProxyProtocol
        pvInit.SetLocalizedText ucsFscLciInternalErrors, GetConfigValue(sProtocol, CFG_INTERNAL_ERRORS, STR_PROXY_INTERNAL)
    Case Else
        Err.Raise vbObjectError, , Printf(ERR_UNKNOWN_PROTOCOL, sProtocol)
    End Select
    If Not SkipInit Then
        If Not pvInit.Init(DeviceString, Me) Then
            Err.Raise vbObjectError, , Zn(pvInit.GetLastError(), Printf(pvInternal(ucsErrInitProtocol), Zn(sProtocol, STR_NONE)))
        End If
        If LenB(pvInit.GetLastError()) Then
            Err.Raise vbObjectError, , pvInit.GetLastError()
        End If
    End If
End Function

Private Function pvGetRowData(oRow As Object) As UcsRowData
    Const FUNC_NAME     As String = "pvGetRowData"
    Dim vRow            As Variant
    
    On Error GoTo EH
    If Not IsEmpty(JsonItem(oRow, "Price")) Then
        pvGetRowData.RowType = ucsRwtSell
        pvGetRowData.ItemName = C_Str(JsonItem(oRow, "ItemName"))
        pvGetRowData.Price = C_Str(JsonItem(oRow, "Price"))
        pvGetRowData.TaxGroup = C_Str(JsonItem(oRow, "TaxGroup"))
        pvGetRowData.Quantity = C_Str(JsonItem(oRow, "Quantity"))
        pvGetRowData.Discount = C_Str(JsonItem(oRow, "Discount"))
        pvGetRowData.UnitOfMeasure = C_Str(JsonItem(oRow, "UnitOfMeasure"))
        pvGetRowData.DepartmentNo = C_Str(JsonItem(oRow, "DepartmentNo"))
        GoTo QH
    End If
    If Not IsEmpty(JsonItem(oRow, "PaymentType")) Then
        pvGetRowData.RowType = ucsRwtPayment
        pvGetRowData.PaymentName = C_Str(JsonItem(oRow, "PaymentName"))
        pvGetRowData.PaymentType = C_Str(JsonItem(oRow, "PaymentType"))
        pvGetRowData.Amount = C_Str(JsonItem(oRow, "Amount"))
        GoTo QH
    End If
    If Not IsEmpty(JsonItem(oRow, "BarcodeType")) Then
        pvGetRowData.RowType = ucsRwtBarcode
        pvGetRowData.Text = C_Str(JsonItem(oRow, "Text"))
        pvGetRowData.BarcodeType = C_Str(JsonItem(oRow, "BarcodeType"))
        pvGetRowData.Height = C_Str(JsonItem(oRow, "Height"))
        GoTo QH
    End If
    If Not IsEmpty(JsonItem(oRow, "Text")) Then
        pvGetRowData.RowType = ucsRwtText
        pvGetRowData.Text = C_Str(JsonItem(oRow, "Text"))
        GoTo QH
    End If
    If Not IsEmpty(JsonItem(oRow, "DiscountPercent")) Then
        pvGetRowData.RowType = ucsRwtDiscount
        pvGetRowData.Discount = ucsFscDscSubtotal
        pvGetRowData.Amount = C_Str(JsonItem(oRow, "DiscountPercent"))
        GoTo QH
    End If
    If Not IsEmpty(JsonItem(oRow, "DiscountAmount")) Then
        pvGetRowData.RowType = ucsRwtDiscount
        pvGetRowData.Discount = ucsFscDscSubtotalAbs
        pvGetRowData.Amount = C_Str(JsonItem(oRow, "DiscountAmount"))
        GoTo QH
    End If
    vRow = JsonItem(oRow, "*")
    If UBound(vRow) = ucsFscPixText Then
        pvGetRowData.RowType = ucsRwtText
        pvGetRowData.Text = At(vRow, ucsFscPixText)
        GoTo QH
    End If
    If UBound(vRow) = ucsFscPixAmount Then
        pvGetRowData.RowType = ucsRwtPayment
        pvGetRowData.PaymentName = At(vRow, ucsFscPixPaymentName)
        pvGetRowData.PaymentType = At(vRow, ucsFscPixPaymentType)
        pvGetRowData.Amount = At(vRow, ucsFscPixAmount)
        GoTo QH
    End If
    pvGetRowData.RowType = ucsRwtSell
    pvGetRowData.ItemName = At(vRow, ucsFscPixItemName)
    pvGetRowData.Price = At(vRow, ucsFscPixPrice)
    pvGetRowData.TaxGroup = At(vRow, ucsFscPixTaxGroup)
    pvGetRowData.Quantity = At(vRow, ucsFscPixQuantity)
    pvGetRowData.Discount = At(vRow, ucsFscPixDiscount)
    pvGetRowData.UnitOfMeasure = At(vRow, ucsFscPixUnitOfMeasure)
    pvGetRowData.DepartmentNo = At(vRow, ucsFscPixDepartmentNo)
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function pvGetTaxGroup(sTaxGroup As String) As Long
    pvGetTaxGroup = Asc(sTaxGroup & " ") - 191 '--- 191 = Asc("А") - 1
    If pvGetTaxGroup < MIN_TAX_GROUP Or pvGetTaxGroup > MAX_TAX_GROUP Then
        pvGetTaxGroup = C_Lng(sTaxGroup)
        If pvGetTaxGroup < MIN_TAX_GROUP Or pvGetTaxGroup > MAX_TAX_GROUP Then
            pvGetTaxGroup = DEF_TAX_GROUP
        End If
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

Private Function pvCommonProlog(sRequest As String, oJson As Object, sLastErr As String, oFP As IDeviceProtocol, oRetVal As Object) As Boolean
    Dim sDeviceString   As String
    
    JsonItem(oRetVal, "Ok") = True
    JsonItem(oRetVal, "ErrorText") = vbNullString
    If Not JsonParse(sRequest, oJson, Error:=sLastErr) Or oJson Is Nothing Then
        sLastErr = Zn(sLastErr, pvInternal(ucsErrInvalidJsonRequest))
        GoTo QH
    End If
    sDeviceString = JsonItem(oJson, "DeviceString")
    Set oFP = pvInit(sDeviceString)
    If sDeviceString <> JsonItem(oJson, "DeviceString") Then
        JsonItem(oJson, "DeviceString") = sDeviceString
    End If
    '--- success
    pvCommonProlog = True
QH:
End Function

Private Function pvCommonEpilog(oFP As IDeviceProtocol, sLastErr As String, ByVal eLastNo As UcsFiscalErrorsEnum, oRetVal As Object, sResponse As String) As Boolean
    If LenB(sLastErr) <> 0 Then
        JsonItem(oRetVal, "Ok") = False
        JsonItem(oRetVal, "ErrorText") = sLastErr
    End If
    Select Case eLastNo
    Case ucsFerGeneralError
        JsonItem(oRetVal, "Ok") = False
    Case ucsFerInvalidPassword
        JsonItem(oRetVal, "Ok") = False
        JsonItem(oRetVal, "ErrorCode") = "InvalidPassword"
    Case ucsFerPasswordNotSet
        JsonItem(oRetVal, "Ok") = False
        JsonItem(oRetVal, "ErrorCode") = "PasswordNotSet"
    End Select
    If LenB(JsonItem(oRetVal, "ErrorText")) = 0 And Not oFP Is Nothing Then
        JsonItem(oRetVal, "ErrorText") = Zn(oFP.GetLastError(), Empty)
    End If
    sResponse = JsonDump(oRetVal, Minimize:=True)
    If Not oFP Is Nothing Then
        If oFP.IsConnected Then
            oFP.Disconnect
        End If
        m_sCommandLog = oFP.GetCommandLog()
    End If
    '--- success
    pvCommonEpilog = True
End Function

Private Function pvCommonErrHandler(sFunction As String, oFP As IDeviceProtocol, sLastErr As String, eLastNo As UcsFiscalErrorsEnum) As Boolean
    If Not oFP Is Nothing Then
        sLastErr = Zn(oFP.GetLastError(eLastNo), Err.Description)
        DebugLog sFunction, "sLastErr=" & sLastErr & ", eLastNo=" & eLastNo
    Else
        sLastErr = Err.Description
        DebugLog sFunction, "sLastErr=" & sLastErr
    End If
    '--- success
    pvCommonErrHandler = True
End Function

Private Sub pvCommonIncludesMiddleware(oFP As IDeviceProtocol, oJson As Object, oRetVal As Object)
    Dim bIncludeAll     As Boolean
    Dim lIdx            As Long
    Dim dLastDateTime   As Date
    Dim oItem           As Object
    Dim sName           As String
    Dim lTaxGroup       As Long
    
    bIncludeAll = JsonBoolItem(oJson, "IncludeAll")
    If JsonBoolItem(oJson, "IncludeTaxNo") Or bIncludeAll Then
        JsonItem(oRetVal, "TaxNo") = Trim$(oFP.GetTaxNo())
        JsonItem(oRetVal, "TaxCaption") = Replace(Trim$(oFP.GetTaxCaption()), ":", vbNullString)
    End If
    If JsonBoolItem(oJson, "IncludeHeaders") Or bIncludeAll Then
        For lIdx = 0 To 5
            JsonItem(oRetVal, "Headers/" & lIdx) = Trim$(oFP.GetHeaderText(lIdx + 1))
        Next
        For lIdx = 5 To 0 Step -1
            If LenB(JsonItem(oRetVal, "Headers/" & lIdx)) <> 0 Then
                Exit For
            End If
            JsonItem(oRetVal, "Headers/" & lIdx) = Empty
        Next
    End If
    If JsonBoolItem(oJson, "IncludeFooters") Or bIncludeAll Then
        For lIdx = 0 To 1
            JsonItem(oRetVal, "Footers/" & lIdx) = Trim$(oFP.GetFooterText(lIdx + 1))
        Next
        For lIdx = 1 To 0 Step -1
            If LenB(JsonItem(oRetVal, "Footers/" & lIdx)) <> 0 Then
                Exit For
            End If
            JsonItem(oRetVal, "Footers/" & lIdx) = Empty
        Next
    End If
    If JsonBoolItem(oJson, "IncludeReceiptNo") Or bIncludeAll Then
        JsonItem(oRetVal, "LastReceiptNo") = StripZeros(oFP.GetLastReceiptNo())
        If LenB(JsonItem(oRetVal, "LastReceiptNo")) = 0 Then
            dLastDateTime = 0
            JsonItem(oRetVal, "LastReceiptNo") = StripZeros(oFP.GetLastQRCodeInfo(dLastDateTime))
            JsonItem(oRetVal, "LastReceiptDateTime") = dLastDateTime
        Else
            JsonItem(oRetVal, "LastReceiptDateTime") = oFP.GetClock()
        End If
    End If
    If JsonBoolItem(oJson, "IncludePaymentNames") Or bIncludeAll Then
        For lIdx = 1 To 12
            JsonItem(oRetVal, "PaymentNames/-1") = Trim$(oFP.GetPaymentName(lIdx))
        Next
    End If
    If JsonBoolItem(oJson, "IncludeTaxRates") Or bIncludeAll Then
        JsonItem(oRetVal, "TaxRates/*") = oFP.GetTaxRates()
    End If
    If JsonBoolItem(oJson, "IncludeOperators") Then
        For lIdx = 0 To 100
            sName = oFP.GetOperatorName(C_Str(lIdx))
            If LenB(sName) = 0 Then
                If lIdx > 0 Then
                    Exit For
                End If
            Else
                Set oItem = Nothing
                JsonItem(oItem, "Code") = C_Str(lIdx)
                JsonItem(oItem, "Name") = sName
                JsonItem(oItem, "DefaultPassword") = oFP.GetDefaultPassword(C_Str(lIdx))
                JsonItem(oRetVal, "Operators/-1") = oItem
            End If
        Next
    End If
    If JsonBoolItem(oJson, "IncludeDepartments") Then
        For lIdx = 0 To 100
            sName = oFP.GetDepartmentName(C_Str(lIdx), lTaxGroup)
            If LenB(sName) = 0 Then
                If lIdx > 0 Then
                    Exit For
                End If
            Else
                Set oItem = Nothing
                JsonItem(oItem, "Code") = C_Str(lIdx)
                JsonItem(oItem, "Name") = sName
                JsonItem(oItem, "TaxGroup") = lTaxGroup
                JsonItem(oRetVal, "Departments/-1") = oItem
            End If
        Next
    End If
End Sub

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Initialize()
    LocalizedText(ucsFscLciInternalErrors) = GetConfigValue(MODULE_NAME, CFG_INTERNAL_ERRORS, STR_INTERNAL)
    LocalizedText(ucsFscLciReceiptTexts) = GetConfigValue(MODULE_NAME, CFG_RECEIPT_TEXTS, STR_TEXTS)
End Sub

Private Sub Class_Terminate()
    mdGlobals.FlushDebugLog
End Sub
