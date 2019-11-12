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
' UcsFP20 (c) 2008-2019 by Unicontsoft
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
' Public enums
'=========================================================================

Public Enum UcsFiscalPrintReportTypeEnum
    ucsFscRptDaily = 1
    ucsFscRptNumber
    ucsFscRptDate
    ucsFscRptOperator
End Enum

'=========================================================================
' Constants and member variables
'=========================================================================

'--- strings
Private Const STR_TEXTS                 As String = "Касиер"
Private Const STR_INTERNAL              As String = "Грешка при инициализация на протокол %1|Командата не се поддържа от протокола|Невалидна JSON заявка"
Private Const STR_DATECS_FP_INTERNAL    As String = "Време за достъп изтече в очакване на отговор|Невалидна дължина на отговора: %1|Невалидна сума за проверка: %1|Липсва символ за край на отговора: %1|Липсва символ за начало на отговора: %1|Няма започната бележка| или невалидна парола на оператор|Грешка на ред %1: %2|Невъзможно отказване на предишна транзакция|Невалиден тип справка|Грешка %1|Памет %1|Неподдържан вид плащане %1"
Private Const STR_DATECS_FP_STATUSES    As String = "Получените данни имат синктактична грешка|Кодът на получената команда е невалиден|Не е сверен часовника|3|Механизмът на печатащото устройство има неизправност|5|Отворен е капакът на принтера|7|При изпълнение на командата се е получило препълване на някои полета от сумите|Изпълнението на командата не е позволено в текущия фискален режим|2|3|4|5|6|7|Свършила е хартията|Останала е малко хартия|Край на КЛЕН|Отворен е фискален бон|Близък край на КЛЕН|Отворен е служебен бон"
Private Const STR_DATECS_FP_ERRORS      As String = "Получените данни имат синктактична грешка|Кодът на получената команда е невалиден|2|3|Механизмът на печатащото устройство има неизправност|5|6|7|0|Изпълнението на командата не е позволено в текущия фискален режим|2|3|4|5|6|7|Свършила е хартията"
Private Const STR_DATECS_FP_MEMORY      As String = "Грешка при запис във фискалната памет|Зададен е ЕИК|Зададени са индивидуален номер на ФУ и номер на ФП|Има място за по-малко от 50 записа във ФП|ФП е пълна|5|ФП не е намерена или е повредена|7|0|ФП е форматирана|2|ФУ е във фискален режим|Зададени са поне веднъж данъчните ставки"
Private Const STR_DATECS_FP_TEXTS       As String = "В БРОЙ|С КАРТА|ПО БАНКА|С ЧЕК|НАДБАВКА %1|ОТСТЪПКА %1|МЕЖДИННА СУМА|Продажби %1|ЕИК"
Private Const STR_ISL_FP_INTERNAL       As String = "Време за достъп изтече в очакване на отговор|Невалидна дължина на отговора: %1|Невалидна сума за проверка: %1|Липсва символ за край на отговора: %1|Липсва символ за начало на отговора: %1|Няма започната бележка| или невалидна парола на оператор|Грешка на ред %1: %2|Невъзможно отказване на предишна транзакция|Невалиден тип справка|Грешка %1|Памет %1|Неподдържан вид плащане %1"
Private Const STR_ISL_FP_STATUSES_A     As String = "Получените данни имат синтактична грешка|Кодът на получената команда е невалиден|Часовникът не е установен|Не е свързан клиентски дисплей|Механизмът на печатащото устройство има неизправност|5|Отворен е капакът на принтера|7|При изпълнение на командата се е получило препълване на някои полета от сумите|Изпълнението на командата не е позволено в текущия фискален режим|Извършено е зануляване на оперативната памет|Слаба батерия (Часовникът за реално време е в състояние RESET)|Отворен сторно бон|Отворен е служебен бон за печат на завъртян на 90 градуса текст|Вграденият данъчен терминал не отговаря|7|Свършила е хартията|Останала е малко хартия|Край на КЛЕН (по-малко от 1 MB от КЛЕН свободни)|Отворен е фискален бон|Близък край на КЛЕН (по-малко от 10 MB от КЛЕН свободни)|Отворен е служебен бон|Много близък край на КЛЕН (допускат се само определени бонове)"
Private Const STR_ISL_FP_ERRORS_A       As String = "Получените данни имат синтактична грешка|Кодът на получената команда е невалиден|2|3|Механизмът на печатащото устройство има неизправност|5|6|7|0|Изпълнението на командата не е позволено в текущия фискален режим|Извършено е зануляване на оперативната памет|Слаба батерия (Часовникът за реално време е в състояние RESET)|4|5|6|7|Свършила е хартията"
Private Const STR_ISL_FP_MEMORY_A       As String = "Има грешка при запис във фискалната памет|Зададен е ЕИК по БУЛСТАТ|Зададени са индивидуален номер на принтера и номер на фискалната памет|Има място за по-малко от 50 записа във ФП|Фискалната памет е пълна|5|Печатащата глава е прегряла|7|Фискалната памет е установена в режим READONLY (заключена)|Фискалната памет е форматирана|Последният запис във фискалната памет не е успешен|Принтерът е във фискален режим|Зададени са поне веднъж данъчните ставки|Грешка при четене от фискалната памет"
Private Const STR_ISL_FP_STATUSES_B     As String = "Получените данни имат синктактична грешка|Кодът на получената команда е невалиден|Не е сверен часовника|Не е свързан клиентски дисплей|4|5|6|7|При изпълнение на командата се е получило препълване на някои полета от сумите|Изпълнението на командата не е позволено в текущия фискален режим|2|3|4|Има неизпратени документи за повече от настроеното време за предупреждение|Вграденият данъчен терминал не отговаря|7|Свършила е хартията|1|Край на КЛЕН (по-малко от 1 MB от КЛЕН свободни)|Отворен е фискален бон|Близък край на КЛЕН (по-малко от 10 MB от КЛЕН свободни)|Отворен е служебен бон"
Private Const STR_ISL_FP_ERRORS_B       As String = "Получените данни имат синктактична грешка|Кодът на получената команда е невалиден|2|3|4|5|6|7|0|Изпълнението на командата не е позволено в текущия фискален режим|2|3|4|5|6|7|Свършила е хартията"
Private Const STR_ISL_FP_MEMORY_B       As String = "Грешка при запис във ФП|Зададен е ЕИК|Зададени са индивидуален номер на ФУ и номер на ФП|Има място за по-малко от 50 записа във ФП|ФП е пълна|5|6|7|0|ФП е форматирана|2|ФУ е във фискален режим|Зададени са поне веднъж данъчните ставки"
Private Const STR_ISL_FP_DIP_SWITCHES   As String = "Автоматично центриране на header и footer|Предварителен header|Sw1.3|Sw1.4|Нулиране на паметта|Прозрачен дисплей|Без данни на дисплея|7"
Private Const STR_ISL_FP_TEXTS          As String = "В БРОЙ|С КАРТА|ПО БАНКА|С ЧЕК|НАДБАВКА %1|ОТСТЪПКА %1|МЕЖДИННА СУМА|Продажби %1|ЕИК"
Private Const STR_TREMOL_FP_INTERNAL    As String = "Време за достъп изтече в очакване на отговор|Невалиден формат на съобщение или сума за проверка (NAK)|Невалидна дължина на отговора: %1|Невалидна сума за проверка: %1|Липсва символ за начало на отговора: %1|Няма започната бележка|Грешка на ред %1: %2|Невалиден тип справка|Устройството е недостъпно|Устройството е заето| или невъзможно установяване парола на оператор|Неподдържан вид плащане %1"
Private Const STR_TREMOL_FP_STATUSES    As String = "Неизвестна грешка|Невалидна команда|Непозволена команда|Непозволена поради ненулев отчет|Синтактична грешка|Препълване на входните регистри|Нулев входен регистър|Липсва транзакция която да се войдира|Недостатъчна налична сума|Конфликт в данъчните групи"
Private Const STR_TREMOL_FP_ERRORS      As String = "Неизвестна грешка|Няма хартия|Препълване на дневните регистри|Несверен/грешен часовник|Отворен фискален бон|Сметка с остатък за плащане|Отворен нефискален бон|Сметка с приключено плащане|Фискална памет само за четене|Грешна парола или непозволена команда|Липсващ външен дисплей|24 часа без дневен отчет|Прегрят принтер|Спад на напрежение във фискален бон|Препълване в електронната контролна лента|Недостатъчни условия"
Private Const STR_TREMOL_FP_TEXTS       As String = "В брой|С карта|По банка|Надбавка%1|Отстъпка%1|Междинна сума|Продажби %1|ЕИК;ЕГН;ЛНЧ;Служебен No"
Private Const STR_DAISY_FP_EXT_ERRORS   As String = "1|Въпросната операция ще доведе до препълване|3|Нямате право на повече продажби в този бон|4|Нямате право на повече плащания в този бон|5|Опит за извършване на нулева транзакция|6|Опит за извършване на продажба, след като е започнато плащане|7|Нямате право на такава операция|8|Забранена за продажби дан. група|9|Грешен диапазон за номера на фактури|11|Въвеждане на повече от една десетична точка|12|Въвеждане на повече от един символ '+' или '-'|13|Символът '+' или '-' не е в първа позиция|14|Недопустим символ, напр. баркод, който съдържа не само цифри|15|Повече от допустимия брой знаци след десетичната точка|16|Въведени са повече от разрешения брой символи|20|В дадената ситуация не се обслужва избраната от Вас команда от PC|21|Стойността е извън допустимите граници|22|Виж системен параметър 8|23|Опит за ""дълбок"" войд след отстъпка/надбавка в/у междинна сума|24|Опит за ""дълбок"" войд на несъществуваща транзакция|" & _
                                                    "25|Опит за извършване на плащане, без да има продажби|26|Опит за продажба на артикул с количество, надвишаващо запаса му|41|Некоректен баркод (грешна контролна сума)|42|Опит за продажба с нулев баркод|43|Опит за програмиране с тегловен баркод|44|Опит за продажба с непрограмиран баркод|45|Опит за програмиране на вече съществуващ баркод|66|Некоректна парола|71|!!! Некоректни данни във ФП !!!!|72|!!! Грешка при запис във ФП !!!!|76|Необходима е информация от сървъра на НАП|90|Не е нулиран периодичният отчет|91|Не е нулиран дневният финансов отчет|92|Не е нулиран отчетът по оператори|93|Не е нулиран отчетът по артикули|94|Не може да се препрограмира това поле|81|Дневният финансов отчет е препълнен|83|Отчетът по оператори е препълнен|84|Отчетът по артикули е препълнен|84|Периодичният отчет е препълнен|88|КЛЕН е препълнен|102|Няма комуникация между ФУ и ДТ|104|Некоректна комуникация между ФУ и ДТ|" & _
                                                    "110|Подменена SIM карта|111|Грешка при комуникация между ДТ и сървъра на НАП|113|Сървърът на НАП не приема подадените му данни|117|Неуспешен опит за свързване на ДТ с мрежата на мобилния оператор|118|Операцията е забранена|119|Грешна въведена стойност|120|Невъведена стойност"
Private Const STR_ESP_POS_INTERNAL      As String = "Няма започната бележка|Грешка на ред %1: %2"
Private Const STR_ESP_POS_TEXTS         As String = "В БРОЙ|С ДЕБИТНА КАРТА|С ЧЕК|С КРЕДИТНА КАРТА|НАДБАВКА %1|ОТСТЪПКА %1|СУМА|МЕЖДИННА СУМА|ФАКТУРА No %1|Продажби %1|ЕДИНИЧНА ЦЕНА|КОЛИЧЕСТВО|СУМА|ВСИЧКО ГРУПА %1|ДДС %1=%2|НЕТО СТОЙНОСТ|ОБЩО|%1 АРТИКУЛА|1 АРТИКУЛ|РЕСТО|ПРОДАВАЧ: |ПОЛУЧАТЕЛ: |КУПУВАЧ: |ИДЕНТ. No: |ЗДДС No: |СИСТЕМЕН БОН|ПОРЪЧКА КЪМ КУХНЯ|Дата:     %1|Оператор: %1|Маса:     %1|ЕИК №: %1"
Private Const STR_CONNECTOR_ERRORS      As String = "Не е указано устройство|Грешка при отваряне: %1|Грешка при SetCommTimeouts: %1|Грешка при BuildCommDCB: %1|Грешка при SetCommState: %1|Грешка при WriteFile: %1|Време за достъп изтече в очакване на отговор|Грешка при ReadFile: %1|Грешка при WaitCommEvent: %1"
Private Const ERR_UNKNOWN_PROTOCOL      As String = "Неизвестен протокол: %1"
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
    OutputDebugLog MODULE_NAME, sFunction & "(" & Erl & ")", "Run-time error: " & Err.Description
End Sub

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
    Dim lIdx            As Long
    Dim oItem           As Object
    Dim oRetVal         As Object
    
    On Error GoTo EH
    JsonItem(oRetVal, "Ok") = True
    vPorts = EnumSerialPorts
    For Each vElem In Array(pvInit("Protocol=" & STR_PROTOCOL_TREMOL, SkipLastError:=True), _
                            pvInit("Protocol=" & STR_PROTOCOL_ESCPOS, SkipLastError:=True), _
                            pvInit("Protocol=" & STR_PROTOCOL_DATECS_FP, SkipLastError:=True), _
                            pvInit("Protocol=" & STR_PROTOCOL_DATECS_X, SkipLastError:=True))
        Set oFP = vElem
        vPorts = oFP.AutodetectDevices(vPorts)
    Next
    For lIdx = 0 To UBound(vPorts)
        Set oItem = Nothing
        If IsArray(vPorts(lIdx)) Then
            JsonItem(oItem, "Port") = vPorts(lIdx)(0)
            JsonItem(oItem, "Speed") = vPorts(lIdx)(1)
            JsonItem(oItem, "Protocol") = vPorts(lIdx)(2)
            JsonItem(oItem, "Model") = vPorts(lIdx)(3)
            JsonItem(oItem, "Firmware") = vPorts(lIdx)(4)
        Else
            JsonItem(oItem, "Port") = vPorts(lIdx)
        End If
        JsonItem(oRetVal, "SerialPorts/-1") = oItem
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
    JsonItem(oRetVal, "DeviceSerialNo") = Trim$(oFP.GetDeviceSerialNo())
    JsonItem(oRetVal, "FiscalMemoryNo") = Trim$(oFP.GetFiscalMemoryNo())
    JsonItem(oRetVal, "DeviceProtocol") = Trim$(oFP.GetDeviceProtocol())
    JsonItem(oRetVal, "DeviceModel") = Trim$(oFP.GetDeviceModel())
    JsonItem(oRetVal, "FirmwareVersion") = Trim$(oFP.GetDeviceFirmware())
    JsonItem(oRetVal, "CharsPerLine") = oFP.GetCharsPerLine()
    JsonItem(oRetVal, "CommentTextMaxLength") = oFP.GetCommentTextMaxLength()
    JsonItem(oRetVal, "ItemNameMaxLength") = oFP.GetItemNameMaxLength()
    If LenB(C_Str(JsonItem(oJson, "Operator/Code"))) <> 0 Then
        JsonItem(oRetVal, "Operator/Code") = JsonItem(oJson, "Operator/Code")
        JsonItem(oRetVal, "Operator/Name") = Trim$(oFP.GetOperatorName(JsonItem(oJson, "Operator/Code")))
        JsonItem(oRetVal, "Operator/DefaultPassword") = oFP.GetDefaultPassword(JsonItem(oJson, "Operator/Code"))
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
    Dim sResumeToken   As String
    Dim sPrevReceiptNo  As String
    Dim dLastDateTime   As Date
    
    On Error GoTo EH
    If Not pvCommonProlog(sRequest, oJson, sLastErr, oFP, oRetVal) Then
        GoTo QH
    End If
    sResumeToken = JsonItem(oJson, "ResumeToken")
    If LenB(sResumeToken) = 0 Then
        If Not oFP.CancelReceipt() Then
            sLastErr = oFP.GetLastError(eLastNo)
            DebugLog FUNC_NAME, "Not oFP.CancelReceipt(), sLastErr=" & sLastErr & ", eLastNo=" & eLastNo
            GoTo QH
        End If
    End If
    sPrevReceiptNo = JsonItem(oJson, "PrevReceiptNo")
    If LenB(sPrevReceiptNo) <> 0 Then
        JsonItem(oRetVal, "PrevReceiptNo") = StripZeros(oFP.GetLastReceiptNo())
        '--- check if receipt already printed
        If LenB(Trim$(sPrevReceiptNo)) <> 0 Then
            If Pad(sPrevReceiptNo, -Len(JsonItem(oRetVal, "PrevReceiptNo"))) <> JsonItem(oRetVal, "PrevReceiptNo") And C_Lng(JsonItem(oRetVal, "PrevReceiptCount")) <> 0 Then
                If C_Str(JsonItem(oRetVal, "PrevReceiptCount")) <> At(oFP.GetTotalsStats(), 0) Then
                    JsonItem(oRetVal, "PrevReceiptNo") = Empty
                    GoTo ReturnReceiptInfo
                End If
            End If
        End If
    End If
    If IsObject(JsonItem(oJson, "Rows")) Then
        If Not oFP.StartReceipt(ToEnumValue(JsonItem(oJson, "ReceiptType"), "|Sale|Reversal|Invoice|CreditNote|OrderList"), _
                JsonItem(oJson, "Operator/Code"), _
                Zn(JsonItem(oJson, "Operator/Name"), pvText(ucsTxtCashier)), _
                JsonItem(oJson, "Operator/Password"), _
                UniqueSaleNo:=JsonItem(oJson, "UniqueSaleNo"), _
                TableNo:=JsonItem(oJson, "TableNo"), _
                DisablePrinting:=JsonBoolItem(oJson, "DisablePrinting"), _
                InvDocNo:=JsonItem(oJson, "Invoice/DocNo"), _
                InvCgTaxNo:=JsonItem(oJson, "Invoice/CgTaxNo"), _
                InvCgTaxNoType:=ToEnumValue(JsonItem(oJson, "Invoice/CgTaxNoType"), "EIC|CitizenNo|ForeignerNo|OfficialNo"), _
                InvCgVatNo:=JsonItem(oJson, "Invoice/CgVatNo"), _
                InvCgName:=JsonItem(oJson, "Invoice/CgName"), _
                InvCgAddress:=JsonItem(oJson, "Invoice/CgAddress"), _
                InvCgPrsReceive:=JsonItem(oJson, "Invoice/CgPrsReceive"), _
                OwnData:=JsonItem(oJson, "OwnData"), _
                RevType:=ToEnumValue(JsonItem(oJson, "Reversal/ReversalType"), "OperatorError|Refund|TaxBaseReduction"), _
                RevReceiptNo:=JsonItem(oJson, "Reversal/ReceiptNo"), _
                RevReceiptDate:=C_Date(JsonItem(oJson, "Reversal/ReceiptDateTime")), _
                RevFiscalMemoryNo:=JsonItem(oJson, "Reversal/FiscalMemoryNo"), _
                RevInvoiceNo:=JsonItem(oJson, "Reversal/InvoiceNo"), _
                RevReason:=JsonItem(oJson, "Reversal/Reason")) Then
            sLastErr = oFP.GetLastError(eLastNo)
            DebugLog FUNC_NAME, "Not oFP.StartReceipt(), sLastErr=" & sLastErr & ", eLastNo=" & eLastNo
            GoTo QH
        End If
        For Each vElem In JsonKeys(oJson, "Rows")
            With pvGetRowData(JsonItem(oJson, "Rows/" & vElem))
                Select Case .RowType
                Case ucsRwtSell
                    oFP.AddPLU .ItemName, C_Dbl(.Price), IIf(LenB(.Quantity) <> 0, C_Dbl(.Quantity), 1), pvGetTaxGroup(.TaxGroup), .UnitOfMeasure, C_Lng(.DepartmentNo)
                    If Abs(C_Dbl(.Discount)) > DBL_EPSILON Then
                        oFP.AddDiscount ucsFscDscPlu, -C_Dbl(.Discount)
                    End If
                Case ucsRwtText
                    oFP.AddLine .Text
                Case ucsRwtBarcode
                    oFP.AddBarcode ToEnumValue(.BarcodeType, "|Ean8|Ean13|Code128|QRcode"), .Text, C_Lng(.Height)
                Case ucsRwtPayment
                    lIdx = ToEnumValue(.PaymentType, "|Cash|Card|Bank|Cheque|Custom1|Custom2|Custom3|Custom4")
                    If lIdx < 0 Then
                        lIdx = LimitLong(lIdx, -4, -1)
                    Else
                        lIdx = LimitLong(lIdx, 1, 8)
                    End If
                    oFP.AddPayment lIdx, Zn(.PaymentName, Trim$(oFP.GetPaymentName(lIdx))), C_Dbl(.Amount)
                End Select
            End With
        Next
        If Not oFP.EndReceipt(sResumeToken) Then
            sLastErr = oFP.GetLastError(eLastNo)
            DebugLog FUNC_NAME, "Not oFP.EndReceipt(), sLastErr=" & sLastErr & ", eLastNo=" & eLastNo
            GoTo QH
        End If
        oFP.OpenDrawer
    End If
    If C_Bool(JsonItem(oJson, "PrintDuplicate")) Then
        If Not oFP.CopyLastReceipt(JsonItem(oJson, "Invoice/DocNo")) Then
            sLastErr = oFP.GetLastError(eLastNo)
            DebugLog FUNC_NAME, "Not oFP.CopyLastReceipt(), sLastErr=" & sLastErr & ", eLastNo=" & eLastNo
            GoTo QH
        End If
    End If
ReturnReceiptInfo:
    JsonItem(oRetVal, "ReceiptNo") = StripZeros(oFP.GetLastQRCodeInfo(dLastDateTime))
    JsonItem(oRetVal, "ReceiptDateTime") = dLastDateTime
    If LenB(JsonItem(oRetVal, "ReceiptNo")) = 0 Then
        JsonItem(oRetVal, "ReceiptNo") = StripZeros(oFP.GetLastReceiptNo())
        JsonItem(oRetVal, "ReceiptDateTime") = oFP.GetClock()
    End If
    JsonItem(oRetVal, "DeviceSerialNo") = Trim$(oFP.GetDeviceSerialNo())
    JsonItem(oRetVal, "FiscalMemoryNo") = Trim$(oFP.GetFiscalMemoryNo())
    pvCommonIncludesMiddleware oFP, oJson, oRetVal
QH:
    PrintReceipt = pvCommonEpilog(oFP, sLastErr, eLastNo, oRetVal, sResponse)
    Exit Function
EH:
    If pvCommonErrHandler(FUNC_NAME, oFP, sLastErr, eLastNo) Then
        If Not oFP Is Nothing Then
            JsonItem(oRetVal, "ResumeToken") = Zn(oFP.GetResumeToken(), Empty)
        End If
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
    
    On Error GoTo EH
    If Not pvCommonProlog(sRequest, oJson, sLastErr, oFP, oRetVal) Then
        GoTo QH
    End If
    Select Case ToEnumValue(JsonItem(oJson, "ReportType"), "|Daily|Number|Date|Operator")
    Case ucsFscRptDaily
        If C_Bool(JsonItem(oJson, "IsItems")) And C_Bool(JsonItem(oJson, "IsDepartments")) Then
            If C_Bool(JsonItem(oJson, "IsClear")) Then
                bResult = oFP.RunZReport(ucsFscRstDailyByItemsAndDepartment)
            Else
                bResult = oFP.RunXReport(ucsFscRstDailyByItemsAndDepartment)
            End If
        ElseIf C_Bool(JsonItem(oJson, "IsItems")) Then
            If C_Bool(JsonItem(oJson, "IsClear")) Then
                bResult = oFP.RunZReport(ucsFscRstDailyByItems)
            Else
                bResult = oFP.RunXReport(ucsFscRstDailyByItems)
            End If
        ElseIf C_Bool(JsonItem(oJson, "IsDepartments")) Then
            If C_Bool(JsonItem(oJson, "IsClear")) Then
                bResult = oFP.RunZReport(ucsFscRstDailyByDepartment)
            Else
                bResult = oFP.RunXReport(ucsFscRstDailyByDepartment)
            End If
        Else
            If C_Bool(JsonItem(oJson, "IsClear")) Then
                bResult = oFP.RunZReport(ucsFscRstDaily)
            Else
                bResult = oFP.RunXReport(ucsFscRstDaily)
            End If
        End If
    Case ucsFscRptNumber
        '--- ToDo: impl
    Case ucsFscRptDate
        If C_Date(JsonItem(oJson, "FromDate")) <> 0 And C_Date(JsonItem(oJson, "ToDate")) <> 0 Then
            bResult = oFP.RunPeriodReport(IIf(C_Bool(JsonItem(oJson, "IsDetailed")), ucsFscRstPeriodDetailed, ucsFscRstPeriodShort), C_Date(JsonItem(oJson, "FromDate")), C_Date(JsonItem(oJson, "ToDate")))
        End If
    Case ucsFscRptOperator
        '--- ToDo: impl
    End Select
    If Not bResult Then
        sLastErr = oFP.GetLastError(eLastNo)
        DebugLog FUNC_NAME, "Not bResult, sLastErr=" & sLastErr & ", eLastNo=" & eLastNo
        GoTo QH
    End If
    JsonItem(oRetVal, "ReceiptNo") = StripZeros(oFP.GetLastReceiptNo())
    JsonItem(oRetVal, "ReceiptDateTime") = oFP.GetClock()
    JsonItem(oRetVal, "DeviceSerialNo") = Trim$(oFP.GetDeviceSerialNo())
    JsonItem(oRetVal, "FiscalMemoryNo") = Trim$(oFP.GetFiscalMemoryNo())
    pvCommonIncludesMiddleware oFP, oJson, oRetVal
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
    If Not oFP.OpenDrawer() Then
        sLastErr = oFP.GetLastError(eLastNo)
        DebugLog FUNC_NAME, "Not oFP.OpenDrawer(), sLastErr=" & sLastErr & ", eLastNo=" & eLastNo
        GoTo QH
    End If
    pvCommonIncludesMiddleware oFP, oJson, oRetVal
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
    dDateTime = C_Date(JsonItem(oJson, "DeviceDateTime"))
    lTolerance = C_Lng(JsonItem(oJson, "AdjustTolerance"))
    dCurrent = oFP.GetClock()
    If dDateTime = 0 And lTolerance = 0 Then
        GoTo SkipSetClock
    End If
    If dDateTime = 0 Then
        dDateTime = Now
    End If
    If lTolerance = 0 Then
        bResult = oFP.SetClock(dDateTime)
    ElseIf Abs(DateDiff("s", dCurrent, dDateTime)) >= lTolerance Then
'        bResult = oFP.SetClock(dDateTime)
    Else
        GoTo SkipSetClock
    End If
    If Not bResult Then
        sLastErr = oFP.GetLastError(eLastNo)
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
    pvCommonIncludesMiddleware oFP, oJson, oRetVal
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
    
    On Error GoTo EH
    If Not pvCommonProlog(sRequest, oJson, sLastErr, oFP, oRetVal) Then
        GoTo QH
    End If
    bIncludeAll = C_Bool(JsonItem(oJson, "IncludeAllTotals"))
    If C_Bool(JsonItem(oJson, "IncludeStats")) Or bIncludeAll Then
        vTotals = oFP.GetTotalsStats()
        If IsArray(vTotals) Then
            JsonItem(oRetVal, "NumReceipts") = vTotals(0)
            JsonItem(oRetVal, "LastZReportDateTime") = vTotals(1)
        End If
    End If
    If C_Bool(JsonItem(oJson, "IncludeServiceDeposit")) Or bIncludeAll Then
        vTotals = oFP.PrintServiceDeposit(JsonItem(oJson, "Operator/Code"), JsonItem(oJson, "Operator/Password"), 0)
        If IsArray(vTotals) Then
            JsonItem(oRetVal, "TotalAvailable") = vTotals(0)
            JsonItem(oRetVal, "TotalDeposits") = vTotals(1)
            JsonItem(oRetVal, "TotalWithdraws") = vTotals(2)
            JsonItem(oRetVal, "TotalReversal") = vTotals(3)
        End If
    End If
    If C_Bool(JsonItem(oJson, "IncludeTaxGroups")) Or bIncludeAll Then
        vTotals = oFP.GetTotalsByTaxGroups()
        If IsArray(vTotals) Then
            For lIdx = 0 To UBound(vTotals)
                If IsArray(vTotals(lIdx)) Then
                    Set oItem = Nothing
                    JsonItem(oItem, "TaxGroup") = lIdx + 1
                    If vTotals(lIdx)(0) >= 0 And vTotals(lIdx)(0) < 100 Then
                        JsonItem(oItem, "VatRate") = vTotals(lIdx)(0)
                    End If
                    JsonItem(oItem, "Amount") = vTotals(lIdx)(1)
                    JsonItem(oItem, "Reversal") = vTotals(lIdx)(2)
                    JsonItem(oRetVal, "TotalsByTaxGroups/-1") = oItem
                End If
            Next
        End If
    End If
    If C_Bool(JsonItem(oJson, "IncludePaymentTypes")) Or bIncludeAll Then
        vTotals = oFP.GetTotalsByPaymentTypes()
        If IsArray(vTotals) Then
            For lIdx = 0 To UBound(vTotals)
                If IsArray(vTotals(lIdx)) Then
                    Set oItem = Nothing
                    JsonItem(oItem, "PaymentType") = lIdx + 1
                    JsonItem(oItem, "PaymentName") = vTotals(lIdx)(0)
                    JsonItem(oItem, "Amount") = vTotals(lIdx)(1)
                    JsonItem(oItem, "Reversal") = vTotals(lIdx)(2)
                    JsonItem(oRetVal, "TotalsByPaymentTypes/-1") = oItem
                End If
            Next
        End If
    End If
    pvCommonIncludesMiddleware oFP, oJson, oRetVal
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
    
    On Error GoTo EH
    If Not pvCommonProlog(sRequest, oJson, sLastErr, oFP, oRetVal) Then
        GoTo QH
    End If
    vTotals = oFP.PrintServiceDeposit(JsonItem(oJson, "Operator/Code"), JsonItem(oJson, "Operator/Password"), C_Dbl(JsonItem(oJson, "Amount")))
    If Not IsArray(vTotals) Then
        sLastErr = pvInternal(ucsErrNotImplemented)
        GoTo QH
    ElseIf UBound(vTotals) < 0 Then
        sLastErr = oFP.GetLastError(eLastNo)
        DebugLog FUNC_NAME, "UBound(vTotals) < 0, sLastErr=" & sLastErr & ", eLastNo=" & eLastNo
        GoTo QH
    End If
    If Abs(C_Dbl(JsonItem(oJson, "Amount"))) > DBL_EPSILON Then
        JsonItem(oRetVal, "ReceiptNo") = StripZeros(oFP.GetLastReceiptNo())
        JsonItem(oRetVal, "ReceiptDateTime") = oFP.GetClock()
        JsonItem(oRetVal, "DeviceSerialNo") = Trim$(oFP.GetDeviceSerialNo())
        JsonItem(oRetVal, "FiscalMemoryNo") = Trim$(oFP.GetFiscalMemoryNo())
    End If
    JsonItem(oRetVal, "TotalAvailable") = vTotals(0)
    JsonItem(oRetVal, "TotalDeposits") = vTotals(1)
    JsonItem(oRetVal, "TotalWithdraws") = vTotals(2)
    JsonItem(oRetVal, "TotalReversal") = vTotals(3)
    pvCommonIncludesMiddleware oFP, oJson, oRetVal
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
    Dim sStatus         As String
    
    On Error GoTo EH
    If Not pvCommonProlog(sRequest, oJson, sLastErr, oFP, oRetVal) Then
        GoTo QH
    End If
    If oFP.GetDeviceStatus(sStatus) Then
        JsonItem(oRetVal, "DeviceStatus") = sStatus
        JsonItem(oRetVal, "DeviceDateTime") = oFP.GetClock()
    Else
        sLastErr = Zn(oFP.GetLastError(eLastNo), STR_NONE)
    End If
    pvCommonIncludesMiddleware oFP, oJson, oRetVal
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

'= private ===============================================================

Private Function pvInit(DeviceString As String, Optional ByVal SkipLastError As Boolean) As IDeviceProtocol
    Dim oOptions        As Object
    Dim sProtocol       As String
    Dim lConfigIdx      As Long
    
    Set oOptions = ParseDeviceString(DeviceString)
    sProtocol = UCase$(JsonItem(oOptions, "Protocol"))
    '--- figure out model
    Select Case sProtocol
    Case STR_PROTOCOL_DATECS_X
        Set pvInit = New cDatecsProtocol
        pvInit.SetLocalizedText ucsFscLciInternalErrors, GetConfigValue(sProtocol, CFG_INTERNAL_ERRORS, STR_DATECS_FP_INTERNAL)
        pvInit.SetLocalizedText ucsFscLciPrinterStatuses, GetConfigValue(sProtocol, CFG_PRINTER_STATUSES, STR_DATECS_FP_STATUSES)
        pvInit.SetLocalizedText ucsFscLciPrinterErrors, GetConfigValue(sProtocol, CFG_PRINTER_ERRORS, STR_DATECS_FP_ERRORS)
        pvInit.SetLocalizedText ucsFscLciMemoryTexts, GetConfigValue(sProtocol, CFG_MEMORY_TEXTS, STR_DATECS_FP_MEMORY)
        pvInit.SetLocalizedText ucsFscLciReceiptTexts, GetConfigValue(sProtocol, CFG_RECEIPT_TEXTS, STR_DATECS_FP_TEXTS)
        pvInit.SetLocalizedText ucsFscLciConnectorErrors, GetConfigValue(sProtocol, CFG_CONNECTOR_ERRORS, STR_CONNECTOR_ERRORS)
    Case STR_PROTOCOL_DATECS_FP, STR_PROTOCOL_DAISY, STR_PROTOCOL_INCOTEX
        Set pvInit = New cIslProtocol
        pvInit.SetLocalizedText ucsFscLciInternalErrors, GetConfigValue(sProtocol, CFG_INTERNAL_ERRORS, STR_ISL_FP_INTERNAL)
        pvInit.SetLocalizedText ucsFscLciDipSwitchesTexts, GetConfigValue(sProtocol, CFG_DIP_SWITCHES_TEXTS, STR_ISL_FP_DIP_SWITCHES)
        pvInit.SetLocalizedText ucsFscLciReceiptTexts, GetConfigValue(sProtocol, CFG_RECEIPT_TEXTS, STR_ISL_FP_TEXTS)
        pvInit.SetLocalizedText ucsFscLciExtendedErrors, GetConfigValue(sProtocol, CFG_EXTENDED_ERRORS, STR_DAISY_FP_EXT_ERRORS)
        pvInit.SetLocalizedText ucsFscLciConnectorErrors, GetConfigValue(sProtocol, CFG_CONNECTOR_ERRORS, STR_CONNECTOR_ERRORS)
        lConfigIdx = 0 '--- 0 = Datecs models type A (FPs)
        pvInit.SetLocalizedText ucsFscLciPrinterStatuses, GetConfigValue(sProtocol, CFG_PRINTER_STATUSES, STR_ISL_FP_STATUSES_A), lConfigIdx
        pvInit.SetLocalizedText ucsFscLciPrinterErrors, GetConfigValue(sProtocol, CFG_PRINTER_ERRORS, STR_ISL_FP_ERRORS_A), lConfigIdx
        pvInit.SetLocalizedText ucsFscLciMemoryTexts, GetConfigValue(sProtocol, CFG_MEMORY_TEXTS, STR_ISL_FP_MEMORY_A), lConfigIdx
        lConfigIdx = 1 '--- 1 = Datecs models type B (ECRs)
        pvInit.SetLocalizedText ucsFscLciPrinterStatuses, GetConfigValue(sProtocol, CFG_PRINTER_STATUSES, STR_ISL_FP_STATUSES_B), lConfigIdx
        pvInit.SetLocalizedText ucsFscLciPrinterErrors, GetConfigValue(sProtocol, CFG_PRINTER_ERRORS, STR_ISL_FP_ERRORS_B), lConfigIdx
        pvInit.SetLocalizedText ucsFscLciMemoryTexts, GetConfigValue(sProtocol, CFG_MEMORY_TEXTS, STR_ISL_FP_MEMORY_B), lConfigIdx
    Case STR_PROTOCOL_TREMOL
        Set pvInit = New cTremolProtocol
        pvInit.SetLocalizedText ucsFscLciInternalErrors, GetConfigValue(sProtocol, CFG_INTERNAL_ERRORS, STR_TREMOL_FP_INTERNAL)
        pvInit.SetLocalizedText ucsFscLciPrinterStatuses, GetConfigValue(sProtocol, CFG_PRINTER_STATUSES, STR_TREMOL_FP_STATUSES)
        pvInit.SetLocalizedText ucsFscLciPrinterErrors, GetConfigValue(sProtocol, CFG_PRINTER_ERRORS, STR_TREMOL_FP_ERRORS)
        pvInit.SetLocalizedText ucsFscLciReceiptTexts, GetConfigValue(sProtocol, CFG_RECEIPT_TEXTS, STR_TREMOL_FP_TEXTS)
        pvInit.SetLocalizedText ucsFscLciConnectorErrors, GetConfigValue(sProtocol, CFG_CONNECTOR_ERRORS, STR_CONNECTOR_ERRORS)
    Case STR_PROTOCOL_ESCPOS
        Set pvInit = New cEscPosProtocol
        pvInit.SetLocalizedText ucsFscLciInternalErrors, GetConfigValue(sProtocol, CFG_INTERNAL_ERRORS, STR_ESP_POS_INTERNAL)
        pvInit.SetLocalizedText ucsFscLciReceiptTexts, GetConfigValue(sProtocol, CFG_RECEIPT_TEXTS, STR_ESP_POS_TEXTS)
        pvInit.SetLocalizedText ucsFscLciConnectorErrors, GetConfigValue(sProtocol, CFG_CONNECTOR_ERRORS, STR_CONNECTOR_ERRORS)
    Case Else
        Err.Raise vbObjectError, , Printf(ERR_UNKNOWN_PROTOCOL, sProtocol)
    End Select
    If UBound(JsonKeys(oOptions)) > 0 Then
        If Not pvInit.Init(DeviceString) Then
            Err.Raise vbObjectError, , Zn(pvInit.GetLastError(), Printf(pvInternal(ucsErrInitProtocol), Zn(sProtocol, STR_NONE)))
        End If
        If Not SkipLastError Then
            If LenB(pvInit.GetLastError()) Then
                Err.Raise vbObjectError, , pvInit.GetLastError()
            End If
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
    JsonItem(oRetVal, "Ok") = True
    If Not JsonParse(sRequest, oJson, Error:=sLastErr) Or oJson Is Nothing Then
        sLastErr = Zn(sLastErr, pvInternal(ucsErrInvalidJsonRequest))
        GoTo QH
    End If
    Set oFP = pvInit(JsonItem(oJson, "DeviceString"))
    '--- success
    pvCommonProlog = True
QH:
End Function

Private Function pvCommonEpilog(oFP As IDeviceProtocol, sLastErr As String, ByVal eLastNo As UcsFiscalErrorsEnum, oRetVal As Object, sResponse As String) As Boolean
    If Not oFP Is Nothing Then
        If oFP.IsConnected Then
            oFP.Disconnect
        End If
        m_sCommandLog = oFP.GetCommandLog()
    End If
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
    sResponse = JsonDump(oRetVal, Minimize:=True)
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
    
    bIncludeAll = C_Bool(JsonItem(oJson, "IncludeAll"))
    If C_Bool(JsonItem(oJson, "IncludeTaxNo")) Or bIncludeAll Then
        JsonItem(oRetVal, "TaxNo") = Trim$(oFP.GetTaxNo())
        JsonItem(oRetVal, "TaxCaption") = Replace(Trim$(oFP.GetTaxCaption()), ":", vbNullString)
    End If
    If C_Bool(JsonItem(oJson, "IncludeHeaders")) Or bIncludeAll Then
        For lIdx = 1 To 6
            JsonItem(oRetVal, "Headers/-1") = Trim$(oFP.GetHeaderText(lIdx))
        Next
    End If
    If C_Bool(JsonItem(oJson, "IncludeFooters")) Or bIncludeAll Then
        For lIdx = 1 To 2
            JsonItem(oRetVal, "Footers/-1") = Trim$(oFP.GetFooterText(lIdx))
        Next
    End If
    If C_Bool(JsonItem(oJson, "IncludeReceiptNo")) Or bIncludeAll Then
        JsonItem(oRetVal, "LastReceiptNo") = StripZeros(oFP.GetLastQRCodeInfo(dLastDateTime))
        JsonItem(oRetVal, "LastReceiptDateTime") = dLastDateTime
        If LenB(JsonItem(oRetVal, "LastReceiptNo")) = 0 Then
            JsonItem(oRetVal, "LastReceiptNo") = StripZeros(oFP.GetLastReceiptNo())
            JsonItem(oRetVal, "LastReceiptDateTime") = oFP.GetClock()
        End If
    End If
    If C_Bool(JsonItem(oJson, "IncludePaymentNames")) Or bIncludeAll Then
        For lIdx = 1 To 4
            JsonItem(oRetVal, "PaymentNames/-1") = Trim$(oFP.GetPaymentName(lIdx))
        Next
        For lIdx = -1 To -4 Step -1
            JsonItem(oRetVal, "PaymentNames/-1") = Trim$(oFP.GetPaymentName(lIdx))
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
    FlushDebugLog
End Sub
