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
Private Const STR_DATECS_FP_INTERNAL    As String = "Време за достъп изтече в очакване на отговор|Невалиден формат на съобщение или сума за проверка (NAK)|Невалидна дължина на отговора: %1|Липсва символ за край на отговора: %1|Липсва символ за начало на отговора: %1|Няма започната бележка| или невалидна парола на оператор|Грешка на ред %1: %2|Невъзможно отказване на предишна транзакция|Невалиден тип справка|Грешка %1|Памет %1"
Private Const STR_DATECS_FP_STATUSES    As String = "Синтактична грешка|Невалидна команда|Неустановени дата и час|3|Неизправност в механизма на печатащото устройство|5|6|7|Аритметично препълване|Непозволена команда|Зануляване на оперативна памет|3|Разрушено съдържание на оперативна памет|5|6|7|Няма хартия|1|2|Отворен фискален бон|4|Отворен служебен (нефискален) бон|6|7"
Private Const STR_DATECS_FP_ERRORS      As String = "Синтактична грешка|Невалидна команда|2|3|Неизправност в механизма на печатащото устройство|5|6|7|1|Непозволена команда|Зануляване на оперативна памет|3|Разрушено съдържание на оперативна памет|5|6|7|Няма хартия|1|2|3|4|5|6|7"
Private Const STR_DATECS_FP_DIP_SWITCHES As String = "Автоматично центриране на header и footer|Предварителен header|Sw1.3|Sw1.4|Нулиране на паметта|Прозрачен дисплей|Без данни на дисплея|7"
Private Const STR_DATECS_FP_MEMORY      As String = "Грешка при запис|1|Няма модул фискална памет|Малко свободно място във фискалната памет|Пълна фискална памет|5|6|7|Фискалната памет забранена за запис|Форматирана фискална памет|2|Принтерът е във фискален режим|Зададени данъчни ставки|Програмирани индивидуален номер и номер на фискалната памет|6|7"
Private Const STR_DATECS_FP_TEXTS       As String = "В БРОЙ|С ДЕБИТНА КАРТА|С ЧЕК|С КРЕДИТНА КАРТА|НАДБАВКА %1|ОТСТЪПКА %1|Продажби %1|ЕИК"
Private Const STR_TREMOL_FP_INTERNAL    As String = "Време за достъп изтече в очакване на отговор|Невалиден формат на съобщение или сума за проверка (NAK)|Невалидна дължина на отговора: %1|Липсва символ за начало на отговора: %1|Няма започната бележка|Грешка на ред %1: %2|Невалиден тип справка|Устройството е недостъпно|Устройството е заето| или невъзможно установяване парола на оператор"
Private Const STR_TREMOL_FP_STATUSES    As String = "Неизвестна грешка|Невалидна команда|Непозволена команда|Непозволена поради ненулев отчет|Синтактична грешка|Препълване на входните регистри|Нулев входен регистър|Липсва транзакция която да се войдира|Недостатъчна налична сума|Конфликт в данъчните групи"
Private Const STR_TREMOL_FP_ERRORS      As String = "Неизвестна грешка|Няма хартия|Препълване на дневните регистри|Несверен/грешен часовник|Отворен фискален бон|Сметка с остатък за плащане|Отворен нефискален бон|Сметка с приключено плащане|Фискална памет само за четене|Грешна парола или непозволена команда|Липсващ външен дисплей|24 часа без дневен отчет|Прегрят принтер|Спад на напрежение във фискален бон|Препълване в електронната контролна лента|Недостатъчни условия"
Private Const STR_TREMOL_FP_TEXTS       As String = "В брой|С карта|С чек|Надбавка%1|Отстъпка%1|Междинна сума|Продажби %1|ЕИК;ЕГН;ЛНЧ;Служебен No"
Private Const STR_DAISY_FP_EXT_ERRORS   As String = "1|Въпросната операция ще доведе до препълване|3|Нямате право на повече продажби в този бон|4|Нямате право на повече плащания в този бон|5|Опит за извършване на нулева транзакция|6|Опит за извършване на продажба, след като е започнато плащане|7|Нямате право на такава операция|8|Забранена за продажби дан. група|11|Въвеждане на повече от една десетична точка|12|Въвеждане на повече от един символ '+' или '-'|13|Символът '+' или '-' не е в първа позиция|14|Недопустим символ, напр. баркод, който съдържа не само цифри|15|Повече от допустимия брой знаци след десетичната точка|16|Въведени са повече от разрешения брой символи|20|В дадената ситуация не се обслужва избраната от Вас команда от PC|21|Стойността е извън допустимите граници|22|Виж системен параметър 8|23|Опит за ""дълбок"" войд след отстъпка/надбавка в/у междинна сума|24|Опит за ""дълбок"" войд на несъществуваща транзакция|25|Опит за извършване на плащане, без да има продажби|" & _
                                                    "26|Опит за продажба на артикул с количество, надвишаващо запаса му|41|Некоректен баркод (грешна контролна сума)|42|Опит за продажба с нулев баркод|43|Опит за програмиране с тегловен баркод|44|Опит за продажба с непрограмиран баркод|45|Опит за програмиране на вече съществуващ баркод|66|Некоректна парола|71|!!! Некоректни данни във ФП !!!!|72|!!! Грешка при запис във ФП !!!!|76|Необходима е информация от сървъра на НАП|90|Не е нулиран периодичният отчет|91|Не е нулиран дневният финансов отчет|92|Не е нулиран отчетът по оператори|93|Не е нулиран отчетът по артикули|94|Не може да се препрограмира това поле|81|Дневният финансов отчет е препълнен|83|Отчетът по оператори е препълнен|84|Отчетът по артикули е препълнен|84|Периодичният отчет е препълнен|88|КЛЕН е препълнен|102|Няма комуникация между ФУ и ДТ|104|Некоректна комуникация между ФУ и ДТ|110|Подменена SIM карта|111|Грешка при комуникация между ДТ и сървъра на НАП|" & _
                                                    "113|Сървърът на НАП не приема подадените му данни|117|Неуспешен опит за свързване на ДТ с мрежата на мобилния оператор|118|Операцията е забранена|119|Грешна въведена стойност|120|Невъведена стойност"
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
End Enum

Private Type UcsRowData
    RowType         As UcsRowTypeEnum
    Item            As String
    Price           As String
    Discount        As String
    TaxGroup        As String
    Quantity        As String
    PmtType         As String
    Amount          As String
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
    ucsFscPixItem = 0
    ucsFscPixPrice
    ucsFscPixTaxGroup
    ucsFscPixQuantity
    ucsFscPixDiscount
    ucsFscPixWayOfPmt = ucsFscPixPrice
    ucsFscPixAmount
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
    For Each vElem In Array(pvInit("Protocol=" & STR_PROTOCOL_TREMOL_FP, SkipLastError:=True), _
                            pvInit("Protocol=" & STR_PROTOCOL_ESP_POS, SkipLastError:=True), _
                            pvInit("Protocol=" & STR_PROTOCOL_DATECS_FP, SkipLastError:=True))
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
    Dim lIdx            As Long
    Dim bIncludeAll     As Boolean
    
    On Error GoTo EH
    If Not pvCommonProlog(sRequest, oJson, sLastErr, oFP, oRetVal) Then
        GoTo QH
    End If
    JsonItem(oRetVal, "DeviceSerialNo") = oFP.GetDeviceSerialNo()
    JsonItem(oRetVal, "FiscalMemoryNo") = oFP.GetFiscalMemoryNo()
    JsonItem(oRetVal, "DeviceProtocol") = oFP.GetDeviceProtocol()
    JsonItem(oRetVal, "DeviceModel") = oFP.GetDeviceModel()
    JsonItem(oRetVal, "FirmwareVersion") = oFP.GetDeviceFirmware()
    JsonItem(oRetVal, "CharsPerLine") = oFP.GetCharsPerLine()
    bIncludeAll = C_Bool(JsonItem(oJson, "IncludeAll"))
    If C_Bool(JsonItem(oJson, "IncludeHeaders")) Or bIncludeAll Then
        For lIdx = 1 To 6
            JsonItem(oRetVal, "Header/-1") = oFP.GetHeaderText(lIdx)
        Next
    End If
    If C_Bool(JsonItem(oJson, "IncludeFooters")) Or bIncludeAll Then
        For lIdx = 1 To 2
            JsonItem(oRetVal, "Footer/-1") = oFP.GetFooterText(lIdx)
        Next
    End If
    If C_Bool(JsonItem(oJson, "IncludeTaxNo")) Or bIncludeAll Then
        JsonItem(oRetVal, "TaxNo") = oFP.GetTaxNo()
        JsonItem(oRetVal, "TaxCaption") = Replace(Trim$(oFP.GetTaxCaption()), ":", vbNullString)
    End If
    If C_Bool(JsonItem(oJson, "IncludeReceiptNo")) Or bIncludeAll Then
        JsonItem(oRetVal, "ReceiptNo") = oFP.GetLastReceiptNo()
        JsonItem(oRetVal, "DeviceDateTime") = oFP.GetClock()
    End If
    If C_Bool(JsonItem(oJson, "IncludePaymentNames")) Or bIncludeAll Then
        For lIdx = 1 To 3
            JsonItem(oRetVal, "PaymentName/-1") = oFP.GetPaymentName(lIdx)
        Next
        For lIdx = -1 To -4 Step -1
            JsonItem(oRetVal, "PaymentName/-1") = oFP.GetPaymentName(lIdx)
        Next
    End If
    If LenB(C_Str(JsonItem(oJson, "Operator/Code"))) <> 0 Then
        JsonItem(oRetVal, "Operator/Code") = JsonItem(oJson, "Operator/Code")
        JsonItem(oRetVal, "Operator/Name") = oFP.GetOperatorName(JsonItem(oJson, "Operator/Code"))
        JsonItem(oRetVal, "Operator/Password") = oFP.GetDefaultPassword(JsonItem(oJson, "Operator/Code"))
    End If
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
    Dim sResumeTicket   As String
    Dim sPrevReceiptNo  As String
    
    On Error GoTo EH
    If Not pvCommonProlog(sRequest, oJson, sLastErr, oFP, oRetVal) Then
        GoTo QH
    End If
    sResumeTicket = JsonItem(oJson, "ResumeTicket")
    If LenB(sResumeTicket) = 0 Then
        If Not oFP.CancelReceipt() Then
            sLastErr = oFP.GetLastError(eLastNo)
            DebugLog FUNC_NAME, "Not oFP.CancelReceipt(), sLastErr=" & sLastErr & ", eLastNo=" & eLastNo
            GoTo QH
        End If
    End If
    sPrevReceiptNo = JsonItem(oJson, "PrevReceiptNo")
    If LenB(sPrevReceiptNo) <> 0 Then
        JsonItem(oRetVal, "PrevReceiptNo") = oFP.GetLastReceiptNo()
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
        If Not oFP.StartReceipt(C_Lng(JsonItem(oJson, "ReceiptType")), _
                JsonItem(oJson, "Operator/Code"), _
                Zn(JsonItem(oJson, "Operator/Name"), pvText(ucsTxtCashier)), _
                JsonItem(oJson, "Operator/Password"), _
                JsonItem(oJson, "UniqueSaleNo"), _
                JsonItem(oJson, "TableNo"), _
                JsonItem(oJson, "Invoice/DocNo"), _
                JsonItem(oJson, "Invoice/CgTaxNo"), _
                JsonItem(oJson, "Invoice/CgVatNo"), _
                JsonItem(oJson, "Invoice/CgName"), _
                vbNullString, _
                JsonItem(oJson, "Invoice/CgAddress"), _
                JsonItem(oJson, "Invoice/CgPrsReceive"), _
                JsonItem(oJson, "OwnData"), _
                C_Lng(JsonItem(oJson, "Reversal/Type")), _
                JsonItem(oJson, "Reversal/ReceiptNo"), _
                C_Date(JsonItem(oJson, "Reversal/ReceiptDate")), _
                JsonItem(oJson, "Reversal/FiscalMemoryNo"), _
                JsonItem(oJson, "Reversal/InvoiceNo"), _
                JsonItem(oJson, "Reversal/Reason")) Then
            sLastErr = oFP.GetLastError(eLastNo)
            DebugLog FUNC_NAME, "Not oFP.StartReceipt(), sLastErr=" & sLastErr & ", eLastNo=" & eLastNo
            GoTo QH
        End If
        For Each vElem In JsonKeys(oJson, "Rows")
            With pvGetRowData(JsonItem(oJson, "Rows/" & vElem))
                Select Case .RowType
                Case ucsRwtSell
                    oFP.AddPLU .Item, C_Dbl(.Price), IIf(LenB(.Quantity) <> 0, C_Dbl(.Quantity), 1), pvGetTaxGroup(.TaxGroup)
                    If C_Dbl(.Discount) <> 0 Then
                        oFP.AddDiscount ucsFscDscPlu, -C_Dbl(.Discount)
                    End If
                Case ucsRwtText
                    oFP.AddLine .Item
                Case ucsRwtPayment
                    lIdx = C_Lng(.PmtType)
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
    JsonItem(oRetVal, "ReceiptNo") = oFP.GetLastReceiptNo()
    JsonItem(oRetVal, "ReceiptDateTime") = oFP.GetClock()
    JsonItem(oRetVal, "DeviceSerialNo") = oFP.GetDeviceSerialNo()
    JsonItem(oRetVal, "FiscalMemoryNo") = oFP.GetFiscalMemoryNo()
QH:
    PrintReceipt = pvCommonEpilog(oFP, sLastErr, eLastNo, oRetVal, sResponse)
    Exit Function
EH:
    If Not oFP Is Nothing Then
        JsonItem(oRetVal, "ResumeTicket") = Zn(oFP.GetResumeTicket(), Empty)
    End If
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
    
    On Error GoTo EH
    If Not pvCommonProlog(sRequest, oJson, sLastErr, oFP, oRetVal) Then
        GoTo QH
    End If
    Select Case JsonItem(oJson, "ReportType")
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
    JsonItem(oRetVal, "ReceiptNo") = oFP.GetLastReceiptNo()
    JsonItem(oRetVal, "ReceiptDateTime") = oFP.GetClock()
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
        bResult = oFP.SetClock(dDateTime)
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
QH:
    InitDateTime = pvCommonEpilog(oFP, sLastErr, eLastNo, oRetVal, sResponse)
    Exit Function
EH:
    If pvCommonErrHandler(FUNC_NAME, oFP, sLastErr, eLastNo) Then
        Resume QH
    End If
End Function

''
' Used to retrieve accumulated totals by payment types since last z-report
'
Public Function GetTotalsByPayments(sRequest As String, sResponse As String) As Boolean
    Const FUNC_NAME     As String = "GetTotalsByPayments"
    Dim oFP             As IDeviceProtocol
    Dim oJson           As Object
    Dim oRetVal         As Object
    Dim sLastErr        As String
    Dim eLastNo         As UcsFiscalErrorsEnum
    Dim vTotals         As Variant
    Dim lIdx            As Long
    Dim oItem           As Object
    
    On Error GoTo EH
    If Not pvCommonProlog(sRequest, oJson, sLastErr, oFP, oRetVal) Then
        GoTo QH
    End If
    vTotals = oFP.GetTotalsByPayments()
    If Not IsArray(vTotals) Then
        sLastErr = pvInternal(ucsErrNotImplemented)
        GoTo QH
    ElseIf UBound(vTotals) < 0 Then
        sLastErr = oFP.GetLastError(eLastNo)
        DebugLog FUNC_NAME, "UBound(vTotals) < 0, sLastErr=" & sLastErr & ", eLastNo=" & eLastNo
        GoTo QH
    End If
    For lIdx = 0 To UBound(vTotals)
        Set oItem = Nothing
        JsonItem(oItem, "PaymentName") = oFP.GetPaymentName(lIdx + 1)
        JsonItem(oItem, "Amount") = vTotals(lIdx)
        JsonItem(oRetVal, "Totals/-1") = oItem
    Next
QH:
    GetTotalsByPayments = pvCommonEpilog(oFP, sLastErr, eLastNo, oRetVal, sResponse)
    Exit Function
EH:
    If pvCommonErrHandler(FUNC_NAME, oFP, sLastErr, eLastNo) Then
        Resume QH
    End If
End Function

''
' Used to retrieve accumulated totals by VAT groups (1-8) since last z-report
'
Public Function GetTotalsByTaxGroups(sRequest As String, sResponse As String) As Boolean
    Const FUNC_NAME     As String = "GetTotalsByTaxGroups"
    Dim oFP             As IDeviceProtocol
    Dim oJson           As Object
    Dim oRetVal         As Object
    Dim sLastErr        As String
    Dim eLastNo         As UcsFiscalErrorsEnum
    Dim vTotals         As Variant
    Dim lIdx            As Long
    Dim oItem           As Object
    
    On Error GoTo EH
    If Not pvCommonProlog(sRequest, oJson, sLastErr, oFP, oRetVal) Then
        GoTo QH
    End If
    vTotals = oFP.GetTotalsByTaxGroups()
    If Not IsArray(vTotals) Then
        sLastErr = pvInternal(ucsErrNotImplemented)
        GoTo QH
    ElseIf UBound(vTotals) < 0 Then
        sLastErr = oFP.GetLastError(eLastNo)
        DebugLog FUNC_NAME, "UBound(vTotals) < 0, sLastErr=" & sLastErr & ", eLastNo=" & eLastNo
        GoTo QH
    End If
    For lIdx = 0 To UBound(vTotals)
        Set oItem = Nothing
        JsonItem(oItem, "TaxGroup") = lIdx + 1
        JsonItem(oItem, "VatPercent") = vTotals(lIdx)(1)
        JsonItem(oItem, "Amount") = vTotals(lIdx)(0)
        JsonItem(oRetVal, "Totals/-1") = oItem
    Next
QH:
    GetTotalsByTaxGroups = pvCommonEpilog(oFP, sLastErr, eLastNo, oRetVal, sResponse)
    Exit Function
EH:
    If pvCommonErrHandler(FUNC_NAME, oFP, sLastErr, eLastNo) Then
        Resume QH
    End If
End Function

''
' Used to retrieve number of receipts printed since last z-report and the date/time of the last receipt printed
'
Public Function GetTotalsStats(sRequest As String, sResponse As String) As Boolean
    Const FUNC_NAME     As String = "GetTotalsStats"
    Dim oFP             As IDeviceProtocol
    Dim oJson           As Object
    Dim oRetVal         As Object
    Dim sLastErr        As String
    Dim eLastNo         As UcsFiscalErrorsEnum
    Dim vStats          As Variant

    On Error GoTo EH
    If Not pvCommonProlog(sRequest, oJson, sLastErr, oFP, oRetVal) Then
        GoTo QH
    End If
    vStats = oFP.GetTotalsStats()
    If Not IsArray(vStats) Then
        sLastErr = pvInternal(ucsErrNotImplemented)
        GoTo QH
    ElseIf UBound(vStats) < 0 Then
        sLastErr = oFP.GetLastError(eLastNo)
        DebugLog FUNC_NAME, "UBound(vStats) < 0, sLastErr=" & sLastErr & ", eLastNo=" & eLastNo
        GoTo QH
    End If
    JsonItem(oRetVal, "NumReceipts") = vStats(0)
    JsonItem(oRetVal, "LastZReportDateTime") = vStats(1)
QH:
    GetTotalsStats = pvCommonEpilog(oFP, sLastErr, eLastNo, oRetVal, sResponse)
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
        JsonItem(oRetVal, "ReceiptNo") = oFP.GetLastReceiptNo()
        JsonItem(oRetVal, "ReceiptDateTime") = oFP.GetClock()
    End If
    JsonItem(oRetVal, "Available") = vTotals(0)
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
    
    Set oOptions = ParseDeviceString(DeviceString)
    sProtocol = UCase$(JsonItem(oOptions, "Protocol"))
    '--- figure out model
    Select Case sProtocol
    Case STR_PROTOCOL_DATECS_FP, STR_PROTOCOL_DAISY_ECR, STR_PROTOCOL_INCOTEX_ECR
        Set pvInit = New cIslProtocol
        pvInit.SetLocalizedText ucsFscLciInternalErrors, GetConfigValue(sProtocol, CFG_INTERNAL_ERRORS, STR_DATECS_FP_INTERNAL)
        pvInit.SetLocalizedText ucsFscLciPrinterStatuses, GetConfigValue(sProtocol, CFG_PRINTER_STATUSES, STR_DATECS_FP_STATUSES)
        pvInit.SetLocalizedText ucsFscLciPrinterErrors, GetConfigValue(sProtocol, CFG_PRINTER_ERRORS, STR_DATECS_FP_ERRORS)
        pvInit.SetLocalizedText ucsFscLciDipSwitchesTexts, GetConfigValue(sProtocol, CFG_DIP_SWITCHES_TEXTS, STR_DATECS_FP_DIP_SWITCHES)
        pvInit.SetLocalizedText ucsFscLciMemoryTexts, GetConfigValue(sProtocol, CFG_MEMORY_TEXTS, STR_DATECS_FP_MEMORY)
        pvInit.SetLocalizedText ucsFscLciReceiptTexts, GetConfigValue(sProtocol, CFG_RECEIPT_TEXTS, STR_DATECS_FP_TEXTS)
        pvInit.SetLocalizedText ucsFscLciExtendedErrors, GetConfigValue(sProtocol, CFG_EXTENDED_ERRORS, STR_DAISY_FP_EXT_ERRORS)
        pvInit.SetLocalizedText ucsFscLciConnectorErrors, GetConfigValue(sProtocol, CFG_CONNECTOR_ERRORS, STR_CONNECTOR_ERRORS)
    Case STR_PROTOCOL_TREMOL_FP
        Set pvInit = New cTremolProtocol
        pvInit.SetLocalizedText ucsFscLciInternalErrors, GetConfigValue(sProtocol, CFG_INTERNAL_ERRORS, STR_TREMOL_FP_INTERNAL)
        pvInit.SetLocalizedText ucsFscLciPrinterStatuses, GetConfigValue(sProtocol, CFG_PRINTER_STATUSES, STR_TREMOL_FP_STATUSES)
        pvInit.SetLocalizedText ucsFscLciPrinterErrors, GetConfigValue(sProtocol, CFG_PRINTER_ERRORS, STR_TREMOL_FP_ERRORS)
        pvInit.SetLocalizedText ucsFscLciReceiptTexts, GetConfigValue(sProtocol, CFG_RECEIPT_TEXTS, STR_TREMOL_FP_TEXTS)
        pvInit.SetLocalizedText ucsFscLciConnectorErrors, GetConfigValue(sProtocol, CFG_CONNECTOR_ERRORS, STR_CONNECTOR_ERRORS)
    Case STR_PROTOCOL_ESP_POS
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
    pvGetRowData.Price = C_Str(JsonItem(oRow, "Price"))
    If LenB(pvGetRowData.Price) <> 0 Then
        pvGetRowData.RowType = ucsRwtSell
        pvGetRowData.Item = C_Str(JsonItem(oRow, "ItemName"))
        pvGetRowData.TaxGroup = C_Str(JsonItem(oRow, "TaxGroup"))
        pvGetRowData.Quantity = C_Str(JsonItem(oRow, "Quantity"))
        pvGetRowData.Discount = C_Str(JsonItem(oRow, "Discount"))
        GoTo QH
    End If
    If Not IsEmpty(JsonItem(oRow, "PaymentType")) Then
        pvGetRowData.RowType = ucsRwtPayment
        pvGetRowData.Amount = C_Str(JsonItem(oRow, "Amount"))
        pvGetRowData.PmtType = C_Str(JsonItem(oRow, "PaymentType"))
        pvGetRowData.Item = C_Str(JsonItem(oRow, "PaymentName"))
        GoTo QH
    End If
    pvGetRowData.Item = C_Str(JsonItem(oRow, "Text"))
    If LenB(pvGetRowData.Item) <> 0 Then
        pvGetRowData.RowType = ucsRwtText
        GoTo QH
    End If
    vRow = JsonItem(oRow, "*")
    If UBound(vRow) = ucsFscPixItem Then
        pvGetRowData.RowType = ucsRwtText
        pvGetRowData.Item = C_Str(vRow(ucsFscPixItem))
        GoTo QH
    End If
    If UBound(vRow) = ucsFscPixAmount Then
        pvGetRowData.RowType = ucsRwtPayment
        pvGetRowData.Item = C_Str(vRow(ucsFscPixItem))
        pvGetRowData.PmtType = C_Str(vRow(ucsFscPixWayOfPmt))
        pvGetRowData.Amount = C_Str(vRow(ucsFscPixAmount))
        GoTo QH
    End If
    pvGetRowData.RowType = ucsRwtSell
    pvGetRowData.Item = C_Str(vRow(ucsFscPixItem))
    pvGetRowData.Price = C_Str(vRow(ucsFscPixPrice))
    pvGetRowData.TaxGroup = C_Str(vRow(ucsFscPixTaxGroup))
    pvGetRowData.Quantity = C_Str(vRow(ucsFscPixQuantity))
    pvGetRowData.Discount = At(vRow, ucsFscPixDiscount)
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function pvGetTaxGroup(sTaxGroup As String) As Long
    pvGetTaxGroup = Asc(sTaxGroup & " ") - 191 '--- 191 = Asc("А") - 1
    If pvGetTaxGroup < 1 Or pvGetTaxGroup > 8 Then
        pvGetTaxGroup = C_Lng(sTaxGroup)
        If pvGetTaxGroup < 1 Or pvGetTaxGroup > 8 Then
            pvGetTaxGroup = 2
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
    If eLastNo = ucsFerInvalidPassword Then
        JsonItem(oRetVal, "Ok") = False
        JsonItem(oRetVal, "ErrorCode") = "InvalidPassword"
    End If
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


