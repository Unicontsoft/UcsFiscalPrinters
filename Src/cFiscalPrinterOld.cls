VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFiscalPrinterOld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=========================================================================
' $Header: /UcsFiscalPrinter/Src/cFiscalPrinterOld.cls 27    22.02.11 10:07 Wqw $
'
'   Unicontsoft Fiscal Printers Project
'   Copyright (c) 2008-2011 Unicontsoft
'
'   Uprawlenie na fiscalni ustrojstwa prez cICLProtocol
'
' $Log: /UcsFiscalPrinter/Src/cFiscalPrinterOld.cls $
' 
' 27    22.02.11 10:07 Wqw
' REF: polzwa string functions
'
' 26    21.02.11 13:45 Wqw
' REF: prazen string, pechat periodichen report
'
' 25    14.02.11 18:14 Wqw
' REF: param timeout impl w DeviceString
'
' 24    29.11.10 10:51 Wqw
' REF: dublikat na nefiskalen bon
'
' 23    26.11.10 13:20 Wqw
' REF: impl daisy FP
'
' 22    17.08.10 15:34 Wqw
' REF: finalizirane fiskalen bon pri pechat na prazen tekst samo
'
' 21    5.08.10 15:55 Wqw
' REF: impl pechat na text mejdu redowe na fiskalna belejka
'
' 20    17.05.10 16:16 Wqw
' REF: debug log
'
' 19    23.03.09 19:05 Wqw
' ADD: pvCalcRow i pvCalcDiscount za unifitsirane na smqtane na otstypka
'
' 18    14.03.09 16:14 Wqw
' REF: impl non-invoice
'
' 17    19.01.09 17:37 Wqw
' REF: log-wa fiscalno plashtane
'
' 16    17.12.08 18:28 Wqw
' REF: impl pechat na kreditno izwestie, DebugLog prowerqwa dali log
' file-a syshtestwuwa
'
' 15    3.12.08 23:49 Wqw
' REF: custom impl na Round, sumi s natrupwaniq se zakryglqwat
'
' 14    2.12.08 0:15 Wqw
' REF: logging
'
' 13    1.12.08 23:02 Wqw
' REF: logging
'
' 12    1.12.08 16:14 Wqw
' REF: impl debug log
'
' 11    1.12.08 0:43 Wqw
' REF: polzwat pvReset
'
' 10    28.11.08 2:06 Wqw
' REF: manual open door pri nefiskalni bonowe, impl open door
'
' 9     11.11.08 18:20 Wqw
' REF: pri slujebnite bonowe moje da ima otritsatelni plashtaniq
'
' 8     11.11.08 17:19 Wqw
' ADD: InitDateTime. REF: impl na redowe za plashtane i free text pri
' pechat na bonowe
'
' 7     4.11.08 14:19 Wqw
' REF: internal errors
'
' 6     30.10.08 22:49 Wqw
' REF: cancel outstanding transaction razbira ot neplaten bon
'
' 5     29.10.08 13:52 Wqw
' REF: impl pechat na dublikat prez PrintReceipt (bez Rows)
'
' 4     15.10.08 16:13 Wqw
' REF: suma za plashtane
'
' 3     14.10.08 19:10 Wqw
' ADD: Function PrintReport, OpenDrawer. REF: GetHeaderFooter wryshta
' selektiwno rezultat
'
' 2     19.06.08 18:26 Wqw
' REF: ne prowerqwa za is numeric quantity
'
' 1     18.06.08 16:45 Wqw
' Initial implementation
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "cFiscalPrinter"

'=========================================================================
' API
'=========================================================================

Private Const LOCALE_SDECIMAL                       As Long = &HE   ' decimal separator

Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Declare Function GetUserDefaultLCID Lib "kernel32" () As Long

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_TOTAL                 As String = "ОБЩО"
Private Const STR_CASH                  As String = "В БРОЙ"
Private Const STR_CARD                  As String = "С ДЕБИТНА КАРТА"
Private Const STR_CHANGE                As String = "РЕСТО"
Private Const STR_INTERNAL              As String = "Не е указано устройство|Грешка при отваряне: |Грешка при SetCommTimeouts: |Грешка при BuildCommDCB: |Грешка при SetCommState: |Грешка при WriteFile: |Време за достъп изтече в очакване на отговор|Грешка при ReadFile: |Невалиден формат на съобщение или сума за проверка (NAK)|Вече има отворена връзка|Няма отворена връзка|Невалидна дължина на отговора|Липсва символ за край на отговора|Липсва символ за начало на отговора|Грешка при WaitCommEvent: "
Private Const STR_STATUSES              As String = "Синтактична грешка|Невалидна команда|Неустановени дата и час|3|Неизправност в механизма на печатащото устройство|5|6|7|Аритметично препълване|Непозволена команда|Зануляване на опертивна памет|3|Разрушено съдържание на опертивна памет|5|6|7|Няма хартия|1|2|Отворен фискален бон|4|Отворен служебен (нефискален) бон|6|7"
Private Const STR_ERRORS_ONLY           As String = "Синтактична грешка|Невалидна команда|2|3|Неизправност в механицва на печатащото устройство|5|6|7|1|Непозволена команда|Зануляване на опертивна памет|3|Разрушено съдържание на опертивна памет|5|6|7|Няма хартия|1|2|3|4|5|6|7"
Private Const STR_DIP_SWITCHES          As String = "Автоматично центриране на header и footer|Предварителен header|Sw1.3|Sw1.4|Нулиране на паметта|Прозрачен дисплей|Без данни на дисплея|7"
Private Const STR_MEMORY                As String = "Грешка при запис|1|Няма модул фискална памет|Малко свободно място във фискалната памет|Пълна фискална памет|5|6|7|Фискалната памет забранена за запис|Форматирана фискална памет|2|Принтерът е във фискален режим|Зададени данъчни ставки|Програмирани индивидуален номер и номер на фискалната памет|6|7"
Private Const STR_ERROR_LINE            As String = "Грешка при печат на ред "
Private Const STR_OR_INCORRECT_LOGIN    As String = " или невалидна парола на оператор"
Private Const STR_HEAD_CREDIT_NOTE      As String = "КРЕДИТНО ИЗВЕСТИЕ N:"
Private Const STR_HEAD_CREDIT_NOTE2     As String = "към ф-ра "
Private Const STR_DUPLICATE             As String = "ДУБЛИКАТ"
Private Const STR_ORIGINAL              As String = "ОРИГИНАЛ"
Private Const STR_SELLER                As String = "ПРОДАВАЧ: "
Private Const STR_RECEIVER              As String = "ПОЛУЧАТЕЛ: "
Private Const STR_BUYER                 As String = "КУПУВАЧ: "
Private Const STR_BULSTAT               As String = "ИДЕНТ. No: "
Private Const STR_VATNO                 As String = "ЗДДС No: "
Private Const STR_ADDRESS               As String = "АДРЕС: "
Private Const STR_UNIT_PRICE            As String = "ЕДИНИЧНА ЦЕНА"
Private Const STR_UNIT_QUO              As String = "КОЛИЧЕСТВО"
Private Const STR_UNIT_SUM              As String = "СУМА"
Private Const STR_GROUP_TOTAL           As String = "ВСИЧКО ГРУПА "
Private Const STR_GROUP_VAT             As String = "ДДС "
Private Const STR_GROUP_NETO            As String = "НЕТО СТОЙНОСТ"
Private Const STR_ARTICLES              As String = " АРТИКУЛА"
Private Const STR_ARTICLE_SINGLE        As String = "1 АРТИКУЛ"
Private Const STR_DISCOUNT              As String = "ОТСТЪПКА "
Private Const STR_DEFAULT_OPER_PASS     As String = "0000"
Private Const DEF_COMMENT_LEN           As Long = 30

Private WithEvents m_oFP            As cICLProtocol
Attribute m_oFP.VB_VarHelpID = -1
Private m_sLog                      As String
Private m_bLogFiscalPayment         As Boolean
Private m_sDecimalSeparator         As String

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

Private Type UcsContext
    Total           As Double
    Amount          As Double
    RowPrinted      As Boolean
    PmtPrinted      As Boolean
    ChangePrinted   As Boolean
    GrpTotal(0 To 10) As Double
    Count           As Long
    FpSerNum        As String
    FpCommentLen    As Long
End Type

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunc As String)
    Debug.Print MODULE_NAME & "." & sFunc & ": " & Error
    DebugLog sFunc, "Run-time error: " & Error
End Sub

Private Sub DebugLog(sFunc As String, sText As String)
    OutputDebugLog MODULE_NAME, sFunc, sText
End Sub

'=========================================================================
' Methods
'=========================================================================

Public Function EnumPorts() As Variant
'    With New cICLProtocol
'        EnumPorts = .EnumPorts
'    End With
    EnumPorts = EnumSerialPorts
End Function

Public Function GetHeaderFooter( _
            DeviceString As String, _
            Optional OperatorData As Variant, _
            Optional RetVal As Variant, _
            Optional LastError As String, _
            Optional CommandLog As String) As Boolean
    Dim vInput      As Variant
    Dim vElem       As Variant
    
    On Error GoTo EH
    '--- init & localize fiscal printer
    pvInit DeviceString
    m_oFP.Exceptions = False
    '--- fetch header & footer
    vInput = RetVal
    ReDim RetVal(0 To 18) As Variant
    If LenB(At(vInput, 0, "1")) <> 0 Then
        RetVal(0) = m_oFP.SendCommand(ucsFpcInitHeaderFooter, "I0")
    End If
    If LenB(At(vInput, 1, "1")) <> 0 Then
        RetVal(1) = m_oFP.SendCommand(ucsFpcInitHeaderFooter, "I1")
    End If
    If LenB(At(vInput, 2, "1")) <> 0 Then
        RetVal(2) = m_oFP.SendCommand(ucsFpcInfoBulstat)
    End If
    If LenB(At(vInput, 3, "1")) <> 0 Then
        RetVal(3) = m_oFP.SendCommand(ucsFpcInitHeaderFooter, "I2")
    End If
    If LenB(At(vInput, 4, "1")) <> 0 Then
        RetVal(4) = m_oFP.SendCommand(ucsFpcInitHeaderFooter, "I3")
    End If
    If LenB(At(vInput, 5, "1")) <> 0 Then
        RetVal(5) = m_oFP.SendCommand(ucsFpcInitHeaderFooter, "I4")
    End If
    If LenB(At(vInput, 6, "1")) <> 0 Then
        RetVal(6) = m_oFP.SendCommand(ucsFpcInitHeaderFooter, "I5")
    End If
    If LenB(At(vInput, 7, "1")) <> 0 Then
        m_oFP.Exceptions = False
        RetVal(7) = m_oFP.SendCommand(ucsFpcInitHeaderFooter, "I6")
        m_oFP.Exceptions = True
    End If
    If LenB(At(vInput, 8, "1")) <> 0 Then
        m_oFP.Exceptions = False
        RetVal(8) = m_oFP.SendCommand(ucsFpcInitHeaderFooter, "I7")
        m_oFP.Exceptions = True
    End If
    If LenB(At(vInput, 9, "1")) <> 0 Then
        RetVal(9) = m_oFP.SendCommand(ucsFpcInfoDiagnostics, "0")
    End If
    If LenB(At(vInput, 10, "1")) <> 0 Then
        If LenB(At(OperatorData, 0)) <> 0 Then
            RetVal(10) = m_oFP.SendCommand(ucsFpcInfoOperator, At(OperatorData, 0))
        End If
    End If
    If LenB(At(vInput, 11, "1")) <> 0 Then
        RetVal(11) = m_oFP.SendCommand(ucsFpcInfoDateTime)
    End If
    If LenB(At(vInput, 12, "1")) <> 0 Then
        RetVal(12) = m_oFP.SendCommand(ucsFpcInfoLastReceipt)
    End If
    If LenB(At(vInput, 13, "1")) <> 0 Then
        m_oFP.Exceptions = False
        RetVal(13) = m_oFP.SendCommand(ucsFpcInitPaymentType, "I")
        m_oFP.Exceptions = True
        If m_oFP.Status(ucsStbPrintingError) Then
            RetVal(13) = m_oFP.SendCommand(ucsFpcInitDaisyText, "R61")
        End If
    End If
    If LenB(At(vInput, 14, "1")) <> 0 Then
        m_oFP.Exceptions = False
        RetVal(14) = m_oFP.SendCommand(ucsFpcInitPaymentType, "J")
        m_oFP.Exceptions = True
        If m_oFP.Status(ucsStbPrintingError) Then
            RetVal(14) = m_oFP.SendCommand(ucsFpcInitDaisyText, "R62")
        End If
    End If
    If LenB(At(vInput, 15, "1")) <> 0 Then
        m_oFP.Exceptions = False
        RetVal(15) = m_oFP.SendCommand(ucsFpcInitPaymentType, "K")
        m_oFP.Exceptions = True
        If m_oFP.Status(ucsStbPrintingError) Then
            RetVal(15) = m_oFP.SendCommand(ucsFpcInitDaisyText, "R63")
        End If
    End If
    If LenB(At(vInput, 16, "1")) <> 0 Then
        m_oFP.Exceptions = False
        RetVal(16) = m_oFP.SendCommand(ucsFpcInitPaymentType, "L")
        m_oFP.Exceptions = True
        If m_oFP.Status(ucsStbPrintingError) Then
            RetVal(16) = m_oFP.SendCommand(ucsFpcInitDaisyText, "R64")
        End If
    End If
    If LenB(At(vInput, 17, "1")) <> 0 Then
        m_oFP.Exceptions = False
        If Left$(At(Split(m_oFP.SendCommand(ucsFpcInfoDiagnostics, "0"), ","), 4), 2) = "DY" Then '--- daisy
            vElem = Split(m_oFP.SendCommand(ucsFpcInfoDaisyConsts), ",")
        End If
        RetVal(17) = At(vElem, 10, DEF_COMMENT_LEN)
        m_oFP.Exceptions = True
    End If
    If LenB(At(vInput, 18, "1")) <> 0 Then
        RetVal(18) = STR_DEFAULT_OPER_PASS
        If Left$(At(Split(m_oFP.SendCommand(ucsFpcInfoDiagnostics, "0"), ","), 4), 2) = "DY" Then '--- daisy
            RetVal(18) = RetVal(18) & "1"
        End If
    End If
    '--- disconnect
    Set m_oFP = Nothing
    CommandLog = m_sLog
    '--- success
    GetHeaderFooter = True
    Exit Function
EH:
    LastError = Err.Description
    Resume NextLine
NextLine:
    On Error Resume Next
    Set m_oFP = Nothing
    CommandLog = m_sLog
End Function

Public Function GetOperators( _
            DeviceString As String, _
            Optional ByVal Count As Long, _
            Optional RetVal As Variant, _
            Optional LastError As String, _
            Optional CommandLog As String) As Boolean
    Dim lIdx            As Long
    
    On Error GoTo EH
    '--- init & localize fiscal printer
    pvInit DeviceString
    '--- fetch operators
    ReDim RetVal(0 To IIf(Count > 0, Count, 24) - 1) As Variant
    For lIdx = 1 To UBound(RetVal) + 1
        RetVal(lIdx - 1) = m_oFP.SendCommand(ucsFpcInfoOperator, C_Str(lIdx))
    Next
    '--- disconnect
    Set m_oFP = Nothing
    CommandLog = m_sLog
    '--- success
    GetOperators = True
    Exit Function
EH:
    LastError = Err.Description
    Resume NextLine
NextLine:
    On Error Resume Next
    Set m_oFP = Nothing
    CommandLog = m_sLog
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
    Dim vElem           As Variant
    Dim bInTrans        As Boolean
    Dim bInLogin        As Boolean
    Dim lCurrentRow     As Long
    Dim sOperLogin      As String
    Dim sText           As String
    Dim vCmd            As Variant
    Dim uCtx            As UcsContext
    
    On Error GoTo EH
    '--- init & localize fiscal printer
    pvInit DeviceString
    '--- cancel outstanding transaction
    pvReset
    '--- gather FP info
    m_oFP.Exceptions = False
    uCtx.FpSerNum = At(Split(m_oFP.SendCommand(ucsFpcInfoDiagnostics, "0"), ","), 4)
    If Left$(uCtx.FpSerNum, 2) = "DY" Then '--- daisy
        vElem = Split(m_oFP.SendCommand(ucsFpcInfoDaisyConsts), ",")
    End If
    uCtx.FpCommentLen = At(vElem, 10, DEF_COMMENT_LEN)
    m_oFP.Exceptions = True
    '--- open transaction
    sOperLogin = At(OperatorData, ucsFscPixOperatorCode, "1") & "," & At(OperatorData, ucsFscPixOperatorPass, "0000") & "," & At(OperatorData, ucsFscPixCashDeskNo, "1")
    Select Case ReceiptType
    Case ucsFscRetNonfiscal
        m_oFP.SendCommand ucsFpcNonFiscalOpen
        If PrintDuplicate Then
            sText = STR_DUPLICATE
            m_oFP.SendCommand ucsFpcNonFiscalText, pvSpace((uCtx.FpCommentLen - Len(sText)) \ 2) & sText      '--- center
        End If
        If LenB(At(OperatorData, ucsFscPixOperatorCode)) <> 0 Then
            vElem = Split(m_oFP.SendCommand(ucsFpcInfoOperator, At(OperatorData, ucsFscPixOperatorCode)), ",")
            sText = Format$(At(OperatorData, ucsFscPixOperatorCode), "0000") & " " & At(vElem, 5)
            If LenB(At(OperatorData, ucsFscPixCashDeskNo)) <> 0 Then
                sText = sText & Right$(pvSpace(uCtx.FpCommentLen) & Format$(At(OperatorData, ucsFscPixCashDeskNo), "00"), uCtx.FpCommentLen - Len(sText))
            End If
            m_oFP.SendCommand ucsFpcNonFiscalText, sText
        End If
    Case ucsFscRetFiscal
        bInLogin = True
        m_oFP.SendCommand ucsFpcFiscalOpen, sOperLogin
        bInLogin = False
    Case ucsFscRetInvoiceFiscal
        If LenB(At(InvoiceData, ucsFscPixInvoiceNo)) <> 0 Then
            m_oFP.Exceptions = False
            m_oFP.SendCommand ucsFpcInitInvoiceNo, At(InvoiceData, ucsFscPixInvoiceNo) & "," & At(InvoiceData, ucsFscPixInvoiceNo)
            m_oFP.Exceptions = True
        End If
        bInLogin = True
        m_oFP.SendCommand ucsFpcFiscalOpen, sOperLogin & ",I"
        bInLogin = False
    Case ucsFscRetInvoiceNonFiscal
        m_oFP.SendCommand ucsFpcNonFiscalOpen
        If LenB(At(InvoiceData, ucsFscPixInvoiceNo)) <> 0 Then
            sText = At(InvoiceData, ucsFscPixInvoiceNoPrefix, STR_HEAD_CREDIT_NOTE) & At(InvoiceData, ucsFscPixInvoiceNo)
            m_oFP.SendCommand ucsFpcNonFiscalText, pvSpace((uCtx.FpCommentLen - Len(sText)) \ 2) & sText  '--- center
        End If
        If LenB(At(InvoiceData, ucsFscPixLnkDoc)) <> 0 Then
            sText = STR_HEAD_CREDIT_NOTE2 & At(InvoiceData, ucsFscPixLnkDoc)
            m_oFP.SendCommand ucsFpcNonFiscalText, pvSpace((uCtx.FpCommentLen - Len(sText)) \ 2) & sText  '--- center
        End If
        sText = IIf(PrintDuplicate, STR_DUPLICATE, STR_ORIGINAL)
        m_oFP.SendCommand ucsFpcNonFiscalText, pvSpace((uCtx.FpCommentLen - Len(sText)) \ 2) & sText      '--- center
        If LenB(At(OperatorData, ucsFscPixOperatorCode)) <> 0 Then
            vElem = Split(m_oFP.SendCommand(ucsFpcInfoOperator, At(OperatorData, ucsFscPixOperatorCode)), ",")
            sText = Format$(At(OperatorData, ucsFscPixOperatorCode), "0000") & " " & At(vElem, 5)
            If LenB(At(OperatorData, ucsFscPixCashDeskNo)) <> 0 Then
                sText = sText & Right$(pvSpace(uCtx.FpCommentLen) & Format$(At(OperatorData, ucsFscPixCashDeskNo), "00"), uCtx.FpCommentLen - Len(sText))
            End If
            m_oFP.SendCommand ucsFpcNonFiscalText, sText
        End If
    End Select
    bInTrans = True
    If IsArray(Rows) Then
        '--- print rows
        For Each vElem In Rows
            lCurrentRow = lCurrentRow + 1
            Select Case ReceiptType
            Case ucsFscRetNonfiscal
                For Each vCmd In pvGetNonFiscalRow(pvGetRowData(vElem), uCtx)
                    Rows(lCurrentRow - 1) = m_oFP.SendCommand(ucsFpcNonFiscalText, C_Str(vCmd))
                Next
            Case ucsFscRetFiscal, ucsFscRetInvoiceFiscal
                For Each vCmd In pvGetFiscalRow(pvGetRowData(vElem), uCtx)
                    Rows(lCurrentRow - 1) = m_oFP.SendCommand(vCmd(0), At(vCmd, 1))
                Next
            Case ucsFscRetInvoiceNonFiscal
                For Each vCmd In pvGetCreditNoteRow(pvGetRowData(vElem), uCtx)
                    Rows(lCurrentRow - 1) = m_oFP.SendCommand(ucsFpcNonFiscalText, C_Str(vCmd))
                Next
            End Select
        Next
        lCurrentRow = 0
    End If
    '--- commit transaction
    Select Case ReceiptType
    Case ucsFscRetNonfiscal
        If uCtx.RowPrinted Then
            '--- print payment
            If Not uCtx.PmtPrinted Then
                For Each vCmd In pvGetNonFiscalRow(pvGetRowData(Array(vbNullString, "P", uCtx.Total)), uCtx)
                    m_oFP.SendCommand ucsFpcNonFiscalText, C_Str(vCmd)
                Next
            End If
            '--- print change
            If Not uCtx.ChangePrinted Then
                For Each vCmd In pvGetNonFiscalRow(pvGetRowData(Array(vbNullString)), uCtx)
                    m_oFP.SendCommand ucsFpcNonFiscalText, C_Str(vCmd)
                Next
            End If
            '--- print custom footer
            m_oFP.Exceptions = False
            sText = m_oFP.SendCommand(ucsFpcInitHeaderFooter, "I6")
            m_oFP.Exceptions = True
            If LenB(sText) <> 0 Then
                m_oFP.SendCommand ucsFpcNonFiscalText, pvSpace((uCtx.FpCommentLen - Len(sText)) \ 2) & sText    '--- center
            End If
            m_oFP.Exceptions = False
            sText = m_oFP.SendCommand(ucsFpcInitHeaderFooter, "I7")
            m_oFP.Exceptions = True
            If LenB(sText) <> 0 Then
                m_oFP.SendCommand ucsFpcNonFiscalText, pvSpace((uCtx.FpCommentLen - Len(sText)) \ 2) & sText    '--- center
            End If
            sText = IIf(uCtx.Count = 1, STR_ARTICLE_SINGLE, uCtx.Count & STR_ARTICLES)
            m_oFP.SendCommand ucsFpcNonFiscalText, pvSpace((uCtx.FpCommentLen - Len(sText)) \ 2) & sText    '--- center
        End If
        '--- closing footer
        m_oFP.SendCommand ucsFpcNonFiscalClose
        '--- open door
        m_oFP.SendCommand ucsFpcInitOpenDrawer
        '--- success
        PrintReceipt = True
    Case ucsFscRetFiscal
        '--- print outstanding payment
        For Each vCmd In pvGetFiscalRow(pvGetRowData(Array(vbNullString)), uCtx)
            If vCmd(0) <> ucsFpcFiscalText Then
                m_oFP.SendCommand vCmd(0), At(vCmd, 1)
            End If
        Next
        m_oFP.SendCommand ucsFpcFiscalClose
        '--- success
        PrintReceipt = True
    Case ucsFscRetInvoiceFiscal
        '--- print outstanding payment
        For Each vCmd In pvGetFiscalRow(pvGetRowData(Array(vbNullString)), uCtx)
            If vCmd(0) <> ucsFpcFiscalText Then
                m_oFP.SendCommand vCmd(0), At(vCmd, 1)
            End If
        Next
        If LenB(At(InvoiceData, ucsFscPixCgBulstat)) <> 0 Or LenB(At(InvoiceData, ucsFscPixCgName)) <> 0 Then
            m_oFP.SendCommand ucsFpcFiscalCgInfo, _
                Left$(At(InvoiceData, ucsFscPixCgBulstat), 13) & vbTab & _
                Left$(At(InvoiceData, ucsFscPixOwnPersonSeller), 30) & vbTab & _
                Left$(At(InvoiceData, ucsFscPixCgPersonReceiver), 25) & vbTab & _
                Left$(At(InvoiceData, ucsFscPixCgName), 27) & vbTab & _
                Left$(At(InvoiceData, ucsFscPixCgVatNo), 12) & vbTab & _
                pvSplitCgAddress(At(InvoiceData, ucsFscPixCgAddress), uCtx)
        End If
        m_oFP.SendCommand ucsFpcFiscalClose
        '--- success
        PrintReceipt = True
    Case ucsFscRetInvoiceNonFiscal
        If uCtx.RowPrinted Then
            '--- print payment
            If Not uCtx.PmtPrinted Then
                For Each vCmd In pvGetCreditNoteRow(pvGetRowData(Array(vbNullString, "P", uCtx.Total)), uCtx)
                    m_oFP.SendCommand ucsFpcNonFiscalText, C_Str(vCmd)
                Next
            End If
            '--- print change
            If Not uCtx.ChangePrinted Then
                For Each vCmd In pvGetCreditNoteRow(pvGetRowData(Array(vbNullString)), uCtx)
                    m_oFP.SendCommand ucsFpcNonFiscalText, C_Str(vCmd)
                Next
            End If
            '--- print credit note info
            sText = STR_SELLER
            m_oFP.SendCommand ucsFpcNonFiscalText, sText & String$(uCtx.FpCommentLen - Len(sText), ".")  '--- pad with dots
            sText = "/" & At(InvoiceData, ucsFscPixOwnPersonSeller) & "/"
            m_oFP.SendCommand ucsFpcNonFiscalText, Right$(pvSpace(uCtx.FpCommentLen) & sText, uCtx.FpCommentLen)   '--- right align
            sText = STR_RECEIVER
            m_oFP.SendCommand ucsFpcNonFiscalText, sText & String$(uCtx.FpCommentLen - Len(sText), ".")  '--- pad with dots
            sText = "/" & At(InvoiceData, ucsFscPixCgPersonReceiver) & "/"
            m_oFP.SendCommand ucsFpcNonFiscalText, Right$(pvSpace(uCtx.FpCommentLen) & sText, uCtx.FpCommentLen)   '--- right align
            sText = STR_BUYER
            For Each vElem In pvWrapText(At(InvoiceData, ucsFscPixCgName), uCtx.FpCommentLen - Len(sText))
                m_oFP.SendCommand ucsFpcNonFiscalText, Left$(sText & vElem, uCtx.FpCommentLen)
                sText = Space$(Len(STR_BUYER))
            Next
            sText = STR_BULSTAT & At(InvoiceData, ucsFscPixCgBulstat)
            m_oFP.SendCommand ucsFpcNonFiscalText, Left$(sText, uCtx.FpCommentLen)
            sText = STR_VATNO & At(InvoiceData, ucsFscPixCgVatNo)
            m_oFP.SendCommand ucsFpcNonFiscalText, Left$(sText, uCtx.FpCommentLen)
            sText = STR_ADDRESS
            For Each vElem In Split(pvSplitCgAddress(At(InvoiceData, ucsFscPixCgAddress), uCtx), vbLf)
                m_oFP.SendCommand ucsFpcNonFiscalText, Left$(sText & vElem, uCtx.FpCommentLen)
                sText = Space$(Len(STR_ADDRESS))
            Next
            m_oFP.SendCommand ucsFpcNonFiscalText, vbNullString
            '--- print custom footer
            m_oFP.Exceptions = False
            sText = m_oFP.SendCommand(ucsFpcInitHeaderFooter, "I6")
            m_oFP.Exceptions = True
            If LenB(sText) <> 0 Then
                m_oFP.SendCommand ucsFpcNonFiscalText, pvSpace((uCtx.FpCommentLen - Len(sText)) \ 2) & sText    '--- center
            End If
            m_oFP.Exceptions = False
            sText = m_oFP.SendCommand(ucsFpcInitHeaderFooter, "I7")
            m_oFP.Exceptions = True
            If LenB(sText) <> 0 Then
                m_oFP.SendCommand ucsFpcNonFiscalText, pvSpace((uCtx.FpCommentLen - Len(sText)) \ 2) & sText    '--- center
            End If
            sText = IIf(uCtx.Count = 1, STR_ARTICLE_SINGLE, uCtx.Count & STR_ARTICLES)
            m_oFP.SendCommand ucsFpcNonFiscalText, pvSpace((uCtx.FpCommentLen - Len(sText)) \ 2) & sText    '--- center
        End If
        '--- closing footer
        m_oFP.SendCommand ucsFpcNonFiscalClose
        '--- open door
        m_oFP.SendCommand ucsFpcInitOpenDrawer
        '--- success
        PrintReceipt = True
    Case Else
        '--- do nothing. moje da se pechata dublikat
    End Select
    bInTrans = False
    If LenB(At(OperatorData, ucsFscPixLastReceipt, vbNullString)) <> 0 Then
        OperatorData(ucsFscPixLastReceipt) = m_oFP.SendCommand(ucsFpcInfoLastReceipt)
    End If
    If LenB(At(OperatorData, ucsFscPixFiscalPayment, vbNullString)) <> 0 Then
        OperatorData(ucsFscPixFiscalPayment) = m_bLogFiscalPayment
    End If
    If PrintDuplicate Then
        Select Case ReceiptType
        Case ucsFscRetNonfiscal, ucsFscRetInvoiceNonFiscal
        Case Else
            m_oFP.SendCommand ucsFpcPrintDuplicate, "1"
            '--- success
            PrintReceipt = True
        End Select
    End If
    '--- disconnect
    Set m_oFP = Nothing
    CommandLog = m_sLog
    Exit Function
EH:
'    Resume
    LastError = Err.Description
    Resume NextLine
NextLine:
    On Error Resume Next
    If bInLogin Then
        LastError = LastError & STR_OR_INCORRECT_LOGIN
        OperatorData(ucsFscPixOperatorPass) = Empty
    End If
    If lCurrentRow > 0 Then
        LastError = STR_ERROR_LINE & lCurrentRow & ": " & LastError
    End If
    If LenB(At(OperatorData, ucsFscPixFiscalPayment, vbNullString)) <> 0 Then
        OperatorData(ucsFscPixFiscalPayment) = m_bLogFiscalPayment
    End If
    If Not m_oFP Is Nothing Then
        If m_oFP.IsConnected Then
            m_oFP.Exceptions = False
            If LenB(At(OperatorData, ucsFscPixLastReceipt, vbNullString)) <> 0 Then
                OperatorData(ucsFscPixLastReceipt) = m_oFP.SendCommand(ucsFpcInfoLastReceipt)
            End If
            If bInTrans Then
                If ReceiptType = ucsFscRetNonfiscal Then
                    m_oFP.SendCommand ucsFpcNonFiscalClose
                Else
                    DebugLog FUNC_NAME, "EH bInTrans And m_oFP.IsConnected, ReceiptType = ucsFscRetFiscal"
                    '--- note: when printing invoice, if no contragent info set then cancel fails!
                    m_oFP.SendCommand ucsFpcFiscalCancel
                    m_oFP.SendCommand ucsFpcFiscalPayment, vbTab
                    m_oFP.SendCommand ucsFpcFiscalCgInfo, "0000000000"
                    m_oFP.SendCommand ucsFpcFiscalClose
                End If
                bInTrans = False
            End If
        End If
        Set m_oFP = Nothing
    End If
    DebugLog FUNC_NAME, "EH LastError=" & LastError & ", m_sLog=" & m_sLog
    CommandLog = m_sLog
End Function

Public Function PrintReport( _
            DeviceString As String, _
            ByVal ReportType As UcsFiscalReportTypeEnum, _
            Optional ReportData As Variant, _
            Optional LastError As String, _
            Optional CommandLog As String) As Boolean
    Const FUNC_NAME     As String = "PrintReport"
    Dim eCmd            As UcsIclCommandsEnum
    Dim vResult         As Variant
    Dim bIsClear        As Boolean
    Dim lIdx            As Long
    
    On Error GoTo EH
    '--- init & localize fiscal printer
    pvInit DeviceString
    pvReset
    Select Case ReportType
    Case ucsFscRptDaily
        '-- ReportData = { IsClear, IsItems, IsDepartments }
        bIsClear = C_Bool(At(ReportData, 0))
        If C_Bool(At(ReportData, 1)) And C_Bool(At(ReportData, 2)) Then
            eCmd = ucsFpcPrintDailyReportItemsDepartments
        ElseIf C_Bool(At(ReportData, 1)) Then
            eCmd = ucsFpcPrintDailyReportItems
        ElseIf C_Bool(At(ReportData, 2)) Then
            eCmd = ucsFpcPrintDailyReportDepartments
        Else
            eCmd = ucsFpcPrintDailyReport
        End If
        m_oFP.Exceptions = False
        '--- zachistwane na elektronna kontrolna lenta
        If bIsClear Then
            vResult = Split(m_oFP.SendCommand(ucsFpcInitEcTape, "I"), ",")
            '--- print
            For lIdx = 1 To C_Lng(At(vResult, 1))
                m_oFP.SendCommand ucsFpcInitEcTape, IIf(lIdx = 1, "PS", "CS")
            Next
            '--- erase
            m_oFP.SendCommand ucsFpcInitEcTape, "E"
        End If
        vResult = m_oFP.SendCommand(eCmd, IIf(bIsClear, "0", "2") & "N")
        If m_oFP.Status(ucsStbPrintingError) Then
            '--- daisy: pechat po depatamenti
            If eCmd = ucsFpcPrintDailyReportDepartments Then
                vResult = m_oFP.SendCommand(ucsFpcPrintDailyReport, IIf(bIsClear, "8", "9") & "N")
            ElseIf eCmd = ucsFpcPrintDailyReportItemsDepartments Then
                vResult = m_oFP.SendCommand(ucsFpcPrintDailyReportItems, IIf(bIsClear, "8", "9") & "N")
            End If
        End If
        m_oFP.Exceptions = True
    Case ucsFscRptNumber
        '-- ReportData = { FromNum, ToNum, IsDetailed }
        If C_Bool(At(ReportData, 2)) Then
            eCmd = ucsFpcPrintReportByNumberDetailed
        Else
            eCmd = ucsFpcPrintReportByNumberShort
        End If
        vResult = m_oFP.SendCommand(eCmd, At(ReportData, 0) & "," & At(ReportData, 1))
    Case ucsFscRptDate
        '-- ReportData = { FromDate, ToDate, IsDetailed }
        If C_Bool(At(ReportData, 2)) Then
            eCmd = ucsFpcPrintReportByDateDetailed
        Else
            eCmd = ucsFpcPrintReportByDateShort
        End If
        vResult = m_oFP.SendCommand(eCmd, Format$(C_Date(At(ReportData, 0)), "ddmmyy") & "," & Format$(C_Date(At(ReportData, 1)), "ddmmyy"))
    Case ucsFscRptOperator
        '-- ReportData = Empty
        vResult = m_oFP.SendCommand(ucsFpcPrintReportByOperators)
    End Select
    ReportData = vResult
    '--- disconnect
    Set m_oFP = Nothing
    CommandLog = m_sLog
    '--- success
    PrintReport = True
    Exit Function
EH:
    LastError = Err.Description
    Resume NextLine
NextLine:
    On Error Resume Next
    Set m_oFP = Nothing
    DebugLog FUNC_NAME, "EH LastError=" & LastError & ", m_sLog=" & m_sLog
    CommandLog = m_sLog
End Function

Public Function OpenDrawer( _
            DeviceString As String, _
            Optional LastError As String, _
            Optional CommandLog As String) As Boolean
    Const FUNC_NAME     As String = "OpenDrawer"
    
    On Error GoTo EH
    '--- init & localize fiscal printer
    pvInit DeviceString
    pvReset
    m_oFP.SendCommand ucsFpcInitOpenDrawer
    '--- disconnect
    Set m_oFP = Nothing
    CommandLog = m_sLog
    '--- success
    OpenDrawer = True
    Exit Function
EH:
    LastError = Err.Description
    Resume NextLine
NextLine:
    On Error Resume Next
    Set m_oFP = Nothing
    DebugLog FUNC_NAME, "EH LastError=" & LastError & ", m_sLog=" & m_sLog
    CommandLog = m_sLog
End Function

Public Function InitDateTime( _
            DeviceString As String, _
            sDateTime As String, _
            Optional LastError As String, _
            Optional CommandLog As String) As Boolean
    Const FUNC_NAME     As String = "InitDateTime"
    
    On Error GoTo EH
    '--- init & localize fiscal printer
    pvInit DeviceString
    m_oFP.Exceptions = False
    pvReset
    If LenB(sDateTime) = 0 Then
        sDateTime = Format$(Now, "dd-MM-yy") & " " & Format$(Now, "hh:mm:ss")
    End If
    m_oFP.SendCommand ucsFpcInitDateTime, sDateTime
    sDateTime = m_oFP.SendCommand(ucsFpcInfoDateTime)
    '--- disconnect
    Set m_oFP = Nothing
    CommandLog = m_sLog
    '--- success
    InitDateTime = True
    Exit Function
EH:
    LastError = Err.Description
    Resume NextLine
NextLine:
    On Error Resume Next
    Set m_oFP = Nothing
    DebugLog FUNC_NAME, "EH LastError=" & LastError & ", m_sLog=" & m_sLog
    CommandLog = m_sLog
End Function

'= private ===============================================================

Private Sub pvInit(DeviceString As String)
    Dim vSplit          As Variant
    
    m_sLog = vbNullString
    m_bLogFiscalPayment = False
    Set m_oFP = New cICLProtocol
    m_oFP.LocalizedText(ucsFscLciInternalErrors) = STR_INTERNAL
    m_oFP.LocalizedText(ucsFscLciPrinterStatuses) = STR_STATUSES
    m_oFP.LocalizedText(ucsFscLciPrinterErrors) = STR_ERRORS_ONLY
    m_oFP.LocalizedText(ucsFscLciDipSwitchesTexts) = STR_DIP_SWITCHES
    m_oFP.LocalizedText(ucsFscLciMemoryTexts) = STR_MEMORY
    vSplit = Split(DeviceString, ";")
    m_oFP.Init At(vSplit, 1), C_Lng(At(vSplit, 2, 3000))
End Sub

Private Sub pvReset()
    Dim bExceptions     As Boolean
    
    bExceptions = m_oFP.Exceptions
    m_oFP.Exceptions = False
    If Left$(m_oFP.SendCommand(ucsFpcInfoTransaction), 1) = "1" Then
        If m_oFP.Status(ucsStbFiscalPrinting) Then
            '--- note: FP3530 moje da anulira winagi, FP550F ne moje
            m_oFP.SendCommand ucsFpcFiscalCancel
            If m_oFP.Status(ucsStbFiscalPrinting) Then
                '--- note: ako platenata suma e po-malka ot sumata na bona, towa pomaga
                m_oFP.SendCommand ucsFpcFiscalPayment, vbTab
                '--- note: when printing invoice, if no contragent info set then cancel fails!
                m_oFP.SendCommand ucsFpcFiscalCgInfo, "0000000000"
                '--- zaradi FP550F
                m_oFP.SendCommand ucsFpcFiscalClose
            End If
        Else
            m_oFP.SendCommand ucsFpcNonFiscalClose
        End If
    End If
    m_oFP.Exceptions = bExceptions
End Sub

Private Function pvGetRowData(vRow As Variant) As UcsRow
    On Error Resume Next
    If UBound(vRow) = ucsFscPixItem Then
        pvGetRowData.Type = ucsRwtText
        pvGetRowData.Item = C_Str(vRow(ucsFscPixItem))
    ElseIf UBound(vRow) = ucsFscPixAmount Then
        pvGetRowData.Type = ucsRwtPayment
        pvGetRowData.Item = C_Str(vRow(ucsFscPixItem))
        pvGetRowData.WayOfPayment = At(Array("P", "D", "C", "N"), C_Lng(vRow(ucsFscPixWayOfPmt)) - 1, "P")
        pvGetRowData.Amount = C_Str(vRow(ucsFscPixAmount))
    Else
        pvGetRowData.Type = ucsRwtSell
        pvGetRowData.Item = C_Str(vRow(ucsFscPixItem))
        pvGetRowData.Price = C_Str(vRow(ucsFscPixPrice))
        pvGetRowData.Group = C_Str(vRow(ucsFscPixGroup))
        pvGetRowData.Quantity = C_Str(vRow(ucsFscPixQuantity))
        pvGetRowData.Discount = At(vRow, ucsFscPixDiscount)
    End If
    On Error GoTo 0
End Function

Private Function pvGetCreditNoteRow(uRow As UcsRow, uCtx As UcsContext) As Variant
    Const FUNC_NAME     As String = "pvGetCreditNoteRow"
    Dim lIdx            As Long
    Dim sItem           As String
    Dim sQuantity       As String
    Dim sPrice          As String
    Dim sTotal          As String
    Dim sDiscount       As String
    Dim vRet            As Variant
    Dim sWayOfPmt       As String
    Dim lGroup          As Long
    Dim vPecents        As Variant
    Dim dblNeto         As Double
    Dim dblTotal        As Double
    Dim dblQuantity     As Double
    
    On Error GoTo EH
    With uRow
        If .Type = ucsRwtSell Then
            If uCtx.ChangePrinted Or uCtx.PmtPrinted Then
                On Error GoTo 0
                Err.Raise vbObjectError, , m_oFP.frGetFlagsText(ucsStbInvalidCommand, m_oFP.LocalizedText(ucsFscLciPrinterStatuses))
            End If
            uCtx.RowPrinted = True
            sPrice = pvSafeFormat(Round(C_Dbl(.Price), 2), "0.00", ",")
            If LenB(.Group) <> 0 Then
                lGroup = Asc(UCase(.Group)) - Asc("А")
            Else
                lGroup = 1
            End If
            If LenB(.Quantity) <> 0 And C_Dbl(.Quantity) <> 1 Then
                dblQuantity = Round(C_Dbl(.Quantity), 3)
            Else
                dblQuantity = 1
            End If
            sQuantity = pvSafeFormat(dblQuantity, "0.000", ",")
            sTotal = pvSafeFormat(pvCalcRow(dblQuantity, C_Dbl(.Price)), "0.00", ",") & " " & .Group
            sDiscount = pvSafeFormat(-pvCalcDiscount(dblQuantity, C_Dbl(.Price), C_Dbl(.Discount)), "0.00", ",") & " " & .Group
            dblTotal = pvCalcRow(dblQuantity, C_Dbl(.Price), C_Dbl(.Discount))
            uCtx.Total = Round(uCtx.Total + dblTotal, 2)
            uCtx.GrpTotal(lGroup) = Round(uCtx.GrpTotal(lGroup) + dblTotal, 2)
            vRet = pvWrapText(RTrim(pvSafeText(.Item)), uCtx.FpCommentLen)
            ReDim Preserve vRet(0 To UBound(vRet) + 3) As Variant
            vRet(UBound(vRet) - 2) = STR_UNIT_PRICE & Right$(pvSpace(uCtx.FpCommentLen) & sPrice, uCtx.FpCommentLen - 2 - Len(STR_UNIT_PRICE))
            vRet(UBound(vRet) - 1) = STR_UNIT_QUO & Right$(pvSpace(uCtx.FpCommentLen) & sQuantity, uCtx.FpCommentLen - 2 - Len(STR_UNIT_QUO))
            vRet(UBound(vRet) - 0) = STR_UNIT_SUM & Right$(pvSpace(uCtx.FpCommentLen) & sTotal, uCtx.FpCommentLen - Len(STR_UNIT_SUM))
            If LenB(.Discount) <> 0 Then
                ReDim Preserve vRet(0 To UBound(vRet) + 1) As Variant
                sItem = STR_DISCOUNT & pvSafeFormat(Round(C_Dbl(.Discount), 2), vbNullString, ",") & "%"
                vRet(UBound(vRet)) = sItem & Right$(pvSpace(uCtx.FpCommentLen) & sDiscount, uCtx.FpCommentLen - Len(sItem))
            End If
            uCtx.Count = uCtx.Count + 1
        ElseIf .Type = ucsRwtPayment Then
            If uCtx.ChangePrinted Then
                On Error GoTo 0
                Err.Raise vbObjectError, , m_oFP.frGetFlagsText(ucsStbInvalidCommand, m_oFP.LocalizedText(ucsFscLciPrinterStatuses))
            End If
            If Not uCtx.PmtPrinted Then
                uCtx.PmtPrinted = True
                ReDim vRet(0 To 1) As Variant
                vRet(0) = String$(uCtx.FpCommentLen, "-")
                vRet(1) = STR_TOTAL & Right$(pvSpace(uCtx.FpCommentLen) & pvSafeFormat(uCtx.Total, "0.00", ","), uCtx.FpCommentLen - Len(STR_TOTAL) - 2)
                vPecents = Split(m_oFP.SendCommand(ucsFpcInfoTaxRates), ",")
                For lIdx = 0 To UBound(vPecents)
                    If lIdx <= UBound(uCtx.GrpTotal) Then
                    If uCtx.GrpTotal(lIdx) <> 0 Then
                        dblNeto = Round(uCtx.GrpTotal(lIdx) / (1 + C_Dbl(At(vPecents, lIdx)) / 100#), 2)
                        ReDim Preserve vRet(0 To UBound(vRet) + 3) As Variant
                        sItem = STR_GROUP_TOTAL & Chr$(Asc("А") + lIdx)
                        vRet(UBound(vRet) - 2) = sItem & Right$(pvSpace(uCtx.FpCommentLen) & pvSafeFormat(uCtx.GrpTotal(lIdx), "0.00", ","), uCtx.FpCommentLen - Len(sItem) - 2)
                        sItem = STR_GROUP_VAT & Chr$(Asc("А") + lIdx) & "=" & pvSafeFormat(C_Dbl(At(vPecents, lIdx)), "0.00", ",") & "%"
                        vRet(UBound(vRet) - 1) = sItem & Right$(pvSpace(uCtx.FpCommentLen) & pvSafeFormat(uCtx.GrpTotal(lIdx) - dblNeto, "0.00", ","), uCtx.FpCommentLen - Len(sItem) - 2)
                        sItem = STR_GROUP_NETO
                        vRet(UBound(vRet) - 0) = sItem & Right$(pvSpace(uCtx.FpCommentLen) & pvSafeFormat(dblNeto, "0.00", ","), uCtx.FpCommentLen - Len(sItem) - 2)
                    End If
                    End If
                Next
                ReDim Preserve vRet(0 To UBound(vRet) + 2) As Variant
                vRet(UBound(vRet) - 1) = Replace(String$(uCtx.FpCommentLen / 2, "-"), "-", "- ")
            Else
                ReDim vRet(0 To 0) As Variant
            End If
            sWayOfPmt = IIf(.WayOfPayment = "D", STR_CARD, STR_CASH)
            vRet(UBound(vRet)) = sWayOfPmt & Right$(pvSpace(uCtx.FpCommentLen) & pvSafeFormat(C_Dbl(.Amount), "0.00", ","), uCtx.FpCommentLen - Len(sWayOfPmt) - 2)
            uCtx.Amount = Round(uCtx.Amount + Round(C_Dbl(.Amount), 2), 2)
        ElseIf .Type = ucsRwtText Then
            vRet = pvWrapText(RTrim(pvSafeText(.Item)), uCtx.FpCommentLen)
            If Not uCtx.ChangePrinted And uCtx.PmtPrinted Then
                uCtx.ChangePrinted = True
                If Round(uCtx.Amount, 2) > Round(uCtx.Total, 2) Then
                    ReDim Preserve vRet(0 To UBound(vRet) + 1) As Variant
                    For lIdx = UBound(vRet) To 1 Step -1
                        vRet(lIdx) = vRet(lIdx - 1)
                    Next
                    vRet(0) = STR_CHANGE & Right$(pvSpace(uCtx.FpCommentLen) & pvSafeFormat(uCtx.Amount - uCtx.Total, "0.00", ","), uCtx.FpCommentLen - Len(STR_CHANGE) - 2)
                End If
            End If
        End If
    End With
    pvGetCreditNoteRow = vRet
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function pvGetNonFiscalRow(uRow As UcsRow, uCtx As UcsContext) As Variant
    Const FUNC_NAME     As String = "pvGetNonFiscalRow"
    Dim lIdx            As Long
    Dim sItem           As String
    Dim sQuantity       As String
    Dim sPrice          As String
    Dim sDiscount       As String
    Dim vSplit          As Variant
    Dim vRet            As Variant
    Dim sWayOfPmt       As String
    Dim dblQuantity     As Double
    Dim dblTotal        As Double
    
    On Error GoTo EH
    With uRow
        If .Type = ucsRwtSell Then
            If uCtx.ChangePrinted Or uCtx.PmtPrinted Then
                On Error GoTo 0
                Err.Raise vbObjectError, , m_oFP.frGetFlagsText(ucsStbInvalidCommand, m_oFP.LocalizedText(ucsFscLciPrinterStatuses))
            End If
            uCtx.RowPrinted = True
            sPrice = pvSafeFormat(Round(C_Dbl(.Price), 2), "0.00", ",")
            If LenB(.Quantity) <> 0 And C_Dbl(.Quantity) <> 1 Then
                dblQuantity = Round(C_Dbl(.Quantity), 3)
                sQuantity = Right$(" " & pvSafeFormat(dblQuantity, "0.0##", ",") & " x " & pvSpace(7 - Len(sPrice)) & sPrice & "  ", uCtx.FpCommentLen)
            Else
                dblQuantity = 1
            End If
            sPrice = pvSafeFormat(pvCalcRow(dblQuantity, C_Dbl(.Price)), "0.00", ",") & " " & .Group
            sDiscount = pvSafeFormat(-pvCalcDiscount(dblQuantity, C_Dbl(.Price), C_Dbl(.Discount)), "0.00", ",") & " " & .Group
            dblTotal = pvCalcRow(dblQuantity, C_Dbl(.Price), C_Dbl(.Discount))
            uCtx.Total = Round(uCtx.Total + dblTotal, 2)
            '--- split long item name
            sItem = RTrim(pvSafeText(.Item))
            ReDim vSplit(-1 To -1) As Variant
            If Len(sItem) > uCtx.FpCommentLen - Len(sPrice) - 1 Then
                ReDim Preserve vSplit(0 To UBound(vSplit) + 1) As Variant
                vSplit(UBound(vSplit)) = pvChopOff(sItem, uCtx.FpCommentLen - Len(sPrice) - 1)
            End If
            If Len(sItem) > uCtx.FpCommentLen - Len(sQuantity) - 1 Then
                ReDim Preserve vSplit(0 To UBound(vSplit) + 1) As Variant
                vSplit(UBound(vSplit)) = pvChopOff(sItem, uCtx.FpCommentLen - Len(sQuantity) - 1)
            End If
            Do While Len(sItem) > uCtx.FpCommentLen
                ReDim Preserve vSplit(0 To UBound(vSplit) + 1) As Variant
                vSplit(UBound(vSplit)) = pvChopOff(sItem, uCtx.FpCommentLen)
            Loop
            If LenB(sItem) <> 0 Then
                ReDim Preserve vSplit(0 To UBound(vSplit) + 1) As Variant
                vSplit(UBound(vSplit)) = sItem
            End If
            '--- append price & quantity
            If LenB(sPrice) <> 0 Then
                If UBound(vSplit) < 0 Then
                    ReDim Preserve vSplit(0 To 0) As Variant
                End If
                vSplit(0) = vSplit(0) & pvSpace(uCtx.FpCommentLen - Len(sPrice) - Len(vSplit(0))) & sPrice
            End If
            If LenB(sQuantity) <> 0 Then
                If UBound(vSplit) < 1 Then
                    ReDim Preserve vSplit(0 To 1) As Variant
                End If
                vSplit(1) = vSplit(1) & pvSpace(uCtx.FpCommentLen - Len(sQuantity) - Len(vSplit(1))) & sQuantity
            End If
            '--- retval
            ReDim vRet(0 To UBound(vSplit)) As Variant
            For lIdx = 0 To UBound(vRet)
                vRet(lIdx) = vSplit(UBound(vSplit) - lIdx)
            Next
            '--- append discount
            If C_Dbl(.Discount) <> 0 Then
                ReDim Preserve vRet(0 To UBound(vRet) + 1) As Variant
                sItem = STR_DISCOUNT & pvSafeFormat(Round(C_Dbl(.Discount), 2), vbNullString, ",") & "%"
                vRet(UBound(vRet)) = sItem & pvSpace(uCtx.FpCommentLen - Len(sItem) - Len(sDiscount)) & sDiscount
            End If
            uCtx.Count = uCtx.Count + 1
        ElseIf .Type = ucsRwtPayment Then
            If uCtx.ChangePrinted Then
                On Error GoTo 0
                Err.Raise vbObjectError, , m_oFP.frGetFlagsText(ucsStbInvalidCommand, m_oFP.LocalizedText(ucsFscLciPrinterStatuses))
            End If
            If Not uCtx.PmtPrinted Then
                uCtx.PmtPrinted = True
                ReDim vRet(0 To 2) As Variant
                vRet(0) = String$(uCtx.FpCommentLen, "-")
                vRet(1) = STR_TOTAL & Right$(pvSpace(uCtx.FpCommentLen) & pvSafeFormat(uCtx.Total, "0.00", ","), uCtx.FpCommentLen - Len(STR_TOTAL) - 2)
            Else
                ReDim vRet(0 To 0) As Variant
            End If
            sWayOfPmt = IIf(.WayOfPayment = "D", STR_CARD, STR_CASH)
            vRet(UBound(vRet)) = sWayOfPmt & Right$(pvSpace(uCtx.FpCommentLen) & pvSafeFormat(C_Dbl(.Amount), "0.00", ","), uCtx.FpCommentLen - Len(sWayOfPmt) - 2)
            uCtx.Amount = Round(uCtx.Amount + Round(C_Dbl(.Amount), 2), 2)
        ElseIf .Type = ucsRwtText Then
            vRet = pvWrapText(RTrim(pvSafeText(.Item)), uCtx.FpCommentLen)
            If Not uCtx.ChangePrinted And uCtx.PmtPrinted Then
                uCtx.ChangePrinted = True
                If Round(uCtx.Amount, 2) > Round(uCtx.Total, 2) Then
                    ReDim Preserve vRet(0 To UBound(vRet) + 1) As Variant
                    For lIdx = UBound(vRet) To 1 Step -1
                        vRet(lIdx) = vRet(lIdx - 1)
                    Next
                    vRet(0) = STR_CHANGE & Right$(pvSpace(uCtx.FpCommentLen) & pvSafeFormat(uCtx.Amount - uCtx.Total, "0.00", ","), uCtx.FpCommentLen - Len(STR_CHANGE) - 2)
                End If
            End If
        End If
    End With
    pvGetNonFiscalRow = vRet
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function pvGetFiscalRow(uRow As UcsRow, uCtx As UcsContext) As Variant
    Const FUNC_NAME     As String = "pvGetFiscalRow"
    Dim lIdx            As Long
    Dim sSign           As String
    Dim sText           As String
    Dim vSplit          As Variant
    Dim vRet            As Variant
    Dim dblPrice        As Double
    Dim dblQuantity     As Double
    Dim dblDiscount     As Double
    Dim dblAbsPrice     As Double
    Dim dblAbsQuantity  As Double
    Dim lItemWidth      As Long
    
    On Error GoTo EH
    lItemWidth = uCtx.FpCommentLen - 5
    With uRow
        If .Type = ucsRwtSell Then
            If uCtx.ChangePrinted Or uCtx.PmtPrinted Then
                On Error GoTo 0
                Err.Raise vbObjectError, , m_oFP.frGetFlagsText(ucsStbInvalidCommand, m_oFP.LocalizedText(ucsFscLciPrinterStatuses))
            End If
            uCtx.RowPrinted = True
            ReDim vRet(0 To 0) As Variant
            '--- split item
            vSplit = pvWrapText(RTrim(pvSafeText(.Item)), lItemWidth)
            If UBound(vSplit) = 0 Then
                sText = vSplit(0)
            Else
                sText = vSplit(UBound(vSplit) - 1) & vbLf & vSplit(UBound(vSplit))
            End If
            '--- long item name -> prepend ucsFpcFiscalText
            If UBound(vSplit) > 1 Then
                ReDim vRet(0 To UBound(vSplit) - 1) As Variant
                For lIdx = 0 To UBound(vRet) - 1
                    vRet(lIdx) = Array(ucsFpcFiscalText, vSplit(lIdx))
                Next
            End If
            dblPrice = Round(C_Dbl(.Price), 2)
            dblQuantity = Round(C_Dbl(.Quantity), 3)
            dblDiscount = Round(C_Dbl(.Discount), 2)
            dblAbsPrice = Round(Abs(C_Dbl(.Price)), 2)
            dblAbsQuantity = Round(Abs(C_Dbl(.Quantity)), 3)
            '--- figure out sign
            If LenB(.Quantity) <> 0 Then
                sSign = IIf(pvCalcRow(dblQuantity, dblPrice) < 0, "-", vbNullString)
            Else
                sSign = IIf(pvCalcRow(1, dblPrice) < 0, "-", vbNullString)
            End If
            '--- append rest
            sText = sText & vbTab & .Group & sSign & pvSafeFormat(dblAbsPrice, "0.00")
            If LenB(.Quantity) <> 0 Then
                sText = sText & "*" & pvSafeFormat(dblAbsQuantity, "0.000")
                uCtx.Total = Round(uCtx.Total + pvCalcRow(dblQuantity, dblPrice, dblDiscount), 2)
            Else
                uCtx.Total = Round(uCtx.Total + pvCalcRow(1, dblPrice, dblDiscount), 2)
            End If
            If LenB(.Discount) <> 0 Then
                sText = sText & "," & pvSafeFormat(-dblDiscount, "0.00")
            End If
            vRet(UBound(vRet)) = Array(ucsFpcFiscalSell, sText)
            pvGetFiscalRow = vRet
            uCtx.Count = uCtx.Count + 1
        ElseIf .Type = ucsRwtPayment Then
            If uCtx.ChangePrinted Then
                On Error GoTo 0
                Err.Raise vbObjectError, , m_oFP.frGetFlagsText(ucsStbInvalidCommand, m_oFP.LocalizedText(ucsFscLciPrinterStatuses))
            End If
            uCtx.PmtPrinted = True
            If LenB(.Amount) = 0 Then
                uCtx.Amount = uCtx.Total
            Else
                uCtx.Amount = Round(uCtx.Amount + Round(Abs(C_Dbl(.Amount)), 2), 2)
            End If
            ReDim vRet(0 To 0) As Variant
            '--- split item
            vSplit = pvWrapText(RTrim(pvSafeText(.Item)), lItemWidth)
            If UBound(vSplit) = 0 Then
                sText = vSplit(0)
            Else
                sText = vSplit(UBound(vSplit) - 1) & vbLf & vSplit(UBound(vSplit))
            End If
            '--- long item name -> prepend ucsFpcFiscalText
            If UBound(vSplit) > 1 Then
                ReDim vRet(0 To UBound(vSplit) - 1) As Variant
                For lIdx = 0 To UBound(vRet) - 1
                    vRet(lIdx) = Array(ucsFpcFiscalText, vSplit(lIdx))
                Next
            End If
            '--- append rest
            sText = sText & vbTab
            If LenB(.Amount) <> 0 Then
                sText = sText & .WayOfPayment & pvSafeFormat(Round(Abs(C_Dbl(.Amount)), 2), "0.00")
            End If
            vRet(UBound(vRet)) = Array(ucsFpcFiscalPayment, sText)
        ElseIf .Type = ucsRwtText Then
            '--- split item
            vSplit = pvWrapText(RTrim(pvSafeText(.Item)), uCtx.FpCommentLen)
            ReDim vRet(0 To UBound(vSplit)) As Variant
            For lIdx = 0 To UBound(vRet)
                vRet(lIdx) = Array(ucsFpcFiscalText, vSplit(lIdx))
            Next
            '--- finalizirane
            If LenB(.Item) = 0 Then
                If Not uCtx.PmtPrinted Or Round(uCtx.Amount, 2) < Round(uCtx.Total, 2) Then
                    uCtx.PmtPrinted = True
                    uCtx.ChangePrinted = True
                    uCtx.Amount = uCtx.Total
                    ReDim Preserve vRet(0 To UBound(vRet) + 1) As Variant
                    For lIdx = UBound(vRet) To 1 Step -1
                        vRet(lIdx) = vRet(lIdx - 1)
                    Next
                    vRet(0) = Array(ucsFpcFiscalPayment, vbTab)
                End If
            End If
        End If
    End With
    pvGetFiscalRow = vRet
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function pvChopOff(sItem As String, ByVal lSize As Long) As String
    Dim lIdx            As Long
    
    lIdx = Len(pvWrapText(StrReverse(sItem), lSize)(0))
    pvChopOff = Right$(sItem, lIdx)
    For lIdx = Len(sItem) - lIdx To 0 Step -1
        If Not pvIsWhiteSpace(Mid$(sItem, lIdx, 1)) Then
            Exit For
        End If
    Next
    sItem = Left$(sItem, lIdx)
End Function

Private Function pvSplitCgAddress(ByVal sAddress As String, uCtx As UcsContext) As String
    Const FUNC_NAME     As String = "pvSplitCgAddress"
    Dim vSplit          As Variant
    Dim vElem           As Variant
    Dim sRow1           As String
    Dim sRow2           As String
    Dim sTemp           As String
    
    On Error GoTo EH
    vSplit = Split(sAddress, vbCrLf)
    If UBound(vSplit) <= 1 Then
        On Error Resume Next
        sRow1 = vSplit(0)
        sRow2 = vSplit(1)
        On Error GoTo 0
    Else
        For Each vElem In vSplit
            If LenB(vElem) <> 0 Then
                If LenB(sRow2) <> 0 Then
                    sRow2 = sRow2 & ";" & vElem
                Else
                    sTemp = sRow1 & IIf(LenB(sRow1) <> 0, ";", vbNullString) & vElem
                    If Len(sTemp) > uCtx.FpCommentLen Then
                        sRow2 = vElem
                    Else
                        sRow1 = sTemp
                    End If
                End If
            End If
        Next
    End If
    pvSplitCgAddress = Left$(sRow1, uCtx.FpCommentLen) & IIf(LenB(sRow2) <> 0, vbLf, vbNullString) & Left$(sRow2, uCtx.FpCommentLen)
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function pvWrapText(ByVal sText As String, ByVal lWidth As Long) As Variant
    Dim lRight          As Long
    Dim lLeft           As Long
    Dim vRet            As Variant
    Dim lCount          As Long
    
    ReDim vRet(0 To Len(sText)) As Variant
    Do While LenB(sText) <> 0
        lRight = lWidth + 1
        If lRight > Len(sText) Then
            lRight = Len(sText) + 1
        Else
            If pvIsDelimiter(Mid$(sText, lRight, 1)) Then
                Do While pvIsWhiteSpace(Mid$(sText, lRight, 1)) And lRight <= Len(sText)
                    lRight = lRight + 1
                Loop
            Else
                Do While lRight > 1
                    If pvIsDelimiter(Mid$(sText, lRight - 1, 1)) Then
                        Exit Do
                    End If
                    lRight = lRight - 1
                Loop
                If lRight = 1 Then
                    lRight = lWidth + 1
                End If
            End If
        End If
        lLeft = lRight - 1
        Do While pvIsWhiteSpace(Mid$(sText, lLeft, 1)) And lLeft > 0
            lLeft = lLeft - 1
        Loop
        vRet(lCount) = Left$(sText, lLeft)
        lCount = lCount + 1
        sText = Mid$(sText, lRight)
    Loop
    If lCount = 0 Then
        pvWrapText = Array(vbNullString)
    Else
        ReDim Preserve vRet(0 To lCount - 1) As Variant
        pvWrapText = vRet
    End If
End Function

Private Function pvIsDelimiter(sText As String) As Boolean
    Const STR_DELIMS As String = "~#$^&*_+-=\|/ " & vbTab & vbCrLf
    If InStr(1, STR_DELIMS, Left$(sText, 1)) > 0 Then
        pvIsDelimiter = True
    End If
End Function

Private Function pvIsWhiteSpace(sText As String) As Boolean
    Const STR_WHITESPACE As String = " " & vbTab & vbCrLf
    If InStr(1, STR_WHITESPACE, Left$(sText, 1)) > 0 Then
        pvIsWhiteSpace = True
    End If
End Function

Private Function pvSafeText(sText As String) As String
    Dim lIdx            As Long
    
    pvSafeText = sText
    For lIdx = 0 To 31
        pvSafeText = Replace(pvSafeText, Chr$(lIdx), vbNullString)
    Next
End Function

Private Function pvSafeFormat(Expression As Variant, Optional Fmt As Variant, Optional sDecimal As String = ".") As String
    pvSafeFormat = Replace(Format$(Expression, Fmt), m_sDecimalSeparator, sDecimal)
End Function

Private Function C_Dbl(v As Variant) As Double
    On Error Resume Next
    C_Dbl = CDbl(Replace(C_Str(v), ".", m_sDecimalSeparator))
    On Error GoTo 0
End Function

Private Function pvGetDecimalSeparator() As String
    Dim sBuffer         As String
    Dim nSize           As Long

    sBuffer = Space$(100)
    nSize = GetLocaleInfo(GetUserDefaultLCID(), LOCALE_SDECIMAL, sBuffer, Len(sBuffer))
    If nSize > 0 Then
        pvGetDecimalSeparator = Left$(sBuffer, nSize - 1)
    Else
        pvGetDecimalSeparator = "."
    End If
End Function

Private Function pvSpace(ByVal lSize As Long) As String
    If lSize > 0 Then
        pvSpace = Space$(lSize)
    End If
End Function

Private Function pvCalcRow(dblQuantity As Double, dblPrice As Double, Optional dblDiscount As Double) As Double
    pvCalcRow = Round(Round(dblQuantity, 3) * Round(dblPrice, 2), 2)
    pvCalcRow = Round(pvCalcRow - Round(pvCalcRow * Round(dblDiscount, 2) / 100#, 2), 2)
End Function

Private Function pvCalcDiscount(dblQuantity As Double, dblPrice As Double, dblDiscount As Double) As Double
    pvCalcDiscount = Round(Round(dblQuantity, 3) * Round(dblPrice, 2) * Round(dblDiscount, 2) / 100#, 2)
End Function

Private Sub Class_Initialize()
    m_sDecimalSeparator = pvGetDecimalSeparator()
End Sub

'=========================================================================
' Event handlers
'=========================================================================

Private Sub m_oFP_CommandComplete(ByVal lCmd As Long, sData As String, sResult As String)
    Const FUNC_NAME     As String = "m_oFP_CommandComplete"
    
    On Error GoTo EH
    If lCmd = ucsFpcFiscalPayment Then
        m_bLogFiscalPayment = True
    End If
    m_sLog = m_sLog & lCmd & IIf(LenB(sData) <> 0, "<-" & sData, vbNullString) & IIf(LenB(sResult) <> 0, "->" & sResult, vbNullString) & vbCrLf
    If LenB(m_oFP.LastError) <> 0 Then
        m_sLog = m_sLog & m_oFP.LastError & vbCrLf
    End If
    If m_oFP.Status(ucsStbPrintingError) Then
        m_sLog = m_sLog & m_oFP.StatusText & vbCrLf ' & m_oFP.DipText & vbCrLf & m_oFP.MemoryText & vbCrLf
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub
