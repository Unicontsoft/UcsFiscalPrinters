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
' $Header: /UcsFiscalPrinter/Src/cFiscalPrinter.cls 14    2.12.08 0:15 Wqw $
'
'   Unicontsoft Fiscal Printers Project
'   Copyright (c) 2008 Unicontsoft
'
'   Interfacen class za uprawlenie na fiscalni printer
'
' $Log: /UcsFiscalPrinter/Src/cFiscalPrinter.cls $
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
' ADD: Function PrintReport, OpenDoor. REF: GetHeaderFooter wryshta
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
Private Const MODULE_NAME As String = "cFiscalPrinter"

'=========================================================================
' API
'=========================================================================

Private Const LOCALE_SDECIMAL                       As Long = &HE   ' decimal separator

Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Declare Function GetUserDefaultLCID Lib "kernel32" () As Long

'=========================================================================
' Public enums
'=========================================================================

Public Enum UcsFiscalPrinterReceiptTypeEnum
    ucsFptNonFiscal = 1
    ucsFptFiscal
    ucsFptInvoice
End Enum

Public Enum UcsFiscalPrinterReportTypeEnum
    ucsFprDaily = 1
    ucsFprNumber
    ucsFprDate
    ucsFprOperator
End Enum

Public Enum UcsFiscalPrinterIndexesEnum
    ucsFpiItem = 0
    ucsFpiPrice
    ucsFpiGroup
    ucsFpiQuantity
    ucsFpiWayOfPmt = ucsFpiPrice
    ucsFpiAmount
    ucsFpiInvoiceNo = 0
    ucsFpiCgName
    ucsFpiCgBulstat
    ucsFpiCgVatNo
    ucsFpiCgAddress
    ucsFpiCgPersonReceiver
    ucsFpiOwnPersonSeller
    ucsFpiOperatorCode = 0
    ucsFpiOperatorPass
    ucsFpiCashDeskNo
    ucsFpiLastReceipt
End Enum

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

Private WithEvents m_oFP            As cDatecsPrinter
Attribute m_oFP.VB_VarHelpID = -1
Private m_sLog                      As String
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
    Group           As String
    Quantity        As String
    WayOfPayment    As String
    Amount          As String
End Type

Private Type UcsContext
    Total           As Double
    Amount          As Double
    PmtPrinted      As Boolean
    ChangePrinted   As Boolean
End Type

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunc As String)
    Debug.Print MODULE_NAME & "." & sFunc & ": " & Error
    DebugLog sFunc, "Run-time error: " & Error
End Sub

'=========================================================================
' Methods
'=========================================================================

Public Function EnumPorts() As Variant
    With New cDatecsPrinter
        EnumPorts = .EnumPorts
    End With
End Function

Public Function GetHeaderFooter( _
            Device As String, _
            Timeout As Long, _
            Optional OperatorData As Variant, _
            Optional RetVal As Variant, _
            Optional LastError As String, _
            Optional CommandLog As String) As Boolean
    Dim vInput      As Variant
    
    On Error GoTo EH
    '--- init & localize fiscal printer
    pvInit Device, Timeout
    '--- fetch header & footer
    vInput = RetVal
    ReDim RetVal(0 To 16)
    If LenB(pvAccess(vInput, 0, "1")) <> 0 Then
        RetVal(0) = m_oFP.SendCommand(ucsFpcInitHeaderFooter, "I0")
    End If
    If LenB(pvAccess(vInput, 1, "1")) <> 0 Then
        RetVal(1) = m_oFP.SendCommand(ucsFpcInitHeaderFooter, "I1")
    End If
    If LenB(pvAccess(vInput, 2, "1")) <> 0 Then
        RetVal(2) = m_oFP.SendCommand(ucsFpcInfoBulstat)
    End If
    If LenB(pvAccess(vInput, 3, "1")) <> 0 Then
        RetVal(3) = m_oFP.SendCommand(ucsFpcInitHeaderFooter, "I2")
    End If
    If LenB(pvAccess(vInput, 4, "1")) <> 0 Then
        RetVal(4) = m_oFP.SendCommand(ucsFpcInitHeaderFooter, "I3")
    End If
    If LenB(pvAccess(vInput, 5, "1")) <> 0 Then
        RetVal(5) = m_oFP.SendCommand(ucsFpcInitHeaderFooter, "I4")
    End If
    If LenB(pvAccess(vInput, 6, "1")) <> 0 Then
        RetVal(6) = m_oFP.SendCommand(ucsFpcInitHeaderFooter, "I5")
    End If
    If LenB(pvAccess(vInput, 7, "1")) <> 0 Then
        RetVal(7) = m_oFP.SendCommand(ucsFpcInitHeaderFooter, "I6")
    End If
    If LenB(pvAccess(vInput, 8, "1")) <> 0 Then
        RetVal(8) = m_oFP.SendCommand(ucsFpcInitHeaderFooter, "I7")
    End If
    If LenB(pvAccess(vInput, 9, "1")) <> 0 Then
        RetVal(9) = m_oFP.SendCommand(ucsFpcInfoDiagnostics, "0")
    End If
    If LenB(pvAccess(vInput, 10, "1")) <> 0 Then
        If LenB(pvAccess(OperatorData, 0)) <> 0 Then
            RetVal(10) = m_oFP.SendCommand(ucsFpcInfoOperator, pvAccess(OperatorData, 0))
        End If
    End If
    If LenB(pvAccess(vInput, 11, "1")) <> 0 Then
        RetVal(11) = m_oFP.SendCommand(ucsFpcInfoDateTime)
    End If
    If LenB(pvAccess(vInput, 12, "1")) <> 0 Then
        RetVal(12) = m_oFP.SendCommand(ucsFpcInfoLastReceipt)
    End If
    If LenB(pvAccess(vInput, 13, "1")) <> 0 Then
        RetVal(13) = m_oFP.SendCommand(ucsFpcInitPaymentType, "I")
    End If
    If LenB(pvAccess(vInput, 14, "1")) <> 0 Then
        RetVal(14) = m_oFP.SendCommand(ucsFpcInitPaymentType, "J")
    End If
    If LenB(pvAccess(vInput, 15, "1")) <> 0 Then
        RetVal(15) = m_oFP.SendCommand(ucsFpcInitPaymentType, "K")
    End If
    If LenB(pvAccess(vInput, 16, "1")) <> 0 Then
        RetVal(16) = m_oFP.SendCommand(ucsFpcInitPaymentType, "L")
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
            Device As String, _
            Timeout As Long, _
            Optional ByVal Count As Long, _
            Optional RetVal As Variant, _
            Optional LastError As String, _
            Optional CommandLog As String) As Boolean
    Dim lIdx            As Long
    
    On Error GoTo EH
    '--- init & localize fiscal printer
    pvInit Device, Timeout
    '--- fetch operators
    ReDim RetVal(0 To IIf(Count > 0, Count, 24) - 1)
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
            Device As String, _
            Timeout As Long, _
            ByVal ReceiptType As UcsFiscalPrinterReceiptTypeEnum, _
            Rows As Variant, _
            Optional OperatorData As Variant, _
            Optional InvoiceData As Variant, _
            Optional PrintDuplicate As Boolean, _
            Optional LastError As String, _
            Optional CommandLog As String) As Boolean
    Const FUNC_NAME     As String = "PrintReceipt"
    Const MAX_WIDTH     As Long = 30
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
    pvInit Device, Timeout
    '--- cancel outstanding transaction
    pvReset
    '--- open transaction
    sOperLogin = pvAccess(OperatorData, ucsFpiOperatorCode, "1") & "," & pvAccess(OperatorData, ucsFpiOperatorPass, "0000") & "," & pvAccess(OperatorData, ucsFpiCashDeskNo, "1")
    Select Case ReceiptType
    Case ucsFptNonFiscal
        m_oFP.SendCommand ucsFpcNonFiscalOpen
        If LenB(pvAccess(OperatorData, ucsFpiOperatorCode)) <> 0 Then
            vElem = Split(m_oFP.SendCommand(ucsFpcInfoOperator, pvAccess(OperatorData, ucsFpiOperatorCode)), ",")
            sText = Format(pvAccess(OperatorData, ucsFpiOperatorCode), "0000") & " " & pvAccess(vElem, 5)
            If LenB(pvAccess(OperatorData, ucsFpiCashDeskNo)) <> 0 Then
                sText = sText & Right(Space(MAX_WIDTH) & Format(pvAccess(OperatorData, ucsFpiCashDeskNo), "00"), MAX_WIDTH - Len(sText))
            End If
            m_oFP.SendCommand ucsFpcNonFiscalText, sText
        End If
    Case ucsFptFiscal
        bInLogin = True
        m_oFP.SendCommand ucsFpcFiscalOpen, sOperLogin
        bInLogin = False
    Case ucsFptInvoice
        m_oFP.Exceptions = False
        If LenB(pvAccess(InvoiceData, ucsFpiInvoiceNo)) <> 0 Then
            m_oFP.SendCommand ucsFpcInitInvoiceNo, pvAccess(InvoiceData, ucsFpiInvoiceNo) & "," & pvAccess(InvoiceData, ucsFpiInvoiceNo)
        End If
        m_oFP.Exceptions = True
        bInLogin = True
        m_oFP.SendCommand ucsFpcFiscalOpen, sOperLogin & ",I"
        bInLogin = False
    End Select
    bInTrans = True
    If IsArray(Rows) Then
        '--- print rows
        For Each vElem In Rows
            lCurrentRow = lCurrentRow + 1
            If ReceiptType = ucsFptNonFiscal Then
                For Each vCmd In pvGetNonFiscalRow(pvGetRowData(vElem), uCtx)
                    Rows(lCurrentRow - 1) = m_oFP.SendCommand(ucsFpcNonFiscalText, C_Str(vCmd))
                Next
            Else
                For Each vCmd In pvGetFiscalRow(pvGetRowData(vElem), uCtx)
                    Rows(lCurrentRow - 1) = m_oFP.SendCommand(vCmd(0), pvAccess(vCmd, 1))
                Next
            End If
        Next
        lCurrentRow = 0
    End If
    '--- commit transaction
    Select Case ReceiptType
    Case ucsFptNonFiscal
        '--- print payment
        If Not uCtx.PmtPrinted Then
            For Each vCmd In pvGetNonFiscalRow(pvGetRowData(Array("", "P", uCtx.Total)), uCtx)
                m_oFP.SendCommand ucsFpcNonFiscalText, C_Str(vCmd)
            Next
        End If
        '--- print change
        If Not uCtx.ChangePrinted Then
            For Each vCmd In pvGetNonFiscalRow(pvGetRowData(Array("")), uCtx)
                m_oFP.SendCommand ucsFpcNonFiscalText, C_Str(vCmd)
            Next
        End If
        '--- print custom footer (centered)
        sText = m_oFP.SendCommand(ucsFpcInitHeaderFooter, "I6")
        If LenB(sText) <> 0 Then
            m_oFP.SendCommand ucsFpcNonFiscalText, Space((MAX_WIDTH - Len(sText)) \ 2) & sText
        End If
        sText = m_oFP.SendCommand(ucsFpcInitHeaderFooter, "I7")
        If LenB(sText) <> 0 Then
            m_oFP.SendCommand ucsFpcNonFiscalText, Space((MAX_WIDTH - Len(sText)) \ 2) & sText
        End If
        m_oFP.SendCommand ucsFpcNonFiscalText, vbNullString
        '--- closing footer
        m_oFP.SendCommand ucsFpcNonFiscalClose
        '--- open door
        m_oFP.SendCommand ucsFpcInitOpenDoor
        '--- success
        PrintReceipt = True
    Case ucsFptFiscal
        '--- print outstanding payment
        For Each vCmd In pvGetFiscalRow(pvGetRowData(Array("")), uCtx)
            If vCmd(0) <> ucsFpcFiscalText Then
                m_oFP.SendCommand vCmd(0), pvAccess(vCmd, 1)
            End If
        Next
        m_oFP.SendCommand ucsFpcFiscalClose
        '--- success
        PrintReceipt = True
    Case ucsFptInvoice
        '--- print outstanding payment
        For Each vCmd In pvGetFiscalRow(pvGetRowData(Array("")), uCtx)
            If vCmd(0) <> ucsFpcFiscalText Then
                m_oFP.SendCommand vCmd(0), pvAccess(vCmd, 1)
            End If
        Next
        If LenB(pvAccess(InvoiceData, ucsFpiCgBulstat)) <> 0 Or LenB(pvAccess(InvoiceData, ucsFpiCgName)) <> 0 Then
            m_oFP.SendCommand ucsFpcFiscalCgInfo, _
                Left(pvAccess(InvoiceData, ucsFpiCgBulstat), 14) & vbTab & _
                Left(pvAccess(InvoiceData, ucsFpiOwnPersonSeller), 30) & vbTab & _
                Left(pvAccess(InvoiceData, ucsFpiCgPersonReceiver), 25) & vbTab & _
                Left(pvAccess(InvoiceData, ucsFpiCgName), 27) & vbTab & _
                Left(pvAccess(InvoiceData, ucsFpiCgVatNo), 16) & vbTab & _
                pvSplitCgAddress(pvAccess(InvoiceData, ucsFpiCgAddress))
        End If
        m_oFP.SendCommand ucsFpcFiscalClose
        '--- success
        PrintReceipt = True
    Case Else
        '--- moje da se pechata dublikat
    End Select
    bInTrans = False
    If LenB(pvAccess(OperatorData, ucsFpiLastReceipt, vbNullString)) <> 0 Then
        OperatorData(ucsFpiLastReceipt) = m_oFP.SendCommand(ucsFpcInfoLastReceipt)
    End If
    If PrintDuplicate Then
        m_oFP.SendCommand ucsFpcPrintDuplicate, "1"
        '--- success
        PrintReceipt = True
    End If
    '--- disconnect
    Set m_oFP = Nothing
    CommandLog = m_sLog
    Exit Function
EH:
    LastError = Err.Description
    Resume NextLine
NextLine:
    On Error Resume Next
    If Not m_oFP Is Nothing Then
        If bInLogin Then
            LastError = LastError & STR_OR_INCORRECT_LOGIN
            OperatorData(ucsFpiOperatorPass) = Empty
        End If
        If lCurrentRow > 0 Then ' m_oFP.Status(ucsStbPrintingError)
            LastError = STR_ERROR_LINE & lCurrentRow & ": " & LastError
        End If
        If bInTrans And m_oFP.IsConnected Then
            If ReceiptType = ucsFptNonFiscal Then
                m_oFP.SendCommand ucsFpcNonFiscalClose
            Else
                DebugLog FUNC_NAME, "EH bInTrans And m_oFP.IsConnected, ReceiptType = ucsFptFiscal"
                '--- note: when printing invoice, if no contragent info set then cancel fails!
                m_oFP.Exceptions = False
                m_oFP.SendCommand ucsFpcFiscalPayment, vbTab
                m_oFP.SendCommand ucsFpcFiscalCgInfo, "0000000000"
                m_oFP.SendCommand ucsFpcFiscalCancel
                m_oFP.SendCommand ucsFpcFiscalClose
                m_oFP.Exceptions = True
            End If
            bInTrans = False
        End If
        Set m_oFP = Nothing
    End If
    DebugLog FUNC_NAME, "EH LastError=" & LastError & ", m_sLog=" & m_sLog
    CommandLog = m_sLog
End Function

Public Function PrintReport( _
            Device As String, _
            Timeout As Long, _
            ByVal ReportType As UcsFiscalPrinterReportTypeEnum, _
            Optional ReportData As Variant, _
            Optional LastError As String, _
            Optional CommandLog As String) As Boolean
    Const FUNC_NAME     As String = "PrintReport"
    Dim eCmd            As UcsFiscalPrinterCommandsEnum
    Dim vResult         As Variant
    
    On Error GoTo EH
    '--- init & localize fiscal printer
    pvInit Device, Timeout
    pvReset
    Select Case ReportType
    Case ucsFprDaily
        '-- ReportData = { IsClear, IsItems, IsDepartments }
        If C_Bool(pvAccess(ReportData, 1)) And C_Bool(pvAccess(ReportData, 2)) Then
            eCmd = ucsFpcPrintDailyReportItemsDepartments
        ElseIf C_Bool(pvAccess(ReportData, 1)) Then
            eCmd = ucsFpcPrintDailyReportItems
        ElseIf C_Bool(pvAccess(ReportData, 2)) Then
            eCmd = ucsFpcPrintDailyReportDepartments
        Else
            eCmd = ucsFpcPrintDailyReport
        End If
        vResult = m_oFP.SendCommand(eCmd, IIf(C_Bool(pvAccess(ReportData, 0)), "0", "2") & "N")
    Case ucsFprNumber
        '-- ReportData = { FromNum, ToNum, IsDetailed }
        If C_Bool(pvAccess(ReportData, 2)) Then
            eCmd = ucsFpcPrintReportByNumberDetailed
        Else
            eCmd = ucsFpcPrintReportByNumberShort
        End If
        vResult = m_oFP.SendCommand(eCmd, pvAccess(ReportData, 0) & "," & pvAccess(ReportData, 1))
    Case ucsFprDate
        '-- ReportData = { FromDate, ToDate, IsDetailed }
        If C_Bool(pvAccess(ReportData, 0)) Then
            eCmd = ucsFpcPrintReportByDateDetailed
        Else
            eCmd = ucsFpcPrintReportByDateShort
        End If
        vResult = m_oFP.SendCommand(eCmd, pvAccess(ReportData, 0) & "," & pvAccess(ReportData, 1))
    Case ucsFprOperator
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

Public Function OpenDoor( _
            Device As String, _
            Timeout As Long, _
            Optional LastError As String, _
            Optional CommandLog As String) As Boolean
    Const FUNC_NAME     As String = "OpenDoor"
    
    On Error GoTo EH
    '--- init & localize fiscal printer
    pvInit Device, Timeout
    pvReset
    m_oFP.SendCommand ucsFpcInitOpenDoor
    '--- disconnect
    Set m_oFP = Nothing
    CommandLog = m_sLog
    '--- success
    OpenDoor = True
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
            Device As String, _
            Timeout As Long, _
            sDateTime As String, _
            Optional LastError As String, _
            Optional CommandLog As String) As Boolean
    Const FUNC_NAME     As String = "InitDateTime"
    
    On Error GoTo EH
    '--- init & localize fiscal printer
    pvInit Device, Timeout
    pvReset
    If LenB(sDateTime) = 0 Then
        sDateTime = Format(Now, "dd-MM-yy") & " " & Format(Now, "hh:mm:ss")
    End If
    m_oFP.Exceptions = False
    m_oFP.SendCommand ucsFpcInitDateTime, sDateTime
    m_oFP.Exceptions = True
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

Private Sub pvInit(Device As String, ByVal Timeout As Long)
    m_sLog = vbNullString
    Set m_oFP = New cDatecsPrinter
    m_oFP.LocalizedText(ucsLciInternalErrors) = STR_INTERNAL
    m_oFP.LocalizedText(ucsLciPrinterStatuses) = STR_STATUSES
    m_oFP.LocalizedText(ucsLciPrinterErrors) = STR_ERRORS_ONLY
    m_oFP.LocalizedText(ucsLciDipSwitchesTexts) = STR_DIP_SWITCHES
    m_oFP.LocalizedText(ucsLciMemoryTexts) = STR_MEMORY
    m_oFP.Init Device, Timeout
End Sub

Private Sub pvReset()
    If Left(m_oFP.SendCommand(ucsFpcInfoTransaction), 1) = "1" Then
        If m_oFP.Status(ucsStbFiscalPrinting) Then
            m_oFP.Exceptions = False
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
            m_oFP.Exceptions = True
        Else
            m_oFP.SendCommand ucsFpcNonFiscalClose
        End If
    End If
End Sub

Private Function pvGetRowData(vRow As Variant) As UcsRow
    On Error Resume Next
    If UBound(vRow) = ucsFpiItem Then
        pvGetRowData.Type = ucsRwtText
        pvGetRowData.Item = C_Str(vRow(ucsFpiItem))
    ElseIf UBound(vRow) = ucsFpiAmount Then
        pvGetRowData.Type = ucsRwtPayment
        pvGetRowData.Item = C_Str(vRow(ucsFpiItem))
        pvGetRowData.WayOfPayment = C_Str(vRow(ucsFpiWayOfPmt))
        pvGetRowData.Amount = C_Str(vRow(ucsFpiAmount))
    Else
        pvGetRowData.Type = ucsRwtSell
        pvGetRowData.Item = C_Str(vRow(ucsFpiItem))
        pvGetRowData.Price = C_Str(vRow(ucsFpiPrice))
        pvGetRowData.Group = C_Str(vRow(ucsFpiGroup))
        pvGetRowData.Quantity = C_Str(vRow(ucsFpiQuantity))
    End If
    On Error GoTo 0
End Function

Private Function pvGetNonFiscalRow(uRow As UcsRow, uCtx As UcsContext) As Variant
    Const FUNC_NAME     As String = "pvGetNonFiscalRow"
    Const MAX_WIDTH     As Long = 30
    Dim lIdx            As Long
    Dim sItem           As String
    Dim sQuantity       As String
    Dim sPrice          As String
    Dim vSplit          As Variant
    Dim vRet            As Variant
    Dim sWayOfPmt       As String
    
    On Error GoTo EH
    With uRow
        If .Type = ucsRwtSell Then
            If uCtx.ChangePrinted Or uCtx.PmtPrinted Then
                On Error GoTo 0
                Err.Raise vbObjectError, , m_oFP.frGetFlagsText(ucsStbInvalidCommand, m_oFP.LocalizedText(ucsLciPrinterStatuses))
            End If
            sPrice = pvSafeFormat(C_Dbl(.Price), "0.00")
            If LenB(.Quantity) <> 0 And C_Dbl(.Quantity) <> 1 Then
                If Len(sPrice) < 7 Then
                    sQuantity = Space(7 - Len(sPrice))
                End If
                sQuantity = Right(" " & pvSafeFormat(C_Dbl(.Quantity), "0.0##") & " x " & sQuantity & sPrice & "  ", MAX_WIDTH)
                sPrice = pvSafeFormat(Round(C_Dbl(.Price), 2) * Round(C_Dbl(.Quantity), 3), "0.00") & " " & .Group
                uCtx.Total = uCtx.Total + Round(Round(C_Dbl(.Price), 2) * Round(C_Dbl(.Quantity), 3), 2)
            Else
                sPrice = sPrice & " " & .Group
                uCtx.Total = uCtx.Total + Round(C_Dbl(.Price), 2)
            End If
            '--- split long item name
            sItem = RTrim(pvSafeText(.Item))
            ReDim vSplit(-1 To -1)
            If Len(sItem) > MAX_WIDTH - Len(sPrice) - 1 Then
                ReDim Preserve vSplit(0 To UBound(vSplit) + 1)
                vSplit(UBound(vSplit)) = pvChopOff(sItem, MAX_WIDTH - Len(sPrice) - 1)
            End If
            If Len(sItem) > MAX_WIDTH - Len(sQuantity) - 1 Then
                ReDim Preserve vSplit(0 To UBound(vSplit) + 1)
                vSplit(UBound(vSplit)) = pvChopOff(sItem, MAX_WIDTH - Len(sQuantity) - 1)
            End If
            Do While Len(sItem) > MAX_WIDTH
                ReDim Preserve vSplit(0 To UBound(vSplit) + 1)
                vSplit(UBound(vSplit)) = pvChopOff(sItem, MAX_WIDTH)
            Loop
            If LenB(sItem) <> 0 Then
                ReDim Preserve vSplit(0 To UBound(vSplit) + 1)
                vSplit(UBound(vSplit)) = sItem
            End If
            '--- append price & quantity
            If LenB(sPrice) <> 0 Then
                If UBound(vSplit) < 0 Then
                    ReDim Preserve vSplit(0 To 0)
                End If
                vSplit(0) = vSplit(0) & Space(MAX_WIDTH - Len(sPrice) - Len(vSplit(0))) & sPrice
            End If
            If LenB(sQuantity) <> 0 Then
                If UBound(vSplit) < 1 Then
                    ReDim Preserve vSplit(0 To 1)
                End If
                vSplit(1) = vSplit(1) & Space(MAX_WIDTH - Len(sQuantity) - Len(vSplit(1))) & sQuantity
            End If
            '--- retval
            ReDim vRet(0 To UBound(vSplit))
            For lIdx = 0 To UBound(vRet)
                vRet(lIdx) = vSplit(UBound(vSplit) - lIdx)
            Next
        ElseIf .Type = ucsRwtPayment Then
            If uCtx.ChangePrinted Then
                On Error GoTo 0
                Err.Raise vbObjectError, , m_oFP.frGetFlagsText(ucsStbInvalidCommand, m_oFP.LocalizedText(ucsLciPrinterStatuses))
            End If
            If Not uCtx.PmtPrinted Then
                uCtx.PmtPrinted = True
                ReDim vRet(0 To 2)
                vRet(0) = String(MAX_WIDTH, "-")
                vRet(1) = STR_TOTAL & Right(Space(MAX_WIDTH) & pvSafeFormat(uCtx.Total, "0.00"), MAX_WIDTH - Len(STR_TOTAL) - 2)
            Else
                ReDim vRet(0 To 0)
            End If
            sWayOfPmt = IIf(.WayOfPayment = "D", STR_CARD, STR_CASH)
            vRet(UBound(vRet)) = sWayOfPmt & Right(Space(MAX_WIDTH) & pvSafeFormat(C_Dbl(.Amount), "0.00"), MAX_WIDTH - Len(sWayOfPmt) - 2)
            uCtx.Amount = uCtx.Amount + Round(C_Dbl(.Amount), 2)
        ElseIf .Type = ucsRwtText Then
            vRet = pvWrapText(RTrim(pvSafeText(.Item)), MAX_WIDTH)
            If Not uCtx.ChangePrinted And uCtx.PmtPrinted Then
                uCtx.ChangePrinted = True
                If Abs(uCtx.Amount - uCtx.Total) >= 0.01 Then
                    ReDim Preserve vRet(0 To UBound(vRet) + 1)
                    For lIdx = UBound(vRet) To 1 Step -1
                        vRet(lIdx) = vRet(lIdx - 1)
                    Next
                    vRet(0) = STR_CHANGE & Right(Space(MAX_WIDTH) & pvSafeFormat(uCtx.Amount - uCtx.Total, "0.00"), MAX_WIDTH - Len(STR_CHANGE) - 2)
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
    Const MAX_WIDTH     As Long = 25
    Const MAX_TEXT_WIDTH As Long = 30
    Dim lIdx            As Long
    Dim sSign           As String
    Dim sText           As String
    Dim vSplit          As Variant
    Dim vRet            As Variant
    
    On Error GoTo EH
    With uRow
        If .Type = ucsRwtSell Then
            If uCtx.ChangePrinted Or uCtx.PmtPrinted Then
                On Error GoTo 0
                Err.Raise vbObjectError, , m_oFP.frGetFlagsText(ucsStbInvalidCommand, m_oFP.LocalizedText(ucsLciPrinterStatuses))
            End If
            ReDim vRet(0 To 0)
            '--- split item
            vSplit = pvWrapText(RTrim(pvSafeText(.Item)), MAX_WIDTH)
            If UBound(vSplit) = 0 Then
                sText = vSplit(0)
            Else
                sText = vSplit(UBound(vSplit) - 1) & vbLf & vSplit(UBound(vSplit))
            End If
            '--- long item name -> prepend ucsFpcFiscalText
            If UBound(vSplit) > 1 Then
                ReDim vRet(0 To UBound(vSplit) - 1)
                For lIdx = 0 To UBound(vRet) - 1
                    vRet(lIdx) = Array(ucsFpcFiscalText, vSplit(lIdx))
                Next
            End If
            '--- figure out sign
            If LenB(.Quantity) <> 0 Then
                sSign = IIf(C_Dbl(.Price) * C_Dbl(.Quantity) < 0, "-", vbNullString)
            Else
                sSign = IIf(C_Dbl(.Price) < 0, "-", vbNullString)
            End If
            '--- append rest
            sText = sText & vbTab & .Group & sSign & pvSafeFormat(Abs(C_Dbl(.Price)), "0.00")
            If LenB(.Quantity) <> 0 Then
                sText = sText & "*" & pvSafeFormat(Abs(C_Dbl(.Quantity)), "0.000")
                uCtx.Total = uCtx.Total + Round(Round(C_Dbl(.Price), 2) * Round(C_Dbl(.Quantity), 3), 2)
            Else
                uCtx.Total = uCtx.Total + Round(C_Dbl(.Price), 2)
            End If
            vRet(UBound(vRet)) = Array(ucsFpcFiscalSell, sText)
            pvGetFiscalRow = vRet
        ElseIf .Type = ucsRwtPayment Then
            If uCtx.ChangePrinted Then
                On Error GoTo 0
                Err.Raise vbObjectError, , m_oFP.frGetFlagsText(ucsStbInvalidCommand, m_oFP.LocalizedText(ucsLciPrinterStatuses))
            End If
            uCtx.PmtPrinted = True
            If LenB(.Amount) = 0 Then
                uCtx.Amount = uCtx.Total
            Else
                uCtx.Amount = uCtx.Amount + Round(Abs(C_Dbl(.Amount)), 2)
            End If
            ReDim vRet(0 To 0)
            '--- split item
            vSplit = pvWrapText(RTrim(pvSafeText(.Item)), MAX_WIDTH)
            If UBound(vSplit) = 0 Then
                sText = vSplit(0)
            Else
                sText = vSplit(UBound(vSplit) - 1) & vbLf & vSplit(UBound(vSplit))
            End If
            '--- long item name -> prepend ucsFpcFiscalText
            If UBound(vSplit) > 1 Then
                ReDim vRet(0 To UBound(vSplit) - 1)
                For lIdx = 0 To UBound(vRet) - 1
                    vRet(lIdx) = Array(ucsFpcFiscalText, vSplit(lIdx))
                Next
            End If
            '--- append rest
            sText = sText & vbTab
            If LenB(.Amount) <> 0 Then
                sText = sText & .WayOfPayment & pvSafeFormat(Abs(C_Dbl(.Amount)), "0.00")
            End If
            vRet(UBound(vRet)) = Array(ucsFpcFiscalPayment, sText)
        ElseIf .Type = ucsRwtText Then
            '--- split item
            vSplit = pvWrapText(RTrim(pvSafeText(.Item)), MAX_TEXT_WIDTH)
            ReDim vRet(0 To UBound(vSplit))
            For lIdx = 0 To UBound(vRet)
                vRet(lIdx) = Array(ucsFpcFiscalText, vSplit(lIdx))
            Next
            If Not uCtx.PmtPrinted Or uCtx.Amount < uCtx.Total Then
                uCtx.PmtPrinted = True
                uCtx.ChangePrinted = True
                uCtx.Amount = uCtx.Total
                ReDim Preserve vRet(0 To UBound(vRet) + 1)
                For lIdx = UBound(vRet) To 1 Step -1
                    vRet(lIdx) = vRet(lIdx - 1)
                Next
                vRet(0) = Array(ucsFpcFiscalPayment, vbTab)
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
    pvChopOff = Right(sItem, lIdx)
    For lIdx = Len(sItem) - lIdx - 1 To 0 Step -1
        If Not pvIsWhitespace(Mid(sItem, lIdx, 1)) Then
            Exit For
        End If
    Next
    sItem = Left(sItem, lIdx)
End Function

Private Function pvSplitCgAddress(ByVal sAddress As String) As String
    Const FUNC_NAME     As String = "pvSplitCgAddress"
    Const MAX_WIDTH     As Long = 30
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
                    If Len(sTemp) > MAX_WIDTH Then
                        sRow2 = vElem
                    Else
                        sRow1 = sTemp
                    End If
                End If
            End If
        Next
    End If
    pvSplitCgAddress = Left(sRow1, MAX_WIDTH) & IIf(LenB(sRow2) <> 0, vbLf, vbNullString) & Left(sRow2, MAX_WIDTH)
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
    
    ReDim vRet(0 To Len(sText))
    Do While LenB(sText) <> 0
        lRight = lWidth + 1
        If lRight > Len(sText) Then
            lRight = Len(sText) + 1
        Else
            If pvIsDelimiter(Mid(sText, lRight, 1)) Then
                Do While pvIsWhitespace(Mid(sText, lRight, 1)) And lRight <= Len(sText)
                    lRight = lRight + 1
                Loop
            Else
                Do While lRight > 1
                    If pvIsDelimiter(Mid(sText, lRight - 1, 1)) Then
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
        Do While pvIsWhitespace(Mid(sText, lLeft, 1)) And lLeft > 0
            lLeft = lLeft - 1
        Loop
        vRet(lCount) = Left(sText, lLeft)
        lCount = lCount + 1
        sText = Mid(sText, lRight)
    Loop
    If lCount = 0 Then
        pvWrapText = Array("")
    Else
        ReDim Preserve vRet(0 To lCount - 1)
        pvWrapText = vRet
    End If
End Function

Private Function pvIsDelimiter(sText As String)
    Const STR_DELIMS As String = "~#$^&*_+-=\|/ " & vbTab & vbCrLf
    If InStr(1, STR_DELIMS, Left(sText, 1)) > 0 Then
        pvIsDelimiter = True
    End If
End Function

Private Function pvIsWhitespace(sText As String)
    Const STR_WHITESPACE As String = " " & vbTab & vbCrLf
    If InStr(1, STR_WHITESPACE, Left(sText, 1)) > 0 Then
        pvIsWhitespace = True
    End If
End Function

Private Function pvSafeText(sText As String) As String
    Dim lIdx            As Long
    
    pvSafeText = sText
    For lIdx = 0 To 31
        pvSafeText = Replace(pvSafeText, Chr$(lIdx), vbNullString)
    Next
End Function

Private Function pvSafeFormat(Expression As Variant, Optional Fmt As Variant) As String
    pvSafeFormat = Replace(Format(Expression, Fmt), m_sDecimalSeparator, ".")
End Function

Private Function pvAccess(vData As Variant, ByVal lIdx As Long, Optional sDefault As String) As String
    On Error Resume Next
    pvAccess = sDefault
    pvAccess = C_Str(vData(lIdx))
    On Error GoTo 0
End Function

Private Function C_Str(v As Variant) As String
    On Error Resume Next
    C_Str = CStr(v)
    On Error GoTo 0
End Function

Private Function C_Dbl(v As Variant) As Double
    On Error Resume Next
    C_Dbl = CDbl(Replace(C_Str(v), ".", m_sDecimalSeparator))
    On Error GoTo 0
End Function

Private Function C_Bool(v As Variant) As Boolean
    On Error Resume Next
    C_Bool = CBool(v)
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

Private Sub DebugLog(sFunc As String, sText As String)
    Dim nFile           As Integer
    
    On Error Resume Next
    nFile = FreeFile
    Open Environ$("TEMP") & "\UcsFP.log" For Append Access Write As #nFile
    Print #nFile, MODULE_NAME & "." & sFunc & "(" & Now & "): " & sText
    Close #nFile
    On Error GoTo 0
End Sub

Private Sub Class_Initialize()
    m_sDecimalSeparator = pvGetDecimalSeparator()
End Sub

'=========================================================================
' Event handlers
'=========================================================================

Private Sub m_oFP_CommandComplete(ByVal lCmd As Long, sData As String, sResult As String)
    Const FUNC_NAME     As String = "m_oFP_CommandComplete"
    
    On Error GoTo EH
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
