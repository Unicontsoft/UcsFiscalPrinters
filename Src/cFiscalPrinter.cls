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
' $Header: /UcsFiscalPrinter/Src/cFiscalPrinter.cls 2     19.06.08 18:26 Wqw $
'
'   Unicontsoft Fiscal Printers Project
'   Copyright (c) 2008 Unicontsoft
'
'   Interfacen class za uprawlenie na fiscalni printer
'
' $Log: /UcsFiscalPrinter/Src/cFiscalPrinter.cls $
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

Public Enum UcsFiscalPrinterIndexesEnum
    ucsFpiItem = 0
    ucsFpiPrice
    ucsFpiGroup
    ucsFpiQuantity
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
End Enum

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_TOTAL                 As String = "ОБЩО: "
Private Const STR_INTERNAL              As String = "Не е указано устройство|Грешка при отваряне: |Грешка при SetCommTimeouts: |Грешка при BuildCommDCB: |Грешка при SetCommState: |Грешка при WriteFile: |Време за достъп изтече в очакване на отговор|Грешка при ReadFile: |Невалиден формат на съобщение или сума за проверка (NAK)|Вече има отворена връзка|Няма отворена връзка|Невалидна дължина на отговора|Липсва символ за край на отговора|Липсва символ за начало на отговора"
Private Const STR_STATUSES              As String = "Синтактична грешка|Невалидна команда|Неустановени дата и час|3|Неизправност в механизма на печатащото устройство|5|6|7|Аритметично препълване|Непозволена команда|Зануляване на опертивна памет|3|Разрушено съдържание на опертивна памет|5|6|7|Няма хартия|1|2|Отворен фискален бон|4|Отворен служебен (нефискален) бон|6|7"
Private Const STR_ERRORS_ONLY           As String = "Синтактична грешка|Невалидна команда|2|3|Неизправност в механицва на печатащото устройство|5|6|7|1|Непозволена команда|Зануляване на опертивна памет|3|Разрушено съдържание на опертивна памет|5|6|7|Няма хартия|1|2|3|4|5|6|7"
Private Const STR_DIP_SWITCHES          As String = "Автоматично центриране на header и footer|Предварителен header|Sw1.3|Sw1.4|Нулиране на паметта|Прозрачен дисплей|Без данни на дисплея|7"
Private Const STR_MEMORY                As String = "Грешка при запис|1|Няма модул фискална памет|Малко свободно място във фискалната памет|Пълна фискална памет|5|6|7|Фискалната памет забранена за запис|Форматирана фискална памет|2|Принтерът е във фискален режим|Зададени данъчни ставки|Програмирани индивидуален номер и номер на фискалната памет|6|7"
Private Const STR_ERROR_LINE            As String = "Грешка при печат на ред "

Private WithEvents m_oFP            As cDatecsPrinter
Attribute m_oFP.VB_VarHelpID = -1
Private m_sLog                      As String

Private Type UcsRow
    Item            As String
    Price           As String
    Group           As String
    Quantity        As String
End Type

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunc As String)
    Debug.Print MODULE_NAME & "." & sFunc & ": " & Error
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
    On Error GoTo EH
    '--- init & localize fiscal printer
    pvInit Device, Timeout
    '--- fetch header & footer
    ReDim RetVal(0 To 12)
    RetVal(0) = m_oFP.SendCommand(ucsFpcInitHeaderFooter, "I0")
    RetVal(1) = m_oFP.SendCommand(ucsFpcInitHeaderFooter, "I1")
    RetVal(2) = m_oFP.SendCommand(ucsFpcInfoBulstat)
    RetVal(3) = m_oFP.SendCommand(ucsFpcInitHeaderFooter, "I2")
    RetVal(4) = m_oFP.SendCommand(ucsFpcInitHeaderFooter, "I3")
    RetVal(5) = m_oFP.SendCommand(ucsFpcInitHeaderFooter, "I4")
    RetVal(6) = m_oFP.SendCommand(ucsFpcInitHeaderFooter, "I5")
    RetVal(7) = m_oFP.SendCommand(ucsFpcInitHeaderFooter, "I6")
    RetVal(8) = m_oFP.SendCommand(ucsFpcInitHeaderFooter, "I7")
    RetVal(9) = m_oFP.SendCommand(ucsFpcInfoDiagnostics, "0")
    If LenB(pvAccess(OperatorData, 0)) <> 0 Then
        RetVal(10) = m_oFP.SendCommand(ucsFpcInfoOperator, pvAccess(OperatorData, 0))
    End If
    RetVal(11) = m_oFP.SendCommand(ucsFpcInfoDateTime)
    RetVal(12) = m_oFP.SendCommand(ucsFpcInfoLastReceipt)
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
    Const MAX_WIDTH     As Long = 30
    Dim vElem           As Variant
    Dim sSecondRow      As String
    Dim bInTrans        As Boolean
    Dim lCurrentRow     As Long
    Dim sOperLogin      As String
    Dim dblTotal        As Double
    Dim sText           As String
    
    On Error GoTo EH
    '--- init & localize fiscal printer
    pvInit Device, Timeout
    '--- cancel outstanding transaction
    If Left(m_oFP.SendCommand(ucsFpcInfoTransaction), 1) = "1" Then
        If m_oFP.Status(ucsStbFiscalPrinting) Then
            m_oFP.Exceptions = False
            '--- note: when printing invoice, if no contragent info set then cancel fails!
            m_oFP.SendCommand ucsFpcFiscalCgInfo, "0000000000"
            '--- note: FP3530 moje da anulira winagi, FP550F ne moje
            m_oFP.SendCommand ucsFpcFiscalCancel
            '--- zaradi FP550F
            m_oFP.SendCommand ucsFpcFiscalClose
            m_oFP.Exceptions = True
        Else
            m_oFP.SendCommand ucsFpcNonFiscalClose
        End If
    End If
    '--- open transaction
    sOperLogin = pvAccess(OperatorData, ucsFpiOperatorCode, "1") & "," & pvAccess(OperatorData, ucsFpiOperatorPass, "0000") & "," & pvAccess(OperatorData, ucsFpiCashDeskNo, "1")
    Select Case ReceiptType
    Case ucsFptNonFiscal
        m_oFP.SendCommand ucsFpcNonFiscalOpen
        If Len(pvAccess(OperatorData, 0)) <> 0 Then
            vElem = Split(m_oFP.SendCommand(ucsFpcInfoOperator, pvAccess(OperatorData, 0)), ",")
            sText = Format(pvAccess(OperatorData, 0), "0000") & " " & pvAccess(vElem, 5)
            If LenB(pvAccess(OperatorData, 2)) <> 0 Then
                sText = sText & Right(Space(MAX_WIDTH) & pvAccess(OperatorData, 2), MAX_WIDTH - Len(sText))
            End If
            m_oFP.SendCommand ucsFpcNonFiscalText, sText
        End If
        m_oFP.SendCommand ucsFpcNonFiscalText, vbNullString
    Case ucsFptFiscal
        m_oFP.SendCommand ucsFpcFiscalOpen, sOperLogin
    Case ucsFptInvoice
        m_oFP.Exceptions = False
        If LenB(pvAccess(InvoiceData, ucsFpiInvoiceNo)) <> 0 Then
            m_oFP.SendCommand ucsFpcInitInvoiceNo, pvAccess(InvoiceData, ucsFpiInvoiceNo) & "," & pvAccess(InvoiceData, ucsFpiInvoiceNo)
        End If
        m_oFP.Exceptions = True
        m_oFP.SendCommand ucsFpcFiscalOpen, sOperLogin & ",I"
    End Select
    bInTrans = True
    '--- print rows
    For Each vElem In Rows
        lCurrentRow = lCurrentRow + 1
        If ReceiptType = ucsFptNonFiscal Then
            m_oFP.SendCommand ucsFpcNonFiscalText, pvGetNonFiscalRow(pvGetRowData(vElem), sSecondRow, dblTotal)
            If LenB(sSecondRow) <> 0 Then
                m_oFP.SendCommand ucsFpcNonFiscalText, sSecondRow
            End If
        Else
            m_oFP.SendCommand ucsFpcFiscalSell, pvGetFiscalRow(pvGetRowData(vElem))
        End If
    Next
    lCurrentRow = 0
    '--- commit transaction
    Select Case ReceiptType
    Case ucsFptNonFiscal
        '--- print total
        m_oFP.SendCommand ucsFpcNonFiscalText, String(MAX_WIDTH, "-")
        m_oFP.SendCommand ucsFpcNonFiscalText, STR_TOTAL & Right(Space(MAX_WIDTH) & Format(dblTotal, "0.00"), MAX_WIDTH - Len(STR_TOTAL) - 2)
        m_oFP.SendCommand ucsFpcNonFiscalText, vbNullString
        '--- print custom footer (centered)
        sText = m_oFP.SendCommand(ucsFpcInitHeaderFooter, "I6")
        If LenB(sText) <> 0 Then
            m_oFP.SendCommand ucsFpcNonFiscalText, Space((MAX_WIDTH - Len(sText)) \ 2) & sText
            sText = m_oFP.SendCommand(ucsFpcInitHeaderFooter, "I7")
            m_oFP.SendCommand ucsFpcNonFiscalText, Space((MAX_WIDTH - Len(sText)) \ 2) & sText
        End If
        '--- closing footer
        m_oFP.SendCommand ucsFpcNonFiscalClose
    Case ucsFptFiscal
        m_oFP.SendCommand ucsFpcFiscalPayment, vbTab
        m_oFP.SendCommand ucsFpcFiscalClose
    Case ucsFptInvoice
        m_oFP.SendCommand ucsFpcFiscalPayment, vbTab
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
    End Select
    bInTrans = False
    If PrintDuplicate Then
        m_oFP.SendCommand ucsFpcPrintDuplicate, "1"
    End If
    '--- disconnect
    Set m_oFP = Nothing
    CommandLog = m_sLog
    '--- success
    PrintReceipt = True
    Exit Function
EH:
    LastError = Err.Description
    Resume NextLine
NextLine:
    On Error Resume Next
    If Not m_oFP Is Nothing Then
        If m_oFP.Status(ucsStbPrintingError) And lCurrentRow > 0 Then
            LastError = STR_ERROR_LINE & lCurrentRow & ": " & LastError
        End If
        If bInTrans And m_oFP.IsConnected Then
            If ReceiptType = ucsFptNonFiscal Then
                m_oFP.SendCommand ucsFpcNonFiscalClose
            Else
                '--- note: when printing invoice, if no contragent info set then cancel fails!
                m_oFP.Exceptions = False
                m_oFP.SendCommand ucsFpcFiscalCgInfo, "0000000000"
                m_oFP.Exceptions = True
                m_oFP.SendCommand ucsFpcFiscalCancel
            End If
            bInTrans = False
        End If
        Set m_oFP = Nothing
    End If
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

Private Function pvGetRowData(vRow As Variant) As UcsRow
    On Error Resume Next
    pvGetRowData.Item = C_Str(vRow(ucsFpiItem))
    pvGetRowData.Price = C_Str(vRow(ucsFpiPrice))
    pvGetRowData.Group = C_Str(vRow(ucsFpiGroup))
    pvGetRowData.Quantity = C_Str(vRow(ucsFpiQuantity))
    On Error GoTo 0
End Function

Private Function pvGetNonFiscalRow(uRow As UcsRow, sSecondRow As String, dblTotal As Double) As String
    Const FUNC_NAME     As String = "pvGetNonFiscalRow"
    Const MAX_WIDTH     As Long = 30
    Dim lIdx            As Long
    Dim sItem           As String
    Dim sQuantity       As String
    Dim sFirstRow       As String
    Dim sRow            As String
    
    On Error GoTo EH
    uRow.Price = Format(C_Dbl(uRow.Price), "0.00")
    If LenB(uRow.Quantity) <> 0 And C_Dbl(uRow.Quantity) <> 1 Then
        If Len(uRow.Price) < 7 Then
            sQuantity = Space(7 - Len(uRow.Price))
        End If
        sQuantity = Right(" " & C_Dbl(uRow.Quantity) & " x " & sQuantity & uRow.Price & "  ", MAX_WIDTH)
        uRow.Price = Format(C_Dbl(uRow.Price) * C_Dbl(uRow.Quantity), "0.00")
    End If
    '--- split long items
    sItem = RTrim(Replace(uRow.Item, vbTab, vbNullString))
    If Len(sItem) > MAX_WIDTH - 2 - Len(uRow.Price) Then
        For lIdx = MAX_WIDTH - Len(sQuantity) To 0 Step -1
            If pvIsWhitespace(Mid(sItem, lIdx, 1)) Then
                Exit For
            End If
        Next
        If lIdx <= 0 Then
            lIdx = MAX_WIDTH - Len(sQuantity)
        End If
        sFirstRow = Left(sItem, lIdx) & Right(Space(MAX_WIDTH) & sQuantity, MAX_WIDTH - lIdx)
        sItem = Mid(sItem, lIdx + 1)
    ElseIf LenB(sQuantity) <> 0 Then
        sFirstRow = Right(Space(MAX_WIDTH) & sQuantity, MAX_WIDTH)
    End If
    '--- format
    sRow = Space(MAX_WIDTH)
    Mid(sRow, 1, Len(sItem)) = Left(sItem, MAX_WIDTH - 3 - Len(uRow.Price))
    Mid(sRow, MAX_WIDTH - 1 - Len(uRow.Price), Len(uRow.Price)) = uRow.Price
    Mid(sRow, MAX_WIDTH, 1) = Left(uRow.Group, 1)
    dblTotal = dblTotal + C_Dbl(uRow.Price)
    '--- figure out result
    If LenB(sFirstRow) <> 0 Then
        pvGetNonFiscalRow = RTrim(sFirstRow)
        sSecondRow = RTrim(sRow)
    Else
        pvGetNonFiscalRow = RTrim(sRow)
        sSecondRow = vbNullString
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function pvGetFiscalRow(uRow As UcsRow) As String
    Const FUNC_NAME     As String = "pvGetFiscalRow"
    Const MAX_WIDTH     As Long = 25
    Dim lIdx            As Long
    
    On Error GoTo EH
    With uRow
        '--- split item
        pvGetFiscalRow = RTrim(Replace(uRow.Item, vbTab, vbNullString))
        If Len(pvGetFiscalRow) > MAX_WIDTH Then
            For lIdx = MAX_WIDTH To 0 Step -1
                If pvIsWhitespace(Mid(pvGetFiscalRow, lIdx, 1)) Then
                    Exit For
                End If
            Next
            If lIdx <= 0 Then
                lIdx = MAX_WIDTH
            End If
            pvGetFiscalRow = Left(pvGetFiscalRow, lIdx) & vbLf & Mid(pvGetFiscalRow, lIdx + 1, MAX_WIDTH)
        End If
        '--- append rest
        pvGetFiscalRow = pvGetFiscalRow & vbTab & .Group & Replace(Format(C_Dbl(.Price), "0.00"), ",", ".")
        If LenB(.Quantity) <> 0 Then
            pvGetFiscalRow = pvGetFiscalRow & "*" & Replace(Format(C_Dbl(.Quantity), "0.000"), ",", ".")
        End If
    End With
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
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

Private Function pvIsWhitespace(sText As String)
    Select Case Asc(sText)
    Case 32, 95 '--- " ", "_"
        pvIsWhitespace = True
    End Select
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
    C_Dbl = CDbl(Replace(C_Str(v), ".", pvGetDecimalSeparator()))
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
        m_sLog = m_sLog & m_oFP.StatusText & vbCrLf & m_oFP.DipText & vbCrLf & m_oFP.MemoryText & vbCrLf
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub
