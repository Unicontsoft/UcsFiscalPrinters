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
' $Header: /UcsFiscalPrinter/Src/cFiscalPrinter.cls 1     14.02.11 18:13 Wqw $
'
'   Unicontsoft Fiscal Printers Project
'   Copyright (c) 2008-2011 Unicontsoft
'
'   Uprawlenie na fiscalni ustrojstwa prez IDeviceProtocol protocoli
'
' $Log: /UcsFiscalPrinter/Src/cFiscalPrinter.cls $
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

Public Enum UcsFiscalReportTypeEnum
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
    ucsFscPixFiscalPayment
End Enum

'=========================================================================
' Constants and member variables
'=========================================================================

'Private Const STR_PAYMENT               As String = "В БРОЙ|КАРТА|С ЧЕК|КРЕДИТ"
Private Const STR_ELTRADE_ECR_INTERNAL  As String = "Не е указано устройство|Грешка при отваряне: |Грешка при SetCommTimeouts: |Грешка при BuildCommDCB: |Грешка при SetCommState: |Грешка при WriteFile: |Време за достъп изтече в очакване на отговор|Грешка при ReadFile: |Вече има отворена връзка|Няма отворена връзка|Невалидно заглавие на отговор от устройството|Грешка при WaitCommEvent: |Няма започната бележка|Липсва хартия или друга фатална грешка|Невалиден параметър 'Command'|Непозволена команда за печат на повече от един дубликат"
Private Const ERR_INIT_PROTOCOL         As String = "Грешка при инициализация на протокол %1"
Private Const ERR_MISSING_ROWS          As String = "Липсват редове за печат"
Private Const STR_DUPLICATE             As String = "ДУБЛИКАТ"
Private Const STR_ORIGINAL              As String = "ОРИГИНАЛ"
Private Const STR_HEAD_CREDIT_NOTE2     As String = "към ф-ра "
Private Const STR_CASHIER               As String = "Касиер"

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

'Private Sub PrintError(sFunc As String)
'    Debug.Print MODULE_NAME & "." & sFunc & ": " & Error
'    DebugLog sFunc, "Run-time error: " & Error
'End Sub

Private Sub DebugLog(sFunc As String, sText As String)
    OutputDebugLog MODULE_NAME, sFunc, sText
End Sub

'=========================================================================
' Methods
'=========================================================================

Public Function EnumPorts() As Variant
    EnumPorts = EnumSerialPorts
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
    Set oFP = pvInit(DeviceString)
    vInput = RetVal
    ReDim RetVal(0 To 18) As Variant
    If LenB(At(vInput, 0, "1")) <> 0 Then
        RetVal(0) = Trim(oFP.GetHeaderText(1))
    End If
    If LenB(At(vInput, 1, "1")) <> 0 Then
        RetVal(1) = Trim(oFP.GetHeaderText(2))
    End If
    If LenB(At(vInput, 2, "1")) <> 0 Then
        RetVal(2) = oFP.GetTaxNumber() & "," & Replace(Trim(oFP.GetTaxCaption()), ":", vbNullString)
    End If
    If LenB(At(vInput, 3, "1")) <> 0 Then
        RetVal(3) = Trim(oFP.GetHeaderText(3))
    End If
    If LenB(At(vInput, 4, "1")) <> 0 Then
        RetVal(4) = Trim(oFP.GetHeaderText(4))
    End If
    If LenB(At(vInput, 5, "1")) <> 0 Then
        RetVal(5) = Trim(oFP.GetHeaderText(5))
    End If
    If LenB(At(vInput, 6, "1")) <> 0 Then
        RetVal(6) = Trim(oFP.GetHeaderText(6))
    End If
    If LenB(At(vInput, 7, "1")) <> 0 Then
        RetVal(7) = Trim(oFP.GetFooterText(1))
    End If
    If LenB(At(vInput, 8, "1")) <> 0 Then
        RetVal(8) = Trim(oFP.GetFooterText(2))
    End If
    If LenB(At(vInput, 10, "1")) <> 0 Then
        RetVal(10) = At(OperatorData, 0)
    End If
    If LenB(At(vInput, 11, "1")) <> 0 Then
        RetVal(11) = oFP.GetClock()
    End If
    If LenB(At(vInput, 12, "1")) <> 0 Then
        RetVal(12) = oFP.GetLastBonNumber()
    End If
    If LenB(At(vInput, 13, "1")) <> 0 Then
        RetVal(13) = oFP.GetPaymentType(1)
    End If
    If LenB(At(vInput, 14, "1")) <> 0 Then
        RetVal(14) = oFP.GetPaymentType(2)
    End If
    If LenB(At(vInput, 15, "1")) <> 0 Then
        RetVal(15) = oFP.GetPaymentType(3)
    End If
    If LenB(At(vInput, 16, "1")) <> 0 Then
        RetVal(16) = oFP.GetPaymentType(4)
    End If
    If LenB(At(vInput, 17, "1")) <> 0 Then
        RetVal(17) = oFP.GetCharsPerLine()
    End If
    If LenB(At(vInput, 18, "1")) <> 0 Then
        RetVal(18) = oFP.GetDefaultPassword()
    End If
    '--- success
    GetHeaderFooter = True
QH:
    CommandLog = oFP.GetCommandLog()
    Exit Function
EH:
    LastError = Err.Description
    DebugLog FUNC_NAME, "LastError=" & LastError
    If Not oFP Is Nothing Then
        DebugLog FUNC_NAME, "oFP.GetLastError=" & oFP.GetLastError()
    End If
    CommandLog = oFP.GetCommandLog()
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
            
    On Error GoTo EH
    Set oFP = pvInit(DeviceString)
    If IsArray(Rows) Then
        oFP.StartBon ReceiptType, True, At(OperatorData, ucsFscPixOperatorName, STR_CASHIER), _
            At(InvoiceData, ucsFscPixInvoiceNoPrefix) & At(InvoiceData, ucsFscPixInvoiceNo), At(InvoiceData, ucsFscPixCgBulstat), _
            At(InvoiceData, ucsFscPixCgVatNo), At(InvoiceData, ucsFscPixCgName), _
            vbNullString, At(InvoiceData, ucsFscPixCgAddress), _
            At(InvoiceData, ucsFscPixCgPersonReceiver), vbNullString
        If LenB(At(InvoiceData, ucsFscPixLnkDoc)) <> 0 Then
            oFP.AddLine STR_HEAD_CREDIT_NOTE2 & At(InvoiceData, ucsFscPixLnkDoc)
        End If
        If PrintDuplicate Then
            oFP.AddLine STR_DUPLICATE
        ElseIf ReceiptType = ucsFscRetInvoiceNonFiscal Then
            oFP.AddLine STR_ORIGINAL
        End If
        For Each vElem In Rows
            With pvGetRowData(vElem)
                Select Case .Type
                Case ucsRwtSell
                    oFP.AddPLU 0, .Item, C_Dbl(.Price), IIf(LenB(.Quantity) <> 0, C_Dbl(.Quantity), 1), pvGetVatGroup(.Group)
                    If C_Dbl(.Discount) <> 0 Then
                        oFP.AddDiscount ucsFscDstPlu, -C_Dbl(.Discount)
                    End If
                Case ucsRwtText
                    oFP.AddLine .Item
                Case ucsRwtPayment
                    lIdx = LimitLong(C_Lng(.WayOfPayment), 1, 8)
                    oFP.AddPayment lIdx, Zn(.Item, oFP.GetPaymentType(lIdx)), C_Dbl(.Amount), 1
                End Select
            End With
        Next
        oFP.EndBon
        '--- success
        PrintReceipt = True
    End If
    If PrintDuplicate Then
        Select Case ReceiptType
        Case ucsFscRetNonfiscal, ucsFscRetInvoiceNonFiscal
            If IsEmpty(Rows) Then
                LastError = ERR_MISSING_ROWS
                GoTo QH
            End If
        Case Else
            If Not oFP.CopyLastBon() Then
                LastError = oFP.GetLastError()
                GoTo QH
            End If
            '--- success
            PrintReceipt = True
        End Select
    End If
    If LenB(At(OperatorData, ucsFscPixLastReceipt, vbNullString)) <> 0 Then
        OperatorData(ucsFscPixLastReceipt) = oFP.GetLastBonNumber()
    End If
    If LenB(At(OperatorData, ucsFscPixFiscalPayment, vbNullString)) <> 0 Then
        OperatorData(ucsFscPixFiscalPayment) = oFP.GetPaymentExecuted()
    End If
QH:
    CommandLog = oFP.GetCommandLog()
    Exit Function
EH:
    LastError = Err.Description
    DebugLog FUNC_NAME, "LastError=" & LastError
    If Not oFP Is Nothing Then
        DebugLog FUNC_NAME, "oFP.GetLastError=" & oFP.GetLastError()
    End If
    CommandLog = oFP.GetCommandLog()
End Function

Public Function PrintReport( _
            DeviceString As String, _
            ByVal ReportType As UcsFiscalReportTypeEnum, _
            Optional ReportData As Variant, _
            Optional LastError As String, _
            Optional CommandLog As String) As Boolean
    Const FUNC_NAME     As String = "PrintReport"
    Dim oFP             As IDeviceProtocol
    
    On Error GoTo EH
    Set oFP = pvInit(DeviceString)
    Select Case ReportType
    Case ucsFscRptDaily
        '-- ReportData = { IsClear, IsItems, IsDepartments }
        If C_Bool(At(ReportData, 1)) And C_Bool(At(ReportData, 2)) Then
            
        ElseIf C_Bool(At(ReportData, 1)) Then
            
        ElseIf C_Bool(At(ReportData, 2)) Then
            
        Else
            If C_Bool(At(ReportData, 0)) Then
                oFP.RunZReport
            Else
                oFP.RunXReport
            End If
        End If
    Case ucsFscRptNumber
        '-- ReportData = { FromNum, ToNum, IsDetailed }
    Case ucsFscRptDate
        '-- ReportData = { FromDate, ToDate, IsDetailed }
    Case ucsFscRptOperator
        '-- ReportData = Empty
        
    End Select
    '--- success
    PrintReport = True
QH:
    CommandLog = oFP.GetCommandLog()
    Exit Function
EH:
    LastError = Err.Description
    DebugLog FUNC_NAME, "LastError=" & LastError
    If Not oFP Is Nothing Then
        DebugLog FUNC_NAME, "oFP.GetLastError=" & oFP.GetLastError()
    End If
    CommandLog = oFP.GetCommandLog()
End Function

Public Function OpenDrawer( _
            DeviceString As String, _
            Optional LastError As String, _
            Optional CommandLog As String) As Boolean
    Const FUNC_NAME     As String = "OpenDrawer"
    Dim oFP             As IDeviceProtocol
    
    On Error GoTo EH
    Set oFP = pvInit(DeviceString)
    oFP.OpenDrawer
    '--- success
    OpenDrawer = True
QH:
    CommandLog = oFP.GetCommandLog()
    Exit Function
EH:
    LastError = Err.Description
    DebugLog FUNC_NAME, "LastError=" & LastError
    If Not oFP Is Nothing Then
        DebugLog FUNC_NAME, "oFP.GetLastError=" & oFP.GetLastError()
    End If
    CommandLog = oFP.GetCommandLog()
End Function

Public Function InitDateTime( _
            DeviceString As String, _
            sDateTime As String, _
            Optional LastError As String, _
            Optional CommandLog As String) As Boolean
    Const FUNC_NAME     As String = "InitDateTime"
    Dim oFP             As IDeviceProtocol
    
    On Error GoTo EH
    Set oFP = pvInit(DeviceString)
    If C_Date(sDateTime) = 0 Then
        oFP.SetClock Now
    Else
        oFP.SetClock C_Date(sDateTime)
    End If
    sDateTime = oFP.GetClock()
    '--- success
    InitDateTime = True
QH:
    CommandLog = oFP.GetCommandLog()
    Exit Function
EH:
    LastError = Err.Description
    DebugLog FUNC_NAME, "LastError=" & LastError
    If Not oFP Is Nothing Then
        DebugLog FUNC_NAME, "oFP.GetLastError=" & oFP.GetLastError()
    End If
    CommandLog = oFP.GetCommandLog()
End Function


'= private ===============================================================

Private Function pvInit(DeviceString As String) As IDeviceProtocol
    Dim vSplit          As Variant
    
    vSplit = Split(DeviceString, ";")
    '--- figure out model
    Select Case UCase(At(vSplit, 0))
    Case "ELTRADE ECR"
        Set pvInit = New cEltradeProtocol
        pvInit.SetLocalizedText ucsFscLciInternalErrors, STR_ELTRADE_ECR_INTERNAL
    Case Else
        Set pvInit = New cEltradeProtocol
        pvInit.SetLocalizedText ucsFscLciInternalErrors, STR_ELTRADE_ECR_INTERNAL
    End Select
    If Not pvInit.Init(At(vSplit, 1, "COM1"), C_Lng(At(vSplit, 2)), C_Lng(At(vSplit, 3))) Then
        Err.Raise vbObjectError, , Printf(ERR_INIT_PROTOCOL, At(vSplit, 0, STR_NONE))
    End If
End Function

Private Function pvGetRowData(vRow As Variant) As UcsRow
    On Error Resume Next
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
    On Error GoTo 0
End Function

Private Function pvGetVatGroup(sVatGroup As String) As Long
    If LenB(sVatGroup) <> 0 Then
        pvGetVatGroup = Asc(UCase(sVatGroup)) - 191
    End If
    If pvGetVatGroup < 1 Or pvGetVatGroup > 8 Then
        pvGetVatGroup = 2
    End If
End Function

Private Function pvGetPmtNumber(sWayOfPayment As String) As Long
    Select Case UCase(sWayOfPayment)
    Case "N" '--- credit
        pvGetPmtNumber = 4
    Case "C" '--- check
        pvGetPmtNumber = 3
    Case "D" '--- card
        pvGetPmtNumber = 2
    Case "P" '--- cash
        pvGetPmtNumber = 1
    Case Else
        pvGetPmtNumber = 1
    End Select
End Function

