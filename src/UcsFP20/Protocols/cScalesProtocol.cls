VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cScalesProtocol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=========================================================================
'
' UcsFP20 (c) 2008-2023 by Unicontsoft
'
' Unicontsoft Fiscal Printers Component 2.0
'
' This project is licensed under the terms of the MIT license
' See the LICENSE file in the project root for more information
'
'=========================================================================
'
' Implementation of scales protocol
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "cScalesProtocol"
Implements IDeviceProtocol

'=========================================================================
' Constant and member variables
'=========================================================================

Private Const STR_INTERNAL              As String = "Invalid model group (%1)|Critical error: %1|Timeout waiting for device|Zero weight|Overload"
Private Const DEF_CONNECTIONTIMEOUT     As Long = 3000
Private Const DEF_TIMEOUT               As Long = 5000
Private Const DEF_CODEPAGE              As Long = 1251
Private Const DEF_MODEL_GROUP           As String = "BIMCO"
'--- numeric
Private Const MAX_RETRY                 As Long = 3
Private Const DBL_EPSILON               As Double = 0.000001

Private m_uPrintData                As UcsProtocolPrintData
Private m_uConfig                   As UcsConfigValues
Private m_uData                     As UcsScaleDataType

Private Enum UcsInternalErrors
    ucsErrInvalidModelGroup
    ucsErrCriticalError
    ucsErrDeviceTimeout
    ucsErrZeroWeight
    ucsErrOverload
End Enum

Private Type UcsConfigValues
    CodePage            As Long
    ModelGroup          As String
    AllowUnstable       As Boolean
    AllowZero           As Boolean
    DeviceSerialNo      As String
    LocalizedText(0 To [_ucsFscLciMax] - 1) As Variant
End Type

Private Type UcsScaleDataType
    Protocol            As UcsScaleProtocolEnum
    Received            As String
    Request             As String
    Response            As String
    Status              As UcsScaleStatusEnum
    Weight              As Double
    ConnectionTimeout   As Long
    CommandTimeout      As Long
    Device              As String
    Connector           As IDeviceConnector
End Type

Private Enum UcsScaleProtocolEnum
    ucsScaProtocolCas
    ucsScaProtocolElicom
    ucsScaProtocolDibal
    ucsScaProtocolMettler
    ucsScaProtocolDelmac
    ucsScaProtocolBimco
End Enum

Private Enum UcsScaleStatusEnum
    ucsScaStatusStable
    ucsScaStatusUnstable
    ucsScaStatusOverload
    ucsScaStatusUnderload
End Enum

Private Enum UcsControlSymbols
    SOH = 1
    STX = 2
    ETX = 3
    EOT = 4
    ENQ = 5
    ACK = 6
    NAK = &H15
    DC1 = &H11
    BMK_REQ = &H36
    ELI_REQ = &HAA
    ELI_UNS = &HBB
End Enum

Private Enum UcsParseResultEnum
    ucsScaResultContinue = 0     '--- must be 0
    ucsScaResultHasResult
    ucsScaResultRetrySend
    ucsScaResultRetryZero
End Enum

'=========================================================================
' Error handling
'=========================================================================

Private Sub RaiseError(sFunction As String)
    Debug.Print "Critical error: " & Err.Description & " [" & MODULE_NAME & "." & sFunction & "]"
    Logger.Log vbLogEventTypeError, MODULE_NAME, sFunction & "(" & Erl & ")", Err.Description
    Err.Raise Err.Number, MODULE_NAME & "." & sFunction & "(" & Erl & ")" & vbCrLf & Err.Source, Err.Description
End Sub

Private Sub DebugLog(sFunction As String, sText As String)
    Logger.Log vbLogEventTypeDebug, MODULE_NAME, sFunction, sText
End Sub

Private Sub DebugDataDump(sFunction As String, sPrefix As String, sData As String)
    Logger.DataDump MODULE_NAME, sFunction, sPrefix, sData
End Sub

'=========================================================================
' Properties
'=========================================================================

Property Get LastError() As String
    LastError = m_uPrintData.LastError
End Property

Property Get LocalizedText(ByVal eIdx As UcsFiscalLocalizedIndexesEnum) As String
    If IsArray(m_uConfig.LocalizedText(eIdx)) Then
        LocalizedText = Join(m_uConfig.LocalizedText(eIdx), "|")
    End If
End Property

Property Let LocalizedText(ByVal eIdx As UcsFiscalLocalizedIndexesEnum, sValue As String)
    m_uConfig.LocalizedText(eIdx) = Split(sValue, "|")
End Property

Property Get IsConnected() As Boolean
    IsConnected = Not m_uData.Connector Is Nothing
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function Init(sDeviceString As String) As Boolean
    Const FUNC_NAME     As String = "Init"
    Dim oOptions        As Object
    
    On Error GoTo EH
    DebugLog FUNC_NAME, "sDeviceString=" & sDeviceString
    Set oOptions = ParseDeviceString(sDeviceString)
    m_uData.ConnectionTimeout = C_Lng(JsonValue(oOptions, "ConnectionTimeout"))
    If m_uData.ConnectionTimeout <= 0 Then
        m_uData.ConnectionTimeout = DEF_CONNECTIONTIMEOUT
    End If
    m_uData.CommandTimeout = C_Lng(JsonValue(oOptions, "Timeout"))
    If m_uData.CommandTimeout <= 0 Then
        m_uData.CommandTimeout = DEF_TIMEOUT
    End If
    m_uData.Device = ToConnectorDevice(oOptions, 0, Me)
    If Not Connect() Then
        GoTo QH
    End If
    With m_uConfig
        .CodePage = Znl(C_Lng(JsonValue(oOptions, "CodePage")), DEF_CODEPAGE)
        .ModelGroup = UCase$(Zn(C_Str(JsonValue(oOptions, "ModelGroup")), DEF_MODEL_GROUP))
        .AllowUnstable = JsonBoolItem(oOptions, "AllowUnstable")
        .AllowZero = JsonBoolItem(oOptions, "AllowZero", .AllowUnstable)
        .DeviceSerialNo = Zn(C_Str(JsonValue(oOptions, "DeviceSerialNo")), "SCALES_" & At(Split(m_uData.Device, ","), 0))
    End With
    Select Case m_uConfig.ModelGroup
    Case "ELICOM"
        m_uData.Protocol = ucsScaProtocolElicom
    Case "DIBAL"
        m_uData.Protocol = ucsScaProtocolDibal
    Case "METTLER"
        m_uData.Protocol = ucsScaProtocolMettler
    Case "DELMAC"
        m_uData.Protocol = ucsScaProtocolDelmac
    Case "BIMCO"
        m_uData.Protocol = ucsScaProtocolBimco
    Case "CAS"
        m_uData.Protocol = ucsScaProtocolCas
    End Select
    If m_uData.Protocol = 0 Then
        On Error GoTo 0
        Err.Raise vbObjectError, , Printf(pvInternal(ucsErrInvalidModelGroup), Zn(m_uConfig.ModelGroup, STR_NONE))
    End If
    '--- success
    Init = True
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function Connect() As Boolean
    Const FUNC_NAME     As String = "Connect"
    
    On Error GoTo EH
    Set m_uData.Connector = InitDeviceConnector(m_uData.Device, m_uData.ConnectionTimeout, LocalizedText(ucsFscLciConnectorErrors), m_uPrintData.LastError)
    If m_uData.Connector Is Nothing Then
        DebugLog FUNC_NAME, "InitDeviceConnector failed, m_uPrintData.LastError=" & m_uPrintData.LastError
        GoTo QH
    End If
    '--- success
    Connect = True
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function Disconnect() As Boolean
    Const FUNC_NAME     As String = "Disconnect"
    
    On Error GoTo EH
    DebugLog FUNC_NAME, "TypeOf m_oConnector=" & TypeName(m_uData.Connector)
    Set m_uData.Connector = Nothing
    '--- success
    Disconnect = True
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetDeviceModel() As String
    GetDeviceModel = m_uConfig.ModelGroup & "_" & m_uData.Protocol
End Function

Public Function GetDeviceStatus(StatusText As String) As UcsFiscalStatusCodeEnum
    Dim dblTimerEx      As Double
    
    dblTimerEx = TimerEx + m_uData.CommandTimeout / 1000#
    Do While TimerEx < dblTimerEx
        If Not pvReadWeight(dblTimerEx - TimerEx, MAX_RETRY, m_uConfig.AllowZero, m_uPrintData.LastError) Then
            StatusText = m_uPrintData.LastError
            GetDeviceStatus = ucsFscStaFailed
            GoTo QH
        End If
        If Abs(m_uData.Weight) > DBL_EPSILON Or m_uConfig.AllowZero Then
            StatusText = Trim$(Replace(Replace(Str$(m_uData.Weight), " .", "0."), "-.", "-0."))
        End If
        If m_uData.Status = ucsScaStatusStable Then
            GetDeviceStatus = ucsFscStaReady
            GoTo QH
        ElseIf m_uConfig.AllowUnstable Then
            If m_uData.Status = ucsScaStatusOverload Then
                StatusText = pvInternal(ucsErrOverload)
            End If
            GetDeviceStatus = ucsFscStaBusy
            GoTo QH
        End If
    Loop
    StatusText = pvInternal(IIf(m_uData.Status = ucsScaStatusOverload, ucsErrOverload, ucsErrDeviceTimeout))
    GetDeviceStatus = ucsFscStaFailed
QH:
    m_uPrintData.LastError = vbNullString
End Function

'= private ===============================================================

Private Function pvInternal(ByVal lIdx As UcsInternalErrors) As String
    pvInternal = At(m_uConfig.LocalizedText(ucsFscLciInternalErrors), lIdx)
    If LenB(pvInternal) = 0 Then
        pvInternal = At(Split(STR_INTERNAL, "|"), lIdx)
    End If
End Function

Private Function pvReadWeight( _
            ByVal dblTimeout As Double, _
            ByVal lMaxRetry As Long, _
            ByVal bAllowZero As Boolean, _
            sError As String) As Boolean
    Const FUNC_NAME     As String = "pvReadWeight"
    Dim lRetry          As Long
    Dim dblTimerEx      As Double
    Dim eResult         As UcsParseResultEnum
    Dim bHasZero        As Boolean
    Dim baText()        As Byte
    Dim sBuffer         As String
    
    On Error GoTo EH
    For lRetry = 1 To lMaxRetry
        dblTimerEx = TimerEx + dblTimeout
        Select Case m_uData.Protocol
        Case ucsScaProtocolCas
            m_uData.Request = Chr$(ENQ)
        Case ucsScaProtocolElicom
            m_uData.Request = Chr$(ELI_REQ)
        Case ucsScaProtocolDibal
            m_uData.Request = Chr$(STX)
        Case ucsScaProtocolMettler
            m_uData.Request = "SI" & vbCrLf
        Case ucsScaProtocolDelmac
            m_uData.Request = vbNullString
        Case ucsScaProtocolBimco
            m_uData.Request = Chr$(BMK_REQ)
        End Select
        m_uData.Received = vbNullString
        If LenB(m_uData.Request) <> 0 Then
            DebugDataDump FUNC_NAME, "[SEND ] ", m_uData.Request
            If Not m_uData.Connector.WriteData(ToMultiByte(m_uData.Request, m_uConfig.CodePage), 100) Then
                DebugLog FUNC_NAME, "WriteData failed, GetLastError=" & m_uData.Connector.GetLastError()
            End If
        End If
        Do While TimerEx < dblTimerEx
            If m_uData.Connector.ReadData(baText, 100) Then
                sBuffer = FromMultiByte(baText, m_uConfig.CodePage)
                DebugDataDump FUNC_NAME, "[RECV ] ", sBuffer
                m_uData.Received = m_uData.Received & sBuffer
                Select Case m_uData.Protocol
                Case ucsScaProtocolCas
                    eResult = pvParseCasResponse(m_uData, bAllowZero)
                Case ucsScaProtocolElicom
                    eResult = pvParseElicomResponse(m_uData, bAllowZero)
                Case ucsScaProtocolDibal
                    eResult = pvParseDibalResponse(m_uData, bAllowZero)
                Case ucsScaProtocolMettler
                    eResult = pvParseMettlerResponse(m_uData, bAllowZero)
                Case ucsScaProtocolDelmac
                    eResult = pvParseDelmacResponse(m_uData, bAllowZero)
                Case ucsScaProtocolBimco
                    eResult = pvParseBimcoResponse(m_uData, bAllowZero)
                Case Else
                    eResult = ucsScaResultContinue
                End Select
                Select Case eResult
                Case ucsScaResultHasResult
                    '--- success
                    pvReadWeight = True
                    GoTo QH
                Case ucsScaResultRetrySend, ucsScaResultRetryZero
                    If eResult = ucsScaResultRetryZero Then
                        bHasZero = True
                    End If
                    m_uData.Received = vbNullString
                    If LenB(m_uData.Request) <> 0 Then
                        DebugDataDump FUNC_NAME, "[SEND ] ", m_uData.Request
                        m_uData.Connector.WriteData ToMultiByte(m_uData.Request, m_uConfig.CodePage), m_uData.CommandTimeout
                    End If
                End Select
            End If
        Loop
    Next
    sError = pvInternal(IIf(bHasZero, ucsErrZeroWeight, ucsErrDeviceTimeout))
QH:
    Exit Function
EH:
    sError = Printf(pvInternal(ucsErrCriticalError), Err.Description)
    Resume QH
End Function

Private Function pvParseCasResponse(uData As UcsScaleDataType, ByVal bAllowZero As Boolean) As UcsParseResultEnum
    Dim lStart          As Long
    Dim lEnd            As Long
    Dim lSign           As Long
    
    With uData
        If InStr(.Received, Chr$(NAK)) > 0 Then
            .Request = Chr$(ENQ)
            pvParseCasResponse = ucsScaResultRetrySend
        ElseIf InStr(.Received, Chr$(ACK)) > 0 And .Request = Chr$(ENQ) Then
            .Request = Chr$(DC1)
            pvParseCasResponse = ucsScaResultRetrySend
        Else
            lStart = InStr(.Received, Chr$(SOH) & Chr$(STX))
            lEnd = InStr(.Received, Chr$(ETX) & Chr$(EOT))
            If lStart > 0 And lEnd > lStart + 2 Then
                .Response = Replace(Mid$(.Received, lStart + 2, lEnd - lStart - 2), Chr$(ACK), vbNullString)
                .Status = ucsScaStatusStable
                .Weight = 0
                Select Case Left$(.Response, 1)
                Case "F"
                    .Status = ucsScaStatusOverload
                Case "U"
                    .Status = ucsScaStatusUnstable
                End Select
                lSign = IIf(Mid$(.Response, 2, 1) = "-", -1, 1)
                .Weight = lSign * Val(Mid$(.Response, 3))
                If .Status = ucsScaStatusStable And Abs(.Weight) < DBL_EPSILON And Not bAllowZero Then
                    .Request = Chr$(ENQ)
                    pvParseCasResponse = ucsScaResultRetryZero
                Else
                    If lSign < 0 Then
                        .Status = ucsScaStatusUnderload
                    End If
                    '--- success
                    pvParseCasResponse = ucsScaResultHasResult
                End If
            ElseIf lEnd > 0 Then
                .Received = Mid$(.Received, lEnd + 2)
            End If
        End If
    End With
End Function

Private Function pvParseElicomResponse(uData As UcsScaleDataType, ByVal bAllowZero As Boolean) As UcsParseResultEnum
    Dim lIdx            As Long
    Dim lChar           As Long
    Dim lSum            As Long
    
    With uData
        If InStr(.Received, Chr$(ELI_UNS)) > 0 Then
            If Not bAllowZero Then
                pvParseElicomResponse = ucsScaResultRetryZero
            Else
                .Status = ucsScaStatusUnstable
                '--- success
                pvParseElicomResponse = ucsScaResultHasResult
            End If
        ElseIf Len(.Received) >= 4 Then
            .Response = vbNullString
            For lIdx = 1 To 3
                lChar = Asc(Mid$(.Received, lIdx, 1))
                lSum = (lSum + lChar) And &HFF&
                .Response = .Response & Right$("0" & Hex$(lChar), 2)
            Next
            If lSum <> Asc(Mid$(.Received, 4, 1)) Then
                pvParseElicomResponse = ucsScaResultRetrySend
            Else
                .Status = ucsScaStatusStable
                .Weight = Val(.Response) / 1000#
                '--- success
                pvParseElicomResponse = ucsScaResultHasResult
            End If
        End If
    End With
End Function

Private Function pvParseDibalResponse(uData As UcsScaleDataType, ByVal bAllowZero As Boolean) As UcsParseResultEnum
    Dim lEnd            As Long

    With uData
        If InStr(.Received, Chr$(ACK)) > 0 Then
            .Request = "10" & vbCrLf
            pvParseDibalResponse = ucsScaResultRetrySend
        ElseIf InStr(.Received, Chr$(STX)) > 0 Then
            .Request = Chr$(ACK)
            pvParseDibalResponse = ucsScaResultRetrySend
        Else
            lEnd = InStr(.Received, " ")
            If lEnd > 0 Then
                .Response = Mid$(.Received, 1, lEnd - 1)
                If Not IsNumeric(.Response) And Not bAllowZero Then
                    .Request = Chr$(STX)
                    pvParseDibalResponse = ucsScaResultRetryZero
                Else
                    .Status = IIf(IsNumeric(.Response), ucsScaStatusStable, ucsScaStatusUnstable)
                    .Weight = Val(.Response) / 1000#
                    If .Status = ucsScaStatusStable And .Weight < -DBL_EPSILON Then
                        .Status = ucsScaStatusUnderload
                    End If
                    '--- success
                    pvParseDibalResponse = ucsScaResultHasResult
                End If
            End If
        End If
    End With
End Function

Private Function pvParseMettlerResponse(uData As UcsScaleDataType, ByVal bAllowZero As Boolean) As UcsParseResultEnum
    Dim lStart          As Long
    Dim lEnd            As Long
    Dim lSign           As Long

    With uData
        lStart = InStr(.Received, "S ")
        lEnd = InStr(.Received, " Kg")
        If lStart > 0 And lEnd > lStart + 2 Then
            .Response = Mid$(.Received, lStart + 2, lEnd - lStart - 2)
            .Status = ucsScaStatusStable
            .Weight = 0
            Select Case Left$(.Response, 1)
            Case "X"
                .Status = ucsScaStatusOverload
            Case "D"
                .Status = ucsScaStatusUnstable
            End Select
            lSign = IIf(Mid$(.Response, 3, 1) = "-", -1, 1)
            .Weight = lSign * Val(Mid$(.Response, 4))
            If .Status = ucsScaStatusStable And Abs(.Weight) < DBL_EPSILON And Not bAllowZero Then
                pvParseMettlerResponse = ucsScaResultRetryZero
            Else
                If lSign < 0 Then
                    .Status = ucsScaStatusUnderload
                End If
                '--- success
                pvParseMettlerResponse = ucsScaResultHasResult
            End If
        ElseIf lEnd > 0 Then
            .Received = Mid$(.Received, lEnd + 3)
        End If
    End With
End Function

Private Function pvParseDelmacResponse(uData As UcsScaleDataType, ByVal bAllowZero As Boolean) As UcsParseResultEnum
    Const FLAG_UNDRLOAD As Long = &H80
    Const FLAG_OVERLOAD As Long = &H40
    Const FLAG_STABLE   As Long = &H10
    Const FLAG_ZERO     As Long = 8
    Const MASK_SCALE    As Long = 7
    Const NUM_DIGITS    As Long = 5
    Dim lStart          As Long
    Dim lEnd            As Long
    Dim lControl        As Long
    
    With uData
        lStart = InStr(.Received, Chr$(STX))
        lEnd = InStr(.Received, Chr$(ETX))
        If lStart > 0 And lEnd > lStart + 1 Then
            .Response = Mid$(.Received, lStart + 1, lEnd - lStart - 1)
            lControl = Asc(Right$(.Response, 1))
            If (lControl And FLAG_ZERO) <> 0 And Not bAllowZero Then
                pvParseDelmacResponse = ucsScaResultRetryZero
            Else
                If (lControl And FLAG_UNDRLOAD) <> 0 Then
                    .Status = ucsScaStatusUnderload
                ElseIf (lControl And FLAG_OVERLOAD) <> 0 Then
                    .Status = ucsScaStatusOverload
                ElseIf (lControl And FLAG_STABLE) <> 0 Then
                    .Status = ucsScaStatusStable
                Else
                    .Status = ucsScaStatusUnstable
                End If
                .Weight = Val(.Response) / (10 ^ (NUM_DIGITS - (lControl And MASK_SCALE)))
                '--- success
                pvParseDelmacResponse = ucsScaResultHasResult
            End If
        ElseIf lEnd > 0 Then
            .Received = Mid$(.Received, lEnd + 1)
        End If
    End With
End Function

Private Function pvParseBimcoResponse(uData As UcsScaleDataType, ByVal bAllowZero As Boolean) As UcsParseResultEnum
    Const IDX_WIEGHT    As Long = 0
    Const IDX_POWER10   As Long = 3
    Const IDX_STATUS    As Long = 4
    Const IDX_ERROR     As Long = 5
    Const IDX_CRC       As Long = 6
    Const FLAG_STABLE   As Long = 1
    Const FLAG_UNDRLOAD As Long = 4
    Dim lStart          As Long
    Dim lIdx            As Long
    Dim lSum            As Long
    Dim baRecv()        As Byte
    
    With uData
        lStart = Len(.Received) - 7
        If lStart > 0 Then
            .Response = Mid$(.Received, lStart)
            baRecv = StrConv(.Response, vbFromUnicode)
            For lIdx = 0 To IDX_CRC - 1
                lSum = lSum + baRecv(lIdx)
            Next
            If lSum <> (baRecv(IDX_CRC) * &H100& Or baRecv(IDX_CRC + 1)) Or baRecv(IDX_POWER10) > 10 Then
                pvParseBimcoResponse = ucsScaResultRetrySend
            Else
                If baRecv(IDX_ERROR) <> 0 Then
                    .Status = ucsScaStatusOverload
                ElseIf (baRecv(IDX_STATUS) And FLAG_STABLE) <> 0 Then
                    .Status = ucsScaStatusStable
                ElseIf (baRecv(IDX_STATUS) And FLAG_UNDRLOAD) = 0 Then
                    .Status = ucsScaStatusUnderload
                Else
                    .Status = ucsScaStatusUnstable
                End If
                .Weight = (IIf(baRecv(IDX_WIEGHT) And &H80, &HFF000000, 0) _
                    Or baRecv(IDX_WIEGHT) * &H10000 _
                    Or baRecv(IDX_WIEGHT + 1) * &H100& _
                    Or baRecv(IDX_WIEGHT + 2)) / (10 ^ baRecv(IDX_POWER10))
                If .Status = ucsScaStatusStable And Abs(.Weight) < DBL_EPSILON And Not bAllowZero Then
                    pvParseBimcoResponse = ucsScaResultRetryZero
                Else
                    '--- success
                    pvParseBimcoResponse = ucsScaResultHasResult
                End If
            End If
        End If
    End With
End Function

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Initialize()
    m_uData.ConnectionTimeout = DEF_CONNECTIONTIMEOUT
    m_uData.CommandTimeout = DEF_TIMEOUT
    LocalizedText(ucsFscLciInternalErrors) = STR_INTERNAL
    m_uConfig.CodePage = DEF_CODEPAGE
End Sub

'=========================================================================
' IDeviceProtocol interface
'=========================================================================

Private Function IDeviceProtocol_AddBarcode(ByVal BarcodeType As UcsFiscalBarcodeTypeEnum, Text As String, Optional ByVal Height As Long) As Boolean
    
End Function

Private Function IDeviceProtocol_AddDiscount(ByVal DiscType As UcsFiscalDiscountTypeEnum, ByVal Value As Double) As Boolean
    
End Function

Private Function IDeviceProtocol_AddLine(Line As String, Optional Command As String, Optional ByVal WordWrap As Boolean = True) As Boolean
    
End Function

Private Function IDeviceProtocol_AddPayment(ByVal PmtType As UcsFiscalPaymentTypeEnum, Optional PmtName As String, Optional ByVal Amount As Double, Optional ByVal Rate As Double) As Boolean
    
End Function

Private Function IDeviceProtocol_AddPLU(Name As String, ByVal Price As Double, Optional ByVal Quantity As Double = 1, Optional ByVal TaxGroup As Long = 2, Optional UnitOfMeasure As String, Optional ByVal DepartmentNo As Long) As Boolean
    
End Function

Private Function IDeviceProtocol_AutodetectDevices(Ports As Variant) As Variant
    IDeviceProtocol_AutodetectDevices = Ports
End Function

Private Function IDeviceProtocol_CancelReceipt() As Boolean
    IDeviceProtocol_CancelReceipt = True
End Function

Private Function IDeviceProtocol_PrintServiceDeposit(OperatorCode As String, OperatorPassword As String, ByVal Amount As Double, Optional ReceiptNo As String, Optional ReceiptDate As Date, Optional ReceiptAmount As Double) As Variant
    
End Function

Private Function IDeviceProtocol_Connect() As Boolean
    IDeviceProtocol_Connect = Connect()
End Function

Private Function IDeviceProtocol_PrintDuplicate(Optional DuplicateNo As String, Optional ReceiptNo As String, Optional ReceiptDate As Date) As Boolean
    
End Function

Private Function IDeviceProtocol_Disconnect() As Boolean
    IDeviceProtocol_Disconnect = Disconnect()
End Function

Private Function IDeviceProtocol_PrintReceipt(Optional ResumeToken As String, Optional ReceiptNo As String, Optional ReceiptDate As Date, Optional ReceiptAmount As Double, Optional InvoiceNo As String) As Boolean
    
End Function

Private Function IDeviceProtocol_GetCharsPerLine() As Long
    
End Function

Private Function IDeviceProtocol_GetCommentTextMaxLength() As Long
    
End Function

Private Function IDeviceProtocol_GetItemNameMaxLength() As Long
    
End Function

Private Function IDeviceProtocol_GetClock() As Date
    
End Function

Private Function IDeviceProtocol_GetCommandLog() As String
    
End Function

Private Function IDeviceProtocol_GetOperatorName(OperatorCode As String) As String
    
End Function

Private Function IDeviceProtocol_GetDefaultPassword(OperatorCode As String) As String

End Function

Private Function IDeviceProtocol_GetDeviceModel() As String
    IDeviceProtocol_GetDeviceModel = GetDeviceModel
End Function

Private Function IDeviceProtocol_GetDeviceFirmware() As String
    
End Function

Private Function IDeviceProtocol_GetDeviceProtocol() As String
    IDeviceProtocol_GetDeviceProtocol = STR_PROTOCOL_SCALES
End Function

Private Function IDeviceProtocol_GetDeviceStatus(StatusText As String) As UcsFiscalStatusCodeEnum
    IDeviceProtocol_GetDeviceStatus = GetDeviceStatus(StatusText)
End Function

Private Function IDeviceProtocol_GetFiscalMemoryNo() As String

End Function

Private Function IDeviceProtocol_GetFooterText(ByVal Index As Long) As String

End Function

Private Function IDeviceProtocol_GetHeaderText(ByVal Index As Long) As String

End Function

Private Function IDeviceProtocol_GetLastError(Optional ErrNo As UcsFiscalErrorsEnum) As String
    ErrNo = m_uPrintData.LastErrNo
    IDeviceProtocol_GetLastError = m_uPrintData.LastError
End Function

Private Function IDeviceProtocol_GetLastReceiptNo() As String

End Function

Private Function IDeviceProtocol_GetLastInvoiceNo() As String

End Function

Private Function IDeviceProtocol_GetLastQRCodeInfo(Optional LastReceiptDate As Date) As String

End Function

Private Function IDeviceProtocol_GetPaymentName(ByVal Index As Long) As String

End Function

Private Function IDeviceProtocol_GetResumeToken() As String

End Function

Private Function IDeviceProtocol_GetDeviceSerialNo() As String
    IDeviceProtocol_GetDeviceSerialNo = m_uConfig.DeviceSerialNo
End Function

Private Function IDeviceProtocol_GetTaxCaption() As String

End Function

Private Function IDeviceProtocol_GetTaxNo() As String

End Function

Private Function IDeviceProtocol_GetTotalsByPaymentTypes() As Variant

End Function

Private Function IDeviceProtocol_GetTotalsByTaxGroups() As Variant

End Function

Private Function IDeviceProtocol_GetTotalsStats() As Variant

End Function

Private Function IDeviceProtocol_Init(DeviceString As String, Optional EventSink As Object) As Boolean
    IDeviceProtocol_Init = Init(DeviceString)
End Function

Private Property Get IDeviceProtocol_IsConnected() As Boolean
    IDeviceProtocol_IsConnected = IsConnected
End Property

Private Function IDeviceProtocol_OpenDrawer() As Boolean
    IDeviceProtocol_OpenDrawer = True
End Function

Private Function IDeviceProtocol_PrintMonthlyReport(ByVal ReportType As UcsFiscalReportsTypeEnum, ByVal StartDate As Date, ByVal EndDate As Date, Optional ReceiptNo As String, Optional ReceiptDate As Date) As Boolean
    
End Function

Private Function IDeviceProtocol_PrintDailyXReport(ByVal ReportType As UcsFiscalReportsTypeEnum, Optional ReceiptNo As String, Optional ReceiptDate As Date) As Boolean
    
End Function

Private Function IDeviceProtocol_PrintDailyZReport(ByVal ReportType As UcsFiscalReportsTypeEnum, Optional ReceiptNo As String, Optional ReceiptDate As Date) As Boolean
    
End Function

Private Function IDeviceProtocol_SetClock(ByVal NewDate As Date) As Boolean

End Function

Private Sub IDeviceProtocol_SetLocalizedText(ByVal Index As UcsFiscalLocalizedIndexesEnum, Text As String, Optional ByVal ConfigIndex As Long)
    LocalizedText(Index) = Text
End Sub

Private Function IDeviceProtocol_StartReceipt(ByVal ReceiptType As UcsFiscalReceiptTypeEnum, OperatorCode As String, OperatorName As String, OperatorPassword As String, Optional UniqueSaleNo As String, Optional TableNo As String, Optional ByVal DisablePrinting As Boolean, Optional InvDocNo As String, Optional InvCgTaxNo As String, Optional ByVal InvCgTaxNoType As UcsFiscalTaxNoTypeEnum, Optional InvCgVatNo As String, Optional InvCgName As String, Optional InvCgCity As String, Optional InvCgAddress As String, Optional InvCgPrsReceive As String, Optional OwnData As String, Optional ByVal RevType As UcsFiscalReversalTypeEnum, Optional RevReceiptNo As String, Optional ByVal RevReceiptDate As Date, Optional RevFiscalMemoryNo As String, Optional RevInvoiceNo As String, Optional RevReason As String) As Boolean
    
End Function

Private Function IDeviceProtocol_GetDepartmentName(DepartmentCode As String, Optional TaxGroup As Long) As String
    
End Function

Private Function IDeviceProtocol_GetTaxRates() As Variant

End Function
