VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLabelProtocol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=========================================================================
'
' UcsFP20 (c) 2008-2022 by Unicontsoft
'
' Unicontsoft Fiscal Printers Component 2.0
'
' This project is licensed under the terms of the MIT license
' See the LICENSE file in the project root for more information
'
'=========================================================================
'
' Implementation of label protocol for Datecs LP-50MX
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "cLabelProtocol"
Implements IDeviceProtocol

'=========================================================================
' Public events
'=========================================================================

Event CommandComplete(sData As String, sResult As String)

'=========================================================================
' API
'=========================================================================

'--- error codes
Private Const ERROR_IO_PENDING              As Long = 997
'--- for WaitCommEvent
Private Const EV_RXCHAR                     As Long = &H1
Private Const EV_ERR                        As Long = &H80

Private Declare Function SetCommMask Lib "kernel32" (ByVal hFile As Long, ByVal dwEvtMask As Long) As Long
Private Declare Function WaitCommEvent Lib "kernel32" (ByVal hFile As Long, lpEvtMask As Long, lpOverlapped As Any) As Long
Private Declare Function CreateEvent Lib "kernel32" Alias "CreateEventA" (ByVal lpEventAttributes As Long, ByVal bManualReset As Long, ByVal bInitialState As Long, ByVal lpName As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function WaitForMultipleObjects Lib "kernel32" (ByVal nCount As Long, lpHandles As Long, ByVal bWaitAll As Long, ByVal dwMilliseconds As Long) As Long

Private Type OVERLAPPED
    Internal                    As Long
    InternalHigh                As Long
    Offset                      As Long
    OffsetHigh                  As Long
    hEvent                      As Long
End Type

'=========================================================================
' Constant and member variables
'=========================================================================

Private Const STR_INTERNAL              As String = "Invalid model group (%1)|Cannot load templates from %1. %2|Template not found (%1)|ESC/POS not implemented"
Private Const STR_ERRORS_EPL2           As String = "1|Syntax error|2|Object exceeded label border|3|Bar code data length error|4|Insufficient memory to store data|5|Memory configuration error|6|RS-232 interface error|7|Paper or ribbon empty|8|Duplicate name: form, graphic or soft font|9|Name not found: form, graphic or soft font|10|Not in data entry mode|11|Print head up (open)|12|Pause mode or paused in peel mode|50|Does not fit in area specified|51|Data length too long|80|Undefined|81|Cutter jammed or not installed|82|AutoSense or sensor failure|83|Illegal interrupt occurred|84|Excessive media feeding|93|PDF-417 coded data to large to fit in bar code"
Private Const DEF_AUTODETECTTIMEOUT     As Long = 53
Private Const DEF_CONNECTIONTIMEOUT     As Long = 3000
Private Const DEF_TIMEOUT               As Long = 5000
Private Const DEF_PING_TIMEOUT          As Long = 400
Private Const DEF_AUTODETECTSPEEDS      As String = "115200|9600"
Private Const DEF_SOCKET_PORT           As Long = 9100
Private Const DEF_TEMPLATE_FILE         As String = "LabelTemplates.conf"
Private Const DEF_CODEPAGE              As Long = 1251
Private Const DEF_MIN_DISCOUNT          As Double = -100
Private Const DEF_MAX_DISCOUNT          As Double = 100
Private Const DEF_MAX_RECEIPT_LINES     As Long = 1000
Private Const DEF_MAX_ITEM_LINES        As Long = 2
Private Const STR_GROUP_EPL2            As String = "EPL-2"
Private Const STR_GROUP_ESCPOS          As String = "ESC/POS"
Private Const LNG_ACK                   As Long = &H6
Private Const LNG_NACK                  As Long = &H15

Private m_sDevice                   As String
Private m_oConnector                As IDeviceConnector
Private m_lConnectionTimeout        As Long
Private m_lCommandTimeout           As Long
Private m_lPingTimeout              As Long
Private m_uPrintData                As UcsProtocolPrintData
Private m_uConfig                   As UcsConfigValues

Private Enum UcsInternalErrors
    ucsErrInvalidModelGroup
    ucsErrCannotLoadTemplates
    ucsErrTemplateNotFound
    ucsErrNotImplemented
End Enum

Private Type UcsAutodetect
    Connector           As cSerialPortConnector
    WaitOver            As OVERLAPPED
    EventMask           As Long
End Type

Private Type UcsConfigValues
    CodePage            As Long
    ModelGroup          As String
    Templates           As Object
    LocalizedText(0 To [_ucsFscLciMax] - 1) As Variant
End Type

'=========================================================================
' Error handling
'=========================================================================

'Private Sub PrintError(sFunction As String)
'    Debug.Print "Critical error: " & Err.Description & " [" & MODULE_NAME & "." & sFunction & "]"
'    Logger.Log vbLogEventTypeError, MODULE_NAME, sFunction & "(" & Erl & ")", Err.Description
'End Sub

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
    IsConnected = Not m_oConnector Is Nothing
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function AutodetectDevices(vPorts As Variant) As Variant
    Const FUNC_NAME     As String = "AutodetectDevices"
    Dim vRetVal         As Variant
    Dim vSpeed          As Variant
    Dim aPorts()        As UcsAutodetect
    Dim lIdx            As Long
    Dim baBuffer()      As Byte
    Dim sModel          As String
    Dim sCmd            As String
    Dim sReply          As String
    Dim lResult         As Long
    Dim hEvents()       As Long
    Dim lEventsCount    As Long
    Dim lTimeout        As Long
    
    On Error GoTo EH
    vRetVal = vPorts
    If Not IsArray(vPorts) Then
        GoTo QH
    End If
    If UBound(vPorts) < 0 Then
        GoTo QH
    End If
    lTimeout = GetConfigNumber(STR_PROTOCOL_LABEL, "AutodetectTimeout", DEF_AUTODETECTTIMEOUT)
    For Each vSpeed In Split(DEF_AUTODETECTSPEEDS, "|")
        ReDim aPorts(0 To UBound(vPorts)) As UcsAutodetect
        ReDim hEvents(0 To UBound(vPorts)) As Long
        For lIdx = 0 To UBound(aPorts)
            If Not IsArray(vRetVal(lIdx)) Then
                With aPorts(lIdx)
                    Set .Connector = InitDeviceConnector(vRetVal(lIdx) & "," & vSpeed, lTimeout, LocalizedText(ucsFscLciConnectorErrors))
                    If Not .Connector Is Nothing Then
                        sCmd = vbLf & "UF" & vbLf '--- Form Information Inquiry
                        .Connector.WriteData ToMultiByte(sCmd, m_uConfig.CodePage), 1
                        .WaitOver.hEvent = CreateEvent(0, 1, 0, 0)
                        Call SetCommMask(.Connector.hComm, EV_RXCHAR Or EV_ERR)
                        lResult = WaitCommEvent(.Connector.hComm, .EventMask, .WaitOver)
                        If lResult = 0 And Err.LastDllError = ERROR_IO_PENDING Then
                            hEvents(lEventsCount) = .WaitOver.hEvent
                            lEventsCount = lEventsCount + 1
                        End If
                    End If
                End With
            End If
        Next
        If lEventsCount > 0 Then
            lResult = WaitForMultipleObjects(lEventsCount, hEvents(0), 1, lTimeout)
            lEventsCount = 0
        End If
        For lIdx = 0 To UBound(aPorts)
            With aPorts(lIdx)
                If Not .Connector Is Nothing Then
                    If .Connector.ReadData(baBuffer, lTimeout) And UBound(baBuffer) >= 0 Then
                        sReply = FromMultiByte(baBuffer, m_uConfig.CodePage)
                        If sReply Like "*[0-9][0-9][0-9]" & vbCrLf & "*" Or sReply Like "*[0-9][0-9][0-9]" & vbLf & "*" Then
                            DebugLog FUNC_NAME, "Auto-detecting " & .Connector.Device & " reply 0x" & ToHexDump(StrConv(baBuffer, vbUnicode))
                            sModel = At(pvRawGetDeviceInfo(.Connector), 0)
                            vRetVal(lIdx) = Array(At(vPorts, lIdx), vSpeed, STR_PROTOCOL_LABEL, sModel, vbNullString, vbNullString, vbNullString, STR_GROUP_EPL2)
                        End If
                    End If
                    Set .Connector = Nothing
                    If .WaitOver.hEvent <> 0 Then
                        Call CloseHandle(.WaitOver.hEvent)
                    End If
                End If
            End With
        Next
    Next
QH:
    AutodetectDevices = vRetVal
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function Init(sDeviceString As String) As Boolean
    Const FUNC_NAME     As String = "Init"
    Dim oOptions        As Object
    Dim sFile           As String
    Dim sError          As String
    
    On Error GoTo EH
    DebugLog FUNC_NAME, "sDeviceString=" & sDeviceString
    Set oOptions = ParseDeviceString(sDeviceString)
    m_lConnectionTimeout = C_Lng(JsonValue(oOptions, "ConnectionTimeout"))
    If m_lConnectionTimeout <= 0 Then
        m_lConnectionTimeout = DEF_CONNECTIONTIMEOUT
    End If
    m_lCommandTimeout = C_Lng(JsonValue(oOptions, "Timeout"))
    If m_lCommandTimeout <= 0 Then
        m_lCommandTimeout = DEF_TIMEOUT
    End If
    m_lPingTimeout = C_Lng(JsonValue(oOptions, "PingTimeout"))
    If m_lPingTimeout <= 0 Then
        m_lPingTimeout = DEF_PING_TIMEOUT
    End If
    m_sDevice = ToConnectorDevice(oOptions, DEF_SOCKET_PORT, Me)
    If Not Connect() Then
        GoTo QH
    End If
    sDeviceString = ToDeviceString(oOptions)
    With m_uConfig
        .CodePage = Znl(C_Lng(JsonValue(oOptions, "CodePage")), DEF_CODEPAGE)
        .ModelGroup = UCase$(Zn(C_Str(JsonValue(oOptions, "ModelGroup")), STR_GROUP_EPL2))
        If .ModelGroup <> STR_GROUP_EPL2 And .ModelGroup <> STR_GROUP_ESCPOS Then
            On Error GoTo 0
            Err.Raise vbObjectError, , Printf(pvInternal(ucsErrInvalidModelGroup), Zn(.ModelGroup, STR_NONE))
        End If
        sFile = Zn(C_Str(JsonValue(oOptions, "TemplateFile")), DEF_TEMPLATE_FILE)
        Set .Templates = JsonParseObject(pvReadTextFile(sFile), Error:=sError)
        If .Templates Is Nothing Then
            On Error GoTo 0
            Err.Raise vbObjectError, , Printf(pvInternal(ucsErrCannotLoadTemplates), sFile, sError)
        End If
    End With
    With m_uPrintData.Config
        .RowChars = C_Lng(JsonValue(oOptions, "RowChars"))
        .CommentChars = C_Lng(JsonValue(oOptions, "CommentChars"))
        .ItemChars = C_Lng(JsonValue(oOptions, "ItemChars"))
        .AbsoluteDiscount = True
        .NegativePrices = True
        .MinDiscount = Zndbl(C_Dbl(JsonValue(oOptions, "MinDiscount")), DEF_MIN_DISCOUNT)
        .MaxDiscount = Zndbl(C_Dbl(JsonValue(oOptions, "MaxDiscount")), DEF_MAX_DISCOUNT)
        .MaxReceiptLines = Znl(C_Lng(JsonValue(oOptions, "MaxReceiptLines")), DEF_MAX_RECEIPT_LINES)
        .MaxItemLines = Znl(C_Lng(JsonValue(oOptions, "MaxItemLines")), DEF_MAX_ITEM_LINES)
    End With
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
    DebugLog FUNC_NAME, "m_sDevice=" & m_sDevice & ", m_lConnectionTimeout=" & m_lConnectionTimeout & ", m_uConfig.CodePage=" & m_uConfig.CodePage
    Set m_oConnector = InitDeviceConnector(m_sDevice, m_lConnectionTimeout, LocalizedText(ucsFscLciConnectorErrors), m_uPrintData.LastError)
    If m_oConnector Is Nothing Then
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
    DebugLog FUNC_NAME, "TypeOf m_oConnector=" & TypeName(m_oConnector)
    Set m_oConnector = Nothing
    '--- success
    Disconnect = True
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetClock() As Date
    GetClock = GetCurrentNow
End Function

Public Function GetDeviceProtocol() As String
    GetDeviceProtocol = STR_PROTOCOL_LABEL
End Function

Public Function GetDeviceModel() As String
    GetDeviceModel = At(pvRawGetDeviceInfo(m_oConnector), 0)
End Function

Public Function GetDeviceStatus(StatusText As String) As UcsFiscalStatusCodeEnum
    StatusText = vbNullString
    GetDeviceStatus = ucsFscStaReady
QH:
    pvSetLastError vbNullString
End Function

'= private ===============================================================

Private Function pvRawSendCommand( _
            oConnector As IDeviceConnector, _
            ByVal lTimeout As Long, _
            Data As String, _
            LogName As String, _
            Optional Error As String) As String
    Const FUNC_NAME     As String = "pvRawSendCommand"
    Dim dblTimerEx      As Double
    Dim sSend           As String
    Dim baText()        As Byte
    Dim sBuffer         As String
    
    On Error GoTo EH
    dblTimerEx = TimerEx
    sSend = Data
    DebugDataDump Zn(LogName, FUNC_NAME), "[SEND ] ", sSend
    baText = ToMultiByte(sSend, m_uConfig.CodePage)
    If Not oConnector.WriteData(baText, lTimeout) Then
        Error = oConnector.GetLastError()
        DebugLog Zn(LogName, FUNC_NAME), "WriteData failed, Error=" & Error
        GoTo QH
    End If
    '--- read response
    Do
        If TimerEx > dblTimerEx + lTimeout / 1000# Then
            Error = pvConnectorError(ucsErrTimeoutWaitingForResponse)
            GoTo QH
        End If
        If Not oConnector.ReadData(baText, lTimeout) Then
            Error = oConnector.GetLastError()
            DebugLog Zn(LogName, FUNC_NAME), "ReadData failed, Error=" & Error
            GoTo QH
        End If
        sBuffer = FromMultiByte(baText, m_uConfig.CodePage)
        DebugDataDump Zn(LogName, FUNC_NAME), "[RECV ] ", sBuffer
        pvRawSendCommand = pvRawSendCommand & sBuffer
        If InStr(pvRawSendCommand, pvChr(LNG_NACK)) > 0 Or InStr(pvRawSendCommand, pvChr(LNG_ACK)) > 0 Then
            Exit Do
        End If
    Loop
QH:
    RaiseEvent CommandComplete(Data, pvRawSendCommand)
    Exit Function
EH:
    RaiseError FUNC_NAME
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

Private Function pvConnectorError(ByVal lIdx As UcsConnectorErrorsEnum) As String
    Const FUNC_NAME     As String = "pvConnectorError"
    
    On Error GoTo EH
    pvConnectorError = At(m_uConfig.LocalizedText(ucsFscLciConnectorErrors), lIdx)
    If LenB(pvConnectorError) = 0 Then
        pvConnectorError = At(Split(STR_CONNECTOR_ERRORS, "|"), lIdx)
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Function pvPrintReceiptEpl2(uCtx As UcsPpdExecuteContext, uRow() As UcsPpdRowData, ByVal lRowCount As Long) As Boolean
    Const FUNC_NAME     As String = "pvPrintReceiptEpl2"
    Dim vSplit          As Variant
    Dim lCopies         As Long
    Dim oTemplate       As Object
    Dim cOutput         As Collection
    Dim sName           As String
    Dim sText           As String
    Dim lWidth          As Long
    Dim lHeight         As Long
    Dim sHorAlign       As String
    Dim sWrap           As String
    Dim lIdx            As Long
    Dim vElem           As Variant
    Dim vErrors         As Variant
    Dim sResult         As String
    Dim sError          As String
    Dim oCanvas         As Object
    Dim oItem           As Object
    Dim lTop            As Long
    
    On Error GoTo EH
    Do While uCtx.Row < lRowCount
        With uRow(uCtx.Row)
        Select Case .RowType
        Case ucsRowLine
            vSplit = Split(.LineText, "=", Limit:=2)
            sName = At(vSplit, 0)
            sText = Replace(At(vSplit, 1), "^p", vbCrLf)
            If oTemplate Is Nothing Then
                vSplit = Split(sText, ",")
                lCopies = Znl(C_Lng(At(vSplit, 0)), 1)
                Set oTemplate = C_Obj(JsonValue(m_uConfig.Templates, sName))
                If oTemplate Is Nothing Then
                    On Error GoTo EH
                    Err.Raise vbObjectError, , Printf(pvInternal(ucsErrTemplateNotFound), sName)
                End If
                Set cOutput = New Collection
                '--- FR Command - Retrieve Form
                cOutput.Add "FR" & """" & sName & """"
                '--- ? Command - Download Variables
                cOutput.Add "?"
                lTop = 0
                For Each vElem In JsonKeys(oTemplate, "Controls")
                    Set oItem = C_Obj(JsonValue(oTemplate, "Controls/" & vElem))
                    If IsEmpty(JsonValue(oItem, "Top")) Then
                        JsonValue(oItem, "Top") = lTop
                    End If
                    lTop = lTop + C_Lng(JsonValue(oItem, "Height"))
                Next
                For lIdx = 0 To lTop - 1
                    JsonValue(oCanvas, lIdx) = vbNullString
                Next
            ElseIf IsObject(JsonValue(oTemplate, "Controls/" & sName)) Then
                Set oItem = C_Obj(JsonValue(oTemplate, "Controls/" & sName))
                lTop = C_Lng(JsonValue(oItem, "Top"))
                lWidth = C_Lng(JsonValue(oItem, "Width"))
                lHeight = C_Lng(JsonValue(oItem, "Height"))
                sHorAlign = LCase$(C_Str(JsonValue(oItem, "HorAlign")))
                sWrap = LCase$(C_Str(JsonValue(oItem, "Wrap")))
                If lWidth > 0 Then
                    If sWrap = "barcode" Then
                        ReDim vSplit(0 To Clamp(lHeight - 1, 0)) As String
                        For lIdx = 0 To UBound(vSplit)
                            vSplit(lIdx) = Mid$(sText, lIdx * lWidth + 1, lWidth)
                        Next
                        For lIdx = 1 To UBound(vSplit)
                            If LenB(vSplit(lIdx)) = 0 And Len(vSplit(lIdx - 1)) > 1 Then
                                vSplit(lIdx) = Mid$(vSplit(lIdx - 1), 2)
                                vSplit(lIdx - 1) = Left$(vSplit(lIdx - 1), 1)
                            End If
                        Next
                    Else
                        vSplit = WrapMultiline(sText, lWidth)
                    End If
                Else
                    vSplit = Split(sText, vbCrLf)
                End If
                For lIdx = 0 To Clamp(lHeight - 1, 0)
                    If sHorAlign = "center" And lWidth > 0 Then
                        JsonValue(oCanvas, lTop + lIdx) = RTrim$(CenterText(At(vSplit, lIdx), lWidth))
                    Else
                        JsonValue(oCanvas, lTop + lIdx) = At(vSplit, lIdx)
                    End If
                Next
            End If
        Case ucsRowInit
        Case ucsRowPlu
        Case ucsRowDiscount
        Case ucsRowBarcode
        Case ucsRowPayment
            If .PmtType = 0 Then
                For Each vElem In JsonValue(oCanvas, "*")
                    cOutput.Add Zn(C_Str(vElem), " ")
                Next
                '--- P Command - Print
                cOutput.Add "P" & 1 & "," & lCopies
                If m_lPingTimeout > 1 Then
                    '--- US Command - Enable Error Reporting
                    Call pvRawSendCommand(m_oConnector, m_lPingTimeout, "US1" & vbLf, FUNC_NAME, sError)
                    If LenB(sError) <> 0 Then
                        m_uPrintData.LastError = sError
                        GoTo QH
                    End If
                End If
                sResult = pvRawSendCommand(m_oConnector, m_lCommandTimeout, ConcatCollection(cOutput, vbLf) & vbLf, FUNC_NAME, sError)
                If LenB(sError) <> 0 Then
                    m_uPrintData.LastError = sError
                    GoTo QH
                End If
                If InStr(sResult, pvChr(LNG_NACK)) > 0 Then
                    vErrors = SplitOrReindex(STR_ERRORS_EPL2, "|")
                    For Each vElem In Split(sResult, pvChr(LNG_NACK))
                        sText = At(vErrors, C_Lng(vElem))
                        If LenB(sText) <> 0 Then
                            m_uPrintData.LastError = sText
                            GoTo QH
                        End If
                    Next
                End If
            End If
        End Select
        End With
        uCtx.Row = uCtx.Row + 1
    Loop
    '--- success
    pvPrintReceiptEpl2 = True
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Function pvPrintReceiptEscPos(uCtx As UcsPpdExecuteContext, uRow() As UcsPpdRowData, ByVal lRowCount As Long) As Boolean
    #If uCtx And uRow And lRowCount Then '--- touch args
    #End If
    Err.Raise vbObjectError, , pvInternal(ucsErrNotImplemented)
End Function

Private Function pvRawGetDeviceInfo(oConnector As IDeviceConnector) As Variant
    Const FUNC_NAME     As String = "pvRawGetDeviceInfo"
    Const DEF_MODEL     As String = "LP-50"
    Dim sCmd            As String
    Dim baBuffer()      As Byte
    Dim sReply          As String
    
    On Error GoTo EH
    sCmd = "UQ" & vbLf '--- Configuration Inquiry
    If oConnector.WriteData(ToMultiByte(sCmd, m_uConfig.CodePage), 100) Then
        If oConnector.ReadData(baBuffer, 100) Then
            sReply = FromMultiByte(baBuffer, m_uConfig.CodePage)
            If sReply = pvChr(LNG_NACK) Then
                pvRawGetDeviceInfo = Array(DEF_MODEL)
            ElseIf InStr(sReply, vbCrLf) Then
                pvRawGetDeviceInfo = Split(sReply, vbCrLf)
            Else
                pvRawGetDeviceInfo = Split(sReply, vbLf)
            End If
        End If
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Private Sub pvSetLastError(sError As String)
    m_uPrintData.LastError = sError
End Sub

Private Function pvReadTextFile(ByVal sFile As String) As String
    If Left$(sFile, 1) <> "\" And Mid$(sFile, 2, 1) <> ":" Then
        sFile = PathCombine(App.Path, sFile)
    End If
    sFile = LocateFile(sFile)
    If LenB(sFile) <> 0 Then
        pvReadTextFile = ReadTextFile(sFile)
    End If
End Function

Private Function pvChr(ByVal lChar As Long) As String
    pvChr = FromMultiByteChar(lChar, m_uConfig.CodePage)
End Function

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Initialize()
    m_lConnectionTimeout = DEF_CONNECTIONTIMEOUT
    m_lCommandTimeout = DEF_TIMEOUT
    LocalizedText(ucsFscLciInternalErrors) = STR_INTERNAL
    m_uConfig.CodePage = DEF_CODEPAGE
End Sub

'=========================================================================
' IDeviceProtocol interface
'=========================================================================

Private Function IDeviceProtocol_AddBarcode(ByVal BarcodeType As UcsFiscalBarcodeTypeEnum, Text As String, Optional ByVal Height As Long) As Boolean
    IDeviceProtocol_AddBarcode = PpdAddBarcode(m_uPrintData, BarcodeType, Text, Height)
End Function

Private Function IDeviceProtocol_AddDiscount(ByVal DiscType As UcsFiscalDiscountTypeEnum, ByVal Value As Double) As Boolean
    IDeviceProtocol_AddDiscount = PpdAddDiscount(m_uPrintData, DiscType, Value)
End Function

Private Function IDeviceProtocol_AddLine(Line As String, Optional Command As String, Optional ByVal WordWrap As Boolean = True) As Boolean
    IDeviceProtocol_AddLine = PpdAddLine(m_uPrintData, Line, WordWrap)
End Function

Private Function IDeviceProtocol_AddPayment(ByVal PmtType As UcsFiscalPaymentTypeEnum, Optional PmtName As String, Optional ByVal Amount As Double, Optional ByVal Rate As Double) As Boolean
    IDeviceProtocol_AddPayment = PpdAddPayment(m_uPrintData, PmtType, PmtName, Amount)
End Function

Private Function IDeviceProtocol_AddPLU(Name As String, ByVal Price As Double, Optional ByVal Quantity As Double = 1, Optional ByVal TaxGroup As Long = 2, Optional UnitOfMeasure As String, Optional ByVal DepartmentNo As Long) As Boolean
    IDeviceProtocol_AddPLU = PpdAddPLU(m_uPrintData, Name, Price, Quantity, TaxGroup, UnitOfMeasure, DepartmentNo)
End Function

Private Function IDeviceProtocol_AutodetectDevices(Ports As Variant) As Variant
    IDeviceProtocol_AutodetectDevices = AutodetectDevices(Ports)
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
    If PpdEndReceipt(m_uPrintData, ResumeToken) Then
        Select Case m_uConfig.ModelGroup
        Case STR_GROUP_EPL2
            IDeviceProtocol_PrintReceipt = pvPrintReceiptEpl2(m_uPrintData.ExecCtx, m_uPrintData.Row, m_uPrintData.RowCount)
        Case STR_GROUP_ESCPOS
            IDeviceProtocol_PrintReceipt = pvPrintReceiptEscPos(m_uPrintData.ExecCtx, m_uPrintData.Row, m_uPrintData.RowCount)
        End Select
        With m_uPrintData.ExecCtx
            ReceiptNo = .ReceiptNo
            ReceiptDate = .ReceiptDate
            ReceiptAmount = .ReceiptAmount
            InvoiceNo = .InvoiceNo
        End With
    End If
End Function

Private Function IDeviceProtocol_GetCharsPerLine() As Long
    IDeviceProtocol_GetCharsPerLine = m_uPrintData.Config.RowChars
End Function

Private Function IDeviceProtocol_GetCommentTextMaxLength() As Long
    IDeviceProtocol_GetCommentTextMaxLength = m_uPrintData.Config.CommentChars
End Function

Private Function IDeviceProtocol_GetItemNameMaxLength() As Long
    IDeviceProtocol_GetItemNameMaxLength = m_uPrintData.Config.ItemChars
End Function

Private Function IDeviceProtocol_GetClock() As Date
    IDeviceProtocol_GetClock = GetClock
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
    IDeviceProtocol_GetDeviceProtocol = STR_PROTOCOL_LABEL
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
    IDeviceProtocol_GetDeviceSerialNo = "LABEL_" & At(Split(m_sDevice, ","), 0)
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
    IDeviceProtocol_StartReceipt = PpdStartReceipt(m_uPrintData, ReceiptType, OperatorCode, OperatorName, OperatorPassword, TableNo, UniqueSaleNo, DisablePrinting, InvDocNo, InvCgTaxNo, InvCgTaxNoType, InvCgVatNo, InvCgName, InvCgCity, InvCgAddress, InvCgPrsReceive, RevType, RevReceiptNo, RevReceiptDate, RevFiscalMemoryNo, RevInvoiceNo, RevReason, OwnData)
End Function

Private Function IDeviceProtocol_GetDepartmentName(DepartmentCode As String, Optional TaxGroup As Long) As String
    
End Function

Private Function IDeviceProtocol_GetTaxRates() As Variant

End Function
