VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cServiceController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================================================
'
' UcsFPHub (c) 2019-2022 by Unicontsoft
'
' Unicontsoft Fiscal Printers Hub
'
' This project is licensed under the terms of the MIT license
' See the LICENSE file in the project root for more information
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "cServiceController"

'=========================================================================
' Public events
'=========================================================================

Event BeforeCommitReceipt(oCtx As Object, Cancel As Boolean)

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const LANG_INVALID_JSON      As Long = 48 ' Невалидна JSON заявка: %1
Private Const LANG_INVALID_XML       As Long = 49 ' Невалидна XML заявка
Private Const LANG_PRINTER_NOT_FOUND As Long = 54 ' Принтер %1 не е намерен
Private Const LANG_MISSING_PAYLOAD   As Long = 50 ' Празна заявка
Private Const LANG_INVALID_URL       As Long = 51 ' Неизвестен адрес
Private Const LANG_MISSING_PARAMETER As Long = 52 ' Липсва параметър %1

Private m_lLastCode                 As Long
Private m_sLastError                As String
Private m_oPrinters                 As Object
Private m_oRequestsCache            As Object
Private WithEvents m_oFP            As cFiscalPrinter
Attribute m_oFP.VB_VarHelpID = -1

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunction As String)
    m_sLastError = Err.Description
    #If USE_DEBUG_LOG <> 0 Then
        DebugLog MODULE_NAME, sFunction & "(" & Erl & ")", Err.Description & " &H" & Hex$(Err.Number), vbLogEventTypeError
    #Else
        Debug.Print "Critical error: " & Err.Description & " [" & MODULE_NAME & "." & sFunction & "]"
    #End If
End Sub

'=========================================================================
' Properties
'=========================================================================

Property Get LastCode() As Long
    LastCode = m_lLastCode
End Property

Property Get LastError() As String
    LastError = m_sLastError
End Property

Property Get Printers() As Object
    Set Printers = m_oPrinters
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function Init(oPrinters As Object, oRequestsCache As Object) As Boolean
    Const FUNC_NAME     As String = "Init"
    
    On Error GoTo EH
    Set m_oPrinters = oPrinters
    Set m_oRequestsCache = oRequestsCache
    Set m_oFP = New cFiscalPrinter
    '--- success
    Init = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Public Function ServiceRequest(sPath As String, sQueryString As String, sRequest As String, sResponse As String) As Boolean
    Const FUNC_NAME     As String = "ServiceRequest"
    Dim vMatches        As Variant
    Dim oInfo           As Object
    Dim vJson           As Variant
    Dim oRequest        As Object
    Dim oParams         As Object
    Dim sRequestId      As String
    Dim oEntry          As Object
    
    On Error GoTo EH
    m_lLastCode = 0
    m_sLastError = vbNullString
    Set oParams = ParseQueryString(sQueryString)
    If Not m_oRequestsCache Is Nothing Then
        sRequestId = JsonValue(oParams, "request_id")
    End If
    If LenB(sRequestId) <> 0 Then
        Set oEntry = C_Obj(JsonValue(m_oRequestsCache, sRequestId))
        If Not oEntry Is Nothing Then
            If JsonValue(oEntry, "Path") = sPath And JsonValue(oEntry, "Request") = sRequest Then
                sResponse = JsonValue(oEntry, "Response")
                '--- success
                ServiceRequest = True
                Exit Function
            End If
        End If
    End If
    If preg_match("/^/autoupdate$/i", sPath) Then
        Set oEntry = Nothing
        JsonValue(oEntry, "StartAutoUpdate") = MainForm.DelayStartAutoUpdate()
        sResponse = JsonDump(oEntry, Minimize:=True)
        '--- success
        ServiceRequest = True
        GoTo QH
    End If
    If preg_match("/^/shutdown$/i", sPath) Then
        Set oEntry = Nothing
        JsonValue(oEntry, "ShutDown") = MainForm.DelayShutDown()
        sResponse = JsonDump(oEntry, Minimize:=True)
        '--- success
        ServiceRequest = True
        GoTo QH
    End If
    If preg_match("/^/printers(?:/([^/]+))?(?:/(status|receipt|deposit|report|datetime|totals|drawer))?$/i", sPath, vMatches) = 0 Then
        m_lLastCode = 404
        m_sLastError = T(LANG_INVALID_URL)
        sResponse = pvEncodeErrorText(m_sLastError)
        GoTo QH
    End If
    If LenB(At(vMatches, 0)) = 0 Then
        sResponse = JsonDump(m_oPrinters, Minimize:=True)
    Else
        Set oInfo = pvGetPrinterInfo(At(vMatches, 0))
        If oInfo Is Nothing Then
            m_lLastCode = 404
            m_sLastError = Printf(T(LANG_PRINTER_NOT_FOUND), At(vMatches, 0))
            sResponse = pvEncodeErrorText(m_sLastError)
            GoTo QH
        End If
        If Left$(sRequest, 1) = "<" Then
            AssignVariant vJson, JsonFromXmlDocument(sRequest)
            If IsEmpty(vJson) Then
                m_lLastCode = 400
                m_sLastError = T(LANG_INVALID_XML)
                sResponse = pvEncodeErrorText(m_sLastError)
                GoTo QH
            End If
        ElseIf Not JsonParse(sRequest, vJson, Error:=m_sLastError) Then
            m_lLastCode = 400
            m_sLastError = Printf(T(LANG_INVALID_JSON), m_sLastError)
            sResponse = pvEncodeErrorText(m_sLastError)
            GoTo QH
        End If
        Set oRequest = C_Obj(vJson)
        If oRequest Is Nothing Then
            Select Case LCase$(At(vMatches, 1))
            Case vbNullString
                JsonValue(oRequest, "IncludeAll") = True
            Case "totals"
                JsonValue(oRequest, "IncludeAllTotals") = True
            Case "status", "datetime", "deposit"
                '--- do nothing
            Case Else
                m_lLastCode = 400
                m_sLastError = T(LANG_MISSING_PAYLOAD)
                sResponse = pvEncodeErrorText(m_sLastError)
                GoTo QH
            End Select
        End If
        JsonValue(oRequest, "DeviceString") = JsonValue(oInfo, "DeviceString")
        Select Case LCase$(At(vMatches, 1))
        Case vbNullString
            If UBound(JsonKeys(oRequest)) = 0 Then
                JsonValue(oInfo, "Ok") = True
                sResponse = JsonDump(oInfo, Minimize:=True)
                JsonValue(oInfo, "Ok") = Empty
            ElseIf Not m_oFP.GetDeviceInfo(JsonDump(oRequest, Minimize:=True), sResponse) Then
                GoTo QH
            End If
        Case "status"
            If Not m_oFP.GetDeviceStatus(JsonDump(oRequest, Minimize:=True), sResponse) Then
                GoTo QH
            End If
        Case "receipt"
            If Not m_oFP.PrintReceipt(JsonDump(oRequest, Minimize:=True), sResponse) Then
                GoTo QH
            End If
        Case "deposit"
            If Not m_oFP.PrintServiceDeposit(JsonDump(oRequest, Minimize:=True), sResponse) Then
                GoTo QH
            End If
        Case "report"
            If Not m_oFP.PrintReport(JsonDump(oRequest, Minimize:=True), sResponse) Then
                GoTo QH
            End If
        Case "datetime"
            If Not m_oFP.InitDateTime(JsonDump(oRequest, Minimize:=True), sResponse) Then
                GoTo QH
            End If
        Case "totals"
            If Not m_oFP.GetDailyTotals(JsonDump(oRequest, Minimize:=True), sResponse) Then
                GoTo QH
            End If
        Case "drawer"
            If Not C_Bool(JsonValue(oRequest, "IsOpen")) Then
                m_lLastCode = 400
                m_sLastError = Printf(T(LANG_MISSING_PARAMETER), "IsOpen")
                sResponse = pvEncodeErrorText(m_sLastError)
                GoTo QH
            End If
            If Not m_oFP.OpenDrawer(JsonDump(oRequest, Minimize:=True), sResponse) Then
                GoTo QH
            End If
        End Select
    End If
    '--- success
    ServiceRequest = True
QH:
    If LenB(sResponse) <> 0 Then
        Select Case LCase$(JsonValue(oParams, "format"))
        Case "xml"
            If JsonParse(sResponse, vJson) Then
                sResponse = JsonToXmlDocument(vJson).Xml
            End If
        Case "json"
            If JsonParse(sResponse, vJson) Then
                sResponse = JsonDump(vJson)
            End If
        End Select
    End If
    If ServiceRequest And LenB(sRequestId) <> 0 And InStr(sResponse, """Ok"":true") <> 0 Then
        Set oEntry = Nothing
        JsonValue(oEntry, "Path") = sPath
        JsonValue(oEntry, "QueryString") = sQueryString
        JsonValue(oEntry, "Request") = sRequest
        JsonValue(oEntry, "Response") = sResponse
        JsonValue(oEntry, "DateTime") = GetCurrentNow
        JsonValue(m_oRequestsCache, sRequestId) = oEntry
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Err.Raise Err.Number, MODULE_NAME & "." & FUNC_NAME & vbCrLf & Err.Source, Err.Description
End Function

Private Function pvGetPrinterInfo(sPrinterID As String) As Object
    Dim sKey            As String
    
    sKey = Zn(JsonValue(m_oPrinters, "Aliases/" & sPrinterID), sPrinterID)
    If LenB(sKey) <> 0 Then
        Set pvGetPrinterInfo = C_Obj(JsonValue(m_oPrinters, sKey))
    End If
End Function

Private Function pvEncodeErrorText(sErrorText As String) As String
    Dim oResponse       As Object
    
    JsonValue(oResponse, "Ok") = False
    JsonValue(oResponse, "ErrorText") = sErrorText
    pvEncodeErrorText = JsonDump(oResponse, Minimize:=True)
End Function

'=========================================================================
' Events
'=========================================================================

Private Sub m_oFP_BeforeCommitReceipt(oCtx As Object, Cancel As Boolean)
    RaiseEvent BeforeCommitReceipt(oCtx, Cancel)
End Sub
