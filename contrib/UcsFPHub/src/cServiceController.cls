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
' UcsFPHub (c) 2019 by Unicontsoft
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
' Constants and member variables
'=========================================================================

Private Const ERR_INVALID_JSON      As String = "Невалидна JSON заявка: %1"
Private Const ERR_PRINTER_NOT_FOUND As String = "Принтер %1 не е намерен"
Private Const ERR_MISSING_REQUEST   As String = "Празна JSON заявка"

Private m_sLastError                As String
Private m_oPrinters                 As Object
Private m_oFP                       As cFiscalPrinter

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunction As String)
    m_sLastError = Err.Description
    Debug.Print "Critical error: " & Err.Description & " [" & MODULE_NAME & "." & sFunction & "]"
End Sub

'=========================================================================
' Properties
'=========================================================================

Property Get LastError() As String
    LastError = m_sLastError
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function Init(oPrinters As Object) As Boolean
    Const FUNC_NAME     As String = "Init"
    
    On Error GoTo EH
    Set m_oPrinters = oPrinters
    Set m_oFP = New cFiscalPrinter
    '--- success
    Init = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Public Function ServiceRequest(sPath As String, sRequest As String, sResponse As String) As Boolean
    Const FUNC_NAME     As String = "ServiceRequest"
    Dim vMatches        As Variant
    Dim oInfo           As Object
    Dim vJson           As Variant
    Dim oRequest        As Object
    Dim oResponse       As Object
    Dim sError          As String
    
    On Error GoTo EH
    m_sLastError = vbNullString
    If preg_match("/^/printers(?:/([^/]+))?(?:/(status|receipt|deposit|report|datetime|totals))?$/i", sPath, vMatches) = 0 Then
        GoTo QH
    End If
    If LenB(At(vMatches, 0)) = 0 Then
        sResponse = JsonDump(m_oPrinters, Minimize:=True)
    Else
        Set oInfo = pvGetPrinterInfo(At(vMatches, 0))
        If oInfo Is Nothing Then
            sResponse = pvEncodeErrorText(Printf(ERR_PRINTER_NOT_FOUND, At(vMatches, 0)))
            GoTo QH
        End If
        If Not JsonParse(sRequest, vJson, Error:=m_sLastError) Then
            sResponse = pvEncodeErrorText(Printf(ERR_INVALID_JSON, m_sLastError))
            GoTo QH
        End If
        Set oRequest = C_Obj(vJson)
        If oRequest Is Nothing Then
            Select Case LCase$(At(vMatches, 1))
            Case vbNullString
                JsonItem(oRequest, "IncludeAll") = True
            Case "status", "datetime", "deposit", "totals"
                '--- do nothing
            Case Else
                sResponse = pvEncodeErrorText(ERR_MISSING_REQUEST)
                GoTo QH
            End Select
        End If
        JsonItem(oRequest, "DeviceString") = JsonItem(oInfo, "DeviceString")
        Select Case LCase$(At(vMatches, 1))
        Case vbNullString
            If Not m_oFP.GetDeviceInfo(JsonDump(oRequest, Minimize:=True), sResponse) Then
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
            If Not m_oFP.GetTotalsStats(JsonDump(oRequest, Minimize:=True), sResponse) Then
                GoTo QH
            End If
            If Not JsonParse(sResponse, oResponse, Error:=sError) Then
                Err.Raise vbObjectError, , sError
            End If
            If Not m_oFP.GetTotalsByPayments(JsonDump(oRequest, Minimize:=True), sResponse) Then
                GoTo QH
            End If
            If Not JsonParse(sResponse, vJson, Error:=sError) Then
                Err.Raise vbObjectError, , sError
            End If
            If C_Bool(JsonItem(C_Obj(vJson), "Ok")) Then
                JsonItem(oResponse, "TotalsByPayments") = JsonItem(C_Obj(vJson), "TotalsByPayments")
            End If
            If Not m_oFP.GetTotalsByTaxGroups(JsonDump(oRequest, Minimize:=True), sResponse) Then
                GoTo QH
            End If
            If Not JsonParse(sResponse, vJson, Error:=sError) Then
                Err.Raise vbObjectError, , sError
            End If
            If C_Bool(JsonItem(C_Obj(vJson), "Ok")) Then
                JsonItem(oResponse, "TotalsByTaxGroups") = JsonItem(C_Obj(vJson), "TotalsByTaxGroups")
            End If
            sResponse = JsonDump(oResponse, Minimize:=True)
        End Select
    End If
    '--- success
    ServiceRequest = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
    Err.Raise Err.Number, Err.Source & vbCrLf & MODULE_NAME & "." & FUNC_NAME, Err.Description
End Function

Private Function pvGetPrinterInfo(sPrinterId As String) As Object
    Dim sKey            As String
    
    sKey = Zn(JsonItem(m_oPrinters, "Aliases/" & sPrinterId), sPrinterId)
    If LenB(sKey) <> 0 Then
        Set pvGetPrinterInfo = C_Obj(JsonItem(m_oPrinters, sKey))
    End If
End Function

Private Function pvEncodeErrorText(sErrorText As String) As String
    Dim oResponse       As Object
    
    JsonItem(oResponse, "Ok") = False
    JsonItem(oResponse, "ErrorText") = sErrorText
    pvEncodeErrorText = JsonDump(oResponse, Minimize:=True)
End Function
