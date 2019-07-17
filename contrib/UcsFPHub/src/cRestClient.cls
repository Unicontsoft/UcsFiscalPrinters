VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRestClient"
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
Private Const MODULE_NAME As String = "cRestClient"

'=========================================================================
' API
'=========================================================================

Private Declare Function vbaObjSetAddref Lib "msvbvm60" Alias "__vbaObjSetAddref" (oDest As Any, ByVal lSrcPtr As Long) As Long

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const HDR_CONTENT_TYPE          As String = "Content-Type"
Private Const HDR_CONTENT_LENGTH        As String = "Content-Length"

Private m_sID                       As String
Private WithEvents m_oSocket        As cAsyncSocket
Attribute m_oSocket.VB_VarHelpID = -1
Private m_lParentPtr                As Long
Private m_oPrinters                 As Object
Private m_oHeaders                  As Object
Private m_sBuffer                   As String
Private m_sHttpMethod               As String
Private m_sRawUrl                   As String
Private m_sProtocolVersion          As String
Private m_sLastError                As String

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunction As String)
    m_sLastError = Err.Description
    Debug.Print "Critical error: Connection " & m_sID & ": " & Err.Description & " [" & MODULE_NAME & "." & sFunction & "]"
End Sub

'=========================================================================
' Properties
'=========================================================================

Property Get ID() As String
    ID = m_sID
End Property

Property Get Socket() As cAsyncSocket
    Set Socket = m_oSocket
End Property

Property Get HttpMethod() As String
    HttpMethod = m_sHttpMethod
End Property

Property Get RawUrl() As String
    RawUrl = m_sRawUrl
End Property

Property Get ProtocolVersion() As String
    ProtocolVersion = m_sProtocolVersion
End Property

Property Get Headers() As Object
    Set Headers = m_oHeaders
End Property

Property Get Body() As String
    Body = m_sBuffer
End Property

Property Get LastError() As String
    LastError = m_sLastError
End Property

Private Property Get pvParent() As cRestEndpoint
    Call vbaObjSetAddref(pvParent, m_lParentPtr)
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function Init(sKey As String, oSocket As cAsyncSocket, oParent As cRestEndpoint) As Boolean
    Const FUNC_NAME     As String = "Init"
    
    On Error GoTo EH
    m_sID = sKey
    m_lParentPtr = ObjPtr(oParent)
    If Not JsonParse(JsonDump(oParent.Printers, Minimize:=True), m_oPrinters, Error:=m_sLastError) Then
        GoTo QH
    End If
    Set m_oSocket = New cAsyncSocket
    If Not oSocket.Accept(m_oSocket) Then
        m_sLastError = oSocket.GetErrorDescription(oSocket.LastError)
        Set m_oSocket = Nothing
        GoTo QH
    End If
    '--- success
    Init = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Private Sub pvClearState()
    Set m_oHeaders = Nothing
    m_sBuffer = vbNullString
    m_sHttpMethod = vbNullString
    m_sRawUrl = vbNullString
    m_sProtocolVersion = vbNullString
End Sub

Private Function pvHandleRequest() As Boolean
    Const FUNC_NAME     As String = "pvHandleRequest"
    Dim sToLowerUrl     As String
    Dim sBody           As String
    Dim baBuffer()      As Byte
    Dim vMatches        As Variant
    Dim oInfo           As Object
    Dim oRequest        As Object
    Dim oFP             As New cFiscalPrinter
    Dim sResponse       As String
    
    On Error GoTo EH
    sToLowerUrl = LCase$(m_sRawUrl)
    If sToLowerUrl = "/printers" Then
        sBody = JsonDump(m_oPrinters, Minimize:=True)
    ElseIf preg_match("/^/printers/([^/]+)$/i", m_sRawUrl, vMatches) Then
        sBody = JsonDump(pvGetPrinterInfo(At(vMatches, 0)), Minimize:=True)
    ElseIf preg_match("/^/printers/([^/]+)/status$/i", m_sRawUrl, vMatches) Then
        Set oInfo = pvGetPrinterInfo(At(vMatches, 0))
        If Not oInfo Is Nothing Then
            JsonItem(oRequest, "DeviceString") = JsonItem(oInfo, "DeviceString")
            If oFP.GetDeviceStatus(JsonDump(oRequest, Minimize:=True), sResponse) Then
                sBody = sResponse
            End If
        End If
    Else
        If pvParent.ClientReceive(m_sID) Then
            '--- success
            pvHandleRequest = True
            GoTo QH
        End If
    End If
    If LenB(sBody) <> 0 Then
        baBuffer = m_oSocket.ToTextArray(sBody)
        m_oSocket.SendText "HTTP/1.1 200 OK" & vbCrLf & _
            HDR_CONTENT_TYPE & ": application/json; charset=utf-8" & vbCrLf & _
            HDR_CONTENT_LENGTH & ": " & UBound(baBuffer) + 1 & vbCrLf & vbCrLf
        m_oSocket.SendArray baBuffer
    Else
        baBuffer = m_oSocket.ToTextArray("404 Not Found")
        m_oSocket.SendText "HTTP/1.1 404 Not Found" & vbCrLf & _
            HDR_CONTENT_LENGTH & ": " & UBound(baBuffer) + 1 & vbCrLf & vbCrLf
        m_oSocket.SendArray baBuffer
    End If
    '--- success
    pvHandleRequest = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
    baBuffer = m_oSocket.ToTextArray(m_sLastError)
    m_oSocket.SendText "HTTP/1.1 500 Internal Server Error" & vbCrLf & _
        HDR_CONTENT_LENGTH & ": " & UBound(baBuffer) + 1 & vbCrLf & vbCrLf
    m_oSocket.SendArray baBuffer
    Resume QH
End Function

Private Function pvGetPrinterInfo(sPrinterId As String) As Object
    Dim sKey            As String
    
    sKey = Zn(JsonItem(m_oPrinters, "Aliases/" & sPrinterId), sPrinterId)
    If LenB(sKey) <> 0 Then
        Set pvGetPrinterInfo = C_Obj(JsonItem(m_oPrinters, sKey))
    End If
End Function

'=========================================================================
' Events
'=========================================================================

Private Sub m_oSocket_OnReceive()
    Const FUNC_NAME     As String = "m_oSocket_OnReceive"
    Dim lHdrPos         As Long
    Dim vElem           As Variant
    Dim lPos            As Long
    Dim lIdx            As Long
    
    On Error GoTo EH
    m_sBuffer = m_sBuffer & m_oSocket.ReceiveText()
    If m_oHeaders Is Nothing Then
        lHdrPos = InStr(m_sBuffer, vbCrLf & vbCrLf)
        If lHdrPos > 0 Then
            For Each vElem In Split(Left$(m_sBuffer, lHdrPos), vbCrLf)
                If lIdx = 0 Then
                    vElem = Split(vElem, " ")
                    m_sHttpMethod = At(vElem, 0)
                    m_sRawUrl = At(vElem, 1)
                    m_sProtocolVersion = At(vElem, 2)
                Else
                    lPos = InStr(vElem, ":")
                    If lPos > 1 Then
                        JsonItem(m_oHeaders, Trim$(Left$(vElem, lPos - 1))) = Trim$(Mid$(vElem, lPos + 1))
                    End If
                End If
                lIdx = lIdx + 1
            Next
        End If
        m_sBuffer = Mid$(m_sBuffer, lHdrPos + 4)
    End If
    If Not m_oHeaders Is Nothing Then
        If Not IsEmpty(JsonItem(m_oHeaders, HDR_CONTENT_LENGTH)) Then
            If Len(m_sBuffer) < C_Lng(JsonItem(m_oHeaders, HDR_CONTENT_LENGTH)) Then
                GoTo QH
            End If
        End If
        pvHandleRequest
        pvClearState
    End If
QH:
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub m_oSocket_OnClose()
    Const FUNC_NAME     As String = "m_oSocket_OnClose"
    
    On Error GoTo EH
    pvParent.ClientClose m_sID
    pvClearState
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub m_oSocket_OnError(ByVal ErrorCode As Long, ByVal EventMask As UcsAsyncSocketEventMaskEnum)
    Const FUNC_NAME     As String = "m_oSocket_OnError"
    
    On Error GoTo EH
    m_sLastError = m_oSocket.GetErrorDescription(ErrorCode)
    pvParent.ClientError m_sID
    pvClearState
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub
