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
Private m_oController               As cServiceController

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
    Set m_oController = New cServiceController
    If Not m_oController.Init(m_oPrinters) Then
        m_sLastError = m_oController.LastError
        Set m_oSocket = Nothing
        Set m_oController = Nothing
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

Private Function pvHandleRequest(sPath As String, sQueryString As String) As Boolean
    Const FUNC_NAME     As String = "pvHandleRequest"
    Dim sResponse       As String
    Dim baBuffer()      As Byte
    
    On Error GoTo EH
    #If sQueryString Then '--- touch args
    #End If
    If m_sHttpMethod <> "GET" And m_sHttpMethod <> "POST" Then
        baBuffer = m_oSocket.ToTextArray("400 Bad request")
        m_oSocket.SendText "HTTP/1.1 400 Bad request" & vbCrLf & _
            HDR_CONTENT_LENGTH & ": " & UBound(baBuffer) + 1 & vbCrLf & vbCrLf
        m_oSocket.SendArray baBuffer
        GoTo QH
    End If
    If Not m_oController.ServiceRequest(sPath, m_sBuffer, sResponse) Then
        If LenB(sResponse) = 0 Then
            If LenB(m_oController.LastError) <> 0 Then
                baBuffer = m_oSocket.ToTextArray(m_oController.LastError)
                GoTo InternalServerError
            End If
            baBuffer = m_oSocket.ToTextArray("404 Not Found")
            m_oSocket.SendText "HTTP/1.1 404 Not Found" & vbCrLf & _
                HDR_CONTENT_LENGTH & ": " & UBound(baBuffer) + 1 & vbCrLf & vbCrLf
            m_oSocket.SendArray baBuffer
            GoTo QH
        End If
    End If
    baBuffer = m_oSocket.ToTextArray(sResponse)
    m_oSocket.SendText "HTTP/1.1 200 OK" & vbCrLf & _
        HDR_CONTENT_TYPE & ": application/json; charset=utf-8" & vbCrLf & _
        HDR_CONTENT_LENGTH & ": " & UBound(baBuffer) + 1 & vbCrLf & vbCrLf
    m_oSocket.SendArray baBuffer
    '--- success
    pvHandleRequest = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
    baBuffer = m_oSocket.ToTextArray(m_sLastError)
InternalServerError:
    m_oSocket.SendText "HTTP/1.1 500 Internal Server Error" & vbCrLf & _
        HDR_CONTENT_LENGTH & ": " & UBound(baBuffer) + 1 & vbCrLf & vbCrLf
    m_oSocket.SendArray baBuffer
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
        vElem = Split2(m_sRawUrl, "?")
        pvHandleRequest At(vElem, 0), At(vElem, 1)
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
