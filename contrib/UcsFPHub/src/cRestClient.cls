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
' UcsFPHub (c) 2019-2020 by Unicontsoft
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

'--- for MultiByteToWideChar
Private Const CP_UTF8                       As Long = 65001

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function vbaObjSetAddref Lib "msvbvm60" Alias "__vbaObjSetAddref" (oDest As Any, ByVal lSrcPtr As Long) As Long
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const HDR_CONTENT_TYPE          As String = "Content-Type"
Private Const HDR_CONTENT_LENGTH        As String = "Content-Length"
Private Const HDR_EXPECT                As String = "Expect"
Private Const STR_STATUS_CODE_BAD_REQUEST As String = "400 Bad request"
Private Const STR_STATUS_CODE_INTERNAL_SERVER_ERROR As String = "500 Internal Server Error"
Private Const STR_STATUS_CODE_NOT_FOUND As String = "404 Not Found"
Private Const STR_STATUS_CODE_OK        As String = "200 OK"
Private Const STR_CONTENT_TYPE_TEXT_PLAIN As String = "text/plain; charset=utf-8"
Private Const STR_CONTENT_TYPE_TEXT_XML As String = "text/xml; charset=utf-8"
Private Const STR_CONTENT_TYPE_APPLICATION_JSON As String = "application/json"
Private Const STR_CONN_ELAPSED          As String = "%1 изпълни заявка %2 за %3 сек."

Private m_sLastError                As String
Private m_sID                       As String
Private WithEvents m_oSocket        As cAsyncSocket
Attribute m_oSocket.VB_VarHelpID = -1
Private m_lParentPtr                As Long
Private m_oController               As cServiceController
Private m_baBuffer()                As Byte
Private m_lHdrSize                  As Long
'--- http request members
Private m_oHeaders                  As Object
Private m_sBody                     As String
Private m_sHttpMethod               As String
Private m_sRawUrl                   As String
Private m_sProtocolVersion          As String
Private m_dblStartTimerEx           As Double

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunction As String)
    m_sLastError = Err.Description
    #If USE_DEBUG_LOG <> 0 Then
        DebugLog MODULE_NAME, sFunction & "(" & m_sID & ")(" & Erl & ")", Err.Description & " &H" & Hex$(Err.Number), vbLogEventTypeError
    #Else
        Debug.Print "Critical error: " & Err.Description & " [" & MODULE_NAME & "." & sFunction & "]"
    #End If
End Sub

'=========================================================================
' Properties
'=========================================================================

Property Get LastError() As String
    LastError = m_sLastError
End Property

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
    Body = m_sBody
End Property

Private Property Get pvParent() As cRestEndpoint
    Call vbaObjSetAddref(pvParent, m_lParentPtr)
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function Init(sKey As String, oSocket As cAsyncSocket, oParent As cRestEndpoint) As Boolean
    Const FUNC_NAME     As String = "Init"
    Dim oPrinters       As Object

    On Error GoTo EH
    m_sID = sKey
    m_lParentPtr = ObjPtr(oParent)
    Set m_oController = oParent.ServiceController
    pvClearState
    '--- init socket
    Set m_oSocket = New cAsyncSocket
    If Not oSocket.Accept(m_oSocket) Then
        m_sLastError = oSocket.GetErrorDescription(oSocket.LastError)
        Set m_oSocket = Nothing
        Set m_oController = Nothing
        GoTo QH
    End If
    '--- deep-copy fiscal printers config
    If Not JsonParse(JsonDump(oParent.Printers, Minimize:=True), oPrinters, Error:=m_sLastError) Then
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
    m_baBuffer = vbNullString
    Set m_oHeaders = Nothing
    m_sBody = vbNullString
    m_sHttpMethod = vbNullString
    m_sRawUrl = vbNullString
    m_sProtocolVersion = vbNullString
    m_dblStartTimerEx = 0
End Sub

Private Function pvHandleRequest(sPath As String, sQueryString As String) As Boolean
    Const FUNC_NAME     As String = "pvHandleRequest"
    Dim sResponse       As String
    Dim sStatusCode     As String
    Dim sContentType    As String
    Dim baBody()        As Byte
    
    On Error GoTo EH
    If IsLogDebugEnabled Then
        DebugLog MODULE_NAME, FUNC_NAME, "m_sHttpMethod=" & m_sHttpMethod & ", sPath=" & sPath & ", sQueryString=" & sQueryString & ", m_sBody=" & Replace(m_sBody, vbCrLf, "^p"), vbLogEventTypeDebug
    End If
    If m_sHttpMethod <> "GET" And m_sHttpMethod <> "POST" Then
        sStatusCode = STR_STATUS_CODE_BAD_REQUEST
        sContentType = STR_CONTENT_TYPE_TEXT_PLAIN
        baBody = m_oSocket.ToTextArray(sStatusCode)
        GoTo QH
    End If
    If Not m_oController.ServiceRequest(sPath, sQueryString, m_sBody, sResponse) Then
        Select Case m_oController.LastCode
        Case 400
            sStatusCode = STR_STATUS_CODE_BAD_REQUEST
        Case 404
            sStatusCode = STR_STATUS_CODE_NOT_FOUND
        Case Else
            sStatusCode = STR_STATUS_CODE_INTERNAL_SERVER_ERROR
        End Select
        If LenB(sResponse) <> 0 Then
            If Left$(sResponse, 1) = "<" Then
                sContentType = STR_CONTENT_TYPE_TEXT_XML
            Else
                sContentType = STR_CONTENT_TYPE_APPLICATION_JSON
            End If
            baBody = m_oSocket.ToTextArray(sResponse)
        Else
            sContentType = STR_CONTENT_TYPE_TEXT_PLAIN
            baBody = m_oSocket.ToTextArray(Zn(m_oController.LastError, sStatusCode))
        End If
        GoTo QH
    End If
    sStatusCode = STR_STATUS_CODE_OK
    If Left$(sResponse, 1) = "<" Then
        sContentType = STR_CONTENT_TYPE_TEXT_XML
    Else
        sContentType = STR_CONTENT_TYPE_APPLICATION_JSON
    End If
    baBody = m_oSocket.ToTextArray(sResponse)
    '--- success
    pvHandleRequest = True
QH:
    If LenB(sStatusCode) <> 0 And Not m_oSocket Is Nothing Then
        m_oSocket.SendText "HTTP/1.1 " & sStatusCode & vbCrLf & _
            HDR_CONTENT_TYPE & ": " & sContentType & vbCrLf & _
            HDR_CONTENT_LENGTH & ": " & UBound(baBody) + 1 & vbCrLf & vbCrLf
        m_oSocket.SendArray baBody
    End If
    If IsLogDebugEnabled Then
        DebugLog MODULE_NAME, FUNC_NAME, "sStatusCode=" & sStatusCode & ", sResponse=" & Replace(sResponse, vbCrLf, "^p"), vbLogEventTypeDebug
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    sStatusCode = STR_STATUS_CODE_INTERNAL_SERVER_ERROR
    sContentType = STR_CONTENT_TYPE_TEXT_PLAIN
    baBody = m_oSocket.ToTextArray(m_sLastError)
    GoTo QH
End Function

'=========================================================================
' Events
'=========================================================================

Private Sub m_oSocket_OnReceive()
    Const FUNC_NAME     As String = "m_oSocket_OnReceive"
    Dim lSize           As Long
    Dim vElem           As Variant
    Dim lPos            As Long
    Dim lIdx            As Long
    Dim baBuffer()      As Byte
    
    On Error GoTo EH
    If Not m_oSocket.ReceiveArray(baBuffer) Then
        GoTo QH
    End If
    If UBound(baBuffer) < 0 Then
        GoTo QH
    End If
    lIdx = UBound(m_baBuffer) + 1
    If lIdx = 0 Then
        m_dblStartTimerEx = TimerEx
        m_baBuffer = baBuffer
    Else
        ReDim Preserve m_baBuffer(0 To lIdx + UBound(baBuffer)) As Byte
        Call CopyMemory(m_baBuffer(lIdx), baBuffer(0), UBound(baBuffer) + 1)
    End If
    If m_oHeaders Is Nothing Then
        m_lHdrSize = InStrB(1, m_baBuffer, StrConv(vbCrLf & vbCrLf, vbFromUnicode))
        If m_lHdrSize > 0 Then
            For Each vElem In Split(Left$(StrConv(m_baBuffer, vbUnicode), m_lHdrSize - 1), vbCrLf)
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
            m_lHdrSize = m_lHdrSize + 3
            If LCase$(JsonItem(m_oHeaders, HDR_EXPECT)) = "100-continue" Then
                m_oSocket.SendText "HTTP/1.1 100 Continue" & vbCrLf & vbCrLf
            End If
        End If
    End If
    If Not m_oHeaders Is Nothing Then
        lSize = UBound(m_baBuffer) + 1 - m_lHdrSize
        If Not IsEmpty(JsonItem(m_oHeaders, HDR_CONTENT_LENGTH)) Then
            If lSize < C_Lng(JsonItem(m_oHeaders, HDR_CONTENT_LENGTH)) Then
                GoTo QH
            End If
        End If
        If lSize > 0 Then
            m_sBody = String$(2 * lSize, 0)
            lSize = MultiByteToWideChar(CP_UTF8, 0, m_baBuffer(m_lHdrSize), lSize, StrPtr(m_sBody), Len(m_sBody) + 1)
            If lSize <> Len(m_sBody) Then
                m_sBody = Left$(m_sBody, lSize)
            End If
        End If
        vElem = Split2(m_sRawUrl, "?")
        pvHandleRequest At(vElem, 0), At(vElem, 1)
        DebugLog MODULE_NAME, FUNC_NAME, Printf(STR_CONN_ELAPSED, m_sID, m_sRawUrl, Format(TimerEx - m_dblStartTimerEx, FORMAT_BASE_3))
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
