VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSocketConnector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================================================
' $Header: /UcsFiscalPrinter/Src/Connectors/cSocketConnector.cls 1     25.04.18 10:39 Wqw $
'
'   Unicontsoft Fiscal Printers Project
'   Copyright (c) 2008-2018 Unicontsoft
'
'   Network socket connector
'
' $Log: /UcsFiscalPrinter/Src/Connectors/cSocketConnector.cls $
' 
' 1     25.04.18 10:39 Wqw
' Initial implementation
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "cSerialPortConnector"
Implements IDeviceConnector

'=========================================================================
' Constants and member variables
'=========================================================================

Private m_oSocket               As cAsyncSocket
Private m_lLastError            As Long

'=========================================================================
' Error handling
'=========================================================================

'Private Sub PrintError(sFunc As String)
'    Debug.Print MODULE_NAME & "." & sFunc & ": " & Err.Description
'    OutputDebugLog MODULE_NAME, sFunc & "(" & Erl & ")", "Run-time error: " & Err.Description
'End Sub

Private Sub RaiseError(sFunc As String)
    Debug.Print MODULE_NAME & "." & sFunc & ": " & Err.Description
    OutputDebugLog MODULE_NAME, sFunc & "(" & Erl & ")", "Run-time error: " & Err.Description
    Err.Raise Err.Number, MODULE_NAME & "." & sFunc & "(" & Erl & ")" & vbCrLf & Err.Source, Err.Description
End Sub

Private Sub DebugLog(sFunc As String, sText As String)
    OutputDebugLog MODULE_NAME, sFunc, sText
End Sub

'=========================================================================
' Methods
'=========================================================================

Public Function Init(Device As String, Optional ByVal Timeout As Long) As Boolean
    Const FUNC_NAME     As String = "Init"
    Dim vSplit          As Variant
    
    On Error GoTo EH
    vSplit = Split(Device, ":")
    Set m_oSocket = New cAsyncSocket
    If Not m_oSocket.Create(SocketType:=ucsSckStream) Then
        DebugLog FUNC_NAME, "m_oSocket.Create failed, m_oSocket.LastError=" & m_oSocket.LastError
        m_lLastError = m_oSocket.LastError
        GoTo QH
    End If
    If Not m_oSocket.Connect(At(vSplit, 0), At(vSplit, 1)) Then
        DebugLog FUNC_NAME, "m_oSocket.Connect failed, Device=" & Device & ", m_oSocket.LastError=" & m_oSocket.LastError
        m_lLastError = m_oSocket.LastError
        GoTo QH
    End If
    If Not m_oSocket.SyncWaitForEvent(Timeout, ucsSfdConnect) Then
        DebugLog FUNC_NAME, "m_oSocket.SyncWaitForEvent failed, Timeout=" & Timeout & ", m_oSocket.LastError=" & m_oSocket.LastError
        m_lLastError = m_oSocket.LastError
        GoTo QH
    End If
    '--- success
    Init = True
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function WriteData(Data() As Byte, Optional ByVal Timeout As Long) As Boolean
    Const FUNC_NAME     As String = "WriteData"
    
    On Error GoTo EH
    If Not m_oSocket.SyncSendArray(Data, Timeout) Then
        DebugLog FUNC_NAME, "m_oSocket.SyncSendArray failed, m_oSocket.LastError=" & m_oSocket.LastError
        m_lLastError = m_oSocket.LastError
        GoTo QH
    End If
    '--- success
    WriteData = True
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function ReadData(Result() As Byte, Optional ByVal Timeout As Long) As Boolean
    Const FUNC_NAME     As String = "ReadData"
    
    On Error GoTo EH
    If Not m_oSocket.SyncReceiveArray(Result, Timeout) Then
        DebugLog FUNC_NAME, "m_oSocket.SyncReceiveArray failed, m_oSocket.LastError=" & m_oSocket.LastError
        m_lLastError = m_oSocket.LastError
        GoTo QH
    End If
    '--- success
    ReadData = True
QH:
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetLastError() As String
    If Not m_oSocket Is Nothing Then
        GetLastError = m_oSocket.GetErrorDescription(m_lLastError)
    End If
End Function

'=========================================================================
' IDeviceConnector interface
'=========================================================================

Private Function IDeviceConnector_ReadData(Result() As Byte, Optional ByVal Timeout As Long) As Boolean
    IDeviceConnector_ReadData = ReadData(Result, Timeout)
End Function

Private Function IDeviceConnector_WriteData(Data() As Byte, Optional ByVal Timeout As Long) As Boolean
    IDeviceConnector_WriteData = WriteData(Data, Timeout)
End Function

Private Function IDeviceConnector_GetLastError(Optional ErrCategory As UcsConnectorErrorsEnum) As String
    IDeviceConnector_GetLastError = GetLastError()
End Function

