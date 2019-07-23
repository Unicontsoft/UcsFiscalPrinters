VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMssqlEndpoint"
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
Private Const MODULE_NAME As String = "cMssqlEndpoint"

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_CONN_OPEN             As String = "Свързан с %1"
Private Const STR_QUEUE_SETUP           As String = "Слуша на опашка %2 за %1 услуги"
Private Const STR_QUEUE_ELAPSED         As String = "%1 на принтер %2 изпълни заявка %3 за %4 сек."
Private Const STR_QUEUE_ERROR           As String = "Грешка в %1: %2"
Private Const STR_CROSS_PRINTER_REQUEST As String = "Заявка %1 не е към устройството на услуга %2"
'--- errors
Private Const ERR_INVALID_SERVICE       As String = "Невалидна Service Broker услуга: %1"
Private Const ERR_INVALID_URL           As String = "Невалиден адрес: %1"
'--- queue
Private Const DEF_QUEUE_TIMEOUT         As Long = 5000
Private Const STR_QUEUE_LOOP            As String = "DECLARE     @QueueName SYSNAME = '%1', @Result INT, @Handle UNIQUEIDENTIFIER, @Request NVARCHAR(MAX), @MsgType SYSNAME, @SvcName SYSNAME" & vbCrLf & _
                                                    "EXEC        @Result = dbo.usp_sys_ServiceBrokerWaitRequest @QueueName, %2, @Handle OUTPUT, @Request OUTPUT, @MsgType OUTPUT, @SvcName OUTPUT" & vbCrLf & _
                                                    "SELECT      @Result AS Result, @Handle AS Handle, @Request AS Request, @MsgType AS MsgType, @SvcName AS SvcName" & vbCrLf & _
                                                    "RAISERROR   ('Result=%d', 10, 0, @Result) WITH NOWAIT" & vbCrLf

Private m_oConfig                   As Object
Private m_oPrinters                 As Object
Private WithEvents m_oConn          As ADODB.Connection
Attribute m_oConn.VB_VarHelpID = -1
Private m_sQueueName                As String
Private m_lQueueTimeout             As Long
Private m_oController               As cServiceController

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunction As String)
    Debug.Print "Critical error: " & Err.Description & " [" & MODULE_NAME & "." & sFunction & "]"
    DebugLog Err.Description & " [" & MODULE_NAME & "." & sFunction & "]", vbLogEventTypeError
End Sub

'=========================================================================
' Properties
'=========================================================================

'=========================================================================
' Methods
'=========================================================================

Public Function Init(oConfig As Object, oPrinters As Object) As Boolean
    Const FUNC_NAME     As String = "Init"
    Dim vElem           As Variant
    Dim lIdx            As Long

    On Error GoTo EH
    Set m_oConfig = oConfig
    Set m_oPrinters = oPrinters
    '--- init DB connection
    Set m_oConn = New ADODB.Connection
    m_oConn.CursorLocation = adUseClient
    m_oConn.Open JsonItem(oConfig, "ConnectString")
    DebugLog Printf(STR_CONN_OPEN & " [" & MODULE_NAME & "." & FUNC_NAME & "]", JsonItem(oConfig, "ConnectString"))
    '--- init Service Broker queue
    m_sQueueName = Zn(C_Str(JsonItem(oConfig, "QueueName")), "UcsFpTargetQueue/" & GetErrorComputerName())
    m_lQueueTimeout = Znl(C_Lng(JsonItem(oConfig, "QueueTimeout")), DEF_QUEUE_TIMEOUT)
    For Each vElem In JsonItem(m_oPrinters, "*/DeviceSerialNo")
        If LenB(C_Str(vElem)) <> 0 Then
            m_oConn.usp_sys_ServiceBrokerSetupService m_sQueueName, "UcsFpTargetService/" & C_Str(vElem), IIf(lIdx = 0, "DROP_EXISTING", "DROP_SERVICE")
            lIdx = lIdx + 1
        End If
    Next
    DebugLog Printf(STR_QUEUE_SETUP & " [" & MODULE_NAME & "." & FUNC_NAME & "]", lIdx, m_sQueueName)
    '--- init service controller
    Set m_oController = New cServiceController
    If Not m_oController.Init(oPrinters) Then
        Set m_oConn = Nothing
        Set m_oController = Nothing
        GoTo QH
    End If
    pvQueueLoopAsync
    '--- success
    Init = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Private Sub pvQueueLoopAsync()
    Const FUNC_NAME     As String = "pvQueueLoopAsync"
    
    On Error GoTo EH
    m_oConn.Execute Printf(STR_QUEUE_LOOP, m_sQueueName, m_lQueueTimeout), , adAsyncExecute Or adAsyncFetch
    '--- ToDo: start heart-beat timer
QH:
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume QH
End Sub

Private Function pvHandleRequest(sSvcName As String, sRequest As String, sResponse As String) As Boolean
    Const FUNC_NAME     As String = "pvHandleRequest"
    Dim vJson           As Variant
    Dim sPrinterId      As String
    Dim sError          As String
    Dim sRawUrl        As String
    Dim dblStartTimerEx As Double
    Dim vSplit          As Variant
    
    On Error GoTo EH
    dblStartTimerEx = TimerEx
    sPrinterId = At(Split(sSvcName, "/"), 1)
    If LenB(sPrinterId) = 0 Then
        sResponse = pvEncodeErrorText(Printf(ERR_INVALID_SERVICE, sSvcName))
        GoTo QH
    End If
    If Left$(sRequest, 1) = "<" Then
        AssignVariant vJson, JsonFromXmlDocument(sRequest)
    ElseIf Not JsonParse(sRequest, vJson, Error:=sError) Then
        sResponse = pvEncodeErrorText(sError)
        GoTo QH
    End If
    sRawUrl = JsonItem(C_Obj(vJson), "Url")
    If Left$(sRawUrl, 1) <> "/" Then
        sResponse = pvEncodeErrorText(Printf(ERR_INVALID_URL, sRawUrl))
        GoTo QH
    End If
    If LCase$("/printers/" & sPrinterId) <> LCase$(Left$(sRawUrl, Len("/printers/" & sPrinterId))) Then
        DebugLog Printf(STR_CROSS_PRINTER_REQUEST, sRawUrl, sSvcName)
    End If
    vSplit = Split2(sRawUrl, "?")
    If Not m_oController.ServiceRequest(At(vSplit, 0), At(vSplit, 1), sRequest, sResponse) Then
        If LenB(sResponse) = 0 Then
            sResponse = pvEncodeErrorText(m_oController.LastError)
            GoTo QH
        End If
    End If
    DebugLog Printf(STR_QUEUE_ELAPSED & " [" & MODULE_NAME & "." & FUNC_NAME & "]", m_sQueueName, sPrinterId, sRawUrl, Format(TimerEx - dblStartTimerEx, "0.000"))
    '--- success
    pvHandleRequest = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
    sResponse = pvEncodeErrorText(Err.Description)
    GoTo QH
End Function

Private Function pvEncodeErrorText(sErrorText As String) As String
    Const FUNC_NAME     As String = "pvEncodeErrorText"
    Dim oResponse       As Object
    
    DebugLog Printf(STR_QUEUE_ERROR & " [" & MODULE_NAME & "." & FUNC_NAME & "]", m_sQueueName, sErrorText), vbLogEventTypeError
    JsonItem(oResponse, "Ok") = False
    JsonItem(oResponse, "ErrorText") = sErrorText
    pvEncodeErrorText = JsonDump(oResponse, Minimize:=True)
End Function

'=========================================================================
' Events
'=========================================================================

Private Sub m_oConn_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
    Const FUNC_NAME     As String = "m_oConn_ExecuteComplete"
    Dim sResponse       As String
    
    On Error GoTo EH
    If Not pRecordset Is Nothing Then
        If pRecordset.State = adStateOpen Then
            If pRecordset.Fields.Count = 5 Then
                If C_Lng(pRecordset!Result.Value) = 0 And LenB(C_Str(pRecordset!Handle.Value)) <> 0 Then
                    '--- note: system msgs are "__" prefixed (e.g. "__FIN__", "__ACK__")
                    If Left$(C_Str(pRecordset!Request.Value), 2) <> "__" Then
                        pvHandleRequest C_Str(pRecordset!SvcName.Value), C_Str(pRecordset!Request.Value), sResponse
                        m_oConn.Execute Printf("SEND ON CONVERSATION '%1' (N'%2')", pRecordset!Handle.Value, Replace(sResponse, "'", "''"))
                    End If
                End If
QH:
                pvQueueLoopAsync
            End If
        End If
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume QH
End Sub
