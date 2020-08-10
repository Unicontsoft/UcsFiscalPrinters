VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMysqlQueue"
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
Private Const MODULE_NAME As String = "cMysqlQueue"
Implements IQueue

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_SVC_PREFIX            As String = "UcsFpTargetService"
Private Const STR_QUEUE_SETUP           As String = "Опашка %1 слуша за %2 услуги"
Private Const STR_QUEUE_TERMINATE       As String = "Oпашка %1 е терминирана"
Private Const STR_QUEUE_HEARTBEAT       As String = "Oпашка %1 ще бъде рестартирана (не отговори в рамките на %2 сек.)"
Private Const STR_QUEUE_ERROR           As String = "Опашка %1 за принтер %2 върна грешка %3"
Private Const STR_QUEUE_LOOP            As String = "CALL        `usp_umq_wait_request`('%1', %2, @handle, @request, @svc_name, @error_text, @result)"
Private Const LNG_QUEUE_FIELDS          As Long = 5
Private Const STR_SEND_ON_CONV          As String = "INSERT      `umq_messages`(`conversation_id`, `service_id`, `queue_id`, `message_type`, `message_body`, `created_at`)" & vbCrLf & _
                                                    "SELECT      %1 AS `conversation_id`" & vbCrLf & _
                                                    "            , (SELECT `service_id` FROM umq_conversations WHERE `id` = %1) AS `service_id`" & vbCrLf & _
                                                    "            , (SELECT `queue_id` FROM umq_services WHERE `id` IN (SELECT `service_id` FROM umq_conversations WHERE `id` = %1)) AS `queue_id`" & vbCrLf & _
                                                    "            , 'DEFAULT' AS `message_type`" & vbCrLf & _
                                                    "            , '%2' AS `message_body`" & vbCrLf & _
                                                    "            , CURRENT_TIMESTAMP(3)  AS `created_at`;"
Private Const STR_END_CONV_WITH_ERROR   As String = "INSERT      `umq_messages`(`conversation_id`, `service_id`, `queue_id`, `message_type`, `message_body`, `created_at`)" & vbCrLf & _
                                                    "SELECT      %1 AS `conversation_id`" & vbCrLf & _
                                                    "            , (SELECT `service_id` FROM umq_conversations WHERE `id` = %1) AS `service_id`" & vbCrLf & _
                                                    "            , (SELECT `queue_id` FROM umq_services WHERE `id` IN (SELECT `service_id` FROM umq_conversations WHERE `id` = %1)) AS `queue_id`" & vbCrLf & _
                                                    "            , 'ERROR' AS `message_type`" & vbCrLf & _
                                                    "            , '%3' AS `message_body`" & vbCrLf & _
                                                    "            , CURRENT_TIMESTAMP(3)  AS `created_at`;"
Private Const ERR_UNKNOWN               As String = "Неизвестна грешка %1"
'--- numeric
Private Const ERR_DB_E_NOTABLE          As Long = &H80040E37

Private m_pTimeoutTimer             As IUnknown
Private m_pCompleteTimer            As IUnknown
Private m_oConn                     As ADODB.Connection
Private WithEvents m_oRecvConn      As ADODB.Connection
Attribute m_oRecvConn.VB_VarHelpID = -1
Private m_vPrinters                 As Variant
Private m_sQueueName                As String
Private m_lQueueTimeout             As Long
Private m_oCallback                 As cQueueEndpoint
Private m_sLastError                As String
Private m_sSqlInstance              As String
Private m_sSqlDatabase              As String
Private m_lSqlProcessID             As Long
Private m_pRecordset                As Recordset

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunction As String)
    m_sLastError = Err.Description
    #If USE_DEBUG_LOG <> 0 Then
        DebugLog MODULE_NAME, sFunction & "(" & m_sSqlInstance & "/" & m_sSqlDatabase & ")(" & Erl & ")", Err.Description & " &H" & Hex$(Err.Number), vbLogEventTypeError
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

Private Property Get pvAddressOfTimerProc() As cMysqlQueue
    Set pvAddressOfTimerProc = InitAddressOfMethod(Me, 0)
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function Init(oConn As Object, oRecvConn As Object, vPrinters As Variant, sQueueName As String, ByVal lQueueTimeout As Long, oCallback As Object) As Boolean
    Const FUNC_NAME     As String = "Init"
    
    On Error GoTo EH
    Set m_oConn = oConn
    Set m_oRecvConn = oRecvConn
    m_vPrinters = vPrinters
    m_sQueueName = sQueueName
    m_lQueueTimeout = lQueueTimeout
    Set m_oCallback = oCallback
    If pvIsConnAlive(m_oConn) Then
        pvQueueSetup
        pvQueueLoopAsync
        pvRestartTimer
        '--- success
        Init = True
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Private Function pvQueueSetup() As Boolean
    Const FUNC_NAME     As String = "pvQueueSetup"
    Dim lIdx            As Long
    Dim vElem           As Variant
    
    On Error GoTo EH
    m_oConn.usp_umq_setup_service m_sQueueName, STR_SVC_PREFIX & "/" & m_lSqlProcessID, "DROP_EXISTING"
    lIdx = lIdx + 1
    For Each vElem In m_vPrinters
        If LenB(C_Str(vElem)) <> 0 Then
            m_oConn.usp_umq_setup_service m_sQueueName, STR_SVC_PREFIX & "/" & C_Str(vElem), "DROP_SERVICE"
            lIdx = lIdx + 1
        End If
    Next
    DebugLog MODULE_NAME, FUNC_NAME, Printf(STR_QUEUE_SETUP, m_sQueueName, lIdx)
    '--- success
    pvQueueSetup = True
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Private Function pvQueueLoopAsync() As Boolean
    Const FUNC_NAME     As String = "pvQueueLoopAsync"
    
    On Error GoTo EH
    m_oRecvConn.Execute Printf(STR_QUEUE_LOOP, m_sQueueName, m_lQueueTimeout), Options:=adCmdText Or adAsyncExecute Or adAsyncFetch
    '--- success
    pvQueueLoopAsync = True
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Private Function pvQueueTerminate() As Boolean
    Const FUNC_NAME     As String = "pvQueueTerminate"
    
    On Error GoTo EH
    If pvIsConnBusy(m_oRecvConn) Then
        m_oRecvConn.Cancel
    End If
    If pvIsConnAlive(m_oConn) Then
        m_oConn.usp_umq_setup_service m_sQueueName, STR_SVC_PREFIX & "/" & m_lSqlProcessID, "DROP_ONLY"
        DebugLog MODULE_NAME, FUNC_NAME, Printf(STR_QUEUE_TERMINATE, m_sQueueName)
    End If
    '--- success
    pvQueueTerminate = True
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Public Function pvIsConnAlive(cn As ADODB.Connection) As Boolean
    On Error GoTo QH
    With cn.Execute("SELECT @@HOSTNAME, DATABASE(), CONNECTION_ID();", Options:=adCmdText)
        m_sSqlInstance = .Fields.Item(0).Value
        m_sSqlDatabase = .Fields.Item(1).Value
        m_lSqlProcessID = .Fields.Item(2).Value
    End With
    '--- success
    pvIsConnAlive = True
QH:
End Function

Public Function pvIsConnBusy(cn As ADODB.Connection) As Boolean
    On Error GoTo QH
    pvIsConnBusy = (cn.State And adStateExecuting) <> 0
QH:
End Function

Private Sub pvRestartTimer()
    Const FUNC_NAME     As String = "pvRestartTimer"
    
    On Error GoTo EH
    '--- start heart-beat timer
    Set m_pTimeoutTimer = InitFireOnceTimerThunk(Me, pvAddressOfTimerProc.TimeoutTimerProc, Delay:=m_lQueueTimeout + 3000)
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Public Function TimeoutTimerProc() As Long
Attribute TimeoutTimerProc.VB_MemberFlags = "40"
    Const FUNC_NAME     As String = "TimeoutTimerProc"
    Dim oConn           As Object
    Dim oRecvConn       As Object
    
    On Error GoTo EH
    '--- if heart-beat timer is reached -> bad things must have happened to connection so re-open it
    DebugLog MODULE_NAME, FUNC_NAME, Printf(STR_QUEUE_HEARTBEAT, m_sQueueName, (m_lQueueTimeout + 3000) / 1000)
    pvQueueTerminate
    If m_oCallback.frReopenConn(oConn, oRecvConn) Then
        Set m_oConn = oConn
        Set m_oRecvConn = oRecvConn
        pvQueueLoopAsync
    End If
    pvRestartTimer
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Public Function CompleteTimerProc() As Long
Attribute CompleteTimerProc.VB_MemberFlags = "40"
    Const FUNC_NAME     As String = "CompleteTimerProc"
    Dim pRecordset      As ADODB.Recordset
    Dim sHandle         As String
    Dim sResponse       As String
    Dim lErrorCode      As Long
    Dim sPrinterID      As String
    
    On Error GoTo EH
    Set pRecordset = m_pRecordset
    If C_Lng(pRecordset!Result.Value) = 98 Then
        pvQueueSetup
    End If
    pvQueueLoopAsync
    Set m_pTimeoutTimer = Nothing
    sHandle = C_Str(pRecordset!Handle.Value)
    If C_Lng(pRecordset!Result.Value) = 0 And LenB(sHandle) <> 0 Then
        If m_oCallback.frHandleRequest(C_Str(pRecordset!SvcName.Value), C_Str(pRecordset!Request.Value), sResponse, lErrorCode) Then
            m_oConn.Execute Printf(STR_SEND_ON_CONV, Quote(sHandle), Quote(sResponse)), Options:=adCmdText Or adExecuteNoRecords
        Else
            lErrorCode = Znl(lErrorCode, 500)
            m_oConn.Execute Printf(STR_END_CONV_WITH_ERROR, Quote(sHandle), lErrorCode, _
                Quote(Zn(sResponse, Printf(ERR_UNKNOWN, lErrorCode)))), Options:=adCmdText Or adExecuteNoRecords
        End If
    ElseIf C_Lng(pRecordset!Result.Value) <> 99 Then
        sResponse = C_Str(pRecordset!ErrorText.Value)
        If LenB(sResponse) <> 0 Then
            sPrinterID = At(Split(C_Str(pRecordset!SvcName.Value), "/"), 1)
            DebugLog MODULE_NAME, FUNC_NAME, Printf(STR_QUEUE_ERROR, m_sQueueName, sPrinterID, sResponse), vbLogEventTypeError
        End If
    End If
    pvRestartTimer
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

'=========================================================================
' Events
'=========================================================================

Private Sub m_oRecvConn_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
    Const FUNC_NAME     As String = "m_oRecvConn_ExecuteComplete"
    Static bHere        As Boolean
    Dim sHandle         As String
    Dim lErrorCode      As Long
    Dim vErr            As Variant
    
    On Error GoTo EH
    If bHere Then
        Exit Sub
    End If
    bHere = True
    If Not pError Is Nothing Then
        vErr = Array(pError.Number, pError.Source, pError.Description)
        DebugLog MODULE_NAME, FUNC_NAME & "(" & m_sSqlInstance & "/" & m_sSqlDatabase & ")", C_Str(vErr(2)) & " &H" & Hex$(vErr(0)), vbLogEventTypeError
        GoTo NextLine
    ElseIf Not pRecordset Is Nothing Then
        If (pRecordset.State And adStateOpen) <> 0 Then
            If pRecordset.Fields.Count = LNG_QUEUE_FIELDS Then
                sHandle = C_Str(pRecordset!Handle.Value)
                Set m_pRecordset = pRecordset
                Set m_pCompleteTimer = InitFireOnceTimerThunk(Me, pvAddressOfTimerProc.CompleteTimerProc)
            End If
        End If
    End If
QH:
    bHere = False
    Exit Sub
EH:
    vErr = Array(Err.Number, Err.Source, Err.Description)
    PrintError FUNC_NAME
    Resume NextLine
NextLine:
    On Error Resume Next
    If pvIsConnAlive(m_oConn) Then
        If vErr(0) = ERR_DB_E_NOTABLE Then
            pvQueueSetup
        End If
        If LenB(sHandle) <> 0 Then
            lErrorCode = 500
            m_oConn.Execute Printf(STR_END_CONV_WITH_ERROR, Quote(sHandle), lErrorCode, _
                Quote(Zn(m_sLastError, Printf(ERR_UNKNOWN, lErrorCode)))), Options:=adCmdText Or adExecuteNoRecords
        End If
        pvQueueLoopAsync
    Else
        Set m_oConn = Nothing
        Set m_oRecvConn = Nothing
    End If
    If m_pTimeoutTimer Is Nothing Then
        pvRestartTimer
    End If
    GoTo QH
End Sub

'=========================================================================
' IQueue interface
'=========================================================================

Private Function IQueue_Setup() As Boolean
    IQueue_Setup = pvQueueSetup
End Function

Private Function IQueue_LoopAsync() As Boolean
    IQueue_LoopAsync = pvQueueLoopAsync
    pvRestartTimer
End Function

Private Function IQueue_Terminate() As Boolean
    IQueue_Terminate = pvQueueTerminate
End Function

Private Function IQueue_IsConversationAlive() As Boolean
    IQueue_IsConversationAlive = True
End Function
