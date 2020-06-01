VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMssqlQueue"
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
Private Const MODULE_NAME As String = "cMssqlQueue"
Implements IQueue

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_SVC_PREFIX            As String = "UcsFpTargetService"
Private Const STR_QUEUE_SETUP           As String = "Опашка %1 слуша за %2 услуги"
Private Const STR_QUEUE_TERMINATE       As String = "Oпашка %1 е терминирана"
Private Const STR_QUEUE_HEARTBEAT       As String = "Oпашка %1 ще бъде рестартирана (не отговори в рамките на %2 сек.)"
Private Const STR_QUEUE_ERROR           As String = "Опашка %1 за принтер %2 върна грешка %3"
Private Const STR_QUEUE_LOOP            As String = "DECLARE     @QueueName SYSNAME = '%1', @Result INT, @Handle UNIQUEIDENTIFIER, @Request NVARCHAR(MAX), @MsgType SYSNAME, @SvcName SYSNAME, @ErrorText NVARCHAR(255)" & vbCrLf & _
                                                    "EXEC        @Result = dbo.usp_sys_ServiceBrokerWaitRequest @QueueName, %2, @Handle OUTPUT, @Request OUTPUT, @MsgType OUTPUT, @SvcName OUTPUT, @ErrorText OUTPUT" & vbCrLf & _
                                                    "SELECT      @Result AS Result, @Handle AS Handle, @Request AS Request, @MsgType AS MsgType, @SvcName AS SvcName, @ErrorText AS ErrorText" & vbCrLf & _
                                                    "RAISERROR   ('Result=%d', 10, 0, @Result) WITH NOWAIT" & vbCrLf
Private Const LNG_QUEUE_FIELDS          As Long = 6
Private Const STR_SEND_ON_CONV          As String = "SEND ON CONVERSATION '%1' (N'%2')"
Private Const STR_END_CONV_WITH_ERROR   As String = "END CONVERSATION '%1' WITH ERROR = %2 DESCRIPTION = '%3'"
Private Const ERR_UNKNOWN               As String = "Неизвестна грешка %1"
Private Const ERR_COMMUNCATION_FAILURE  As String = "%Communication link failure%"
'--- numeric
Private Const ERR_DB_E_NOTABLE          As Long = &H80040E37

Private m_pTimeoutTimer             As IUnknown
Attribute m_pTimeoutTimer.VB_VarHelpID = -1
Private m_oConn                     As ADODB.Connection
Attribute m_oConn.VB_VarHelpID = -1
Private WithEvents m_oRecvConn      As ADODB.Connection
Attribute m_oRecvConn.VB_VarHelpID = -1
Private m_vPrinters                 As Variant
Private m_sQueueName                As String
Private m_lQueueTimeout             As Long
Private m_oCallback                 As cQueueEndpoint
Private m_sLastError                As String
Private m_sSqlInstance              As String
Private m_sSqlDatabase              As String

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

Private Property Get pvAddressOfTimerProc() As cMssqlQueue
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
    Dim lPos            As Long
    Dim lIdx            As Long
    Dim vElem           As Variant
    
    On Error GoTo EH
    '-- setup initiator queue
    m_oConn.usp_sys_ServiceBrokerSetupService
    '-- setup target queue
    lPos = InStr(m_sQueueName, "/")
    m_oConn.usp_sys_ServiceBrokerSetupService m_sQueueName, STR_SVC_PREFIX & Mid$(m_sQueueName, lPos), "DROP_EXISTING"
    lIdx = lIdx + 1
    For Each vElem In m_vPrinters
        If LenB(C_Str(vElem)) <> 0 Then
            m_oConn.usp_sys_ServiceBrokerSetupService m_sQueueName, STR_SVC_PREFIX & "/" & C_Str(vElem), "DROP_SERVICE"
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
    m_oRecvConn.Execute Printf(STR_QUEUE_LOOP, m_sQueueName, -1), , adAsyncExecute Or adAsyncFetch
    '--- success
    pvQueueLoopAsync = True
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Private Function pvQueueTerminate() As Boolean
    Const FUNC_NAME     As String = "pvQueueTerminate"
    Dim lPos            As Long
    
    On Error GoTo EH
    If pvIsConnBusy(m_oRecvConn) Then
        m_oRecvConn.Cancel
    End If
    If pvIsConnAlive(m_oConn) Then
        lPos = InStr(m_sQueueName, "/")
        m_oConn.usp_sys_ServiceBrokerSetupService m_sQueueName, STR_SVC_PREFIX & Mid$(m_sQueueName, lPos), "DROP_ONLY"
        DebugLog MODULE_NAME, FUNC_NAME, Printf(STR_QUEUE_TERMINATE, m_sQueueName)
    End If
    '--- success
    pvQueueTerminate = True
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Private Function pvIsConnAlive(cn As ADODB.Connection) As Boolean
    On Error GoTo QH
    With cn.Execute("SELECT @@SERVERNAME, DB_NAME()")
        m_sSqlInstance = .Fields.Item(0).Value
        m_sSqlDatabase = .Fields.Item(1).Value
    End With
    '--- success
    pvIsConnAlive = True
QH:
End Function

Private Function pvIsConnBusy(cn As ADODB.Connection) As Boolean
    On Error GoTo QH
    pvIsConnBusy = (cn.State And adStateExecuting) <> 0
QH:
End Function

Private Function pvIsQueueAlive(cn As ADODB.Connection) As Boolean
    Dim lPos            As Long
    Dim sHandle         As String
    
    On Error GoTo QH
    lPos = InStr(m_sQueueName, "/")
    With pvPrepareServiceBrokerSendCommand(cn, Request:="__PING__", TargetSvc:=STR_SVC_PREFIX & Mid$(m_sQueueName, lPos))
        .Execute
        If .Parameters("RetVal").Value <> 0 Then
            GoTo QH
        End If
        If .Parameters("Response").Value <> "__PONG__" Then
            GoTo QH
        End If
        sHandle = .Parameters("Handle").Value
    End With
    '--- close conversation
    With pvPrepareServiceBrokerSendCommand(cn, Handle:=sHandle)
        .Execute
    End With
    '--- success
    pvIsQueueAlive = True
QH:
End Function

Private Function pvPrepareServiceBrokerSendCommand( _
            cn As ADODB.Connection, _
            Optional Request As String, _
            Optional TargetSvc As String, _
            Optional Handle As String) As ADODB.Command
    Const FUNC_NAME     As String = "pvPrepareServiceBrokerSendCommand"
    
    On Error GoTo EH
    Set pvPrepareServiceBrokerSendCommand = New ADODB.Command
    With pvPrepareServiceBrokerSendCommand
        .CommandText = "usp_sys_ServiceBrokerSend"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("RetVal", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("Request", adVarWChar, adParamInput, 1000)
        .Parameters.Append .CreateParameter("Response", adVarWChar, adParamOutput, 1000)
        .Parameters.Append .CreateParameter("TargetSvc", adVarWChar, adParamInput, 120)
        .Parameters.Append .CreateParameter("Handle", adGUID, adParamInputOutput)
        If LenB(Request) <> 0 Then
            .Parameters("Request").Value = Request
        End If
        If LenB(TargetSvc) <> 0 Then
            .Parameters("TargetSvc").Value = TargetSvc
        End If
        If LenB(Handle) <> 0 Then
            .Parameters("Handle").Value = Handle
        End If
        Set .ActiveConnection = cn
        .CommandTimeout = 3
    End With
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Public Function TimerProc() As Long
    Const FUNC_NAME     As String = "TimerProc"
    Dim oConn           As Object
    Dim oRecvConn       As Object
    
    On Error GoTo EH
    If pvIsQueueAlive(m_oConn) Then
        GoTo QH
    End If
    '--- if heart-beat timer is reached -> bad things must have happened to connection so re-open it
    DebugLog MODULE_NAME, FUNC_NAME, Printf(STR_QUEUE_HEARTBEAT, m_sQueueName, (m_lQueueTimeout) / 1000)
    pvQueueTerminate
    If m_oCallback.frReopenConn(oConn, oRecvConn) Then
        Set m_oConn = oConn
        Set m_oRecvConn = oRecvConn
        pvQueueLoopAsync
    End If
QH:
    pvRestartTimer
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Sub pvRestartTimer()
    Const FUNC_NAME     As String = "pvRestartTimer"
    
    On Error GoTo EH
    '--- start heart-beat timer
    Set m_pTimeoutTimer = InitFireOnceTimerThunk(Me, pvAddressOfTimerProc.TimerProc, Delay:=m_lQueueTimeout)
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

'=========================================================================
' Events
'=========================================================================

Private Sub m_oRecvConn_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
    Const FUNC_NAME     As String = "m_oRecvConn_ExecuteComplete"
    Static bHere        As Boolean
    Dim sHandle         As String
    Dim sResponse       As String
    Dim lErrorCode      As Long
    Dim sPrinterID      As String
    Dim vErr            As Variant
    Dim oConn           As ADODB.Connection
    Dim oRecvConn       As ADODB.Connection
    
    On Error GoTo EH
    If bHere Then
        Exit Sub
    End If
    bHere = True
    If Not pError Is Nothing Then
        vErr = Array(pError.Number, pError.Source, pError.Description)
        If vErr(0) <> ERR_DB_E_NOTABLE And vErr(0) <> adErrOperationCancelled Then
            DebugLog MODULE_NAME, FUNC_NAME & "(" & m_sSqlInstance & "/" & m_sSqlDatabase & ")", C_Str(vErr(2)) & " &H" & Hex$(vErr(0)), vbLogEventTypeError
        End If
        GoTo NextLine
    ElseIf Not pRecordset Is Nothing Then
        If (pRecordset.State And adStateOpen) <> 0 Then
            If pRecordset.Fields.Count = LNG_QUEUE_FIELDS Then
                pvQueueLoopAsync
                Set m_pTimeoutTimer = Nothing
                sHandle = C_Str(pRecordset!Handle.Value)
                If C_Lng(pRecordset!Result.Value) = 0 And LenB(sHandle) <> 0 Then
                    If m_oCallback.frHandleRequest(C_Str(pRecordset!SvcName.Value), C_Str(pRecordset!Request.Value), sResponse, lErrorCode) Then
                        m_oConn.Execute Printf(STR_SEND_ON_CONV, Quote(sHandle), Quote(sResponse))
                    Else
                        lErrorCode = Znl(lErrorCode, 500)
                        m_oConn.Execute Printf(STR_END_CONV_WITH_ERROR, Quote(sHandle), lErrorCode, _
                            Quote(Zn(sResponse, Printf(ERR_UNKNOWN, lErrorCode))))
                    End If
                ElseIf C_Lng(pRecordset!Result.Value) <> 99 Then
                    sResponse = C_Str(pRecordset!ErrorText.Value)
                    If LenB(sResponse) <> 0 Then
                        sPrinterID = At(Split(C_Str(pRecordset!SvcName.Value), "/"), 1)
                        DebugLog MODULE_NAME, FUNC_NAME, Printf(STR_QUEUE_ERROR, m_sQueueName, sPrinterID, sResponse), vbLogEventTypeError
                    End If
                End If
                pvRestartTimer
            End If
        End If
    End If
QH:
    bHere = False
    Exit Sub
EH:
    vErr = Array(Err.Number, Err.Source, Err.Description)
    If vErr(0) <> ERR_DB_E_NOTABLE And vErr(0) <> adErrOperationCancelled Then
        PrintError FUNC_NAME
    End If
    Resume NextLine
NextLine:
    On Error Resume Next
    If vErr(2) Like ERR_COMMUNCATION_FAILURE Then
        Set m_oConn = Nothing
        Set m_oRecvConn = Nothing
    End If
    If pvIsConnAlive(m_oConn) Then
        If vErr(0) = ERR_DB_E_NOTABLE Then
            pvQueueSetup
        End If
        If LenB(sHandle) <> 0 Then
            lErrorCode = Znl(lErrorCode, 500)
            m_oConn.Execute Printf(STR_END_CONV_WITH_ERROR, Quote(sHandle), lErrorCode, _
                Quote(Zn(m_sLastError, Printf(ERR_UNKNOWN, lErrorCode))))
        End If
        pvQueueLoopAsync
    Else
        Set m_oConn = Nothing
        Set m_oRecvConn = Nothing
        If Not m_oCallback.frReopenConn(oConn, oRecvConn) Then
            '--- do nothing
        Else
            Set m_oConn = oConn
            Set m_oRecvConn = oRecvConn
            If LenB(sHandle) <> 0 Then
                lErrorCode = Znl(lErrorCode, 500)
                m_oConn.Execute Printf(STR_END_CONV_WITH_ERROR, Quote(sHandle), lErrorCode, _
                    Quote(Zn(m_sLastError, Printf(ERR_UNKNOWN, lErrorCode))))
            End If
            pvQueueLoopAsync
        End If
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
