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

#Const ImplNoIdeProtection = (MST_NO_IDE_PROTECTION <> 0)
#Const ImplSelfContained = True

'=========================================================================
' API
'=========================================================================

'--- for thunks
Private Const MEM_COMMIT                    As Long = &H1000
Private Const PAGE_EXECUTE_READWRITE        As Long = &H40
Private Const CRYPT_STRING_BASE64           As Long = 1

Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function CryptStringToBinary Lib "crypt32" Alias "CryptStringToBinaryA" (ByVal pszString As String, ByVal cchString As Long, ByVal dwFlags As Long, ByVal pbBinary As Long, pcbBinary As Long, Optional ByVal pdwSkip As Long, Optional ByVal pdwFlags As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
#If Not ImplNoIdeProtection Then
    Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, ByVal lpszClass As String, ByVal lpszWindow As String) As Long
    Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
#End If
#If ImplSelfContained Then
    Private Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long
    Private Declare Function SetEnvironmentVariable Lib "kernel32" Alias "SetEnvironmentVariableA" (ByVal lpName As String, ByVal lpValue As String) As Long
#End If

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_CONN_OPEN             As String = "Свързан с %1"
Private Const STR_CONN_REOPEN           As String = "Възстановена връзка с %1"
Private Const STR_QUEUE_SETUP           As String = "Слуша на опашка %2 за %1 услуги"
Private Const STR_QUEUE_ELAPSED         As String = "%1 на принтер %2 изпълни заявка %3 за %4 сек."
Private Const STR_QUEUE_ERROR           As String = "Грешка в %1 на принтер %2: %3"
Private Const STR_CROSS_PRINTER_REQUEST As String = "Заявка %1 не е към устройството на услуга %2"
'--- errors
Private Const ERR_INVALID_SERVICE       As String = "Невалидна Service Broker услуга: %1"
Private Const ERR_INVALID_URL           As String = "Невалиден адрес: %1"
Private Const ERR_INVALID_XML           As String = "Невалиден XML"
Private Const ERR_UNKNOWN               As String = "Неизвестна грешка %1"
'--- queue
Private Const DEF_QUEUE_TIMEOUT         As Long = 5000
Private Const STR_QUEUE_LOOP            As String = "DECLARE     @QueueName SYSNAME = '%1', @Result INT, @Handle UNIQUEIDENTIFIER, @Request NVARCHAR(MAX), @MsgType SYSNAME, @SvcName SYSNAME" & vbCrLf & _
                                                    "EXEC        @Result = dbo.usp_sys_ServiceBrokerWaitRequest @QueueName, %2, @Handle OUTPUT, @Request OUTPUT, @MsgType OUTPUT, @SvcName OUTPUT" & vbCrLf & _
                                                    "SELECT      @Result AS Result, @Handle AS Handle, @Request AS Request, @MsgType AS MsgType, @SvcName AS SvcName" & vbCrLf & _
                                                    "RAISERROR   ('Result=%d', 10, 0, @Result) WITH NOWAIT" & vbCrLf
Private Const STR_SEND_ON_CONV          As String = "SEND ON CONVERSATION '%1' (N'%2')"
Private Const STR_END_CONV_WITH_ERROR   As String = "END CONVERSATION '%1' WITH ERROR = %2 DESCRIPTION = '%3'"

Private m_sLastError                As String
Private m_oConfig                   As Object
Private m_oPrinters                 As Object
Private m_pTimer                    As IUnknown
Private WithEvents m_oConn          As ADODB.Connection
Attribute m_oConn.VB_VarHelpID = -1
Private m_sQueueName                As String
Private m_lQueueTimeout             As Long
Private m_oController               As cServiceController

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunction As String)
    m_sLastError = Err.Description
    Debug.Print "Critical error: " & Err.Description & " [" & MODULE_NAME & "." & sFunction & "]"
    DebugLog Err.Description & " [" & MODULE_NAME & "." & sFunction & "]", vbLogEventTypeError
End Sub

'=========================================================================
' Properties
'=========================================================================

Property Get LastError() As String
    LastError = m_sLastError
End Property

Private Property Get pvAddressOfTimerProc() As cMssqlEndpoint
    Set pvAddressOfTimerProc = InitAddressOfMethod(Me, 0)
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function Init(oConfig As Object, oPrinters As Object) As Boolean
    Const FUNC_NAME     As String = "Init"
    Dim vElem           As Variant
    Dim lIdx            As Long
    Dim oRequestsCache  As Object

    On Error GoTo EH
    Set m_oConfig = oConfig
    Set m_oPrinters = oPrinters
    JsonItem(oRequestsCache, vbNullString) = Empty
    '--- init service controller
    Set m_oController = New cServiceController
    If Not m_oController.Init(oPrinters, oRequestsCache) Then
        m_sLastError = m_oController.LastError
        Set m_oController = Nothing
        GoTo QH
    End If
    '--- init DB connection
    If Not pvOpenConn(oConfig) Then
        Set m_oController = Nothing
        GoTo QH
    End If
    DebugLog Printf(STR_CONN_OPEN & " [" & MODULE_NAME & "." & FUNC_NAME & "]", C_Str(m_oConn))
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
    pvQueueLoopAsync
    '--- success
    Init = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Private Function pvOpenConn(oConfig As Object) As Boolean
    Const FUNC_NAME     As String = "pvOpenConn"
    
    On Error GoTo EH
    '--- start heart-beat timer
    Set m_pTimer = InitFireOnceTimerThunk(Me, pvAddressOfTimerProc.TimerProc, Delay:=m_lQueueTimeout + 3000)
    Set m_oConn = New ADODB.Connection
    m_oConn.CursorLocation = adUseClient
    m_oConn.Open JsonItem(oConfig, "ConnectString")
    '--- success
    pvOpenConn = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Private Function pvQueueLoopAsync() As Boolean
    Const FUNC_NAME     As String = "pvQueueLoopAsync"
    
    On Error GoTo EH
    '--- start heart-beat timer
    Set m_pTimer = InitFireOnceTimerThunk(Me, pvAddressOfTimerProc.TimerProc, Delay:=m_lQueueTimeout + 3000)
    m_oConn.Execute Printf(STR_QUEUE_LOOP, m_sQueueName, m_lQueueTimeout), , adAsyncExecute Or adAsyncFetch
    '--- success
    pvQueueLoopAsync = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Private Function pvHandleRequest(sSvcName As String, sRequest As String, sResponse As String, lErrorCode As Long) As Boolean
    Const FUNC_NAME     As String = "pvHandleRequest"
    Dim vJson           As Variant
    Dim sPrinterId      As String
    Dim sError          As String
    Dim sRawUrl         As String
    Dim dblStartTimerEx As Double
    Dim vSplit          As Variant
    
    On Error GoTo EH
    dblStartTimerEx = TimerEx
    sPrinterId = At(Split(sSvcName, "/"), 1)
    If LenB(sPrinterId) = 0 Then
        lErrorCode = 400
        sResponse = Printf(ERR_INVALID_SERVICE, sSvcName)
        GoTo QH
    End If
    If Left$(sRequest, 1) = "<" Then
        AssignVariant vJson, JsonFromXmlDocument(sRequest)
        If IsEmpty(vJson) Then
            lErrorCode = 400
            sResponse = ERR_INVALID_XML
            GoTo QH
        End If
    ElseIf Not JsonParse(sRequest, vJson, Error:=sError) Then
        lErrorCode = 400
        sResponse = sError
        GoTo QH
    End If
    sRawUrl = JsonItem(C_Obj(vJson), "Url")
    If Left$(sRawUrl, 1) <> "/" Then
        lErrorCode = 400
        sResponse = Printf(ERR_INVALID_URL, sRawUrl)
        GoTo QH
    End If
    If LCase$("/printers/" & sPrinterId) <> LCase$(Left$(sRawUrl, Len("/printers/" & sPrinterId))) Then
        DebugLog Printf(STR_CROSS_PRINTER_REQUEST, sRawUrl, sSvcName)
    End If
    vSplit = Split2(sRawUrl, "?")
    If Not m_oController.ServiceRequest(At(vSplit, 0), At(vSplit, 1), sRequest, sResponse) Then
        lErrorCode = m_oController.LastCode
        sResponse = m_oController.LastError
        GoTo QH
    End If
    DebugLog Printf(STR_QUEUE_ELAPSED & " [" & MODULE_NAME & "." & FUNC_NAME & "]", m_sQueueName, sPrinterId, sRawUrl, Format(TimerEx - dblStartTimerEx, "0.000"))
    '--- success
    pvHandleRequest = True
QH:
    If Not pvHandleRequest Then
        DebugLog Printf(STR_QUEUE_ERROR & " [" & MODULE_NAME & "." & FUNC_NAME & "]", m_sQueueName, sPrinterId, sResponse)
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    sResponse = m_sLastError
    GoTo QH
End Function

Public Function TimerProc() As Long
Attribute TimerProc.VB_MemberFlags = "40"
    Const FUNC_NAME     As String = "TimerProc"
    
    On Error GoTo EH
    '--- if heart-beat timer is reached -> bad things must have happened to connection so re-open it
    If pvOpenConn(m_oConfig) Then
        DebugLog Printf(STR_CONN_REOPEN & " [" & MODULE_NAME & "." & FUNC_NAME & "]", C_Str(m_oConn))
        pvQueueLoopAsync
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

'=========================================================================
' MST
'=========================================================================

Private Function InitAddressOfMethod(pObj As Object, ByVal MethodParamCount As Long) As Object
    Const STR_THUNK     As String = "6AAAAABag+oFV4v6ge9QEMEAgcekEcEAuP9EJAS5+QcAAPOri8LB4AgFuQAAAKuLwsHoGAUAjYEAq7gIAAArq7hEJASLq7hJCIsEq7iBi1Qkq4tEJAzB4AIFCIkCM6uLRCQMweASBcDCCACriTrHQgQBAAAAi0QkCIsAiUIIi0QkEIlCDIHqUBDBAIvCBTwRwQCri8IFUBHBAKuLwgVgEcEAq4vCBYQRwQCri8IFjBHBAKuLwgWUEcEAq4vCBZwRwQCri8IFpBHBALn5BwAAq4PABOL6i8dfgcJQEMEAi0wkEIkRK8LCEAAPHwCLVCQE/0IEi0QkDIkQM8DCDABmkItUJAT/QgSLQgTCBAAPHwCLVCQE/0oEi0IEg/gAfgPCBABZWotCDGgAgAAAagBSUf/gZpC4AUAAgMIIALgBQACAwhAAuAFAAIDCGAC4AUAAgMIkAA==" ' 25.3.2019 14:01:08
    Const THUNK_SIZE    As Long = 16728
    Dim hThunk          As Long
    Dim lSize           As Long
    
    hThunk = VirtualAlloc(0, THUNK_SIZE, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
    If hThunk = 0 Then
        Exit Function
    End If
    Call CryptStringToBinary(STR_THUNK, Len(STR_THUNK), CRYPT_STRING_BASE64, hThunk, THUNK_SIZE)
    lSize = CallWindowProc(hThunk, ObjPtr(pObj), MethodParamCount, GetProcAddress(GetModuleHandle("kernel32"), "VirtualFree"), VarPtr(InitAddressOfMethod))
    Debug.Assert lSize = THUNK_SIZE
End Function

Private Function InitFireOnceTimerThunk(pObj As Object, ByVal pfnCallback As Long, Optional Delay As Long) As IUnknown
    Const STR_THUNK     As String = "6AAAAABag+oFgeogERkAV1aLdCQUg8YIgz4AdCqL+oHHBBMZAIvCBSgSGQCri8IFZBIZAKuLwgV0EhkAqzPAq7kIAAAA86WBwgQTGQBSahj/UhBai/iLwqu4AQAAAKszwKuri3QkFKWlg+8Yi0IMSCX/AAAAUItKDDsMJHULWIsPV/9RFDP/62P/QgyBYgz/AAAAjQTKjQTIjUyIMIB5EwB101jHAf80JLiJeQTHQQiJRCQEi8ItBBMZAAWgEhkAUMHgCAW4AAAAiUEMWMHoGAUA/+CQiUEQiU8MUf90JBRqAGoAiw//URiJRwiLRCQYiTheX7g0ExkALSARGQAFABQAAMIQAGaQi0QkCIM4AHUqg3gEAHUkgXgIwAAAAHUbgXgMAAAARnUSi1QkBP9CBItEJAyJEDPAwgwAuAJAAIDCDACQi1QkBP9CBItCBMIEAA8fAItUJAT/SgSLQgR1HYtCDMZAEwCLCv9yCGoA/1Eci1QkBIsKUv9RFDPAwgQAi1QkBIsKi0EohcB0J1L/0FqD+AF3SYsKUv9RLFqFwHU+iwpSavD/cSD/USRaqQAAAAh1K4sKUv9yCGoA/1EcWv9CBDPAUFT/chD/UhSLVCQIx0IIAAAAAFLodv///1jCFABmkA==" ' 27.3.2019 9:14:57
    Const THUNK_SIZE    As Long = 5652
    Static hThunk       As Long
    Dim aParams(0 To 9) As Long
    Dim lSize           As Long
    
    aParams(0) = ObjPtr(pObj)
    aParams(1) = pfnCallback
    #If ImplSelfContained Then
        If hThunk = 0 Then
            hThunk = pvThunkGlobalData("InitFireOnceTimerThunk")
        End If
    #End If
    If hThunk = 0 Then
        hThunk = VirtualAlloc(0, THUNK_SIZE, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
        If hThunk = 0 Then
            Exit Function
        End If
        Call CryptStringToBinary(STR_THUNK, Len(STR_THUNK), CRYPT_STRING_BASE64, hThunk, THUNK_SIZE)
        aParams(2) = GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemAlloc")
        aParams(3) = GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemFree")
        aParams(4) = GetProcAddress(GetModuleHandle("user32"), "SetTimer")
        aParams(5) = GetProcAddress(GetModuleHandle("user32"), "KillTimer")
        '--- for IDE protection
        Debug.Assert pvGetIdeOwner(aParams(6))
        If aParams(6) <> 0 Then
            aParams(7) = GetProcAddress(GetModuleHandle("user32"), "GetWindowLongA")
            aParams(8) = GetProcAddress(GetModuleHandle("vba6"), "EbMode")
            aParams(9) = GetProcAddress(GetModuleHandle("vba6"), "EbIsResetting")
        End If
        #If ImplSelfContained Then
            pvThunkGlobalData("InitFireOnceTimerThunk") = hThunk
        #End If
    End If
    lSize = CallWindowProc(hThunk, 0, Delay, VarPtr(aParams(0)), VarPtr(InitFireOnceTimerThunk))
    Debug.Assert lSize = THUNK_SIZE
End Function

Private Function pvGetIdeOwner(hIdeOwner As Long) As Boolean
    #If Not ImplNoIdeProtection Then
        Dim lProcessId      As Long
        
        Do
            hIdeOwner = FindWindowEx(0, hIdeOwner, "IDEOwner", vbNullString)
            Call GetWindowThreadProcessId(hIdeOwner, lProcessId)
        Loop While hIdeOwner <> 0 And lProcessId <> GetCurrentProcessId()
    #End If
    pvGetIdeOwner = True
End Function

#If ImplSelfContained Then
Private Property Get pvThunkGlobalData(sKey As String) As Long
    Dim sBuffer     As String
    
    sBuffer = String$(50, 0)
    Call GetEnvironmentVariable("_MST_GLOBAL" & GetCurrentProcessId() & "_" & sKey, sBuffer, Len(sBuffer) - 1)
    pvThunkGlobalData = Val(Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1))
End Property

Private Property Let pvThunkGlobalData(sKey As String, ByVal lValue As Long)
    Call SetEnvironmentVariable("_MST_GLOBAL" & GetCurrentProcessId() & "_" & sKey, lValue)
End Property
#End If

'=========================================================================
' Events
'=========================================================================

Private Sub m_oConn_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
    Const FUNC_NAME     As String = "m_oConn_ExecuteComplete"
    Dim sHandle         As String
    Dim sResponse       As String
    Dim lErrorCode      As Long
    
    On Error GoTo EH
    If Not pRecordset Is Nothing Then
        If pRecordset.State = adStateOpen Then
            If pRecordset.Fields.Count = 5 Then
                If C_Lng(pRecordset!Result.Value) = 0 And LenB(C_Str(pRecordset!Handle.Value)) <> 0 Then
                    sHandle = pRecordset!Handle.Value
                    '--- note: system msgs are "__" prefixed (e.g. "__FIN__", "__ACK__")
                    If Left$(C_Str(pRecordset!Request.Value), 2) <> "__" Then
                        If pvHandleRequest(C_Str(pRecordset!SvcName.Value), C_Str(pRecordset!Request.Value), sResponse, lErrorCode) Then
                            m_oConn.Execute Printf(STR_SEND_ON_CONV, Quote(sHandle), Quote(sResponse))
                        Else
                            m_oConn.Execute Printf(STR_END_CONV_WITH_ERROR, Quote(sHandle), Znl(lErrorCode, 500), _
                                Quote(Zn(sResponse, Printf(ERR_UNKNOWN, lErrorCode))))
                        End If
                    End If
                End If
                pvQueueLoopAsync
            End If
        End If
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume NextLine
NextLine:
    On Error Resume Next
    If LenB(sHandle) <> 0 Then
        m_oConn.Execute Printf(STR_END_CONV_WITH_ERROR, Quote(sHandle), Znl(lErrorCode, 500), _
            Quote(Zn(m_sLastError, Printf(ERR_UNKNOWN, lErrorCode))))
    End If
    pvQueueLoopAsync
End Sub
