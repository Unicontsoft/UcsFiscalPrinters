Attribute VB_Name = "mdNtService"
'=========================================================================
'
' NtService Helpers (c) 2019 by wqweto@gmail.com
'
' Based on NT Service module © 2000-2004 Sergey Merzlikin (sm@smsoft.ru)
'
' This project is licensed under the terms of the MIT license
' See the LICENSE file in the project root for more information
'
'=========================================================================
Option Explicit
DefObj A-Z

'=========================================================================
' API
'=========================================================================

Private Const INFINITE                      As Long = -1
Private Const QS_ALLINPUT                   As Long = &H4FF
Private Const SC_MANAGER_CONNECT            As Long = 1
Private Const SERVICE_QUERY_STATUS          As Long = 4

Private Declare Function CreateThread Lib "kernel32" (ByVal lpThreadAttributes As Long, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lpParameter As Long, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function MsgWaitForMultipleObjects Lib "user32" (ByVal nCount As Long, pHandles As Long, ByVal fWaitAll As Long, ByVal dwMilliseconds As Long, ByVal dwWakeMask As Long) As Long
Private Declare Function OpenSCManager Lib "advapi32" Alias "OpenSCManagerW" (ByVal lpMachineName As Long, ByVal lpDatabaseName As Long, ByVal dwDesiredAccess As Long) As Long
Private Declare Function OpenService Lib "advapi32" Alias "OpenServiceW" (ByVal hSCManager As Long, ByVal lpServiceName As Long, ByVal dwDesiredAccess As Long) As Long
Private Declare Function QueryServiceStatus Lib "advapi32" (ByVal hService As Long, lpServiceStatus As SERVICE_STATUS) As Long
Private Declare Function CloseServiceHandle Lib "advapi32" (ByVal hSCObject As Long) As Long

'=========================================================================
' Constants and member variables
'=========================================================================

Private m_sServiceName          As String
Private m_lTimeout              As Long
Private m_lServiceNamePtr       As Long
Private m_hStartEvent           As Long
Private m_hStopEvent            As Long
Private m_hStopPendingEvent     As Long
Private m_hThread               As Long
Private m_uService              As SERVICE_STATUS
Private m_hStatus               As Long

'=========================================================================
' Functions
'=========================================================================

Public Function NtServiceInit(sServiceName As String, Optional ByVal Timeout As Long = 30000) As Boolean
    Dim aHandles(0 To 1)    As Long
    
    '--- check if service installed
    If NtServiceGetStatus(sServiceName) = 0 Then
        GoTo QH
    End If
    '--- init member vars
    m_sServiceName = sServiceName
    m_lTimeout = Timeout
    '--- init extra vars
    m_lServiceNamePtr = StrPtr(m_sServiceName)
    m_hStartEvent = CreateEventW(0, 1, 0, 0)
    m_hStopEvent = CreateEventW(0, 1, 0, 0)
    m_hStopPendingEvent = CreateEventW(0, 1, 0, 0)
    With m_uService
        .dwServiceType = SERVICE_WIN32_OWN_PROCESS
        .dwControlsAccepted = SERVICE_ACCEPT_STOP Or SERVICE_ACCEPT_SHUTDOWN
        .dwWin32ExitCode = 0
        .dwServiceSpecificExitCode = 0
        .dwCheckPoint = 0
        .dwWaitHint = 0
    End With
    m_hThread = CreateThread(0, 0, AddressOf pvThreadProc, 0, 0, 0)
    If m_hThread = 0 Then
        GoTo QH
    End If
    aHandles(0) = m_hStartEvent
    aHandles(1) = m_hThread
    If pvMsgWaitWithDoEvents(2, aHandles(0), m_lTimeout) <> 0 Then
        GoTo QH
    End If
    pvSetStatus SERVICE_RUNNING
    '--- success
    NtServiceInit = True
    Exit Function
QH:
    If m_hStartEvent <> 0 Then
        Call CloseHandle(m_hStartEvent)
        m_hStartEvent = 0
    End If
    If m_hStopEvent <> 0 Then
        Call CloseHandle(m_hStopEvent)
        m_hStopEvent = 0
    End If
    If m_hStopPendingEvent <> 0 Then
        Call CloseHandle(m_hStopPendingEvent)
        m_hStopPendingEvent = 0
    End If
    m_hThread = 0
End Function

Public Function NtServiceQueryStop() As Boolean
    If m_hStopPendingEvent = 0 Then
        NtServiceQueryStop = True
    ElseIf pvMsgWaitWithDoEvents(1, m_hStopPendingEvent, 1000) = 0 Then
        pvSetStatus SERVICE_STOP_PENDING
        NtServiceQueryStop = True
    End If
End Function

Public Function NtServiceTerminate() As Boolean
    If m_hStopEvent <> 0 Then
        Call SetEvent(m_hStopEvent)
        If pvMsgWaitWithDoEvents(1, m_hThread, m_lTimeout) = 0 Then
            NtServiceTerminate = True
        End If
    End If
End Function

Public Function NtServiceGetStatus(sServiceName As String) As SERVICE_STATE
    Dim hSCManager      As Long
    Dim hService        As Long
    Dim uStatus         As SERVICE_STATUS
    
    hSCManager = OpenSCManager(0, 0, SC_MANAGER_CONNECT)
    If hSCManager = 0 Then
        GoTo QH
    End If
    hService = OpenService(hSCManager, StrPtr(sServiceName), SERVICE_QUERY_STATUS)
    If hService = 0 Then
        GoTo QH
    End If
    If QueryServiceStatus(hService, uStatus) = 0 Then
        GoTo QH
    End If
    '--- success
    NtServiceGetStatus = uStatus.dwCurrentState
QH:
    If hService <> 0 Then
        Call CloseServiceHandle(hService)
    End If
    If hSCManager <> 0 Then
        Call CloseServiceHandle(hSCManager)
    End If
End Function

'= private ===============================================================

Private Sub pvThreadProc(ByVal dwDummy As Long)
    Dim uEntry          As SERVICE_TABLE
    
    uEntry.lpServiceName = m_lServiceNamePtr
    uEntry.lpServiceProc = pvAddr(AddressOf pvServiceProc)
    Call StartServiceCtrlDispatcherW(uEntry)
End Sub

Private Sub pvServiceProc(ByVal dwArgc As Long, ByVal lpszArgv As Long)
    m_hStatus = RegisterServiceCtrlHandlerW(m_lServiceNamePtr, AddressOf pvHandlerProc)
    If m_hStatus = 0 Then
        GoTo QH
    End If
    pvSetStatus SERVICE_START_PENDING
    Call SetEvent(m_hStartEvent)
    Call WaitForSingleObject(m_hStopEvent, INFINITE)
    pvSetStatus SERVICE_STOPPED
QH:
End Sub
   
Private Sub pvHandlerProc(ByVal dwControl As Long)
    Select Case dwControl
    Case SERVICE_CONTROL_SHUTDOWN, SERVICE_CONTROL_STOP
        Call SetEvent(m_hStopPendingEvent)
    Case Else
        pvSetStatus m_uService.dwCurrentState
    End Select
End Sub

Private Function pvSetStatus(ByVal eNewState As SERVICE_STATE) As Boolean
    If m_hStatus <> 0 Then
        m_uService.dwCurrentState = eNewState
        '--- success (or failure)
        pvSetStatus = SetServiceStatus(m_hStatus, m_uService) <> 0
    End If
End Function

Private Function pvAddr(ByVal pfn As Long) As Long
    pvAddr = pfn
End Function

Private Function pvMsgWaitWithDoEvents(ByVal nCount As Long, pHandles As Long, ByVal dwMilliseconds As Long) As Long
    Dim dblEndTimer     As Double
    Dim lWaitMs         As Long
    
    dblEndTimer = TimerEx + dwMilliseconds / 1000
    Do
        If dwMilliseconds < 0 Then
            lWaitMs = INFINITE
        Else
            lWaitMs = (dblEndTimer - TimerEx) * 1000
            If lWaitMs < 0 Then
                Exit Do
            End If
        End If
        pvMsgWaitWithDoEvents = MsgWaitForMultipleObjects(nCount, pHandles, 0, lWaitMs, QS_ALLINPUT)
        If pvMsgWaitWithDoEvents <> nCount Then
            Exit Do
        End If
        DoEvents
    Loop
End Function

Private Property Get TimerEx() As Double
    Dim cFreq           As Currency
    Dim cValue          As Currency
    
    Call QueryPerformanceFrequency(cFreq)
    Call QueryPerformanceCounter(cValue)
    TimerEx = cValue / cFreq
End Property

