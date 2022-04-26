VERSION 5.00
Begin VB.Form frmLocalEndpoint 
   Caption         =   "Form1"
   ClientHeight    =   2316
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3624
   LinkTopic       =   "Form1"
   ScaleHeight     =   2316
   ScaleWidth      =   3624
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmLocalEndpoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================
'
' UcsFPHub (c) 2019-2022 by Unicontsoft
'
' Unicontsoft Fiscal Printers Hub
'
' This project is licensed under the terms of the MIT license
' See the LICENSE file in the project root for more information
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "frmLocalEndpoint"
Implements IEndpoint

'=========================================================================
' API
'=========================================================================

Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_MONIKER               As String = "UcsFPHub.LocalEndpoint"
Private Const STR_COM_SETUP             As String = "Слуша на COM сървър с моникер %1"
Private Const ERR_REGISTATION_FAILED    As String = "Невъзможна COM регистрация на моникер %1"

Private m_sLastError                As String
Private m_oController               As cServiceController
Private m_lCookie                   As Long
Private m_pTimerAutodetect          As IUnknown

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunction As String)
    m_sLastError = Err.Description
    #If USE_DEBUG_LOG <> 0 Then
        DebugLog MODULE_NAME, sFunction & "(" & Erl & ")", Err.Description & " &H" & Hex$(Err.Number), vbLogEventTypeError
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

Property Get Moniker() As String
    Moniker = STR_MONIKER
End Property

Property Get ProcessID() As String
    ProcessID = GetCurrentProcessId()
End Property

Private Property Get pvAddressOfTimerProc() As frmLocalEndpoint
    Set pvAddressOfTimerProc = InitAddressOfMethod(Me, 0)
End Property

Private Property Get pvSettingsForm() As frmSettings
    Dim oForm       As Object
    
    For Each oForm In Forms
        If TypeOf oForm Is frmSettings Then
            Set pvSettingsForm = oForm
            Exit Property
        End If
    Next
End Property

'=========================================================================
' Methods
'=========================================================================

Friend Function frInit(oConfig As Object, oPrinters As Object) As Boolean
    Const FUNC_NAME     As String = "frInit"
    Const ROTFLAGS_ALLOWANYCLIENT As Long = 2
    Dim oRequestsCache  As Object
    
    On Error GoTo EH
    #If oConfig Then '--- touch args
    #End If
    JsonValue(oRequestsCache, vbNullString) = Empty
    Set m_oController = New cServiceController
    If Not m_oController.Init(oPrinters, oRequestsCache) Then
        m_sLastError = m_oController.LastError
        Set m_oController = Nothing
        GoTo QH
    End If
    m_lCookie = PutObject(Me, STR_MONIKER, IIf(IsRunningAsService, ROTFLAGS_ALLOWANYCLIENT, 0))
    If m_lCookie = 0 Then
        m_sLastError = Printf(ERR_REGISTATION_FAILED, STR_MONIKER)
        Set m_oController = Nothing
        GoTo QH
    End If
    DebugLog MODULE_NAME, FUNC_NAME, Printf(STR_COM_SETUP, STR_MONIKER)
    '--- success
    frInit = True
QH:
    If LenB(m_sLastError) <> 0 Then
        DebugLog MODULE_NAME, FUNC_NAME, m_sLastError, vbLogEventTypeError
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Friend Sub frTerminate()
    Const FUNC_NAME     As String = "frTerminate"
    
    On Error GoTo EH
    If m_lCookie <> 0 Then
        RevokeObject m_lCookie
        m_lCookie = 0
    End If
    Set m_pTimerAutodetect = Nothing
QH:
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume QH
End Sub

Public Function ServiceRequest(sRawUrl As String, sRequest As String, sResponse As String) As Boolean
    Const FUNC_NAME     As String = "ServiceRequest"
    Dim vSplit          As Variant
    
    On Error GoTo EH
    If IsLogDebugEnabled Then
        DebugLog MODULE_NAME, FUNC_NAME, "sRequest=" & Replace(sRequest, vbCrLf, "^p") & ", sRawUrl=" & sRawUrl, vbLogEventTypeDebug
    End If
    vSplit = Split(sRawUrl, "?", Limit:=2)
    ServiceRequest = m_oController.ServiceRequest(At(vSplit, 0), At(vSplit, 1), sRequest, sResponse)
QH:
    If IsLogDebugEnabled Then
        DebugLog MODULE_NAME, FUNC_NAME, "sResponse=" & Replace(sResponse, vbCrLf, "^p"), vbLogEventTypeDebug
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Public Function CreateObject(sProgID As String) As Object
    Const FUNC_NAME     As String = "CreateObject"
    Const LIB_UCSFP     As String = "UcsFP20"
    
    On Error GoTo EH
    Select Case LCase$(sProgID)
    Case LCase$(LIB_UCSFP & ".cFiscalPrinter")
        Set CreateObject = New cFiscalPrinter
    Case LCase$(LIB_UCSFP & ".cIslProtocol")
        Set CreateObject = New cIslProtocol
    Case LCase$(LIB_UCSFP & ".cTremolProtocol")
        Set CreateObject = New cTremolProtocol
    Case LCase$(LIB_UCSFP & ".cEscPosProtocol")
        Set CreateObject = New cEscPosProtocol
    Case Else
        Set CreateObject = VBA.CreateObject(sProgID)
    End Select
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Public Sub ShowConfig(Optional OwnerForm As Object)
    Const FUNC_NAME     As String = "ShowConfig"
    Dim oForm           As frmMain
    
    On Error GoTo EH
    Set oForm = MainForm
    If Not oForm Is Nothing Then
        oForm.ShowConfig OwnerForm
    End If
QH:
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume QH
End Sub

Public Sub ShutDown()
    Const FUNC_NAME     As String = "Shutdown"
    Dim oForm           As frmMain
    
    On Error GoTo EH
    If IsRunningAsService Then
        NtServiceStop
        GoTo QH
    End If
    Set oForm = MainForm
    If Not oForm Is Nothing Then
        oForm.ShutDown
    End If
QH:
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume QH
End Sub

Public Sub Restart()
    Const FUNC_NAME     As String = "Restart"
    Dim oSettings       As frmSettings
    Dim oForm           As frmMain
    
    On Error GoTo EH
    If IsRunningAsService Then
        NtServiceStop
        ShellExec "net", "start " & STR_SERVICE_NAME
        GoTo QH
    End If
    Set oSettings = pvSettingsForm
    If Not oSettings Is Nothing Then
        oSettings.frRestart
        GoTo QH
    End If
    Set oForm = MainForm
    If Not oForm Is Nothing Then
        oForm.Restart
    End If
QH:
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume QH
End Sub

Public Function Autodetect(Optional ByVal Async As Boolean, Optional ByVal Delay As Long) As Boolean
    Const FUNC_NAME     As String = "Autodetect"
    Dim vElem           As Variant
    
    On Error GoTo EH
    For Each vElem In JsonValue(m_oController.Printers, "*/Autodetected")
        If C_Bool(vElem) Then
            '--- success
            Autodetect = True
            Exit For
        End If
    Next
    If Autodetect Then
        If Async Then
            Set m_pTimerAutodetect = InitFireOnceTimerThunk(Me, pvAddressOfTimerProc.TimerAutodetect, Delay:=Delay)
        Else
            Restart
        End If
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Public Function TimerAutodetect() As Long
    Const FUNC_NAME     As String = "TimerAutodetect"
    
    On Error GoTo EH
    Restart
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

'=========================================================================
' Base class events
'=========================================================================

Private Sub Form_Terminate()
    frTerminate
End Sub

'=========================================================================
' IEndpoint interface
'=========================================================================

Private Sub IEndpoint_Terminate()
    frTerminate
End Sub
