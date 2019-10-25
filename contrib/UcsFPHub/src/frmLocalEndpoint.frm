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
Private Const MODULE_NAME As String = "frmLocalEndpoint"
Implements IEndpoint

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_MONIKER               As String = "UcsFPHub.LocalEndpoint"
Private Const STR_COM_SETUP             As String = "Слуша на COM сървър с моникер %1"
Private Const ERR_REGISTATION_FAILED    As String = "Невъзможна COM регистрация на моникер %1"

Private m_sLastError                As String
Private m_oController               As cServiceController
Private m_lCookie                   As Long

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunction As String)
    m_sLastError = Err.Description
    #If USE_DEBUG_LOG <> 0 Then
        DebugLog Err.Description & " &H" & Hex$(Err.Number) & " [" & MODULE_NAME & "." & sFunction & "(" & Erl & ")]", vbLogEventTypeError
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
    JsonItem(oRequestsCache, vbNullString) = Empty
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
    DebugLog Printf(STR_COM_SETUP, STR_MONIKER) & " [" & MODULE_NAME & "." & FUNC_NAME & "]"
    '--- success
    frInit = True
QH:
    If LenB(m_sLastError) <> 0 Then
        DebugLog m_sLastError, vbLogEventTypeError
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
    DebugDataDump "sRawUrl=" & sRawUrl & ", sRequest=" & Replace(sRequest, vbCrLf, "^p") & " [" & MODULE_NAME & "." & FUNC_NAME & "]"
    vSplit = Split2(sRawUrl, "?")
    ServiceRequest = m_oController.ServiceRequest(At(vSplit, 0), At(vSplit, 1), sRequest, sResponse)
QH:
    DebugDataDump "sResponse=" & sResponse & " [" & MODULE_NAME & "." & FUNC_NAME & "]"
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

Public Sub ShowConfig()
    Const FUNC_NAME     As String = "ShowConfig"
    Dim oForm           As Object
    
    On Error GoTo EH
    For Each oForm In Forms
        If TypeOf oForm Is frmIcon Then
            oForm.ShowConfig
        End If
    Next
QH:
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume QH
End Sub

Public Sub ShutDown()
    Const FUNC_NAME     As String = "Shutdown"
    Dim oForm           As Object
    
    For Each oForm In Forms
        If TypeOf oForm Is frmIcon Then
            oForm.ShutDown
        End If
    Next
    If IsRunningAsService Then
        NtServiceStop
    End If
QH:
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume QH
End Sub

Public Sub Restart(Optional AddParam As Variant)
    Const FUNC_NAME     As String = "Restart"
    Dim oForm           As Object
    
    For Each oForm In Forms
        If TypeOf oForm Is frmIcon Then
            oForm.Restart AddParam
        End If
    Next
QH:
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume QH
End Sub

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
