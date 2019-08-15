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
    Debug.Print "Critical error: " & Err.Description & " [" & MODULE_NAME & "." & sFunction & "]"
    DebugLog Err.Description & " [" & MODULE_NAME & "." & sFunction & "]", vbLogEventTypeError
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

Public Function Init(oConfig As Object, oPrinters As Object) As Boolean
    Const FUNC_NAME     As String = "Init"
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
    m_lCookie = PutObject(Me, STR_MONIKER)
    If m_lCookie = 0 Then
        m_sLastError = Printf(ERR_REGISTATION_FAILED, STR_MONIKER)
        Set m_oController = Nothing
        GoTo QH
    End If
    DebugLog Printf(STR_COM_SETUP & " [" & MODULE_NAME & "." & FUNC_NAME & "]", STR_MONIKER)
    '--- success
    Init = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Public Sub Terminate()
    If m_lCookie <> 0 Then
        RevokeObject m_lCookie
        m_lCookie = 0
    End If
End Sub

Public Function ServiceRequest(sPath As String, sQueryString As String, sRequest As String, sResponse As String) As Boolean
    ServiceRequest = m_oController.ServiceRequest(sPath, sQueryString, sRequest, sResponse)
End Function

'=========================================================================
' Base class events
'=========================================================================

Private Sub Form_Terminate()
    Terminate
End Sub
