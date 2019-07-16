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

'=========================================================================
' API
'=========================================================================

Private Declare Function vbaObjSetAddref Lib "msvbvm60" Alias "__vbaObjSetAddref" (oDest As Any, ByVal lSrcPtr As Long) As Long

'=========================================================================
' Constants and member variables
'=========================================================================

Public ID                           As String
Public WithEvents Socket            As cAsyncSocket
Private m_lParentPtr                As Long

'=========================================================================
' Properties
'=========================================================================

Private Property Get pvParent() As cRestEndpoint
    Call vbaObjSetAddref(pvParent, m_lParentPtr)
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function Init(sKey As String, oSocket As cAsyncSocket, oParent As cRestEndpoint)
    ID = sKey
    Set Socket = oSocket
    m_lParentPtr = ObjPtr(oParent)
End Function

'=========================================================================
' Events
'=========================================================================

Private Sub Socket_OnReceive()
    pvParent.frClientReceive ID, Socket
End Sub

Private Sub Socket_OnClose()
    pvParent.frClientClose ID
End Sub

Private Sub Socket_OnError(ByVal ErrorCode As Long, ByVal EventMask As UcsAsyncSocketEventMaskEnum)
    pvParent.frClientError ID, Socket, ErrorCode, EventMask
End Sub
