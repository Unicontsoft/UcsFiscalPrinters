VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSysTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================================================
'
' Poor Man's SysTray Support (c) 2019-2022 by wqweto@gmail.com
'
' Based on systray by Ellis Dee of vbforums.com
'
' This project is licensed under the terms of the MIT license
' See the LICENSE file in the project root for more information
'
'=========================================================================
Option Explicit
DefObj A-Z

#Const HasPtrSafe = (VBA7 <> 0)

'=========================================================================
' Public events
'=========================================================================

Public Event Click()
Public Event ContextMenu()
Public Event DblClick()
Public Event BalloonHide()
Public Event BalloonTimeout()
Public Event BalloonClick()

'=========================================================================
' Public enums
'=========================================================================

Public Enum UcsBalloonIconEnum
   ucsBicNone
   ucsBicInformation
   ucsBicWarning
   ucsBicError
   ucsBicNoSound = &H10
End Enum

'=========================================================================
' API
'=========================================================================

'--- Windows Messages
Private Const WM_LBUTTONDOWN                    As Long = &H201
Private Const WM_LBUTTONUP                      As Long = &H202
Private Const WM_LBUTTONDBLCLK                  As Long = &H203
Private Const WM_RBUTTONUP                      As Long = &H205
Private Const WM_USER                           As Long = &H400
Private Const NIN_BALLOONHIDE                   As Long = (WM_USER + 3)
Private Const NIN_BALLOONTIMEOUT                As Long = (WM_USER + 4)
Private Const NIN_BALLOONUSERCLICK              As Long = (WM_USER + 5)
'--- for Shell_NotifyIcon
Private Const NIM_ADD                           As Long = &H0
Private Const NIM_MODIFY                        As Long = &H1
Private Const NIM_DELETE                        As Long = &H2
'--- for NOTIFYICONDATA.uFlags
Private Const NIF_ICON                          As Long = &H2
Private Const NIF_MESSAGE                       As Long = &H1
Private Const NIF_TIP                           As Long = &H4
Private Const NIF_STATE                         As Long = &H8
Private Const NIF_INFO                          As Long = &H10
'--- for NOTIFYICONDATA.dwState
Private Const NIS_HIDDEN                        As Long = 1

#If HasPtrSafe Then
Private Declare PtrSafe Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconW" (ByVal dwMessage As Long, ByVal pnid As LongPtr) As Boolean
Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function SetCapture Lib "user32" (ByVal hWnd As LongPtr) As Long
#Else
Private Enum LongPtr
    [_]
End Enum
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconW" (ByVal dwMessage As Long, ByVal pnid As LongPtr) As Boolean
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As LongPtr) As Long
#End If

Private Type NOTIFYICONDATA
    cbSize              As Long
    hWnd                As LongPtr
    uID                 As Long
    uFlags              As Long
    uCallbackMessage    As Long
    hIcon               As LongPtr
    szTip               As String * 128
    dwState             As Long
    dwStateMask         As Long
    szInfo              As String * 256
    uTimeoutOrVersion   As Long
    szInfoTitle         As String * 64
    dwInfoFlags         As Long
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Private WithEvents m_oForm      As VB.Form
Attribute m_oForm.VB_VarHelpID = -1
Private m_bVisible              As Boolean
Private m_sTooltip              As String
Private m_oIcon                 As StdPicture
Private m_uData                 As NOTIFYICONDATA

'=========================================================================
' Properties
'=========================================================================

Property Get Visible() As Boolean
    Visible = m_bVisible
End Property

Property Let Visible(ByVal bValue As Boolean)
    m_bVisible = bValue
    m_uData.uFlags = NIF_STATE
    m_uData.dwState = IIf(bValue, 0, NIS_HIDDEN)
    m_uData.dwStateMask = NIS_HIDDEN
    If Shell_NotifyIcon(NIM_MODIFY, VarPtr(m_uData)) = 0 Then
        Err.Raise Err.LastDllError
    End If
End Property

Property Get Icon() As StdPicture
    Set Icon = m_oIcon
End Property

Property Set Icon(oValue As StdPicture)
    Debug.Assert oValue.Type = vbPicTypeIcon
    Set m_oIcon = oValue
    m_uData.uFlags = NIF_ICON
    m_uData.hIcon = oValue.Handle
    If Shell_NotifyIcon(NIM_MODIFY, VarPtr(m_uData)) = 0 Then
        Err.Raise Err.LastDllError
    End If
End Property

Property Get TooltipText() As String
    TooltipText = m_sTooltip
End Property

Property Let TooltipText(ByVal sValue As String)
    m_sTooltip = sValue
    m_uData.uFlags = NIF_TIP
    m_uData.szTip = m_sTooltip & vbNullChar
    If Shell_NotifyIcon(NIM_MODIFY, VarPtr(m_uData)) = 0 Then
        Err.Raise Err.LastDllError
    End If
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function Init(oForm As Form, sTooltip As String, Optional ByVal uID As Long) As Boolean
    Set m_oForm = oForm
    With m_uData
        .cbSize = Len(m_uData)
        .hWnd = m_oForm.hWnd
        .uID = uID
        .uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
        .uCallbackMessage = WM_LBUTTONDOWN
        #If TWINBASIC = 0 Then
            .hIcon = m_oForm.Icon.Handle
        #End If
        .szTip = sTooltip & vbNullChar
    End With
    Debug.Print LenB(m_uData), VarPtr(m_uData.dwInfoFlags) - VarPtr(m_uData)
    If Shell_NotifyIcon(NIM_ADD, VarPtr(m_uData)) = 0 Then
        Err.Raise Err.LastDllError
    End If
    m_bVisible = True
    '--- success
    Init = True
End Function

Public Sub Terminate()
    Shell_NotifyIcon NIM_DELETE, VarPtr(m_uData)
    Set m_oForm = Nothing
End Sub

Public Sub ShowBalloonTip( _
            sMessage As String, _
            Optional ByVal IconType As UcsBalloonIconEnum = ucsBicInformation, _
            Optional Title As String, _
            Optional ByVal Timeout As Long = 30000)
    m_uData.szInfo = sMessage & vbNullChar
    m_uData.szInfoTitle = Title & vbNullChar
    m_uData.uTimeoutOrVersion = Timeout
    m_uData.dwInfoFlags = IconType
    m_uData.uFlags = NIF_INFO
    Shell_NotifyIcon NIM_MODIFY, VarPtr(m_uData)
End Sub

'=========================================================================
' Events
'=========================================================================

Private Sub m_oForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Y = 0 Then
        #If TWINBASIC Then
            X = X * m_oForm.DpiScaleFactorX
        #End If
        Select Case CLng(m_oForm.ScaleX(X, m_oForm.ScaleMode, vbPixels))
        Case WM_LBUTTONUP
            Call SetCapture(0)
            Call SetForegroundWindow(m_uData.hWnd)
            RaiseEvent Click
        Case WM_LBUTTONDBLCLK
            Call SetCapture(0)
            Call SetForegroundWindow(m_uData.hWnd)
            RaiseEvent DblClick
        Case WM_RBUTTONUP
            Call SetForegroundWindow(m_uData.hWnd)
            RaiseEvent ContextMenu
        Case NIN_BALLOONHIDE
            RaiseEvent BalloonHide
        Case NIN_BALLOONTIMEOUT
            RaiseEvent BalloonTimeout
        Case NIN_BALLOONUSERCLICK
            RaiseEvent BalloonClick
        End Select
    End If
End Sub

Private Sub m_oForm_Unload(Cancel As Integer)
    If Cancel = 0 Then
        Terminate
    End If
End Sub

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Terminate()
    Terminate
End Sub
