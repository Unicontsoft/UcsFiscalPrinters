VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFiscalAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=========================================================================
' $Header: /UcsFiscalPrinter/Src/cFiscalAdmin.cls 7     9.08.11 23:25 Wqw $
'
'   Unicontsoft Fiscal Printers Project
'   Copyright (c) 2008-2011 Unicontsoft
'
'   Nastrojka na fiscalni ustrojstwa
'
' $Log: /UcsFiscalPrinter/Src/cFiscalAdmin.cls $
' 
' 7     9.08.11 23:25 Wqw
' REF: supports zeka protocol
'
' 6     4.07.11 15:48 Wqw
' REF: err handling
'
' 5     24.02.11 16:19 Wqw
' REF: IsShared [let]
'
' 4     24.02.11 16:05 Wqw
' ADD: Property CanShare
'
' 3     22.02.11 13:51 Wqw
' ADD: Function CreateObject
'
' 2     22.02.11 10:26 Wqw
' DEL: EnumPorts
'
' 1     22.02.11 10:05 Wqw
' Initial implementation
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "cFiscalAdmin"

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_COMPLUS_APP_NAME      As String = "Unicontsoft Fiscal Printers Component"
Private Const STR_COMPLUS_IMPORT_PROGID As String = "cFiscalPrinter|cICLProtocol|cEltradeProtocol|cZekaProtocol"

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunc As String)
    Debug.Print MODULE_NAME & "." & sFunc & ": " & Err.Description
    OutputDebugLog MODULE_NAME, sFunc & "(" & Erl & ")", "Run-time error: " & Err.Description
End Sub

'Private Sub DebugLog(sFunc As String, sText As String)
'    OutputDebugLog MODULE_NAME, sFunc, sText
'End Sub

'=========================================================================
' Properties
'=========================================================================

Public Property Get IsShared() As Boolean
    Const FUNC_NAME     As String = "IsShared [get]"
    Dim lIdx            As Long
    
    On Error GoTo EH
    With VBA.CreateObject("COMAdmin.COMAdminCatalog")
        With .GetCollection("Applications")
            .Populate
            For lIdx = .Count - 1 To 0 Step -1
                If LCase$(.Item(lIdx).Name) <> LCase$(STR_COMPLUS_APP_NAME) Then
                Else
                    IsShared = True
                    Exit For
                End If
            Next
        End With
    End With
    Exit Property
EH:
    PrintError FUNC_NAME
    Resume Next
End Property

Public Property Let IsShared(ByVal bValue As Boolean)
    Const FUNC_NAME     As String = "IsShared [let]"
    Dim sDllHostImage   As String
    Dim oApp            As Object
    Dim oRole           As Object
    Dim lIdx            As Long
    Dim vElem           As Variant
    
    On Error GoTo EH
    With VBA.CreateObject("COMAdmin.COMAdminCatalog")
        If bValue Then
            For Each vElem In Split(STR_COMPLUS_IMPORT_PROGID, "|")
                If LenB(pvGetDll(LIB_NAME & "." & vElem)) = 0 Then
                    Exit Property
                End If
            Next
            With .GetCollection("Applications")
                .Populate
                For lIdx = .Count - 1 To 0 Step -1
                    If LCase$(.Item(lIdx).Name) = LCase$(STR_COMPLUS_APP_NAME) Then
                        Set oApp = .Item(lIdx)
                        Exit For
                    End If
                Next
                If oApp Is Nothing Then
                    Set oApp = .Add
                    With oApp
                        .Value("Name") = STR_COMPLUS_APP_NAME
    '                    If Len(sUser) <> 0 Then
    '                        .Value("Identity") = sUser
    '                        .Value("Password") = sPass
    '                    End If
                        .Value("Authentication") = 2 '--- "Connect". for win98 clients
                    End With
                    .SaveChanges
                    With .GetCollection("Roles", oApp.Key)
                        .Populate
                        Set oRole = .Add
                        oRole.Value("Name") = "CreatorOwner"
                        .SaveChanges
                        With .GetCollection("UsersInRole", oRole.Key)
                            .Populate
                            .Add.Value("User") = "Everyone"
                            .SaveChanges
                        End With
                    End With
                End If
            End With
            .ShutdownApplication oApp.Key
            For Each vElem In Split(STR_COMPLUS_IMPORT_PROGID, "|")
                .ImportComponent oApp.Key, LIB_NAME & "." & vElem
            Next
            '--- add dllhost.exe to allowed firewall applications
            sDllHostImage = GetSystemDirectory() & "\dllhost.exe"
            If LenB(Dir(sDllHostImage)) Then
                Set oApp = VBA.CreateObject("HNetCfg.FwAuthorizedApplication")
                oApp.ProcessImageFileName = sDllHostImage
                oApp.Name = STR_COMPLUS_APP_NAME
                oApp.Scope = 1      ' NET_FW_SCOPE_LOCAL_SUBNET
                oApp.IpVersion = 2  ' NET_FW_IP_VERSION_ANY
                oApp.Enabled = False
                VBA.CreateObject("HNetCfg.FwMgr").LocalPolicy.CurrentProfile.AuthorizedApplications.Add oApp
            End If
        Else
            .ShutdownApplication STR_COMPLUS_APP_NAME
            With .GetCollection("Applications")
                .Populate
                Set oApp = Nothing
                For lIdx = .Count - 1 To 0 Step -1
                    If LCase$(.Item(lIdx).Name) <> LCase$(STR_COMPLUS_APP_NAME) Then
                    Else
                        Set oApp = .Item(lIdx)
                        .Remove lIdx
                        .SaveChanges
                        Exit For
                    End If
                Next
            End With
        End If
    End With
    Exit Property
EH:
    PrintError FUNC_NAME
    Resume Next
End Property

Public Property Get CanShare() As Boolean
    Dim vElem           As Variant
    
    For Each vElem In Split(STR_COMPLUS_IMPORT_PROGID, "|")
        If LenB(pvGetDll(LIB_NAME & "." & vElem)) = 0 Then
            Exit Property
        End If
    Next
    CanShare = True
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function ShowSetup(DeviceString As String, Server As String, OwnerForm As Object) As Boolean
    Dim oFrmIcl         As New frmIclSetup
    Dim oFrmEltrade     As New frmEltradeSetup
    Dim oFrmZeka        As New frmZekaSetup
    
    '--- figure out protocol
    Select Case UCase$(At(Split(DeviceString, ";"), 0))
    Case STR_PROTOCOL_ELTRADE_ECR
        ShowSetup = oFrmEltrade.frInit(DeviceString, Server, OwnerForm)
    Case STR_PROTOCOL_DATECS_FP, STR_PROTOCOL_DAISY_ECR
        ShowSetup = oFrmIcl.frInit(DeviceString, Server, OwnerForm)
    Case STR_PROTOCOL_ZEKA_FP
        ShowSetup = oFrmZeka.frInit(DeviceString, Server, OwnerForm)
    End Select
End Function

Public Function CreateObject(sProgID As String) As Object
    Select Case LCase$(sProgID)
    Case LCase$(LIB_NAME & ".cFiscalPrinter")
        Set CreateObject = New cFiscalPrinter
    Case LCase$(LIB_NAME & ".cICLProtocol")
        Set CreateObject = New cICLProtocol
    Case LCase$(LIB_NAME & ".cEltradeProtocol")
        Set CreateObject = New cEltradeProtocol
    Case LCase$(LIB_NAME & ".cZekaProtocol")
        Set CreateObject = New cZekaProtocol
    End Select
End Function

Private Function pvGetDll(sProgID As String) As String
    pvGetDll = RegReadString(HKEY_CLASSES_ROOT, sProgID & "\Clsid", vbNullString)
    If LenB(pvGetDll) <> 0 Then
        pvGetDll = Dir(RegReadString(HKEY_CLASSES_ROOT, "CLSID\" & pvGetDll & "\InProcServer32", vbNullString))
    End If
End Function
