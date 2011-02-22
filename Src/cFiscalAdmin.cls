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
' $Header: /UcsFiscalPrinter/Src/cFiscalAdmin.cls 2     22.02.11 10:26 Wqw $
'
'   Unicontsoft Fiscal Printers Project
'   Copyright (c) 2008-2011 Unicontsoft
'
'   Nastrojka na fiscalni ustrojstwa
'
' $Log: /UcsFiscalPrinter/Src/cFiscalAdmin.cls $
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
Private Const STR_COMPLUS_IMPORT_PROGID As String = "UcsFiscalPrinters.cFiscalPrinter|UcsFiscalPrinters.cICLProtocol|UcsFiscalPrinters.cEltradeProtocol"

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunc As String)
    Debug.Print MODULE_NAME & "." & sFunc & ": " & Error
    DebugLog sFunc, "Run-time error: " & Error
End Sub

Private Sub DebugLog(sFunc As String, sText As String)
    OutputDebugLog MODULE_NAME, sFunc, sText
End Sub

'=========================================================================
' Properties
'=========================================================================

Public Property Get IsShared() As Boolean
    Const FUNC_NAME     As String = "IsShared [get]"
    Dim lIdx            As Long
    
    On Error GoTo EH
    With CreateObject("COMAdmin.COMAdminCatalog")
        With .GetCollection("Applications")
            .Populate
            For lIdx = .Count - 1 To 0 Step -1
                If LCase(.Item(lIdx).Name) = LCase(STR_COMPLUS_APP_NAME) Then
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
    Dim sDll            As String
    Dim sDllHostImage   As String
    Dim oApp            As Object
    Dim oRole           As Object
    Dim lIdx            As Long
    Dim vElem           As Variant
    
    On Error GoTo EH
    With CreateObject("COMAdmin.COMAdminCatalog")
        If bValue Then
            sDll = RegReadString(HKEY_CLASSES_ROOT, "UcsFiscalPrinters.cFiscalPrinter\Clsid", vbNullString)
            If LenB(sDll) <> 0 Then
                sDll = RegReadString(HKEY_CLASSES_ROOT, "CLSID\" & sDll & "\InProcServer32", vbNullString)
            End If
            If LenB(sDll) <> 0 Then
                With .GetCollection("Applications")
                    .Populate
                    For lIdx = .Count - 1 To 0 Step -1
                        If LCase(.Item(lIdx).Name) = LCase(STR_COMPLUS_APP_NAME) Then
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
                    .ImportComponent oApp.Key, C_Str(vElem)
                Next
            End If
            '--- add dllhost.exe to allowed firewall applications
            sDllHostImage = GetSystemDirectory() & "\dllhost.exe"
            If LenB(Dir(sDllHostImage)) Then
                Set oApp = CreateObject("HNetCfg.FwAuthorizedApplication")
                oApp.ProcessImageFileName = sDllHostImage
                oApp.Name = STR_COMPLUS_APP_NAME
                oApp.Scope = 1      ' NET_FW_SCOPE_LOCAL_SUBNET
                oApp.IpVersion = 2  ' NET_FW_IP_VERSION_ANY
                oApp.Enabled = False
                CreateObject("HNetCfg.FwMgr").LocalPolicy.CurrentProfile.AuthorizedApplications.Add oApp
            End If
        Else
            .ShutdownApplication STR_COMPLUS_APP_NAME
            With .GetCollection("Applications")
                .Populate
                Set oApp = Nothing
                For lIdx = .Count - 1 To 0 Step -1
                    If LCase(.Item(lIdx).Name) = LCase(STR_COMPLUS_APP_NAME) Then
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

'=========================================================================
' Methods
'=========================================================================

Public Function ShowSetup(DeviceString As String, Server As String, OwnerForm As Object) As Boolean
    Dim oFrmIcl         As New frmIclSetup
    Dim oFrmEltrade     As New frmEltradeSetup
    
    '--- figure out model
    Select Case UCase(At(Split(DeviceString, ";"), 0))
    Case STR_MODEL_ELTRADE_ECR
        ShowSetup = oFrmEltrade.frInit(DeviceString, Server, OwnerForm)
    Case STR_MODEL_DATECS_FP, STR_MODEL_DAISY_ECR
        ShowSetup = oFrmIcl.frInit(DeviceString, Server, OwnerForm)
    End Select
End Function

