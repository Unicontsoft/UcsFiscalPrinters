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
Attribute VB_Description = "cFiscalAdmin class is used for library administration and as a class factory for fiscal printers components of the library"
''
' cFiscalAdmin class is used for library administration and as a class factory for fiscal printers components of the
' library.
'
'=========================================================================
' $Header: /UcsFiscalPrinter/Src/cFiscalAdmin.cls 8     20.11.15 16:40 Wqw $
'
'   Unicontsoft Fiscal Printers Project
'   Copyright (c) 2008-2015 Unicontsoft
'
'   Administration of fiscal devices and handlers
'
' $Log: /UcsFiscalPrinter/Src/cFiscalAdmin.cls $
' 
' 8     20.11.15 16:40 Wqw
' REF: add VBDOX documentation and typelib descriptions
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

''
' Returns whether the library is registered under COM+ as an application for remote access
'
Public Property Get IsShared() As Boolean
Attribute IsShared.VB_Description = "Returns whether the library is registered under COM+ as an application for remote access"
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

''
' Used to registers FP classes in a COM+ application for remote access. Needs admin permissions for modifying
' system COM+ configuration.
'
' @return `True` if already registered, `False` if COM+ application is not found
' @remark Can be used only if `CanShare` property is `True`
'
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

''
' Used to check if FP classes are available for COM+ registration. If the user application is using reg-free COM these
' classes are not available in registry and cannot by registered for remoting
'
' @return `True` is fiscal component can be remoted, `False` is registration unavailable (reg-free COM)
'
Public Property Get CanShare() As Boolean
Attribute CanShare.VB_Description = "Used to check if FP classes are available for COM+ registration"
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

''
' Used to show device's specific configuration dialog where the end-user can configure operators, departments
' and other device settings
'
' @param `DeviceString` see `cFiscalPrinter` class documentation
' @param `Server`       network address of DCOM server to be used for device communication. Empty for local machine
' @param `OwnerForm`    reference to owner VB form
' @return `True` if form is confirmed, `False` if cancelled
'
Public Function ShowSetup(DeviceString As String, Server As String, OwnerForm As Object) As Boolean
Attribute ShowSetup.VB_Description = "Used to show device's specific configuration dialog where the end-user can configure operators, departments and other device settings"
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

''
' Class factory for fiscal printers classes. Can be used for reg-free COM instantiation
'
' @param `sProgID` ProgID of FP class
' @return Object reference to FP class instance or `Nothing` is unknown `ProgID`
'
Public Function CreateObject(sProgID As String) As Object
Attribute CreateObject.VB_Description = "Class factory for fiscal printers classes. Can be used for reg-free COM instantiation"
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
