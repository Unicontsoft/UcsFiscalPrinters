Attribute VB_Name = "mdNtSvcControl"
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

Private Const SW_HIDE                       As Long = 0
Private Const SW_SHOWDEFAULT                As Long = 10
'--- for ShellExecuteEx
Private Const SEE_MASK_NOCLOSEPROCESS       As Long = &H40
Private Const SEE_MASK_NOASYNC              As Long = &H100
Private Const SEE_MASK_FLAG_NO_UI           As Long = &H400
'--- for WaitForSingleObject
Private Const INFINITE                      As Long = -1
'--- for FormatMessage
Private Const FORMAT_MESSAGE_FROM_SYSTEM    As Long = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200

Private Declare Function ShellExecuteEx Lib "shell32" Alias "ShellExecuteExA" (lpExecInfo As SHELLEXECUTEINFO) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Args As Any) As Long

Private Type SHELLEXECUTEINFO
    cbSize              As Long
    fMask               As Long
    hWnd                As Long
    lpVerb              As String
    lpFile              As String
    lpParameters        As String
    lpDirectory         As Long
    nShow               As Long
    hInstApp            As Long
    '  optional fields
    lpIDList            As Long
    lpClass             As Long
    hkeyClass           As Long
    dwHotKey            As Long
    hIcon               As Long
    hProcess            As Long
End Type

'=========================================================================
' Function
'=========================================================================

Public Function NtServiceInstall(sServiceName As String, sDisplayName As String, sExeFile As String, Optional Error As String) As Boolean
    Dim sParams             As String
    Dim lExitCode           As Long
    
    Select Case NtServiceGetStatus(sServiceName)
    Case SERVICE_RUNNING, SERVICE_START_PENDING
        Call ShellWait("net", "stop " & ArgvQuote(sServiceName), StartHidden:=True)
    End Select
    sParams = "create " & ArgvQuote(sServiceName) & " binPath= " & ArgvQuote(sExeFile) & " DisplayName= " & ArgvQuote(sDisplayName) & " start= auto"
    If Not ShellWait("sc", sParams, StartHidden:=True, ExitCode:=lExitCode) Or lExitCode <> 0 Then
        Error = GetErrorDescription(lExitCode)
        GoTo QH
    End If
    If Not ShellWait("net", "start " & ArgvQuote(sServiceName), StartHidden:=True, ExitCode:=lExitCode) Or lExitCode <> 0 Then
        Error = GetErrorDescription(lExitCode)
        GoTo QH
    End If
    '--- succes
    NtServiceInstall = True
QH:
End Function

Public Function NtServiceUninstall(sServiceName As String, Optional Error As String) As Boolean
    Dim lExitCode           As Long
    
    Select Case NtServiceGetStatus(sServiceName)
    Case SERVICE_RUNNING, SERVICE_START_PENDING
        Call ShellWait("net", "stop " & ArgvQuote(sServiceName), StartHidden:=True)
    End Select
    If Not ShellWait("sc", "delete " & ArgvQuote(sServiceName), StartHidden:=True, ExitCode:=lExitCode) Or lExitCode <> 0 Then
        Error = "Error " & lExitCode
        GoTo QH
    End If
    '--- success
    NtServiceUninstall = True
QH:
End Function

Private Function ShellWait( _
            sFile As String, _
            sParameters As String, _
            Optional ByVal StartHidden As Boolean, _
            Optional Verb As String, _
            Optional ExitCode As Long) As Boolean
    Dim uShell          As SHELLEXECUTEINFO
    
    With uShell
        .cbSize = Len(uShell)
        .fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_NOASYNC Or SEE_MASK_FLAG_NO_UI
        .lpVerb = Verb
        .lpFile = sFile
        .lpParameters = sParameters
        .nShow = IIf(StartHidden, SW_HIDE, SW_SHOWDEFAULT)
    End With
    If ShellExecuteEx(uShell) <> 0 Then
        Call WaitForSingleObject(uShell.hProcess, INFINITE)
        Call GetExitCodeProcess(uShell.hProcess, ExitCode)
        Call CloseHandle(uShell.hProcess)
        '--- success
        ShellWait = True
    Else
        ExitCode = Err.LastDllError
    End If
    If ExitCode <> 0 And LenB(Verb) = 0 Then
        ShellWait = ShellWait(sFile, sParameters, StartHidden, "runas", ExitCode)
    End If
QH:
End Function

Public Function ArgvQuote(sArg As String, Optional ByVal Force As Boolean) As String
    Const WHITESPACE As String = "*[ " & vbTab & vbVerticalTab & vbCrLf & "]*"
    
    If Not Force And LenB(sArg) <> 0 And Not sArg Like WHITESPACE Then
        ArgvQuote = sArg
    Else
        With CreateObject("VBScript.RegExp")
            .Global = True
            .Pattern = "(\\+)($|"")|(\\+)"
            ArgvQuote = """" & Replace(.Replace(sArg, "$1$1$2$3"), """", "\""") & """"
        End With
    End If
End Function

Public Function GetErrorDescription(ByVal ErrorCode As Long) As String
    Dim lSize           As Long
    
    GetErrorDescription = Space$(2000)
    lSize = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0&, ErrorCode, 0, GetErrorDescription, Len(GetErrorDescription) + 1, 0)
    If lSize > 2 Then
        If Mid$(GetErrorDescription, lSize - 1, 2) = vbCrLf Then
            lSize = lSize - 2
        End If
    End If
    GetErrorDescription = Left$(GetErrorDescription, lSize)
End Function
