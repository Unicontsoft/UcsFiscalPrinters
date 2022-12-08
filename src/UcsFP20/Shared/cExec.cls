VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cExec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================================================
'
' UcsFP20 (c) 2008-2022 by Unicontsoft
'
' Unicontsoft Fiscal Printers Component 2.0
'
' This project is licensed under the terms of the MIT license
' See the LICENSE file in the project root for more information
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "cExec"

'=========================================================================
' API
'=========================================================================

'--- for CreateProcess
Private Const STARTF_USESHOWWINDOW          As Long = 1
Private Const STARTF_USESTDHANDLES          As Long = &H100
Private Const SW_HIDE                       As Long = 0
Private Const SW_MINIMIZE                   As Long = 6
Private Const NORMAL_PRIORITY_CLASS         As Long = &H20
'--- for WaitForSingleObject
Private Const INFINITE                      As Long = &HFFFFFFFF
'--- for DuplicateHandle
Private Const DUPLICATE_SAME_ACCESS         As Long = &H2
'--- for GetExitCodeProcess
Private Const STATUS_PENDING                As Long = &H103
'--- for SetInformationJobObject
Private Const JobObjectExtendedLimitInformation As Long = 9
Private Const JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE As Long = &H2000

Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessW" (ByVal lpApplicationName As Long, ByVal lpCommandLine As Long, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As SECURITY_ATTRIBUTES, ByVal nSize As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function DuplicateHandle Lib "kernel32" (ByVal hSourceProcessHandle As Long, ByVal hSourceHandle As Long, ByVal hTargetProcessHandle As Long, lpTargetHandle As Long, ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwOptions As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function FlushFileBuffers Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Function PeekNamedPipe Lib "kernel32" (ByVal hNamedPipe As Long, lpBuffer As Any, ByVal nBufferSize As Long, lpBytesRead As Long, lpTotalBytesAvail As Long, lpBytesLeftThisMessage As Long) As Long
Private Declare Function CreateJobObject Lib "kernel32" Alias "CreateJobObjectW" (ByVal lpJobAttributes As Long, ByVal lpName As Long) As Long
Private Declare Function SetInformationJobObject Lib "kernel32" (ByVal hJob As Long, ByVal JobObjectInformationClass As Long, lpJobObjectInformation As Any, ByVal cbJobObjectInformationLength As Long) As Long
Private Declare Function AssignProcessToJobObject Lib "kernel32" (ByVal hJob As Long, ByVal hProcess As Long) As Long
Private Declare Function OemToCharBuffA Lib "user32" (lpszSrc As Any, ByVal lpszDst As String, ByVal cchDstLength As Long) As Long

Private Type STARTUPINFO
    cb                  As Long
    lpReserved          As String
    lpDesktop           As String
    lpTitle             As String
    dwX                 As Long
    dwY                 As Long
    dwXSize             As Long
    dwYSize             As Long
    dwXCountChars       As Long
    dwYCountChars       As Long
    dwFillAttribute     As Long
    dwFlags             As Long
    wShowWindow         As Integer
    cbReserved2         As Integer
    lpReserved2         As Long
    hStdInput           As Long
    hStdOutput          As Long
    hStdError           As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess            As Long
    hThread             As Long
    dwProcessID         As Long
    dwThreadID          As Long
End Type

Private Type SECURITY_ATTRIBUTES
    nLength             As Long
    lpSecurityDescriptor As Long
    bInheritHandle      As Long
End Type

Private Type JOBOBJECT_EXTENDED_LIMIT_INFORMATION
    '--- BasicLimitInformation
    PerProcessUserTimeLimit As Currency
    PerJobUserTimeLimit     As Currency
    LimitFlags              As Long
    MinimumWorkingSetSize   As Long
    MaximumWorkingSetSize   As Long
    ActiveProcessLimit      As Long
    Affinity                As Long
    PriorityClass           As Long
    SchedulingClass         As Long
    dwPadding1              As Long
    '--- IoInfo
    ReadOperationCount      As Currency
    WriteOperationCount     As Currency
    OtherOperationCount     As Currency
    ReadTransferCount       As Currency
    WriteTransferCount      As Currency
    OtherTransferCount      As Currency
    '--- rest
    ProcessMemoryLimit      As Long
    JobMemoryLimit          As Long
    PeakProcessMemoryUsed   As Long
    PeakJobMemoryUsed       As Long
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Private m_hProcess              As Long
Private m_hReadOutput           As Long
Private m_hReadError            As Long
Private m_hWriteInput           As Long

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunction As String)
    #If USE_DEBUG_LOG <> 0 Then
        DebugLog MODULE_NAME, sFunction & "(" & Erl & ")", Err.Description & " &H" & Hex$(Err.Number), vbLogEventTypeError
    #Else
        Debug.Print "Critical error: " & Err.Description & " [" & MODULE_NAME & "." & sFunction & "]"
    #End If
End Sub

'=========================================================================
' Methods
'=========================================================================

Public Function Run( _
            sFile As String, _
            sParams As String, _
            Optional ByVal StartHidden As Boolean, _
            Optional ByVal StartMinimized As Boolean, _
            Optional ByVal LimitFlags As Long = JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE, _
            Optional CurrentDir As String) As Boolean
    Const FUNC_NAME     As String = "Run"
    Dim uProcInfo       As PROCESS_INFORMATION
    Dim uStart          As STARTUPINFO
    Dim sCommandLine    As String
    Dim uAttr           As SECURITY_ATTRIBUTES
    Dim hTmp            As Long
    Dim hWriteOutput    As Long
    Dim hWriteError     As Long
    Dim hReadInput      As Long
    Dim hJob            As Long
    Dim uInfo           As JOBOBJECT_EXTENDED_LIMIT_INFORMATION
    
    On Error GoTo EH
    '--- cleanup previous
    If m_hProcess <> 0 Then
        Call CloseHandle(m_hProcess)
        m_hProcess = 0
    End If
    If m_hReadOutput <> 0 Then
        Call CloseHandle(m_hReadOutput)
        m_hReadOutput = 0
    End If
    If m_hReadError <> 0 Then
        Call CloseHandle(m_hReadError)
        m_hReadError = 0
    End If
    '--- argv quote executable if not already quoted
    If LenB(sFile) <> 0 And Left$(sFile, 1) <> """" Then
        sCommandLine = ArgvQuote(sFile)
    Else
        sCommandLine = sFile
    End If
    If LenB(sParams) <> 0 Then
        sCommandLine = IIf(LenB(sCommandLine) <> 0, sCommandLine & " ", vbNullString) & sParams
    End If
    '--- create pipes
    uAttr.nLength = LenB(uAttr)
    uAttr.bInheritHandle = 1
    Call CreatePipe(hTmp, hWriteOutput, uAttr, 0)
    Call DuplicateHandle(GetCurrentProcess(), hTmp, GetCurrentProcess(), m_hReadOutput, 0, False, DUPLICATE_SAME_ACCESS)
    Call CloseHandle(hTmp)
    Call CreatePipe(hTmp, hWriteError, uAttr, 0)
    Call DuplicateHandle(GetCurrentProcess(), hTmp, GetCurrentProcess(), m_hReadError, 0, False, DUPLICATE_SAME_ACCESS)
    Call CloseHandle(hTmp)
    Call CreatePipe(hReadInput, hTmp, uAttr, 0)
    Call DuplicateHandle(GetCurrentProcess(), hTmp, GetCurrentProcess(), m_hWriteInput, 0, False, DUPLICATE_SAME_ACCESS)
    Call CloseHandle(hTmp)
    '--- setup start info
    uStart.cb = LenB(uStart)
    uStart.dwFlags = STARTF_USESTDHANDLES
    uStart.hStdInput = hReadInput
    uStart.hStdOutput = hWriteOutput
    uStart.hStdError = hWriteError
    If StartHidden Then
        uStart.dwFlags = uStart.dwFlags Or STARTF_USESHOWWINDOW
        uStart.wShowWindow = SW_HIDE
    End If
    If StartMinimized Then
        uStart.dwFlags = uStart.dwFlags Or STARTF_USESHOWWINDOW
        uStart.wShowWindow = SW_MINIMIZE
    End If
    '-- init job object
    If LimitFlags <> 0 Then
        hJob = CreateJobObject(0, 0)
        uInfo.LimitFlags = LimitFlags
        Call SetInformationJobObject(hJob, JobObjectExtendedLimitInformation, uInfo, LenB(uInfo))
    End If
    If LenB(CurrentDir) = 0 Then
        CurrentDir = vbNullString
    End If
    If CreateProcess(0, StrPtr(sCommandLine), 0, 0, 1, NORMAL_PRIORITY_CLASS, 0, StrPtr(CurrentDir), uStart, uProcInfo) <> 0 Then
        Call CloseHandle(uProcInfo.hThread)
        m_hProcess = uProcInfo.hProcess
        If hJob <> 0 Then
            Call AssignProcessToJobObject(hJob, m_hProcess)
        End If
    Else
        Call CloseHandle(m_hReadOutput)
        m_hReadOutput = 0
        Call CloseHandle(m_hReadError)
        m_hReadError = 0
    End If
    Call CloseHandle(hReadInput)
    Call CloseHandle(hWriteOutput)
    Call CloseHandle(hWriteError)
    '--- success (or failure)
    Run = (m_hProcess <> 0)
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Public Sub Detach()
    If m_hProcess <> 0 Then
        Call CloseHandle(m_hProcess)
        m_hProcess = 0
    End If
    If m_hReadOutput <> 0 Then
        Call CloseHandle(m_hReadOutput)
        m_hReadOutput = 0
    End If
    If m_hReadError <> 0 Then
        Call CloseHandle(m_hReadError)
        m_hReadError = 0
    End If
    If m_hWriteInput <> 0 Then
        Call CloseHandle(m_hWriteInput)
        m_hWriteInput = 0
    End If
End Sub

Public Function AtEndOfError() As Boolean
    Dim lTotal2         As Long
    
    If m_hReadError <> 0 Then
        If PeekNamedPipe(m_hReadError, ByVal 0, 0, 0, lTotal2, 0) = 0 Then
            Call CloseHandle(m_hReadError)
            m_hReadError = 0
        End If
    End If
    AtEndOfError = (m_hReadError = 0)
End Function

Public Function ReadError(ByVal lSize As Long) As String
    Dim baBuffer()      As Byte
    Dim lRead           As Long
    
    If m_hReadError <> 0 Then
        ReDim baBuffer(0 To lSize) As Byte
        If ReadFile(m_hReadError, baBuffer(0), lSize, lRead, 0) <> 0 And lRead > 0 Then
            ReadError = String$(lRead, 0)
            Call OemToCharBuffA(baBuffer(0), ReadError, lRead + 1)
        Else
            Call CloseHandle(m_hReadError)
            m_hReadError = 0
        End If
    End If
End Function

Public Function ReadLineError(Optional Terminator As String) As String
    Do While Not AtEndOfError
        ReadLineError = ReadLineError & ReadError(1)
        If Right$(ReadLineError, 2) = vbCrLf Then
            Exit Function
        ElseIf LenB(Terminator) <> 0 And Right$(ReadLineError, Len(Terminator)) = Terminator Then
            Exit Function
        End If
    Loop
End Function

Public Function ReadAllError() As String
    Do While Not AtEndOfError
        ReadAllError = ReadAllError & ReadError(100)
    Loop
End Function

Public Function ReadPendingError() As String
    Dim lTotal          As Long
    
    Do While Not AtEndOfError
        lTotal = 0
        If PeekNamedPipe(m_hReadError, ByVal 0, 0, 0, lTotal, 0) = 0 Then
            Call CloseHandle(m_hReadError)
            m_hReadError = 0
        End If
        If lTotal > 0 Then
            ReadPendingError = ReadPendingError & ReadError(lTotal)
        Else
            Exit Function
        End If
    Loop
End Function

Public Function AtEndOfOutput() As Boolean
    Dim lTotal2         As Long

    If m_hReadOutput <> 0 Then
        If PeekNamedPipe(m_hReadOutput, ByVal 0, 0, 0, lTotal2, 0) = 0 Then
            Call CloseHandle(m_hReadOutput)
            m_hReadOutput = 0
        End If
    End If
    AtEndOfOutput = (m_hReadOutput = 0)
End Function

Public Function ReadOutput(ByVal lSize As Long) As String
    Dim baBuffer()      As Byte
    Dim lRead           As Long
    
    If m_hReadOutput <> 0 Then
        ReDim baBuffer(0 To lSize) As Byte
        If ReadFile(m_hReadOutput, baBuffer(0), lSize, lRead, 0) <> 0 And lRead > 0 Then
            ReadOutput = String$(lRead, 0)
            Call OemToCharBuffA(baBuffer(0), ReadOutput, lRead + 1)
        Else
            Call CloseHandle(m_hReadOutput)
            m_hReadOutput = 0
        End If
    End If
End Function

Public Function ReadLineOutput() As String
    Do While Not AtEndOfOutput
        ReadLineOutput = ReadLineOutput & ReadOutput(1)
        If Right$(ReadLineOutput, 2) = vbCrLf Then
            Exit Function
        End If
    Loop
End Function

Public Function ReadAllOutput() As String
    Do While Not AtEndOfOutput
        ReadAllOutput = ReadAllOutput & ReadOutput(100)
    Loop
End Function

Public Function ReadPendingOutput() As String
    Dim lTotal          As Long
    
    Do While Not AtEndOfOutput
        lTotal = 0
        If PeekNamedPipe(m_hReadOutput, ByVal 0, 0, 0, lTotal, 0) = 0 Then
            Call CloseHandle(m_hReadOutput)
            m_hReadOutput = 0
        End If
        If lTotal > 0 Then
            ReadPendingOutput = ReadPendingOutput & ReadOutput(lTotal)
        Else
            Exit Function
        End If
    Loop
End Function

Public Function WriteInput(sValue As String) As Boolean
    Dim lWritten        As Long
    
    If m_hWriteInput <> 0 Then
        If WriteFile(m_hWriteInput, ByVal sValue, Len(sValue), lWritten, 0) <> 0 Then
            Call FlushFileBuffers(m_hWriteInput)
            WriteInput = True
        End If
    End If
End Function

Public Function GetExitCode() As Long
    If m_hProcess <> 0 Then
        Call GetExitCodeProcess(m_hProcess, GetExitCode)
        If GetExitCode = STATUS_PENDING Then
            If m_hReadOutput <> 0 Then
                Call CloseHandle(m_hReadOutput)
                m_hReadOutput = 0
            End If
            If m_hReadError <> 0 Then
                Call CloseHandle(m_hReadError)
                m_hReadError = 0
            End If
            If m_hWriteInput <> 0 Then
                Call CloseHandle(m_hWriteInput)
                m_hWriteInput = 0
            End If
            Call WaitForSingleObject(m_hProcess, INFINITE)
            Call GetExitCodeProcess(m_hProcess, GetExitCode)
        End If
        Call CloseHandle(m_hProcess)
        m_hProcess = 0
    End If
End Function

Public Function KillProcess() As Boolean
    If Not AtEndOfOutput And Not AtEndOfError Then
        If TerminateProcess(m_hProcess, 0) <> 0 Then
            KillProcess = True
        End If
    End If
End Function

'= private ===============================================================

' based on https://blogs.msdn.microsoft.com/twistylittlepassagesallalike/2011/04/23/everyone-quotes-command-line-arguments-the-wrong-way
Public Function ArgvQuote(sArg As String, Optional ByVal Force As Boolean) As String
    Const WHITESPACE As String = "*[ " & vbTab & vbVerticalTab & vbCrLf & "]*"
    
    If Not Force And LenB(sArg) <> 0 And Not sArg Like WHITESPACE Then
        ArgvQuote = sArg
    Else
        With VBA.CreateObject("VBScript.RegExp")
            .Global = True
            .Pattern = "(\\+)($|"")|(\\+)"
            ArgvQuote = """" & Replace(.Replace(sArg, "$1$1$2$3"), """", "\""") & """"
        End With
    End If
End Function

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Terminate()
    If m_hProcess <> 0 Then
        Call TerminateProcess(m_hProcess, 0)
    End If
    Detach
End Sub
