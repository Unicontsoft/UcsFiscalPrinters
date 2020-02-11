VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFileLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================================================
'
' UcsFP20 (c) 2008-2020 by Unicontsoft
'
' Unicontsoft Fiscal Printers Component 2.0
'
' This project is licensed under the terms of the MIT license
' See the LICENSE file in the project root for more information
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "cFileLogger"

'=========================================================================
' API
'=========================================================================

Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const MAX_MEMORY_SIZE       As Long = 1& * 1024 * 1024
Private Const MAX_FILE_SIZE         As Long = 10& * 1024 * 1024
Private Const FORMAT_DATETIME_LOG   As String = "yyyy.MM.dd hh:nn:ss"
Private Const FORMAT_BASE_3         As String = "0.000"

Private m_nLogFile              As Integer
Private m_sLogFileName          As String
Private m_eLogLevel             As LogEventTypeConstants
Private m_bIsService            As Boolean
Private m_cMemoryLog            As Collection
Private m_lMemorySize           As Long
Private m_lMemoryCount          As Long

'=========================================================================
' Properties
'=========================================================================

Property Get LogFile() As Integer
    LogFile = m_nLogFile
End Property

Property Get LogFileName() As String
    LogFileName = m_sLogFileName
End Property

Property Let LogFileName(sValue As String)
    Flush
    m_sLogFileName = sValue
    m_nLogFile = (LenB(sValue) = 0)
End Property

Property Get LogLevel() As LogEventTypeConstants
    LogLevel = m_eLogLevel
End Property

Property Let LogLevel(ByVal lValue As LogEventTypeConstants)
    m_eLogLevel = lValue
End Property

Property Get IsService() As Boolean
    IsService = m_bIsService
End Property

Property Let IsService(ByVal bValue As Boolean)
    m_bIsService = bValue
End Property

Property Get MemoryLog() As Object
    Set MemoryLog = m_cMemoryLog
End Property

Property Get MemoryCount() As Long
    MemoryCount = m_lMemoryCount
End Property

'=========================================================================
' Methods
'=========================================================================

Public Sub Log(ByVal eType As LogEventTypeConstants, sModule As String, sFunction As String, sText As String)
    Const FUNC_NAME     As String = "Log"
    Dim vErr            As Variant
    Dim sPrefix         As String
    
    If eType > m_eLogLevel Then
        Exit Sub
    End If
    vErr = Array(Err.Number, Err.Description, Err.Source)
    On Error GoTo EH
    If eType = vbLogEventTypeError Then
        sPrefix = "[ERROR] "
    ElseIf eType = vbLogEventTypeWarning Then
        sPrefix = "[WARN ] "
    ElseIf eType = vbLogEventTypeInformation Then
        sPrefix = "[INFO ] "
    ElseIf eType = vbLogEventTypeDebug Then
        sPrefix = "[DEBUG] "
    End If
    If pvOutputLog(sModule, sFunction, sPrefix & sText) Then
        GoTo QH
    End If
    If m_bIsService Then
        App.LogEvent sText, LimitLong(eType, 0, vbLogEventTypeInformation)
    End If
QH:
    Err.Number = vErr(0)
    Err.Description = vErr(1)
    Err.Source = vErr(2)
    Exit Sub
EH:
    Debug.Print "Critical error: " & Err.Description & " [" & MODULE_NAME & "." & FUNC_NAME & "(" & Erl & ")]"
    Resume QH
End Sub

Public Sub DataDump(sModule As String, sFunction As String, ByVal sPrefix As String, sData As String)
    Const FUNC_NAME     As String = "DataDump"
    Dim vErr            As Variant
    Dim baData()        As Byte
    Dim lIdx            As Long
    Dim sText           As String
    Dim sHext           As String
    Dim lCount          As Long
    
    If m_eLogLevel < vbLogEventTypeDataDump Then
        Exit Sub
    End If
    If m_nLogFile = -1 Then
        Exit Sub
    End If
    vErr = Array(Err.Number, Err.Description, Err.Source)
    On Error GoTo EH
    baData = StrConv(sData, vbFromUnicode)
    lCount = ((UBound(baData) + 16) \ 16) * 16
    For lIdx = 0 To lCount
        If lIdx Mod 16 = 0 And LenB(sHext) <> 0 Then
            pvOutputLog IIf(lIdx = lCount, sModule, vbNullString), sFunction, sPrefix & Right$("0000" & Hex$(lIdx - 16), 4) & ": " & sHext & " " & sText
            sHext = vbNullString
            sText = vbNullString
            sPrefix = Space$(Len(sPrefix))
        End If
        If lIdx <= UBound(baData) Then
            sHext = sHext & Right$("0" & Hex$(baData(lIdx)), 2) & " "
            sText = sText & IIf(baData(lIdx) >= 32, Chr$(baData(lIdx)), ".")
        Else
            sHext = sHext & "   "
        End If
    Next
QH:
    On Error GoTo 0
    Err.Number = vErr(0)
    Err.Description = vErr(1)
    Err.Source = vErr(2)
    Exit Sub
EH:
    Debug.Print "Critical error: " & Err.Description & " [" & MODULE_NAME & "." & FUNC_NAME & "(" & Erl & ")]"
    Resume QH
End Sub

Public Sub Flush()
    Const FUNC_NAME     As String = "Flush"
    Dim vErr            As Variant
    
    vErr = Array(Err.Number, Err.Description, Err.Source)
    On Error GoTo EH
    If m_nLogFile <> 0 And m_nLogFile <> -1 Then
        Close #m_nLogFile
    End If
    m_nLogFile = 0
QH:
    On Error GoTo 0
    Err.Number = vErr(0)
    Err.Description = vErr(1)
    Err.Source = vErr(2)
    Exit Sub
EH:
    Debug.Print "Critical error: " & Err.Description & " [" & MODULE_NAME & "." & FUNC_NAME & "(" & Erl & ")]"
    Resume QH
End Sub

'= private ===============================================================

Private Function pvOutputLog(sModule As String, sFunction As String, sText As String) As Boolean
    Const FUNC_NAME     As String = "pvOutputLog"
    Dim sFile           As String
    Dim sNewFile        As String
    Dim nLogFile        As Integer
    Dim sPrefix         As String
    Dim sSuffix         As String
    Dim lIdx            As Long
    
    sPrefix = GetCurrentProcessId() & ": " & GetCurrentThreadId() & ": " & "(" & Format$(GetCurrentNow, FORMAT_DATETIME_LOG) & Right$(Format$(GetCurrentTimer, FORMAT_BASE_3), 4) & "): "
    sSuffix = IIf(LenB(sModule) <> 0, " [" & sModule & "." & sFunction & "]", vbNullString)
    m_lMemoryCount = m_lMemoryCount + 1
    m_lMemorySize = m_lMemorySize + Len(sPrefix & sText & sSuffix) + 2
    m_cMemoryLog.Add sPrefix & sText & sSuffix, "#" & m_lMemoryCount
    Do While m_lMemorySize >= MAX_MEMORY_SIZE
        m_lMemorySize = m_lMemorySize - Len(m_cMemoryLog.Item(1)) - 2
        m_cMemoryLog.Remove 1
    Loop
    If m_nLogFile = -1 Or LenB(m_sLogFileName) = 0 Then
        Exit Function
    End If
    On Error GoTo EH
    If m_nLogFile = 0 Then
        sFile = m_sLogFileName
        If FileExists(sFile) Then
            If FileLen(sFile) > MAX_FILE_SIZE Then
                If InStrRev(sFile, ".") > InStrRev(sFile, "\") Then
                    sNewFile = Left$(sFile, InStrRev(sFile, ".") - 1) & Format$(Date, "_yyyy_mm_dd") & Mid$(sFile, InStrRev(sFile, "."))
                Else
                    sNewFile = sFile & Format$(Date, "_yyyy_mm_dd")
                End If
                If FileExists(sNewFile) Then
                    For lIdx = 2 To 100
                        If Not FileExists(sNewFile & "." & lIdx) Then
                            sNewFile = sNewFile & "." & lIdx
                            Exit For
                        End If
                    Next
                End If
                Name sFile As sNewFile
            End If
        End If
        nLogFile = FreeFile
        m_nLogFile = -1
        Open sFile For Append Access Write Shared As #nLogFile
        m_nLogFile = nLogFile
    End If
    Print #m_nLogFile, sPrefix & sText & sSuffix
    If LOF(m_nLogFile) > MAX_FILE_SIZE Then
        Close #m_nLogFile
        m_nLogFile = 0
    End If
    '--- success
    pvOutputLog = True
QH:
    Exit Function
EH:
    Debug.Print "Critical error: " & Err.Description & " [" & MODULE_NAME & "." & FUNC_NAME & "(" & Erl & ")]"
    Resume QH
End Function

Private Function FileExists(sFile As String) As Boolean
    If GetFileAttributes(sFile) = -1 Then ' INVALID_FILE_ATTRIBUTES
    Else
        FileExists = True
    End If
End Function

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Initialize()
    m_nLogFile = -1
    Set m_cMemoryLog = New Collection
End Sub

Private Sub Class_Terminate()
    Flush
End Sub
