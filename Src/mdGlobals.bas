Attribute VB_Name = "mdGlobals"
'=========================================================================
' $Header: /UcsFiscalPrinter/Src/mdGlobals.bas 1     14.02.11 18:13 Wqw $
'
'   Unicontsoft Fiscal Printers Project
'   Copyright (c) 2008-2011 Unicontsoft
'
'   Globalni funktsii, constanti i promenliwi
'
' $Log: /UcsFiscalPrinter/Src/mdGlobals.bas $
' 
' 1     14.02.11 18:13 Wqw
' Initial implementation
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "mdGlobals"

'=========================================================================
' API
'=========================================================================

'--- for CreateFile
Private Const GENERIC_READ                  As Long = &H80000000
Private Const GENERIC_WRITE                 As Long = &H40000000
Private Const OPEN_EXISTING                 As Long = 3
Private Const INVALID_HANDLE_VALUE          As Long = -1
'--- for FormatMessage
Private Const FORMAT_MESSAGE_FROM_SYSTEM    As Long = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200
'--- for GetVersionEx
Private Const VER_PLATFORM_WIN32_NT         As Long = 2
'--- error codes
Private Const ERROR_ACCESS_DENIED           As Long = 5&
Private Const ERROR_GEN_FAILURE             As Long = 31&
Private Const ERROR_SHARING_VIOLATION       As Long = 32&
Private Const ERROR_SEM_TIMEOUT             As Long = 121&
'--- for GetLocaleInfo
Private Const LOCALE_SDECIMAL               As Long = &HE   ' decimal separator

Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Args As Any) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function QueryDosDevice Lib "kernel32" Alias "QueryDosDeviceA" (ByVal lpDeviceName As Long, ByVal lpTargetPath As String, ByVal ucchMax As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Declare Function GetUserDefaultLCID Lib "kernel32" () As Long

Private Type OSVERSIONINFO
    dwOSVersionInfoSize         As Long
    dwMajorVersion              As Long
    dwMinorVersion              As Long
    dwBuildNumber               As Long
    dwPlatformID                As Long
    szCSDVersion                As String * 128      '  Maintenance string for PSS usage
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Public Const STR_NONE               As String = "(Няма)"

Public g_sDecimalSeparator      As String

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunc As String)
    Debug.Print MODULE_NAME & "." & sFunc & ": " & Error
    OutputDebugLog MODULE_NAME, sFunc, "Run-time error: " & Error
End Sub

'=========================================================================
' Functions
'=========================================================================

Public Sub Main()
    g_sDecimalSeparator = GetDecimalSeparator()
End Sub

Public Function At(vData As Variant, ByVal lIdx As Long, Optional sDefault As String) As String
    On Error Resume Next
    At = sDefault
    At = C_Str(vData(lIdx))
    On Error GoTo 0
End Function

Public Function C_Lng(v As Variant) As Long
    On Error Resume Next
    C_Lng = CLng(v)
    On Error GoTo 0
End Function

Public Function C_Str(v As Variant) As String
    On Error Resume Next
    C_Str = CStr(v)
    On Error GoTo 0
End Function

Public Function C_Bool(v As Variant) As Boolean
    On Error Resume Next
    C_Bool = CBool(v)
    On Error GoTo 0
End Function

'Public Function C_Dbl(v As Variant) As Double
'    On Error Resume Next
'    C_Dbl = CDbl(v)
'    On Error GoTo 0
'End Function

Public Function C_Dbl(v As Variant) As Double
    On Error Resume Next
    C_Dbl = CDbl(Replace(C_Str(v), ".", g_sDecimalSeparator))
    On Error GoTo 0
End Function

Public Function C_Date(v As Variant) As Boolean
    On Error Resume Next
    C_Date = CDate(v)
    On Error GoTo 0
End Function

Public Function Zn(sText As String, Optional IfEmptyString As Variant = Null) As Variant
    Zn = IIf(LenB(sText) = 0, IfEmptyString, sText)
End Function

Public Function GetApiErr(ByVal lLastDllError As Long) As String
    Dim lRet            As Long
   
    GetApiErr = Space$(2000)
    lRet = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0&, lLastDllError, 0&, GetApiErr, Len(GetApiErr), 0&)
    If lRet > 2 Then
        If Mid$(GetApiErr, lRet - 1, 2) = vbCrLf Then
            lRet = lRet - 2
        End If
    End If
    GetApiErr = Left$(GetApiErr, lRet)
End Function

Public Function IsNT() As Boolean
    Const FUNC_NAME     As String = "IsNT"
    Dim udtVer          As OSVERSIONINFO
    
    On Error GoTo EH
    udtVer.dwOSVersionInfoSize = Len(udtVer)
    If GetVersionEx(udtVer) Then
        If udtVer.dwPlatformID = VER_PLATFORM_WIN32_NT Then
            IsNT = True
        End If
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Public Function EnumSerialPorts() As Variant
    Const FUNC_NAME     As String = "EnumSerialPorts"
    Dim sBuffer         As String
    Dim lIdx            As Long
    Dim hFile           As Long
    Dim vRet            As Variant
    Dim lCount          As Long
    
    On Error GoTo EH
    ReDim vRet(0 To 255) As Variant
    If IsNT Then
        sBuffer = String(100000, 1)
        Call QueryDosDevice(0, sBuffer, Len(sBuffer))
        sBuffer = Chr$(0) & sBuffer
        For lIdx = 1 To 255
            If InStr(1, sBuffer, Chr$(0) & "COM" & lIdx & Chr$(0), vbTextCompare) > 0 Then
                vRet(lCount) = "COM" & lIdx
                lCount = lCount + 1
            End If
        Next
    Else
        For lIdx = 1 To 255
            hFile = CreateFile("COM" & lIdx, GENERIC_READ Or GENERIC_WRITE, 0, 0, OPEN_EXISTING, 0, 0)
            If hFile = INVALID_HANDLE_VALUE Then
                Select Case Err.LastDllError
                Case ERROR_ACCESS_DENIED, ERROR_GEN_FAILURE, ERROR_SHARING_VIOLATION, ERROR_SEM_TIMEOUT
                    hFile = 0
                End Select
            Else
                Call CloseHandle(hFile)
                hFile = 0
            End If
            If hFile = 0 Then
                vRet(lCount) = "COM" & lIdx
                lCount = lCount + 1
            End If
        Next
    End If
    If lCount = 0 Then
        EnumSerialPorts = Split("")
    Else
        ReDim Preserve vRet(0 To lCount - 1) As Variant
        EnumSerialPorts = vRet
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Public Sub OutputDebugLog(sModule As String, sFunc As String, sText As String)
    Dim sFile           As String
    Dim nFile           As Integer
    
    On Error Resume Next
    sFile = Environ$("_UCS_FISCAL_PRINTER_LOG")
    If LenB(sFile) = 0 Then
        sFile = Environ$("TEMP") & "\UcsFP.log"
        If GetAttr(sFile) = -1 Then
            Exit Sub
        End If
    End If
    nFile = FreeFile
    Open sFile For Append Access Write As #nFile
    Print #nFile, sModule & "." & sFunc & "(" & Now & "): " & sText
    Close #nFile
    On Error GoTo 0
End Sub

Public Function Round(ByVal Value As Double, Optional ByVal NumDigits As Long) As Double
    On Error Resume Next
    Round = VBA.Round(Value + IIf(Value > 0, 10 ^ -13, -10 ^ -13), NumDigits)
    On Error GoTo 0
End Function

Public Function Ceil(ByVal Value As Double) As Long
    Ceil = -Int(-Value)
End Function

Public Function Floor(ByVal Value As Double) As Long
    Floor = Int(Value)
End Function

Public Function GetDecimalSeparator() As String
    Dim sBuffer         As String
    Dim nSize           As Long

    sBuffer = Space$(100)
    nSize = GetLocaleInfo(GetUserDefaultLCID(), LOCALE_SDECIMAL, sBuffer, Len(sBuffer))
    If nSize > 0 Then
        GetDecimalSeparator = Left$(sBuffer, nSize - 1)
    Else
        GetDecimalSeparator = "."
    End If
End Function

Public Function IsDelimiter(sText As String) As Boolean
    Const STR_DELIMS As String = "~#$^&*_+-=\|/ " & vbTab & vbCrLf
    If InStr(1, STR_DELIMS, Left(sText, 1)) > 0 Then
        IsDelimiter = True
    End If
End Function

Public Function IsWhitespace(sText As String) As Boolean
    Const STR_WHITESPACE As String = " " & vbTab & vbCrLf
    If InStr(1, STR_WHITESPACE, Left(sText, 1)) > 0 Then
        IsWhitespace = True
    End If
End Function

Public Function WrapText(ByVal sText As String, ByVal lWidth As Long) As Variant
    Dim lRight          As Long
    Dim lLeft           As Long
    Dim vRet            As Variant
    Dim lCount          As Long
    
    ReDim vRet(0 To Len(sText)) As Variant
    Do While LenB(sText) <> 0
        lRight = lWidth + 1
        If lRight > Len(sText) Then
            lRight = Len(sText) + 1
        Else
            If IsDelimiter(Mid(sText, lRight, 1)) Then
                Do While IsWhitespace(Mid(sText, lRight, 1)) And lRight <= Len(sText)
                    lRight = lRight + 1
                Loop
            Else
                Do While lRight > 1
                    If IsDelimiter(Mid(sText, lRight - 1, 1)) Then
                        Exit Do
                    End If
                    lRight = lRight - 1
                Loop
                If lRight = 1 Then
                    lRight = lWidth + 1
                End If
            End If
        End If
        lLeft = lRight - 1
        Do While IsWhitespace(Mid(sText, lLeft, 1)) And lLeft > 0
            lLeft = lLeft - 1
        Loop
        vRet(lCount) = Left(sText, lLeft)
        lCount = lCount + 1
        sText = Mid(sText, lRight)
    Loop
    If lCount = 0 Then
        WrapText = Array("")
    Else
        ReDim Preserve vRet(0 To lCount - 1) As Variant
        WrapText = vRet
    End If
End Function

Public Function Printf(ByVal sText As String, ParamArray A() As Variant) As String
    Const LNG_PRIVATE   As Long = &HE1B6 '-- U+E000 to U+F8FF - Private Use Area (PUA)
    Dim lIdx            As Long
    
    For lIdx = UBound(A) To LBound(A) Step -1
        sText = Replace(sText, "%" & (lIdx - LBound(A) + 1), Replace(A(lIdx), "%", ChrW$(LNG_PRIVATE)))
    Next
    Printf = Replace(sText, ChrW$(LNG_PRIVATE), "%")
End Function

Public Function LimitLong( _
            ByVal lValue As Long, _
            Optional ByVal lMin As Long = -2147483647, _
            Optional ByVal lMax As Long = 2147483647) As Long
    If lValue < lMin Then
        LimitLong = lMin
    ElseIf lValue > lMax Then
        LimitLong = lMax
    Else
        LimitLong = lValue
    End If
End Function

