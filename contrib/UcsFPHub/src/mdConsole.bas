Attribute VB_Name = "mdConsole"
'=========================================================================
'
' VbPeg (c) 2018 by wqweto@gmail.com
'
' PEG parser generator for VB6
'
' mdConsole.bas - Console I/O functions
'
'=========================================================================
Option Explicit
DefObj A-Z

'=========================================================================
' API
'=========================================================================

Private Const STD_OUTPUT_HANDLE             As Long = -11&
Private Const STD_ERROR_HANDLE              As Long = -12&

Private Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Private Declare Function CharToOemBuff Lib "user32" Alias "CharToOemBuffA" (ByVal lpszSrc As String, lpszDst As Any, ByVal cchDstLength As Long) As Long
Private Declare Function SetConsoleTextAttribute Lib "kernel32" (ByVal hConsoleOutput As Long, ByVal wAttributes As Long) As Long
Private Declare Function GetConsoleScreenBufferInfo Lib "kernel32" (ByVal hConsoleOutput As Long, lpConsoleScreenBufferInfo As CONSOLE_SCREEN_BUFFER_INFO) As Long

Private Type COORD
    X                   As Integer
    Y                   As Integer
End Type

Private Type SMALL_RECT
    Left                As Integer
    Top                 As Integer
    Right               As Integer
    Bottom              As Integer
End Type

Private Type CONSOLE_SCREEN_BUFFER_INFO
    dwSize              As COORD
    dwCursorPosition    As COORD
    wAttributes         As Integer
    srWindow            As SMALL_RECT
    dwMaximumWindowSize As COORD
End Type

'=========================================================================
' Functions
'=========================================================================

Public Function ConsolePrint(ByVal sText As String, ParamArray A() As Variant) As String
    ConsolePrint = pvConsoleOutput(GetStdHandle(STD_OUTPUT_HANDLE), sText, CVar(A))
End Function

Public Function ConsoleError(ByVal sText As String, ParamArray A() As Variant) As String
    ConsoleError = pvConsoleOutput(GetStdHandle(STD_ERROR_HANDLE), sText, CVar(A))
End Function

Public Function ConsoleColorPrint(ByVal wAttr As Long, ByVal wMask As Long, ByVal sText As String, ParamArray A() As Variant) As String
    Dim hConsole        As Long
    Dim uInfo           As CONSOLE_SCREEN_BUFFER_INFO
    
    hConsole = GetStdHandle(STD_OUTPUT_HANDLE)
    Call GetConsoleScreenBufferInfo(hConsole, uInfo)
    Call SetConsoleTextAttribute(hConsole, (uInfo.wAttributes And Not wMask) Or (wAttr And wMask))
    ConsoleColorPrint = pvConsoleOutput(hConsole, sText, CVar(A))
    Call SetConsoleTextAttribute(hConsole, uInfo.wAttributes)
End Function

Public Function ConsoleColorError(ByVal wAttr As Long, ByVal wMask As Long, ByVal sText As String, ParamArray A() As Variant) As String
    Dim hConsole        As Long
    Dim uInfo           As CONSOLE_SCREEN_BUFFER_INFO
    
    hConsole = GetStdHandle(STD_ERROR_HANDLE)
    Call GetConsoleScreenBufferInfo(hConsole, uInfo)
    Call SetConsoleTextAttribute(hConsole, (uInfo.wAttributes And Not wMask) Or (wAttr And wMask))
    ConsoleColorError = pvConsoleOutput(hConsole, sText, CVar(A))
    Call SetConsoleTextAttribute(hConsole, uInfo.wAttributes)
End Function

Private Function pvConsoleOutput(ByVal hOut As Long, ByVal sText As String, A As Variant) As String
    Const LNG_PRIVATE   As Long = &HE1B6 '-- U+E000 to U+F8FF - Private Use Area (PUA)
    Dim lIdx            As Long
    Dim sArg            As String
    Dim baBuffer()      As Byte
    Dim dwDummy         As Long

    If LenB(sText) = 0 Then
        Exit Function
    End If
    '--- format
    For lIdx = UBound(A) To LBound(A) Step -1
        sArg = Replace(A(lIdx), "%", ChrW$(LNG_PRIVATE))
        sText = Replace(sText, "%" & (lIdx - LBound(A) + 1), sArg)
    Next
    pvConsoleOutput = Replace(sText, ChrW$(LNG_PRIVATE), "%")
    '--- output
    If hOut = 0 Then
        Debug.Print pvConsoleOutput;
    Else
        ReDim baBuffer(0 To Len(pvConsoleOutput) - 1) As Byte
        If CharToOemBuff(pvConsoleOutput, baBuffer(0), UBound(baBuffer) + 1) Then
            Call WriteFile(hOut, baBuffer(0), UBound(baBuffer) + 1, dwDummy, ByVal 0&)
        End If
    End If
End Function

