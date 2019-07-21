Attribute VB_Name = "mdConsole"
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

Private Const STD_INPUT_HANDLE              As Long = -10&
Private Const STD_OUTPUT_HANDLE             As Long = -11&
Private Const STD_ERROR_HANDLE              As Long = -12&
Private Const INPUT_RECORD_SIZE             As Long = 20
Private Const KEY_EVENT                     As Long = 1

Private Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Private Declare Function CharToOemBuff Lib "user32" Alias "CharToOemBuffA" (ByVal lpszSrc As String, lpszDst As Any, ByVal cchDstLength As Long) As Long
Private Declare Function OemToCharBuff Lib "user32" Alias "OemToCharBuffA" (lpszSrc As Any, ByVal lpszDst As String, ByVal cchDstLength As Long) As Long
Private Declare Function SetConsoleTextAttribute Lib "kernel32" (ByVal hConsoleOutput As Long, ByVal wAttributes As Long) As Long
Private Declare Function GetConsoleScreenBufferInfo Lib "kernel32" (ByVal hConsoleOutput As Long, lpConsoleScreenBufferInfo As CONSOLE_SCREEN_BUFFER_INFO) As Long
Private Declare Function PeekNamedPipe Lib "kernel32" (ByVal hNamedPipe As Long, lpBuffer As Any, ByVal nBufferSize As Long, lpBytesRead As Long, lpTotalBytesAvail As Long, lpBytesLeftThisMessage As Long) As Long
Private Declare Function GetNumberOfConsoleInputEvents Lib "kernel32" (ByVal hConsoleInput As Long, lpNumberOfEvents As Long) As Long
Private Declare Function PeekConsoleInput Lib "kernel32" Alias "PeekConsoleInputW" (ByVal hConsoleInput As Long, Buffer As Any, ByVal Length As Long, ByRef NumberOfEventsRead As Long) As Long
Private Declare Function FlushConsoleInputBuffer Lib "kernel32" (ByVal hConsoleInput As Long) As Long

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
' Constants and member variables
'=========================================================================

Public Const FOREGROUND_GREEN               As Long = &H2
Public Const FOREGROUND_RED                 As Long = &H4 + 8
Public Const FOREGROUND_MASK                As Long = &HF

Private m_sLastConsoleOutput            As String

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
        m_sLastConsoleOutput = pvConsoleOutput
        Debug.Print pvConsoleOutput;
    Else
        ReDim baBuffer(0 To Len(pvConsoleOutput) - 1) As Byte
        If CharToOemBuff(pvConsoleOutput, baBuffer(0), UBound(baBuffer) + 1) Then
            Call WriteFile(hOut, baBuffer(0), UBound(baBuffer) + 1, dwDummy, ByVal 0&)
        End If
    End If
End Function

Public Function ConsoleRead(Optional ByVal lSize As Long = 1) As String
    Dim hIn             As Long
    Dim baBuffer()      As Byte
    Dim sText           As String
    Dim lTotal          As Long
    Dim lIdx            As Long
    
    hIn = GetStdHandle(STD_INPUT_HANDLE)
    If hIn = 0 Then
        sText = InputBox(m_sLastConsoleOutput, "Console")
        If StrPtr(sText) = 0 Then
            End
        End If
        sText = sText & vbLf
    Else
        If PeekNamedPipe(hIn, ByVal 0, 0, 0, lTotal, 0) <> 0 Then
            If lTotal < lSize Then
                Exit Function
            End If
        ElseIf GetNumberOfConsoleInputEvents(hIn, lTotal) <> 0 Then
            If lTotal > 0 Then
                ReDim baBuffer(0 To lTotal * INPUT_RECORD_SIZE - 1) As Byte
                If PeekConsoleInput(hIn, baBuffer(0), UBound(baBuffer) + 1, lIdx) <> 0 Then
                    lTotal = 0
                    For lIdx = 0 To lIdx - 1
                        If baBuffer(lIdx * INPUT_RECORD_SIZE) = KEY_EVENT Then
                            lTotal = lTotal + 1
                        End If
                    Next
                End If
                If lTotal = 0 Then
                    Call FlushConsoleInputBuffer(hIn)
                End If
            End If
            If lTotal < lSize Then
                Exit Function
            End If
        End If
        ReDim baBuffer(0 To lSize - 1) As Byte
        If ReadFile(hIn, baBuffer(0), UBound(baBuffer) + 1, lSize, ByVal 0) <> 0 And lSize > 0 Then
            sText = String$(lSize, 0)
            Call OemToCharBuff(baBuffer(0), sText, lSize + 1)
        End If
    End If
    ConsoleRead = sText
End Function

Public Function ConsoleReadLine() As String
    Dim sChar           As String
    Dim sText           As String
    
    Do
        sChar = ConsoleRead
        Do While LenB(sChar) <> 0
            If Left$(sChar, 1) = vbLf Then
                ConsoleReadLine = sText
                Exit Function
            ElseIf Left$(sChar, 1) <> vbCr Then
                sText = sText & Left$(sChar, 1)
            End If
            sChar = Mid$(sChar, 2)
        Loop
    Loop
End Function
