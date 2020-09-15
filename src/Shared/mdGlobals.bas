Attribute VB_Name = "mdGlobals"
'=========================================================================
'
' UcsFP20 (c) 2008-2019 by Unicontsoft
'
' Unicontsoft Fiscal Printers Component 2.0
'
' This project is licensed under the terms of the MIT license
' See the LICENSE file in the project root for more information
'
'=========================================================================
'
' Global functions, constants and variables
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "mdGlobals"

'=========================================================================
' Public Enums
'=========================================================================

Public Enum UcsRegistryRootsEnum
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_LOCAL_MACHINE = &H80000002
End Enum

Public Type UcsParsedUrl
    Protocol        As String
    Host            As String
    Port            As Long
    Path            As String
    User            As String
    Pass            As String
End Type

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
'--- error codes
Private Const ERROR_ACCESS_DENIED           As Long = 5&
Private Const ERROR_GEN_FAILURE             As Long = 31&
Private Const ERROR_SHARING_VIOLATION       As Long = 32&
Private Const ERROR_SEM_TIMEOUT             As Long = 121&
'--- for GetLocaleInfo
Private Const LOCALE_USER_DEFAULT           As Long = &H400
Private Const LOCALE_SDECIMAL               As Long = &HE   ' decimal separator
'--- windows messages
Private Const WM_PRINTCLIENT                As Long = &H318
Private Const WM_MOUSELEAVE                 As Long = &H2A3
'--- registry
Private Const REG_SZ                        As Long = 1
Private Const REG_EXPAND_SZ                 As Long = 2
Private Const REG_DWORD                     As Long = 4
'--- for GetOpenFileNameA
Private Const OFN_HIDEREADONLY              As Long = &H4&
Private Const OFN_EXTENSIONDIFFERENT        As Long = &H400
Private Const OFN_CREATEPROMPT              As Long = &H2000&
Private Const OFN_EXPLORER                  As Long = &H80000
Private Const OFN_LONGNAMES                 As Long = &H200000
Private Const OFN_ENABLESIZING              As Long = &H800000
'--- for VariantChangeType
Private Const VARIANT_ALPHABOOL             As Long = 2

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal ByteLength As Long)
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Args As Any) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function GetVersion Lib "kernel32" () As Long
Private Declare Function QueryDosDevice Lib "kernel32" Alias "QueryDosDeviceA" (ByVal lpDeviceName As Long, ByVal lpTargetPath As String, ByVal ucchMax As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Declare Function DllGetVersion Lib "comctl32" (pdvi As DLLVERSIONINFO) As Long
Private Declare Function SetWindowSubclass Lib "comctl32" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
Private Declare Function DefSubclassProc Lib "comctl32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function APIGetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long
Private Declare Function GetOpenFileName Lib "comdlg32" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function ApiEmptyDoubleArray Lib "oleaut32" Alias "SafeArrayCreateVector" (Optional ByVal vt As VbVarType = vbDouble, Optional ByVal lLow As Long = 0, Optional ByVal lCount As Long = 0) As Double()
Private Declare Function IsTextUnicode Lib "advapi32" (lpBuffer As Any, ByVal cb As Long, lpi As Long) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function VariantChangeType Lib "oleaut32" (vDest As Variant, Src As Variant, ByVal wFlags As Integer, ByVal vt As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Type OPENFILENAME
    lStructSize         As Long     ' size of type/structure
    hWndOwner           As Long     ' Handle of owner window
    hInstance           As Long
    lpstrFilter         As Long     ' Filters used to select files
    lpstrCustomFilter   As Long
    nMaxCustomFilter    As Long
    nFilterIndex        As Long     ' index of Filter to start with
    lpstrFile           As Long     ' Holds filepath and name
    nMaxFile            As Long     ' Maximum Filepath and name length
    lpstrFileTitle      As Long     ' Filename
    nMaxFileTitle       As Long     ' Max Length of filename
    lpstrInitialDir     As Long     ' Starting Directory
    lpstrTitle          As Long     ' Title of window
    Flags               As Long
    nFileOffset         As Integer
    nFileExtension      As Integer
    lpstrDefExt         As Long
    lCustData           As Long
    lpfnHook            As Long
    lpTemplateName      As Long
    pvReserved          As Long
    dwReserved          As Long
    FlagsEx             As Long
End Type

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion      As Long
    dwMinorVersion      As Long
    dwBuildNumber       As Long
    dwPlatformID        As Long
    szCSDVersion        As String * 128      '  Maintenance string for PSS usage
End Type

Private Type DLLVERSIONINFO
    cbSize              As Long
    dwMajor             As Long
    dwMinor             As Long
    dwBuildNumber       As Long
    dwPlatformID        As Long
End Type

Private Type VBGUID
    Data1               As Long
    Data2               As Integer
    Data3               As Integer
    Data4(7)            As Byte
End Type

Private Type DISPPARAMS
    rgPointerToVariantArray As Long
    rgPointerToLongNamedArgs As Long
    cArgs               As Long
    cNamedArgs          As Long
End Type

Private Type EXCEPINFO
    wCode               As Integer
    wReserved           As Integer
    Source              As String
    Description         As String
    HelpFile            As String
    dwHelpContext       As Long
    pvReserved          As Long
    pfnDeferredFillIn   As Long
    sCode               As Long
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Public Const LIB_NAME                   As String = "UcsFP20"
Public Const STR_NONE                   As String = "(Ќ€ма)"
Public Const STR_PROTOCOL_DATECS_X      As String = "DATECS/X"
Public Const STR_PROTOCOL_DATECS_FP     As String = "DATECS"
Public Const STR_PROTOCOL_DAISY         As String = "DAISY"
Public Const STR_PROTOCOL_INCOTEX       As String = "INCOTEX"
Public Const STR_PROTOCOL_ELTRADE       As String = "ELTRADE"
Public Const STR_PROTOCOL_TREMOL        As String = "TREMOL"
Public Const STR_PROTOCOL_ESCPOS        As String = "ESC/POS"
Public Const STR_PROTOCOL_PROXY         As String = "PROXY"
Public Const STR_CHR1                   As String = "" '--- CHAR(1)
Public Const DBL_EPSILON                As Double = 0.0000000001
Public Const FORMAT_BASE_2              As String = "0.00"
Public Const FORMAT_BASE_3              As String = "0.000"
Public Const STR_CONNECTOR_ERRORS       As String = "No device info set|%1 failed: %2|Timeout waiting for response"
Public Const vbLogEventTypeDebug        As Long = vbLogEventTypeInformation + 1
Public Const vbLogEventTypeDataDump     As Long = vbLogEventTypeInformation + 2
Public Const STR_ENUM_STATUS_CODE       As String = "ready|busy|failed"

Private m_sDecimalSeparator         As String
Private m_oConfig                   As Object
Private m_oProtocolConfig           As Object
Private m_oPortWrapper              As cPortWrapper
Private m_oLogger                   As cFileLogger
Private m_dCurrentStartDate         As Date
Private m_dblCurrentStartTimer      As Double
Private m_cRegExpCache              As Collection

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunction As String)
    Debug.Print "Critical error: " & Err.Description & " [" & MODULE_NAME & "." & sFunction & "]"
    Logger.Log vbLogEventTypeError, MODULE_NAME, sFunction & "(" & Erl & ")", Err.Description
End Sub

Private Sub RaiseError(sFunction As String)
    Debug.Print "Critical error: " & Err.Description & " [" & MODULE_NAME & "." & sFunction & "]"
    Logger.Log vbLogEventTypeError, MODULE_NAME, sFunction & "(" & Erl & ")", Err.Description
    Err.Raise Err.Number, MODULE_NAME & "." & sFunction & "(" & Erl & ")" & vbCrLf & Err.Source, Err.Description
End Sub

'=========================================================================
' Properties
'=========================================================================

Property Get DecimalSeparator() As String
    DecimalSeparator = m_sDecimalSeparator
End Property

Property Get PortWrapper() As cPortWrapper
    Set PortWrapper = m_oPortWrapper
End Property

'=========================================================================
' Functions
'=========================================================================

Private Sub Main()
    Const FUNC_NAME     As String = "Main"
    Dim sFile           As String
    Dim vJson           As Variant
    Dim sError          As String
    
    On Error GoTo EH
    m_sDecimalSeparator = GetDecimalSeparator()
    sFile = LocateFile(App.Path & "\" & App.EXEName & ".conf")
    If LenB(sFile) <> 0 Then
        Logger.Log vbLogEventTypeDebug, MODULE_NAME, FUNC_NAME, "Loading config file " & sFile
        If Not JsonParse(ReadTextFile(sFile), vJson, Error:=sError) Then
            Logger.Log vbLogEventTypeError, MODULE_NAME, FUNC_NAME, "Error in config: " & sError
        End If
    End If
    Set m_oConfig = C_Obj(vJson)
    Set m_oPortWrapper = New cPortWrapper
    SetCurrentDateTimer VBA.Now, TimerEx
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Public Property Get At(vData As Variant, ByVal lIdx As Long, Optional sDefault As String) As String
    On Error GoTo QH
    At = sDefault
    If IsArray(vData) Then
        If lIdx < LBound(vData) Then
            '--- lIdx = -1 for last element
            lIdx = UBound(vData) + 1 + lIdx
        End If
        If LBound(vData) <= lIdx And lIdx <= UBound(vData) Then
            At = C_Str(vData(lIdx))
        End If
    End If
QH:
End Property

Public Property Let ValueAt(vData As Variant, ByVal lIdx As Long, vValue As Variant)
    On Error GoTo QH
    If IsArray(vData) Then
        If LBound(vData) <= lIdx And lIdx <= UBound(vData) Then
            vData(lIdx) = vValue
        End If
    End If
QH:
End Property

Public Function C_Lng(Value As Variant) As Long
    Dim vDest           As Variant
    
    If VarType(Value) = vbLong Then
        C_Lng = Value
    ElseIf VariantChangeType(vDest, Value, 0, vbLong) = 0 Then
        C_Lng = vDest
    End If
End Function

Public Function C_Str(Value As Variant) As String
    Dim vDest           As Variant
    
    If VarType(Value) = vbString Then
        C_Str = Value
    ElseIf VariantChangeType(vDest, Value, VARIANT_ALPHABOOL, vbString) = 0 Then
        C_Str = vDest
    End If
End Function

Public Function C_Bool(Value As Variant) As Boolean
    Dim vDest           As Variant
    
    If VarType(Value) = vbBoolean Then
        C_Bool = Value
    ElseIf VariantChangeType(vDest, Value, VARIANT_ALPHABOOL, vbBoolean) = 0 Then
        C_Bool = vDest
    End If
End Function

Public Function C_Dbl(Value As Variant) As Double
    Dim vDest           As Variant
    
    If VarType(Value) = vbDouble Then
        C_Dbl = Value
    ElseIf VariantChangeType(vDest, Replace(C_Str(Value), ".", m_sDecimalSeparator), 0, vbDouble) = 0 Then
        C_Dbl = vDest
    End If
End Function

Public Function C_Date(Value As Variant) As Date
    Dim vDest           As Variant
    
    If VarType(Value) = vbDate Then
        C_Date = Value
    ElseIf VariantChangeType(vDest, Value, 0, vbDate) = 0 Then
        C_Date = vDest
    End If
End Function

Public Function C_Obj(Value As Variant) As Object
    Dim vDest       As Variant

    If VarType(Value) = vbObject Then
        Set C_Obj = Value
    ElseIf VariantChangeType(vDest, Value, 0, vbObject) = 0 Then
        Set C_Obj = vDest
    End If
End Function

Public Function C_Val(Text As String) As Double
    On Error GoTo QH
    C_Val = Val(Text)
QH:
End Function

Public Function Zn(sText As String, Optional IfEmptyString As Variant = Null) As Variant
    Zn = IIf(LenB(sText) = 0, IfEmptyString, sText)
End Function

Public Function Znl(ByVal lValue As Long, Optional IfEmptyLong As Variant = Null, Optional ByVal EmptyLong As Long = 0) As Variant
    Znl = IIf(lValue = EmptyLong, IfEmptyLong, lValue)
End Function

Public Function Zndbl(ByVal dblValue As Double, Optional IfZeroDouble As Variant = Null) As Variant
    Zndbl = IIf(C_Dbl(CStr(dblValue)) = 0, IfZeroDouble, dblValue)
End Function

Public Function GetErrorDescription(ByVal ErrorCode As Long) As String
    Dim lSize           As Long
   
    GetErrorDescription = Space$(2000)
    lSize = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0, ErrorCode, 0, GetErrorDescription, Len(GetErrorDescription), 0)
    If lSize > 2 Then
        If Mid$(GetErrorDescription, lSize - 1, 2) = vbCrLf Then
            lSize = lSize - 2
        End If
    End If
    GetErrorDescription = Left$(GetErrorDescription, lSize)
End Function

Public Property Get IsNT() As Boolean
    Static lVersion     As Long
    
    If lVersion = 0 Then
        lVersion = GetVersion()
    End If
    IsNT = ((lVersion And &H80000000) = 0)
End Property

Public Property Get OsVersion() As Long
    Static lVersion     As Long
    Dim uVer            As OSVERSIONINFO

    If lVersion = 0 Then
        uVer.dwOSVersionInfoSize = Len(uVer)
        If GetVersionEx(uVer) Then
            lVersion = uVer.dwMajorVersion * 100 + uVer.dwMinorVersion
        End If
    End If
    OsVersion = lVersion
End Property

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
        sBuffer = String$(100000, 1)
        Call QueryDosDevice(0, sBuffer, Len(sBuffer))
        sBuffer = vbNullChar & sBuffer
        For lIdx = 1 To 255
            If InStr(1, sBuffer, vbNullChar & "COM" & lIdx & vbNullChar, vbTextCompare) > 0 Then
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
        EnumSerialPorts = Array()
    Else
        ReDim Preserve vRet(0 To lCount - 1) As Variant
        EnumSerialPorts = vRet
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Property Get Logger() As cFileLogger
    Const FUNC_NAME     As String = "Logger [get]"
    Dim sFileName       As String
    Dim vErr            As Variant
    
    If m_oLogger Is Nothing Then
        vErr = Array(Err.Number, Err.Description, Err.Source)
        On Error GoTo EH
        sFileName = GetEnvironmentVar("_UCS_FISCAL_PRINTER_LOG")
        If LenB(sFileName) = 0 Then
            sFileName = GetErrorTempPath() & "\UcsFP.log"
            If Not FileExists(sFileName) Then
                sFileName = vbNullString
            End If
        End If
        Set m_oLogger = New cFileLogger
        m_oLogger.LogFileName = sFileName
        m_oLogger.LogLevel = IIf(Val(GetEnvironmentVar("_UCS_FISCAL_PRINTER_DATA_DUMP")) <> 0, vbLogEventTypeDataDump, vbLogEventTypeDebug)
        On Error GoTo 0
QH:
        Err.Number = vErr(0)
        Err.Description = vErr(1)
        Err.Source = vErr(2)
    End If
    Set Logger = m_oLogger
    Exit Property
EH:
    Debug.Print "Critical error: " & Err.Description & " [" & MODULE_NAME & "." & FUNC_NAME & "(" & Erl & ")]"
    Resume QH
End Property

Property Set Logger(oValue As cFileLogger)
    Set m_oLogger = oValue
End Property

Public Sub DebugLog(sModule As String, sFunction As String, sText As String, Optional ByVal eType As LogEventTypeConstants = vbLogEventTypeInformation)
    Logger.Log eType, sModule, sFunction, sText
End Sub

Public Sub FlushDebugLog()
    If Not m_oLogger Is Nothing Then
        m_oLogger.Flush
    End If
End Sub

Public Function Round(ByVal Value As Double, Optional ByVal NumDigits As Long) As Double
    Round = FormatNumber(Value, NumDigits, vbTrue, vbFalse, vbFalse)
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
    nSize = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SDECIMAL, sBuffer, Len(sBuffer))
    If nSize > 0 Then
        GetDecimalSeparator = Left$(sBuffer, nSize - 1)
    Else
        GetDecimalSeparator = "."
    End If
End Function

Public Function IsDelimiter(sText As String) As Boolean
    Const STR_DELIMS As String = "~#$^&*_+-=\|, " & vbTab & vbCrLf
    If InStr(1, STR_DELIMS, Left$(sText, 1)) > 0 Then
        IsDelimiter = True
    End If
End Function

Public Function IsWhiteSpace(sText As String) As Boolean
    Const STR_WHITESPACE As String = " " & vbTab & vbCrLf
    If InStr(1, STR_WHITESPACE, Left$(sText, 1)) > 0 Then
        IsWhiteSpace = True
    End If
End Function

Public Function WrapMultiline(ByVal sText As String, ByVal lWidth As Long) As Variant
    Dim vElem           As Variant
    Dim vRetVal         As Variant
    Dim lIdx            As Long
    
    For Each vElem In Split(sText, vbCrLf)
        vElem = WrapText(C_Str(vElem), lWidth)
        If Not IsArray(vRetVal) Then
            vRetVal = vElem
        Else
            ReDim Preserve vRetVal(0 To UBound(vRetVal) + UBound(vElem) + 1) As Variant
            For lIdx = 0 To UBound(vElem)
                vRetVal(UBound(vRetVal) - UBound(vElem) + lIdx) = vElem(lIdx)
            Next
        End If
    Next
    WrapMultiline = vRetVal
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
            If IsDelimiter(Mid$(sText, lRight, 1)) Then
                Do While IsWhiteSpace(Mid$(sText, lRight, 1)) And lRight <= Len(sText)
                    lRight = lRight + 1
                Loop
            Else
                Do While lRight > 1
                    If IsDelimiter(Mid$(sText, lRight - 1, 1)) Then
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
        Do While IsWhiteSpace(Mid$(sText, lLeft, 1)) And lLeft > 0
            lLeft = lLeft - 1
        Loop
        vRet(lCount) = Left$(sText, lLeft)
        lCount = lCount + 1
        sText = Mid$(sText, lRight)
    Loop
    If lCount = 0 Then
        WrapText = Array(vbNullString)
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

Public Function Clamp( _
            ByVal lValue As Long, _
            Optional ByVal lMin As Long = -2147483647, _
            Optional ByVal lMax As Long = 2147483647) As Long
    Select Case lValue
    Case lMin To lMax
        Clamp = lValue
    Case Is < lMin
        Clamp = lMin
    Case Is > lMax
        Clamp = lMax
    End Select
End Function

Public Function Limit( _
            ByVal Value As Double, _
            Optional Min As Variant, _
            Optional Max As Variant) As Double
    Limit = Value
    If Not IsMissing(Min) Then
        If Value < C_Dbl(Min) Then
            Limit = C_Dbl(Min)
        End If
    End If
    If Not IsMissing(Max) Then
        If Value > C_Dbl(Max) Then
            Limit = C_Dbl(Max)
        End If
    End If
End Function

Public Function SplitCgAddress( _
            ByVal sAddress As String, _
            sRow1 As String, _
            sRow2 As String, _
            ByVal lRowChars As Long) As String
    Dim vSplit          As Variant
    
    sAddress = Replace(sAddress, "є", "N")
    Do While Left$(sAddress, 2) = vbCrLf
        sAddress = LTrim$(Mid$(sAddress, 3))
    Loop
    Do While Right$(sAddress, 2) = vbCrLf
        sAddress = RTrim$(Left$(sAddress, Len(sAddress) - 2))
    Loop
    Do While InStr(sAddress, " " & vbCrLf) > 0
        sAddress = Replace(sAddress, " " & vbCrLf, vbCrLf)
    Loop
    sAddress = Replace(sAddress, vbCrLf, "; ")
    vSplit = WrapText(sAddress, lRowChars)
    sRow1 = Trim$(At(vSplit, 0))
    If Right$(sRow1, 1) = ";" Then
        sRow1 = Left$(sRow1, Len(sRow1) - 1)
    End If
    sRow2 = Trim$(At(vSplit, 1))
    If Right$(sRow2, 1) = ";" Then
        sRow2 = Left$(sRow2, Len(sRow2) - 1)
    End If
End Function

Public Function AlignText( _
            ByVal sLeft As String, _
            ByVal sRight As String, _
            ByVal lWidth As Long) As String
    sLeft = Left$(sLeft, lWidth)
    If Left$(sRight, 1) = STR_CHR1 Then
        sRight = String$(lWidth - Len(sLeft), Right$(sRight, 1))
    Else
        sRight = Right$(sRight, lWidth)
    End If
    AlignText = sLeft & Space$(lWidth - Len(sLeft))
    If LenB(sRight) <> 0 Then
        Mid$(AlignText, lWidth - Len(sRight) + 1, Len(sRight)) = sRight
    End If
End Function

Public Function CenterText(ByVal sText As String, ByVal lWidth As Long) As String
    sText = Left$(sText, lWidth)
    CenterText = Space$(Clamp((lWidth - Len(sText)) \ 2, 0)) & sText
End Function

Public Function SumArray(vArray As Variant) As Double
    Dim vElem           As Variant
    
    For Each vElem In vArray
        SumArray = SumArray + C_Dbl(vElem)
    Next
End Function

Public Function IsComCtl6Loaded() As Boolean
    Dim uVer            As DLLVERSIONINFO
    
    uVer.cbSize = Len(uVer)
    Call DllGetVersion(uVer)
    IsComCtl6Loaded = (uVer.dwMajor >= 6)
End Function

Public Function FixThemeSupport(oControls As Object) As Boolean
    Const FUNC_NAME     As String = "FixThemeSupport"
    Dim oCtl            As Object
    
    On Error GoTo EH
    If IsComCtl6Loaded() Then
        For Each oCtl In oControls
            If TypeOf oCtl Is VB.Frame Then
                SetWindowSubclass oCtl.hWnd, AddressOf pvRedirectFrame, 0, 0
            End If
        Next
        '--- success
        FixThemeSupport = True
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function pvRedirectFrame( _
            ByVal hWnd As Long, _
            ByVal wMsg As Long, _
            ByVal wParam As Long, _
            ByVal lParam As Long, _
            ByVal uIdSubclass As Long, _
            ByVal dwRefData As Long) As Long
    Const FUNC_NAME     As String = "pvRedirectFrame"
    
    On Error GoTo EH
    #If uIdSubclass And dwRefData Then '--- touch args
    #End If
    Select Case wMsg
    Case WM_PRINTCLIENT, WM_MOUSELEAVE
        pvRedirectFrame = DefWindowProc(hWnd, wMsg, wParam, lParam)
    Case Else
        pvRedirectFrame = DefSubclassProc(hWnd, wMsg, wParam, lParam)
    End Select
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Public Function RegReadString(ByVal hRoot As UcsRegistryRootsEnum, sKey As String, sValue As String) As String
    Dim hKey            As Long
    Dim lType           As Long
    Dim lNeeded         As Long
    Dim sBuffer         As String
    Dim dwBuffer        As Long
    
    If RegOpenKeyEx(hRoot, sKey, 0, &H20001, hKey) = 0 Then '--- &H20001 = READ_CONTROL Or KEY_QUERY_VALUE
        Call RegQueryValueEx(hKey, sValue, 0, lType, ByVal vbNullString, lNeeded)
        If lType = REG_SZ Or lType = REG_EXPAND_SZ Then
            sBuffer = String$(lNeeded + 1, 0)
            If RegQueryValueEx(hKey, sValue, 0, lType, ByVal sBuffer, Len(sBuffer)) = 0 Then
                sBuffer = Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1)
                If lType = REG_EXPAND_SZ Then
                    RegReadString = String$(ExpandEnvironmentStrings(sBuffer, vbNullString, 0), 0)
                    If ExpandEnvironmentStrings(sBuffer, RegReadString, Len(RegReadString)) > 0 Then
                        RegReadString = Left$(RegReadString, InStr(RegReadString, vbNullChar) - 1)
                    Else
                        RegReadString = sBuffer
                    End If
                Else
                    RegReadString = sBuffer
                End If
            End If
        ElseIf lType = REG_DWORD Then
            If RegQueryValueEx(hKey, sValue, 0, lType, dwBuffer, 4) = 0 Then
                RegReadString = dwBuffer
            End If
        End If
        Call RegCloseKey(hKey)
    End If
End Function

Public Sub RegWriteValue(ByVal hRoot As UcsRegistryRootsEnum, sKey As String, sValue As String, vValue As Variant)
    Dim hKey            As Long
    Dim lTemp           As Long
    Dim sTemp           As String
    
    If RegOpenKeyEx(hRoot, sKey, 0, &H20002, hKey) = 0 Then '--- &H20002 = READ_CONTROL Or KEY_SET_VALUE
        Select Case VarType(vValue)
        Case vbLong, vbInteger, vbByte
            lTemp = C_Lng(vValue)
            Call RegSetValueEx(hKey, sValue, 0, REG_DWORD, lTemp, 4)
        Case vbBoolean
            lTemp = -C_Lng(vValue)
            Call RegSetValueEx(hKey, sValue, 0, REG_DWORD, lTemp, 4)
        Case Else
            sTemp = C_Str(vValue)
            Call RegSetValueEx(hKey, sValue, 0, REG_SZ, ByVal sTemp, Len(sTemp))
        End Select
        Call RegCloseKey(hKey)
    End If
End Sub

Public Function GetSystemDirectory() As String
    GetSystemDirectory = String$(1000, 0)
    APIGetSystemDirectory GetSystemDirectory, Len(GetSystemDirectory) - 1
    GetSystemDirectory = Left$(GetSystemDirectory, InStr(GetSystemDirectory, vbNullChar) - 1)
End Function

Public Function OpenSaveDialog(ByVal hWndOwner As Long, ByVal sFilter As String, ByVal sTitle As String, sFile As String) As Boolean
    Const FUNC_NAME     As String = "OpenSaveDialog"
    Dim uOFN            As OPENFILENAME
    Dim sBuffer         As String
    Dim baFilter()      As Byte
    Dim baTitle()       As Byte
    
    On Error GoTo EH
    baFilter = ToAscii(Replace(sFilter, "|", vbNullChar))
    baTitle = ToAscii(sTitle)
    sBuffer = String$(1000, 0)
    If OsVersion >= 500 Then
        uOFN.lStructSize = Len(uOFN)
    Else
        uOFN.lStructSize = Len(uOFN) - 12
    End If
    uOFN.Flags = OFN_LONGNAMES Or OFN_CREATEPROMPT Or OFN_HIDEREADONLY Or OFN_EXTENSIONDIFFERENT Or OFN_EXPLORER Or OFN_ENABLESIZING
    uOFN.hWndOwner = hWndOwner
    uOFN.lpstrFilter = VarPtr(baFilter(0))
    uOFN.nFilterIndex = 1
    uOFN.lpstrTitle = VarPtr(baTitle(0))
    uOFN.lpstrFile = StrPtr(sBuffer)
    uOFN.nMaxFile = Len(sBuffer)
    If GetOpenFileName(uOFN) <> 0 Then
        sFile = StrConv(sBuffer, vbUnicode)
        sFile = Left$(sFile, InStr(sFile, vbNullChar) - 1)
        '--- success
        OpenSaveDialog = True
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Public Function Pad(ByVal sText As String, ByVal lSize As Long, Optional ByVal sFill As String) As String
    If LenB(sFill) = 0 Then
        sFill = IIf(lSize > 0, " ", "0")
    End If
    If lSize > 0 Then
        Pad = Left$(sText & String(lSize, sFill), lSize)
    Else
        Pad = Right$(String(-lSize, sFill) & sText, -lSize)
    End If
End Function

Public Function EmptyDoubleArray() As Double()
    EmptyDoubleArray = ApiEmptyDoubleArray()
End Function

Public Function FileExists(sFile As String) As Boolean
    If GetFileAttributes(sFile) = -1 Then ' INVALID_FILE_ATTRIBUTES
    Else
        FileExists = True
    End If
End Function

Public Function ReadTextFile(sFile As String) As String
    Const BOM_UTF       As String = "пїњ" '--- "\xEF\xBB\xBF"
    Const BOM_UNICODE   As String = "€ю"  '--- "\xFF\xFE"
    Const ForReading    As Long = 1
    Dim lSize           As Long
    Dim sPrefix         As String
    
    With CreateObject("Scripting.FileSystemObject")
        lSize = .GetFile(sFile).Size
        If lSize = 0 Then
            Exit Function
        End If
        sPrefix = .OpenTextFile(sFile, ForReading).Read(IIf(lSize < 50, lSize, 50))
        If Left$(sPrefix, Len(BOM_UTF)) <> BOM_UTF And Left$(sPrefix, Len(BOM_UNICODE)) <> BOM_UNICODE Then
            '--- special xml encoding test
            If InStr(1, sPrefix, "<?xml", vbTextCompare) > 0 And InStr(1, sPrefix, "utf-8", vbTextCompare) > 0 Then
                sPrefix = BOM_UTF
            End If
        End If
        If Left$(sPrefix, Len(BOM_UTF)) <> BOM_UTF Then
            On Error GoTo QH
            ReadTextFile = .OpenTextFile(sFile, ForReading, False, Left$(sPrefix, Len(BOM_UNICODE)) = BOM_UNICODE Or IsTextUnicode(ByVal sPrefix, Len(sPrefix), &HFFFF& - 2) <> 0).ReadAll()
        Else
            With CreateObject("ADODB.Stream")
                .Open
                If Left$(sPrefix, Len(BOM_UNICODE)) = BOM_UNICODE Then
                    .Charset = "Unicode"
                ElseIf Left$(sPrefix, Len(BOM_UTF)) = BOM_UTF Then
                    .Charset = "UTF-8"
                Else
                    .Charset = "_autodetect_all"
                End If
                .LoadFromFile sFile
                ReadTextFile = .ReadText
            End With
        End If
    End With
QH:
End Function

Public Function SetProtocolConfigRoot(oValue As Object)
    Set m_oProtocolConfig = oValue
End Function

Public Function GetConfigValue(sSerial As String, sKey As String, Optional vDefault As Variant) As Variant
    Const FUNC_NAME     As String = "GetConfigValue"
    Dim oItem           As Object
    
    On Error GoTo EH
    If LenB(sSerial) <> 0 Then
        Set oItem = C_Obj(JsonItem(m_oProtocolConfig, sSerial))
        If oItem Is Nothing Then
            Set oItem = C_Obj(JsonItem(m_oConfig, sSerial))
        End If
        If Not IsEmpty(JsonItem(oItem, sKey)) Then
            AssignVariant GetConfigValue, JsonItem(oItem, sKey)
            Exit Function
        End If
    End If
    If IsMissing(vDefault) Then
        Err.Raise vbObjectError, , "Missing default value for " & sKey
    End If
    AssignVariant GetConfigValue, vDefault
    Exit Function
EH:
    RaiseError FUNC_NAME & "(sSerial=" & sSerial & ", sKey=" & sKey & ")"
End Function

Public Function GetConfigNumber(sSerial As String, sKey As String, ByVal dblDefault As Double) As Double
    Const FUNC_NAME     As String = "GetConfigNumber"
    
    On Error GoTo EH
    GetConfigNumber = C_Dbl(GetConfigValue(sSerial, sKey, 0))
    If dblDefault > DBL_EPSILON Then
        If GetConfigNumber <= 0 Then
            GetConfigNumber = dblDefault
        End If
    ElseIf dblDefault < -DBL_EPSILON Then
        If GetConfigNumber >= 0 Then
            GetConfigNumber = dblDefault
        End If
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME & "(sSerial=" & sSerial & ", sKey=" & sKey & ")"
End Function

Public Function GetConfigCollection(sSerial As String, sKey As String) As Collection
    Const FUNC_NAME     As String = "GetConfigCollection"
    Dim vValue          As Variant
    Dim oDict           As Object
    
    On Error GoTo EH
    AssignVariant vValue, GetConfigValue(sSerial, sKey, Empty)
    If IsObject(vValue) Then
        Set oDict = JsonToDictionary(C_Obj(vValue))
        Set GetConfigCollection = New Collection
        pvAppendConfigCollection oDict, vbNullString, GetConfigCollection
    End If
    Exit Function
EH:
    RaiseError FUNC_NAME & "(sSerial=" & sSerial & ", sKey=" & sKey & ")"
End Function

Private Function pvAppendConfigCollection(oDict As Object, sPrefix As String, oCol As Collection)
    Dim vKey            As Variant
    
    For Each vKey In oDict
        If IsObject(oDict(vKey)) Then
            pvAppendConfigCollection oDict(vKey), sPrefix & "\" & vKey, oCol
        Else
            oCol.Add oDict(vKey), sPrefix & "\" & vKey
        End If
    Next
End Function

Public Function LocateFile(sFile As String) As String
    Dim sDir            As String
    Dim sName           As String
    Dim lPos            As Long
    
    If InStrRev(sFile, "\") > 0 Then
        sDir = Left$(sFile, InStrRev(sFile, "\"))
        sName = Mid$(sFile, InStrRev(sFile, "\") + 1)
        Do While Not FileExists(sDir & sName)
            If Len(sDir) > 1 Then
                lPos = InStrRev(sDir, "\", Len(sDir) - 1)
                If lPos > 0 Then
                    sDir = Left$(sDir, lPos)
                    If Left$(sDir, 2) = "\\" And InStrRev(sDir, "\", Len(sDir) - 1) <= 2 Then
                        Exit Function
                    End If
                Else
                    Exit Function
                End If
            Else
                Exit Function
            End If
        Loop
        LocateFile = sDir & sName
    ElseIf FileExists(sFile) Then
        LocateFile = sFile
    End If
End Function

Public Function SafeFormat(Expression As Variant, Optional Fmt As Variant, Optional sDecimal As String = ".") As String
    SafeFormat = Replace(Format$(Expression, Fmt), m_sDecimalSeparator, sDecimal)
End Function

Public Function SafeText(sText As String) As String
    Dim lIdx            As Long
    
    SafeText = sText
    For lIdx = 0 To 31
        SafeText = Replace(SafeText, Chr$(lIdx), vbNullString)
    Next
End Function

Public Sub AssignVariant(vDest As Variant, vSrc As Variant)
    If IsObject(vSrc) Then
        Set vDest = vSrc
    Else
        vDest = vSrc
    End If
End Sub

Public Function preg_match(find_re As String, sText As String, Optional Matches As Variant, Optional Indexes As Variant) As Long
    Dim lIdx            As Long
    
    With InitRegExp(find_re).Execute(sText)
        preg_match = .Count
        If Not IsMissing(Matches) Then
            If .Count = 0 Then
                Matches = Split(vbNullString)
            ElseIf .Count = 1 Then
                ReDim Matches(0 To 0) As String
                Matches(0) = .Item(0).Value
            Else
                ReDim Matches(0 To .Count - 1) As String
                For lIdx = 0 To .Count - 1
                    Matches(lIdx) = .Item(lIdx).Value
                Next
            End If
        End If
        If Not IsMissing(Indexes) Then
            If .Count = 0 Then
                Indexes = Array()
            ElseIf .Count = 1 Then
                Indexes = Array(.Item(0).FirstIndex + 1)
            Else
                ReDim Indexes(0 To .Count - 1) As Variant
                For lIdx = 0 To .Count - 1
                    Indexes(lIdx) = .Item(lIdx).FirstIndex + 1
                Next
            End If
        End If
    End With
End Function

Public Function preg_replace(find_re As String, sText As String, Optional sReplace As String) As String
    preg_replace = InitRegExp(find_re).Replace(sText, sReplace)
End Function

Public Function InitRegExp(sPattern As String) As Object
    Const FUNC_NAME     As String = "InitRegExp"
    Dim lPos            As Long
    
    On Error GoTo EH
    If Not SearchCollection(m_cRegExpCache, sPattern, RetVal:=InitRegExp) Then
        Set InitRegExp = CreateObject("VBScript.RegExp")
        With InitRegExp
            lPos = InStrRev(sPattern, "/")
            If Left$(sPattern, 1) = "/" And lPos > 1 Then
                .Pattern = Mid$(sPattern, 2, lPos - 2)
                .IgnoreCase = (InStr(lPos, sPattern, "i") > 0)
                .MultiLine = (InStr(lPos, sPattern, "m") > 0)
                .Global = (InStr(lPos, sPattern, "l") = 0)
            Else
                .Global = True
                .Pattern = sPattern
            End If
        End With
        If m_cRegExpCache Is Nothing Then
            Set m_cRegExpCache = New Collection
        End If
        m_cRegExpCache.Add InitRegExp, sPattern
        If m_cRegExpCache.Count > 1000 Then
            m_cRegExpCache.Remove 1
        End If
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Public Function GetConfigForCommand(oConfigCmd As Collection, oLocalizedCmd As Collection, sFunction As String, sKey As String, Optional Default As Variant) As Variant
    Dim sMerged         As String
    Dim vItem           As Variant
    
    sMerged = "\" & sFunction & IIf(LenB(sKey) <> 0, "\" & sKey, vbNullString)
    If Not SearchCollection(oConfigCmd, sMerged, vItem) Then
        If Not SearchCollection(oLocalizedCmd, sMerged, vItem) Then
            If Not IsMissing(Default) Then
                GetConfigForCommand = Default
            End If
            Exit Function
        End If
    End If
    Select Case VarType(Default)
    Case vbLong, vbInteger, vbByte
        GetConfigForCommand = C_Lng(vItem)
    Case vbDouble, vbSingle
        GetConfigForCommand = C_Dbl(vItem)
    Case vbString
        GetConfigForCommand = C_Str(vItem)
    Case vbBoolean
        GetConfigForCommand = C_Bool(vItem)
    End Select
End Function

Public Property Get TimerEx() As Double
    Dim cFreq           As Currency
    Dim cValue          As Currency
    
    Call QueryPerformanceFrequency(cFreq)
    Call QueryPerformanceCounter(cValue)
    TimerEx = cValue / cFreq
End Property

Public Function ToHexDump(sText As String) As String
    Dim lIdx            As Long
    
    For lIdx = 1 To Len(sText)
        ToHexDump = ToHexDump & Right$("0" & Hex$(Asc(Mid$(sText, lIdx, 1))), 2)
    Next
End Function

Public Function SearchCollection(ByVal pCol As Object, Index As Variant, Optional RetVal As Variant) As Boolean
    Const DISPID_VALUE  As Long = 0
    Const VT_BYREF      As Long = &H4000
    Const S_OK          As Long = 0
    Dim pVbCol          As IVbCollection
    Dim vItem           As Variant

    If pCol Is Nothing Then
        '--- do nothing
    ElseIf (PeekInt(VarPtr(RetVal)) And VT_BYREF) = 0 Then
        If TypeOf pCol Is IVbCollection Then
            Set pVbCol = pCol
            SearchCollection = (pVbCol.Item(Index, RetVal) = S_OK)
        Else
            SearchCollection = DispInvoke(pCol, DISPID_VALUE, VbMethod Or VbGet, RetVal:=RetVal, Args:=Index)
        End If
    Else
        If TypeOf pCol Is IVbCollection Then
            Set pVbCol = pCol
            SearchCollection = (pVbCol.Item(Index, vItem) = S_OK)
        Else
            SearchCollection = DispInvoke(pCol, DISPID_VALUE, VbMethod Or VbGet, RetVal:=vItem, Args:=Index)
        End If
        If SearchCollection Then
            If IsObject(vItem) Then
                Set RetVal = vItem
            Else
                RetVal = vItem
            End If
        End If
    End If
End Function

Public Function DispInvoke( _
            ByVal pDisp As IVbDispatch, _
            Name As Variant, _
            Optional ByVal CallType As VbCallType, _
            Optional Args As Variant, _
            Optional RetVal As Variant) As Boolean
    Const DISPID_PROPERTYPUT As Long = -3
    Const VT_BYREF      As Long = &H4000
    Dim IID_NULL        As VBGUID
    Dim lDispID         As Long
    Dim hResult         As Long
    Dim uParams         As DISPPARAMS
    Dim uInfo           As EXCEPINFO
    Dim aParams()       As Variant
    Dim lNamedParam     As Long
    Dim lIdx            As Long
    Dim lParamCount     As Long
    Dim lArgErr         As Long
    Dim lPtrResult      As Long
    Dim vRetVal         As Variant

    If pDisp Is Nothing Then
        Exit Function
    End If
    '--- get disp id
    If IsNumeric(Name) Then
        lDispID = C_Lng(Name)
    Else
        hResult = pDisp.GetIDsOfNames(IID_NULL, C_Str(Name), 1, LOCALE_USER_DEFAULT, lDispID)
        If hResult < 0 Then
            GoTo QH
        End If
    End If
    If CallType = 0 Then
        CallType = VbMethod Or IIf(Not IsMissing(RetVal), VbGet, 0)
    End If
    '--- process params
    If Not IsMissing(Args) Then
        If IsArray(Args) Then
            lParamCount = UBound(Args) - LBound(Args)
            ReDim aParams(0 To lParamCount) As Variant
            For lIdx = 0 To lParamCount
                Call AssignVariant(aParams(lParamCount - lIdx), Args(lIdx))
            Next
        Else
            ReDim aParams(0 To 0) As Variant
            Call AssignVariant(aParams(0), Args)
        End If
        With uParams
            .cArgs = lParamCount + 1
            .rgPointerToVariantArray = VarPtr(aParams(0))
        End With
        If (CallType And (VbLet Or VbSet)) <> 0 Then
            lNamedParam = DISPID_PROPERTYPUT
            With uParams
                .cNamedArgs = 1
                .rgPointerToLongNamedArgs = VarPtr(lNamedParam)
            End With
        End If
    End If
    If (CallType And VbGet) <> 0 Or (CallType And VbMethod) <> 0 And Not IsMissing(RetVal) Then
        lPtrResult = VarPtr(RetVal)
        If (PeekInt(lPtrResult) And VT_BYREF) = 0 Then
            If IsObject(RetVal) Then
                Set RetVal = Nothing
            Else
                RetVal = Empty
            End If
        Else
            lPtrResult = VarPtr(vRetVal)
            If IsObject(RetVal) Then
                Set vRetVal = Nothing
            Else
                vRetVal = Empty
            End If
        End If
    End If
    hResult = pDisp.Invoke(lDispID, IID_NULL, LOCALE_USER_DEFAULT, CallType, uParams, ByVal lPtrResult, uInfo, lArgErr)
    If hResult < 0 Then
        GoTo QH
    End If
    If lPtrResult = VarPtr(vRetVal) Then
        If IsObject(vRetVal) Then
            Set RetVal = vRetVal
        Else
            RetVal = vRetVal
        End If
    End If
    '--- success
    DispInvoke = True
    Exit Function
QH:
    If VarType(RetVal) = vbVariant Then
        RetVal = Array(hResult, uInfo.sCode, uInfo.Description, uInfo.Source)
    End If
End Function

Public Function DispPropertyGet(pDisp As Object, PropName As String, Optional RetVal As Variant) As Variant
    If DispInvoke(pDisp, PropName, VbMethod Or VbGet, RetVal:=RetVal) Then
        AssignVariant DispPropertyGet, RetVal
    End If
End Function

Public Property Get LockControl(oCtl As Object) As Boolean
    Dim vResult         As Variant
    
    If DispInvoke(oCtl, "Locked", RetVal:=vResult) Then
        LockControl = vResult
    ElseIf DispInvoke(oCtl, "Enabled", RetVal:=vResult) Then
        LockControl = Not vResult
    End If
End Property

Public Property Let LockControl(oCtl As Object, ByVal bValue As Boolean)
    If DispInvoke(oCtl, "Locked", VbLet, Args:=bValue) Or TypeOf oCtl Is ListBox Then
        DispInvoke oCtl, "BackColor", VbLet, Args:=IIf(bValue, vbButtonFace, vbWindowBackground)
    Else
        DispInvoke oCtl, "Enabled", VbLet, Args:=Not bValue
    End If
End Property

Public Function ParseSum(sValue As String) As Double
    If InStr(sValue, ".") > 0 Then
        ParseSum = C_Dbl(sValue)
    Else
        ParseSum = C_Dbl(sValue) / 100#
    End If
End Function

Public Function ToAscii(sSend As String, Optional ByVal CodePage As Long) As Byte()
    Dim lSize           As Long
    Dim baText()        As Byte
    
    lSize = Len(sSend)
    If lSize > 0 Then
        ReDim baText(0 To lSize - 1) As Byte
        Call WideCharToMultiByte(CodePage, 0, StrPtr(sSend), lSize, baText(0), Len(sSend), 0, 0)
    Else
        baText = " "
    End If
    ToAscii = baText
End Function

Public Function FromAscii(baRecv() As Byte, Optional ByVal CodePage As Long) As String
    Dim lSize           As Long
    
    If UBound(baRecv) >= 0 Then
        FromAscii = String$(2 * (UBound(baRecv) + 1), 0)
        lSize = MultiByteToWideChar(CodePage, 0, baRecv(0), UBound(baRecv) + 1, StrPtr(FromAscii), Len(FromAscii) + 1)
        If lSize <> Len(FromAscii) Then
            FromAscii = Left$(FromAscii, lSize)
        End If
    End If
End Function

Public Function SplitOrReindex(Expression As String, Delimiter As String) As Variant
    Dim vResult         As Variant
    Dim lIdx            As Long
    Dim lSize           As Long
    
    SplitOrReindex = Split(Expression, Delimiter)
    '--- check if reindex needed
    If IsNumeric(At(SplitOrReindex, 0)) Then
        For lIdx = 0 To UBound(SplitOrReindex) Step 2
            lSize = Clamp(lSize, C_Lng(At(SplitOrReindex, lIdx)))
        Next
        ReDim vResult(0 To lSize) As Variant
        For lIdx = 0 To UBound(SplitOrReindex) Step 2
            vResult(C_Lng(At(SplitOrReindex, lIdx))) = At(SplitOrReindex, lIdx + 1)
        Next
        SplitOrReindex = vResult
    End If
End Function

Public Function InitDeviceConnector( _
            sDevice As String, _
            ByVal lTimeout As Long, _
            Optional LocalizedConnectorErrors As String, _
            Optional Error As String) As IDeviceConnector
    Dim oSerialPortConn As cSerialPortConnector
    Dim oSocketConn     As cSocketConnector
    
    If LenB(sDevice) = 0 Then
        Error = At(Split(LocalizedConnectorErrors, "|"), ucsErrNoDeviceInfoSet, "No device info set")
        GoTo QH
    End If
    If LCase$(Left$(sDevice, 3)) = "com" Then
        Set oSerialPortConn = New cSerialPortConnector
        If LenB(LocalizedConnectorErrors) <> 0 Then
            oSerialPortConn.LocalizedText(ucsFscLciInternalErrors) = LocalizedConnectorErrors
        End If
        If Not oSerialPortConn.Init(sDevice, lTimeout) Then
            Error = oSerialPortConn.GetLastError()
            GoTo QH
        End If
        Set InitDeviceConnector = oSerialPortConn
    Else
        Set oSocketConn = New cSocketConnector
        If Not oSocketConn.Init(sDevice, lTimeout) Then
            Error = oSocketConn.GetErrorDescription(oSocketConn.LastError)
            GoTo QH
        End If
        Set InitDeviceConnector = oSocketConn
    End If
QH:
End Function

Public Function Peek(ByVal lPtr As Long) As Long
    Call CopyMemory(Peek, ByVal lPtr, 4)
End Function

Public Function PeekInt(ByVal lPtr As Long) As Integer
    Call CopyMemory(PeekInt, ByVal lPtr, 2)
End Function

Public Function GetErrorTempPath() As String
    Dim sBuffer         As String
    
    sBuffer = String$(2000, 0)
    Call GetTempPath(Len(sBuffer) - 1, sBuffer)
    GetErrorTempPath = Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    If Right$(GetErrorTempPath, 1) = "\" Then
        GetErrorTempPath = Left$(GetErrorTempPath, Len(GetErrorTempPath) - 1)
    End If
End Function

Private Function pvParseTokenByRegExp(sText As String, sPattern As String) As String
    Dim oCol            As Object
    
    Set oCol = InitRegExp(sPattern).Execute(sText)
    If oCol.Count > 0 Then
        pvParseTokenByRegExp = oCol.Item(0).SubMatches(0)
        sText = Mid$(sText, oCol.Item(0).FirstIndex + oCol.Item(0).Length + 1)
    End If
End Function

Public Function ParseDeviceString(ByVal sDeviceString As String, Optional Separator As String = ";") As Object
    Const KEY_PATTERN   As String = "^([^=]+)="
    Const VALUE_PATTERN As String = "^\s*('(?:''|[^'])*'|""(?:""""|[^""])*""|[^;]*)\s*;?"
    Const ASC_QUOTE     As Long = 39
    Const ASC_DBL_QUOTE As Long = 34
    Dim sKey            As String
    Dim sValue          As String
    Dim oRetVal         As Object
    Dim lChar           As Long
    
    Do
        sKey = Trim$(pvParseTokenByRegExp(sDeviceString, KEY_PATTERN))
        If LenB(sKey) = 0 Then
            Exit Do
        End If
        sValue = Trim$(pvParseTokenByRegExp(sDeviceString, Replace(VALUE_PATTERN, ";", Separator)))
        If Len(sValue) >= 2 Then
            If Left$(sValue, 1) = Right$(sValue, 1) Then
                lChar = Asc(sValue)
                Select Case lChar
                Case ASC_QUOTE, ASC_DBL_QUOTE
                    sValue = Replace(Mid$(sValue, 2, Len(sValue) - 2), Chr$(lChar) & Chr$(lChar), Chr$(lChar))
                End Select
            End If
        End If
        JsonItem(oRetVal, sKey) = sValue
    Loop
    Set ParseDeviceString = oRetVal
End Function

Public Function ToDeviceString(oMap As Object, Optional Separator As String = ";") As String
    Dim vKey            As Variant
    Dim sValue          As String
    Dim sRetVal         As String
    
    For Each vKey In JsonKeys(oMap)
        '--- try to escape value
        sValue = C_Str(JsonItem(oMap, vKey))
        Select Case True
        Case InStr(sValue, Separator) > 0, InStr(sValue, """") > 0, InStr(sValue, "'") > 0
            sValue = """" & Replace(sValue, """", """""") & """"
        End Select
        sRetVal = IIf(LenB(sRetVal) <> 0, sRetVal & Separator, vbNullString) & vKey & "=" & sValue
    Next
    ToDeviceString = sRetVal
End Function

Public Function GetEnvironmentVar(sName As String) As String
    Dim sBuffer         As String
    
    sBuffer = String$(2000, 0)
    Call GetEnvironmentVariable(sName, sBuffer, Len(sBuffer) - 1)
    GetEnvironmentVar = Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1)
End Function

Public Function StripZeros(ByVal sText As String) As String
    Dim lIdx            As Long
    
    sText = Trim$(sText)
    For lIdx = 1 To Len(sText) - 1
        If Mid$(sText, lIdx, 1) <> "0" Then
            Exit For
        End If
    Next
    StripZeros = Trim$(Mid$(sText, lIdx))
End Function

Public Function JsonEnumValue(vValue As Variant, sList As String) As Long
    Dim sText           As String
    Dim vElem           As Variant
    Dim lIdx            As Long
    
    If VarType(vValue) = vbString Then
        sText = LCase$(vValue)
        If LenB(sText) <> 0 Then
            If InStr(1, "|" & sList & "|", "|" & sText & "|", vbTextCompare) > 0 Then
                For Each vElem In Split(LCase$(sList), "|")
                    If sText = vElem Then
                        JsonEnumValue = lIdx
                        Exit Function
                    End If
                    lIdx = lIdx + 1
                Next
            End If
        End If
    End If
    JsonEnumValue = C_Lng(vValue)
End Function

Public Function JsonBoolItem(oJson As Object, sKey As String, Optional ByVal Default As Boolean) As Boolean
    Dim vValue          As Variant
    
    AssignVariant vValue, JsonItem(oJson, sKey)
    If VarType(vValue) = vbString Then
        Select Case LCase$(vValue)
        Case "y", "yes", "true", "on", "д", "да"
            JsonBoolItem = True
        Case "n", "no", "false", "off", "н", "не"
            JsonBoolItem = False
        Case Else
            If IsNumeric(vValue) Then
                JsonBoolItem = C_Bool(vValue)
            Else
                JsonBoolItem = Default
            End If
        End Select
    ElseIf IsEmpty(vValue) Then
        JsonBoolItem = Default
    ElseIf IsNumeric(vValue) Then
        JsonBoolItem = C_Bool(vValue)
    Else
        JsonBoolItem = Default
    End If
End Function

Public Function ToConnectorDevice( _
            oOptions As Object, _
            ByVal DefSocketPort As Long, _
            oProtocol As IDeviceProtocol) As String
    Const DEF_SERIAL_PORT As String = "COM1"
    Const DEF_SERIAL_SPEED As Long = 115200
    Dim vPorts          As Variant
    Dim vElem           As Variant
    
    If LenB(JsonItem(oOptions, "IP")) <> 0 Then
        ToConnectorDevice = Trim$(JsonItem(oOptions, "IP")) & _
            ":" & Znl(C_Lng(JsonItem(oOptions, "Port")), DefSocketPort)
    Else
        If Not oProtocol Is Nothing Then
            If LenB(JsonItem(oOptions, "Port")) = 0 Then
                vPorts = EnumSerialPorts
            ElseIf LenB(JsonItem(oOptions, "Speed")) = 0 Then
                vPorts = Array(JsonItem(oOptions, "Port"))
            End If
            If IsArray(vPorts) Then
                vPorts = oProtocol.AutodetectDevices(vPorts)
                For Each vElem In vPorts
                    If IsArray(vElem) Then
                        If JsonItem(oOptions, "Protocol") = Zn(Trim$(At(vElem, 2)), Empty) Then
                            If LenB(JsonItem(oOptions, "Port")) = 0 Then
                                JsonItem(oOptions, "Port") = Zn(Trim$(At(vElem, 0)), Empty)
                            End If
                            If LenB(JsonItem(oOptions, "Speed")) = 0 Then
                                JsonItem(oOptions, "Speed") = Zn(Trim$(At(vElem, 1)), Empty)
                            End If
                            JsonItem(oOptions, "DeviceSerialNo") = Zn(Trim$(At(vElem, 5)), Empty)
                            JsonItem(oOptions, "FiscalMemoryNo") = Zn(Trim$(At(vElem, 6)), Empty)
                            Exit For
                        End If
                    End If
                Next
                If LenB(JsonItem(oOptions, "Speed")) = 0 Then
                    GoTo QH
                End If
            End If
        End If
        ToConnectorDevice = Zn(Trim$(JsonItem(oOptions, "Port")), DEF_SERIAL_PORT) & _
            "," & Znl(C_Lng(JsonItem(oOptions, "Speed")), DEF_SERIAL_SPEED) & _
            "," & JsonBoolItem(oOptions, "Persistent") & _
            "," & Znl(C_Lng(JsonItem(oOptions, "BaudRate")), 8) & _
            "," & IIf(JsonBoolItem(oOptions, "Parity"), "Y", "N") & _
            "," & Znl(C_Lng(JsonItem(oOptions, "StopBits")), 1)
    End If
QH:
End Function

Public Function InitRequest( _
            sType As String, _
            sUrl As String, _
            Optional ByVal Timeout As Single, _
            Optional ByVal Async As Boolean, _
            Optional RetVal As Object) As Object
    Const FUNC_NAME     As String = "InitRequest"
    Const SNG_EPSILON   As Single = 0.0001
    Static lUseXmlHttp  As Long
    Dim lMili           As Long
    
    On Error GoTo EH
    If lUseXmlHttp = 0 Then
        lUseXmlHttp = IIf(Val(GetEnvironmentVar("_UCS_FISCAL_PRINTER_USE_XMLHTTP")) <> 0, 1, -1)
    End If
    '--- first try server-side XMLHTTP because it has timeouts
    If lUseXmlHttp <> 1 And Timeout <> 30 Then
        Set RetVal = CreateObject("MSXML2.ServerXMLHTTP")
    End If
    If Not RetVal Is Nothing Then
        lMili = Switch(Timeout < -SNG_EPSILON, 0, Timeout > SNG_EPSILON, Timeout, True, 30) * 1000
        RetVal.SetTimeouts 5000, 5000, lMili, lMili
    Else
        Set RetVal = CreateObject("MSXML2.XMLHTTP")
    End If
    RetVal.Open sType, sUrl, Async
    Set InitRequest = RetVal
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Public Function ParseUrl(sUrl As String, uParsed As UcsParsedUrl) As Boolean
    Const FUNC_NAME     As String = "ParseUrl"
    
    On Error GoTo EH
    With CreateObject("VBScript.RegExp")
        .Global = True
        .Pattern = "^(.*)://(?:(?:([^:]*):)?([^@]*)@)?([A-Za-z0-9\-\.]+)(:[0-9]+)?(.*)$"
        With .Execute(sUrl)
            If .Count > 0 Then
                With .Item(0).SubMatches
                    uParsed.Protocol = .Item(0)
                    uParsed.User = .Item(1)
                    If LenB(uParsed.User) = 0 Then
                        uParsed.User = .Item(2)
                    Else
                        uParsed.Pass = .Item(2)
                    End If
                    uParsed.Host = .Item(3)
                    uParsed.Port = Val(Mid$(.Item(4), 2))
                    If uParsed.Port = 0 Then
                        Select Case LCase$(uParsed.Protocol)
                        Case "https"
                            uParsed.Port = 443
                        Case "socks5"
                            uParsed.Port = 1080
                        Case Else
                            uParsed.Port = 80
                        End Select
                    End If
                    uParsed.Path = .Item(5)
                    If LenB(uParsed.Path) = 0 Then
                        uParsed.Path = "/"
                    End If
                End With
                ParseUrl = True
            End If
        End With
    End With
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function SetCurrentDateTimer(ByVal dDate As Date, dblTimer As Double, Optional Error As String) As Boolean
    m_dCurrentStartDate = dDate
    m_dblCurrentStartTimer = dblTimer
    Error = vbNullString
    '--- success
    SetCurrentDateTimer = True
End Function

Property Get GetCurrentNow() As Date
    If m_dCurrentStartDate = 0 Then
        GetCurrentNow = VBA.Now
    Else
        GetCurrentNow = DateAdd("s", TimerEx - m_dblCurrentStartTimer, m_dCurrentStartDate)
    End If
End Property

Property Get GetCurrentTimer() As Double
    GetCurrentTimer = TimerEx - m_dblCurrentStartTimer
End Property

Property Get GetCurrentDate() As Date
    GetCurrentDate = Fix(GetCurrentNow)
End Property

Public Function ToDeviceSerialNo(sText As String) As String
    If preg_match("^[a-zA-Z]{2}\d{6}$", sText) > 0 Then
        ToDeviceSerialNo = Trim$(sText)
    End If
End Function

Public Function ConcatArrays(vSrc As Variant, vDst As Variant) As Variant
    Const FUNC_NAME     As String = "ConcatArrays"
    Dim vResult         As Variant
    Dim lIdx            As Long
    Dim vElem           As Variant

    On Error GoTo EH
    If IsArray(vSrc) Then
        lIdx = lIdx + UBound(vSrc) - LBound(vSrc) + 1
    End If
    If IsArray(vDst) Then
        lIdx = lIdx + UBound(vDst) - LBound(vDst) + 1
    End If
    If lIdx > 0 Then
        ReDim vResult(0 To lIdx - 1) As Variant
        lIdx = 0
        If IsArray(vSrc) Then
            For Each vElem In vSrc
                AssignVariant vResult(lIdx), vElem
                lIdx = lIdx + 1
            Next
        End If
        If IsArray(vDst) Then
            For Each vElem In vDst
                AssignVariant vResult(lIdx), vElem
                lIdx = lIdx + 1
            Next
        End If
    End If
    ConcatArrays = vResult
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function ParseExtendedDate(sText As String) As Date
    Dim vMatches        As Variant
    Dim lYear           As Long
    
    On Error GoTo QH
    ParseExtendedDate = C_Date(sText)
    If preg_match("\d+", sText, vMatches) > 0 Then
        If C_Lng(At(vMatches, 0)) < 2000 Then
            lYear = At(vMatches, 2)
            ParseExtendedDate = DateSerial(IIf(lYear < 2000, lYear + 2000, lYear), At(vMatches, 1), At(vMatches, 0))
        Else
            ParseExtendedDate = DateSerial(At(vMatches, 0), At(vMatches, 1), At(vMatches, 2))
        End If
        ParseExtendedDate = ParseExtendedDate + TimeSerial(At(vMatches, 3), At(vMatches, 4), At(vMatches, 5, 0))
    End If
QH:
End Function

Public Property Get IsLogDebugEnabled() As Boolean
    IsLogDebugEnabled = Logger.LogLevel >= vbLogEventTypeDebug
End Property

Public Property Get IsLogDataDumpEnabled() As Boolean
    IsLogDataDumpEnabled = Logger.LogLevel >= vbLogEventTypeDataDump
End Property
