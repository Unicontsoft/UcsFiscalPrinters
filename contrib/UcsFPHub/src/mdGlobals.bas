Attribute VB_Name = "mdGlobals"
'=========================================================================
'
' UcsFPHub (c) 2019-2020 by Unicontsoft
'
' Unicontsoft Fiscal Printers Hub
'
' This project is licensed under the terms of the MIT license
' See the LICENSE file in the project root for more information
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "mdGlobals"

'=========================================================================
' Public enums
'=========================================================================

Public Enum UcsFileTypeEnum
    ucsFltAnsi = 1
    ucsFltUnicode
    ucsFltUtf8
    ucsFltUtf8NoBom
End Enum

Public Enum UcsOpenSaveDirectoryType
    ucsOdtPersonal = &H5                         ' My Documents
    ucsOdtMyMusic = &HD                          ' "My Music" folder
    ucsOdtAppData = &H1A                         ' Application Data, new for NT4
    ucsOdtLocalAppData = &H1C                    ' non roaming, user\Local Settings\Application Data
    ucsOdtInternetCache = &H20
    ucsOdtCookies = &H21
    ucsOdtHistory = &H22
    ucsOdtCommonAppData = &H23                   ' All Users\Application Data
    ucsOdtWindows = &H24                         ' GetWindowsDirectory()
    ucsOdtSystem = &H25                          ' GetSystemDirectory()
    ucsOdtProgramFiles = &H26                    ' C:\Program Files
    ucsOdtMyPictures = &H27                      ' My Pictures, new for Win2K
    ucsOdtSystemX86 = &H29
    ucsOdtProgramFilesCommon = &H2B              ' C:\Program Files\Common
    ucsOdtCommonDocuments = &H2E                 ' All Users\Documents
    ucsOdtResources = &H38                       ' %windir%\Resources\, For theme and other windows resources.
    ucsOdtResourcesLocalized = &H39              ' %windir%\Resources\<LangID>, for theme and other windows specific resources.
    ucsOdtCommonAdminTools = &H2F                ' All Users\Start Menu\Programs\Administrative Tools
    ucsOdtAdminTools = &H30                      ' <user name>\Start Menu\Programs\Administrative Tools
    ucsOdtFlagCreate = &H8000&                   ' new for Win2K, or this in to force creation of folder
End Enum

'=========================================================================
' API
'=========================================================================

'--- for VariantChangeType
Private Const VARIANT_ALPHABOOL             As Long = 2
'--- for GetSystemMetrics
Private Const SM_REMOTESESSION              As Long = &H1000
'--- for UrlUnescapeW
Private Const URL_UNESCAPE_AS_UTF8          As Long = &H40000
Private Const INTERNET_MAX_URL_LENGTH       As Long = 2048
'--- for OpenProcessToken
Private Const TOKEN_READ                    As Long = &H20008
'--- for SystemParametersInfo
Private Const SPI_GETICONTITLELOGFONT       As Long = 31
Private Const FW_NORMAL                     As Long = 400
'--- GetDeviceCaps constants
Private Const LOGPIXELSX                    As Long = 88
Private Const LOGPIXELSY                    As Long = 90
'--- for ShellExecuteEx
Private Const SEE_MASK_NOASYNC              As Long = &H100
Private Const SEE_MASK_FLAG_NO_UI           As Long = &H400

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function CommandLineToArgvW Lib "shell32" (ByVal lpCmdLine As Long, pNumArgs As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function ApiSysAllocString Lib "oleaut32" Alias "SysAllocString" (ByVal Ptr As Long) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function VariantChangeType Lib "oleaut32" (Dest As Variant, Src As Variant, ByVal wFlags As Integer, ByVal vt As VbVarType) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function ProcessIdToSessionId Lib "kernel32" (ByVal dwProcessID As Long, dwSessionID As Long) As Long
Private Declare Function UrlUnescapeW Lib "shlwapi" (ByVal pszURL As Long, ByVal pszUnescaped As Long, ByRef cchUnescaped As Long, ByVal dwFlags As Long) As Long
Private Declare Function IsTextUnicode Lib "advapi32" (lpBuffer As Any, ByVal cb As Long, lpi As Long) As Long
Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryW" (ByVal lpPathName As Long, ByVal lpSecurityAttributes As Long) As Long
Private Declare Function CreateFileMoniker Lib "ole32" (ByVal lpszPathName As Long, pResult As IUnknown) As Long
Private Declare Function GetRunningObjectTable Lib "ole32" (ByVal dwReserved As Long, pResult As IUnknown) As Long
Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As Long, ByVal oVft As Long, ByVal lCc As Long, ByVal vtReturn As VbVarType, ByVal cActuals As Long, prgVt As Any, prgpVarg As Any, pvargResult As Variant) As Long
Private Declare Function ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long
Private Declare Function GetAdaptersInfo Lib "iphlpapi" (lpAdapterInfo As Any, lpSize As Long) As Long
Private Declare Function SHGetFolderPath Lib "shfolder" Alias "SHGetFolderPathA" (ByVal hWnd As Long, ByVal csidl As Long, ByVal hToken As Long, ByVal dwFlags As Long, ByVal szPath As String) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessID As Long) As Long
Private Declare Function GetProcessTimes Lib "kernel32" (ByVal hProcess As Long, lpCreationTime As FILETIME, lpExitTime As FILETIME, lpKernelTime As FILETIME, lpUserTime As FILETIME) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function SystemTimeToVariantTime Lib "oleaut32" (lpSystemTime As SYSTEMTIME, pvTime As Date) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function GetTokenInformation Lib "advapi32" (ByVal TokenHandle As Long, ByVal TokenInformationClass As Long, TokenInformation As Any, ByVal TokenInformationLength As Long, ReturnLength As Long) As Long
Private Declare Function LookupAccountSid Lib "advapi32" Alias "LookupAccountSidA" (ByVal lpSystemName As String, ByVal sID As Long, ByVal Name As String, cbName As Long, ByVal ReferencedDomainName As String, cbReferencedDomainName As Long, peUse As Long) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function lstrlenA Lib "kernel32" (lpString As Any) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function ShellExecuteEx Lib "shell32" Alias "ShellExecuteExA" (lpExecInfo As SHELLEXECUTEINFO) As Long

Private Type FILETIME
    dwLowDateTime   As Long
    dwHighDateTime  As Long
End Type

Private Type SYSTEMTIME
    wYear           As Integer
    wMonth          As Integer
    wDayOfWeek      As Integer
    wDay            As Integer
    wHour           As Integer
    wMinute         As Integer
    wSecond         As Integer
    wMilliseconds   As Integer
End Type

Private Type LOGFONT
    lfHeight            As Long
    lfWidth             As Long
    lfEscapement        As Long
    lfOrientation       As Long
    lfWeight            As Long
    lfItalic            As Byte
    lfUnderline         As Byte
    lfStrikeOut         As Byte
    lfCharSet           As Byte
    lfOutPrecision      As Byte
    lfClipPrecision     As Byte
    lfQuality           As Byte
    lfPitchAndFamily    As Byte
    lfFaceName(1 To 32) As Byte
End Type

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
' Constants and member variables
'=========================================================================

Private m_sErrComputerName          As String
Private m_sngScreenTwipsPerPixelX   As Single
Private m_sngScreenTwipsPerPixelY   As Single
Private m_sngOrigTwipsPerPixelX     As Single
Private m_sngOrigTwipsPerPixelY     As Single
Private m_dCurrentStartDate         As Date
Private m_dblCurrentStartTimer      As Double

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

Private Sub RaiseError(sFunction As String)
    PrintError sFunction
    Err.Raise Err.Number, MODULE_NAME & "." & sFunction & vbCrLf & Err.Source, Err.Description
End Sub

'=========================================================================
' Functions
'=========================================================================

Public Function SplitArgs(sText As String) As Variant
    Dim vRetVal         As Variant
    Dim lPtr            As Long
    Dim lArgc           As Long
    Dim lIdx            As Long
    Dim lArgPtr         As Long

    If LenB(sText) <> 0 Then
        lPtr = CommandLineToArgvW(StrPtr(sText), lArgc)
    End If
    If lArgc > 0 Then
        ReDim vRetVal(0 To lArgc - 1) As String
        For lIdx = 0 To UBound(vRetVal)
            Call CopyMemory(lArgPtr, ByVal lPtr + 4 * lIdx, 4)
            vRetVal(lIdx) = SysAllocString(lArgPtr)
        Next
    Else
        vRetVal = Split(vbNullString)
    End If
    Call LocalFree(lPtr)
    SplitArgs = vRetVal
End Function

Private Function SysAllocString(ByVal lPtr As Long) As String
    Dim lTemp           As Long

    lTemp = ApiSysAllocString(lPtr)
    Call CopyMemory(ByVal VarPtr(SysAllocString), lTemp, 4)
End Function

Public Property Get InIde() As Boolean
    Debug.Assert pvSetTrue(InIde)
End Property

Private Function pvSetTrue(bValue As Boolean) As Boolean
    bValue = True
    pvSetTrue = True
End Function

Public Function ReadTextFile(sFile As String) As String
    Const FUNC_NAME     As String = "ReadTextFile"
    Const ForReading    As Long = 1
    Const BOM_UTF       As String = "ï»¿"   '--- "\xEF\xBB\xBF"
    Const BOM_UNICODE   As String = "ÿþ"    '--- "\xFF\xFE"
    Dim lSize           As Long
    Dim sPrefix         As String
    Dim nFile           As Integer
    Dim sCharset        As String
    Dim oStream         As ADODB.Stream
    
    '--- get file size
    On Error GoTo EH
    If FileExists(sFile) Then
        lSize = FileLen(sFile)
    End If
    If lSize = 0 Then
        Exit Function
    End If
    '--- read first 50 chars
    nFile = FreeFile
    Open sFile For Binary Access Read Shared As nFile
    sPrefix = String$(IIf(lSize < 50, lSize, 50), 0)
    Get nFile, , sPrefix
    Close nFile
    '--- figure out charset
    If Left$(sPrefix, 3) = BOM_UTF Then
        sCharset = "UTF-8"
    ElseIf Left$(sPrefix, 2) = BOM_UNICODE Or IsTextUnicode(ByVal sPrefix, Len(sPrefix), &HFFFF& - 2) <> 0 Then
        sCharset = "Unicode"
    ElseIf InStr(1, sPrefix, "<?xml", vbTextCompare) > 0 And InStr(1, sPrefix, "utf-8", vbTextCompare) > 0 Then
        '--- special xml encoding test
        sCharset = "UTF-8"
    End If
    '--- plain text: direct VB6 read
    If LenB(ReadTextFile) = 0 And LenB(sCharset) = 0 Then
        nFile = FreeFile
        Open sFile For Binary Access Read Shared As nFile
        ReadTextFile = String$(lSize, 0)
        Get nFile, , ReadTextFile
        Close nFile
    End If
    '--- plain text + unicode: use FileSystemObject
    If LenB(ReadTextFile) = 0 And sCharset <> "UTF-8" Then
        On Error Resume Next  '--- checked
        ReadTextFile = CreateObject("Scripting.FileSystemObject").OpenTextFile(sFile, ForReading, False, sCharset = "Unicode").ReadAll()
        On Error GoTo EH
    End If
    '--- plain text + unicode + utf-8: use ADODB.Stream
    If LenB(ReadTextFile) = 0 Then
        Set oStream = New ADODB.Stream
        With oStream
            .Open
            If LenB(sCharset) <> 0 Then
                .Charset = sCharset
            End If
            .LoadFromFile sFile
            ReadTextFile = .ReadText()
        End With
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Err.Raise Err.Number, MODULE_NAME & "." & FUNC_NAME & vbCrLf & Err.Source, Err.Description
End Function

Public Sub WriteTextFile(sFile As String, sText As String, ByVal eType As UcsFileTypeEnum)
    Const FUNC_NAME     As String = "WriteTextFile"
    Dim oStream         As ADODB.Stream
    Dim oBinStream      As ADODB.Stream
    
    On Error GoTo EH
    MkPath Left$(sFile, InStrRev(sFile, "\"))
    Set oStream = New ADODB.Stream
    With oStream
        .Open
        Select Case eType
        Case ucsFltUnicode
            .Charset = "Unicode"
        Case ucsFltUtf8, ucsFltUtf8NoBom
            .Charset = "UTF-8"
        Case Else
            .Charset = "Windows-1251"
        End Select
        .WriteText sText
        If eType = ucsFltUtf8NoBom Then
            .Position = 3
            Set oBinStream = New ADODB.Stream
            oBinStream.Type = adTypeBinary
            oBinStream.Mode = adModeReadWrite
            oBinStream.Open
            .CopyTo oBinStream
            .Close
            '--- don't log save errors
            On Error GoTo 0
            oBinStream.SaveToFile sFile, adSaveCreateOverWrite
            On Error GoTo EH
        Else
            On Error GoTo 0
            .SaveToFile sFile, adSaveCreateOverWrite
            On Error GoTo EH
        End If
    End With
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Public Function MkPath(sPath As String) As Boolean
    Dim lAttrib         As Long
    
    lAttrib = GetFileAttributes(sPath)
    If lAttrib = -1 Then
        If InStrRev(sPath, "\") > 0 Then
            If Not MkPath(Left$(sPath, InStrRev(sPath, "\") - 1)) Then
                Exit Function
            End If
        End If
        If CreateDirectory(StrPtr(sPath), 0) = 0 Then
            Exit Function
        End If
    ElseIf (lAttrib And vbDirectory + vbVolume) = 0 Then
        Exit Function
    End If
    '--- success
    MkPath = True
End Function

Public Function PathCombine(sPath As String, sFile As String) As String
    PathCombine = sPath & IIf(LenB(sPath) <> 0 And Right$(sPath, 1) <> "\" And LenB(sFile) <> 0, "\", vbNullString) & sFile
End Function

Public Function FileExists(sFile As String) As Boolean
    If GetFileAttributes(sFile) = -1 Then ' INVALID_FILE_ATTRIBUTES
    Else
        FileExists = True
    End If
End Function

Public Function GetOpt(vArgs As Variant, Optional OptionsWithArg As String) As Object
    Dim oRetVal         As Object
    Dim lIdx            As Long
    Dim bNoMoreOpt      As Boolean
    Dim vOptArg         As Variant
    Dim vElem           As Variant

    vOptArg = Split(OptionsWithArg, ":")
    Set oRetVal = CreateObject("Scripting.Dictionary")
    With oRetVal
        .CompareMode = vbTextCompare
        For lIdx = 0 To UBound(vArgs)
            Select Case Left$(At(vArgs, lIdx), 1 + bNoMoreOpt)
            Case "-", "/"
                For Each vElem In vOptArg
                    If Mid$(At(vArgs, lIdx), 2, Len(vElem)) = vElem Then
                        If Mid(At(vArgs, lIdx), Len(vElem) + 2, 1) = ":" Then
                            .Item("-" & vElem) = Mid$(At(vArgs, lIdx), Len(vElem) + 3)
                        ElseIf Len(At(vArgs, lIdx)) > Len(vElem) + 1 Then
                            .Item("-" & vElem) = Mid$(At(vArgs, lIdx), Len(vElem) + 2)
                        ElseIf LenB(At(vArgs, lIdx + 1)) <> 0 Then
                            .Item("-" & vElem) = At(vArgs, lIdx + 1)
                            lIdx = lIdx + 1
                        Else
                            .Item("error") = "Option -" & vElem & " requires an argument"
                        End If
                        GoTo Continue
                    End If
                Next
                .Item("-" & Mid$(At(vArgs, lIdx), 2)) = True
            Case Else
                .Item("numarg") = .Item("numarg") + 1
                .Item("arg" & .Item("numarg")) = At(vArgs, lIdx)
            End Select
Continue:
        Next
    End With
    Set GetOpt = oRetVal
End Function

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

Public Function Peek(ByVal lPtr As Long) As Long
    Call CopyMemory(Peek, ByVal lPtr, 4)
End Function

Public Function PeekInt(ByVal lPtr As Long) As Integer
    Call CopyMemory(PeekInt, ByVal lPtr, 2)
End Function

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

Public Function C_Obj(Value As Variant) As Object
    Dim vDest           As Variant

    If VarType(Value) = vbObject Then
        Set C_Obj = Value
    ElseIf VariantChangeType(vDest, Value, 0, vbObject) = 0 Then
        Set C_Obj = vDest
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
    ElseIf VariantChangeType(vDest, Value, 0, vbDouble) = 0 Then
        C_Dbl = vDest
    End If
End Function

Public Function Zn(sText As String, Optional IfEmptyString As Variant = Null, Optional EmptyString As String) As Variant
    Zn = IIf(sText = EmptyString, IfEmptyString, sText)
End Function

Public Function Znl(ByVal lValue As Long, Optional IfEmptyLong As Variant = Null, Optional ByVal EmptyLong As Long = 0) As Variant
    Znl = IIf(lValue = EmptyLong, IfEmptyLong, lValue)
End Function

Public Function preg_match(find_re As String, sText As String, Optional Matches As Variant, Optional Indexes As Variant) As Long
    Const FUNC_NAME     As String = "preg_match"
    Dim lIdx            As Long
    Dim oMatches        As Object
    
    On Error GoTo EH
    Set oMatches = InitRegExp(find_re).Execute(sText)
    With oMatches
        preg_match = .Count
        If Not IsMissing(Matches) Then
            If .Count = 0 Then
                Matches = Split(vbNullString)
            ElseIf .Count = 1 Then
                With .Item(0)
                    If .SubMatches.Count = 0 Then
                        ReDim Matches(0 To 0) As String
                        Matches(0) = .Value
                    Else
                        ReDim Matches(0 To .SubMatches.Count - 1) As String
                        For lIdx = 0 To .SubMatches.Count - 1
                            Matches(lIdx) = .SubMatches(lIdx)
                        Next
                    End If
                End With
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
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Public Function Printf(ByVal sText As String, ParamArray A() As Variant) As String
    Const LNG_PRIVATE   As Long = &HE1B6 '-- U+E000 to U+F8FF - Private Use Area (PUA)
    Dim lIdx            As Long
    
    For lIdx = UBound(A) To LBound(A) Step -1
        sText = Replace(sText, "%" & (lIdx - LBound(A) + 1), Replace(A(lIdx), "%", ChrW$(LNG_PRIVATE)))
    Next
    Printf = Replace(sText, ChrW$(LNG_PRIVATE), "%")
End Function

Public Property Get TimerEx() As Double
    Dim cFreq           As Currency
    Dim cValue          As Currency
    
    Call QueryPerformanceFrequency(cFreq)
    Call QueryPerformanceCounter(cValue)
    TimerEx = cValue / cFreq
End Property

Public Function GetErrorTempPath() As String
    Dim sBuffer         As String
    
    sBuffer = String$(2000, 0)
    Call GetTempPath(Len(sBuffer) - 1, sBuffer)
    GetErrorTempPath = Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    If Right$(GetErrorTempPath, 1) = "\" Then
        GetErrorTempPath = Left$(GetErrorTempPath, Len(GetErrorTempPath) - 1)
    End If
End Function

Public Function GetErrorComputerName(Optional ByVal NoSession As Boolean) As String
    Dim lSize           As Long
    
    If LenB(m_sErrComputerName) = 0 Then
        m_sErrComputerName = Space$(256): lSize = 255
        If GetComputerName(m_sErrComputerName, lSize) > 0 Then
            m_sErrComputerName = Left$(m_sErrComputerName, lSize)
        Else
            m_sErrComputerName = vbNullString
        End If
    End If
    GetErrorComputerName = m_sErrComputerName
    If GetSystemMetrics(SM_REMOTESESSION) <> 0 And Not NoSession Then
        lSize = -1
        On Error Resume Next '--- checked
        Call ProcessIdToSessionId(GetCurrentProcessId(), lSize)
        On Error GoTo 0
        If lSize <> -1 Then
            GetErrorComputerName = GetErrorComputerName & ":" & lSize
        End If
    End If
End Function

Public Function GetErrorProcessCreationTime(Optional ByVal lPID As Long) As Date
    Const PROCESS_QUERY_INFORMATION     As Long = &H400&
    Dim hProcess    As Long
    Dim uCreation   As FILETIME
    Dim uDummy      As FILETIME
    
    If lPID = 0 Then
        lPID = GetCurrentProcessId()
    End If
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, 0, lPID)
    If hProcess <> 0 Then
        If GetProcessTimes(hProcess, uCreation, uDummy, uDummy, uDummy) <> 0 Then
            GetErrorProcessCreationTime = pvFromFileTime(uCreation)
        End If
        Call CloseHandle(hProcess)
    End If
End Function

Private Function pvFromFileTime(uTime As FILETIME) As Date
    Dim uLocalTime      As FILETIME
    Dim uSysTime        As SYSTEMTIME
    
    If FileTimeToLocalFileTime(uTime, uLocalTime) <> 0 Then
        If FileTimeToSystemTime(uLocalTime, uSysTime) <> 0 Then
            Call SystemTimeToVariantTime(uSysTime, pvFromFileTime)
        End If
    End If
End Function

Public Function GetCurrentProcessUser(Optional ByVal IncludeDomain As Boolean = True) As String
    Dim hProcessID      As Long
    Dim hToken          As Long
    Dim lNeeded         As Long
    Dim baBuffer()      As Byte
    Dim sUser           As String
    Dim sDomain         As String
    
    hProcessID = GetCurrentProcess()
    If hProcessID <> 0 Then
        If OpenProcessToken(hProcessID, TOKEN_READ, hToken) = 1 Then
            Call GetTokenInformation(hToken, 1, ByVal 0, 0, lNeeded)
                ReDim baBuffer(0 To lNeeded) As Byte
                '--- enum TokenInformationClass { TokenUser = 1, TokenGroups = 2, ... }
                If GetTokenInformation(hToken, 1, baBuffer(0), UBound(baBuffer), lNeeded) = 1 Then
                    sUser = String$(1000, 0)
                    sDomain = String$(1000, 0)
                    If LookupAccountSid(vbNullString, Peek(VarPtr(baBuffer(0))), sUser, Len(sUser), sDomain, Len(sDomain), 0) = 1 Then
                        If IncludeDomain Then
                            GetCurrentProcessUser = Left$(sDomain, InStr(sDomain, vbNullChar) - 1)
                            If LenB(GetCurrentProcessUser) <> 0 Then
                                GetCurrentProcessUser = GetCurrentProcessUser & "\"
                            End If
                        End If
                        GetCurrentProcessUser = GetCurrentProcessUser & Left$(sUser, InStr(sUser, vbNullChar) - 1)
                    End If
                End If
        End If
        Call CloseHandle(hProcessID)
    End If
End Function

Public Function GetProcessName() As String
    GetProcessName = String$(1000, 0)
    Call GetModuleFileName(0, GetProcessName, Len(GetProcessName) - 1)
    GetProcessName = Left$(GetProcessName, InStr(GetProcessName, vbNullChar) - 1)
End Function

Public Function GetEnvironmentVar(sName As String) As String
    Dim sBuffer         As String
    
    sBuffer = String$(2000, 0)
    Call GetEnvironmentVariable(sName, sBuffer, Len(sBuffer) - 1)
    GetEnvironmentVar = Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1)
End Function

Public Sub AssignVariant(vDest As Variant, vSrc As Variant)
    On Error GoTo QH
    If IsObject(vSrc) Then
        Set vDest = vSrc
    Else
        vDest = vSrc
    End If
QH:
End Sub

Private Function pvParseTokenByRegExp(sText As String, sPattern As String) As String
    Dim oCol            As Object
    
    Set oCol = InitRegExp(sPattern).Execute(sText)
    If oCol.Count > 0 Then
        pvParseTokenByRegExp = oCol.Item(0).SubMatches(0)
        sText = Mid$(sText, oCol.Item(0).FirstIndex + oCol.Item(0).Length + 1)
    End If
End Function

Public Function ParseQueryString(ByVal sQueryString As String) As Object
    Const KEY_PATTERN   As String = "^([^=&#?]+)"
    Const VALUE_PATTERN As String = "^(?:=([^&#?]*))"
    Dim sKey            As String
    Dim oRetVal         As Object
    Dim sBuffer         As String
    Dim lSize           As Long
    
    sBuffer = String$(INTERNET_MAX_URL_LENGTH, 0)
    Do
        sKey = pvParseTokenByRegExp(sQueryString, KEY_PATTERN)
        If LenB(sKey) = 0 Then
            Exit Do
        End If
        lSize = Len(sBuffer)
        Call UrlUnescapeW(StrPtr(pvParseTokenByRegExp(sQueryString, VALUE_PATTERN)), StrPtr(sBuffer), lSize, URL_UNESCAPE_AS_UTF8)
        JsonItem(oRetVal, sKey) = Left$(sBuffer, lSize)
    Loop
    Set ParseQueryString = oRetVal
End Function

Public Function ParseConnectString(ByVal sConnStr As String, Optional Separator As String = ";") As Object
    Set ParseConnectString = ParseDeviceString(sConnStr, Separator)
End Function

Public Function ToConnectString(oMap As Object, Optional Separator As String = ";") As String
    ToConnectString = ToDeviceString(oMap, Separator)
End Function

Public Function Quote(sText As String) As String
    Quote = Replace(sText, "'", "''")
End Function

Public Function CryptRC4(sText As String, sKey As String) As String
    Dim baS(0 To 255) As Byte
    Dim baK(0 To 255) As Byte
    Dim lI          As Long
    Dim lJ          As Long
    Dim lSwap       As Long
    Dim lIdx        As Long

    For lIdx = 0 To 255
        baS(lIdx) = lIdx
        baK(lIdx) = Asc(Mid$(sKey, 1 + (lIdx Mod Len(sKey)), 1))
    Next
    For lI = 0 To 255
        lJ = (lJ + baS(lI) + baK(lI)) Mod 256
        lSwap = baS(lI)
        baS(lI) = baS(lJ)
        baS(lJ) = lSwap
    Next
    lI = 0
    lJ = 0
    For lIdx = 1 To Len(sText)
        lI = (lI + 1) Mod 256
        lJ = (lJ + baS(lI)) Mod 256
        lSwap = baS(lI)
        baS(lI) = baS(lJ)
        baS(lJ) = lSwap
        CryptRC4 = CryptRC4 & Chr$((pvCryptXor(baS((CLng(baS(lI)) + baS(lJ)) Mod 256), Asc(Mid$(sText, lIdx, 1)))))
    Next
End Function

Private Function pvCryptXor(ByVal lI As Long, ByVal lJ As Long) As Long
    If lI = lJ Then
        pvCryptXor = lJ
    Else
        pvCryptXor = lI Xor lJ
    End If
End Function

Public Function LocateFile(sFile As String) As String
    Const FUNC_NAME     As String = "LocateFile"
    Dim sDir            As String
    Dim sName           As String
    Dim lPos            As Long
    
    On Error GoTo EH
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
    Exit Function
EH:
    PrintError FUNC_NAME
    Err.Raise Err.Number, MODULE_NAME & "." & FUNC_NAME & vbCrLf & Err.Source, Err.Description
End Function

Public Function PutObject(oObj As Object, sPathName As String, Optional ByVal Flags As Long) As Long
    Const ROTFLAGS_REGISTRATIONKEEPSALIVE As Long = 1
    Const IDX_REGISTER  As Long = 3
    Dim hResult         As Long
    Dim pROT            As IUnknown
    Dim pMoniker        As IUnknown
    
    hResult = GetRunningObjectTable(0, pROT)
    If hResult < 0 Then
        Err.Raise hResult, "GetRunningObjectTable"
    End If
    hResult = CreateFileMoniker(StrPtr(sPathName), pMoniker)
    If hResult < 0 Then
        Err.Raise hResult, "CreateFileMoniker"
    End If
    DispCallByVtbl pROT, IDX_REGISTER, ROTFLAGS_REGISTRATIONKEEPSALIVE Or Flags, ObjPtr(oObj), ObjPtr(pMoniker), VarPtr(PutObject)
End Function

Public Sub RevokeObject(ByVal lCookie As Long)
    Const IDX_REVOKE    As Long = 4
    Dim hResult         As Long
    Dim pROT            As IUnknown
    
    hResult = GetRunningObjectTable(0, pROT)
    If hResult < 0 Then
        Err.Raise hResult, "GetRunningObjectTable"
    End If
    DispCallByVtbl pROT, IDX_REVOKE, lCookie
End Sub

Public Function IsObjectRunning(sPathName As String) As Boolean
    Const IDX_ISRUNNING As Long = 5
    Const S_OK          As Long = 0
    Dim hResult         As Long
    Dim pROT            As IUnknown
    Dim pMoniker        As IUnknown
    
    hResult = GetRunningObjectTable(0, pROT)
    If hResult < 0 Then
        Err.Raise hResult, "GetRunningObjectTable"
    End If
    hResult = CreateFileMoniker(StrPtr(sPathName), pMoniker)
    If hResult < 0 Then
        Err.Raise hResult, "CreateFileMoniker"
    End If
    If DispCallByVtbl(pROT, IDX_ISRUNNING, ObjPtr(pMoniker)) = S_OK Then
        '--- success
        IsObjectRunning = True
    End If
End Function

Private Function DispCallByVtbl(pUnk As IUnknown, ByVal lIndex As Long, ParamArray A() As Variant) As Variant
    Const CC_STDCALL    As Long = 4
    Dim lIdx            As Long
    Dim vParam()        As Variant
    Dim vType(0 To 63)  As Integer
    Dim vPtr(0 To 63)   As Long
    Dim hResult         As Long
    
    vParam = A
    For lIdx = 0 To UBound(vParam)
        vType(lIdx) = VarType(vParam(lIdx))
        vPtr(lIdx) = VarPtr(vParam(lIdx))
    Next
    hResult = DispCallFunc(ObjPtr(pUnk), lIndex * 4, CC_STDCALL, vbLong, lIdx, vType(0), vPtr(0), DispCallByVtbl)
    If hResult < 0 Then
        Err.Raise hResult, "DispCallFunc"
    End If
End Function

Public Sub JsonExpandEnviron(ByVal oJson As Object)
    Dim vKey            As Variant
    Dim sText           As String
    Dim sExpand         As String
    
    For Each vKey In JsonKeys(oJson)
        If IsObject(JsonItem(oJson, vKey)) Then
            JsonExpandEnviron JsonItem(oJson, vKey)
        Else
            sText = C_Str(JsonItem(oJson, vKey))
            sExpand = String$(ExpandEnvironmentStrings(sText, vbNullString, 0), 0)
            If ExpandEnvironmentStrings(sText, sExpand, Len(sExpand)) > 0 Then
                sExpand = Left$(sExpand, InStr(sExpand, vbNullChar) - 1)
                If sExpand <> sText Then
                    JsonItem(oJson, vKey) = sExpand
                End If
            End If
        End If
    Next
End Sub

Public Function GetMacAddress() As String
    Const OFFSET_LENGTH As Long = 400
    Dim lSize           As Long
    Dim baBuffer()      As Byte
    Dim lIdx            As Long
    Dim sRetVal         As String
    
    Call GetAdaptersInfo(ByVal 0, lSize)
    If lSize <> 0 Then
        ReDim baBuffer(0 To lSize - 1) As Byte
        Call GetAdaptersInfo(baBuffer(0), lSize)
        Call CopyMemory(lSize, baBuffer(OFFSET_LENGTH), 4)
        For lIdx = OFFSET_LENGTH + 4 To OFFSET_LENGTH + 4 + lSize - 1
            sRetVal = IIf(LenB(sRetVal) <> 0, sRetVal & ":", vbNullString) & Right$("0" & Hex$(baBuffer(lIdx)), 2)
        Next
    End If
    GetMacAddress = sRetVal
End Function

Public Function GetSpecialFolder(ByVal eType As UcsOpenSaveDirectoryType) As String
    GetSpecialFolder = String$(1000, 0)
    Call SHGetFolderPath(0, eType, 0, 0, GetSpecialFolder)
    GetSpecialFolder = Left$(GetSpecialFolder, InStr(GetSpecialFolder, vbNullChar) - 1)
End Function

Property Get SystemIconFont() As StdFont
    Dim uFont           As LOGFONT
    Dim sBuffer         As String
    Dim hTempDC         As Long
    
    Call SystemParametersInfo(SPI_GETICONTITLELOGFONT, LenB(uFont), uFont, 0)
    Set SystemIconFont = New StdFont
    With SystemIconFont
        sBuffer = Space$(lstrlenA(uFont.lfFaceName(1)))
        CopyMemory ByVal sBuffer, uFont.lfFaceName(1), Len(sBuffer)
        .Name = sBuffer
        .Bold = (uFont.lfWeight >= FW_NORMAL)
        .Charset = uFont.lfCharSet
        .Italic = (uFont.lfItalic <> 0)
        .Strikethrough = (uFont.lfStrikeOut <> 0)
        .Underline = (uFont.lfUnderline <> 0)
        .Weight = uFont.lfWeight
        hTempDC = GetDC(0)
        .Size = -(uFont.lfHeight * 72) / GetDeviceCaps(hTempDC, LOGPIXELSY)
        Call ReleaseDC(0, hTempDC)
    End With
End Property

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

Public Function ConcatCollection(oCol As Collection, Optional Separator As String = vbCrLf) As String
    Dim lSize           As Long
    Dim vElem           As Variant
    
    For Each vElem In oCol
        lSize = lSize + Len(vElem) + Len(Separator)
    Next
    If lSize > 0 Then
        ConcatCollection = String$(lSize - Len(Separator), 0)
        lSize = 1
        For Each vElem In oCol
            If lSize <= Len(ConcatCollection) Then
                Mid$(ConcatCollection, lSize, Len(vElem) + Len(Separator)) = vElem & Separator
            End If
            lSize = lSize + Len(vElem) + Len(Separator)
        Next
    End If
End Function

Public Sub MoveCtl( _
            oCtl As Object, _
            ByVal Left As Single, _
            Optional Top As Variant, _
            Optional Width As Variant, _
            Optional Height As Variant)
    Const FUNC_NAME     As String = "MoveCtl"
    Dim oCtlExt         As VBControlExtender
    
    On Error GoTo EH
    If oCtl Is Nothing Then
        Exit Sub
    End If
    If TypeOf oCtl Is VBControlExtender Then
        Set oCtlExt = oCtl
        If IsMissing(Top) Then
            If oCtlExt.Left <> Left Then
                oCtlExt.Move Left
            End If
        ElseIf IsMissing(Width) Then
            If oCtlExt.Left <> Left Or oCtlExt.Top <> Top Then
                oCtlExt.Move Left, Top
            End If
        ElseIf IsMissing(Height) Then
            If oCtlExt.Left <> Left Or oCtlExt.Top <> Top Or oCtlExt.Width <> Limit(Width, 0) Then
                If 1440 \ ScreenTwipsPerPixelX = 1440 / ScreenTwipsPerPixelX Then
                    oCtlExt.Move Left, Top, Limit(Width, 0)
                ElseIf oCtlExt.Left <> Left Or oCtlExt.Top <> Top Then
                    oCtlExt.Move oCtlExt.Left, oCtlExt.Top, Limit(Width, 0)
                    oCtlExt.Move Left, Top
                Else
                    oCtlExt.Move Left + ScreenTwipsPerPixelX, Top, Limit(Width, 0)
                    oCtlExt.Move Left
                End If
            End If
        Else
            If oCtlExt.Left <> Left Or oCtlExt.Top <> Top Or oCtlExt.Width <> Limit(Width, 0) Or oCtlExt.Height <> Limit(Height, 0) Then
                If 1440 \ ScreenTwipsPerPixelX = 1440 / ScreenTwipsPerPixelX Then
                    oCtlExt.Move Left, Top, Limit(Width, 0), Limit(Height, 0)
                ElseIf oCtlExt.Left <> Left Or oCtlExt.Top <> Top Then
                    oCtlExt.Move oCtlExt.Left, oCtlExt.Top, Limit(Width, 0), Limit(Height, 0)
                    oCtlExt.Move Left, Top
                Else
                    oCtlExt.Move Left + ScreenTwipsPerPixelX, Top, Limit(Width, 0), Limit(Height, 0)
                    oCtlExt.Move Left
                End If
            End If
        End If
    Else
        If IsMissing(Top) Then
            If oCtl.Left <> Left Then
                oCtl.Move Left
            End If
        ElseIf IsMissing(Width) Then
            If oCtl.Left <> Left Or oCtl.Top <> Top Then
                oCtl.Move Left, Top
            End If
        ElseIf IsMissing(Height) Then
            If oCtl.Left <> Left Or oCtl.Top <> Top Or oCtl.Width <> Limit(Width, 0) Then
                If 1440 \ ScreenTwipsPerPixelX = 1440 / ScreenTwipsPerPixelX Then
                    oCtl.Move Left, Top, Limit(Width, 0)
                ElseIf oCtl.Left <> Left Or oCtl.Top <> Top Then
                    oCtl.Move oCtl.Left, oCtl.Top, Limit(Width, 0)
                    oCtl.Move Left, Top
                Else
                    oCtl.Move Left + ScreenTwipsPerPixelX, Top, Limit(Width, 0)
                    oCtl.Move Left
                End If
            End If
        Else
            If oCtl.Left <> Left Or oCtl.Top <> Top Or oCtl.Width <> Limit(Width, 0) Or oCtl.Height <> Limit(Height, 0) Then
                If 1440 \ ScreenTwipsPerPixelX = 1440 / ScreenTwipsPerPixelX Then
                    oCtl.Move Left, Top, Limit(Width, 0), Limit(Height, 0)
                ElseIf oCtl.Left <> Left Or oCtl.Top <> Top Then
                    oCtl.Move oCtl.Left, oCtl.Top, Limit(Width, 0), Limit(Height, 0)
                    oCtl.Move Left, Top
                Else
                    oCtl.Move Left + ScreenTwipsPerPixelX, Top, Limit(Width, 0), Limit(Height, 0)
                    oCtl.Move Left
                End If
            End If
        End If
    End If
    Exit Sub
EH:
    RaiseError FUNC_NAME
End Sub

Property Get ScreenTwipsPerPixelX() As Single
    If m_sngScreenTwipsPerPixelX <> 0 Then
        ScreenTwipsPerPixelX = m_sngScreenTwipsPerPixelX
    ElseIf Screen.TwipsPerPixelX <> 0 Then
        ScreenTwipsPerPixelX = Screen.TwipsPerPixelX
    Else
        ScreenTwipsPerPixelX = 15
    End If
End Property

Property Get ScreenTwipsPerPixelY() As Single
    If m_sngScreenTwipsPerPixelY <> 0 Then
        ScreenTwipsPerPixelY = m_sngScreenTwipsPerPixelY
    ElseIf Screen.TwipsPerPixelY <> 0 Then
        ScreenTwipsPerPixelY = Screen.TwipsPerPixelY
    Else
        ScreenTwipsPerPixelY = 15
    End If
End Property

Property Get OrigTwipsPerPixelX() As Single
    If m_sngOrigTwipsPerPixelX <> 0 Then
        OrigTwipsPerPixelX = m_sngOrigTwipsPerPixelX
    Else
        OrigTwipsPerPixelX = ScreenTwipsPerPixelX
    End If
End Property

Property Get OrigTwipsPerPixelY() As Single
    If m_sngOrigTwipsPerPixelY <> 0 Then
        OrigTwipsPerPixelY = m_sngOrigTwipsPerPixelY
    Else
        OrigTwipsPerPixelY = ScreenTwipsPerPixelY
    End If
End Property

Public Function Limit( _
            ByVal Value As Double, _
            Optional Min As Variant, _
            Optional Max As Variant) As Double
    Const FUNC_NAME     As String = "Limit"
    
    On Error GoTo EH
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
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Sub ApplyTheme()
    Dim hScreenDC           As Long
    
    hScreenDC = GetDC(0)
    m_sngScreenTwipsPerPixelX = GetDeviceCaps(hScreenDC, LOGPIXELSX)
    m_sngScreenTwipsPerPixelY = GetDeviceCaps(hScreenDC, LOGPIXELSY)
    Call ReleaseDC(0, hScreenDC)
    If m_sngScreenTwipsPerPixelX <> 0 Then
        m_sngOrigTwipsPerPixelX = Int(1440 / m_sngScreenTwipsPerPixelX)
        m_sngScreenTwipsPerPixelX = 1440 / m_sngScreenTwipsPerPixelX
    End If
    If m_sngScreenTwipsPerPixelY <> 0 Then
        m_sngOrigTwipsPerPixelY = Int(1440 / m_sngScreenTwipsPerPixelY)
        m_sngScreenTwipsPerPixelY = 1440 / m_sngScreenTwipsPerPixelY
    End If
End Sub

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

Public Function IconScale(ByVal sngSize As Single) As Long
    If ScreenTwipsPerPixelX < 6.5 Then
        IconScale = Int(sngSize * 3)
    ElseIf ScreenTwipsPerPixelX < 9.5 Then
        IconScale = Int(sngSize * 2)
    ElseIf ScreenTwipsPerPixelX < 11.5 Then
        IconScale = Int(sngSize * 3 \ 2)
    Else
        IconScale = Int(sngSize * 1)
    End If
End Function

Public Function ShellExec(sExeFile As String, sParams As String) As Boolean
    Dim uShell          As SHELLEXECUTEINFO
    
    With uShell
        .cbSize = Len(uShell)
        .fMask = SEE_MASK_NOASYNC Or SEE_MASK_FLAG_NO_UI
        .lpFile = sExeFile
        .lpParameters = sParams
    End With
    Call ShellExecuteEx(uShell)
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

Public Function IsOnlyDigits(ByVal sText As String) As Boolean
    If LenB(sText) <> 0 Then
        IsOnlyDigits = Not (sText Like "*[!0-9]*")
    End If
End Function
