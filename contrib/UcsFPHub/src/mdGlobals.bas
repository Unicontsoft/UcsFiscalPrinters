Attribute VB_Name = "mdGlobals"
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

'=========================================================================
' Constants and member variables
'=========================================================================

Private m_cRegExpCache              As Collection
Private m_sErrComputerName          As String

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
    PrintError FUNC_NAME
    Err.Raise Err.Number, MODULE_NAME & "." & FUNC_NAME & vbCrLf & Err.Source, Err.Description
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

Public Function At(vData As Variant, ByVal lIdx As Long, Optional sDefault As String) As String
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
End Function

Public Function Peek(ByVal lPtr As Long) As Long
    Call CopyMemory(Peek, ByVal lPtr, 4)
End Function

Public Function PeekInt(ByVal lPtr As Long) As Integer
    Call CopyMemory(PeekInt, ByVal lPtr, 2)
End Function

Public Function SearchCollection(ByVal pCol As Object, Index As Variant, Optional RetVal As Variant) As Boolean
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
            Err.Raise vbObjectError, , "Not implemented"
        End If
    Else
        If TypeOf pCol Is IVbCollection Then
            Set pVbCol = pCol
            SearchCollection = (pVbCol.Item(Index, vItem) = S_OK)
        Else
            Err.Raise vbObjectError, , "Not implemented"
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

Public Function Split2(sText As String, sDelim As String) As Variant
    Dim lPos            As Long
    
    lPos = InStr(sText, sDelim)
    If lPos > 0 Then
        Split2 = Array(Left$(sText, lPos - 1), Mid$(sText, lPos + Len(sDelim)))
    Else
        Split2 = Array(sText)
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

Public Function ParseConnectString(ByVal sDeviceString As String) As Object
    Const KEY_PATTERN   As String = "^([^=]+)="
    Const VALUE_PATTERN As String = "^\s*('[^']*'|""[^""]*""|[^;]*)\s*;?"
    Dim sKey            As String
    Dim sValue          As String
    Dim oRetVal         As Object
    
    Do
        sKey = Trim$(pvParseTokenByRegExp(sDeviceString, KEY_PATTERN))
        If LenB(sKey) = 0 Then
            Exit Do
        End If
        sValue = Trim$(pvParseTokenByRegExp(sDeviceString, VALUE_PATTERN))
        If Len(sValue) >= 2 Then
            If Left$(sValue, 1) = Right$(sValue, 1) Then
                Select Case Asc(sValue)
                Case 34, 39 '--- ' and "
                    sValue = Mid$(sValue, 2, Len(sValue) - 2)
                End Select
            End If
        End If
        JsonItem(oRetVal, sKey) = sValue
    Loop
    Set ParseConnectString = oRetVal
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

Public Function ParseDeviceString(ByVal sDeviceString As String) As Object
    Const KEY_PATTERN   As String = "^([^=]+)="
    Const VALUE_PATTERN As String = "^\s*('[^']*'|""[^""]*""|[^;]*)\s*;?"
    Dim sKey            As String
    Dim sValue          As String
    Dim oRetVal         As Object
    
    Do
        sKey = Trim$(pvParseTokenByRegExp(sDeviceString, KEY_PATTERN))
        If LenB(sKey) = 0 Then
            Exit Do
        End If
        sValue = Trim$(pvParseTokenByRegExp(sDeviceString, VALUE_PATTERN))
        If Len(sValue) >= 2 Then
            If Left$(sValue, 1) = Right$(sValue, 1) Then
                Select Case Asc(sValue)
                Case 34, 39 '--- ' and "
                    sValue = Mid$(sValue, 2, Len(sValue) - 2)
                End Select
            End If
        End If
        JsonItem(oRetVal, sKey) = sValue
    Loop
    Set ParseDeviceString = oRetVal
End Function
