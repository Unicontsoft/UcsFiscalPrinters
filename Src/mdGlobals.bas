Attribute VB_Name = "mdGlobals"
'=========================================================================
' $Header: /UcsFiscalPrinter/Src/mdGlobals.bas 32    30.01.15 15:30 Wqw $
'
'   Unicontsoft Fiscal Printers Project
'   Copyright (c) 2008-2015 Unicontsoft
'
'   Global functions, constants and variables
'
' $Log: /UcsFiscalPrinter/Src/mdGlobals.bas $
' 
' 32    30.01.15 15:30 Wqw
' REF: date timer precision
'
' 31    30.01.15 15:06 Wqw
' ADD: Sub FlushDebugLog
'
' 30    29.01.15 11:45 Wqw
' REF: public debug log file handle
'
' 29    28.01.15 15:34 Wqw
' REF: output debug log keeps log file open for performance
'
' 28    23.01.15 17:58 Wqw
' ADD: global persistent port object
'
' 27    23.01.15 17:11 Wqw
' REF: impl persistent com ports
'
' 26    17.12.14 16:07 Wqw
' REF: impl to ascii
'
' 25    31.07.14 16:06 Wqw
' ADD: Consts
'
' 24    31.03.14 17:34 Wqw
' REF: impl GetConfigForCommand params for localized commands
'
' 23    16.10.13 15:47 Wqw
' REF: impl log rotate on output debug log
'
' 22    18.06.13 17:20 Wqw
' REF: break on all errors. ADD: Function ToHexDump, SearchCollection,
' DispInvoke, DispPropertyGet, ParseSum, Property LockControl
'
' 21    26.02.13 19:38 Wqw
' REF: debug print error on loading config
'
' 20    3.01.13 16:38 Wqw
' ADD: Function SafeFormat, SafeText, AssignVariant, preg_replace,
' GetConfigForCommand, Property DateTimer
'
' 19    10.10.12 15:11 Wqw
' REF: config loading error handling
'
' 18    5.10.12 14:15 Wqw
' REF: datecs protocol name
'
' 17    29.08.12 15:04 Wqw
' ADD: Function FileExists, ReadTextFile, GetConfigValue
'
' 16    6.08.12 18:36 Wqw
' ADD: Function EmptyVariantArray
'
' 15    23.03.12 15:25 Wqw
' ADD: EmptyDoubleArray. REF: err handling
'
' 14    8.12.11 15:47 Wqw
' ADD: Property Let ValueAt
'
' 13    9.08.11 23:25 Wqw
' ADD: Sub RegWriteValue, OpenSaveDialog, ConvertToBW, GdipLoadImage,
' GdipReleaseImage, Pad
'
' 12    4.07.11 15:48 Wqw
' REF: err handling
'
' 11    11.04.11 14:50 Wqw
' ADD: Function Znl
'
' 10    8.03.11 13:19 Wqw
' REF: split cg address replaces non-printable characters
'
' 9     8.03.11 13:04 Wqw
' ADD: Function Limit
'
' 8     7.03.11 19:21 Wqw
' REF: impl milliseconds w OutputDebugLog
'
' 7     24.02.11 16:13 Wqw
' REF: fix RegReadString
'
' 6     24.02.11 16:05 Wqw
' REF: RegReadString razbira ot expand string-owe
'
' 5     22.02.11 13:53 Wqw
' ADD: Consts
'
' 4     22.02.11 10:06 Wqw
' REF: polzwa string functions
'
' 3     21.02.11 16:28 Wqw
' ADD: Function RegReadString, GetSystemDirectory
'
' 2     21.02.11 13:43 Wqw
' ADD: Function SplitCgAddress, AlignText, CenterText, SumArray,
' IsComCtl6Loaded, FixThemeSupport
'
' 1     14.02.11 18:13 Wqw
' Initial implementation
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

Public Enum UcsInvokeCallEnum
    UcsIclMethod = 1
    UcsIclPropGet = 2
    UcsIclPropLet = 4
    UcsIclPropSet = 8
End Enum

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
'--- for CreateDIBSection
Private Const DIB_RGB_COLORS                As Long = 0
'--- for VariantChangeType
Private Const VT_I4                         As Long = 3
Private Const VT_R8                         As Long = 5
'Private Const VT_CY                         As Long = 6
Private Const VT_DATE                       As Long = 7
Private Const VT_BSTR                       As Long = 8
Private Const VT_BOOL                       As Long = 11
'Private Const VT_UI1                        As Long = 17
Private Const VARIANT_ALPHABOOL             As Long = 2
'--- for invoke
Private Const LOCALE_SYSTEM_DEFAULT         As Long = &H800
'--- hresults
Private Const S_OK                          As Long = 0

Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Args As Any) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function QueryDosDevice Lib "kernel32" Alias "QueryDosDeviceA" (ByVal lpDeviceName As Long, ByVal lpTargetPath As String, ByVal ucchMax As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Declare Function GetUserDefaultLCID Lib "kernel32" () As Long
Private Declare Function DllGetVersion Lib "comctl32.dll" (pdvi As DLLVERSIONINFO) As Long
Private Declare Function SetWindowSubclass Lib "comctl32" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
Private Declare Function DefSubclassProc Lib "comctl32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function APIGetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, lpBitsInfo As BITMAPINFOHEADER, ByVal wUsage As Long, lpBits As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal ByteLength As Long)
Private Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal mFilename As Long, ByRef mImage As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal mGraphics As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDC As Long, hGraphics As Long) As Long
Private Declare Function GdipDrawImageRectI Lib "gdiplus" (ByVal Graphics As Long, ByVal img As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long
Private Declare Function GdiplusStartup Lib "gdiplus" (Token As Long, inputbuf As GDIPLUS_STARTUP_INPUT, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Function GdipGetImageDimension Lib "gdiplus" (ByVal Image As Long, ByRef Width As Single, ByRef Height As Single) As Long
Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal Color As Long, ByRef Brush As Long) As Long
Private Declare Function GdipFillRectangleI Lib "gdiplus" (ByVal Graphics As Long, ByVal Brush As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As Long
Private Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal Brush As Long) As Long
Private Declare Function ApiEmptyDoubleArray Lib "oleaut32" Alias "SafeArrayCreateVector" (Optional ByVal vt As VbVarType = vbDouble, Optional ByVal lLow As Long = 0, Optional ByVal lCount As Long = 0) As Double()
Private Declare Function ApiEmptyVariantArray Lib "oleaut32" Alias "SafeArrayCreateVector" (Optional ByVal vt As VbVarType = vbVariant, Optional ByVal lLow As Long = 0, Optional ByVal lCount As Long = 0) As Variant()
Private Declare Function IsTextUnicode Lib "advapi32" (lpBuffer As Any, ByVal cb As Long, lpi As Long) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function VariantChangeType Lib "oleaut32" (dest As Variant, src As Variant, ByVal wFlags As Integer, ByVal vt As Long) As Long
Private Declare Function VariantCopy Lib "oleaut32" (dest As Variant, src As Variant) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function GetSystemTimeAsFileTime Lib "kernel32.dll" (lpSystemTimeAsFileTime As Currency) As Long

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
    flags               As Long
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

Private Type BITMAPINFOHEADER
    biSize              As Long
    biWidth             As Long
    biHeight            As Long
    biPlanes            As Integer
    biBitCount          As Integer
    biCompression       As Long
    biSizeImage         As Long
    biXPelsPerMeter     As Long
    biYPelsPerMeter     As Long
    biClrUsed           As Long
    biClrImportant      As Long
End Type

Private Type RGBQUAD
    B                   As Byte
    G                   As Byte
    R                   As Byte
    A                   As Byte
End Type

Private Type SAFEARRAYBOUND
    cElements           As Long
    lLbound             As Long
End Type

Private Type SAFEARRAY2D
    cDims               As Integer
    fFeatures           As Integer
    cbElements          As Long
    cLocks              As Long
    pvData              As Long
    Bounds(0 To 1)      As SAFEARRAYBOUND
End Type

Private Type GDIPLUS_STARTUP_INPUT
    GdiplusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type

Private Type BMPFILE_HEADER
    filesz              As Long
    creator1            As Integer
    creator2            As Integer
    bmp_offset          As Long
End Type

Private Type GUID
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
    Source              As Long
    Description         As Long
    HelpFile            As Long
    dwHelpContext       As Long
    pvReserved          As Long
    pfnDeferredFillIn   As Long
    sCode               As Long
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Public Const LIB_NAME               As String = "UcsFiscalPrinters"
Public Const STR_NONE               As String = "(Íÿìà)"
Public Const STR_PROTOCOL_ELTRADE_ECR As String = "ELTRADE ECR"
Public Const STR_PROTOCOL_DATECS_FP As String = "DATECS FP/ECR"
Public Const STR_PROTOCOL_DAISY_ECR As String = "DAISY MICRO"
Public Const STR_PROTOCOL_ZEKA_FP   As String = "TREMOL ZEKA"
Public Const CHR1                   As String = "" '--- Chr$(1)
Public Const DBL_EPSILON            As Double = 0.0000000001

Private m_sDecimalSeparator     As String
Private m_hGdiPlus              As Long
Private m_oConfig               As Object
Private m_oPortWrapper          As cPortWrapper
Private m_nDebugLogFile         As Integer

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunc As String)
    Debug.Print MODULE_NAME & "." & sFunc & ": " & Err.Description
    OutputDebugLog MODULE_NAME, sFunc & "(" & Erl & ")", "Run-time error: " & Err.Description
End Sub

Private Sub RaiseError(sFunc As String)
    Debug.Print MODULE_NAME & "." & sFunc & ": " & Err.Description
    OutputDebugLog MODULE_NAME, sFunc & "(" & Erl & ")", "Run-time error: " & Err.Description
    Err.Raise Err.Number, MODULE_NAME & "." & sFunc & "(" & Erl & ")" & vbCrLf & Err.Source, Err.Description
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
        OutputDebugLog MODULE_NAME, FUNC_NAME, "Loading config file " & sFile
        If Not JsonParse(ReadTextFile(sFile), vJson, Error:=sError) Then
            OutputDebugLog MODULE_NAME, FUNC_NAME, "Error in config: " & sError
            Debug.Print "Error in config: " & sError
        End If
    End If
    If Not IsObject(vJson) Then
        JsonParse "{}", vJson
    End If
    Set m_oConfig = vJson
    Set m_oPortWrapper = New cPortWrapper
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Public Function At(vData As Variant, ByVal lIdx As Long, Optional sDefault As String) As String
    On Error GoTo QH
    At = sDefault
    If IsArray(vData) Then
        If LBound(vData) <= lIdx And lIdx <= UBound(vData) Then
            At = C_Str(vData(lIdx))
        End If
    End If
QH:
End Function

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
    ElseIf VariantChangeType(vDest, Value, 0, VT_I4) = 0 Then
        C_Lng = vDest
    End If
End Function

Public Function C_Str(Value As Variant) As String
    Dim vDest           As Variant
    
    If VarType(Value) = vbString Then
        C_Str = Value
    ElseIf VariantChangeType(vDest, Value, VARIANT_ALPHABOOL, VT_BSTR) = 0 Then
        C_Str = vDest
    End If
End Function

Public Function C_Bool(Value As Variant) As Boolean
    Dim vDest           As Variant
    
    If VarType(Value) = vbBoolean Then
        C_Bool = Value
    ElseIf VariantChangeType(vDest, Value, VARIANT_ALPHABOOL, VT_BOOL) = 0 Then
        C_Bool = vDest
    End If
End Function

Public Function C_Dbl(Value As Variant) As Double
    Dim vDest           As Variant
    
    If VarType(Value) = vbDouble Then
        C_Dbl = Value
    ElseIf VariantChangeType(vDest, Replace(C_Str(Value), ".", m_sDecimalSeparator), 0, VT_R8) = 0 Then
        C_Dbl = vDest
    End If
End Function

Public Function C_Date(Value As Variant) As Date
    Dim vDest           As Variant
    
    If VarType(Value) = vbDate Then
        C_Date = Value
    ElseIf VariantChangeType(vDest, Value, 0, VT_DATE) = 0 Then
        C_Date = vDest
    End If
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
    Dim udtVer          As OSVERSIONINFO
    
    udtVer.dwOSVersionInfoSize = Len(udtVer)
    If GetVersionEx(udtVer) Then
        If udtVer.dwPlatformID = VER_PLATFORM_WIN32_NT Then
            IsNT = True
        End If
    End If
End Function

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
        EnumSerialPorts = Split(vbNullString)
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
    Const LNG_MAX_SIZE  As Long = 10& * 1024 * 1024
    Dim vErr            As Variant
    Dim sFile           As String
    Dim sNewFile        As String
    
    vErr = Array(Err.Number, Err.Description, Err.Source)
    If m_nDebugLogFile = -1 Then
        Exit Sub
    End If
    On Error Resume Next '--- checked
    If m_nDebugLogFile = 0 Then
        sFile = Environ$("_UCS_FISCAL_PRINTER_LOG")
        If LenB(sFile) = 0 Then
            sFile = Environ$("TEMP") & "\UcsFP.log"
            If Not FileExists(sFile) Then
                m_nDebugLogFile = -1
                GoTo QH
            End If
        End If
        If FileExists(sFile) Then
            If FileLen(sFile) > LNG_MAX_SIZE Then
                If InStrRev(sFile, ".") > InStrRev(sFile, "\") Then
                    sNewFile = Left$(sFile, InStrRev(sFile, ".") - 1) & Format$(Date, "_yyyy_mm_dd") & Mid$(sFile, InStrRev(sFile, "."))
                Else
                    sNewFile = sFile & Format$(Date, "_yyyy_mm_dd")
                End If
                Name sFile As sNewFile
            End If
        End If
        m_nDebugLogFile = FreeFile
        Open sFile For Append Access Write Shared As #m_nDebugLogFile
    End If
    Print #m_nDebugLogFile, sModule & "." & sFunc & "(" & Now & Right$(Format$(DateTimer, "0.000"), 4) & "): " & sText
    If LOF(m_nDebugLogFile) > LNG_MAX_SIZE Then
        Close #m_nDebugLogFile
        m_nDebugLogFile = 0
    End If
QH:
    On Error GoTo 0
    Err.Number = vErr(0)
    Err.Description = vErr(1)
    Err.Source = vErr(2)
End Sub

Public Sub FlushDebugLog()
    Dim vErr            As Variant
    
    vErr = Array(Err.Number, Err.Description, Err.Source)
    On Error GoTo QH
    If m_nDebugLogFile <> 0 And m_nDebugLogFile <> -1 Then
        Close #m_nDebugLogFile
        m_nDebugLogFile = 0
    End If
QH:
    Err.Number = vErr(0)
    Err.Description = vErr(1)
    Err.Source = vErr(2)
End Sub

Public Function Round(ByVal Value As Double, Optional ByVal NumDigits As Long) As Double
    On Error GoTo QH
    Round = VBA.Round(Value + IIf(Value > 0, 10 ^ -13, -10 ^ -13), NumDigits)
QH:
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
    
    sAddress = Replace(sAddress, "¹", "N")
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
    If Left$(sRight, 1) = Chr$(1) Then
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
    CenterText = Space$(LimitLong((lWidth - Len(sText)) \ 2, 0)) & sText
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
                sBuffer = Left$(sBuffer, InStr(sBuffer, Chr$(0)) - 1)
                If lType = REG_EXPAND_SZ Then
                    RegReadString = String$(ExpandEnvironmentStrings(sBuffer, vbNullString, 0), 0)
                    If ExpandEnvironmentStrings(sBuffer, RegReadString, Len(RegReadString)) > 0 Then
                        RegReadString = Left$(RegReadString, InStr(RegReadString, Chr$(0)) - 1)
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
    GetSystemDirectory = Left$(GetSystemDirectory, InStr(GetSystemDirectory, Chr$(0)) - 1)
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
    uOFN.flags = OFN_LONGNAMES Or OFN_CREATEPROMPT Or OFN_HIDEREADONLY Or OFN_EXTENSIONDIFFERENT Or OFN_EXPLORER Or OFN_ENABLESIZING
    uOFN.hWndOwner = hWndOwner
    uOFN.lpstrFilter = VarPtr(baFilter(0))
    uOFN.nFilterIndex = 1
    uOFN.lpstrTitle = VarPtr(baTitle(0))
    uOFN.lpstrFile = StrPtr(sBuffer)
    uOFN.nMaxFile = Len(sBuffer)
    If GetOpenFileName(uOFN) <> 0 Then
        sFile = StrConv(sBuffer, vbUnicode)
        sFile = Left$(sFile, InStr(sFile, Chr$(0)) - 1)
        '--- success
        OpenSaveDialog = True
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Public Function ConvertToBW( _
            ByVal hBitmap As Long, _
            ByVal lWidth As Long, _
            ByVal lHeight As Long, _
            ByVal lThreshold As Long, _
            ByVal bCenter As Boolean) As Byte()
    Const FUNC_NAME     As String = "ConvertToBW"
    Dim uBIH            As BITMAPINFOHEADER
    Dim hDC             As Long
    Dim hDIB            As Long
    Dim lpBits          As Long
    Dim hOldDIB         As Long
    Dim uSA             As SAFEARRAY2D
    Dim aBitsRGB()      As RGBQUAD
    Dim hGraphics       As Long
    Dim uHdr            As BMPFILE_HEADER
    Dim baRetVal()      As Byte
    Dim lX              As Long
    Dim lY              As Long
    Dim lOffset         As Long
    Dim lLum            As Long
    Dim lScanline       As Long
    Dim sngWidth        As Single
    Dim sngHeight       As Single
    Dim hBrush          As Long

    On Error GoTo EH
    hDC = CreateCompatibleDC(0)
    If hDC <> 0 Then
        With uBIH
            .biSize = Len(uBIH)
            .biPlanes = 1
            .biBitCount = 32
            .biWidth = lWidth
            .biHeight = -lHeight
            .biSizeImage = (4 * lWidth) * lHeight
        End With
        hDIB = CreateDIBSection(hDC, uBIH, DIB_RGB_COLORS, lpBits, 0, 0)
        If hDIB <> 0 Then
            hOldDIB = SelectObject(hDC, hDIB)
            With uSA
                .cbElements = 4
                .cDims = 2
                .Bounds(0).lLbound = 0
                .Bounds(0).cElements = lHeight
                .Bounds(1).lLbound = 0
                .Bounds(1).cElements = lWidth
                .pvData = lpBits
            End With
            Call CopyMemory(ByVal ArrPtr(aBitsRGB()), VarPtr(uSA), 4)
            '--- stretch bitmap to DIB
            If GdipCreateFromHDC(hDC, hGraphics) = 0 Then
                If bCenter Then
                    If GdipCreateSolidFill(&HFFFFFFFF, hBrush) = 0 Then
                        Call GdipFillRectangleI(hGraphics, hBrush, 0, 0, lWidth, lHeight)
                        Call GdipDeleteBrush(hBrush)
                    End If
                    Call GdipGetImageDimension(hBitmap, sngWidth, sngHeight)
                    Call GdipDrawImageRectI(hGraphics, hBitmap, (lWidth - sngWidth) \ 2, (lHeight - sngHeight) \ 2, sngWidth, sngHeight)
                Else
                    Call GdipDrawImageRectI(hGraphics, hBitmap, 0, 0, lWidth, lHeight)
                End If
                Call GdipDeleteGraphics(hGraphics)
                '--- prepare headers
                lScanline = ((lWidth + 31) \ 32) * 4
                With uHdr
                    .bmp_offset = 2 + Len(uHdr) + Len(uBIH) + 2 * 4
                    .filesz = .bmp_offset + lScanline * lHeight
                End With
                With uBIH
                    .biSize = Len(uBIH)
                    .biPlanes = 1
                    .biBitCount = 1
                    .biWidth = lWidth
                    .biHeight = lHeight
                    .biSizeImage = lScanline * lHeight
                    .biClrUsed = 2
                End With
                ReDim baRetVal(0 To uHdr.filesz - 1)
                Call CopyMemory(baRetVal(0), &H4D42, 2) '--- BM
                Call CopyMemory(baRetVal(2), uHdr, Len(uHdr))
                Call CopyMemory(baRetVal(2 + Len(uHdr)), uBIH, Len(uBIH))
                '--- color palette
                lX = &HFFFFFF
                Call CopyMemory(baRetVal(2 + Len(uHdr) + Len(uBIH) + 4), lX, 4)
                '--- calc luminance and set bits
                For lY = 0 To lHeight - 1
                    lOffset = uHdr.bmp_offset + lScanline * (lHeight - lY - 1)
                    For lX = 0 To lWidth - 1
                        With aBitsRGB(lX, lY)
                            lLum = .R * 0.299 + .G * 0.587 + .B * 0.114
                        End With
                        If lLum >= lThreshold Then
                            baRetVal(lOffset + lX \ 8) = baRetVal(lOffset + lX \ 8) Or 2 ^ (7 - (lX Mod 8))
                        End If
                    Next
                Next
            End If
            '--- cleanup
            Call CopyMemory(ByVal ArrPtr(aBitsRGB()), 0&, 4)
            Call SelectObject(hDC, hOldDIB)
            Call DeleteObject(hDIB)
        End If
        Call DeleteObject(hDC)
    End If
    ConvertToBW = baRetVal
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Public Function GdipLoadImage(sFile As String) As Long
    Dim uStartup        As GDIPLUS_STARTUP_INPUT
    
    If m_hGdiPlus = 0 Then
        uStartup.GdiplusVersion = 1&
        Call GdiplusStartup(m_hGdiPlus, uStartup)
    End If
    Call GdipLoadImageFromFile(StrPtr(sFile), GdipLoadImage)
End Function

Public Sub GdipReleaseImage(hBitmap As Long)
    Call GdipDisposeImage(hBitmap)
End Sub

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

Public Function EmptyVariantArray() As Variant()
    EmptyVariantArray = ApiEmptyVariantArray()
End Function

Public Function FileExists(sFile As String) As Boolean
    If GetFileAttributes(sFile) = -1 Then ' INVALID_FILE_ATTRIBUTES
    Else
        FileExists = True
    End If
End Function

Public Function ReadTextFile(sFile As String) As String
    Const BOM_UTF       As String = "ï»¿" '--- "\xEF\xBB\xBF"
    Const BOM_UNICODE   As String = "ÿþ"  '--- "\xFF\xFE"
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

Public Function GetConfigValue(sSerial As String, sKey As String, Optional vDefault As Variant) As Variant
    Const FUNC_NAME     As String = "GetConfigValue"
    
    On Error GoTo EH
    If LenB(sSerial) <> 0 Then
        If m_oConfig.Exists(sSerial) Then
            If IsObject(m_oConfig(sSerial)) Then
                If m_oConfig(sSerial).Exists(sKey) Then
                    AssignVariant GetConfigValue, m_oConfig(sSerial).Item(sKey)
                    Exit Function
                End If
            End If
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
    If dblDefault > 0 Then
        If GetConfigNumber <= 0 Then
            GetConfigNumber = dblDefault
        End If
    ElseIf dblDefault < 0 Then
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
        Set oDict = vValue
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

Public Function preg_replace(sPattern As String, sReplace As String, sText As String) As String
    preg_replace = pvInitRegExp(sPattern).Replace(sText, sReplace)
End Function

Private Function pvInitRegExp(sPattern As String) As Object
    Dim lIdx            As Long
    
    Set pvInitRegExp = CreateObject("VBScript.RegExp")
    With pvInitRegExp
        .Global = True
        If Left$(sPattern, 1) = "/" Then
            lIdx = InStrRev(sPattern, "/")
            .Pattern = Mid$(sPattern, 2, lIdx - 2)
            .IgnoreCase = (InStr(lIdx, sPattern, "i") > 0)
            .MultiLine = (InStr(lIdx, sPattern, "m") > 0)
        Else
            .Pattern = sPattern
        End If
    End With
End Function

Public Function GetConfigForCommand(oConfigCmd As Collection, oLocalizedCmd As Collection, sFunc As String, sKey As String, Optional Default As Variant) As Variant
    Dim sMerged         As String
    Dim vItem           As Variant
    
    sMerged = "\" & sFunc & IIf(LenB(sKey) <> 0, "\" & sKey, vbNullString)
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

Public Property Get DateTimer() As Double
    Dim cDateTime       As Currency
    
    Call GetSystemTimeAsFileTime(cDateTime)
    DateTimer = CDbl(cDateTime - 9435304800000@) / 1000#
End Property

Public Function ToHexDump(sText As String) As String
    Dim lIdx            As Long
    
    For lIdx = 1 To Len(sText)
        ToHexDump = ToHexDump & Right$("0" & Hex(Asc(Mid$(sText, lIdx, 1))), 2)
    Next
End Function

Public Function SearchCollection(ByVal pCol As Object, Index As Variant, Optional RetVal As Variant) As Boolean
    Const DISPID_VALUE  As Long = 0
    Dim pVbCol          As IVbCollection
    
    If pCol Is Nothing Then
        '--- do nothing
    ElseIf TypeOf pCol Is IVbCollection Then
        Set pVbCol = pCol
        SearchCollection = pVbCol.Item(Index, RetVal) = 0
    Else
        SearchCollection = DispInvoke(pCol, DISPID_VALUE, UcsIclPropGet, Result:=RetVal, Args:=Index)
        If Not SearchCollection Then
            '--- some weird collections have default (Item) method
            SearchCollection = DispInvoke(pCol, DISPID_VALUE, UcsIclMethod, Result:=RetVal, Args:=Index)
        End If
    End If
End Function

Public Function DispInvoke( _
            ByVal pDisp As IVbDispatch, _
            Name As Variant, _
            Optional ByVal CallType As UcsInvokeCallEnum = UcsIclMethod, _
            Optional Result As Variant, _
            Optional Args As Variant) As Boolean
    Const DISPID_PROPERTYPUT As Long = -3
    Dim IID_NULL        As GUID
    Dim lDispID         As Long
    Dim uParams         As DISPPARAMS
    Dim uInfo           As EXCEPINFO
    Dim aParams()       As Variant
    Dim lNamedParam     As Long
    Dim lIdx            As Long
    Dim lParamCount     As Long
    Dim lArgErr         As Long
    Dim lPtrResult      As Long

    If pDisp Is Nothing Then
        Exit Function
    End If
    '--- get disp id
    If IsNumeric(Name) Then
        lDispID = C_Lng(Name)
    Else
        If pDisp.GetIDsOfNames(IID_NULL, C_Str(Name), 1, LOCALE_SYSTEM_DEFAULT, lDispID) <> S_OK Then
            Exit Function
        End If
    End If
    '--- process params
    If Not IsMissing(Args) Then
        If IsArray(Args) Then
            lParamCount = UBound(Args) - LBound(Args)
            ReDim aParams(0 To lParamCount) As Variant
            For lIdx = 0 To lParamCount
                Call VariantCopy(aParams(lParamCount - lIdx), Args(lIdx))
            Next
        Else
            ReDim aParams(0 To 0) As Variant
            Call VariantCopy(aParams(0), Args)
        End If
        With uParams
            .cArgs = lParamCount + 1
            .rgPointerToVariantArray = VarPtr(aParams(0))
        End With
        If CallType = UcsIclPropLet Or CallType = UcsIclPropSet Then
            lNamedParam = DISPID_PROPERTYPUT
            With uParams
                .cNamedArgs = 1
                .rgPointerToLongNamedArgs = VarPtr(lNamedParam)
            End With
        End If
    End If
    If CallType = UcsIclPropGet Or CallType = UcsIclMethod And Not IsMissing(Result) Then
        Result = Empty
        lPtrResult = VarPtr(Result)
    End If
    DispInvoke = (pDisp.Invoke(lDispID, IID_NULL, LOCALE_SYSTEM_DEFAULT, CallType, uParams, ByVal lPtrResult, uInfo, lArgErr) = S_OK)
End Function

Public Function DispPropertyGet(pDisp As Object, PropName As String, Optional Result As Variant) As Variant
    If DispInvoke(pDisp, PropName, UcsIclPropGet, Result) Then
        AssignVariant DispPropertyGet, Result
    End If
End Function

Public Property Get LockControl(oCtl As Object) As Boolean
    Dim vResult         As Variant
    
    If DispInvoke(oCtl, "Locked", UcsIclPropGet, vResult) Then
        LockControl = vResult
    ElseIf DispInvoke(oCtl, "Enabled", UcsIclPropGet, vResult) Then
        LockControl = Not vResult
    End If
End Property

Public Property Let LockControl(oCtl As Object, ByVal bValue As Boolean)
    If DispInvoke(oCtl, "Locked", UcsIclPropLet, Args:=bValue) Then
        DispInvoke oCtl, "BackColor", UcsIclPropLet, Args:=IIf(bValue, vbButtonFace, vbWindowBackground)
    Else
        DispInvoke oCtl, "Enabled", UcsIclPropLet, Args:=Not bValue
    End If
End Property

Public Function ParseSum(sValue As String) As Double
    If InStr(sValue, ".") > 0 Then
        ParseSum = C_Dbl(sValue)
    Else
        ParseSum = C_Dbl(sValue) / 100#
    End If
End Function

Public Function ToAscii(sSend As String) As Byte()
    Dim lSize           As Long
    Dim baText()        As Byte
    
    lSize = Len(sSend)
    If lSize > 0 Then
        ReDim baText(0 To lSize - 1) As Byte
        Call WideCharToMultiByte(0, 0, StrPtr(sSend), lSize, baText(0), Len(sSend), 0, 0)
    Else
        baText = " "
    End If
    ToAscii = baText
End Function

