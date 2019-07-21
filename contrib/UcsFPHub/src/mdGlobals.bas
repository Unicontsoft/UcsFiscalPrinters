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
' API
'=========================================================================

'--- for VariantChangeType
Private Const VARIANT_ALPHABOOL             As Long = 2

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function CommandLineToArgvW Lib "shell32" (ByVal lpCmdLine As Long, pNumArgs As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function ApiSysAllocString Lib "oleaut32" Alias "SysAllocString" (ByVal Ptr As Long) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function VariantChangeType Lib "oleaut32" (Dest As Variant, Src As Variant, ByVal wFlags As Integer, ByVal vt As VbVarType) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

'=========================================================================
' Constants and member variables
'=========================================================================

Private m_cRegExpCache              As Collection

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunction As String)
    Debug.Print "Critical error: " & Err.Description & " [" & MODULE_NAME & "." & sFunction & "]"
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

Public Function ReadBinaryFile(sFile As String) As Byte()
    With CreateObject("ADODB.Stream")
        .Open
        .Type = 1
        .LoadFromFile sFile
        ReadBinaryFile = .Read
    End With
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

Public Function FromUtf8Array(baData() As Byte) As String
    With New cAsyncSocket
        FromUtf8Array = .FromTextArray(baData)
    End With
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
                        GoTo Conitnue
                    End If
                Next
                .Item("-" & Mid$(At(vArgs, lIdx), 2)) = True
            Case Else
                .Item("numarg") = .Item("numarg") + 1
                .Item("arg" & .Item("numarg")) = At(vArgs, lIdx)
            End Select
Conitnue:
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
                    If .Submatches.Count = 0 Then
                        ReDim Matches(0 To 0) As String
                        Matches(0) = .Value
                    Else
                        ReDim Matches(0 To .Submatches.Count - 1) As String
                        For lIdx = 0 To .Submatches.Count - 1
                            Matches(lIdx) = .Submatches(lIdx)
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
