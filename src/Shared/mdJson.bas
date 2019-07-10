Attribute VB_Name = "mdJson"
'=========================================================================
'
' UcsFP20 (c) 2008-2019 by wqweto@gmail.com
'
' Unicontsoft Fiscal Printers Component 2.0
'
' This project is licensed under the terms of the MIT license
' See the LICENSE file in the project root for more information
'
'=========================================================================
'
' JSON parsing and dumping functions
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "mdJson"

#Const ImplScripting = JSON_USE_SCRIPTING <> 0
#Const ImplUseShared = DebugMode <> 0

#Const HasPtrSafe = (VBA7 <> 0)
#Const LargeAddressAware = (Win64 = 0 And VBA7 = 0 And VBA6 = 0 And VBA5 = 0)

'=========================================================================
' API
'=========================================================================

#If HasPtrSafe Then
    Private Const NULL_PTR                  As LongPtr = 0
#Else
    Private Const NULL_PTR                  As Long = 0
#End If
#If Win64 Then
    Private Const PTR_SIZE                  As Long = 8
#Else
    Private Const PTR_SIZE                  As Long = 4
    Private Const SIGN_BIT                  As Long = &H80000000
#End If
'--- for CompareStringW
Private Const LOCALE_USER_DEFAULT           As Long = &H400
Private Const NORM_IGNORECASE               As Long = 1
Private Const CSTR_LESS_THAN                As Long = 1
Private Const CSTR_EQUAL                    As Long = 2
Private Const CSTR_GREATER_THAN             As Long = 3
'--- for VariantChangeType
Private Const VARIANT_ALPHABOOL             As Long = 2

#If HasPtrSafe Then
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
    Private Declare PtrSafe Function ArrPtr Lib "vbe7" Alias "VarPtr" (Ptr() As Any) As LongPtr
    Private Declare PtrSafe Function CompareStringW Lib "kernel32" (ByVal Locale As Long, ByVal dwCmpFlags As Long, lpString1 As Any, ByVal cchCount1 As Long, lpString2 As Any, ByVal cchCount2 As Long) As Long
    Private Declare PtrSafe Function VariantChangeType Lib "oleaut32" (Dest As Variant, Src As Variant, ByVal wFlags As Integer, ByVal vt As VbVarType) As Long
#Else
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
    Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
    Private Declare Function CompareStringW Lib "kernel32" (ByVal Locale As Long, ByVal dwCmpFlags As Long, lpString1 As Any, ByVal cchCount1 As Long, lpString2 As Any, ByVal cchCount2 As Long) As Long
    Private Declare Function VariantChangeType Lib "oleaut32" (Dest As Variant, Src As Variant, ByVal wFlags As Integer, ByVal vt As VbVarType) As Long
#End If

Private Const STR_PREFIX            As String = "__json__"
Private Const STR_ATTR_EMPTY        As String = STR_PREFIX & "empty"
Private Const STR_ATTR_ARRAY        As String = STR_PREFIX & "array"
Private Const STR_ATTR_NAME         As String = STR_PREFIX & "name"
Private Const STR_ATTR_NIL          As String = STR_PREFIX & "nil"
Private Const STR_ATTR_BOOL         As String = STR_PREFIX & "bool"
Private Const STR_NODE_ARRAY        As String = STR_PREFIX & "array"
Private Const ERR_EXTRA_SYMBOL      As String = "Extra '%1' found at position %2"
Private Const ERR_DUPLICATE_KEY     As String = "Duplicate key '%1' found at position %2"
Private Const ERR_UNEXPECTED_SYMBOL As String = "Unexpected '%1' at position %2"
Private Const ERR_CONVERSION        As String = "%1 at position %2"
Private Const ERR_UNTERMIN_COMMENT  As String = "Unterminated comment at position %1"
Private Const ERR_MISSING_EOS       As String = "Missing end of string at position %1"
Private Const ERR_INVALID_ESCAPE    As String = "Invalid escape at position %1"
Private Const ERR_MISSING_KEY       As String = "Missing key at position %1"
Private Const ERR_EXPECTED_SYMBOL   As String = "Expected '%1' at position %2"
Private Const ERR_EXPECTED_TWO      As String = "Expected '%1' or '%2' at position %3"
Private Const DEF_IGNORE_CASE       As Boolean = True
#If ImplScripting Then
    Private Const IDX_OFFSET        As Long = 0
#Else
    Private Const IDX_OFFSET        As Long = 1
#End If

Private Type SAFEARRAY1D
    cDims               As Integer
    fFeatures           As Integer
    cbElements          As Long
    cLocks              As Long
    #If HasPtrSafe Then
        pvData          As LongPtr
    #Else
        pvData          As Long
    #End If
    cElements           As Long
    lLbound             As Long
End Type

Private Type JsonContext
    StrictMode          As Boolean
    Text()              As Integer
    Pos                 As Long
    Error               As String
    LastChar            As Integer
    TextArray           As SAFEARRAY1D
End Type

#If Win64 Then
    Private Enum VbCollectionOffsets
        o_pFirstIndexedItem = &H28
        o_pRootTreeItem = &H40
        o_pEndTreePtr = &H48
        o_pvUnk5 = &H50
        '--- item
        o_KeyPtr = &H18
        o_pNextIndexedItem = o_pFirstIndexedItem '--- Coincidence?
        o_pRightBranch = &H38
        o_pLeftBranch = &H40
    End Enum
#Else
    Private Enum VbCollectionOffsets
        o_pFirstIndexedItem = &H18
        o_pRootTreeItem = &H24
        o_pEndTreePtr = &H28
        o_pvUnk5 = &H2C
        '--- item
        o_KeyPtr = &H10
        o_pNextIndexedItem = o_pFirstIndexedItem '--- Again?
        o_pRightBranch = &H24
        o_pLeftBranch = &H28
    End Enum
#End If

'=========================================================================
' Error management
'=========================================================================

Private Function RaiseError(sFunction As String) As VbMsgBoxResult
    #If ImplUseShared Then
        Dim vErr            As Variant
        
        PushError vErr
        RaiseError = GApp.HandleOutOfMemory(vErr)
        If RaiseError <> vbRetry Then
            PopRaiseError sFunction, MODULE_NAME, vErr
        End If
    #Else
        Err.Raise Err.Number, MODULE_NAME & "." & sFunction & vbCrLf & Err.Source, Err.Description
    #End If
End Function

Private Function PrintError(sFunction As String) As VbMsgBoxResult
    #If ImplUseShared Then
        Dim vErr            As Variant
        
        PushError vErr
        PrintError = GApp.HandleOutOfMemory(vErr)
        If PrintError <> vbRetry Then
            PopPrintError sFunction, MODULE_NAME, vErr
        End If
    #Else
        Debug.Print "Critical error: " & Err.Description & " [" & MODULE_NAME & "." & sFunction & "]"
        Debug.Assert MsgBox(Err.Description, vbCritical, MODULE_NAME & "." & sFunction) <> -1
    #End If
End Function

'=========================================================================
' Functions
'=========================================================================

Public Function JsonParse( _
            sText As String, _
            Optional RetVal As Variant, _
            Optional Error As String, _
            Optional ByVal StrictMode As Boolean) As Boolean
    Const FUNC_NAME     As String = "JsonParse"
    Dim uCtx            As JsonContext

    On Error GoTo EH
    With uCtx
        .StrictMode = StrictMode
        '--- map array over input string
        With .TextArray
            .cDims = 1
            .fFeatures = 1 ' FADF_AUTO
            .cbElements = 2
            .cLocks = 1
            .pvData = StrPtr(sText)
            If .pvData = 0 Then
                .pvData = StrPtr("")
            End If
            .cElements = Len(sText) + 1
        End With
        Call CopyMemory(ByVal ArrPtr(.Text), VarPtr(.TextArray), PTR_SIZE)
        AssignVariant RetVal, pvJsonParse(uCtx)
        If LenB(.Error) Then
            Error = .Error
            GoTo QH
        End If
        If pvJsonGetChar(uCtx) <> 0 Then
            Error = Printf(ERR_EXTRA_SYMBOL, ChrW$(.LastChar), .Pos)
            GoTo QH
        End If
        '--- success
        JsonParse = True
QH:
        .TextArray.pvData = 0
        .TextArray.cElements = 0
    End With
    Exit Function
EH:
    If PrintError(FUNC_NAME) = vbRetry Then
        Resume
    End If
    Resume Next
End Function

Private Function pvJsonParse(uCtx As JsonContext) As Variant
    Const FUNC_NAME     As String = "pvJsonParse"
    Dim lIdx            As Long
    Dim sKey            As String
    Dim sText           As String
    Dim vValue          As Variant
    #If ImplScripting Then
        Dim oRetVal     As Scripting.Dictionary
    #Else
        Dim oRetVal     As VBA.Collection
    #End If
    
    On Error GoTo EH
    With uCtx
        Select Case pvJsonGetChar(uCtx)
        Case 34 '--- "
            pvJsonParse = pvJsonGetString(uCtx)
            If .LastChar = 0 Then
                GoTo QH
            End If
        Case 91 '--- [
            Set oRetVal = pvJsonCreateObject(vbTextCompare)
            Do
                AssignVariant vValue, pvJsonParse(uCtx)
                If LenB(.Error) <> 0 Then
                    If .LastChar = 93 Then '--- ]
                        If Not .StrictMode Then
                            Exit Do
                        End If
                        lIdx = oRetVal.Count
                        If lIdx = 0 Then
                            Exit Do
                        End If
                    End If
                    GoTo QH
                End If
                #If ImplScripting Then
                    oRetVal.Add lIdx, vValue
                #Else
                    oRetVal.Add vValue
                #End If
                Select Case pvJsonGetChar(uCtx)
                Case 44 '--- ,
                    lIdx = lIdx + 1
                Case 93 '--- ]
                    Exit Do
                Case Else
                    .Error = Printf(ERR_EXPECTED_TWO, ",", "]", .Pos)
                    Exit Function
                End Select
            Loop
            .Error = vbNullString
            Set pvJsonParse = oRetVal
        Case 123 '--- {
            Set oRetVal = pvJsonCreateObject(vbBinaryCompare)
            Do
                If pvJsonGetChar(uCtx) <> 34 Then '--- "
                    If .LastChar = 125 Then '--- }
                        If Not .StrictMode Then
                            Exit Do
                        End If
                        lIdx = oRetVal.Count
                        If lIdx = 0 Then
                            Exit Do
                        End If
                    End If
                    .Error = Printf(ERR_MISSING_KEY, .Pos)
                    GoTo QH
                End If
                sKey = pvJsonGetString(uCtx)
                If .LastChar = 0 Then
                    GoTo QH
                End If
                If pvJsonGetChar(uCtx) <> 58 Then '--- :
                    .Error = Printf(ERR_EXPECTED_SYMBOL, ":", .Pos)
                    GoTo QH
                End If
                AssignVariant vValue, pvJsonParse(uCtx)
                If LenB(.Error) <> 0 Then
                    GoTo QH
                End If
                Select Case pvJsonGetChar(uCtx)
                Case 44, 125 '--- , }
                    #If ImplScripting Then
                        If oRetVal.Exists(sKey) Then
                    #Else
                        If CollectionIndexByKey(oRetVal, sKey, DEF_IGNORE_CASE) > 0 Then
                    #End If
                        .Error = Printf(ERR_DUPLICATE_KEY, sKey, .Pos)
                        GoTo QH
                    End If
                    #If ImplScripting Then
                        oRetVal.Add sKey, vValue
                    #Else
                        oRetVal.Add vValue, sKey
                    #End If
                    If .LastChar = 125 Then '--- }
                        Exit Do
                    End If
                Case Else
                    .Error = Printf(ERR_EXPECTED_TWO, ",", "}", .Pos)
                    GoTo QH
                End Select
            Loop
            .Error = vbNullString
            Set pvJsonParse = oRetVal
        Case 116, 84  '--- "t", "T"
            If (.Text(.Pos + 0) Or &H20) <> 114 Then    '--- r
                GoTo UnexpectedSymbol
            End If
            If (.Text(.Pos + 1) Or &H20) <> 117 Then    '--- u
                GoTo UnexpectedSymbol
            End If
            If (.Text(.Pos + 2) Or &H20) <> 101 Then    '--- e
                GoTo UnexpectedSymbol
            End If
            .Pos = .Pos + 3
            pvJsonParse = True
        Case 102, 70 '--- "f", "F"
            If (.Text(.Pos + 0) Or &H20) <> 97 Then     '--- a
                GoTo UnexpectedSymbol
            End If
            If (.Text(.Pos + 1) Or &H20) <> 108 Then    '--- l
                GoTo UnexpectedSymbol
            End If
            If (.Text(.Pos + 2) Or &H20) <> 115 Then    '--- s
                GoTo UnexpectedSymbol
            End If
            If (.Text(.Pos + 3) Or &H20) <> 101 Then    '--- e
                GoTo UnexpectedSymbol
            End If
            .Pos = .Pos + 4
            pvJsonParse = False
        Case 110, 78 '--- "n", "N"
            If (.Text(.Pos + 0) Or &H20) <> 117 Then    '--- u
                GoTo UnexpectedSymbol
            End If
            If (.Text(.Pos + 1) Or &H20) <> 108 Then    '--- l
                GoTo UnexpectedSymbol
            End If
            If (.Text(.Pos + 2) Or &H20) <> 108 Then    '--- l
                GoTo UnexpectedSymbol
            End If
            .Pos = .Pos + 3
            pvJsonParse = Null
        Case 48 To 57, 43, 45, 46 '--- 0-9 + - .
            For lIdx = 0 To 1000
                Select Case .Text(.Pos + lIdx)
                Case 48 To 57, 43, 45, 46, 101, 69, 120, 88, 97 To 102, 65 To 70 '--- 0-9 + - . e E x X a-f A-F
                Case Else
                    Exit For
                End Select
            Next
            sText = Space$(lIdx + 1)
            Call CopyMemory(ByVal StrPtr(sText), .Text(.Pos - 1), LenB(sText))
            If LCase$(Left$(sText, 2)) = "0x" Then
                Mid$(sText, 1, 2) = "&H"
            End If
            On Error GoTo ErrorConvert
            pvJsonParse = CDbl(sText)
            On Error GoTo 0
            .Pos = .Pos + lIdx
        Case 0
            If LenB(.Error) <> 0 Then
                GoTo QH
            End If
        Case Else
            GoTo UnexpectedSymbol
        End Select
QH:
        Exit Function
UnexpectedSymbol:
        .Error = Printf(ERR_UNEXPECTED_SYMBOL, ChrW$(.LastChar), .Pos)
        Exit Function
ErrorConvert:
        .Error = Printf(ERR_CONVERSION, Err.Description, .Pos)
    End With
    Exit Function
EH:
    If RaiseError(FUNC_NAME) = vbRetry Then
        Resume
    End If
End Function

Private Function pvJsonGetChar(uCtx As JsonContext) As Integer
    Const FUNC_NAME     As String = "pvJsonGetChar"
    Dim lIdx            As Long
    
    On Error GoTo EH
    With uCtx
        Do While .Pos <= UBound(.Text)
            .LastChar = .Text(.Pos)
            .Pos = .Pos + 1
            Select Case .LastChar
            Case 0
                Exit Function
            Case 9, 10, 13, 32 '--- vbTab, vbCr, vbLf, " "
                '--- do nothing
            Case 47 '--- /
                If Not .StrictMode Then
                    Select Case .Text(.Pos)
                    Case 47 '--- //
                        .Pos = .Pos + 1
                        Do
                            .LastChar = .Text(.Pos)
                            .Pos = .Pos + 1
                            If .LastChar = 0 Then
                                Exit Function
                            End If
                        Loop While Not (.LastChar = 10 Or .LastChar = 13)  '--- vbLf or vbCr
                    Case 42 '--- /*
                        lIdx = .Pos + 1
                        Do
                            .LastChar = .Text(lIdx)
                            lIdx = lIdx + 1
                            If .LastChar = 0 Then
                                .Error = Printf(ERR_UNTERMIN_COMMENT, .Pos)
                                Exit Function
                            End If
                        Loop While Not (.LastChar = 42 And .Text(lIdx) = 47) '--- */
                        .LastChar = .Text(lIdx)
                        .Pos = lIdx + 1
                    Case Else
                        pvJsonGetChar = .LastChar
                        Exit Do
                    End Select
                Else
                    pvJsonGetChar = .LastChar
                    Exit Do
                End If
            Case Else
                pvJsonGetChar = .LastChar
                Exit Do
            End Select
        Loop
    End With
    Exit Function
EH:
    If RaiseError(FUNC_NAME) = vbRetry Then
        Resume
    End If
End Function

Private Function pvJsonGetString(uCtx As JsonContext) As String
    Const FUNC_NAME     As String = "pvJsonGetString"
    Dim lIdx            As Long
    Dim nChar           As Integer
    Dim sText           As String
    
    On Error GoTo EH
    With uCtx
        For lIdx = 0 To &H7FFFFFFF
            nChar = .Text(.Pos + lIdx)
            Select Case nChar
            Case 0, 34, 92 '--- " \
                sText = Space$(lIdx)
                Call CopyMemory(ByVal StrPtr(sText), .Text(.Pos), LenB(sText))
                pvJsonGetString = pvJsonGetString & sText
                If nChar = 34 Then '--- "
                    .Pos = .Pos + lIdx + 1
                    Exit For
                ElseIf nChar <> 92 Then '--- \
                    nChar = 0
                    .Pos = .Pos + lIdx + 1
                    .Error = Printf(ERR_MISSING_EOS, .Pos)
                    Exit For
                End If
                lIdx = lIdx + 1
                nChar = .Text(.Pos + lIdx)
                Select Case nChar
                Case 98  '--- b
                    pvJsonGetString = pvJsonGetString & ChrW$(8)
                Case 102 '--- f
                    pvJsonGetString = pvJsonGetString & ChrW$(12)
                Case 110 '--- n
                    pvJsonGetString = pvJsonGetString & vbLf
                Case 114 '--- r
                    pvJsonGetString = pvJsonGetString & vbCr
                Case 116 '--- t
                    pvJsonGetString = pvJsonGetString & vbTab
                Case 34  '--- "
                    pvJsonGetString = pvJsonGetString & """"
                Case 92  '--- \
                    pvJsonGetString = pvJsonGetString & "\"
                Case 47  '--- /
                    pvJsonGetString = pvJsonGetString & "/"
                Case 117 '--- u
                    pvJsonGetString = pvJsonGetString & ChrW$(CLng("&H" & ChrW$(.Text(.Pos + lIdx + 1)) & ChrW$(.Text(.Pos + lIdx + 2)) & ChrW$(.Text(.Pos + lIdx + 3)) & ChrW$(.Text(.Pos + lIdx + 4))))
                    lIdx = lIdx + 4
                Case 120 '--- x
                    pvJsonGetString = pvJsonGetString & ChrW$(CLng("&H" & ChrW$(.Text(.Pos + lIdx + 1)) & ChrW$(.Text(.Pos + lIdx + 2))))
                    lIdx = lIdx + 2
                Case Else
                    nChar = 0
                    .Pos = .Pos + lIdx + 1
                    .Error = Printf(ERR_INVALID_ESCAPE, .Pos)
                    Exit For
                End Select
                .Pos = .Pos + lIdx + 1
                lIdx = -1
            End Select
        Next
        .LastChar = nChar
    End With
    Exit Function
EH:
    If RaiseError(FUNC_NAME) = vbRetry Then
        Resume
    End If
End Function

Public Function JsonDump(vJson As Variant, Optional ByVal Level As Long, Optional ByVal Minimize As Boolean) As String
    Const STR_CODES     As String = "\u0000|\u0001|\u0002|\u0003|\u0004|\u0005|\u0006|\u0007|\b|\t|\n|\u000B|\f|\r|\u000E|\u000F|\u0010|\u0011|" & _
                                    "\u0012|\u0013|\u0014|\u0015|\u0016|\u0017|\u0018|\u0019|\u001A|\u001B|\u001C|\u001D|\u001E|\u001F"
    Const LNG_INDENT    As Long = 4
    Static vTranscode   As Variant
    Dim vKeys           As Variant
    Dim vItems          As Variant
    Dim lIdx            As Long
    Dim lSize           As Long
    Dim sCompound       As String
    Dim sSpace          As String
    Dim lAsc            As Long
    Dim lCompareMode    As VbCompareMethod
    Dim lCount          As Long
    #If ImplScripting Then
        Dim oJson       As Scripting.Dictionary
    #Else
        Dim oJson       As VBA.Collection
    #End If
    
    '--- note: skip error handling not to clear Err because used in error handlers
'    On Error GoTo EH
    Select Case VarType(vJson)
    Case vbObject
        Set oJson = vJson
        If oJson Is Nothing Then
            Exit Function
        End If
        lCompareMode = pvJsonCompareMode(oJson)
        sCompound = IIf(lCompareMode = vbBinaryCompare, "{}", "[]")
        lCount = oJson.Count
        If lCount <= 0 Then
            JsonDump = sCompound
        Else
            sSpace = IIf(Minimize, vbNullString, " ")
            ReDim vItems(0 To lCount - 1) As String
            If lCompareMode = vbBinaryCompare Then
                #If ImplScripting Then
                    vKeys = oJson.Keys
                #Else
                    vKeys = CollectionAllKeys(oJson)
                #End If
            End If
            For lIdx = 0 To lCount - 1
                If lCompareMode = vbBinaryCompare Then
                    vItems(lIdx) = JsonDump(vKeys(lIdx)) & ":" & sSpace & JsonDump(oJson.Item(vKeys(lIdx)), Level + 1, Minimize)
                Else
                    vItems(lIdx) = JsonDump(oJson.Item(lIdx + IDX_OFFSET), Level + 1, Minimize)
                End If
                lSize = lSize + Len(vItems(lIdx))
            Next
            If lSize > 100 And Not Minimize Then
                JsonDump = Left$(sCompound, 1) & vbCrLf & _
                    Space$((Level + 1) * LNG_INDENT) & Join(vItems, "," & vbCrLf & Space$((Level + 1) * LNG_INDENT)) & vbCrLf & _
                    Space$(Level * LNG_INDENT) & Right$(sCompound, 1)
            Else
                JsonDump = Left$(sCompound, 1) & sSpace & Join(vItems, "," & sSpace) & sSpace & Right$(sCompound, 1)
            End If
        End If
    Case vbNull
        JsonDump = "Null"
    Case vbEmpty
        JsonDump = "Empty"
    Case vbString
        '--- one-time initialization of transcoding array
        If IsEmpty(vTranscode) Then
            vTranscode = Split(STR_CODES, "|")
        End If
        For lIdx = 1 To Len(vJson)
            lAsc = AscW(Mid$(vJson, lIdx, 1))
            If lAsc = 92 Or lAsc = 34 Then '--- \ and "
                JsonDump = JsonDump & "\" & ChrW$(lAsc)
            ElseIf lAsc >= 32 And lAsc < 256 Then
                JsonDump = JsonDump & ChrW$(lAsc)
            ElseIf lAsc >= 0 And lAsc < 32 Then
                JsonDump = JsonDump & vTranscode(lAsc)
            ElseIf Asc(Mid$(vJson, lIdx, 1)) <> 63 Or Mid$(vJson, lIdx, 1) = "?" Then '--- ?
                JsonDump = JsonDump & ChrW$(AscW(Mid$(vJson, lIdx, 1)))
            Else
                JsonDump = JsonDump & "\u" & Right$("0000" & Hex$(lAsc), 4)
            End If
        Next
        JsonDump = """" & JsonDump & """"
    Case Else
        If IsArray(vJson) Then
            JsonDump = Join(vJson)
        Else
            JsonDump = vJson & vbNullString
        End If
    End Select
End Function

Public Property Get JsonItem(oJson As Object, ByVal sKey As String) As Variant
    Const FUNC_NAME     As String = "JsonItem [get]"
    Dim vSplit          As Variant
    Dim lIdx            As Long
    Dim lJdx            As Long
    Dim vKey            As Variant
    Dim vItem           As Variant
    #If ImplScripting Then
        Dim oParam      As Scripting.Dictionary
    #Else
        Dim oParam      As VBA.Collection
    #End If
    
    On Error GoTo EH
    If oJson Is Nothing Then
        GoTo ReturnEmpty
    End If
    If LenB(sKey) = 0 Then
        vSplit = Array(vbNullString)
    Else
        vSplit = Split(sKey, "/")
    End If
    Set oParam = oJson
    For lIdx = 0 To UBound(vSplit)
        vKey = vSplit(lIdx)
        If C_Str(vKey) = "-1" Then
            JsonItem = oParam.Count
            GoTo QH
        ElseIf IsOnlyDigits(vKey) Then
            If pvJsonCompareMode(oParam) <> vbBinaryCompare Then
                vKey = C_Lng(vKey)
            End If
        End If
        AssignVariant vItem, pvJsonItem(oParam, vKey)
        If Not IsEmpty(vItem) Then
            If lIdx < UBound(vSplit) Then
                If Not IsObject(vItem) Then
                    GoTo ReturnEmpty
                End If
                Set oParam = vItem
            Else
                AssignVariant JsonItem, vItem
            End If
        ElseIf C_Str(vKey) = "0" Then
            '--- do nothing & continue
        Else
            If LenB(vKey) = 0 Then
                Set JsonItem = oParam
            ElseIf C_Str(vKey) = "*" Then
                vKey = vbNullString
                For lJdx = lIdx + 1 To UBound(vSplit)
                    vKey = vKey & "/" & vSplit(lJdx)
                Next
                If oParam.Count > 0 Then
                    ReDim vItem(0 To oParam.Count - 1) As Variant
                Else
                    vItem = Array()
                End If
                lJdx = 0
                For Each vSplit In JsonKeys(oParam)
                    AssignVariant vItem(lJdx), JsonItem(oParam, vSplit & vKey)
                    lJdx = lJdx + 1
                Next
                If lJdx = 0 Then
                    JsonItem = Array()
                Else
                    If lJdx - 1 <> UBound(vItem) Then
                        ReDim Preserve vItem(0 To lJdx - 1) As Variant
                    End If
                    JsonItem = vItem
                End If
            Else
ReturnEmpty:
                If Right$(sKey, 1) = "/" Then
                    Set JsonItem = pvJsonCreateObject(vbBinaryCompare)
                ElseIf Right$(sKey, 3) = "/-1" Then
                    JsonItem = 0&
                ElseIf InStr(sKey, "*") > 0 Then
                    JsonItem = Array()
                End If
            End If
            GoTo QH
        End If
    Next
QH:
    Exit Property
EH:
    If RaiseError(FUNC_NAME & "(sKey=" & sKey & ")") = vbRetry Then
        Resume
    End If
End Property

Public Property Let JsonItem(oJson As Object, ByVal sKey As String, vValue As Variant)
    Const FUNC_NAME     As String = "JsonItem [let]"
    Dim vSplit          As Variant
    Dim lIdx            As Long
    Dim lJdx            As Long
    Dim vKey            As Variant
    Dim lKey            As Long
    Dim vItem           As Variant
    #If ImplScripting Then
        Dim oParam      As Scripting.Dictionary
    #Else
        Dim oParam      As VBA.Collection
    #End If

    On Error GoTo EH
    If LenB(sKey) = 0 Then
        vSplit = Array(vbNullString)
    Else
        vSplit = Split(sKey, "/")
    End If
    If oJson Is Nothing Then
        If UBound(vSplit) < 0 Then
            Set oJson = pvJsonCreateObject(vbBinaryCompare)
        Else
            Set oJson = pvJsonCreateObject(-(IsOnlyDigits(vSplit(0)) Or vSplit(0) = "*" Or vSplit(0) = "-1"))
        End If
    End If
    Set oParam = oJson
    For lIdx = 0 To UBound(vSplit)
        vKey = vSplit(lIdx)
        If C_Str(vKey) = "-1" Then
            vKey = oParam.Count
        ElseIf IsOnlyDigits(vKey) Then
            If pvJsonCompareMode(oParam) <> vbBinaryCompare Then
                vKey = C_Lng(vKey)
            End If
        End If
        If C_Str(vKey) = "*" Then
HandleArray:
            vKey = vbNullString
            For lJdx = lIdx + 1 To UBound(vSplit)
                vKey = vKey & "/" & vSplit(lJdx)
            Next
            If IsEmpty(vValue) Then
                For Each vItem In JsonKeys(oParam)
                    JsonItem(oParam, vItem & vKey) = Empty
                Next
            Else
                lKey = 0
                For Each vItem In vValue
                    JsonItem(oParam, lKey & vKey) = vItem
                    lKey = lKey + 1
                Next
            End If
            Exit For
        ElseIf lIdx < UBound(vSplit) Then
            If Not IsObject(pvJsonItem(oParam, vKey)) Then
                pvJsonItem(oParam, vKey) = pvJsonCreateObject(-(IsOnlyDigits(vSplit(lIdx + 1)) Or vSplit(lIdx + 1) = "*" Or vSplit(lIdx + 1) = "-1"))
            End If
            Set oParam = pvJsonItem(oParam, vKey)
        ElseIf IsEmpty(vValue) Then
            #If ImplScripting Then
                If oParam.Exists(vKey) Then
                    oParam.Remove vKey
                End If
            #Else
                If VarType(vKey) = vbLong Then
                    lKey = vKey + IDX_OFFSET
                    If lKey > 0 And lKey <= oParam.Count Then
                        oParam.Remove lKey
                    End If
                Else
                    If CollectionIndexByKey(oParam, vKey, DEF_IGNORE_CASE) > 0 Then
                        oParam.Remove vKey
                    End If
                End If
            #End If
        ElseIf IsArray(vValue) Then
            pvJsonItem(oParam, vKey) = pvJsonCreateObject(vbTextCompare)
            Set oParam = pvJsonItem(oParam, vKey)
            GoTo HandleArray
        Else
            pvJsonItem(oParam, vKey) = vValue
        End If
    Next
    Exit Property
EH:
    If RaiseError(FUNC_NAME & "(sKey=" & sKey & ", vValue=" & C_Str(vValue) & ")") = vbRetry Then
        Resume
    End If
End Property

Public Function JsonKeys(oJson As Object, Optional ByVal Key As String) As Variant
    Const FUNC_NAME     As String = "JsonKeys"
    Dim vSplit          As Variant
    Dim lIdx            As Long
    Dim vKey            As Variant
    Dim vItem           As Variant
    Dim lCount          As Long
    #If ImplScripting Then
        Dim oParam      As Scripting.Dictionary
    #Else
        Dim oParam      As VBA.Collection
    #End If
    
    On Error GoTo EH
    If oJson Is Nothing Then
        JsonKeys = Array()
        Exit Function
    End If
    vSplit = Split(Key, "/")
    Set oParam = oJson
    For lIdx = 0 To UBound(vSplit)
        vKey = vSplit(lIdx)
        If IsOnlyDigits(vKey) Then
            If pvJsonCompareMode(oParam) <> vbBinaryCompare Then
                vKey = C_Lng(vKey)
            End If
        End If
        AssignVariant vItem, pvJsonItem(oParam, vKey)
        If IsObject(vItem) Then
            Set oParam = vItem
        Else
            JsonKeys = Array()
            Exit Function
        End If
    Next
    lCount = oParam.Count
    If lCount = 0 Then
        JsonKeys = Array()
        Exit Function
    End If
    ReDim vItem(0 To lCount - 1) As Variant
    If pvJsonCompareMode(oParam) = vbBinaryCompare Then
        #If ImplScripting Then
            vItem = oParam.Keys
        #Else
            vItem = CollectionAllKeys(oParam)
        #End If
    Else
        For lIdx = 0 To UBound(vItem)
            vItem(lIdx) = lIdx
        Next
    End If
    JsonKeys = vItem
    Exit Function
EH:
    If RaiseError(FUNC_NAME & "(Key=" & Key & ")") = vbRetry Then
        Resume
    End If
End Function

Public Function JsonObjectType(oJson As Object, Optional ByVal Key As String) As String
    Const FUNC_NAME     As String = "JsonObjectType"
    Dim vSplit          As Variant
    Dim lIdx            As Long
    Dim vKey            As Variant
    Dim vItem           As Variant
    #If ImplScripting Then
        Dim oParam      As Scripting.Dictionary
    #Else
        Dim oParam      As VBA.Collection
    #End If
    
    On Error GoTo EH
    If oJson Is Nothing Then
        Exit Function
    End If
    vSplit = Split(Key, "/")
    Set oParam = oJson
    For lIdx = 0 To UBound(vSplit)
        vKey = vSplit(lIdx)
        If IsOnlyDigits(vKey) Then
            If pvJsonCompareMode(oParam) <> vbBinaryCompare Then
                vKey = C_Lng(vKey)
            End If
        End If
        AssignVariant vItem, pvJsonItem(oParam, vKey)
        If IsObject(vItem) Then
            Set oParam = vItem
        Else
            Exit Function
        End If
    Next
    JsonObjectType = IIf(pvJsonCompareMode(oParam) = vbBinaryCompare, "object", "array")
    Exit Function
EH:
    If RaiseError(FUNC_NAME & "(Key=" & Key & ")") = vbRetry Then
        Resume
    End If
End Function

Public Function JsonToXmlDocument(vJson As Variant, Optional Root As Object, Optional Doc As Object) As Object
    Const FUNC_NAME     As String = "JsonToXmlDocument"
    Dim vElem           As Variant
    Dim vItem           As Variant
    Dim oArray          As Object
    Dim oItem           As Object
    #If ImplScripting Then
        Dim oJson       As Scripting.Dictionary
    #Else
        Dim oJson       As VBA.Collection
    #End If
    Dim lCount          As Long
    Dim lIdx            As Long
    
    On Error GoTo EH
    If Doc Is Nothing Then
        Set Doc = CreateObject("MSXML2.DOMDocument")
        Doc.appendChild Doc.createProcessingInstruction("xml", "version=""1.0"" encoding=""UTF-16""")
    End If
    If Root Is Nothing Then
        Set Root = Doc.appendChild(Doc.createElement("Root"))
    End If
    If IsObject(vJson) Then
        Set oJson = vJson
        If Not oJson Is Nothing Then
            lCount = oJson.Count
            If lCount = 0 Then
                If Not Root Is Doc.documentElement Then
                    Root.setAttribute STR_ATTR_EMPTY & pvJsonCompareMode(oJson), 1
                End If
                If pvJsonCompareMode(oJson) <> vbBinaryCompare Then
                    Root.setAttribute STR_ATTR_ARRAY, 1
                End If
            Else
                If pvJsonCompareMode(oJson) <> vbBinaryCompare Then
                    Set oArray = Root
                    If oArray Is Doc.documentElement Then
                        Set oArray = oArray.appendChild(Doc.createElement(STR_NODE_ARRAY))
                    ElseIf lCount = 1 Then
                        oArray.setAttribute STR_ATTR_ARRAY, 1
                    End If
                    Set oItem = oArray
                    For lIdx = 0 To lCount - 1
                        AssignVariant vItem, oJson.Item(lIdx + IDX_OFFSET)
                        If oItem Is Nothing Then
                            Set oItem = oArray.ParentNode.appendChild(Doc.createElement(oArray.nodeName))
                        End If
                        JsonToXmlDocument vItem, oItem, Doc
                        Set oItem = Nothing
                    Next
                Else
                    #If ImplScripting Then
                        For Each vElem In oJson.Keys
                    #Else
                        For Each vElem In CollectionAllKeys(oJson)
                    #End If
                        AssignVariant vItem, pvJsonItem(oJson, vElem)
                        If Left$(vElem, 1) = "@" Then
                            If IsObject(vItem) Then
                                Root.setAttribute Mid$(vElem, 2), JsonDump(vItem, Minimize:=True)
                            Else
                                Root.setAttribute Mid$(vElem, 2), vItem
                            End If
                        Else
                            If IsOnlyDigits(vElem) Or LenB(vElem) = 0 Then
                                Set oItem = Root.appendChild(Doc.createElement(STR_PREFIX & Root.childNodes.Length))
                                oItem.setAttribute STR_ATTR_NAME, vElem
                            Else
                                Set oItem = Root.appendChild(Doc.createElement(vElem))
                            End If
                            JsonToXmlDocument vItem, oItem, Doc
                        End If
                    Next
                End If
            End If
        End If
    ElseIf IsEmpty(vJson) Then
        Root.setAttribute STR_ATTR_EMPTY, 1
    ElseIf IsNull(vJson) Then
        Root.setAttribute STR_ATTR_NIL, 1
    ElseIf VarType(vJson) = vbBoolean Then
        Root.setAttribute STR_ATTR_BOOL, 1
        Root.NodeTypedValue = -vJson
    ElseIf VarType(vJson) = vbDate Then
        Root.DataType = "dateTime.tz"
        Root.NodeTypedValue = vJson
    ElseIf IsArray(vJson) Then
        Root.NodeTypedValue = Join(vJson)
    Else
        Root.NodeTypedValue = vJson
    End If
    Set JsonToXmlDocument = Root
    Exit Function
EH:
    If RaiseError(FUNC_NAME) = vbRetry Then
        Resume
    End If
End Function

Public Function JsonFromXmlDocument(vXml As Variant) As Variant
    Const FUNC_NAME     As String = "JsonFromXmlDocument"
    Dim oRoot           As Object
    Dim oNode           As Object
    Dim sKey            As String
    Dim bHasAttributes  As Boolean
    Dim vItem           As Variant
    #If ImplScripting Then
        Dim oDict       As Scripting.Dictionary
        Dim oArray      As Scripting.Dictionary
    #Else
        Dim oDict       As VBA.Collection
        Dim oArray      As VBA.Collection
    #End If

    On Error GoTo EH
    If IsObject(vXml) Then
        Set oRoot = vXml
    Else
        With CreateObject("MSXML2.DOMDocument")
            .LoadXml C_Str(vXml)
            Set oRoot = .documentElement
        End With
        If oRoot Is Nothing Then
            Exit Function
        End If
    End If
    If Not oRoot.firstChild Is Nothing Then
        bHasAttributes = Not oRoot.firstChild.Attributes Is Nothing
    Else
        bHasAttributes = oRoot.Attributes.Length > 0 Or oRoot Is oRoot.ownerDocument.documentElement
    End If
    If C_Bool(oRoot.getAttribute(STR_ATTR_EMPTY & vbBinaryCompare)) Then
        Set JsonFromXmlDocument = pvJsonCreateObject(vbBinaryCompare)
    ElseIf C_Bool(oRoot.getAttribute(STR_ATTR_EMPTY)) Then
        JsonFromXmlDocument = Empty
    ElseIf C_Bool(oRoot.getAttribute(STR_ATTR_NIL)) Then
        JsonFromXmlDocument = Null
    ElseIf C_Bool(oRoot.getAttribute(STR_ATTR_BOOL)) Then
        JsonFromXmlDocument = C_Bool(oRoot.Text)
    ElseIf bHasAttributes Then
        If oRoot.firstChild Is Nothing Then
            Set oDict = pvJsonCreateObject(-C_Bool(oRoot.getAttribute(STR_ATTR_ARRAY)))
        Else
            Set oDict = pvJsonCreateObject(vbBinaryCompare)
        End If
        For Each oNode In oRoot.Attributes
            sKey = C_Str(oNode.nodeName)
            If Left$(sKey, Len(STR_PREFIX)) <> STR_PREFIX Then
                sKey = "@" & sKey
                If Left$(oNode.Text, 1) = "{" Or Left$(oNode.Text, 1) = "[" Then
                    If JsonParse(oNode.Text, vItem, StrictMode:=True) Then
                        pvJsonItem(oDict, sKey) = vItem
                    Else
                        pvJsonItem(oDict, sKey) = oNode.Text
                    End If
                ElseIf C_Str(C_Lng(oNode.Text)) = oNode.Text Then
                    pvJsonItem(oDict, sKey) = C_Lng(oNode.Text)
                ElseIf C_Str(C_Dbl(oNode.Text)) = oNode.Text Then
                    pvJsonItem(oDict, sKey) = C_Dbl(oNode.Text)
                Else
                    pvJsonItem(oDict, sKey) = oNode.NodeTypedValue
                End If
            End If
        Next
        sKey = vbNullString
        For Each oNode In oRoot.childNodes
            If Not IsNull(oNode.getAttribute(STR_ATTR_NAME)) Then
                sKey = C_Str(oNode.getAttribute(STR_ATTR_NAME))
            Else
                sKey = C_Str(oNode.nodeName)
            End If
            If Not IsEmpty(pvJsonItem(oDict, sKey)) Or sKey = STR_NODE_ARRAY Or C_Bool(oNode.getAttribute(STR_ATTR_ARRAY)) Then
                AssignVariant vItem, pvJsonItem(oDict, sKey)
                If IsEmpty(vItem) Then
                    Set oArray = pvJsonCreateObject(vbTextCompare)
                    pvJsonItem(oDict, sKey) = oArray
                ElseIf Not IsObject(vItem) Then
CreateArray:
                    Set oArray = pvJsonCreateObject(vbTextCompare)
                    #If ImplScripting Then
                        oArray.Add 0&, vItem
                    #Else
                        oArray.Add vItem
                    #End If
                    pvJsonItem(oDict, sKey) = oArray
                ElseIf pvJsonCompareMode(C_Obj(vItem)) = vbBinaryCompare Then
                    GoTo CreateArray
                Else
                    Set oArray = C_Obj(vItem)
                End If
                If Not C_Bool(oNode.getAttribute(STR_ATTR_EMPTY & 0)) Then
                    #If ImplScripting Then
                        oArray.Add oArray.Count, JsonFromXmlDocument(oNode)
                    #Else
                        oArray.Add JsonFromXmlDocument(oNode)
                    #End If
                End If
            Else
                pvJsonItem(oDict, sKey) = JsonFromXmlDocument(oNode)
            End If
        Next
        If sKey = STR_NODE_ARRAY Then
            Set JsonFromXmlDocument = pvJsonItem(oDict, sKey)
        Else
            Set JsonFromXmlDocument = oDict
        End If
    ElseIf C_Str(C_Lng(oRoot.Text)) = oRoot.Text Then
        JsonFromXmlDocument = C_Lng(oRoot.Text)
    ElseIf C_Str(C_Dbl(oRoot.Text)) = oRoot.Text Then
        JsonFromXmlDocument = C_Dbl(oRoot.Text)
    ElseIf oRoot.Text Like "####-##-##T##:##:##*" Then
        vItem = Split(Replace(Replace(Replace(Replace(oRoot.Text, "T", "-"), ":", "-"), ".", "-"), "+", "-"), "-")
        JsonFromXmlDocument = DateSerial(C_Lng(vItem(0)), C_Lng(vItem(1)), C_Lng(vItem(2))) + TimeSerial(C_Lng(vItem(3)), C_Lng(vItem(4)), Val(vItem(5)))
    Else
        JsonFromXmlDocument = oRoot.NodeTypedValue
    End If
    Exit Function
EH:
    If RaiseError(FUNC_NAME) = vbRetry Then
        Resume
    End If
End Function

#If ImplScripting Then
    Public Function JsonToDictionary(oJson As Object) As Object
        Set JsonToDictionary = oJson
    End Function
#Else
    Public Function JsonToDictionary(oJson As Object) As Object
        Const FUNC_NAME     As String = "JsonToDictionary"
        Dim oRetVal         As Object
        Dim vKeys           As Variant
        Dim lKey            As Long
        Dim vKey            As Variant
        Dim vElem           As Variant
        
        On Error GoTo EH
        If oJson Is Nothing Then
            Exit Function
        End If
        Set oRetVal = CreateObject("Scripting.Dictionary")
        vKeys = JsonKeys(oJson)
        If UBound(vKeys) < 0 And oJson.Count > 0 Then
            For lKey = 0 To oJson.Count - 1
                AssignVariant vElem, JsonItem(oJson, lKey)
                If IsObject(vElem) Then
                    Set oRetVal.Item(lKey) = JsonToDictionary(C_Obj(vElem))
                Else
                    oRetVal.Item(lKey) = vElem
                End If
            Next
        Else
            For Each vKey In vKeys
                AssignVariant vElem, JsonItem(oJson, vKey)
                If IsObject(vElem) Then
                    Set oRetVal.Item(vKey) = JsonToDictionary(C_Obj(vElem))
                Else
                    oRetVal.Item(vKey) = vElem
                End If
            Next
        End If
        Set JsonToDictionary = oRetVal
        Exit Function
EH:
        If RaiseError(FUNC_NAME) = vbRetry Then
            Resume
        End If
    End Function
#End If

#If ImplScripting Then
    Private Function pvJsonCreateObject(ByVal lCompareMode As VbCompareMethod) As Scripting.Dictionary
        Set pvJsonCreateObject = New Scripting.Dictionary
        pvJsonCreateObject.CompareMode = lCompareMode
    End Function
    
    Private Function pvJsonCompareMode(oJson As Scripting.Dictionary) As VbCompareMethod
        pvJsonCompareMode = oJson.CompareMode
    End Function
    
    Private Property Get pvJsonItem(oParam As Scripting.Dictionary, vKey As Variant) As Variant
        If oParam.Exists(vKey) Then
            AssignVariant pvJsonItem, oParam.Item(vKey)
        End If
    End Property

    Private Property Let pvJsonItem(oParam As Scripting.Dictionary, vKey As Variant, vValue As Variant)
        If IsObject(vValue) Then
            Set oParam.Item(vKey) = vValue
        Else
            oParam.Item(vKey) = vValue
        End If
    End Property
#Else
    Private Function pvJsonCreateObject(ByVal lCompareMode As VbCompareMethod) As VBA.Collection
        Set pvJsonCreateObject = New VBA.Collection
        #If LargeAddressAware Then
            Call CopyMemory(ByVal (ObjPtr(pvJsonCreateObject) Xor SIGN_BIT) + o_pvUnk5 Xor SIGN_BIT, lCompareMode, 4)
        #Else
            Call CopyMemory(ByVal ObjPtr(pvJsonCreateObject) + o_pvUnk5, lCompareMode, 4)
        #End If
    End Function
    
    Private Function pvJsonCompareMode(oJson As VBA.Collection) As VbCompareMethod
        #If LargeAddressAware Then
            Call CopyMemory(pvJsonCompareMode, ByVal (ObjPtr(oJson) Xor SIGN_BIT) + o_pvUnk5 Xor SIGN_BIT, 4)
        #Else
            Call CopyMemory(pvJsonCompareMode, ByVal ObjPtr(oJson) + o_pvUnk5, 4)
        #End If
    End Function

    Private Property Get pvJsonItem(oParam As VBA.Collection, vKey As Variant) As Variant
        #If ImplUseShared Then
            If VarType(vKey) = vbLong Then
                SearchCollection oParam, vKey + 1, RetVal:=pvJsonItem
            Else
                SearchCollection oParam, vKey, RetVal:=pvJsonItem
            End If
        #Else
            Const FUNC_NAME     As String = "pvJsonItem [get]"
            Dim lKey            As Long
            
            On Error GoTo EH
            If VarType(vKey) = vbLong Then
                lKey = vKey + 1
                If lKey > 0 And lKey <= oParam.Count Then
                    AssignVariant pvJsonItem, oParam.Item(lKey)
                End If
            Else
                If CollectionIndexByKey(oParam, vKey, DEF_IGNORE_CASE) > 0 Then
                    AssignVariant pvJsonItem, oParam.Item(vKey)
                End If
            End If
            Exit Property
EH:
            If RaiseError(FUNC_NAME & "(vKey=" & C_Str(vKey) & ")") = vbRetry Then
                Resume
            End If
        #End If
    End Property

    Private Property Let pvJsonItem(oParam As VBA.Collection, vKey As Variant, vValue As Variant)
        Const FUNC_NAME     As String = "pvJsonItem [let]"
        Dim lKey            As Long
        
        On Error GoTo EH
        If VarType(vKey) = vbLong Then
            lKey = vKey + 1
            If lKey > 0 And lKey <= oParam.Count Then
                oParam.Remove lKey
            End If
            If lKey > 0 And lKey <= oParam.Count Then
                oParam.Add vValue, Before:=lKey
            Else
                oParam.Add vValue
            End If
        Else
            lKey = CollectionIndexByKey(oParam, vKey, DEF_IGNORE_CASE)
            If lKey > 0 Then
                oParam.Remove lKey
            End If
            If lKey > 0 And lKey <= oParam.Count Then
                oParam.Add vValue, vKey, Before:=lKey
            Else
                oParam.Add vValue, vKey
            End If
        End If
        
        Exit Property
EH:
        If RaiseError(FUNC_NAME & "(vKey=" & C_Str(vKey) & ", lKey=" & lKey & ", vValue=" & C_Str(vValue) & ")") = vbRetry Then
            Resume
        End If
    End Property
#End If

#If Not ImplUseShared Then
Private Function IsOnlyDigits(ByVal sText As String) As Boolean
    If LenB(sText) <> 0 Then
        IsOnlyDigits = Not (sText Like "*[!0-9]*")
    End If
End Function

Private Sub AssignVariant(vDest As Variant, vSrc As Variant)
    On Error GoTo QH
    If IsObject(vSrc) Then
        Set vDest = vSrc
    Else
        vDest = vSrc
    End If
QH:
End Sub

Private Function C_Str(Value As Variant) As String
    Dim vDest           As Variant
    
    If VarType(Value) = vbString Then
        C_Str = Value
    ElseIf VariantChangeType(vDest, Value, VARIANT_ALPHABOOL, vbString) = 0 Then
        C_Str = vDest
    End If
End Function

Private Function C_Bool(Value As Variant) As Boolean
    Dim vDest           As Variant
    
    If VarType(Value) = vbBoolean Then
        C_Bool = Value
    ElseIf VariantChangeType(vDest, Value, VARIANT_ALPHABOOL, vbBoolean) = 0 Then
        C_Bool = vDest
    End If
End Function

Private Function C_Lng(Value As Variant) As Long
    Dim vDest       As Variant
    
    If VarType(Value) = vbLong Then
        C_Lng = Value
    ElseIf VariantChangeType(vDest, Value, 0, vbLong) = 0 Then
        C_Lng = vDest
    End If
End Function

Private Function C_Dbl(Value As Variant) As Double
    Dim vDest       As Variant
    
    If VarType(Value) = vbDouble Then
        C_Dbl = Value
    ElseIf VariantChangeType(vDest, Value, 0, vbDouble) = 0 Then
        C_Dbl = vDest
    End If
End Function

Private Function C_Obj(Value As Variant) As Object
    Dim vDest       As Variant

    If VarType(Value) = vbObject Then
        Set C_Obj = Value
    ElseIf VariantChangeType(vDest, Value, 0, vbObject) = 0 Then
        Set C_Obj = vDest
    End If
End Function

Private Function Printf(ByVal sText As String, ParamArray A() As Variant) As String
    Const LNG_PRIVATE   As Long = &HE1B6 '-- U+E000 to U+F8FF - Private Use Area (PUA)
    Dim lIdx            As Long
    
    For lIdx = UBound(A) To LBound(A) Step -1
        sText = Replace(sText, "%" & (lIdx - LBound(A) + 1), Replace(A(lIdx), "%", ChrW$(LNG_PRIVATE)))
    Next
    Printf = Replace(sText, ChrW$(LNG_PRIVATE), "%")
End Function

Private Function CollectionAllKeys(oCol As VBA.Collection, Optional ByVal StartIndex As Long) As String()
    #If HasPtrSafe Then
        Dim lPtr        As LongPtr
    #Else
        Dim lPtr        As Long
    #End If
    Dim aRetVal()       As String
    Dim lIdx            As Long
    Dim sTemp           As String
    
    If oCol.Count = 0 Then
        aRetVal = Split(vbNullString)
    Else
        ReDim aRetVal(StartIndex To StartIndex + oCol.Count - 1) As String
        lPtr = ObjPtr(oCol)
        For lIdx = LBound(aRetVal) To UBound(aRetVal)
            #If LargeAddressAware Then
                Call CopyMemory(lPtr, ByVal (lPtr Xor SIGN_BIT) + o_pNextIndexedItem Xor SIGN_BIT, PTR_SIZE)
                Call CopyMemory(ByVal VarPtr(sTemp), ByVal (lPtr Xor SIGN_BIT) + o_KeyPtr Xor SIGN_BIT, PTR_SIZE)
            #Else
                Call CopyMemory(lPtr, ByVal lPtr + o_pNextIndexedItem, PTR_SIZE)
                Call CopyMemory(ByVal VarPtr(sTemp), ByVal lPtr + o_KeyPtr, PTR_SIZE)
            #End If
            aRetVal(lIdx) = sTemp
        Next
        Call CopyMemory(ByVal VarPtr(sTemp), NULL_PTR, PTR_SIZE)
    End If
    CollectionAllKeys = aRetVal
End Function

Private Function CollectionIndexByKey(oCol As VBA.Collection, ByVal sKey As String, Optional ByVal IgnoreCase As Boolean = True) As Long
    #If HasPtrSafe Then
        Dim lItemPtr    As LongPtr
        Dim lEofPtr     As LongPtr
        Dim lPtr        As LongPtr
    #Else
        Dim lItemPtr    As Long
        Dim lEofPtr     As Long
        Dim lPtr        As Long
    #End If
    Dim sTemp           As String
    
    If Not oCol Is Nothing Then
        #If LargeAddressAware Then
            Call CopyMemory(lItemPtr, ByVal (ObjPtr(oCol) Xor SIGN_BIT) + o_pRootTreeItem Xor SIGN_BIT, PTR_SIZE)
            Call CopyMemory(lEofPtr, ByVal (ObjPtr(oCol) Xor SIGN_BIT) + o_pEndTreePtr Xor SIGN_BIT, PTR_SIZE)
        #Else
            Call CopyMemory(lItemPtr, ByVal ObjPtr(oCol) + o_pRootTreeItem, PTR_SIZE)
            Call CopyMemory(lEofPtr, ByVal ObjPtr(oCol) + o_pEndTreePtr, PTR_SIZE)
        #End If
    End If
    Do While lItemPtr <> lEofPtr
        #If LargeAddressAware Then
            Call CopyMemory(ByVal VarPtr(sTemp), ByVal (lItemPtr Xor SIGN_BIT) + o_KeyPtr Xor SIGN_BIT, PTR_SIZE)
        #Else
            Call CopyMemory(ByVal VarPtr(sTemp), ByVal lItemPtr + o_KeyPtr, PTR_SIZE)
        #End If
        Select Case CompareStringW(LOCALE_USER_DEFAULT, -IgnoreCase * NORM_IGNORECASE, ByVal StrPtr(sKey), Len(sKey), ByVal StrPtr(sTemp), Len(sTemp))
        Case CSTR_LESS_THAN
            #If LargeAddressAware Then
                Call CopyMemory(lItemPtr, ByVal (lItemPtr Xor SIGN_BIT) + o_pLeftBranch Xor SIGN_BIT, PTR_SIZE)
            #Else
                Call CopyMemory(lItemPtr, ByVal lItemPtr + o_pLeftBranch, PTR_SIZE)
            #End If
        Case CSTR_GREATER_THAN
            #If LargeAddressAware Then
                Call CopyMemory(lItemPtr, ByVal (lItemPtr Xor SIGN_BIT) + o_pRightBranch Xor SIGN_BIT, PTR_SIZE)
            #Else
                Call CopyMemory(lItemPtr, ByVal lItemPtr + o_pRightBranch, PTR_SIZE)
            #End If
        Case CSTR_EQUAL
            lPtr = ObjPtr(oCol)
            Do While lPtr <> lItemPtr
                #If LargeAddressAware Then
                    Call CopyMemory(lPtr, ByVal (lPtr Xor SIGN_BIT) + o_pNextIndexedItem Xor SIGN_BIT, PTR_SIZE)
                #Else
                    Call CopyMemory(lPtr, ByVal lPtr + o_pNextIndexedItem, PTR_SIZE)
                #End If
                CollectionIndexByKey = CollectionIndexByKey + 1
            Loop
            GoTo QH
        Case Else
            Call CopyMemory(ByVal VarPtr(sTemp), NULL_PTR, PTR_SIZE)
            Err.Raise vbObjectError, , "Unexpected result from CompareStringW"
        End Select
    Loop
QH:
    Call CopyMemory(ByVal VarPtr(sTemp), NULL_PTR, PTR_SIZE)
End Function
#End If


