Attribute VB_Name = "mdJson"
'=========================================================================
' $Header: /UcsFiscalPrinter/Src/mdJson.bas 3     21.11.13 16:37 Wqw $
'
'   Unicontsoft Fiscal Printers Project
'   Copyright (c) 2008-2013 Unicontsoft
'
'   JSON parsing and dumping functions
'
' $Log: /UcsFiscalPrinter/Src/mdJson.bas $
' 
' 3     21.11.13 16:37 Wqw
' REF: impl error handling
'
' 2     16.11.12 18:52 Wqw
' REF: description
'
' 1     5.10.12 10:22 Wqw
' Initial implementation
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "mdJson"

'=========================================================================
' API
'=========================================================================

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Type JsonContext
    Text()              As Integer
    Pos                 As Long
    Error               As String
    LastChar            As Integer
End Type

'=========================================================================
' Error management
'=========================================================================

Private Sub RaiseError(sFunc As String)
    Debug.Print MODULE_NAME & "." & sFunc & ": " & Err.Description
    OutputDebugLog MODULE_NAME, sFunc & "(" & Erl & ")", "Run-time error: " & Err.Description
    Err.Raise Err.Number, MODULE_NAME & "." & sFunc & "(" & Erl & ")" & vbCrLf & Err.Source, Err.Description
End Sub

Private Sub PrintError(sFunc As String)
    Debug.Print MODULE_NAME & "." & sFunc & ": " & Err.Description
    OutputDebugLog MODULE_NAME, sFunc & "(" & Erl & ")", "Run-time error: " & Err.Description
End Sub

'=========================================================================
' Functions
'=========================================================================

Public Function JsonParse(sText As String, vResult As Variant, Optional Error As String) As Boolean
    Const FUNC_NAME     As String = "JsonParse"
    Dim uCtx            As JsonContext
    Dim oResult         As Object
    
    On Error GoTo EH
    With uCtx
        ReDim .Text(0 To Len(sText)) As Integer
        Call CopyMemory(.Text(0), ByVal StrPtr(sText), LenB(sText))
        JsonParse = pvJsonParse(uCtx, vResult, oResult)
        If Not oResult Is Nothing Then
            Set vResult = oResult
        End If
        Error = .Error
    End With
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function pvJsonMissing(Optional vMissing As Variant) As Variant
    pvJsonMissing = vMissing
End Function

Private Function pvJsonParse(uCtx As JsonContext, vResult As Variant, oResult As Object) As Boolean
    '--- note: when using collections change type of parameter oResult to Collection
    #Const USE_RICHCLIENT = False
    #Const USE_COLLECTION = False
    Const FUNC_NAME     As String = "pvJsonParse"
    Dim lIdx            As Long
    Dim vKey            As Variant
    Dim vValue          As Variant
    Dim oValue          As Object
    Dim sText           As String
    
    On Error GoTo EH
    vValue = pvJsonMissing
    With uCtx
        Select Case pvJsonGetChar(uCtx)
        Case 34 ' "
            vResult = pvJsonGetString(uCtx)
        Case 91 ' [
            #If USE_RICHCLIENT Then
                #If USE_COLLECTION Then
                    Set oResult = New cCollection
                #Else
                    Set oResult = New cSortedDictionary
                #End If
            #Else
                #If USE_COLLECTION Then
                    Set oResult = New Collection
                #Else
                    Set oResult = CreateObject("Scripting.Dictionary")
                #End If
            #End If
            Do
                Select Case pvJsonGetChar(uCtx)
                Case 0, 44, 93 ' , ]
                    If Not oValue Is Nothing Then
                        #If USE_COLLECTION Then
                            oResult.Add oValue
                        #Else
                            oResult.Add lIdx, oValue
                        #End If
                    ElseIf Not IsMissing(vValue) Then
                        #If USE_COLLECTION Then
                            oResult.Add vValue
                        #Else
                            oResult.Add lIdx, vValue
                        #End If
                    End If
                    If .LastChar <> 44 Then ' ,
                        Exit Do
                    End If
                    lIdx = lIdx + 1
                    vValue = pvJsonMissing
                    Set oValue = Nothing
                Case Else
                    .Pos = .Pos - 1
                    If Not pvJsonParse(uCtx, vValue, oValue) Then
                        GoTo QH
                    End If
                End Select
            Loop
        Case 123 ' {
            #If USE_RICHCLIENT Then
                #If USE_COLLECTION Then
                    Set oResult = New cCollection
                #Else
                    Set oResult = New cSortedDictionary
                    oResult.StringCompareMode = 1 ' TextCompare
                #End If
            #Else
                #If USE_COLLECTION Then
                    Set oResult = New Collection
                #Else
                    Set oResult = CreateObject("Scripting.Dictionary")
                    oResult.CompareMode = 1 ' TextCompare
                #End If
            #End If
            Do
                Select Case pvJsonGetChar(uCtx)
                Case 34 ' "
                    vKey = pvJsonGetString(uCtx)
                Case 58 ' :
                    If Not oValue Is Nothing Then
                        .Error = "Value already specified at position " & .Pos
                        GoTo QH
                    ElseIf Not IsMissing(vValue) Then
                        vKey = vValue
                        vValue = pvJsonMissing
                    End If
                    lIdx = .Pos
                    If Not pvJsonParse(uCtx, vValue, oValue) Then
                        .Pos = lIdx
                        vValue = Empty
                        Set oValue = Nothing
                    End If
                Case 0, 44, 125 ' , }
                    If IsMissing(vValue) And oValue Is Nothing Then
                        If IsEmpty(vKey) Then
                            GoTo NoProp
                        End If
                        vValue = vKey
                        vKey = vbNullString
                    End If
                    If IsEmpty(vKey) Then
                        vKey = vbNullString
                    ElseIf IsNull(vKey) Then
                        vKey = "null"
                    End If
                    If Not oValue Is Nothing Then
                        #If USE_COLLECTION Then
                            oResult.Add oValue, vKey & ""
                        #Else
                            oResult.Add vKey & "", oValue
                        #End If
                    Else
                        #If USE_COLLECTION Then
                            oResult.Add vValue, vKey & ""
                        #Else
                            oResult.Add vKey & "", vValue
                        #End If
                    End If
NoProp:
                    If .LastChar = 0 Then
                        GoTo QH
                    ElseIf .LastChar <> 44 Then ' ,
                        Exit Do
                    End If
                    vKey = Empty
                    vValue = pvJsonMissing
                    Set oValue = Nothing
                Case Else
                    .Pos = .Pos - 1
                    If Not pvJsonParse(uCtx, vValue, oValue) Then
                        GoTo QH
                    End If
                End Select
            Loop
        Case 116, 84  ' "t", "T"
            If Not ((.Text(.Pos + 0) Or &H20) = 114 And (.Text(.Pos + 1) Or &H20) = 117 And (.Text(.Pos + 2) Or &H20) = 101) Then
                GoTo UnexpectedSymbol
            End If
            .Pos = .Pos + 3
            vResult = True
        Case 102, 70 ' "f", "F"
            If Not ((.Text(.Pos + 0) Or &H20) = 97 And (.Text(.Pos + 1) Or &H20) = 108 And (.Text(.Pos + 2) Or &H20) = 115 And (.Text(.Pos + 3) Or &H20) = 101) Then
                GoTo UnexpectedSymbol
            End If
            .Pos = .Pos + 4
            vResult = False
        Case 110, 78 ' "n", "N"
            If Not ((.Text(.Pos + 0) Or &H20) = 117 And (.Text(.Pos + 1) Or &H20) = 108 And (.Text(.Pos + 2) Or &H20) = 108) Then
                GoTo UnexpectedSymbol
            End If
            .Pos = .Pos + 3
            vResult = Null
        Case 48 To 57, 43, 45, 46 ' 0-9 + - .
            For lIdx = 0 To 1000
                Select Case .Text(.Pos + lIdx)
                Case 48 To 57, 43, 45, 46, 101, 69, 120, 88, 97 To 102, 65 To 70 ' 0-9 + - . e E x X a-f A-F
                Case Else
                    Exit For
                End Select
            Next
            sText = Space$(lIdx + 1)
            Call CopyMemory(ByVal StrPtr(sText), .Text(.Pos - 1), LenB(sText))
            If LCase$(Left$(sText, 2)) = "0x" Then
                sText = "&H" & Mid$(sText, 3)
            End If
            On Error GoTo ErrorConvert
            vResult = CDbl(sText)
            On Error GoTo 0
            .Pos = .Pos + lIdx
        Case 0
            If LenB(.Error) <> 0 Then
                GoTo QH
            End If
        Case Else
            GoTo UnexpectedSymbol
        End Select
        pvJsonParse = True
QH:
        Exit Function
UnexpectedSymbol:
        .Error = "Unexpected symbol '" & ChrW$(.LastChar) & "' at position " & .Pos
        Exit Function
ErrorConvert:
        .Error = Err.Description & " at position " & .Pos
    End With
    Exit Function
EH:
    RaiseError FUNC_NAME
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
            Case 9, 10, 13, 32 ' vbTab, vbCr, vbLf, " "
                '--- do nothing
            Case 47 ' /
                Select Case .Text(.Pos)
                Case 47 ' //
                    .Pos = .Pos + 1
                    Do
                        .LastChar = .Text(.Pos)
                        .Pos = .Pos + 1
                        If .LastChar = 0 Then
                            Exit Function
                        End If
                    Loop While Not (.LastChar = 10 Or .LastChar = 13)  ' vbLf or vbCr
                Case 42 ' /*
                    lIdx = .Pos + 1
                    Do
                        .LastChar = .Text(lIdx)
                        lIdx = lIdx + 1
                        If .LastChar = 0 Then
                            .Error = "Unterminated comment at position " & .Pos
                            Exit Function
                        End If
                    Loop While Not (.LastChar = 42 And .Text(lIdx) = 47) ' */
                    .LastChar = .Text(lIdx)
                    .Pos = lIdx + 1
                Case Else
                    pvJsonGetChar = .LastChar
                    Exit Do
                End Select
            Case Else
                pvJsonGetChar = .LastChar
                Exit Do
            End Select
        Loop
    End With
    Exit Function
EH:
    RaiseError FUNC_NAME
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
            Case 0, 34, 92 ' " \
                sText = Space$(lIdx)
                Call CopyMemory(ByVal StrPtr(sText), .Text(.Pos), LenB(sText))
                pvJsonGetString = pvJsonGetString & sText
                If nChar <> 92 Then ' \
                    .Pos = .Pos + lIdx + 1
                    Exit For
                End If
                lIdx = lIdx + 1
                nChar = .Text(.Pos + lIdx)
                Select Case nChar
                Case 0
                    Exit For
                Case 98  ' b
                    pvJsonGetString = pvJsonGetString & Chr$(8)
                Case 102 ' f
                    pvJsonGetString = pvJsonGetString & Chr$(12)
                Case 110 ' n
                    pvJsonGetString = pvJsonGetString & vbLf
                Case 114 ' r
                    pvJsonGetString = pvJsonGetString & vbCr
                Case 116 ' t
                    pvJsonGetString = pvJsonGetString & vbTab
                Case 117 ' u
                    pvJsonGetString = pvJsonGetString & ChrW$(CLng("&H" & ChrW$(.Text(.Pos + lIdx + 1)) & ChrW$(.Text(.Pos + lIdx + 2)) & ChrW$(.Text(.Pos + lIdx + 3)) & ChrW$(.Text(.Pos + lIdx + 4))))
                    lIdx = lIdx + 4
                Case 120 ' x
                    pvJsonGetString = pvJsonGetString & ChrW$(CLng("&H" & ChrW$(.Text(.Pos + lIdx + 1)) & ChrW$(.Text(.Pos + lIdx + 2))))
                    lIdx = lIdx + 2
                Case Else
                    pvJsonGetString = pvJsonGetString & ChrW$(nChar)
                End Select
                .Pos = .Pos + lIdx + 1
                lIdx = -1
            End Select
        Next
    End With
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function JsonDump(vJson As Variant, Optional ByVal Level As Long, Optional ByVal Minimize As Boolean) As String
    Const FUNC_NAME     As String = "JsonDump"
    Const STR_CODES     As String = "\u0000|\u0001|\u0002|\u0003|\u0004|\u0005|\u0006|\u0007|\b|\t|\n|\u000B|\f|\r|\u000E|\u000F|\u0010|\u0011|\u0012|\u0013|\u0014|\u0015|\u0016|\u0017|\u0018|\u0019|\u001A|\u001B|\u001C|\u001D|\u001E|\u001F"
    Const INDENT        As Long = 4
    Static vTranscode   As Variant
    Dim vKeys           As Variant
    Dim vItems          As Variant
    Dim lIdx            As Long
    Dim lSize           As Long
    Dim sCompound       As String
    Dim sSpace          As String
    Dim lAsc            As Long
    
    On Error GoTo EH
    Select Case VarType(vJson)
    Case vbObject
        sCompound = IIf(vJson.CompareMode = 0, "[]", "{}")
        sSpace = IIf(Minimize, vbNullString, " ")
        If vJson.Count = 0 Then
            JsonDump = sCompound
        Else
            vKeys = vJson.Keys
            vItems = vJson.Items
            For lIdx = 0 To vJson.Count - 1
                vItems(lIdx) = JsonDump(vItems(lIdx), Level + 1, Minimize)
                If vJson.CompareMode = 1 Then
                    vItems(lIdx) = JsonDump(vKeys(lIdx)) & ":" & sSpace & vItems(lIdx)
                End If
                lSize = lSize + Len(vItems(lIdx))
            Next
            If lSize > 100 And Not Minimize Then
                JsonDump = Left$(sCompound, 1) & vbCrLf & _
                    Space$((Level + 1) * INDENT) & Join(vItems, "," & vbCrLf & Space$((Level + 1) * INDENT)) & vbCrLf & _
                    Space$(Level * INDENT) & Right$(sCompound, 1)
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
                JsonDump = JsonDump & "\" & Chr$(lAsc)
            ElseIf lAsc >= 32 And lAsc < 256 Then
                JsonDump = JsonDump & Chr$(lAsc)
            ElseIf lAsc >= 0 And lAsc < 32 Then
                JsonDump = JsonDump & vTranscode(lAsc)
            ElseIf Asc(Mid$(vJson, lIdx, 1)) <> 63 Then '--- ?
                JsonDump = JsonDump & Chr$(Asc(Mid$(vJson, lIdx, 1)))
            Else
                JsonDump = JsonDump & "\u" & Right$("0000" & Hex(lAsc), 4)
            End If
        Next
        JsonDump = """" & JsonDump & """"
    Case Else
        JsonDump = vJson & ""
    End Select
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

