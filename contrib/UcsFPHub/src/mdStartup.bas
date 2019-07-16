Attribute VB_Name = "mdStartup"
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

Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_VERSION           As String = "1.0.0"
Private Const STR_ERROR_CONFIG_NOT_FOUND As String = "Грешка: Конфигурационен файл %1 не е намерен"
Private Const STR_ERROR_PARSING_CONFIG As String = "Грешка: Невалиден %1: %2"
Private Const STR_AUTODETECTING_PRINTERS As String = "Автоматично търсене на принтери..."
Private Const STR_INFO_ERROR_ACCESSING As String = "Информация: Принтер %1: %2"
Private Const STR_PRINTERS_FOUND As String = "Намерени %1 принтера"

Private m_oOpt                  As Object
Private m_oPrinters             As Object

'=========================================================================
' Functions
'=========================================================================

Private Sub Main()
    Dim lExitCode       As Long
    
    lExitCode = Process(SplitArgs(Command$))
    If Not InIde Then
        Call ExitProcess(lExitCode)
    End If
End Sub

Private Function Process(vArgs As Variant) As Long
    Dim sFile           As String
    Dim sError          As String
    Dim oConfig         As Object
    
    On Error GoTo EH
    Set m_oOpt = GetOpt(vArgs)
    If Not m_oOpt.Item("-nologo") Then
        ConsolePrint App.ProductName & " " & STR_VERSION & " (c) 2019 by Unicontsoft" & vbCrLf & vbCrLf
    End If
    sFile = Zn(m_oOpt.Item("-conf"), m_oOpt.Item("c"))
    If LenB(sFile) = 0 Then
        sFile = PathCombine(App.Path, App.EXEName & ".conf")
        If Not FileExists(sFile) Then
            sFile = vbNullString
        End If
    End If
    If LenB(sFile) <> 0 Then
        If Not FileExists(sFile) Then
            ConsoleError STR_ERROR_CONFIG_NOT_FOUND & vbCrLf, sFile
            Process = 1
            GoTo QH
        End If
        If Not JsonParse(FromUtf8Array(ReadBinaryFile(sFile)), oConfig, Error:=sError) Then
            ConsoleError STR_ERROR_PARSING_CONFIG & vbCrLf, sFile, sError
            Process = 1
            GoTo QH
        End If
    Else
        JsonItem(oConfig, "Printers/Autodetect") = True
        JsonItem(oConfig, "Endpoints/0/Binding") = "RestHttp"
        JsonItem(oConfig, "Endpoints/0/Address") = "127.0.0.1:8192"
    End If
    Set m_oPrinters = pvCollectPrinters(oConfig)
    ConsolePrint STR_PRINTERS_FOUND & vbCrLf, JsonItem(m_oPrinters, "Count")
    ConsolePrint JsonDump(m_oPrinters) & vbCrLf
QH:
    Exit Function
EH:
    ConsoleError "Critical error: " & Err.Description & vbCrLf
    Process = 100
End Function

Private Function pvCollectPrinters(oConfig As Object) As Object
    Dim oFP             As cFiscalPrinter
    Dim sResponse       As String
    Dim oJson           As Object
    Dim vKey            As Variant
    Dim oRequest        As Object
    Dim oRetVal         As Object
    Dim sDeviceString   As String
    Dim sKey            As String
    
    On Error GoTo EH
    Set oFP = New cFiscalPrinter
    JsonItem(oRetVal, "Count") = 0
    JsonItem(oRetVal, "Alias") = 0
    If JsonItem(oConfig, "Printers/Autodetect") Then
        ConsolePrint STR_AUTODETECTING_PRINTERS & vbCrLf
        If oFP.EnumPorts(sResponse) And JsonParse(sResponse, oJson) Then
            For Each vKey In JsonKeys(oJson, "SerialPorts")
                If LenB(JsonItem(oJson, "SerialPorts/" & vKey & "/Protocol")) <> 0 Then
                    sDeviceString = "Protocol=" & JsonItem(oJson, "SerialPorts/" & vKey & "/Protocol") & _
                        ";Port=" & JsonItem(oJson, "SerialPorts/" & vKey & "/Port") & _
                        ";Speed=" & JsonItem(oJson, "SerialPorts/" & vKey & "/Speed")
                    Set oRequest = Nothing
                    JsonItem(oRequest, "DeviceString") = sDeviceString
                    If oFP.GetStatus(JsonDump(oRequest, Minimize:=True), sResponse) And JsonParse(sResponse, oJson) Then
                        sKey = JsonItem(oJson, "DeviceSerialNo")
                        If LenB(sKey) <> 0 Then
                            JsonItem(oRetVal, sKey) = sDeviceString
                            JsonItem(oRetVal, "Count") = JsonItem(oRetVal, "Count") + 1
                        End If
                    End If
                End If
            Next
        End If
    End If
    For Each vKey In JsonKeys(oConfig, "Printers")
        sDeviceString = C_Str(JsonItem(oConfig, "Printers/" & vKey & "/DeviceString"))
        If LenB(sDeviceString) <> 0 Then
            Set oRequest = Nothing
            JsonItem(oRequest, "DeviceString") = sDeviceString
            If oFP.GetStatus(JsonDump(oRequest, Minimize:=True), sResponse) And JsonParse(sResponse, oJson) Then
                If Not JsonItem(oJson, "Ok") Then
                    ConsoleError STR_INFO_ERROR_ACCESSING & vbCrLf, vKey, JsonItem(oJson, "ErrorText")
                Else
                    sKey = JsonItem(oJson, "DeviceSerialNo")
                    If LenB(sKey) <> 0 Then
                        JsonItem(oRetVal, sKey) = sDeviceString
                        JsonItem(oRetVal, "__" & vKey) = sDeviceString
                        JsonItem(oRetVal, "Count") = JsonItem(oRetVal, "Count") + 1
                        JsonItem(oRetVal, "Alias") = JsonItem(oRetVal, "Alias") + 1
                    End If
                End If
            End If
        End If
    Next
    Set pvCollectPrinters = oRetVal
    Exit Function
EH:
    ConsoleError "Critical error: " & Err.Description & vbCrLf
    Resume Next
End Function
