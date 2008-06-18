VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDatecsPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================================================
' $Header: /UcsFiscalPrinter/cDatecsPrinter.cls 1     18.06.08 16:45 Wqw $
'
'   Unicontsoft Fiscal Printers Project
'   Copyright (c) 2008 Unicontsoft
'
'   Uprawlenie na DATECS fiscalni printer
'
' $Log: /UcsFiscalPrinter/cDatecsPrinter.cls $
' 
' 1     18.06.08 16:45 Wqw
' Initial implementation
'
'=========================================================================
Option Explicit
Private Const MODULE_NAME As String = "cDatecsPrinter"

'=========================================================================
' Public events
'=========================================================================

Event CommandComplete(ByVal lCmd As Long, sData As String, sResult As String)

'=========================================================================
' Public enums
'=========================================================================

Public Enum UcsStatusBitsEnum
    ucsStbSyntaxError = 2 ^ 0
    ucsStbInvalidCommand = 2 ^ 1
    ucsStbClockNotSet = 2 ^ 2
    ucsStbDeviceError = 2 ^ 4
    ucsStbPrintingError = 2 ^ 5
    ucsStbSumOverflow = 2 ^ (8 + 0)
    ucsStbInvalidFiscalMode = 2 ^ (8 + 1)
    ucsStbRamCleared = 2 ^ (8 + 2)
    ucsStbRamCorrupted = 2 ^ (8 + 4)
    ucsStbOutOfPaper = 2 ^ (16 + 0)
    ucsStbFiscalPrinting = 2 ^ (16 + 3)
    ucsStbNonFiscalPrinting = 2 ^ (16 + 5)
End Enum

Public Enum UcsDipSwitchesEnum
    ucsDipCenterHeaderFooter = 2 ^ 0
    ucsDipPrintHeader = 2 ^ 1
    ucsDipSwitch_1_3 = 2 ^ 2
    ucsDipSwitch_1_4 = 2 ^ 3
    ucsDipZeroMemory = 2 ^ 4
    ucsDipTransparentDisplay = 2 ^ 5
    ucsDipShowToDisplay = 2 ^ 6
End Enum

Public Enum UcsFiscalMemoryEnum
    ucsFmeWriteError = 2 ^ 0
    ucsFmeNoMemoryModule = 2 ^ 2
    ucsFmeLowMemory = 2 ^ 3
    ucsFmeOutOfMemory = 2 ^ 4
    ucsFmeMemoryError = 2 ^ 5
    ucsFmeReadOnly = 2 ^ (8 + 0)
    ucsFmeFormatted = 2 ^ (8 + 1)
    ucsFmeFiscalMode = 2 ^ (8 + 3)
    ucsFmeTaxRatesSet = 2 ^ (8 + 4)
    ucsFmeIdentNoSet = 2 ^ (8 + 5)
End Enum

Public Enum UcsLocalizedIndexesEnum
    ucsLciInternalErrors
    ucsLciPrinterStatuses
    ucsLciPrinterErrors
    ucsLciDipSwitchesTexts
    ucsLciMemoryTexts
End Enum

Public Enum UcsFiscalPrinterCommandsEnum
    '--- init
    ucsFpcInitHeaderFooter = 43
    ucsFpcInitDateTime = 61
    ucsFpcInitInvoiceNo = 66
    ucsFpcInitDecimals = 83
    ucsFpcInitPaymentType = 85
    ucsFpcInitDepartment = 87
    ucsFpcInitTaxRates = 96
    ucsFpcInitBulstat = 98
    ucsFpcInitOperatorPassword = 101
    ucsFpcInitOperatorName = 102
    ucsFpcInitOperatorReset = 104
    ucsFpcInitItem = 107
    '--- transaction
    ucsFpcNonFiscalOpen = 38
    ucsFpcNonFiscalClose = 39
    ucsFpcNonFiscalText = 42
    ucsFpcFiscalOpen = 48
    ucsFpcFiscalSell = 49
    ucsFpcFiscalPayment = 53
    ucsFpcFiscalText = 54
    ucsFpcFiscalClose = 56
    ucsFpcFiscalCgInfo = 57
    ucsFpcFiscalCancel = 60
    ucsFpcAdminCashDebitCredit = 70
    ucsFpcPrintDuplicate = 109
    '--- info
    ucsFpcInfoDateTime = 62
    ucsFpcInfoLastFiscalReceipt = 64
    ucsFpcInfoStatus = 74
    ucsFpcInfoTransaction = 76
    ucsFpcInfoDepartment = 88
    ucsFpcInfoDiagnostics = 90
    ucsFpcInfoTaxRates = 97
    ucsFpcInfoBulstat = 99
    ucsFpcInfoOperator = 112
    ucsFpcInfoLastReceipt = 113
    '--- print
    ucsFpcPrintDailyReport = 69
    ucsFpcPrintDiagnostics = 71
    ucsFpcPrintReportByNumberDetailed = 73
    ucsFpcPrintReportByDateShort = 79
    ucsFpcPrintReportByDateDetailed = 94
    ucsFpcPrintReportByNumberShort = 95
    ucsFpcPrintReportByOperators = 105
    ucsFpcPrintDailyReportItems = 108
    ucsFpcPrintDailyReportDepartments = 117
    ucsFpcPrintDailyReportItemsDepartments = 118
End Enum

'=========================================================================
' API
'=========================================================================

'--- for CreateFile
Private Const GENERIC_READ                  As Long = &H80000000
Private Const GENERIC_WRITE                 As Long = &H40000000
Private Const OPEN_EXISTING                 As Long = 3
Private Const FILE_FLAG_NO_BUFFERING        As Long = &H20000000
'--- for FormatMessage
Private Const FORMAT_MESSAGE_FROM_SYSTEM    As Long = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function SetCommTimeouts Lib "kernel32" (ByVal hFile As Long, lpCommTimeouts As COMMTIMEOUTS) As Long
Private Declare Function BuildCommDCB Lib "kernel32" Alias "BuildCommDCBA" (ByVal lpDef As String, lpDCB As DCB) As Long
Private Declare Function SetCommState Lib "kernel32" (ByVal hCommDev As Long, lpDCB As DCB) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Args As Any) As Long

Private Type COMMTIMEOUTS
    ReadIntervalTimeout         As Long
    ReadTotalTimeoutMultiplier  As Long
    ReadTotalTimeoutConstant    As Long
    WriteTotalTimeoutMultiplier As Long
    WriteTotalTimeoutConstant   As Long
End Type

Private Type DCB
    DCBlength                   As Long
    BaudRate                    As Long
    fBinary                     As Long
    fParity                     As Long
    fOutxCtsFlow                As Long
    fOutxDsrFlow                As Long
    fDtrControl                 As Long
    fDsrSensitivity             As Long
    fTXContinueOnXoff           As Long
    fOutX                       As Long
    fInX                        As Long
    fErrorChar                  As Long
    fNull                       As Long
    fRtsControl                 As Long
    fAbortOnError               As Long
    fDummy2                     As Long
    wReserved                   As Integer
    XonLim                      As Integer
    XoffLim                     As Integer
    ByteSize                    As Byte
    Parity                      As Byte
    StopBits                    As Byte
    XonChar                     As Byte
    XoffChar                    As Byte
    ErrorChar                   As Byte
    EofChar                     As Byte
    EvtChar                     As Byte
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const STR_INTERNAL              As String = "No device info set|CreateFile failed: |SetCommTimeouts failed: |BuildCommDCB failed: |SetCommState failed: |WriteFile failure: |Timeout waiting for response|ReadFile failed: |Invalid message format or checksum (NAK)|Already connected|Not connected|Invalid response length|Response missing postamble|Response missing preamble"
Private Const STR_STATUSES              As String = "Syntax error|Invalid command|Clock not set|3|Feed error|5|6|7|Overflow|Command not allow|Memory reset|3|Memory corrupt|5|6|7|Out of paper|1|2|Fiscal printing|4|Nonfiscal printing|6|7"
Private Const STR_ERRORS_ONLY           As String = "Syntax error|Invalid command|2|3|Feed error|5|6|7|1|Command not allow|Memory reset|3|Memory corrupt|5|6|7|Out of paper|1|2|3|4|5|6|7"
Private Const STR_DIP_SWITCHES          As String = "Center header and footer|Enable header|Sw1.3|Sw1.4|Zero memory|Transparent display|No display|7"
Private Const STR_MEMORY                As String = "Write error|1|No memory module|Memory low|Out of memory|Fiscal memory error|6|7|Read only|Memory formatted|2|Fiscal mode|Tax rates set|Tax number set|6|7"

Private m_sDevice                   As String
Private m_hComm                     As Long
Private m_lTimeout                  As Long
Private m_lSequence                 As Long
Private m_eStatus                   As UcsStatusBitsEnum
Private m_eDip                      As UcsDipSwitchesEnum
Private m_eMemory                   As UcsFiscalMemoryEnum
Private m_sLastError                As String
Private m_bExceptions               As Boolean
Private m_sLocalizedText(0 To 5)    As String

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunc As String)
    Debug.Print MODULE_NAME & "." & sFunc & ": " & Error
End Sub

'=========================================================================
' Properties
'=========================================================================

Property Get Timeout() As Long
    Timeout = m_lTimeout
End Property

Property Let Timeout(ByVal lValue As Long)
    m_lTimeout = lValue
End Property

Property Get Status(ByVal eValue As UcsStatusBitsEnum) As Boolean
    Status = (m_eStatus And eValue) <> 0
End Property

Property Get StatusText() As String
    StatusText = pvGetFlagsText(m_eStatus, LocalizedText(ucsLciPrinterStatuses))
End Property

Property Get ErrorText() As String
    ErrorText = pvGetFlagsText(m_eStatus, LocalizedText(ucsLciPrinterErrors))
End Property

Property Get Dip(ByVal eValue As UcsDipSwitchesEnum) As Boolean
    Dip = (m_eDip And eValue) <> 0
End Property

Property Get DipText() As String
    DipText = pvGetFlagsText(m_eDip, LocalizedText(ucsLciDipSwitchesTexts))
End Property

Property Get Memory(ByVal eValue As UcsFiscalMemoryEnum) As Boolean
    Memory = (m_eMemory And eValue) <> 0
End Property

Property Get MemoryText() As String
    MemoryText = pvGetFlagsText(m_eMemory, LocalizedText(ucsLciMemoryTexts))
End Property

Property Get LastError() As String
    LastError = m_sLastError
End Property

Property Get Exceptions() As Boolean
    Exceptions = m_bExceptions
End Property

Property Let Exceptions(ByVal bValue As Boolean)
    m_bExceptions = bValue
End Property

Property Get Device() As String
    Device = m_sDevice
End Property

Property Get IsConnected() As Boolean
    IsConnected = (m_hComm <> -1)
End Property

Property Get LocalizedText(ByVal eIdx As UcsLocalizedIndexesEnum) As String
    LocalizedText = m_sLocalizedText(eIdx)
    If LenB(LocalizedText) = 0 Then
        Select Case eIdx
        Case ucsLciInternalErrors
            LocalizedText = STR_INTERNAL
        Case ucsLciPrinterStatuses
            LocalizedText = STR_STATUSES
        Case ucsLciPrinterErrors
            LocalizedText = STR_ERRORS_ONLY
        Case ucsLciDipSwitchesTexts
            LocalizedText = STR_DIP_SWITCHES
        Case ucsLciMemoryTexts
            LocalizedText = STR_MEMORY
        End Select
    End If
End Property

Property Let LocalizedText(ByVal eIdx As UcsLocalizedIndexesEnum, sValue As String)
    m_sLocalizedText(eIdx) = sValue
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function EnumPorts() As Variant
    Dim lIdx            As Long
    Dim hComm           As Long
    Dim sError          As String
    Dim sResult         As String
    Dim cPrinters       As Collection
    Dim vResult         As Variant
    
    Set cPrinters = New Collection
    For lIdx = 1 To 9
        sError = vbNullString
        hComm = pvOpenPort("COM" & lIdx & ",9600,8,N,1", sError)
        If hComm <> -1 And LenB(sError) = 0 Then
            If pvWriteData(hComm, ucsFpcInfoStatus, &H20) Then
                sError = vbNullString
                sResult = pvReadData(hComm, 500, sError)
                If LenB(sError) = 0 Then
                    If Asc(Left(sResult, 1)) = 1 Then
                        cPrinters.Add "COM" & lIdx
                    End If
                End If
            End If
            Call CloseHandle(hComm)
        End If
    Next
    If cPrinters.Count = 0 Then
        vResult = Split("")
    Else
        ReDim vResult(0 To cPrinters.Count - 1)
        For lIdx = 1 To cPrinters.Count
            vResult(lIdx - 1) = cPrinters(lIdx)
        Next
    End If
    EnumPorts = vResult
End Function

Public Function Init(sDevice As String, Optional ByVal lTimeout As Long = 3000) As Boolean
    Dim vSplit          As Variant
    
    m_sLastError = vbNullString
    If m_hComm <> -1 Then
        Call CloseHandle(m_hComm)
        m_hComm = -1
    End If
    m_lTimeout = lTimeout
    vSplit = Split(sDevice, ",")
    m_sDevice = Trim(pvAccess(vSplit, 0, "COM1")) & "," & C_Lng(pvAccess(vSplit, 1, "9600")) & "," & C_Lng(pvAccess(vSplit, 2, "8")) & "," & _
        IIf(UCase(pvAccess(vSplit, 3, "N")) = "Y", "Y", "N") & "," & C_Lng(pvAccess(vSplit, 4, "1"))
    m_hComm = pvOpenPort(m_sDevice, m_sLastError)
    If LenB(m_sLastError) <> 0 Then
        GoTo QH
    End If
    '--- success
    Init = True
QH:
    If LenB(m_sLastError) <> 0 And Exceptions Then
        Err.Raise vbObjectError, , m_sLastError
    End If
End Function

Public Function Connect() As Boolean
    m_sLastError = vbNullString
    If m_hComm <> -1 Then
        m_sLastError = pvInternal(9, "Already connected")
        GoTo QH
    End If
    m_hComm = pvOpenPort(m_sDevice, m_sLastError)
    '--- success
    Connect = True
QH:
    If LenB(m_sLastError) <> 0 And Exceptions Then
        Err.Raise vbObjectError, , m_sLastError
    End If
End Function

Public Function Disconnect() As Boolean
    m_sLastError = vbNullString
    If m_hComm = -1 Then
        m_sLastError = pvInternal(10, "Not connected")
        GoTo QH
    End If
    Call CloseHandle(m_hComm)
    m_hComm = -1
    '--- success
    Disconnect = True
QH:
    If LenB(m_sLastError) <> 0 And Exceptions Then
        Err.Raise vbObjectError, , m_sLastError
    End If
End Function

Public Function SendCommand(ByVal eCmd As UcsFiscalPrinterCommandsEnum, Optional sData As String) As String
    Dim sText           As String
    Dim lIdx            As Long
    Dim lSize           As Long
        
    m_sLastError = vbNullString
    If Not pvWriteData(m_hComm, eCmd, m_lSequence, sData, m_sLastError) Then
        GoTo QH
    End If
    '--- increment sequence
    m_lSequence = m_lSequence + 1
    If m_lSequence > &H7F Then
        m_lSequence = &H20
    End If
    '--- check result
    sText = pvReadData(m_hComm, m_lTimeout, m_sLastError)
    If LenB(m_sLastError) = 0 Then
        If Left(sText, 1) = Chr(1) Then
            lIdx = InStr(sText, Chr(4))
            If lIdx > 2 Then
                lSize = Asc(Mid(sText, 2, 1)) - &H20 - 11
                '--- check message length
                If lIdx - 5 = lSize Then
                    pvSetStatus Mid(sText, lIdx + 1, 6)
                    SendCommand = Mid(sText, 5, lSize)
                    GoTo QH
                Else
                    m_sLastError = pvInternal(11, "Invalid response length")
                End If
            Else
                m_sLastError = pvInternal(12, "Response missing postamble")
            End If
        Else
            m_sLastError = pvInternal(13, "Response missing preamble")
        End If
    End If
    pvSetStatus vbNullString '--- clear
QH:
    RaiseEvent CommandComplete(eCmd, sData, SendCommand)
    If Exceptions Then
        If LenB(m_sLastError) <> 0 Then
            Err.Raise vbObjectError, , m_sLastError
        End If
        If Status(ucsStbPrintingError) Then
            Err.Raise vbObjectError, , ErrorText
        End If
    End If
End Function

Public Function ReadData(Optional ByVal lTimeout As Long) As String
    m_sLastError = vbNullString
    ReadData = pvReadData(m_hComm, IIf(lTimeout > 0, lTimeout, m_lTimeout), m_sLastError)
    If LenB(m_sLastError) <> 0 And Exceptions Then
        Err.Raise vbObjectError, , m_sLastError
    End If
End Function

'= private ===============================================================

Private Function pvOpenPort(sDevice As String, Optional sError As String) As Long
    Const FUNC_NAME     As String = "pvOpenPort"
    Dim vSplit          As Variant
    Dim uTimeout        As COMMTIMEOUTS
    Dim uDCB            As DCB
    
    On Error GoTo EH
    vSplit = Split(sDevice, ",")
    If LenB(pvAccess(vSplit, 0)) = 0 Then
        sError = pvInternal(0, "No device info set")
        Exit Function
    End If
    pvOpenPort = CreateFile("\\.\" & pvAccess(vSplit, 0), GENERIC_READ Or GENERIC_WRITE, 0, 0, OPEN_EXISTING, FILE_FLAG_NO_BUFFERING, 0)
    If pvOpenPort = -1 Then
        sError = pvInternal(1, "CreateFile failed: ") & pvGetApiErr(Err.LastDllError)
        Exit Function
    End If
    uTimeout.ReadIntervalTimeout = 65
    uTimeout.ReadTotalTimeoutConstant = 1
    uTimeout.ReadTotalTimeoutMultiplier = 1
    uTimeout.WriteTotalTimeoutConstant = 500
    uTimeout.WriteTotalTimeoutMultiplier = 1
    If SetCommTimeouts(pvOpenPort, uTimeout) = 0 Then
        sError = pvInternal(2, "SetCommTimeouts failed: ") & pvGetApiErr(Err.LastDllError)
        Exit Function
    End If
    If BuildCommDCB("baud=" & pvAccess(vSplit, 1, 9600) & " data=" & pvAccess(vSplit, 2, 8) & " parity=" & pvAccess(vSplit, 3, "N") & " stop=" & pvAccess(vSplit, 4, 1), uDCB) = 0 Then
        sError = pvInternal(3, "BuildCommDCB failed: ") & pvGetApiErr(Err.LastDllError)
        Exit Function
    End If
    If SetCommState(pvOpenPort, uDCB) = 0 Then
        sError = pvInternal(4, "SetCommState failed: ") & pvGetApiErr(Err.LastDllError)
        Exit Function
    End If
    pvSetStatus vbNullString '--- clear
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function pvWriteData(ByVal hComm As Long, ByVal eCmd As UcsFiscalPrinterCommandsEnum, ByVal lSequence As Long, Optional sData As String, Optional sError As String) As Boolean
    Const FUNC_NAME     As String = "pvWriteData"
    Dim sText           As String
    Dim baText()        As Byte
    Dim lIdx            As Long
    Dim lSum            As Long
    Dim lWritten        As Long
    
    On Error GoTo EH
    '--- prepare command and transcode cyrillic
    sText = pvTranscode(sData, True)
    sText = Chr(1) & Chr(&H20 + 4 + Len(sText)) & Chr(lSequence) & Chr(eCmd) & sText & Chr(5)
    '--- calc checksum
    baText = StrConv(sText, vbFromUnicode)
    For lIdx = 1 To UBound(baText)
        lSum = lSum + baText(lIdx)
    Next
    '--- append checksum and terminator
    sText = sText & Chr((lSum \ &H1000 Mod &H10) + &H30) & Chr((lSum \ &H100 Mod &H10) + &H30) & Chr((lSum \ &H10 Mod &H10) + &H30) & Chr((lSum \ &H1 Mod &H10) + &H30) & Chr(3)
    If WriteFile(hComm, ByVal sText, Len(sText), lWritten, 0) = 0 Then
        sError = pvInternal(5, "WriteFile failure: ") & pvGetApiErr(Err.LastDllError)
    Else
        pvWriteData = True
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function pvReadData(ByVal hComm As Long, ByVal lTimeout As Long, Optional sError As String) As String
    Const FUNC_NAME     As String = "pvReadData"
    Dim sBuffer         As String
    Dim dwRead          As Long
    Dim lResult         As Long
    Dim dblTimer        As Double
    
    On Error GoTo EH
    sBuffer = String(1000, 0)
    dblTimer = Timer
    Do
        If Timer - dblTimer > lTimeout / 1000# Then
            sError = pvInternal(6, "Timeout waiting for response")
            Exit Do
        End If
        lResult = ReadFile(hComm, ByVal sBuffer, Len(sBuffer), dwRead, 0)
        If lResult = 0 Then
            sError = pvInternal(7, "ReadFile failed: ") & pvGetApiErr(Err.LastDllError)
            Exit Do
        End If
        Do While dwRead > 0
            Select Case Asc(sBuffer)
            Case &H15 '--- NAK
                sError = pvInternal(8, "Invalid message format or checksum (NAK)")
                Exit Do
            Case &H16 '--- SYN
                dblTimer = Timer
            Case Else
                Exit Do
            End Select
            sBuffer = Mid(sBuffer, 2)
            dwRead = dwRead - 1
        Loop
    Loop While dwRead = 0 And LenB(pvReadData) = 0
'    Call FlushFileBuffers(hComm)
    pvReadData = pvTranscode(Left(sBuffer, dwRead), False)
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function pvTranscode(sText As String, ByVal bToPrinter As Boolean) As String
    Const FUNC_NAME     As String = "pvTranscode"
    Dim lIdx            As Long
    Dim lChar           As Long
    
    On Error GoTo EH
    pvTranscode = sText
    For lIdx = 1 To Len(pvTranscode)
        lChar = Asc(Mid(pvTranscode, lIdx, 1))
        If lChar < &H7F Then
            '--- do nothing
        ElseIf lChar >= &HC0 And lChar <= &HFF And bToPrinter Then
            Mid(pvTranscode, lIdx, 1) = Chr$(lChar - &H40)
        ElseIf lChar >= &H80 And lChar < &HC0 And Not bToPrinter Then
            Mid(pvTranscode, lIdx, 1) = Chr$(lChar + &H40)
        Else
            Mid(pvTranscode, lIdx, 1) = "?"
        End If
    Next
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Sub pvSetStatus(sStatus As String)
    Const FUNC_NAME     As String = "pvSetStatus"
    Dim baStatus()      As Byte
    
    On Error GoTo EH
    baStatus = StrConv(sStatus, vbFromUnicode)
    If UBound(baStatus) >= 2 Then
        Call CopyMemory(m_eStatus, baStatus(0), 3)
    Else
        m_eStatus = 0
    End If
    If UBound(baStatus) >= 3 Then
        Call CopyMemory(m_eDip, baStatus(3), 1)
    Else
        m_eDip = 0
    End If
    If UBound(baStatus) >= 5 Then
        Call CopyMemory(m_eMemory, baStatus(4), 2)
    Else
        m_eMemory = 0
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Function pvGetFlagsText(ByVal lFlags As Long, sTexts As String) As String
    Const FUNC_NAME     As String = "pvGetFlagsText"
    Dim lIdx            As Long
    Dim vSplit          As Variant
    
    On Error GoTo EH
    vSplit = Split(sTexts, "|")
    For lIdx = 0 To UBound(vSplit)
        If (lFlags And (2 ^ lIdx)) <> 0 Then
            If Len(vSplit(lIdx)) > 1 Then
                If LenB(pvGetFlagsText) <> 0 Then
                    pvGetFlagsText = pvGetFlagsText & ", "
                End If
                pvGetFlagsText = pvGetFlagsText & vSplit(lIdx)
            End If
        End If
    Next
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function pvAccess(vData As Variant, ByVal lIdx As Long, Optional sDefault As String) As String
    On Error Resume Next
    pvAccess = sDefault
    pvAccess = C_Str(vData(lIdx))
    On Error GoTo 0
End Function

Private Function C_Lng(v As Variant) As Long
    On Error Resume Next
    C_Lng = CLng(v)
    On Error GoTo 0
End Function

Private Function C_Str(v As Variant) As String
    On Error Resume Next
    C_Str = CStr(v)
    On Error GoTo 0
End Function

Private Function pvGetApiErr(ByVal lLastDllError As Long) As String
    Dim lRet            As Long
   
    pvGetApiErr = Space$(2000)
    lRet = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0&, lLastDllError, 0&, pvGetApiErr, Len(pvGetApiErr), 0&)
    If lRet > 2 Then
        If Mid(pvGetApiErr, lRet - 1, 2) = vbCrLf Then
            lRet = lRet - 2
        End If
    End If
    pvGetApiErr = Left$(pvGetApiErr, lRet)
End Function

Private Function pvInternal(ByVal lIdx As Long, Optional sDefault As String) As String
    pvInternal = pvAccess(Split(LocalizedText(ucsLciInternalErrors), "|"), lIdx, sDefault)
End Function

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Initialize()
    m_hComm = -1
    m_lTimeout = 3000
    m_lSequence = &H20
    m_bExceptions = True
End Sub

Private Sub Class_Terminate()
    If m_hComm <> -1 Then
        Call CloseHandle(m_hComm)
        m_hComm = -1
    End If
End Sub
