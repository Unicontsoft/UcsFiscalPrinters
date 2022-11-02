VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5916
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11496
   LinkTopic       =   "Form1"
   ScaleHeight     =   5916
   ScaleWidth      =   11496
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      Height          =   432
      Left            =   2772
      TabIndex        =   6
      Top             =   2352
      Width           =   2448
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   432
      Left            =   2772
      TabIndex        =   5
      Top             =   1764
      Width           =   2448
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   432
      Left            =   2772
      TabIndex        =   4
      Top             =   1176
      Width           =   2448
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   432
      Left            =   2772
      TabIndex        =   3
      Top             =   588
      Width           =   2448
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   432
      Left            =   504
      TabIndex        =   2
      Top             =   1764
      Width           =   1776
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   432
      Left            =   504
      TabIndex        =   1
      Top             =   1176
      Width           =   1776
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   432
      Left            =   504
      TabIndex        =   0
      Top             =   588
      Width           =   1776
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const STR_PROTOCOL_DATECS_FP As String = "DATECS"
Private Const STR_PROTOCOL_TREMOL_FP As String = "TREMOL"
Private Const STR_DEVICE_STRING     As String = "Protocol=" & STR_PROTOCOL_DATECS_FP & ";Port=COM2;Speed=115200"

Private WithEvents m_oFP As cIslProtocol
Attribute m_oFP.VB_VarHelpID = -1
Private m_oDP As IDeviceProtocol

Private Sub Command4_Click()
    Dim oRequest        As Object
    Dim sResponse       As String
    Dim vJson           As Variant
    Dim sError          As String
    
    Set m_oFP = Nothing
    Set m_oDP = m_oFP
    With New cFiscalPrinter
        If .EnumPorts(sResponse) And JsonParse(sResponse, vJson, sError) Then
            Debug.Print JsonDump(vJson)
        End If
        JsonValue(oRequest, "DeviceString") = STR_DEVICE_STRING
        JsonValue(oRequest, "Operator/Code") = "1"
        If .GetDeviceInfo(JsonDump(oRequest), sResponse) And JsonParse(sResponse, vJson, sError) Then
            Debug.Print JsonDump(vJson)
        End If
    End With
End Sub

Private Sub Command5_Click()
    Dim oRequest        As Object
    Dim sResponse       As String
    Dim vJson           As Variant
    Dim sError          As String
    Dim oRow            As Object
    
    Set m_oFP = Nothing
    Set m_oDP = m_oFP
    With New cFiscalPrinter
        JsonValue(oRequest, "DeviceString") = STR_DEVICE_STRING
        JsonValue(oRequest, "Operator/Code") = "1"
        JsonValue(oRequest, "ReceiptType") = ucsFscRcpSale
        JsonValue(oRow, "ItemName") = "Продукт 1"
        JsonValue(oRow, "Price") = 5.23
        JsonValue(oRequest, "Rows/-1") = oRow
        JsonValue(oRequest, "Rows/-1") = Array("Продукт 2", 2, "Б", 1.345)
        Set oRow = Nothing
        JsonValue(oRow, "Amount") = 1.23
        JsonValue(oRow, "PaymentType") = 2
        JsonValue(oRequest, "Rows/-1") = oRow
        JsonValue(oRequest, "Rows/-1") = Array("С карта", 2, 1.5)
        Debug.Print JsonDump(oRequest)
        If .PrintReceipt(JsonDump(oRequest), sResponse) And JsonParse(sResponse, vJson, sError) Then
            Debug.Print JsonDump(vJson)
        End If
    End With
End Sub

Private Sub Command6_Click()
    Dim oRequest        As Object
    Dim sResponse       As String
    Dim vJson           As Variant
    Dim sError          As String
    
    Set m_oFP = Nothing
    Set m_oDP = m_oFP
    With New cFiscalPrinter
        JsonValue(oRequest, "DeviceString") = STR_DEVICE_STRING
        If .GetDailyTotals(JsonDump(oRequest), sResponse) And JsonParse(sResponse, vJson, sError) Then
            Debug.Print JsonDump(vJson)
        End If
    End With
End Sub

Private Sub Command7_Click()
    Dim oRequest        As Object
    Dim sResponse       As String
    Dim vJson           As Variant
    Dim sError          As String
    
    Set m_oFP = Nothing
    Set m_oDP = m_oFP
    With New cFiscalPrinter
        JsonValue(oRequest, "DeviceString") = STR_DEVICE_STRING
        JsonValue(oRequest, "Operator/Code") = "1"
        JsonValue(oRequest, "Amount") = 123
        If .PrintServiceDeposit(JsonDump(oRequest), sResponse) And JsonParse(sResponse, vJson, sError) Then
            Debug.Print JsonDump(vJson)
        End If
        If .GetDeviceStatus(JsonDump(oRequest), sResponse) And JsonParse(sResponse, vJson, sError) Then
            Debug.Print JsonDump(vJson)
        End If
    End With
End Sub

Private Sub Form_Load()
    On Error GoTo EH
    Set m_oFP = New cIslProtocol
    Set m_oDP = m_oFP
    If m_oFP.Init("Port=COM2;Speed=115200") Then
        Debug.Print m_oFP.GetClock
        Debug.Print m_oFP.GetDeviceSerialNo
        Debug.Print m_oFP.GetFiscalMemoryNo
    End If
    Exit Sub
EH:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub Command1_Click()
    On Error GoTo EH
    m_oDP.CancelReceipt
    m_oDP.StartReceipt ucsFscRcpSale, "1", "Оператор 1", "1", "ZK140945-0001-0000123"
    m_oDP.AddPLU "Продукт 1", 5.12
    m_oDP.PrintReceipt
    Exit Sub
EH:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub Command2_Click()
    On Error GoTo EH
    m_oDP.CancelReceipt
    m_oDP.StartReceipt ucsFscRcpReversal, "1", "Оператор 1", "1", "ZK140945-0001-0000123", _
        RevType:=ucsFscRevRefund, RevReceiptNo:="51", RevReceiptDate:=Now, RevFiscalMemoryNo:="50178759"
    m_oDP.AddPLU "Продукт 1", 5, -1
    m_oDP.PrintReceipt
    Exit Sub
EH:
    MsgBox Err.Description, vbCritical
End Sub


Private Sub Command3_Click()
    On Error GoTo EH
    m_oDP.CancelReceipt
    m_oDP.StartReceipt ucsFscRcpCreditNote, "1", "Оператор 1", "1", "ZK140945-0001-0000124", _
        InvDocNo:="137", InvCgTaxNo:="130395814", InvCgVatNo:="BG130395814", InvCgName:="Униконт Софт ООД", InvCgCity:="София", InvCgAddress:="бул. Тотлебен 85-87", InvCgPrsReceive:="В. Висулчев", _
        RevInvoiceNo:="345", RevReceiptNo:="51", RevReceiptDate:=Now, RevFiscalMemoryNo:="50178759", RevReason:="Корекция на количества"
    m_oDP.AddPLU "Продукт 1", 5, -1
    m_oDP.PrintReceipt
    Exit Sub
EH:
    MsgBox Err.Description, vbCritical
End Sub


