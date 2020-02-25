VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2316
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3624
   LinkTopic       =   "Form1"
   ScaleHeight     =   2316
   ScaleWidth      =   3624
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim oConn As cSerialPortConnector
    Dim baData() As Byte
    
    Set oConn = New cSerialPortConnector
    oConn.Init "COM1"
    oConn.WriteData StrConv(Chr$(4), vbFromUnicode)
    If oConn.ReadData(baData(), 100) Then
        Debug.Print "RECV: " & Right$("0" & Hex$(baData(0)), 2), Timer
    End If
End Sub
