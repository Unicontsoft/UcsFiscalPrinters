VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IDeviceConnector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================================================
' $Header: /UcsFiscalPrinter/Src/Connectors/IDeviceConnector.cls 1     25.04.18 10:39 Wqw $
'
'   Unicontsoft Fiscal Printers Project
'   Copyright (c) 2008-2018 Unicontsoft
'
'   Internal device connector interface
'
' $Log: /UcsFiscalPrinter/Src/Connectors/IDeviceConnector.cls $
' 
' 1     25.04.18 10:39 Wqw
' Initial implementation
'
'=========================================================================
Option Explicit

Public Enum UcsConnectorErrorsEnum
    ucsFscCerNone
End Enum

Public Function GetLastError(Optional ErrCategory As UcsConnectorErrorsEnum) As String

End Function

Public Function WriteData(Data() As Byte, Optional ByVal Timeout As Long) As Boolean

End Function

Public Function ReadData(Result() As Byte, Optional ByVal Timeout As Long) As Boolean

End Function
