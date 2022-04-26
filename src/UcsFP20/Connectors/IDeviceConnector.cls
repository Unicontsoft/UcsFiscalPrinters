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
'
' UcsFP20 (c) 2008-2019 by Unicontsoft
'
' Unicontsoft Fiscal Printers Component 2.0
'
' This project is licensed under the terms of the MIT license
' See the LICENSE file in the project root for more information
'
'=========================================================================
'
' Internal device connector interface
'
'=========================================================================
Option Explicit

Public Enum UcsConnectorErrorsEnum
    ucsErrNoDeviceInfoSet
    ucsErrApiCallFailed
    ucsErrTimeoutWaitingForResponse
End Enum

Public Function GetLastError(Optional ErrNo As UcsConnectorErrorsEnum) As String

End Function

Public Function WriteData(Data() As Byte, Optional ByVal Timeout As Long) As Boolean

End Function

Public Function ReadData(Result() As Byte, Optional ByVal Timeout As Long) As Boolean

End Function
