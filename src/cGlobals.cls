VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGlobals"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=========================================================================
'
' UcsFP20 (c) 2008-2020 by Unicontsoft
'
' Unicontsoft Fiscal Printers Component 2.0
'
' This project is licensed under the terms of the MIT license
' See the LICENSE file in the project root for more information
'
'=========================================================================
'
' Global functions
'
'=========================================================================
Option Explicit
DefObj A-Z
'Private Const MODULE_NAME As String = "cGlobals"

Public Function JsonParse(sText As String, vResult As Variant, Optional Error As String, Optional ByVal StrictMode As Boolean, Optional LastPos As Long) As Boolean
    JsonParse = mdJson.JsonParse(sText, vResult, Error, StrictMode, LastPos)
End Function

Public Function JsonParseObject(sText As String, Optional Error As String, Optional ByVal StrictMode As Boolean) As Object
    Set JsonParseObject = mdJson.JsonParseObject(sText, Error, StrictMode)
End Function

Public Function JsonDump(vJson As Variant, Optional ByVal Level As Long, Optional ByVal Minimize As Boolean, Optional CompoundChars As String) As String
    JsonDump = mdJson.JsonDump(vJson, Level, Minimize, CompoundChars)
End Function

Public Function JsonKeys(oJson As Object, Optional sKey As String) As Variant
    JsonKeys = mdJson.JsonKeys(oJson, sKey)
End Function

Public Property Get JsonItem(oJson As Object, ByVal sKey As String) As Variant
    AssignVariant JsonItem, mdJson.JsonItem(oJson, sKey)
End Property

Public Property Let JsonItem(oJson As Object, ByVal sKey As String, vValue As Variant)
    mdJson.JsonItem(oJson, sKey) = vValue
End Property

Public Function JsonToXmlDocument(vJson As Variant, Optional Root As Object, Optional Doc As Object) As Object
    Set JsonToXmlDocument = mdJson.JsonToXmlDocument(vJson, Root, Doc)
End Function

Public Function JsonFromXmlDocument(vXml As Variant) As Variant
    AssignVariant JsonFromXmlDocument, mdJson.JsonFromXmlDocument(vXml)
End Function

Public Function JsonToDictionary(oJson As Object) As Object
    Set JsonToDictionary = mdJson.JsonToDictionary(oJson)
End Function

Public Property Get Logger() As Object
    Set Logger = mdGlobals.Logger
End Property

Public Property Set Logger(oValue As Object)
    Set mdGlobals.Logger = oValue
End Property

Public Sub FlushDebugLog()
    mdGlobals.FlushDebugLog
End Sub

Public Property Get ProtocolConfig() As Object
    
End Property

Public Property Set ProtocolConfig(oValue As Object)
    SetProtocolConfigRoot oValue
End Property

Public Function EnumSerialPorts() As Variant
    AssignVariant EnumSerialPorts, mdGlobals.EnumSerialPorts
End Function

Public Function ParseDeviceString(ByVal sDeviceString As String, Optional Separator As String = ";") As Object
    Set ParseDeviceString = mdGlobals.ParseDeviceString(sDeviceString, Separator)
End Function

Public Function ToDeviceString(oMap As Object, Optional Separator As String = ";") As String
    ToDeviceString = mdGlobals.ToDeviceString(oMap, Separator)
End Function

Public Function InitRegExp(sPattern As String) As Object
    Set InitRegExp = mdGlobals.InitRegExp(sPattern)
End Function

Public Function SearchCollection(ByVal pCol As Object, Index As Variant, Optional RetVal As Variant) As Boolean
    SearchCollection = mdGlobals.SearchCollection(pCol, Index, RetVal)
End Function
