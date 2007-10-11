VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Public SendTo As String
Public From As String
Public Body As String

Public Sub Load(ByVal strFileName As String)
    Dim strData As String
    Dim lngTo As Long
    Dim lngFrom As Long
    Dim lngEOL As Long
    
    strData = GetFile(strFileName)
    
    lngTo = InStr(1, strData, "To: ") + 4
    lngEOL = InStr(lngTo, strData, vbNewLine)
    SendTo = GetAddress(Mid(strData, lngTo, lngEOL - lngTo))
'    strData = Left(strData, lngTo - 5) & _
                Mid(strData, lngEOL + 2)

    lngFrom = InStr(1, strData, "From: ") + 6
    lngEOL = InStr(lngFrom, strData, vbNewLine)
    From = GetAddress(Mid(strData, lngFrom, lngEOL - lngFrom))
'    strData = Left(strData, lngFrom - 7) & _
                Mid(strData, lngEOL + 2)
    
    Body = strData
End Sub
