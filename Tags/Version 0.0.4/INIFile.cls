VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "INIFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' Ini File Functions Class
' Copyright (C) 1996, Jens Balchen
'
' Uses
'
' Exposes
'     Function GetSetting
'     Function SaveSetting
'     Function GetSection
'
' Comments

Option Explicit
Option Compare Text

' --------
'  Public
' --------
'
' Property for file to read
Public File As String

' ---------
'  Private
' ---------
'
' API to read/write ini's
#If Win32 Then
   Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
   Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal Appname As String, ByVal KeyName As Any, ByVal NewString As Any, ByVal Filename As String) As Integer
#Else
   Private Declare Function GetPrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
   Private Declare Function WritePrivateProfileString Lib "Kernel" (ByVal Appname As String, ByVal KeyName As Any, ByVal NewString As Any, ByVal Filename As String) As Integer
#End If

Sub DeleteSection(ByVal Section As String)

Dim retval As Integer

   retval = WritePrivateProfileString(Section, 0&, "", File)

End Sub
Public Function SaveSetting(ByVal Section$, ByVal Key$, ByVal Value$)

Dim retval As Integer

   SaveSetting = WritePrivateProfileString(Section$, Key$, Value$, File)

End Function

Public Function GetSetting(ByVal Section As String, ByVal KeyName As String, Optional ByVal Default As String = "") As String

Dim retval As Integer
Dim t As String * 255

   ' Get the value
   retval = GetPrivateProfileString(Section, KeyName, Default, t, Len(t), File)

   ' If there is one, return it
   If retval > 0 Then
      GetSetting = Left$(t, retval)
   Else
      GetSetting = Default
   End If

End Function

Public Function GetSection(ByVal Section As String, KeyArray() As String) As Integer

Dim retval As Integer
' Allocate space for return value
Dim t As String * 2500
Dim lastpointer As Integer
Dim nullpointer As Integer
Dim ArrayCount As Integer
Dim keystring As String
   
   ReDim KeyArray(0)
   
   ' Get the value
   retval = GetPrivateProfileString(Section, 0&, "", t, Len(t), File)
   
   ' If there is one, return it
   If retval > 0 Then
      '
      ' Separate the keys and store them in the array
      nullpointer = InStr(t, Chr$(0))
      lastpointer = 1
      Do While (nullpointer <> 0 And nullpointer > lastpointer + 1)
         '
         ' Extract key string
         keystring = Mid$(t, lastpointer, nullpointer - lastpointer)
         '
         ' Now add to array
         ArrayCount = ArrayCount + 1
         ReDim Preserve KeyArray(ArrayCount)
         KeyArray(ArrayCount) = keystring
         '
         ' Find next null
         lastpointer = nullpointer + 1
         nullpointer = InStr(nullpointer + 1, t, Chr$(0))
      Loop
   End If
   '
   ' Return the number of array elements
   GetSection = ArrayCount
   
End Function

