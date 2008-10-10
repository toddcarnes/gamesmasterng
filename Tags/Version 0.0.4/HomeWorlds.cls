VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HomeWorlds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"HomeWorld"
Attribute VB_Ext_KEY = "Member0" ,"HomeWorld"
'********************************************************
'   Copyright 2007,2008 Ian Evans.                      *
'   This program is distributed under the terms of the  *
'       GNU General Public License.                     *
'********************************************************
Option Explicit

Private mCol As Collection

Public Sub Add(ByVal objHomeworld As HomeWorld)
    mCol.Add objHomeworld
End Sub

Public Property Get Item(vntIndexKey As Variant) As HomeWorld
Attribute Item.VB_UserMemId = 0
    On Error Resume Next
    Set Item = mCol(vntIndexKey)
    
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

Public Function MaxSize() As Long
    Dim objHomeworld As HomeWorld
    Dim lngMax As Long
    
    For Each objHomeworld In mCol
        If objHomeworld.Size > lngMax Then
            lngMax = objHomeworld.Size
        End If
    Next objHomeworld
    
    MaxSize = lngMax
End Function

Public Function TotalSize() As Long
    Dim objHomeworld As HomeWorld
    Dim lngTotal As Long
    
    For Each objHomeworld In mCol
        lngTotal = lngTotal + objHomeworld.Size
    Next objHomeworld
    
    TotalSize = lngTotal
End Function

Public Function Text() As String
    Dim objHomeworld As HomeWorld
    Dim strText As String
    Dim i As Long
    
    For Each objHomeworld In mCol
        i = i + 1
        strText = strText & _
                "    Homeworld " & CStr(i) & " size: " & CStr(objHomeworld.Size) & vbNewLine
    Next objHomeworld
    Text = strText
End Function

