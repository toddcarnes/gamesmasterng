VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Messages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"Message"
Attribute VB_Ext_KEY = "Member0" ,"Message"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mColKey As Collection
Private mColIndex As Collection

Public Function Add(ByVal objMessage As Message)
    mColKey.Add objMessage, objMessage.Key
    mColIndex.Add objMessage, objMessage.Index
End Function

Public Property Get Item(vntIndexKey As Variant) As Message
Attribute Item.VB_UserMemId = 0
    If IsNumeric(vntIndexKey) Then
        Set Item = mColIndex(vntIndexKey)
    Else
        Set Item = mColKey(vntIndexKey)
    End If
End Property

Public Property Get Count() As Long
    Count = mColIndex.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    Dim objMessage As Message
    
    If IsNumeric(vntIndexKey) Then
        Set objMessage = mColIndex(vntIndexKey)
    Else
        Set objMessage = mColKey(vntIndexKey)
    End If
    mColIndex.Remove objMessage.Index
    mColKey.Remove objMessage.Key
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mColIndex.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mColKey = New Collection
    Set mColIndex = New Collection
End Sub

Private Sub Class_Terminate()
    Set mColKey = Nothing
    Set mColIndex = Nothing
End Sub

