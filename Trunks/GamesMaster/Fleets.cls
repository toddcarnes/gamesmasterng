VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Fleets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"Fleet"
Attribute VB_Ext_KEY = "Member0" ,"Fleet"
Option Explicit

'local variable to hold collection
Private mcolFleets As Collection

Public Function Add(ByVal objFleet As Fleet) As Fleet
    mcolFleets.Add objFleet
End Function

Public Property Get Item(vntIndexKey As Variant) As Fleet
Attribute Item.VB_UserMemId = 0
    Set Item = mcolFleets(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mcolFleets.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    mcolFleets.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mcolFleets.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mcolFleets = New Collection
End Sub

Private Sub Class_Terminate()
    Set mcolFleets = Nothing
End Sub

