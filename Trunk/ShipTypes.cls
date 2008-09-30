VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ShipTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"ShipType"
Attribute VB_Ext_KEY = "Member0" ,"ShipType"
Option Explicit

Private mcolShipTypes As Collection

Public Function Add(ByVal objShipType As ShipType) As ShipType
    mcolShipTypes.Add objShipType, objShipType.ShipType
End Function

Public Property Get Item(vntIndexKey As Variant) As ShipType
Attribute Item.VB_UserMemId = 0
    Set Item = mcolShipTypes(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mcolShipTypes.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    mcolShipTypes.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mcolShipTypes.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mcolShipTypes = New Collection
End Sub

Private Sub Class_Terminate()
    Set mcolShipTypes = Nothing
End Sub

