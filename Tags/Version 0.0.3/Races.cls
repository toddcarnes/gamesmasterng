VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Races"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "A Collection of the Races playing in the game"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"Race"
Attribute VB_Ext_KEY = "Member0" ,"Race"
Option Explicit
Option Compare Text

'local variable to hold collection
Private mcolRaces As Collection

Public Property Get Item(vntIndexKey As Variant) As Race
Attribute Item.VB_UserMemId = 0
    On Error Resume Next
    Set Item = mcolRaces(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mcolRaces.Count
End Property

Public Sub Add(ByVal Race As Race, ByVal Key As String)
    mcolRaces.Add Race, Key
End Sub

Public Sub Remove(vntIndexKey As Variant)
    mcolRaces.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mcolRaces.[_NewEnum]
End Property


Private Sub Class_Initialize()
    Set mcolRaces = New Collection
End Sub

Private Sub Class_Terminate()
    Dim objRace As Race
    
    For Each objRace In mcolRaces
        Call objRace.Terminate
    Next objRace
    
    Set mcolRaces = Nothing
End Sub

