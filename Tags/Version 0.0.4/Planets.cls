VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Planets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"Planet"
Attribute VB_Ext_KEY = "Member0" ,"Planet"
'********************************************************
'   Copyright 2007,2008 Ian Evans.                      *
'   This program is distributed under the terms of the  *
'       GNU General Public License.                     *
'********************************************************
Option Explicit

Private mcolPlanets As Collection

Public Function Add(ByVal objPlanet As Planet) As Planet
    If objPlanet.Planet = "" Then
        mcolPlanets.Add objPlanet
    Else
        mcolPlanets.Add objPlanet, objPlanet.Planet
    End If
End Function

Public Property Get Item(vntIndexKey As Variant) As Planet
Attribute Item.VB_UserMemId = 0
    On Error Resume Next
    Set Item = mcolPlanets(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mcolPlanets.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    mcolPlanets.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mcolPlanets.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mcolPlanets = New Collection
End Sub

Private Sub Class_Terminate()
    Dim objPlanet As Planet
    
    For Each objPlanet In mcolPlanets
        objPlanet.Terminate
    Next objPlanet
    
    Set mcolPlanets = Nothing
End Sub

