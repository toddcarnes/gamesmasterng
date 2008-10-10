VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Planet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Planet"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'********************************************************
'   Copyright 2007,2008 Ian Evans.                      *
'   This program is distributed under the terms of the  *
'       GNU General Public License.                     *
'********************************************************
Option Explicit

Public PlanetID As Long
Public Owner As Race
Public Routes As Variant
Public RouteIDs As Variant

Public Planet As String
Public OwnerID As Long
Public X As Single
Public Y As Single
Public Size As Single
Public Resources As Single
Public Pop As Single
Public Ind As Single
Public Producing As String
Public Cap As Single
Public Mat As Single
Public Col As Single
Public InProgress As String
Public Spent As Single
Public Flags As Long

Private Sub Class_Initialize()
    Dim i As Long
    
    ReDim Routes(4)
    ReDim RouteIDs(4)
    
    For i = 1 To 4
        Set Routes(i) = Nothing
        RouteIDs(i) = 0
    Next i
End Sub

Private Sub Class_Terminate()
    Call Terminate
End Sub

Public Sub Terminate()
    Routes = Empty
End Sub
