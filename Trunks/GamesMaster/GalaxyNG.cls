VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GalaxyNG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Controlling Class for Controlling GalaxyNG Games by a Gamesmaster"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Games"
Attribute VB_Ext_KEY = "Member1" ,"Registration"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member2" ,"Registrations"
Option Explicit

Private mvarGames As Games

Public Property Get Games() As Games
    If mvarGames Is Nothing Then
        Set mvarGames = New Games
    End If
    Set Games = mvarGames
End Property

Public Property Set Games(vData As Games)
    Set mvarGames = vData
End Property

Private Sub Class_Terminate()
    Set mvarGames = Nothing
End Sub



