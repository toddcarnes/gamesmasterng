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
'********************************************************
'   Copyright 2007,2008 Ian Evans.                      *
'   This program is distributed under the terms of the  *
'       GNU General Public License.                     *
'********************************************************
Option Explicit

Private mcolGames As Games

Public Property Get Games() As Games
    If mcolGames Is Nothing Then
        Set mcolGames = New Games
    End If
    Set Games = mcolGames
End Property

Public Property Set Games(vData As Games)
    Set mcolGames = vData
End Property

Private Sub Class_Terminate()
    Set mcolGames = Nothing
End Sub



