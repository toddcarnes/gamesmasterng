VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Registration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"HomeWorlds"
Option Explicit

Public EMail As String
Public RaceName As String
Private mvarHomeWorlds As HomeWorlds

Public Property Get HomeWorlds() As HomeWorlds
    If mvarHomeWorlds Is Nothing Then
        Set mvarHomeWorlds = New HomeWorlds
    End If
    Set HomeWorlds = mvarHomeWorlds
End Property

Public Property Set HomeWorlds(vData As HomeWorlds)
    Set mvarHomeWorlds = vData
End Property

Private Sub Class_Terminate()
    Set mvarHomeWorlds = Nothing
End Sub



