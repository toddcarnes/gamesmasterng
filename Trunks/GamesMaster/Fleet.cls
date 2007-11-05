VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Fleet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Fleet"
Attribute VB_Ext_KEY = "Member1" ,"Groups"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public FleetID As Long
Public Fleet As String

'local variable(s) to hold property value(s)
Private mvarGroups As Groups 'local copy
Public Property Set Groups(ByVal vData As Groups)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Groups = Form1
    Set mvarGroups = vData
End Property


Public Property Get Groups() As Groups
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Groups
    Set Groups = mvarGroups
End Property



