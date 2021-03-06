VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Registrations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"Registration"
Attribute VB_Ext_KEY = "Member0" ,"Registration"
'********************************************************
'   Copyright 2007,2008 Ian Evans.                      *
'   This program is distributed under the terms of the  *
'       GNU General Public License.                     *
'********************************************************
Option Explicit

Private mcolRegistrations As Collection

Public Function Add(objRegistration As Registration) As Registration
    mcolRegistrations.Add objRegistration
    Set Add = objRegistration
End Function

Public Property Get Item(vntIndexKey As Variant) As Registration
Attribute Item.VB_UserMemId = 0
    Dim objRegistration As Registration
    On Error Resume Next
    If IsNumeric(vntIndexKey) Then
        Set Item = mcolRegistrations(vntIndexKey)
    Else
        For Each objRegistration In mcolRegistrations
            If GetAddress(objRegistration.EMail) = vntIndexKey Then
                Set Item = objRegistration
                Exit For
            End If
        Next objRegistration
    End If
End Property

Public Property Get Count() As Long
    Count = mcolRegistrations.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    mcolRegistrations.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mcolRegistrations.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mcolRegistrations = New Collection
End Sub

Private Sub Class_Terminate()
    Set mcolRegistrations = Nothing
End Sub

