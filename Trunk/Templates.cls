VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Templates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Design Templates for Galaxy NG Games"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"Template"
Attribute VB_Ext_KEY = "Member0" ,"Template"
Option Explicit

Private mCol As Collection

Public Property Get Item(vntIndexKey As Variant) As Template
Attribute Item.VB_UserMemId = 0
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property

Public Function Refresh()
    Dim strFile As String
    Dim objTemplate As Template
    Set mCol = New Collection

    strFile = Dir(GalaxyNGHome & "*.glx")
    
    While strFile <> ""
        If Right(strFile, 1) <> "~" Then
            Set objTemplate = New Template
            With objTemplate
                .Filename = strFile
                .TemplateName = GetFileName(strFile)
                mCol.Add objTemplate, .TemplateName
            End With
        End If
        strFile = Dir()
    Wend
End Function

Private Sub Class_Initialize()
    Call Refresh
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

