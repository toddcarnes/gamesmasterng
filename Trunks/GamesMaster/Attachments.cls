VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Attachments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"Attachment"
Attribute VB_Ext_KEY = "Member0" ,"Attachment"
Option Explicit

Private mcolAttachments As Collection

Public Function Add(objAttachment As Attachment) As Attachment
    mcolAttachments.Add objAttachment
    Set Add = objAttachment
End Function

Public Property Get Item(vntIndexKey As Variant) As Attachment
Attribute Item.VB_UserMemId = 0
  Set Item = mcolAttachments(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mcolAttachments.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    mcolAttachments.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mcolAttachments.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mcolAttachments = New Collection
End Sub

Private Sub Class_Terminate()
    Set mcolAttachments = Nothing
End Sub


