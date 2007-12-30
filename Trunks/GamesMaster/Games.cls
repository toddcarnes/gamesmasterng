VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Games"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "A Collection of Game that are the implementation of Template"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"Game"
Attribute VB_Ext_KEY = "Member0" ,"Game"
Option Explicit

'local variable to hold collection
Private mCol As Collection

Public Property Get Item(vntIndexKey As Variant) As Game
Attribute Item.VB_UserMemId = 0
    On Error Resume Next
    Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property

Public Function Refresh()
    Dim strFile As String
    Dim objGame As Game
    
    Set mCol = New Collection
    
    ' Load templates
    strFile = Dir(Options.GalaxyNGHome & "*.glx")
    
    While strFile <> ""
        If Right(strFile, 1) <> "~" Then
            Set objGame = New Game
            With objGame
                .TemplateFile = Options.GalaxyNGHome & strFile
                .GameName = GetFileName(strFile)
                mCol.Add objGame, .GameName
            End With
        End If
        strFile = Dir()
    Wend
    
    'Load Created Games
    strFile = Dir(Options.GalaxyNGData & "*.*", vbDirectory)
    
    While strFile <> ""
        If strFile <> "." And strFile <> ".." Then
            Set objGame = Me.Item(strFile)
            If objGame Is Nothing Then
                Set objGame = New Game
                With objGame
                    .GameName = strFile
                    mCol.Add objGame, strFile
                End With
            End If
        End If
        strFile = Dir()
    Wend

    For Each objGame In mCol
        If Dir(objGame.GameFile(0)) <> "" Then
            objGame.Created = True
        End If
    Next objGame

End Function

Private Sub Class_Initialize()
    Call Refresh
End Sub


Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

