VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Messages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"Message"
Attribute VB_Ext_KEY = "Member0" ,"Message"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mcolMessages As Collection
Private mColIndex As Collection

Public Function Add(ByVal objMessage As Message)
    mcolMessages.Add objMessage, objMessage.Key
End Function

Public Property Get Item(vntIndexKey As Variant) As Message
Attribute Item.VB_UserMemId = 0
    Set Item = mcolMessages(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mcolMessages.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    mcolMessages.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mcolMessages.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mcolMessages = New Collection
End Sub

Private Sub Class_Terminate()
    Set mcolMessages = Nothing
End Sub

Public Sub Load()
    Dim strMessage As String
    Dim lngNo As Long
    Dim strCategory As String
    Dim strKey As String
    Dim intFN As Integer
    Dim i As Long
    Dim strLine As String
    Dim blnText As Boolean
    Dim objMessage As Message
    
    Set mcolMessages = New Collection
    intFN = FreeFile
    Open App.Path & "\" & App.EXEName & ".txt" For Input As intFN
    blnText = False
    
    While Not EOF(intFN)
        Line Input #intFN, strLine
        strLine = Trim(strLine)
        If blnText Then
            If strLine = "@" Then
                blnText = False
                Set objMessage = New Message
                objMessage.Category = strCategory
                objMessage.Key = strKey
                objMessage.Message = strMessage
                Add objMessage
            Else
                strMessage = strMessage & strLine & vbNewLine
            End If
        Else
            i = InStr(1, strLine, " ")
            strCategory = Val(Left(strLine, i - 1))
            strKey = Trim(Mid(strLine, i + 1))
            strMessage = ""
            blnText = True
        End If
    Wend
    Close #intFN
End Sub

Public Sub Save()
    Dim strMessage As String
    Dim i As Long
    Dim intFN As Integer
    Dim objMessage As Message
    Dim strBackupFileName As String
    Dim strFileName As String
    
    ' remove the backup file name
    strBackupFileName = App.Path & "\" & App.EXEName & ".bak"
    If Dir(strBackupFileName) <> "" Then
        Kill strBackupFileName
    End If
    
    'Rename the file name as a backup
    strFileName = App.Path & "\" & App.EXEName & ".txt"
    Name strFileName As strBackupFileName
    
    'save the messages
    intFN = FreeFile
    Open App.Path & "\" & App.EXEName & ".txt" For Output As intFN
    i = 0
    For Each objMessage In mcolMessages
        i = i + 1
        Print #intFN, CStr(i) & " " & objMessage.Key
        If Right(objMessage.Message, 2) = vbNewLine Then
            Print #intFN, objMessage.Message;
        Else
            Print #intFN, objMessage.Message
        End If
        Print #intFN, "@"
    Next objMessage
    
    Close #intFN
End Sub

Public Function Clone() As Messages
    Dim objMessage As Message
    
    Set Clone = New Messages
    
    For Each objMessage In mcolMessages
        Clone.Add objMessage.Clone
    Next objMessage
End Function
