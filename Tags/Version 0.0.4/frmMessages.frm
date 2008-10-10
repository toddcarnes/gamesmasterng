VERSION 5.00
Begin VB.Form frmMessages 
   Caption         =   "Messages"
   ClientHeight    =   5835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10905
   LinkTopic       =   "Form1"
   ScaleHeight     =   5835
   ScaleWidth      =   10905
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frFooter 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   615
      Left            =   3000
      TabIndex        =   2
      Top             =   4980
      Width           =   3015
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   375
         Left            =   1740
         TabIndex        =   4
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   60
         TabIndex        =   3
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.ListBox lstMessages 
      Height          =   1620
      Left            =   60
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   60
      Width           =   3075
   End
   Begin VB.TextBox txtMessage 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Left            =   3180
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   60
      Width           =   2355
   End
End
Attribute VB_Name = "frmMessages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************
'   Copyright 2007,2008 Ian Evans.                      *
'   This program is distributed under the terms of the  *
'       GNU General Public License.                     *
'********************************************************
Option Explicit

Private mcolMessages As Messages
Private mobjMessage As Message

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Set Options.Messages = mcolMessages
    mcolMessages.Save
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Long
    Dim objMessage As Message
    
    Me.Icon = MainForm.Icon
    Call LoadFormSettings(Me)
    Set mcolMessages = Options.Messages.Clone
    
    With lstMessages
        .Clear
        For Each objMessage In mcolMessages
            i = i + 1
            .AddItem objMessage.Key
            .ItemData(.NewIndex) = i
        Next objMessage
    End With
End Sub

Private Sub Form_Resize()
    Dim L As Single, T As Single, W As Single, H As Single
    
    With frFooter
        L = (Me.ScaleWidth - .Width) / 2
        T = Me.ScaleHeight - .Height
        If T < 0 Then T = 0
        .Move L, T
    End With
    With lstMessages
        L = .Left
        T = .Top
        W = .Width
        H = Me.ScaleHeight - frFooter.Height - lstMessages.Top
        If H < 0 Then H = 0
        .Move L, T, W, H
    End With
    With txtMessage
        L = .Left
        T = .Top
        W = Me.ScaleWidth - .Left - lstMessages.Left
        If W < 0 Then W = 0
        H = Me.ScaleHeight - frFooter.Height - lstMessages.Top
        If H < 0 Then H = 0
        .Move L, T, W, H
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveFormSettings(Me)
End Sub

Private Sub lstMessages_Click()
    With lstMessages
        If .ListIndex < 0 Then Exit Sub
        Set mobjMessage = mcolMessages(.ItemData(.ListIndex))
        txtMessage.Text = mobjMessage.Message
    End With
End Sub

Private Sub txtMessage_Validate(Cancel As Boolean)
    If mobjMessage Is Nothing Then Exit Sub
    mobjMessage.Message = txtMessage.Text
End Sub
