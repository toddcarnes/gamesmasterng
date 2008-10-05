VERSION 5.00
Begin VB.Form frmNewTemplate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Template"
   ClientHeight    =   1080
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1080
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtMaxPlayers 
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Tag             =   "1"
      Top             =   540
      Width           =   495
   End
   Begin VB.TextBox txtName 
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   2355
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      Caption         =   "Max Players:"
      Height          =   255
      Index           =   3
      Left            =   300
      TabIndex        =   5
      Top             =   600
      Width           =   1155
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      Caption         =   "Name:"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   180
      Width           =   1215
   End
End
Attribute VB_Name = "frmNewTemplate"
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
Option Compare Text

Private mblnCancelled

Public Property Get Cancelled() As Boolean
    Cancelled = mblnCancelled
End Property

Public Property Get TemplateName() As String
    TemplateName = txtName.Text
End Property

Public Property Get Players() As Long
    Players = Val(txtMaxPlayers)
End Property

Private Sub CancelButton_Click()
    mblnCancelled = True
    Me.Hide
End Sub

Private Sub Form_Load()
    Me.Icon = MainForm.Icon
End Sub

Private Sub OKButton_Click()
    mblnCancelled = False
    Me.Hide
End Sub
