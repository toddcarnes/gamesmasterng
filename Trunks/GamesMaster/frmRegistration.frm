VERSION 5.00
Begin VB.Form frmRegistration 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registration"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   6555
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtY 
      Height          =   315
      Index           =   6
      Left            =   2700
      TabIndex        =   27
      Top             =   3060
      Width           =   555
   End
   Begin VB.TextBox txtY 
      Height          =   315
      Index           =   5
      Left            =   2700
      TabIndex        =   26
      Top             =   2700
      Width           =   555
   End
   Begin VB.TextBox txtY 
      Height          =   315
      Index           =   4
      Left            =   2700
      TabIndex        =   25
      Top             =   2340
      Width           =   555
   End
   Begin VB.TextBox txtY 
      Height          =   315
      Index           =   3
      Left            =   2700
      TabIndex        =   24
      Top             =   1980
      Width           =   555
   End
   Begin VB.TextBox txtY 
      Height          =   315
      Index           =   2
      Left            =   2700
      TabIndex        =   23
      Top             =   1620
      Width           =   555
   End
   Begin VB.TextBox txtY 
      Height          =   315
      Index           =   1
      Left            =   2700
      TabIndex        =   22
      Top             =   1260
      Width           =   555
   End
   Begin VB.TextBox txtY 
      Height          =   315
      Index           =   0
      Left            =   2700
      TabIndex        =   21
      Top             =   900
      Width           =   555
   End
   Begin VB.TextBox txtX 
      Height          =   315
      Index           =   6
      Left            =   2040
      TabIndex        =   20
      Top             =   3060
      Width           =   555
   End
   Begin VB.TextBox txtX 
      Height          =   315
      Index           =   5
      Left            =   2040
      TabIndex        =   19
      Top             =   2700
      Width           =   555
   End
   Begin VB.TextBox txtX 
      Height          =   315
      Index           =   4
      Left            =   2040
      TabIndex        =   18
      Top             =   2340
      Width           =   555
   End
   Begin VB.TextBox txtX 
      Height          =   315
      Index           =   3
      Left            =   2040
      TabIndex        =   17
      Top             =   1980
      Width           =   555
   End
   Begin VB.TextBox txtX 
      Height          =   315
      Index           =   2
      Left            =   2040
      TabIndex        =   16
      Top             =   1620
      Width           =   555
   End
   Begin VB.TextBox txtX 
      Height          =   315
      Index           =   1
      Left            =   2040
      TabIndex        =   15
      Top             =   1260
      Width           =   555
   End
   Begin VB.TextBox txtX 
      Height          =   315
      Index           =   0
      Left            =   2040
      TabIndex        =   14
      Top             =   900
      Width           =   555
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   3690
      TabIndex        =   11
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2010
      TabIndex        =   10
      Top             =   3600
      Width           =   855
   End
   Begin VB.TextBox txtSize 
      Height          =   315
      Index           =   6
      Left            =   1380
      TabIndex        =   9
      Top             =   3060
      Width           =   555
   End
   Begin VB.TextBox txtSize 
      Height          =   315
      Index           =   5
      Left            =   1380
      TabIndex        =   8
      Top             =   2700
      Width           =   555
   End
   Begin VB.TextBox txtSize 
      Height          =   315
      Index           =   4
      Left            =   1380
      TabIndex        =   7
      Top             =   2340
      Width           =   555
   End
   Begin VB.TextBox txtSize 
      Height          =   315
      Index           =   3
      Left            =   1380
      TabIndex        =   6
      Top             =   1980
      Width           =   555
   End
   Begin VB.TextBox txtSize 
      Height          =   315
      Index           =   2
      Left            =   1380
      TabIndex        =   5
      Top             =   1620
      Width           =   555
   End
   Begin VB.TextBox txtSize 
      Height          =   315
      Index           =   1
      Left            =   1380
      TabIndex        =   4
      Top             =   1260
      Width           =   555
   End
   Begin VB.TextBox txtSize 
      Height          =   315
      Index           =   0
      Left            =   1380
      TabIndex        =   3
      Top             =   900
      Width           =   555
   End
   Begin VB.TextBox txtEMailAddress 
      Height          =   315
      Left            =   1380
      TabIndex        =   2
      Top             =   180
      Width           =   4575
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      Caption         =   "Home Planets:"
      Height          =   255
      Index           =   4
      Left            =   60
      TabIndex        =   28
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "Y"
      Height          =   255
      Index           =   3
      Left            =   2760
      TabIndex        =   13
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "X"
      Height          =   255
      Index           =   2
      Left            =   2100
      TabIndex        =   12
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "Size"
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   1
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      Caption         =   "EMail Address:"
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   180
      Width           =   1215
   End
End
Attribute VB_Name = "frmRegistration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

End Sub

Private Sub Label_Click(Index As Integer)

End Sub