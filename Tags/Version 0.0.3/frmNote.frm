VERSION 5.00
Begin VB.Form frmNote 
   Caption         =   "Notes"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   6135
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frFooter 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   1860
      TabIndex        =   1
      Top             =   3120
      Width           =   2775
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   375
         Left            =   420
         TabIndex        =   3
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.TextBox txtNote 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   300
      Width           =   4275
   End
End
Attribute VB_Name = "frmNote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mstrNote As String
Dim mblnCancelled As Boolean

Public Property Get Text() As String
    Text = txtNote.Text
End Property

Public Property Let Text(ByVal strText As String)
    mstrNote = strText
    txtNote.Text = strText
End Property

Public Property Get Cancelled() As Boolean
    Cancelled = mblnCancelled
End Property

Private Sub cmdCancel_Click()
    mblnCancelled = True
    Me.Hide
End Sub

Private Sub cmdClose_Click()
    mblnCancelled = False
    Me.Hide
End Sub

Private Sub Form_Load()
    Me.Icon = MainForm.Icon
    Call LoadFormSettings(Me)
End Sub

Private Sub Form_Resize()
    Dim L As Single, T As Single, H As Single, W As Single
    
    L = 0
    T = 0
    W = Me.ScaleWidth
    H = Me.ScaleHeight - frFooter.Height
    txtNote.Move L, T, W, H
    
    L = (Me.ScaleWidth - frFooter.Width) / 2
    T = H
    W = frFooter.Width
    H = frFooter.Height
    frFooter.Move L, T, W, H
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveFormSettings(Me)
End Sub
