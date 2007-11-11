VERSION 5.00
Begin VB.Form frmSendMail 
   Caption         =   "SendMail"
   ClientHeight    =   3690
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8910
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3690
   ScaleWidth      =   8910
   Begin VB.TextBox txtLog 
      Height          =   3075
      Left            =   840
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   180
      Width           =   6915
   End
End
Attribute VB_Name = "frmSendMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Public WithEvents mobjSendMail As SendMail
Attribute mobjSendMail.VB_VarHelpID = -1

Private Sub Form_Load()
    Me.Icon = MainForm.Icon
    Set mobjSendMail = MainForm.SendMail
End Sub

Private Sub Form_Resize()
    txtLog.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mobjSendMail = Nothing
    MainForm.mnuMailShowSendMail.Checked = False
End Sub

Private Sub mobjSendMail_LogData(ByVal strData As String)
    txtLog = txtLog & strData
End Sub

Private Sub mobjSendMail_NewConnection()
    txtLog = Format(Now, "hh:mm:ss dddd, dd mmmm yyyy") & vbNewLine & _
             "------------------------------------------------------------"
End Sub
