VERSION 5.00
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "GalaxyNG Games Master"
   ClientHeight    =   4650
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8340
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuGames 
      Caption         =   "&Games"
   End
   Begin VB.Menu mnuGetMail 
      Caption         =   "GetMail"
      Begin VB.Menu mnuCheckMail 
         Caption         =   "Check Mail"
      End
      Begin VB.Menu mnuShowMail 
         Caption         =   "Show"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mobjGetMail As frmGetMail
Attribute mobjGetMail.VB_VarHelpID = -1

Private Sub MDIForm_Load()
    With Me
        .Width = 800 * Screen.TwipsPerPixelX
        .Height = 600 * Screen.TwipsPerPixelY
    End With
End Sub

Private Sub mnuCheckMail_Click()
    GetMail.GetMail
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuGames_Click()
    Dim fForm As Form
    Dim fGames As frmGames
    
    For Each fForm In Forms
        If fForm.Name = "frmGames" Then
            Set fGames = fForm
            Exit For
        End If
    Next fForm
    
    If fGames Is Nothing Then
        Set fGames = New frmGames
        Load fGames
        Set fGames.Games = GalaxyNG.Games
        fGames.Show
    Else
        fGames.Visible = True
        fGames.WindowState = vbNormal
        fGames.SetFocus
    End If
    Set fForm = Nothing
    Set fGames = Nothing
End Sub

Public Property Get GetMail() As frmGetMail
    If mobjGetMail Is Nothing Then
        Set mobjGetMail = New frmGetMail
        Load mobjGetMail
    End If
    Set GetMail = mobjGetMail
End Property

Private Sub mnuShowMail_Click()
    If mnuShowMail.Checked Then
        GetMail.Hide
        mnuShowMail.Checked = False
    Else
        GetMail.Show
        mnuShowMail.Checked = True
    End If
End Sub
