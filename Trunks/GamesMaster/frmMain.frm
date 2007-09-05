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
   Begin VB.Menu mnMail 
      Caption         =   "Mail"
      Begin VB.Menu mnuMailRetreive 
         Caption         =   "Retreive"
      End
      Begin VB.Menu mnuMailShow 
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

Private mobjGetMail As GetMail
Attribute mobjGetMail.VB_VarHelpID = -1

Public Function GetMail() As GetMail
    If mobjGetMail Is Nothing Then
        Set mobjGetMail = New GetMail
    End If
    Set GetMail = mobjGetMail
End Function

Private Sub MDIForm_Load()
    With Me
        .Width = 800 * Screen.TwipsPerPixelX
        .Height = 600 * Screen.TwipsPerPixelY
    End With
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuGames_Click()
    Dim fForm As Form
    Dim fGames As frmGames
    
    For Each fForm In Forms
        If fForm.name = "frmGames" Then
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

Private Sub mnuMailRetreive_Click()
    GetMail.GetMail
End Sub

Private Sub mnuMailShow_Click()
    Dim fForm As Form
    Dim fGetMail As frmGetMail
    
    For Each fForm In Forms
        If fForm.name = "frmGetMail" Then
            Set fGetMail = fForm
            Exit For
        End If
    Next fForm
    
    If fGetMail Is Nothing Then
        Set fGetMail = New frmGetMail
        Load fGetMail
    End If
    If mnuMailShow.Checked Then
        mnuMailShow.Checked = False
        Unload fGetMail
    Else
        mnuMailShow.Checked = True
        fGetMail.Show
    End If
    
    Set fForm = Nothing
    Set fGetMail = Nothing
End Sub
