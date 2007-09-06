VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "GalaxyNG Games Master"
   ClientHeight    =   4650
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8340
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin GamesMaster.cSysTray Systray 
      Left            =   0
      Top             =   0
      _ExtentX        =   900
      _ExtentY        =   900
      InTray          =   0   'False
      TrayIcon        =   "frmMain.frx":0000
      TrayTip         =   "GalaxyNG - Games Master"
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   4335
      Width           =   8340
      _ExtentX        =   14711
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7064
            Text            =   "Status Messages"
            TextSave        =   "Status Messages"
            Key             =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3519
            MinWidth        =   3528
            Key             =   "Progress"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "6/09/2007"
            Key             =   "Date"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1402
            MinWidth        =   1411
            TextSave        =   "13:56"
            Key             =   "Time"
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrGalaxyNG 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   960
      Top             =   0
   End
   Begin VB.Timer tmrMail 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   540
      Top             =   0
   End
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
Private mdtNextMailCheck As Date

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
    With tmrMail
        .Interval = 10000
        .Enabled = True
    End With
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Systray.InTray = True
        Me.Hide
        Cancel = -1
        Exit Sub
    End If
End Sub

Private Sub MDIForm_Resize()
    If Me.WindowState = vbMinimized Then
        Systray.InTray = True
    End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Systray.InTray = False
    tmrMail.Interval = 0
    tmrGalaxyNG.Interval = 0
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

Private Sub SysTray_MouseDblClick(Button As Integer, Id As Long)
    Me.Show
    Me.WindowState = vbNormal
    Systray.InTray = False
End Sub

Private Sub tmrMail_Timer()
    If mdtNextMailCheck < Now Then
        mdtNextMailCheck = DateAdd("n", 5, Now)
        GetMail.GetMail
    End If
End Sub
