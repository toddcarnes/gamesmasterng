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
            TextSave        =   "17/10/2007"
            Key             =   "Date"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1402
            MinWidth        =   1411
            TextSave        =   "6:29"
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
      Begin VB.Menu mnuMailShowGetMail 
         Caption         =   "Show Get Mail"
      End
      Begin VB.Menu mnuMailRetreive 
         Caption         =   "Retreive"
      End
      Begin VB.Menu mnuMailProcess 
         Caption         =   "Process"
      End
      Begin VB.Menu mnuMailSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMailShowSendMail 
         Caption         =   "Show Send Mail"
      End
      Begin VB.Menu mnuMailSend 
         Caption         =   "Send"
      End
      Begin VB.Menu mnuMailSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMailAutoCheck 
         Caption         =   "Auto Check Mail"
      End
      Begin VB.Menu mnuAutoRun 
         Caption         =   "Auto RunGames"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mobjGetMail As GetMail
Attribute mobjGetMail.VB_VarHelpID = -1
Private WithEvents mobjSendMail As SendMail
Attribute mobjSendMail.VB_VarHelpID = -1
Private mdtNextMailCheck As Date
Private mdtNextRunCheck As Date

Public Function GetMail() As GetMail
    If mobjGetMail Is Nothing Then
        Set mobjGetMail = New GetMail
    End If
    Set GetMail = mobjGetMail
End Function

Public Function SendMail() As SendMail
    If mobjSendMail Is Nothing Then
        Set mobjSendMail = New SendMail
    End If
    Set SendMail = mobjSendMail
End Function

Private Sub MDIForm_Load()
    With Me
        .Width = 800 * Screen.TwipsPerPixelX
        .Height = 600 * Screen.TwipsPerPixelY
    End With
    With tmrMail
        .Interval = 150
        .Enabled = False
    End With
    With tmrGalaxyNG
        .Interval = 150
        .Enabled = False
    End With
    Status = ""
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        If tmrMail.Enabled Or tmrGalaxyNG.Enabled Then
            Systray.InTray = True
            Me.Hide
            Cancel = -1
            Exit Sub
        End If
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

Private Sub mnuAutoRun_Click()
    If mnuAutoRun.Checked Then
        tmrGalaxyNG.Enabled = False
        mnuAutoRun.Checked = False
    Else
        mdtNextRunCheck = 0
        tmrGalaxyNG.Interval = 150
        tmrGalaxyNG.Enabled = True
        mnuAutoRun.Checked = True
    End If
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

Public Sub RefreshGamesForm()
    Dim fForm As Form
    Dim fGames As frmGames
    
    For Each fForm In Forms
        If fForm.name = "frmGames" Then
            Set fGames = fForm
            Call fGames.LoadGames
        End If
    Next fForm
End Sub
Private Sub mnuMailAutoCheck_Click()
    If mnuMailAutoCheck.Checked Then
        tmrMail.Enabled = False
        mnuMailAutoCheck.Checked = False
    Else
        tmrMail.Interval = 150
        mdtNextMailCheck = 0
        tmrMail.Enabled = True
        mnuMailAutoCheck.Checked = True
    End If
End Sub

Private Sub mnuMailProcess_Click()
    Call ProcessEMails
End Sub

Private Sub mnuMailRetreive_Click()
    GetMail.GetMail
End Sub

Private Sub mnuMailSend_Click()
    SendMail.Send
End Sub

Private Sub mnuMailShowGetMail_Click()
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
    If mnuMailShowGetMail.Checked Then
        mnuMailShowGetMail.Checked = False
        Unload fGetMail
    Else
        mnuMailShowGetMail.Checked = True
        fGetMail.Show
    End If
    
    Set fForm = Nothing
    Set fGetMail = Nothing
End Sub

Private Sub mnuMailShowSendMail_Click()
    Dim fForm As Form
    Dim fSendMail As frmSendMail
    
    For Each fForm In Forms
        If fForm.name = "frmSendMail" Then
            Set fSendMail = fForm
            Exit For
        End If
    Next fForm
    
    If fSendMail Is Nothing Then
        Set fSendMail = New frmSendMail
        Load fSendMail
    End If
    If mnuMailShowSendMail.Checked Then
        mnuMailShowSendMail.Checked = False
        Unload fSendMail
    Else
        mnuMailShowSendMail.Checked = True
        fSendMail.Show
    End If
    
    Set fForm = Nothing
    Set fSendMail = Nothing

End Sub

Private Sub mobjGetMail_Closing()
    Status = "Closing POP Connection"
End Sub

Private Sub mobjGetMail_Connecting(ByVal strServer As String)
    Status = "Connecting to " & strServer
End Sub

Private Sub mobjGetMail_Disconnected()
    Status = ""
    Call ProcessEMails
    DoEvents
    Call SendMail.Send
End Sub

Private Sub mobjGetMail_Receiving(ByVal lngEMail As Long, ByVal lngTotal As Long)
    Status = "Receiving E-Mail " & CStr(lngEMail) & " of " & CStr(lngTotal) & "."
End Sub

Private Sub mobjGetMail_Validating()
    Status = "Signing onto Mail Server"
End Sub

Private Sub mobjSendMail_Closing()
    Status = "Closing SMTP Mail Connection"
End Sub

Private Sub mobjSendMail_Connecting(ByVal strServer As String)
    Status = "Connecting to " & strServer
End Sub

Private Sub mobjSendMail_Disconnected()
    Status = ""
End Sub

Private Sub mobjSendMail_Sending(ByVal lngEMail As Long, ByVal lngTotal As Long)
    Status = "Sending EMail " & CStr(lngEMail) & " of " & CStr(lngTotal) & "."
End Sub

Private Sub SysTray_MouseDblClick(Button As Integer, Id As Long)
    Me.Show
    Me.WindowState = vbNormal
    Systray.InTray = False
End Sub

Private Sub tmrGalaxyNG_Timer()
    Dim objGames As Games
    Dim objGame As Game
    Dim blnMailTimer As Boolean
    Dim blnProcessed As Boolean
    
    tmrGalaxyNG.Interval = 30000
    
    If mdtNextRunCheck < Now Then
        mdtNextMailCheck = DateAdd("n", 5, Now)
        
        blnMailTimer = tmrMail.Enabled
        tmrMail.Enabled = False
        Set objGames = New Games
        For Each objGame In objGames
            objGame.Refresh
            If objGame.ReadyToRun Then
                Call RunGame(objGame.GameName)
                blnProcessed = True
            ElseIf objGame.NotifyUsers Then
                Call NotifyUsers(objGame.GameName)
                blnProcessed = True
            End If
        Next objGame
        tmrMail.Enabled = blnMailTimer
        
        If blnProcessed Then
            GalaxyNG.Games.Refresh
            Call MainForm.RefreshGamesForm
            Call SendMail.Send
        End If
    End If
End Sub

Private Sub tmrMail_Timer()
    tmrMail.Interval = 10000
    If mdtNextMailCheck < Now Then
        mdtNextMailCheck = DateAdd("n", CheckMailInterval, Now)
        GetMail.GetMail
    End If
End Sub

Public Property Get Status() As String
    Status = Mid(StatusBar.Panels(1).Text, 7)
End Property

Public Property Let Status(ByVal strStatus As String)
    If strStatus = "" Then
        StatusBar.Panels(1).Text = ""
    Else
        StatusBar.Panels(1).Text = Format(Now, "hh:nn") & " " & strStatus
    End If
End Property

