VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "GalaxyNG Games Master"
   ClientHeight    =   4650
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   8340
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin GamesMaster.cSysTray Systray 
      Left            =   0
      Top             =   0
      _ExtentX        =   900
      _ExtentY        =   900
      InTray          =   0   'False
      TrayIcon        =   "frmMain.frx":0CCA
      TrayTip         =   ""
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
            TextSave        =   "11/12/2007"
            Key             =   "Date"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1402
            MinWidth        =   1411
            TextSave        =   "6:01"
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
      Begin VB.Menu mnuFileOptions 
         Caption         =   "Options"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "View"
      Begin VB.Menu mnuViewGames 
         Caption         =   "&Games"
      End
   End
   Begin VB.Menu mnuTemplate 
      Caption         =   "&Template"
      Begin VB.Menu mnuTemplateCreate 
         Caption         =   "&Create"
      End
      Begin VB.Menu mnuTemplateView 
         Caption         =   "&View"
      End
      Begin VB.Menu mnuTemplateEdit 
         Caption         =   "&Edit"
      End
      Begin VB.Menu mnuTemplateDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuTemplateViewSourceFile 
         Caption         =   "View Source File"
      End
      Begin VB.Menu mnutemplateSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTemplateRefresh 
         Caption         =   "&Refresh"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuGame 
      Caption         =   "&Game"
      Begin VB.Menu mnuGameCreate 
         Caption         =   "&Create"
      End
      Begin VB.Menu mnuGameView 
         Caption         =   "&View"
      End
      Begin VB.Menu mnuGameEdit 
         Caption         =   "&Edit"
      End
      Begin VB.Menu mnuGameDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuGameSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameStart 
         Caption         =   "&Start"
      End
      Begin VB.Menu mnuGameRun 
         Caption         =   "Run Turn"
      End
      Begin VB.Menu mnuGameResend 
         Caption         =   "ReSend Reports"
      End
      Begin VB.Menu mnuGameNotify 
         Caption         =   "Notify Users"
      End
   End
   Begin VB.Menu mnuActions 
      Caption         =   "&Actions"
      Begin VB.Menu mnuCreateTemplate 
         Caption         =   "Create Template"
      End
      Begin VB.Menu mnuViewTemplate 
         Caption         =   "View Template"
      End
      Begin VB.Menu mnuEditTemplate 
         Caption         =   "Edit Template"
      End
      Begin VB.Menu mnuDeleteTemplate 
         Caption         =   "Delete Template"
      End
      Begin VB.Menu mnuViewTemplateSourceFile 
         Caption         =   "View Template Source File"
      End
      Begin VB.Menu mnuRefreshTemplate 
         Caption         =   "Refresh Templates"
      End
      Begin VB.Menu mnuActionSeperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCreateGame 
         Caption         =   "Create Game"
      End
      Begin VB.Menu mnuViewGame 
         Caption         =   "View Game"
      End
      Begin VB.Menu mnuEditGame 
         Caption         =   "Edit Game"
      End
      Begin VB.Menu mnuDeleteGame 
         Caption         =   "Delete Game"
      End
      Begin VB.Menu mnuActionSeperator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStartGame 
         Caption         =   "Start Game"
      End
      Begin VB.Menu mnuRunTurn 
         Caption         =   "Run Turn"
      End
      Begin VB.Menu mnuResendReports 
         Caption         =   "ReSend Reports"
      End
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
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuWindowTileHorizontal 
         Caption         =   "Tile &Horizontal"
      End
      Begin VB.Menu mnuWindowTileVerticle 
         Caption         =   "&Tile Vertical"
      End
      Begin VB.Menu mnuWindowArrangeIcons 
         Caption         =   "&Arrange Icons"
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
    Set Systray.TrayIcon = Me.Icon
    mnuActions.Visible = False
    mnuGameView.Visible = False
    mnuGameEdit.Visible = False
    mnuGameDelete.Visible = False
    With Me
        .Width = 1024 * Screen.TwipsPerPixelX
        .Height = 768 * Screen.TwipsPerPixelY
        If (.Top + .Height) > Screen.Height Then
            .Top = Screen.Height - .Height
        End If
        If (.Left + .Width) > Screen.Width Then
            .Left = Screen.Width - .Width
        End If
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
    If Options.MinimizeatStartup Then
        Me.WindowState = vbMinimized
    End If
    If Options.ShowGames Then
        Call mnuViewGames_Click
    End If
    If Options.ShowSendMail Then
        Call mnuMailShowSendMail_Click
    End If
    If Options.ShowGetMail Then
        Call mnuMailShowGetMail_Click
    End If
    If Options.AutoCheckMail Then
        Call mnuMailAutoCheck_Click
    End If
    If Options.AutoRunGames Then
        Call mnuAutoRun_Click
    End If
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
        If tmrMail.Enabled Or tmrGalaxyNG.Enabled Then
            Systray.InTray = True
            Me.Hide
            Exit Sub
        End If
    End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Systray.InTray = False
    tmrMail.Interval = 0
    tmrGalaxyNG.Interval = 0
End Sub

Public Sub mnuActions_Click()
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

Private Sub mnuCreateGame_Click()
    Call mnuGameCreate_Click
End Sub

Private Sub mnuCreateTemplate_Click()
    Call mnuTemplateCreate_Click
End Sub

Private Sub mnuDeleteGame_Click()
    Call mnuGameDelete_Click
End Sub

Private Sub mnuDeleteTemplate_Click()
    Call mnuTemplateDelete_Click
End Sub

Private Sub mnuEditGame_Click()
    Call mnuGameEdit_Click
End Sub

Private Sub mnuEditTemplate_Click()
    Call mnuTemplateEdit_Click
End Sub

Public Sub mnuExit_Click()
    Unload Me
End Sub

Public Sub mnuFileOptions_Click()
    Dim fForm As Form
    Dim fOptions As frmOptions
    
    For Each fForm In Forms
        If fForm.name = "frmOptions" Then
            Set fOptions = fForm
            Exit For
        End If
    Next fForm
    
    If fOptions Is Nothing Then
        Set fOptions = New frmOptions
        Load fOptions
        fOptions.Show
    Else
        fOptions.Visible = True
        fOptions.WindowState = vbNormal
        fOptions.SetFocus
    End If
    Set fForm = Nothing
    Set fOptions = Nothing
End Sub

Private Sub mnuGameCreate_Click()
    Call CreateGame(SelectedGame)
End Sub

Private Sub mnuGameDelete_Click()
'
End Sub

Private Sub mnuGameEdit_Click()
'
End Sub

Private Sub mnuGameNotify_Click()
    Call NotifyUsers(SelectedGame)
End Sub

Private Sub mnuGameResend_Click()
    Call ResendReports(SelectedGame)
End Sub

Private Sub mnuGameRun_Click()
    Call RunGame(SelectedGame)
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

Private Sub mnuGameStart_Click()
    Call StartGame(SelectedGame)
End Sub

Private Sub mnuGameView_Click()
'
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

Public Sub mnuGame_Click()
    Dim strGame As String
    Dim objGame As Game
    
    strGame = SelectedGame
    Set objGame = GalaxyNG.Games(strGame)
    
    If objGame Is Nothing Then
        mnuGameCreate.Enabled = False
        mnuGameDelete.Enabled = False
        mnuGameView.Enabled = False
        mnuGameEdit.Enabled = False
        mnuGameStart.Enabled = False
        mnuGameRun.Enabled = False
        mnuGameResend.Enabled = False
        mnuGameNotify.Enabled = False
        mnuGameSep1.Enabled = True
    Else
        objGame.Refresh
        If objGame.Created Then
            mnuGameCreate.Enabled = False
            mnuGameDelete.Enabled = True
            mnuGameView.Enabled = True
            mnuGameEdit.Enabled = True
            mnuGameStart.Enabled = Not objGame.Started
            mnuGameRun.Enabled = objGame.Started
            mnuGameResend.Enabled = objGame.Started
            mnuGameNotify.Enabled = objGame.Started
        Else
            mnuGameCreate.Enabled = (objGame.Template.Registrations.Count >= objGame.Template.MinPlayers)
            mnuGameDelete.Enabled = False
            mnuGameEdit.Enabled = False
            mnuGameView.Enabled = False
            mnuGameStart.Enabled = False
            mnuGameRun.Enabled = False
            mnuGameResend.Enabled = False
            mnuGameNotify.Enabled = False
        End If
    End If
    
    ' update the sction menu
    mnuCreateGame.Visible = mnuGameCreate.Enabled And mnuGameCreate.Visible
    mnuDeleteGame.Visible = mnuGameDelete.Enabled And mnuGameDelete.Visible
    mnuViewGame.Visible = mnuGameView.Enabled And mnuGameView.Visible
    mnuActionSeperator1.Visible = (mnuCreateGame.Visible Or mnuViewGame.Visible)
    mnuEditGame.Visible = mnuGameEdit.Enabled And mnuGameEdit.Visible
    mnuStartGame.Visible = mnuGameStart.Enabled And mnuGameStart.Visible
    mnuRunTurn.Visible = mnuGameRun.Enabled And mnuGameRun.Visible
    mnuResendReports.Visible = mnuGameResend.Enabled And mnuGameResend.Visible
    mnuActionSeperator2.Visible = (mnuStartGame.Visible Or mnuRunTurn.Visible Or mnuResendReports.Visible)
    
    Set objGame = Nothing
End Sub

Private Sub mnuRefreshTemplate_Click()
    Call mnuTemplateRefresh_Click
End Sub

Private Sub mnuResendReports_Click()
    Call mnuGameResend_Click
End Sub

Private Sub mnuRunTurn_Click()
    Call mnuGameRun_Click
End Sub

Private Sub mnuStartGame_Click()
    Call mnuGameStart_Click
End Sub

Public Sub mnuTemplate_Click()
    Dim strTemplate As String
    Dim objGame As Game
    Dim objtemplate As Template
    
    strTemplate = SelectedGame
    Set objGame = GalaxyNG.Games(strTemplate)
    mnuTemplateCreate.Enabled = True
    mnuTemplateRefresh.Enabled = True
    If objGame Is Nothing Then
        mnuTemplateDelete.Enabled = False
        mnuTemplateEdit.Enabled = False
        mnuTemplateView.Enabled = False
        mnuTemplateView.Enabled = False
        mnuTemplateViewSourceFile.Enabled = False
        mnuTemplateRefresh.Enabled = (Not GamesForm Is Nothing)
    Else
        Set objtemplate = objGame.Template
        If objGame.Created Then
            mnuTemplateDelete.Enabled = False
            mnuTemplateEdit.Enabled = True
            mnuTemplateView.Enabled = True
            mnuTemplateViewSourceFile.Enabled = True
        Else
            mnuTemplateDelete.Enabled = True
            mnuTemplateEdit.Enabled = True
            mnuTemplateView.Enabled = True
            mnuTemplateViewSourceFile.Enabled = True
        End If
    End If
    mnuCreateTemplate.Visible = mnuTemplateCreate.Enabled And mnuTemplateCreate.Visible
    mnuDeleteTemplate.Visible = mnuTemplateDelete.Enabled And mnuTemplateDelete.Visible
    mnuViewTemplate.Visible = mnuTemplateView.Enabled And mnuTemplateView.Visible
    mnuEditTemplate.Visible = mnuTemplateEdit.Enabled And mnuTemplateEdit.Visible
    mnuRefreshTemplate.Visible = mnuTemplateRefresh.Enabled And mnuTemplateRefresh.Visible
    mnuViewTemplateSourceFile.Visible = mnuTemplateView.Enabled And mnuTemplateView.Visible
    
    Set objtemplate = Nothing
    Set objGame = Nothing
End Sub

Private Sub mnuTemplateCreate_Click()
    Dim fNewTemplate As frmNewTemplate
    Dim fTemplate As frmTemplate
    Dim strTemplate As String
    Dim lngPlayers As Long
    Dim blnCancelled As Boolean
    
    Set fNewTemplate = New frmNewTemplate
    With fNewTemplate
        .Show vbModal
        blnCancelled = .Cancelled
        If Not blnCancelled Then
            strTemplate = .TemplateName
            lngPlayers = .Players
        End If
    End With
    Unload fNewTemplate
    Set fNewTemplate = Nothing
    If blnCancelled Then Exit Sub
    
    Call RunGalaxyNG("-template " & strTemplate & " " & lngPlayers)
    
    GalaxyNG.Games.Refresh
    Set fTemplate = New frmTemplate
    Load fTemplate
    Set fTemplate.Template = GalaxyNG.Games(strTemplate).Template
    fTemplate.Show

End Sub

Private Sub mnuTemplateDelete_Click()
    Dim strTemplate As String
    Dim objGame As Game
    
    strTemplate = SelectedGame
    Set objGame = GalaxyNG.Games(strTemplate)
    If objGame.Created Then
        MsgBox "Game has already been created. " & vbNewLine & _
            "Cannot delete the template as it is required to run the game.", vbOKOnly + vbExclamation, "Delete Template"
    ElseIf vbYes = MsgBox("Are you sure that you wish to delete the template " & strTemplate & "?", vbYesNo + vbQuestion, "Delete Template") Then
        Kill objGame.TemplateFile
        GalaxyNG.Games.Refresh
    End If
    Set objGame = Nothing
End Sub

Private Sub mnuTemplateEdit_Click()
    Call GetTemplate(False)
End Sub

Private Sub mnuTemplateRefresh_Click()
    Call GamesForm.LoadGames
End Sub

Public Sub mnuTemplateView_Click()
    Call GetTemplate(True)
End Sub

Private Sub GetTemplate(Optional ByVal blnReadOnly As Boolean = True)
    Dim fForm As Form
    Dim fTemplate As frmTemplate
    Dim strTemplate As String
    
    strTemplate = SelectedGame
    
    For Each fForm In Forms
        If fForm.name = "frmTemplate" Then
            Set fTemplate = fForm
            If Not fTemplate.Template Is Nothing Then
                If fTemplate.Template.TemplateName = strTemplate Then
                    Exit For
                End If
            End If
            Set fTemplate = Nothing
        End If
    Next fForm
    
    If fTemplate Is Nothing Then
        Set fTemplate = New frmTemplate
        Load fTemplate
        Set fTemplate.Template = GalaxyNG.Games(strTemplate).Template
        fTemplate.Show
    Else
        fTemplate.Visible = True
        fTemplate.WindowState = vbNormal
        fTemplate.SetFocus
    End If
    fTemplate.ReadOnly = blnReadOnly
    Set fForm = Nothing
    Set fTemplate = Nothing
End Sub

Private Sub mnuTemplateViewSourceFile_Click()
    Dim strTemplate As String
    Dim objtemplate As Template
    
    strTemplate = SelectedGame
    Set objtemplate = GalaxyNG.Games(strTemplate).Template
    ShellOpen objtemplate.Filename
End Sub

Private Sub mnuViewGame_Click()
    Call mnuGameView_Click
End Sub

Private Sub mnuViewGames_Click()
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

Private Sub mnuViewTemplate_Click()
    Call mnuTemplateView_Click
End Sub

Private Sub mnuViewTemplateSourceFile_Click()
    Call mnuTemplateViewSourceFile_Click
End Sub

Private Sub mnuWindowArrangeIcons_Click()
    Me.Arrange vbArrangeIcons
End Sub

Private Sub mnuWindowCascade_Click()
    Me.Arrange vbCascade
End Sub

Private Sub mnuWindowTileHorizontal_Click()
    Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuWindowTileVerticle_Click()
    Me.Arrange vbTileVertical
End Sub

Private Sub mobjGetMail_Closing()
    Status = "Closing POP Connection"
End Sub

Private Sub mobjGetMail_Connecting(ByVal strServer As String)
    Status = "Connecting to " & strServer
End Sub

Private Sub mobjGetMail_Disconnected()
    Status = ""
    If ProcessEMails Then
        Call MainForm.RefreshGamesForm
        Call SendMail.Send
    End If
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
    Me.Visible = True
    Systray.InTray = False
End Sub

Private Sub tmrGalaxyNG_Timer()
    Dim objGames As Games
    Dim objGame As Game
    Dim blnGalaxyNGTimer As Boolean
    Dim blnProcessed As Boolean
    
    tmrGalaxyNG.Interval = 30000
    If mdtNextRunCheck < Now Then
        mdtNextRunCheck = DateAdd("n", 5, Now)
        
        blnGalaxyNGTimer = tmrGalaxyNG.Enabled
        tmrGalaxyNG.Enabled = False
        Set objGames = New Games
        For Each objGame In objGames
            objGame.Refresh
            If objGame.Template.ScheduleActive Then
            
                If objGame.ReadyToCreate Then
                    Call CreateGame(objGame.GameName)
                End If
                
                If objGame.ReadyToStart Then
                    Call StartGame(objGame.GameName)
                
                ElseIf objGame.Started Then
                    
                    If objGame.NotifyUsers Then
                        Call NotifyUsers(objGame.GameName)
                        blnProcessed = True
                    
                    ElseIf objGame.ReadyToRun Then
                        Call RunGame(objGame.GameName)
                        blnProcessed = True
                    End If
                End If
        End If
        Next objGame
        tmrGalaxyNG.Enabled = blnGalaxyNGTimer
        
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
        mdtNextMailCheck = DateAdd("n", Options.CheckMailInterval, Now)
        GetMail.GetMail
        If Me.Visible = False Then
            Systray.InTray = False
            Systray.InTray = True
        End If
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

Public Function SelectedGame() As String
    Dim fGames As frmGames

    Set fGames = GamesForm
    If fGames Is Nothing Then Exit Function
    
    With fGames.grdGames
        SelectedGame = .TextMatrix(.Row, 2)
    End With
End Function

Public Function GamesForm() As frmGames
    Dim fForm As Form
    Dim fGames As frmGames
    
    For Each fForm In Forms
        If fForm.name = "frmGames" Then
            Set fGames = fForm
            Exit For
        End If
    Next fForm
    
    Set GamesForm = fGames
End Function
