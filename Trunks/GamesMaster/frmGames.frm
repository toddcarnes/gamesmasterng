VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmGames 
   Caption         =   "Games"
   ClientHeight    =   3990
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   6915
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3990
   ScaleWidth      =   6915
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdGames 
      Height          =   2415
      Left            =   420
      TabIndex        =   0
      Top             =   360
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   4260
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOptions 
         Caption         =   "Options"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
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
End
Attribute VB_Name = "frmGames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private mobjGames As Games

Public Property Get Games() As Games
    Set Games = mobjGames
End Property

Public Property Set Games(ByVal objGames As Games)
    Set mobjGames = objGames
    Call LoadGames
End Property

Public Sub LoadGames()
    Dim lngRow As Long
    Dim objGame As Game
    Dim C As Long
    Dim dtNext As Date
    
    With grdGames
        .Clear
        .AllowUserResizing = flexResizeColumns
        .FocusRect = flexFocusNone
        .SelectionMode = flexSelectionByRow
        .Rows = 2
        .RowHeight(1) = 0
        .Cols = 7
        .FixedCols = 1
        .FixedRows = 1
        .ColSel = 3
        C = 0
        .ColWidth(C) = 16 * Screen.TwipsPerPixelX
        C = C + 1
        .TextMatrix(0, C) = "A"
        .ColWidth(C) = 16 * Screen.TwipsPerPixelX
        C = C + 1
        .TextMatrix(0, C) = "Name"
        .ColWidth(C) = 2000
        C = C + 1
        .TextMatrix(0, C) = "Turn"
        .ColWidth(C) = 800
        .ColAlignment(C) = flexAlignCenterTop
        C = C + 1
        .TextMatrix(0, C) = "Players"
        .ColWidth(C) = 800
        .ColAlignment(C) = flexAlignCenterTop
        C = C + 1
        .TextMatrix(0, C) = "Last Run"
        .ColWidth(C) = 1500
        .ColAlignment(C) = flexAlignLeftTop
        C = C + 1
        .TextMatrix(0, C) = "Next Run"
        .ColWidth(C) = 1500
        .ColAlignment(C) = flexAlignLeftTop
        
        lngRow = 1
        For Each objGame In Games
            objGame.Refresh
            lngRow = lngRow + 1
            If lngRow + 1 > .Rows Then .Rows = lngRow + 1
            C = 1
            .TextMatrix(lngRow, C) = IIf(objGame.Template.ScheduleActive, "S", "")
            C = C + 1
            .TextMatrix(lngRow, C) = objGame.GameName
            C = C + 1
            If objGame.Created Then
                If objGame.Started Then
                    .TextMatrix(lngRow, C) = objGame.Turn
                Else
                    .TextMatrix(lngRow, C) = "-"
                End If
                C = C + 1
                .TextMatrix(lngRow, C) = objGame.PlayersReady & "/" & objGame.ActivePlayers & "/" & objGame.Races.Count
                C = C + 1
                .TextMatrix(lngRow, C) = Format(objGame.LastRunDate, "dd-mmm-yyyy hh:nn")
                C = C + 1
                dtNext = objGame.NextRunDate
                .TextMatrix(lngRow, C) = IIf(dtNext = 0, "", Format(objGame.NextRunDate, "dd-mmm-yyyy hh:nn"))
            Else
                .TextMatrix(lngRow, C) = ""
                C = C + 1
                .TextMatrix(lngRow, C) = objGame.Template.Registrations.Count & "/" & objGame.Template.MaxPlayers
                C = C + 1
                .TextMatrix(lngRow, C) = ""
                C = C + 1
                .TextMatrix(lngRow, C) = ""
            End If
        Next objGame
    End With
End Sub

Private Sub Form_Load()
    Me.Icon = MainForm.Icon
    mnuActions.Visible = False
    mnuGameView.Visible = False
    mnuGameEdit.Visible = False
    mnuGameDelete.Visible = False
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    With grdGames
        .Top = 0
        .Left = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight
    End With
End Sub

Private Sub grdGames_DblClick()
    Call mnuTemplateView_Click
End Sub

Private Sub grdGames_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        Call mnuActions_Click
        PopupMenu mnuActions
    End If
End Sub

Private Sub mnuActions_Click()
    Call mnuTemplate_Click
    Call mnuGame_Click
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

Private Sub mnuFileExit_Click()
    Call MainForm.mnuExit_Click
End Sub

Private Sub mnuFileOptions_Click()
    Call MainForm.mnuFileOptions_Click
End Sub

Private Sub mnuGame_Click()
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
        Else
            mnuGameCreate.Enabled = (objGame.Template.Registrations.Count >= objGame.Template.MinPlayers)
            mnuGameDelete.Enabled = False
            mnuGameEdit.Enabled = False
            mnuGameView.Enabled = False
            mnuGameStart.Enabled = False
            mnuRunTurn.Enabled = False
            mnuGameResend.Enabled = False
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

Private Sub mnuGameStart_Click()
    Call StartGame(SelectedGame)
End Sub

Private Sub mnuGameView_Click()
'
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

Private Sub mnuTemplate_Click()
    Dim strTemplate As String
    Dim objGame As Game
    Dim objTemplate As Template
    
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
    Else
        Set objTemplate = objGame.Template
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
    
    Set objTemplate = Nothing
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
    Set Games = GalaxyNG.Games
    DoEvents
    
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
        Set Games = GalaxyNG.Games
    End If
    Set objGame = Nothing
End Sub

Private Sub mnuTemplateEdit_Click()
    Call GetTemplate(False)
End Sub

Private Sub mnuTemplateRefresh_Click()
    Call LoadGames
End Sub

Private Sub mnuTemplateView_Click()
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
    Dim objTemplate As Template
    
    strTemplate = SelectedGame
    Set objTemplate = GalaxyNG.Games(strTemplate).Template
    ShellOpen objTemplate.Filename
End Sub

Private Sub mnuViewGame_Click()
    Call mnuGameView_Click
End Sub

Private Sub mnuViewTemplate_Click()
    Call mnuTemplateView_Click
End Sub

Public Property Get SelectedGame() As String
    With grdGames
        SelectedGame = .TextMatrix(.Row, 2)
    End With
End Property

Private Sub mnuViewTemplateSourceFile_Click()
    Call mnuTemplateViewSourceFile_Click
End Sub
