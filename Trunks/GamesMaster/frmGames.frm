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
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnutemplate 
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
         Caption         =   "Edit template"
      End
      Begin VB.Menu mnuDeleteTemplate 
         Caption         =   "Delete Template"
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

Private Sub LoadGames()
    Dim lngRow As Long
    Dim objGame As Game
    Dim c As Long
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
        c = 0
        .ColWidth(c) = 16 * Screen.TwipsPerPixelX
        c = c + 1
        .TextMatrix(0, c) = "A"
        .ColWidth(c) = 16 * Screen.TwipsPerPixelX
        c = c + 1
        .TextMatrix(0, c) = "Name"
        .ColWidth(c) = 2000
        c = c + 1
        .TextMatrix(0, c) = "Turn"
        .ColWidth(c) = 800
        .ColAlignment(c) = flexAlignCenterTop
        c = c + 1
        .TextMatrix(0, c) = "Players"
        .ColWidth(c) = 800
        .ColAlignment(c) = flexAlignCenterTop
        c = c + 1
        .TextMatrix(0, c) = "Last Run"
        .ColWidth(c) = 1500
        .ColAlignment(c) = flexAlignLeftTop
        c = c + 1
        .TextMatrix(0, c) = "Next Run"
        .ColWidth(c) = 1500
        .ColAlignment(c) = flexAlignLeftTop
        
        lngRow = 1
        For Each objGame In Games
            objGame.Refresh
            lngRow = lngRow + 1
            If lngRow + 1 > .Rows Then .Rows = lngRow + 1
            c = 1
            .TextMatrix(lngRow, c) = IIf(objGame.Template.ScheduleActive, "S", "")
            c = c + 1
            .TextMatrix(lngRow, c) = objGame.GameName
            c = c + 1
            If objGame.Created Then
                If objGame.Started Then
                    .TextMatrix(lngRow, c) = objGame.Turn
                Else
                    .TextMatrix(lngRow, c) = "-"
                End If
                c = c + 1
                .TextMatrix(lngRow, c) = objGame.ActivePlayers & "/" & objGame.Races.Count
                c = c + 1
                .TextMatrix(lngRow, c) = Format(objGame.LastRunDate, "dd-mmm-yyyy hh:nn")
                c = c + 1
                dtNext = objGame.NextRunDate
                .TextMatrix(lngRow, c) = IIf(dtNext = 0, "", Format(objGame.NextRunDate, "dd-mmm-yyyy hh:nn"))
            Else
                .TextMatrix(lngRow, c) = ""
                c = c + 1
                .TextMatrix(lngRow, c) = objGame.Template.Registrations.Count & "/" & objGame.Template.MaxPlayers
                c = c + 1
                .TextMatrix(lngRow, c) = ""
                c = c + 1
                .TextMatrix(lngRow, c) = ""
            End If
        Next objGame
    End With
End Sub

Private Sub Form_Load()
    mnuActions.Visible = False
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

Private Sub grdGames_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
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
    Unload MainForm
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
'            mnuGameSep1.Enabled = (mnuGameStart.Enabled Or mnuGameRun.Enabled)
        Else
            mnuGameCreate.Enabled = (objGame.Template.Registrations.Count >= objGame.Template.MinPlayers)
            mnuGameDelete.Enabled = False
            mnuGameEdit.Enabled = False
            mnuGameView.Enabled = False
            mnuGameStart.Enabled = False
            mnuRunTurn.Enabled = False
            mnuGameResend.Enabled = False
            mnuGameSep1.Enabled = False
        End If
    End If
    
    ' update the sction menu
    mnuActionSeperator1.Visible = (mnuGameCreate.Enabled Or mnuGameView.Enabled)
    mnuCreateGame.Visible = mnuGameCreate.Enabled
    mnuDeleteGame.Visible = mnuGameDelete.Enabled
    mnuViewGame.Visible = mnuGameView.Enabled
    mnuEditGame.Visible = mnuGameEdit.Enabled
    mnuActionSeperator2.Visible = (mnuGameStart.Enabled Or mnuGameRun.Enabled)
    mnuStartGame.Visible = mnuGameStart.Enabled
    mnuRunTurn.Visible = mnuGameRun.Enabled
    mnuResendReports.Visible = mnuGameResend.Enabled
    
    Set objGame = Nothing
End Sub

Private Sub mnuGameCreate_Click()
    Dim strTemplate As String
    Dim objGame As Game
    Dim objtemplate As Template
    
    strTemplate = SelectedGame
    Set objGame = GalaxyNG.Games(strTemplate)
    Set objtemplate = objGame.Template
    Call RunGalaxyNG("-create """ & objGame.TemplateFile & """ >" & strTemplate & ".txt")
End Sub

Private Sub mnuGameDelete_Click()
'
End Sub

Private Sub mnuGameEdit_Click()
'
End Sub

Private Sub mnuGameResend_Click()
    Dim strGame As String
    Dim objGame As Game
    
    strGame = SelectedGame
    Set objGame = GalaxyNG.Games(strGame)
    Call objGame.Refresh
    
    Call SendReports(strGame)
    Call MainForm.SendMail.Send
End Sub

Private Sub mnuGameRun_Click()
    Dim strGame As String
    Dim objGame As Game
    
    strGame = SelectedGame
    GalaxyNG.Games.Refresh
    
    Call RunGame(strGame)
    Call MainForm.SendMail.Send
End Sub

Private Sub mnuGameStart_Click()
    Dim strGame As String
    Dim objGame As Game
    
    strGame = SelectedGame
    GalaxyNG.Games.Refresh
    Set objGame = GalaxyNG.Games(strGame)
    objGame.Refresh
    
    Call RunGalaxyNG("-mail0 " & strGame)
    Call SendReports(strGame)
    Call MainForm.SendMail.Send
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
    Dim objtemplate As Template
    
    strTemplate = SelectedGame
    Set objGame = GalaxyNG.Games(strTemplate)
    mnuTemplateCreate.Enabled = True
    mnuTemplateRefresh.Enabled = True
    If objGame Is Nothing Then
        mnuTemplateDelete.Enabled = False
        mnuTemplateEdit.Enabled = False
        mnuTemplateView.Enabled = False
    Else
        Set objtemplate = objGame.Template
        If objGame.Created Then
            mnuTemplateDelete.Enabled = False
            mnuTemplateEdit.Enabled = True
            mnuTemplateView.Enabled = True
        Else
            mnuTemplateDelete.Enabled = True
            mnuTemplateEdit.Enabled = True
            mnuTemplateView.Enabled = True
        End If
    End If
    mnuCreateTemplate.Visible = mnuTemplateCreate.Enabled
    mnuDeleteTemplate.Visible = mnuTemplateDelete.Enabled
    mnuViewTemplate.Visible = mnuTemplateView.Enabled
    mnuEditTemplate.Visible = mnuTemplateEdit.Enabled
    mnuRefreshTemplate.Visible = mnuTemplateRefresh
    
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

