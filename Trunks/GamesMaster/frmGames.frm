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
   Begin VB.Menu mnuTemplates 
      Caption         =   "&Templates"
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
   End
   Begin VB.Menu mnuTurn 
      Caption         =   "Turn"
      Begin VB.Menu mnuTurnStart 
         Caption         =   "&Start Game"
      End
      Begin VB.Menu mnuTurnRun 
         Caption         =   "&Run Turn"
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
   End
   Begin VB.Menu mnuTest 
      Caption         =   "Test"
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
    
    With grdGames
        .Clear
        .AllowUserResizing = flexResizeColumns
        .FocusRect = flexFocusNone
        .SelectionMode = flexSelectionByRow
        .FixedCols = 1
        .FixedRows = 1
        .Rows = 2
        .Cols = 4
        .ColSel = 3
        .ColWidth(0) = 16 * Screen.TwipsPerPixelX
        .ColWidth(1) = 2000
        .ColWidth(2) = 800
        .ColAlignment(2) = flexAlignCenterTop
        .ColWidth(3) = 800
        .ColAlignment(3) = flexAlignCenterTop
        .TextMatrix(0, 1) = "Name"
        .TextMatrix(0, 2) = "Turn"
        .TextMatrix(0, 3) = "Players"
        
        lngRow = 0
        For Each objGame In Games
            objGame.Refresh
            lngRow = lngRow + 1
            If lngRow + 1 > .Rows Then .Rows = lngRow + 1
            .TextMatrix(lngRow, 1) = objGame.GameName
            If objGame.Created Then
                .TextMatrix(lngRow, 2) = objGame.Turn
                .TextMatrix(lngRow, 3) = objGame.ActivePlayers & "/" & objGame.Races.Count
            Else
                .TextMatrix(lngRow, 2) = ""
                .TextMatrix(lngRow, 3) = objGame.Template.Registrations.Count & "/" & objGame.Template.MaxPlayers
            End If
        Next objGame
    End With
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
    Dim strGame As String
    Dim objGame As Game
    If Button = vbRightButton Then
        strGame = grdGames.TextMatrix(grdGames.Row, 1)
        Set objGame = GalaxyNG.Games(strGame)
        mnuCreateGame.Visible = Not objGame.Created
        mnuViewGame.Visible = objGame.Created
        mnuEditGame.Visible = objGame.Created
        mnuDeleteGame.Visible = objGame.Created
        mnuActionSeperator2.Visible = objGame.Created
        mnuStartGame.Visible = objGame.Created And (objGame.NextTurn < 0)
        mnuRunTurn.Visible = objGame.Created And (objGame.NextTurn >= 0)
        
        PopupMenu mnuActions
    End If
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
    
    With grdGames
        strTemplate = .TextMatrix(.Row, 1)
    End With
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

Private Sub mnuTemplateView_Click()
    Call GetTemplate(True)
End Sub

Private Sub GetTemplate(Optional ByVal blnReadOnly As Boolean = True)
    Dim fForm As Form
    Dim fTemplate As frmTemplate
    Dim strTemplate As String
    
    With grdGames
        strTemplate = .TextMatrix(.Row, 1)
    End With
    
    For Each fForm In Forms
        If fForm.name = "frmTemplate" Then
            Set fTemplate = fForm
            If fTemplate.Template.TemplateName = strTemplate Then
                Exit For
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

