VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmGames 
   Caption         =   "Games"
   ClientHeight    =   3990
   ClientLeft      =   60
   ClientTop       =   450
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
    Call MainForm.mnuTemplateView_Click
End Sub

Private Sub grdGames_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        Call MainForm.mnuTemplate_Click
        Call MainForm.mnuGame_Click
        'Call MainForm.mnuActions_Click
        PopupMenu MainForm.mnuActions
    End If
End Sub

