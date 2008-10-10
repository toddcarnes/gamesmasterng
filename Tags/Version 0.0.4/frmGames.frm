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
'********************************************************
'   Copyright 2007,2008 Ian Evans.                      *
'   This program is distributed under the terms of the  *
'       GNU General Public License.                     *
'********************************************************
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
    Dim c As Long
    Dim dtNext As Date
    
    Games.Refresh
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
        
        Call LoadGridSettings(grdGames, Me.Name)

'        .Visible = False
        lngRow = 1
        For Each objGame In Games
            DoEvents
            If Not objGame.Template.Finished _
            Or (objGame.Template.Finished And MainForm.mnuTemplateShowAll.Checked) Then
                objGame.Refresh
                lngRow = lngRow + 1
                If lngRow + 1 > .Rows Then .Rows = lngRow + 1
                Dim lngForeColour As Long
                Dim lngBackColour As Long
                If objGame.Template.Finished Then
                    lngForeColour = vbWhite
                    lngBackColour = vbRed
                ElseIf objGame.Started Then
                    lngForeColour = vbBlack
                    lngBackColour = vbWhite
                ElseIf objGame.Created Then
                    lngForeColour = vbBlack
                    lngBackColour = vbCyan
                Else
                    lngForeColour = vbWhite
                    lngBackColour = vbBlue
                End If
                For c = 1 To .Cols - 1
                    .Row = lngRow
                    .Col = c
                    .CellBackColor = lngBackColour
                    .CellForeColor = lngForeColour
                Next c
                c = 1
                .TextMatrix(lngRow, c) = IIf(objGame.Template.ScheduleActive, "S", "")
                c = c + 1
                .TextMatrix(lngRow, c) = objGame.GameName
                c = c + 1
                If objGame.Created Then
                    If objGame.Started Then
                        .TextMatrix(lngRow, c) = objGame.Turn
                    Else
                        .TextMatrix(lngRow, c) = "Created"
                    End If
                    c = c + 1
                    .TextMatrix(lngRow, c) = objGame.PlayersReady & "/" & objGame.ActivePlayers & "/" & objGame.Races.Count
                    c = c + 1
                    .TextMatrix(lngRow, c) = Format(objGame.LastRunDate, "dd-mmm-yyyy hh:nn")
                Else
                    .TextMatrix(lngRow, c) = ""
                    c = c + 1
                    .TextMatrix(lngRow, c) = objGame.Template.Registrations.Count & "/" & objGame.Template.MaxPlayers
                    c = c + 1
                    .TextMatrix(lngRow, c) = ""
                End If
                c = c + 1
                If objGame.Template.ScheduleActive Then
                    dtNext = objGame.NextRunDate
                    .TextMatrix(lngRow, c) = IIf(dtNext = 0, "", Format(objGame.NextRunDate, "dd-mmm-yyyy hh:nn"))
                Else
                    .TextMatrix(lngRow, c) = ""
                End If
            End If
        Next objGame
        .Visible = True
        .Col = 1
        .Row = 1
        .ColSel = 1
        .RowSel = 1
    End With
End Sub

Private Sub Form_Load()
    Me.Icon = MainForm.Icon
    Call LoadFormSettings(Me)
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

Private Sub Form_Unload(Cancel As Integer)
    Call SaveFormSettings(Me)
    Call SaveGridSettings(grdGames, Me.Name)
End Sub

Private Sub grdGames_DblClick()
    Dim strGame As String
    
    strGame = MainForm.SelectedGame
    If GalaxyNG.Games(strGame).Created Then
        Call MainForm.mnuGameView_Click
    Else
        Call MainForm.mnuTemplateView_Click
    End If
End Sub

Private Sub grdGames_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        Call MainForm.mnuTemplate_Click
        Call MainForm.mnuGame_Click
        PopupMenu MainForm.mnuActions
    End If
End Sub

