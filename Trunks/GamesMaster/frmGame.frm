VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmGame 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Game"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7215
   LinkTopic       =   "Game"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   7215
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frInitialTechLevels 
      Caption         =   "Initial Tech Levels"
      Height          =   915
      Left            =   60
      TabIndex        =   28
      Top             =   2100
      Width           =   3495
      Begin VB.TextBox txtInitialTechLevel 
         Height          =   315
         Index           =   3
         Left            =   2640
         TabIndex        =   32
         Tag             =   "7"
         Top             =   480
         Width           =   675
      End
      Begin VB.TextBox txtInitialTechLevel 
         Height          =   315
         Index           =   2
         Left            =   1800
         TabIndex        =   31
         Tag             =   "7"
         Top             =   480
         Width           =   675
      End
      Begin VB.TextBox txtInitialTechLevel 
         Height          =   315
         Index           =   1
         Left            =   960
         TabIndex        =   30
         Tag             =   "7"
         Top             =   480
         Width           =   675
      End
      Begin VB.TextBox txtInitialTechLevel 
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   29
         Tag             =   "7"
         Top             =   480
         Width           =   675
      End
      Begin VB.Label Label 
         Caption         =   "Cargo"
         Height          =   255
         Index           =   15
         Left            =   2640
         TabIndex        =   36
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label 
         Caption         =   "Shields"
         Height          =   255
         Index           =   14
         Left            =   1800
         TabIndex        =   35
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label 
         Caption         =   "Weapons"
         Height          =   255
         Index           =   13
         Left            =   960
         TabIndex        =   34
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label 
         Caption         =   "Drive"
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   675
      End
   End
   Begin VB.CheckBox chkFinished 
      Height          =   315
      Left            =   6480
      TabIndex        =   11
      Tag             =   "12"
      Top             =   60
      Width           =   255
   End
   Begin VB.CheckBox chkScheduleActive 
      Height          =   315
      Left            =   4620
      TabIndex        =   1
      Tag             =   "12"
      Top             =   60
      Width           =   255
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   435
      Left            =   2160
      TabIndex        =   25
      Top             =   6420
      Width           =   1155
   End
   Begin VB.Frame frRunOptions 
      Caption         =   "Run Options"
      Height          =   1335
      Left            =   4200
      TabIndex        =   22
      Top             =   2100
      Width           =   2955
      Begin VB.TextBox txtScheduleDays 
         Height          =   315
         Left            =   1560
         TabIndex        =   9
         Tag             =   "1"
         Top             =   720
         Width           =   555
      End
      Begin GamesMaster.DateBox dtRunTime 
         Height          =   315
         Left            =   1560
         TabIndex        =   8
         Top             =   300
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DateFormat      =   ""
         TimeFormat      =   "hh:nn"
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Schedule Days:"
         Height          =   255
         Index           =   21
         Left            =   300
         TabIndex        =   24
         Top             =   780
         Width           =   1155
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Run Time:"
         Height          =   255
         Index           =   19
         Left            =   300
         TabIndex        =   23
         Top             =   360
         Width           =   1155
      End
   End
   Begin VB.Frame frRaces 
      Caption         =   "Races"
      Height          =   2835
      Left            =   0
      TabIndex        =   16
      Top             =   3480
      Width           =   7155
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdRaces 
         Height          =   2475
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   6915
         _ExtentX        =   12197
         _ExtentY        =   4366
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame frGalaxy 
      Caption         =   "Galaxy Options"
      Height          =   1575
      Left            =   60
      TabIndex        =   13
      Top             =   480
      Width           =   7095
      Begin VB.CheckBox chkFullBombing 
         Height          =   315
         Left            =   5220
         TabIndex        =   4
         Tag             =   "8"
         Top             =   240
         Width           =   255
      End
      Begin VB.CheckBox chkKeepproduction 
         Height          =   315
         Left            =   5220
         TabIndex        =   5
         Tag             =   "10"
         Top             =   540
         Width           =   255
      End
      Begin VB.CheckBox chkDontDropDead 
         Height          =   315
         Left            =   5220
         TabIndex        =   6
         Tag             =   "11"
         Top             =   840
         Width           =   255
      End
      Begin VB.CheckBox chkSphericalGalaxy 
         Height          =   315
         Left            =   5220
         TabIndex        =   7
         Tag             =   "12"
         Top             =   1140
         Width           =   255
      End
      Begin VB.TextBox txtPeace 
         Height          =   315
         Left            =   1440
         TabIndex        =   3
         Tag             =   "9"
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtSize 
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Tag             =   "1"
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Full Bombing:"
         Height          =   255
         Index           =   6
         Left            =   4080
         TabIndex        =   21
         Top             =   300
         Width           =   1035
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Keep production:"
         Height          =   255
         Index           =   7
         Left            =   3840
         TabIndex        =   20
         Top             =   600
         Width           =   1275
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Don't Drop Dead:"
         Height          =   255
         Index           =   8
         Left            =   3840
         TabIndex        =   19
         Top             =   900
         Width           =   1275
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Spherical Galaxy:"
         Height          =   255
         Index           =   9
         Left            =   3840
         TabIndex        =   18
         Top             =   1200
         Width           =   1275
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Peace:"
         Height          =   255
         Index           =   12
         Left            =   780
         TabIndex        =   17
         Top             =   660
         Width           =   555
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Galaxy Size:"
         Height          =   255
         Index           =   1
         Left            =   420
         TabIndex        =   14
         Top             =   300
         Width           =   915
      End
   End
   Begin VB.TextBox txtName 
      Height          =   315
      Left            =   720
      TabIndex        =   0
      Top             =   60
      Width           =   2355
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   435
      Left            =   3660
      TabIndex        =   12
      Top             =   6420
      Width           =   1155
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      Caption         =   "Game Finished:"
      Height          =   255
      Index           =   27
      Left            =   5100
      TabIndex        =   27
      Top             =   120
      Width           =   1275
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      Caption         =   "Schedule Active:"
      Height          =   255
      Index           =   26
      Left            =   3240
      TabIndex        =   26
      Top             =   120
      Width           =   1275
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      Caption         =   "Name:"
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   15
      Top             =   120
      Width           =   555
   End
   Begin VB.Menu mnuAction 
      Caption         =   "&Races"
      Visible         =   0   'False
      Begin VB.Menu mnuAdd 
         Caption         =   "&Add"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Edit"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
      End
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mobjGame As Game
Private mblnReadOnly As Boolean

Public Property Get Game() As Game
    Set Game = mobjGame
End Property

Public Property Set Game(ByVal objGame As Game)
    Set mobjGame = objGame
    Call LoadGame
End Property

Public Property Let ReadOnly(ByVal blnReadOnly As Boolean)
    mblnReadOnly = blnReadOnly
    If blnReadOnly Then
        cmdClose.Caption = "Close"
        cmdCancel.Visible = False
    Else
        cmdClose.Caption = "Save"
        cmdCancel.Visible = True
    End If
End Property

Public Property Get ReadOnly() As Boolean
    ReadOnly = mblnReadOnly
End Property

Private Sub chkFinished_Click()
    Game.Template.Finished = (chkFinished.Value = vbChecked)
End Sub

Private Sub chkScheduleActive_Click()
    Game.Template.ScheduleActive = (chkScheduleActive.Value = vbChecked)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClose_Click()
    If Not ReadOnly Then
        Call Game.Save
        Call MainForm.RefreshGamesForm
    End If
    Unload Me
End Sub

Private Sub LoadGame()
    Dim i As Long
    
    With mobjGame
        txtName = .GameName
        chkScheduleActive = IIf(.Template.ScheduleActive, vbChecked, vbUnchecked)
        chkFinished = IIf(.Template.Finished, vbChecked, vbUnchecked)
        txtSize = .Size
        chkFullBombing = IIf(.FullBombing, vbChecked, vbUnchecked)
        txtPeace = .Peace
        chkKeepproduction = IIf(.KeepProduction, vbChecked, vbUnchecked)
        chkDontDropDead = IIf(.DontDropDead, vbChecked, vbUnchecked)
        chkSphericalGalaxy = IIf(.sphericalgalaxy, vbChecked, vbUnchecked)
        
        txtInitialTechLevel(Tech.Drive) = .InitialTechLevels(Tech.Drive)
        txtInitialTechLevel(Tech.Weapons) = .InitialTechLevels(Tech.Weapons)
        txtInitialTechLevel(Tech.Shields) = .InitialTechLevels(Tech.Shields)
        txtInitialTechLevel(Tech.Cargo) = .InitialTechLevels(Tech.Cargo)
        dtRunTime.TimeStamp = .RunTime
        txtScheduleDays = .ScheduleDays
    End With
    Call LoadRaces
End Sub

Private Sub dtRunTime_Change()
    Game.RunTime = dtRunTime.TimeStamp
End Sub

Private Sub Form_Load()
    Me.Icon = MainForm.Icon
    Me.Top = 0
    dtRunTime.DateFormat = ""
    dtRunTime.TimeFormat = "hh:nn"
    dtRunTime.TimeStamp = "00:00"
End Sub

Private Sub grdRaces_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbRightButton And Not ReadOnly Then
        PopupMenu mnuAction
    End If
End Sub

Private Sub mnuAction_Click()
    Dim blnEnable As Boolean
    With grdRaces
        blnEnable = (.Row > 1)
        mnuEdit.Enabled = blnEnable
        mnuDelete.Enabled = blnEnable
    End With
End Sub

Private Sub mnuEdit_Click()
    Dim objRace As Race
    Dim fRace As frmRace
    Dim i As Long
    
    i = grdRaces.Row
    If i <= 1 Then Exit Sub
    
    Set objRace = mobjGame.Races(i - 1)
    
    Set fRace = New frmRace
    Set fRace.Race = objRace
    fRace.Show vbModal
    Set fRace = Nothing
    Set objRace = Nothing
    Call LoadRaces
End Sub

Private Sub txtPeace_Change()
    Game.Peace = Val(txtPeace.Text)
End Sub

Private Sub txtScheduleDays_Change()
    Game.ScheduleDays = Val(txtScheduleDays.Text)
End Sub

Private Sub LoadRaces()
    Dim objRace As Race
    Dim i As Long
    Dim c As Long
    
    With grdRaces
        .Clear
        .Rows = mobjGame.Races.Count + 2
        .Cols = 6
        .FixedRows = 1
        .FixedCols = 1
        .RowHeight(1) = 0
        .AllowUserResizing = flexResizeColumns
        .SelectionMode = flexSelectionByRow
        .FocusRect = flexFocusNone
        .ColSel = 5
        .ColWidth(0) = 16 * Screen.TwipsPerPixelX
        .TextMatrix(0, 1) = "E-Mail Address"
        .ColWidth(1) = 4000
        .TextMatrix(0, 2) = "Size 1"
        .ColWidth(2) = 600
        .TextMatrix(0, 3) = "Size 2"
        .ColWidth(3) = 600
        .TextMatrix(0, 4) = "Size 3"
        .ColWidth(4) = 600
        .TextMatrix(0, 5) = "Size 4"
        .ColWidth(5) = 600
        i = 1
        For Each objRace In mobjGame.Races
            i = i + 1
        Next objRace
    End With

End Sub

