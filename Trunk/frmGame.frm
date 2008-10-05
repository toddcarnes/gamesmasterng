VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmGame 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Game"
   ClientHeight    =   7725
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7215
   LinkTopic       =   "Game"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdViewMap 
      Caption         =   "View Map"
      Height          =   435
      Left            =   180
      TabIndex        =   43
      Top             =   7140
      Width           =   1095
   End
   Begin VB.Frame frInitialTechLevels 
      Caption         =   "Initial Tech Levels"
      Height          =   975
      Left            =   60
      TabIndex        =   28
      Top             =   1980
      Width           =   3495
      Begin VB.TextBox txtInitialTechLevel 
         Height          =   315
         Index           =   3
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   32
         Tag             =   "7"
         Top             =   480
         Width           =   675
      End
      Begin VB.TextBox txtInitialTechLevel 
         Height          =   315
         Index           =   2
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   31
         Tag             =   "7"
         Top             =   480
         Width           =   675
      End
      Begin VB.TextBox txtInitialTechLevel 
         Height          =   315
         Index           =   1
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   30
         Tag             =   "7"
         Top             =   480
         Width           =   675
      End
      Begin VB.TextBox txtInitialTechLevel 
         Height          =   315
         Index           =   0
         Left            =   120
         Locked          =   -1  'True
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
      Left            =   4380
      TabIndex        =   25
      Top             =   7140
      Width           =   1155
   End
   Begin VB.Frame frRunOptions 
      Caption         =   "Run Options"
      Height          =   975
      Left            =   3660
      TabIndex        =   22
      Top             =   1980
      Width           =   3495
      Begin VB.TextBox txtScheduleDays 
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   9
         Tag             =   "1"
         Top             =   540
         Width           =   555
      End
      Begin GamesMaster.DateBox dtRunTime 
         Height          =   315
         Left            =   1560
         TabIndex        =   8
         Top             =   180
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
         Locked          =   -1  'True
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Schedule Days:"
         Height          =   255
         Index           =   21
         Left            =   300
         TabIndex        =   24
         Top             =   600
         Width           =   1155
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Run Time:"
         Height          =   255
         Index           =   19
         Left            =   300
         TabIndex        =   23
         Top             =   240
         Width           =   1155
      End
   End
   Begin VB.Frame frRaces 
      Caption         =   "Races"
      Height          =   3975
      Left            =   0
      TabIndex        =   16
      Top             =   3060
      Width           =   7155
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdRaces 
         Height          =   3615
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   6915
         _ExtentX        =   12197
         _ExtentY        =   6376
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame frGalaxy 
      Caption         =   "Galaxy Options"
      Height          =   1455
      Left            =   60
      TabIndex        =   13
      Top             =   480
      Width           =   7095
      Begin VB.TextBox txtTurn 
         Height          =   315
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   41
         Tag             =   "1"
         Top             =   240
         Width           =   495
      End
      Begin VB.CheckBox chkSaveCopy 
         Height          =   315
         Left            =   6360
         TabIndex        =   39
         Tag             =   "10"
         Top             =   600
         Width           =   255
      End
      Begin VB.CheckBox chkCircle 
         Height          =   315
         Left            =   4320
         TabIndex        =   37
         Tag             =   "10"
         Top             =   960
         Width           =   255
      End
      Begin VB.CheckBox chkFullBombing 
         Height          =   315
         Left            =   4320
         TabIndex        =   4
         Tag             =   "8"
         Top             =   240
         Width           =   255
      End
      Begin VB.CheckBox chkKeepproduction 
         Height          =   315
         Left            =   4320
         TabIndex        =   5
         Tag             =   "10"
         Top             =   600
         Width           =   255
      End
      Begin VB.CheckBox chkDontDropDead 
         Height          =   315
         Left            =   6360
         TabIndex        =   6
         Tag             =   "11"
         Top             =   240
         Width           =   255
      End
      Begin VB.CheckBox chkSphericalGalaxy 
         Height          =   315
         Left            =   6360
         TabIndex        =   7
         Tag             =   "12"
         Top             =   960
         Width           =   255
      End
      Begin VB.TextBox txtPeace 
         Height          =   315
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   3
         Tag             =   "9"
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txtSize 
         Height          =   315
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   2
         Tag             =   "1"
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Turn No:"
         Height          =   255
         Index           =   4
         Left            =   420
         TabIndex        =   42
         Top             =   300
         Width           =   915
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Save Copy:"
         Height          =   255
         Index           =   3
         Left            =   4980
         TabIndex        =   40
         Top             =   660
         Width           =   1275
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Create Circle:"
         Height          =   255
         Index           =   2
         Left            =   2940
         TabIndex        =   38
         Top             =   1020
         Width           =   1275
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Full Bombing:"
         Height          =   255
         Index           =   6
         Left            =   3180
         TabIndex        =   21
         Top             =   300
         Width           =   1035
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Keep production:"
         Height          =   255
         Index           =   7
         Left            =   2940
         TabIndex        =   20
         Top             =   660
         Width           =   1275
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Don't Drop Dead:"
         Height          =   255
         Index           =   8
         Left            =   4980
         TabIndex        =   19
         Top             =   300
         Width           =   1275
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Spherical Galaxy:"
         Height          =   255
         Index           =   9
         Left            =   4980
         TabIndex        =   18
         Top             =   1020
         Width           =   1275
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Peace:"
         Height          =   255
         Index           =   12
         Left            =   780
         TabIndex        =   17
         Top             =   1020
         Width           =   555
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Galaxy Size:"
         Height          =   255
         Index           =   1
         Left            =   420
         TabIndex        =   14
         Top             =   660
         Width           =   915
      End
   End
   Begin VB.TextBox txtName 
      Height          =   315
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   60
      Width           =   2355
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   435
      Left            =   5880
      TabIndex        =   12
      Top             =   7140
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
'********************************************************
'   Copyright 2007,2008 Ian Evans.                      *
'   This program is distributed under the terms of the  *
'       GNU General Public License.                     *
'********************************************************
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
        'Call Game.Template.Save
        'Call Game.Save
        'Call MainForm.RefreshGamesForm
    End If
    Unload Me
End Sub

Private Sub LoadGame()
    Dim i As Long
    
    With mobjGame
        txtName = .GameName
        chkScheduleActive = IIf(.Template.ScheduleActive, vbChecked, vbUnchecked)
        chkFinished = IIf(.Template.Finished, vbChecked, vbUnchecked)
        txtTurn = .Turn
        txtSize = .GalaxySize
        txtPeace = .GalacticPeace
        chkFullBombing = IIf(.flag(G_NONGBOMBING), vbChecked, vbUnchecked)
        chkKeepproduction = IIf(.flag(G_KEEPPRODUCTION), vbChecked, vbUnchecked)
        chkCircle = IIf(.flag(G_CREATECIRCLE), vbChecked, vbUnchecked)
        chkDontDropDead = IIf(.flag(G_NODROP), vbChecked, vbUnchecked)
        chkSaveCopy = IIf(.flag(G_SAVECOPY), vbChecked, vbUnchecked)
        chkSphericalGalaxy = IIf(.flag(G_SPHERICALGALAXY), vbChecked, vbUnchecked)
        
        txtInitialTechLevel(Tech.Drive) = .InitialDrive
        txtInitialTechLevel(Tech.Weapons) = .InitialWeapons
        txtInitialTechLevel(Tech.Shields) = .InitialShield
        txtInitialTechLevel(Tech.Cargo) = .InitialCargo
        dtRunTime.TimeStamp = .Template.RunTime
        txtScheduleDays = .Template.ScheduleDays
    End With
    Call LoadRaces
End Sub

Private Sub cmdViewMap_Click()
    Dim fMap As frmMap
    Dim objPlanet As Planet
    Dim colPlanets As Collection
    
    Set colPlanets = New Collection
    Set fMap = New frmMap
    Load fMap
    fMap.GalaxySize = Game.GalaxySize
    
   
    For Each objPlanet In Game.Planets
        colPlanets.Add objPlanet
    Next objPlanet
    
    Set fMap.Planets = colPlanets
    
    fMap.Show vbModal, Me
    
    Set fMap = Nothing
End Sub

Private Sub Form_Load()
    Me.Icon = MainForm.Icon
    Call LoadFormSettings(Me, True)
    dtRunTime.DateFormat = ""
    dtRunTime.TimeFormat = "hh:nn"
    dtRunTime.TimeStamp = "00:00"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveFormSettings(Me, True)
    Call SaveGridSettings(grdRaces, Me.Name)
End Sub

Private Sub grdRaces_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
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

Private Sub LoadRaces()
    Dim objRace As Race
    Dim R As Long
    Dim c As Long
    
    With grdRaces
        .Clear
        .Rows = mobjGame.Races.Count + 2
        .Cols = 9
        .FixedRows = 1
        .FixedCols = 3
        .RowHeight(1) = 0
        .AllowUserResizing = flexResizeColumns
        .SelectionMode = flexSelectionByRow
        .FocusRect = flexFocusNone
        .ColSel = 7
        c = 0     '------------------------------
        .ColWidth(c) = 16 * Screen.TwipsPerPixelX
        c = c + 1 '------------------------------ Status
        .TextMatrix(0, c) = "S"
        .ColWidth(c) = 16 * Screen.TwipsPerPixelX
        .ColAlignment(c) = flexAlignCenterCenter
        c = c + 1 '------------------------------ Race
        .TextMatrix(0, c) = "Race"
        .ColWidth(c) = 1200
        c = c + 1 '------------------------------ E-Mail
        .TextMatrix(0, c) = "E-Mail"
        .ColWidth(c) = 1 * Screen.TwipsPerPixelX
        c = c + 1 '------------------------------ Player
        .TextMatrix(0, c) = "Player"
        .ColWidth(c) = 1200
        c = c + 1 '------------------------------ last Orders
        .TextMatrix(0, c) = "Lst Ord"
        .ColWidth(c) = 600
        .ColAlignment(c) = flexAlignCenterCenter
        c = c + 1 '------------------------------ Technology
        .TextMatrix(0, c) = "D/W/S/C"
        .ColWidth(c) = 1400
        .ColAlignment(c) = flexAlignLeftCenter
        c = c + 1 '------------------------------ Planets
        .TextMatrix(0, c) = "Planets"
        .ColWidth(c) = 600
        .ColAlignment(c) = flexAlignCenterCenter
        c = c + 1 '------------------------------ Production
        .TextMatrix(0, c) = "Prod L/T"
        .ColWidth(c) = 1000
        .ColAlignment(c) = flexAlignLeftCenter
        
        Call LoadGridSettings(grdRaces, Me.Name)
        
        R = 1
        For Each objRace In mobjGame.Races
            R = R + 1
            c = 1 '------------------------------ Status
            If objRace.flag(R_DEAD) Then
                .TextMatrix(R, c) = "X"
            ElseIf mobjGame.FinalOrdersReceived(objRace.RaceName) Then
                .TextMatrix(R, c) = "F"
            ElseIf mobjGame.OrdersReceived(objRace.RaceName) Then
                .TextMatrix(R, c) = "O"
            Else
                .TextMatrix(R, c) = ""
            End If
            c = c + 1 '------------------------------ Race Name
            .TextMatrix(R, c) = objRace.RaceName
            c = c + 1 '------------------------------ EMail Address
            .TextMatrix(R, c) = objRace.EMail
            c = c + 1 '------------------------------ Player Name
            .TextMatrix(R, c) = objRace.PlayerName
            c = c + 1 '------------------------------ Last Orders
            .TextMatrix(R, c) = objRace.LastOrders
            c = c + 1 '------------------------------ Technology
            .TextMatrix(R, c) = CStr(RoundTech(objRace.Drive)) & " / " & _
                                CStr(RoundTech(objRace.Weapons)) & " / " & _
                                CStr(RoundTech(objRace.Shields)) & " / " & _
                                CStr(RoundTech(objRace.Cargo))
            c = c + 1 '------------------------------ Planets
            .TextMatrix(R, c) = objRace.Planets.Count
            c = c + 1 '------------------------------ Production
            .TextMatrix(R, c) = CStr(Fix(0 & objRace.MassLost)) & " / " & _
                                CStr(Fix(0 & objRace.MassProduced))
        Next objRace
    End With

End Sub

Private Function RoundTech(ByVal sngVal As Single) As Single
    RoundTech = Fix(sngVal * 10) / 10
End Function
