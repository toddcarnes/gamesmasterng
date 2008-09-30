VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmTemplate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Template"
   ClientHeight    =   8220
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdViewMap 
      Caption         =   "View Map"
      Height          =   435
      Left            =   1620
      TabIndex        =   72
      Top             =   7680
      Width           =   915
   End
   Begin VB.CommandButton cmdDesign 
      Caption         =   "Apply Design"
      Height          =   435
      Left            =   180
      TabIndex        =   71
      Top             =   7680
      Width           =   1275
   End
   Begin VB.Frame frDesign 
      Caption         =   "Design"
      Height          =   1395
      Left            =   60
      TabIndex        =   67
      Top             =   4380
      Width           =   4095
      Begin VB.TextBox txtOrbitDistance 
         Height          =   315
         Left            =   3420
         TabIndex        =   23
         Tag             =   "2"
         Top             =   540
         Width           =   495
      End
      Begin VB.CheckBox chkOrbitPlanets 
         Alignment       =   1  'Right Justify
         Caption         =   "Orbit Planets:"
         Height          =   315
         Left            =   120
         TabIndex        =   22
         Tag             =   "12"
         Top             =   540
         Width           =   1275
      End
      Begin VB.ComboBox cboSeed 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   900
         Width           =   2715
      End
      Begin VB.ComboBox cboDesign 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   180
         Width           =   2715
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Seed Type:"
         Height          =   255
         Index           =   31
         Left            =   60
         TabIndex        =   70
         Top             =   960
         Width           =   1035
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Orbit Distance:"
         Height          =   255
         Index           =   30
         Left            =   2100
         TabIndex        =   69
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Design Type:"
         Height          =   255
         Index           =   28
         Left            =   60
         TabIndex        =   68
         Top             =   240
         Width           =   1035
      End
   End
   Begin VB.CheckBox chkFinished 
      Alignment       =   1  'Right Justify
      Caption         =   "Game Finished:"
      Height          =   315
      Left            =   5340
      TabIndex        =   37
      Tag             =   "12"
      Top             =   60
      Width           =   1395
   End
   Begin VB.CheckBox chkScheduleActive 
      Alignment       =   1  'Right Justify
      Caption         =   "Schedule Active:"
      Height          =   315
      Left            =   3360
      TabIndex        =   1
      Tag             =   "12"
      Top             =   60
      Width           =   1515
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   435
      Left            =   4380
      TabIndex        =   63
      Top             =   7680
      Width           =   1155
   End
   Begin VB.Frame frRunOptions 
      Caption         =   "Run Options"
      Height          =   4215
      Left            =   4200
      TabIndex        =   55
      Top             =   1560
      Width           =   2955
      Begin VB.TextBox txtTotalPlanetSize 
         Height          =   315
         Left            =   1560
         TabIndex        =   29
         Tag             =   "1"
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox txtMaxPlanetSize 
         Height          =   315
         Left            =   1560
         TabIndex        =   28
         Tag             =   "1"
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox txtMaxPlanets 
         Height          =   315
         Left            =   1560
         TabIndex        =   27
         Tag             =   "1"
         Top             =   960
         Width           =   495
      End
      Begin VB.CheckBox chkFinalOrders 
         Alignment       =   1  'Right Justify
         Caption         =   "Final Orders:"
         Height          =   315
         Left            =   540
         TabIndex        =   35
         Tag             =   "12"
         Top             =   3840
         Width           =   1215
      End
      Begin VB.TextBox txtScheduleDays 
         Height          =   315
         Left            =   1560
         TabIndex        =   34
         Tag             =   "1"
         Top             =   3480
         Width           =   555
      End
      Begin GamesMaster.DateBox dtRegOpen 
         Height          =   315
         Left            =   1560
         TabIndex        =   30
         Top             =   2040
         Width           =   1035
         _ExtentX        =   1826
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
         DateFormat      =   "dd/mm/yyyy"
         TimeFormat      =   ""
      End
      Begin VB.TextBox txtMinPlayers 
         Height          =   315
         Left            =   1560
         TabIndex        =   26
         Tag             =   "1"
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtMaxPlayers 
         Height          =   315
         Left            =   1560
         TabIndex        =   25
         Tag             =   "1"
         Top             =   240
         Width           =   495
      End
      Begin GamesMaster.DateBox dtRegClose 
         Height          =   315
         Left            =   1560
         TabIndex        =   31
         Top             =   2400
         Width           =   1035
         _ExtentX        =   1826
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
         DateFormat      =   "dd/mm/yyyy"
         TimeFormat      =   ""
      End
      Begin GamesMaster.DateBox dtRunTime 
         Height          =   315
         Left            =   1560
         TabIndex        =   32
         Top             =   2760
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
      Begin GamesMaster.DateBox dtStartDate 
         Height          =   315
         Left            =   1560
         TabIndex        =   33
         Top             =   3120
         Width           =   1035
         _ExtentX        =   1826
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
         DateFormat      =   "dd/mm/yyyy"
         TimeFormat      =   ""
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Planet Size:"
         Height          =   255
         Index           =   25
         Left            =   60
         TabIndex        =   66
         Top             =   1740
         Width           =   1395
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Max Planet Size:"
         Height          =   255
         Index           =   24
         Left            =   240
         TabIndex        =   65
         Top             =   1380
         Width           =   1215
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Max Planets:"
         Height          =   255
         Index           =   23
         Left            =   300
         TabIndex        =   64
         Top             =   1020
         Width           =   1155
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Schedule Days:"
         Height          =   255
         Index           =   21
         Left            =   300
         TabIndex        =   62
         Top             =   3540
         Width           =   1155
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Start Date:"
         Height          =   255
         Index           =   20
         Left            =   300
         TabIndex        =   61
         Top             =   3180
         Width           =   1155
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Run Time:"
         Height          =   255
         Index           =   19
         Left            =   300
         TabIndex        =   60
         Top             =   2820
         Width           =   1155
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Reg Close:"
         Height          =   255
         Index           =   18
         Left            =   300
         TabIndex        =   59
         Top             =   2460
         Width           =   1155
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Reg Open:"
         Height          =   255
         Index           =   17
         Left            =   300
         TabIndex        =   58
         Top             =   2100
         Width           =   1155
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Min Players:"
         Height          =   255
         Index           =   16
         Left            =   300
         TabIndex        =   57
         Top             =   660
         Width           =   1155
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Max Players:"
         Height          =   255
         Index           =   3
         Left            =   300
         TabIndex        =   56
         Top             =   300
         Width           =   1155
      End
   End
   Begin VB.Frame frRegistrations 
      Caption         =   "Registrations"
      Height          =   1755
      Left            =   0
      TabIndex        =   53
      Top             =   5820
      Width           =   7155
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdRegistrations 
         Height          =   1395
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   6915
         _ExtentX        =   12197
         _ExtentY        =   2461
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame frGalaxy 
      Caption         =   "Galaxy Options"
      Height          =   1035
      Left            =   60
      TabIndex        =   49
      Top             =   480
      Width           =   7095
      Begin VB.CheckBox chkFullBombing 
         Alignment       =   1  'Right Justify
         Caption         =   "Full Bombing:"
         Height          =   315
         Left            =   180
         TabIndex        =   5
         Tag             =   "8"
         Top             =   600
         Width           =   1275
      End
      Begin VB.CheckBox chkKeepproduction 
         Alignment       =   1  'Right Justify
         Caption         =   "Keep production:"
         Height          =   315
         Left            =   1740
         TabIndex        =   6
         Tag             =   "10"
         Top             =   600
         Width           =   1515
      End
      Begin VB.CheckBox chkDontDropDead 
         Alignment       =   1  'Right Justify
         Caption         =   "Don't Drop Dead:"
         Height          =   315
         Left            =   3540
         TabIndex        =   7
         Tag             =   "11"
         Top             =   600
         Width           =   1575
      End
      Begin VB.CheckBox chkSphericalGalaxy 
         Alignment       =   1  'Right Justify
         Caption         =   "Spherical Galaxy:"
         Height          =   315
         Left            =   5400
         TabIndex        =   8
         Tag             =   "12"
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtPeace 
         Height          =   315
         Left            =   4920
         TabIndex        =   4
         Tag             =   "9"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtRaceSpacing 
         Height          =   315
         Left            =   3060
         TabIndex        =   3
         Tag             =   "2"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtSize 
         Height          =   315
         Left            =   1260
         TabIndex        =   2
         Tag             =   "1"
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Peace:"
         Height          =   255
         Index           =   12
         Left            =   4260
         TabIndex        =   54
         Top             =   300
         Width           =   555
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Race Spacing:"
         Height          =   255
         Index           =   2
         Left            =   1800
         TabIndex        =   51
         Top             =   300
         Width           =   1155
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Galaxy Size:"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   50
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
   Begin VB.Frame frPlayer 
      Caption         =   "Player Options"
      Height          =   2775
      Left            =   60
      TabIndex        =   39
      Top             =   1560
      Width           =   4095
      Begin VB.Frame frCoreSizes 
         Caption         =   "Planet Core Sizes"
         Height          =   675
         Left            =   120
         TabIndex        =   45
         Top             =   1020
         Width           =   3855
         Begin VB.TextBox txtCoreSizes 
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   12
            Tag             =   "3"
            Top             =   240
            Width           =   675
         End
         Begin VB.TextBox txtCoreSizes 
            Height          =   315
            Index           =   1
            Left            =   840
            TabIndex        =   13
            Tag             =   "3"
            Top             =   240
            Width           =   675
         End
         Begin VB.TextBox txtCoreSizes 
            Height          =   315
            Index           =   2
            Left            =   1560
            TabIndex        =   14
            Tag             =   "3"
            Top             =   240
            Width           =   675
         End
         Begin VB.TextBox txtCoreSizes 
            Height          =   315
            Index           =   3
            Left            =   2280
            TabIndex        =   15
            Tag             =   "3"
            Top             =   240
            Width           =   675
         End
         Begin VB.TextBox txtCoreSizes 
            Height          =   315
            Index           =   4
            Left            =   3000
            TabIndex        =   16
            Tag             =   "3"
            Top             =   240
            Width           =   675
         End
      End
      Begin VB.Frame frInitialTechLevels 
         Caption         =   "Initial Tech Levels"
         Height          =   915
         Left            =   120
         TabIndex        =   40
         Top             =   1740
         Width           =   3495
         Begin VB.TextBox txtInitialTechLevel 
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   17
            Tag             =   "7"
            Top             =   480
            Width           =   675
         End
         Begin VB.TextBox txtInitialTechLevel 
            Height          =   315
            Index           =   1
            Left            =   960
            TabIndex        =   18
            Tag             =   "7"
            Top             =   480
            Width           =   675
         End
         Begin VB.TextBox txtInitialTechLevel 
            Height          =   315
            Index           =   2
            Left            =   1800
            TabIndex        =   19
            Tag             =   "7"
            Top             =   480
            Width           =   675
         End
         Begin VB.TextBox txtInitialTechLevel 
            Height          =   315
            Index           =   3
            Left            =   2640
            TabIndex        =   20
            Tag             =   "7"
            Top             =   480
            Width           =   675
         End
         Begin VB.Label Label 
            Caption         =   "Drive"
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   44
            Top             =   240
            Width           =   675
         End
         Begin VB.Label Label 
            Caption         =   "Weapons"
            Height          =   255
            Index           =   13
            Left            =   960
            TabIndex        =   43
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label 
            Caption         =   "Shields"
            Height          =   255
            Index           =   14
            Left            =   1800
            TabIndex        =   42
            Top             =   240
            Width           =   675
         End
         Begin VB.Label Label 
            Caption         =   "Cargo"
            Height          =   255
            Index           =   15
            Left            =   2640
            TabIndex        =   41
            Top             =   240
            Width           =   675
         End
      End
      Begin VB.TextBox txtStuffPlanets 
         Height          =   315
         Left            =   1260
         TabIndex        =   11
         Tag             =   "6"
         Top             =   660
         Width           =   495
      End
      Begin VB.TextBox txtEmptyRadius 
         Height          =   315
         Left            =   3300
         TabIndex        =   10
         Tag             =   "5"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtEmptyPlanets 
         Height          =   315
         Left            =   1260
         TabIndex        =   9
         Tag             =   "4"
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Stuff Planets:"
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   48
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Empty Radius"
         Height          =   255
         Index           =   5
         Left            =   2160
         TabIndex        =   47
         Top             =   300
         Width           =   1035
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Empty Planets"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   46
         Top             =   300
         Width           =   1035
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   435
      Left            =   5880
      TabIndex        =   38
      Top             =   7680
      Width           =   1155
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      Caption         =   "Name:"
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   52
      Top             =   120
      Width           =   555
   End
   Begin VB.Menu mnuEdit1 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditDescription 
         Caption         =   "Description"
      End
      Begin VB.Menu mnuEditMessage 
         Caption         =   "Message"
      End
   End
   Begin VB.Menu mnuAction 
      Caption         =   "&Registrations"
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
Attribute VB_Name = "frmTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mobjTemplate As Template
Private mblnReadOnly As Boolean

Public Property Get Template() As Template
    Set Template = mobjTemplate
End Property

Public Property Set Template(ByVal objTemplate As Template)
    Set mobjTemplate = objTemplate
    Call LoadTemplate
End Property

Public Property Let ReadOnly(ByVal blnReadOnly As Boolean)
    mblnReadOnly = blnReadOnly
    txtName.Locked = blnReadOnly
    frGalaxy.Enabled = Not blnReadOnly
    frPlayer.Enabled = Not blnReadOnly
    frRunOptions.Enabled = Not blnReadOnly
    frDesign.Enabled = Not blnReadOnly
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

Private Sub cboDesign_Click()
    Template.DesignType = cboDesign.ListIndex
End Sub

Private Sub cboSeed_Click()
    Template.SeedType = cboSeed.ListIndex
End Sub

Private Sub chkDontDropDead_Click()
    Template.DontDropDead = (chkDontDropDead.Value = vbChecked)
End Sub

Private Sub chkFinalOrders_Click()
    Template.FinalOrders = (chkFinalOrders.Value = vbChecked)
End Sub

Private Sub chkFinished_Click()
    Template.Finished = (chkFinished.Value = vbChecked)
End Sub

Private Sub chkFullBombing_Click()
    Template.FullBombing = (chkFullBombing.Value = vbChecked)
End Sub

Private Sub chkKeepproduction_Click()
    Template.KeepProduction = (chkKeepproduction.Value = vbChecked)
End Sub

Private Sub chkOrbitPlanets_Click()
    Template.OrbitPlanets = (chkOrbitPlanets.Value = vbChecked)
End Sub

Private Sub chkScheduleActive_Click()
    Template.ScheduleActive = (chkScheduleActive.Value = vbChecked)
End Sub

Private Sub chkSphericalGalaxy_Click()
    Template.sphericalgalaxy = (chkSphericalGalaxy.Value = vbChecked)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClose_Click()
    If Not ReadOnly Then
        Call Template.Save
        Call MainForm.RefreshGamesForm
    End If
    Unload Me
End Sub

Private Sub LoadTemplate()
    Dim i As Long
    
    If mobjTemplate Is Nothing Then
        txtName = ""
        chkScheduleActive = vbUnchecked
        chkFinished = vbUnchecked
        txtSize = ""
        txtRaceSpacing = ""
        txtEmptyPlanets = ""
        txtEmptyRadius = ""
        txtStuffPlanets = ""
        chkFullBombing = vbUnchecked
        txtPeace = ""
        chkKeepproduction = vbUnchecked
        chkDontDropDead = vbUnchecked
        chkSphericalGalaxy = vbUnchecked
        On Error Resume Next
        For i = txtCoreSizes.LBound To txtCoreSizes.UBound
            txtCoreSizes(i) = ""
            txtCoreSizes(i).Visible = False
        Next i
        On Error GoTo 0
        txtInitialTechLevel(Tech.Drive) = ""
        txtInitialTechLevel(Tech.Weapons) = ""
        txtInitialTechLevel(Tech.Shields) = ""
        txtInitialTechLevel(Tech.Cargo) = ""
        cboDesign.ListIndex = 0
        chkOrbitPlanets = vbUnchecked
        txtOrbitDistance = 0
        cboSeed.ListIndex = 0
        With grdRegistrations
            .Clear
            .Rows = 0
            .Cols = 0
            .AllowUserResizing = flexResizeColumns
            .SelectionMode = flexSelectionByRow
            .FocusRect = flexFocusNone
        End With
    Else
        With mobjTemplate
            txtName = .TemplateName
            chkScheduleActive = IIf(.ScheduleActive, vbChecked, vbUnchecked)
            chkFinished = IIf(.Finished, vbChecked, vbUnchecked)
            txtSize = .Size
            txtRaceSpacing = .race_spacing
            txtEmptyPlanets = .empty_planets
            txtEmptyRadius = .empty_radius
            txtStuffPlanets = .stuff_planets
            chkFullBombing = IIf(.FullBombing, vbChecked, vbUnchecked)
            txtPeace = .Peace
            chkKeepproduction = IIf(.KeepProduction, vbChecked, vbUnchecked)
            chkDontDropDead = IIf(.DontDropDead, vbChecked, vbUnchecked)
            chkSphericalGalaxy = IIf(.sphericalgalaxy, vbChecked, vbUnchecked)
            
            For i = 0 To UBound(.core_sizes)
                If i > txtCoreSizes.UBound Then Exit For
                txtCoreSizes(i) = .core_sizes(i)
            Next i
            txtInitialTechLevel(Tech.Drive) = .InitialTechLevels(Tech.Drive)
            txtInitialTechLevel(Tech.Weapons) = .InitialTechLevels(Tech.Weapons)
            txtInitialTechLevel(Tech.Shields) = .InitialTechLevels(Tech.Shields)
            txtInitialTechLevel(Tech.Cargo) = .InitialTechLevels(Tech.Cargo)
            cboDesign.ListIndex = .DesignType
            chkOrbitPlanets = IIf(.OrbitPlanets, vbChecked, vbUnchecked)
            txtOrbitDistance = .OrbitDistance
            cboSeed.ListIndex = .SeedType
            
            txtMaxPlayers = .MaxPlayers
            txtMinPlayers = .MinPlayers
            txtMaxPlanets = .MaxPlanets
            txtMaxPlanetSize = .MaxPlanetSize
            txtTotalPlanetSize = .TotalPlanetSize
            dtRegOpen.TimeStamp = .RegistrationOpen
            dtRegClose.TimeStamp = .RegistrationClose
            dtRunTime.TimeStamp = .RunTime
            dtStartDate.TimeStamp = .StartDate
            txtScheduleDays = .ScheduleDays
            chkFinalOrders = IIf(.FinalOrders, vbChecked, vbUnchecked)
        End With
        Call LoadRegistrations
    End If
End Sub

Private Sub cmdDesign_Click()
    Call ApplyDesign(Template)
End Sub

Private Sub cmdViewMap_Click()
    Dim fMap As frmMap
    Dim objReg As Registration
    Dim objHomeworld As HomeWorld
    Dim objPlanet As Planet
    Dim colPlanets As Collection
    Dim R As Long
    Dim H As Long
    Dim p As Long
    
    Set colPlanets = New Collection
    Set fMap = New frmMap
    Load fMap
    fMap.GalaxySize = Template.Size
    
    R = 0
    For Each objReg In Template.Registrations
        R = R + 1
        H = 0
        For Each objHomeworld In objReg.HomeWorlds
            H = H + 1
            objHomeworld.Owner = "R" & CStr(R)
            objHomeworld.Planet = objHomeworld.Owner & "_" & CStr(H)
            objHomeworld.Resources = 10
            colPlanets.Add objHomeworld
        Next objHomeworld
    Next objReg
    
    p = 0
    For Each objPlanet In Template.Planets
        p = p + 1
        objPlanet.Planet = "P" & CStr(p)
        colPlanets.Add objPlanet
    Next objPlanet
    
    Set fMap.Planets = colPlanets
    
    fMap.Show vbModal, Me
    
    Set fMap = Nothing
End Sub

Private Sub dtRegClose_Change()
    Template.RegistrationClose = dtRegClose.TimeStamp
End Sub

Private Sub dtRegOpen_Change()
    Template.RegistrationOpen = dtRegOpen.TimeStamp
End Sub

Private Sub dtRunTime_Change()
    Template.RunTime = dtRunTime.TimeStamp
End Sub

Private Sub dtStartDate_Change()
    Template.StartDate = dtStartDate.TimeStamp
End Sub

Private Sub Form_Load()
    Dim i As Long
    Dim vData As Variant
    
    Call LoadFormSettings(Me, True)
    Me.Icon = MainForm.Icon
    dtRegOpen.DateFormat = "Short Date"
    dtRegOpen.TimeFormat = ""
    dtRegClose.DateFormat = "Short Date"
    dtRegClose.TimeFormat = ""
    dtRunTime.DateFormat = ""
    dtRunTime.TimeFormat = "hh:nn"
    dtRunTime.TimeStamp = "00:00"
    dtStartDate.DateFormat = "Short Date"
    dtStartDate.TimeFormat = ""
    
    With cboDesign
        .Clear
        vData = Options.DesignDescriptions
        For i = LBound(vData) To UBound(vData)
            .AddItem vData(i)
        Next i
    End With
    
    With cboSeed
        .Clear
        vData = Options.SeedDescriptions
        For i = LBound(vData) To UBound(vData)
            .AddItem vData(i)
        Next i
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveFormSettings(Me, True)
    Call SaveGridSettings(grdRegistrations, Me.Name)
End Sub

Private Sub grdRegistrations_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton And Not ReadOnly Then
        PopupMenu mnuAction
    End If
End Sub

Private Sub mnuAction_Click()
    Dim blnEnable As Boolean
    With grdRegistrations
        blnEnable = (.Row > 1)
        mnuEdit.Enabled = blnEnable
        mnuDelete.Enabled = blnEnable
    End With
End Sub

Private Sub mnuAdd_Click()
    Dim objRegistration As Registration
    Dim fRegistration As frmRegistration
    
    Set objRegistration = New Registration
    
    Set fRegistration = New frmRegistration
    Set fRegistration.Registration = objRegistration
    fRegistration.Show vbModal
    If objRegistration.EMail <> "" Then
        mobjTemplate.Registrations.Add objRegistration
    End If
    Set fRegistration = Nothing
    Set objRegistration = Nothing
    Call LoadRegistrations
    
End Sub

Private Sub mnuDelete_Click()
    Dim i As Long
    
    i = grdRegistrations.Row
    If i <= 1 Then Exit Sub
    
    mobjTemplate.Registrations.Remove i - 1
    
    Call LoadRegistrations
End Sub

Private Sub mnuEdit_Click()
    Dim objRegistration As Registration
    Dim fRegistration As frmRegistration
    Dim i As Long
    
    i = grdRegistrations.Row
    If i <= 1 Then Exit Sub
    
    Set objRegistration = mobjTemplate.Registrations(i - 1)
    
    Set fRegistration = New frmRegistration
    Set fRegistration.Registration = objRegistration
    fRegistration.Show vbModal
    Set fRegistration = Nothing
    Set objRegistration = Nothing
    Call LoadRegistrations
End Sub

Private Sub mnuEditDescription_Click()
    Dim fNote As frmNote
    
    Set fNote = New frmNote
    Load fNote
    fNote.Caption = Template.TemplateName & " Description"
    fNote.Text = Template.Description
    fNote.Show vbModal, Me
    If Not fNote.Cancelled Then
        Template.Description = fNote.Text
    End If
    Unload fNote
    Set fNote = Nothing
    
End Sub

Private Sub mnuEditMessage_Click()
    Dim fNote As frmNote
    
    Set fNote = New frmNote
    Load fNote
    fNote.Caption = Template.TemplateName & " Message"
    fNote.Text = Template.Message
    fNote.Show vbModal, Me
    If Not fNote.Cancelled Then
        Template.Message = fNote.Text
    End If
    Unload fNote
    Set fNote = Nothing
    
End Sub

Private Sub txtCoreSizes_Change(Index As Integer)
    Template.core_sizes(Index) = Val(txtCoreSizes(Index))
End Sub

Private Sub txtCoreSizes_GotFocus(Index As Integer)
    With txtCoreSizes(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtEmptyPlanets_Change()
    Template.empty_planets = Val(txtEmptyPlanets.Text)
End Sub

Private Sub txtEmptyPlanets_GotFocus()
    With txtEmptyPlanets
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtEmptyRadius_Change()
    Template.empty_radius = Val(txtEmptyRadius.Text)
End Sub

Private Sub txtEmptyRadius_GotFocus()
    With txtEmptyRadius
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtInitialTechLevel_GotFocus(Index As Integer)
    With txtInitialTechLevel(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtMaxPlanets_Change()
    Template.MaxPlanets = Val(txtMaxPlanets.Text)
End Sub

Private Sub txtMaxPlanets_GotFocus()
    With txtMaxPlanets
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtMaxPlanetSize_Change()
    Template.MaxPlanetSize = Val(txtMaxPlanetSize.Text)
End Sub

Private Sub txtMaxPlanetSize_GotFocus()
    With txtMaxPlanetSize
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtMaxPlayers_Change()
    Template.MaxPlayers = Val(txtMaxPlayers.Text)
End Sub

Private Sub txtMaxPlayers_GotFocus()
    With txtMaxPlayers
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtMinPlayers_Change()
    Template.MinPlayers = Val(txtMinPlayers.Text)
End Sub

Private Sub txtMinPlayers_GotFocus()
    With txtMinPlayers
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtName_Change()
    Template.TemplateName = txtName.Text
End Sub

Private Sub txtOrbitDistance_Change()
    Template.OrbitDistance = Val(txtOrbitDistance)
End Sub

Private Sub txtOrbitDistance_GotFocus()
    With txtOrbitDistance
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtPeace_Change()
    Template.Peace = Val(txtPeace.Text)
End Sub

Private Sub txtPeace_GotFocus()
    With txtPeace
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtRaceSpacing_Change()
    Template.race_spacing = Val(txtRaceSpacing.Text)
End Sub

Private Sub txtRaceSpacing_GotFocus()
    With txtRaceSpacing
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtScheduleDays_Change()
    Template.ScheduleDays = Val(txtScheduleDays.Text)
End Sub

Private Sub txtScheduleDays_GotFocus()
    With txtScheduleDays
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtSize_Change()
    Template.Size = Val(txtSize.Text)
End Sub

Private Sub txtSize_GotFocus()
    With txtSize
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtStuffPlanets_Change()
    Template.stuff_planets = Val(txtStuffPlanets.Text)
End Sub

Private Sub LoadRegistrations()
    Dim objRegistration As Registration
    Dim i As Long
    Dim c As Long
    
    With grdRegistrations
        .Clear
        .Rows = mobjTemplate.Registrations.Count + 2
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
        
        Call SaveGridSettings(grdRegistrations, Me.Name)
        
        i = 1
        For Each objRegistration In mobjTemplate.Registrations
            i = i + 1
            .TextMatrix(i, 1) = objRegistration.EMail
            For c = 1 To objRegistration.HomeWorlds.Count
                If c > 4 Then Exit For
                .TextMatrix(i, c + 1) = objRegistration.HomeWorlds(c).Size
            Next c
        Next objRegistration
    End With

End Sub

Private Sub txtStuffPlanets_GotFocus()
    With txtStuffPlanets
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtTotalPlanetSize_Change()
    Template.TotalPlanetSize = Val(txtTotalPlanetSize.Text)
End Sub

Private Sub txtTotalPlanetSize_GotFocus()
    With txtTotalPlanetSize
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
