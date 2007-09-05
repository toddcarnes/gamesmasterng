VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmTemplate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Template"
   ClientHeight    =   7020
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   7215
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   435
      Left            =   2160
      TabIndex        =   59
      Top             =   6540
      Width           =   1155
   End
   Begin VB.Frame frRunOptions 
      Caption         =   "Run Options"
      Height          =   3135
      Left            =   4200
      TabIndex        =   50
      Top             =   1560
      Width           =   2955
      Begin VB.CheckBox chkFinalOrders 
         Height          =   315
         Left            =   1380
         TabIndex        =   27
         Tag             =   "12"
         Top             =   2700
         Width           =   255
      End
      Begin VB.TextBox txtScheduleDays 
         Height          =   315
         Left            =   1380
         TabIndex        =   26
         Tag             =   "1"
         Top             =   2400
         Width           =   555
      End
      Begin GamesMaster.DateBox dtRegOpen 
         Height          =   315
         Left            =   1380
         TabIndex        =   22
         Top             =   960
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
         Left            =   1380
         TabIndex        =   21
         Tag             =   "1"
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtMaxPlayers 
         Height          =   315
         Left            =   1380
         TabIndex        =   20
         Tag             =   "1"
         Top             =   240
         Width           =   495
      End
      Begin GamesMaster.DateBox dtRegClose 
         Height          =   315
         Left            =   1380
         TabIndex        =   23
         Top             =   1320
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
         Left            =   1380
         TabIndex        =   24
         Top             =   1680
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
         Left            =   1380
         TabIndex        =   25
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
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Final Orders:"
         Height          =   255
         Index           =   22
         Left            =   120
         TabIndex        =   58
         Top             =   2760
         Width           =   1155
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Schedule Days:"
         Height          =   255
         Index           =   21
         Left            =   120
         TabIndex        =   57
         Top             =   2460
         Width           =   1155
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Start Date:"
         Height          =   255
         Index           =   20
         Left            =   120
         TabIndex        =   56
         Top             =   2100
         Width           =   1155
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Run Time:"
         Height          =   255
         Index           =   19
         Left            =   120
         TabIndex        =   55
         Top             =   1740
         Width           =   1155
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Reg Close:"
         Height          =   255
         Index           =   18
         Left            =   120
         TabIndex        =   54
         Top             =   1380
         Width           =   1155
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Reg Open:"
         Height          =   255
         Index           =   17
         Left            =   120
         TabIndex        =   53
         Top             =   1020
         Width           =   1155
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Min Players:"
         Height          =   255
         Index           =   16
         Left            =   120
         TabIndex        =   52
         Top             =   660
         Width           =   1155
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Max Players:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   51
         Top             =   300
         Width           =   1155
      End
   End
   Begin VB.Frame frRegistrations 
      Caption         =   "Registrations"
      Height          =   1755
      Left            =   0
      TabIndex        =   44
      Top             =   4740
      Width           =   7155
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdRegistrations 
         Height          =   1395
         Left            =   120
         TabIndex        =   28
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
      TabIndex        =   40
      Top             =   480
      Width           =   7095
      Begin VB.CheckBox chkFullBombing 
         Height          =   315
         Left            =   1440
         TabIndex        =   4
         Tag             =   "8"
         Top             =   600
         Width           =   255
      End
      Begin VB.CheckBox chkKeepproduction 
         Height          =   315
         Left            =   3120
         TabIndex        =   5
         Tag             =   "10"
         Top             =   600
         Width           =   255
      End
      Begin VB.CheckBox chkDontDropDead 
         Height          =   315
         Left            =   4740
         TabIndex        =   6
         Tag             =   "11"
         Top             =   600
         Width           =   255
      End
      Begin VB.CheckBox chkSphericalGalaxy 
         Height          =   315
         Left            =   6360
         TabIndex        =   7
         Tag             =   "12"
         Top             =   600
         Width           =   255
      End
      Begin VB.TextBox txtPeace 
         Height          =   315
         Left            =   4920
         TabIndex        =   3
         Tag             =   "9"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtRaceSpacing 
         Height          =   315
         Left            =   3300
         TabIndex        =   2
         Tag             =   "2"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtSize 
         Height          =   315
         Left            =   1440
         TabIndex        =   1
         Tag             =   "1"
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Full Bombing:"
         Height          =   255
         Index           =   6
         Left            =   300
         TabIndex        =   49
         Top             =   660
         Width           =   1035
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Keep production:"
         Height          =   255
         Index           =   7
         Left            =   1740
         TabIndex        =   48
         Top             =   660
         Width           =   1275
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Don't Drop Dead:"
         Height          =   255
         Index           =   8
         Left            =   3360
         TabIndex        =   47
         Top             =   660
         Width           =   1275
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Spherical Galaxy:"
         Height          =   255
         Index           =   9
         Left            =   4980
         TabIndex        =   46
         Top             =   660
         Width           =   1275
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Peace:"
         Height          =   255
         Index           =   12
         Left            =   4260
         TabIndex        =   45
         Top             =   300
         Width           =   555
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Race Spacing:"
         Height          =   255
         Index           =   2
         Left            =   2040
         TabIndex        =   42
         Top             =   300
         Width           =   1155
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Galaxy Size:"
         Height          =   255
         Index           =   1
         Left            =   420
         TabIndex        =   41
         Top             =   300
         Width           =   915
      End
   End
   Begin VB.TextBox txtName 
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Top             =   60
      Width           =   2355
   End
   Begin VB.Frame frPlayer 
      Caption         =   "Player Options"
      Height          =   3135
      Left            =   60
      TabIndex        =   30
      Top             =   1560
      Width           =   4095
      Begin VB.Frame frCoreSizes 
         Caption         =   "Planet Core Sizes"
         Height          =   675
         Left            =   120
         TabIndex        =   36
         Top             =   1380
         Width           =   3855
         Begin VB.TextBox txtCoreSizes 
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   11
            Tag             =   "3"
            Top             =   240
            Width           =   675
         End
         Begin VB.TextBox txtCoreSizes 
            Height          =   315
            Index           =   1
            Left            =   840
            TabIndex        =   12
            Tag             =   "3"
            Top             =   240
            Width           =   675
         End
         Begin VB.TextBox txtCoreSizes 
            Height          =   315
            Index           =   2
            Left            =   1560
            TabIndex        =   13
            Tag             =   "3"
            Top             =   240
            Width           =   675
         End
         Begin VB.TextBox txtCoreSizes 
            Height          =   315
            Index           =   3
            Left            =   2280
            TabIndex        =   14
            Tag             =   "3"
            Top             =   240
            Width           =   675
         End
         Begin VB.TextBox txtCoreSizes 
            Height          =   315
            Index           =   4
            Left            =   3000
            TabIndex        =   15
            Tag             =   "3"
            Top             =   240
            Width           =   675
         End
      End
      Begin VB.Frame frInitialTechLevels 
         Caption         =   "Initial Tech Levels"
         Height          =   915
         Left            =   120
         TabIndex        =   31
         Top             =   2100
         Width           =   3495
         Begin VB.TextBox txtInitialTechLevel 
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   16
            Tag             =   "7"
            Top             =   480
            Width           =   675
         End
         Begin VB.TextBox txtInitialTechLevel 
            Height          =   315
            Index           =   1
            Left            =   960
            TabIndex        =   17
            Tag             =   "7"
            Top             =   480
            Width           =   675
         End
         Begin VB.TextBox txtInitialTechLevel 
            Height          =   315
            Index           =   2
            Left            =   1800
            TabIndex        =   18
            Tag             =   "7"
            Top             =   480
            Width           =   675
         End
         Begin VB.TextBox txtInitialTechLevel 
            Height          =   315
            Index           =   3
            Left            =   2640
            TabIndex        =   19
            Tag             =   "7"
            Top             =   480
            Width           =   675
         End
         Begin VB.Label Label 
            Caption         =   "Drive"
            Height          =   255
            Index           =   11
            Left            =   120
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
            Caption         =   "Shields"
            Height          =   255
            Index           =   14
            Left            =   1800
            TabIndex        =   33
            Top             =   240
            Width           =   675
         End
         Begin VB.Label Label 
            Caption         =   "Cargo"
            Height          =   255
            Index           =   15
            Left            =   2640
            TabIndex        =   32
            Top             =   240
            Width           =   675
         End
      End
      Begin VB.TextBox txtStuffPlanets 
         Height          =   315
         Left            =   1260
         TabIndex        =   10
         Tag             =   "6"
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txtEmptyRadius 
         Height          =   315
         Left            =   1260
         TabIndex        =   9
         Tag             =   "5"
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtEmptyPlanets 
         Height          =   315
         Left            =   1260
         TabIndex        =   8
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
         TabIndex        =   39
         Top             =   1020
         Width           =   975
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Empty Radius"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   38
         Top             =   660
         Width           =   1035
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Empty Planets"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   37
         Top             =   300
         Width           =   1035
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   435
      Left            =   3660
      TabIndex        =   29
      Top             =   6540
      Width           =   1155
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      Caption         =   "Name:"
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   43
      Top             =   120
      Width           =   1395
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

Public Property Set Template(ByVal objtemplate As Template)
    Set mobjTemplate = objtemplate
    Call LoadTemplate
End Property

Public Property Let ReadOnly(ByVal blnReadOnly As Boolean)
    mblnReadOnly = blnReadOnly
    txtName.Locked = blnReadOnly
    frGalaxy.Enabled = Not blnReadOnly
    frPlayer.Enabled = Not blnReadOnly
    frRunOptions.Enabled = Not blnReadOnly
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

Private Sub chkDontDropDead_Click()
    Template.DontDropDead = (chkDontDropDead.Value = vbChecked)
End Sub

Private Sub chkFinalOrders_Click()
    Template.FinalOrders = (chkFinalOrders.Value = vbChecked)
End Sub

Private Sub chkFullBombing_Click()
    Template.FullBombing = (chkFullBombing.Value = vbChecked)
End Sub

Private Sub chkKeepproduction_Click()
    Template.KeepProduction = (chkKeepproduction.Value = vbChecked)
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
    End If
    Unload Me
End Sub

Private Sub LoadTemplate()
    Dim i As Long
    Dim C As Long
    Dim objRegistration As Registration
    
    If mobjTemplate Is Nothing Then
        txtName = ""
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
            txtMaxPlayers = .MaxPlayers
            txtMinPlayers = .MinPlayers
            dtRegOpen.TimeStamp = .RegistrationOpen
            dtRegClose.TimeStamp = .RegistrationClose
            dtRunTime.TimeStamp = .RunTime
            dtStartDate.TimeStamp = .StartDate
            txtScheduleDays = .ScheduleDays
            chkFinalOrders = IIf(.FinalOrders, vbChecked, vbUnchecked)
            
            With grdRegistrations
                .Clear
                .FixedRows = 1
                .FixedCols = 1
                .AllowUserResizing = flexResizeColumns
                .SelectionMode = flexSelectionByRow
                .FocusRect = flexFocusNone
                .Rows = mobjTemplate.Registrations.Count + 1
                .Cols = 6
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
                i = 0
                For Each objRegistration In mobjTemplate.Registrations
                    i = i + 1
                    .TextMatrix(i, 1) = objRegistration.EMail
                    For C = 1 To objRegistration.HomeWorlds.Count
                        If C > 4 Then Exit For
                        If objRegistration.HomeWorlds(C).x = 0 Then
                            .TextMatrix(i, C + 1) = objRegistration.HomeWorlds(C).Size
                        Else
                            .TextMatrix(i, C + 1) = objRegistration.HomeWorlds(C).Size _
                                                & "/" & objRegistration.HomeWorlds(C).x _
                                                & "/" & objRegistration.HomeWorlds(C).y
                        End If
                    Next C
                Next objRegistration
            End With
        End With
    End If
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
    dtRegOpen.DateFormat = "Short Date"
    dtRegOpen.TimeFormat = ""
    dtRegClose.DateFormat = "Short Date"
    dtRegClose.TimeFormat = ""
    dtRunTime.DateFormat = ""
    dtRunTime.TimeFormat = "hh:nn"
    dtRunTime.TimeStamp = "00:00"
    dtStartDate.DateFormat = "Short Date"
    dtStartDate.TimeFormat = ""
End Sub

Private Sub txtCoreSizes_Change(Index As Integer)
    Template.core_sizes(Index) = Val(txtCoreSizes(Index))
End Sub

Private Sub txtEmptyPlanets_Change()
    Template.empty_planets = Val(txtEmptyPlanets.Text)
End Sub

Private Sub txtEmptyRadius_Change()
    Template.empty_radius = Val(txtEmptyRadius.Text)
End Sub

Private Sub txtMaxPlayers_Change()
    Template.MaxPlayers = Val(txtMaxPlayers.Text)
End Sub

Private Sub txtMinPlayers_Change()
    Template.MinPlayers = Val(txtMinPlayers.Text)
End Sub

Private Sub txtName_Change()
    Template.TemplateName = txtName.Text
End Sub

Private Sub txtPeace_Change()
    Template.Peace = Val(txtPeace.Text)
End Sub

Private Sub txtRaceSpacing_Change()
    Template.race_spacing = Val(txtRaceSpacing.Text)
End Sub

Private Sub txtScheduleDays_Change()
    Template.ScheduleDays = Val(txtScheduleDays.Text)
End Sub

Private Sub txtSize_Change()
    Template.Size = Val(txtSize.Text)
End Sub

Private Sub txtStuffPlanets_Change()
    Template.stuff_planets = Val(txtStuffPlanets.Text)
End Sub