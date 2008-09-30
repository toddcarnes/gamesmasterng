VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   6360
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   7800
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   4920
      TabIndex        =   32
      Top             =   5820
      Width           =   1155
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   31
      Top             =   5820
      Width           =   1155
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   375
      Left            =   1560
      TabIndex        =   30
      Top             =   5820
      Width           =   1155
   End
   Begin TabDlg.SSTab tabOptions 
      Height          =   5535
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   9763
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmOptions.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label(20)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtServerName"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtGalaxyNGHome"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtGamesMasterEMail"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "chkLogErrors"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdViewErrors"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "E-Mail"
      TabPicture(1)   =   "frmOptions.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtOutBox"
      Tab(1).Control(1)=   "txtInbox"
      Tab(1).Control(2)=   "frSMTPServer"
      Tab(1).Control(3)=   "frPopServer"
      Tab(1).Control(4)=   "Label(11)"
      Tab(1).Control(5)=   "Label(10)"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Startup"
      TabPicture(2)   =   "frmOptions.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "chkAutoRunGames"
      Tab(2).Control(1)=   "chkAutoCheckMail"
      Tab(2).Control(2)=   "chkShowGetMail"
      Tab(2).Control(3)=   "chkShowSendMail"
      Tab(2).Control(4)=   "chkShowGames"
      Tab(2).Control(5)=   "chkMinimizeAtStartup"
      Tab(2).Control(6)=   "chkStartWithWindows"
      Tab(2).Control(7)=   "Label(19)"
      Tab(2).Control(8)=   "Label(18)"
      Tab(2).Control(9)=   "Label(17)"
      Tab(2).Control(10)=   "Label(16)"
      Tab(2).Control(11)=   "Label(15)"
      Tab(2).Control(12)=   "Label(14)"
      Tab(2).Control(13)=   "Label(13)"
      Tab(2).ControlCount=   14
      Begin VB.CommandButton cmdViewErrors 
         Caption         =   "View Errors"
         Height          =   375
         Left            =   3060
         TabIndex        =   48
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CheckBox chkLogErrors 
         Height          =   315
         Left            =   2640
         TabIndex        =   7
         Top             =   2040
         Width           =   315
      End
      Begin VB.CheckBox chkAutoRunGames 
         Height          =   255
         Left            =   -72600
         TabIndex        =   46
         Top             =   2820
         Width           =   1455
      End
      Begin VB.CheckBox chkAutoCheckMail 
         Height          =   255
         Left            =   -72600
         TabIndex        =   45
         Top             =   2460
         Width           =   1455
      End
      Begin VB.CheckBox chkShowGetMail 
         Height          =   255
         Left            =   -72600
         TabIndex        =   44
         Top             =   2100
         Width           =   1455
      End
      Begin VB.CheckBox chkShowSendMail 
         Height          =   255
         Left            =   -72600
         TabIndex        =   43
         Top             =   1740
         Width           =   1455
      End
      Begin VB.CheckBox chkShowGames 
         Height          =   255
         Left            =   -72600
         TabIndex        =   38
         Top             =   1380
         Width           =   1455
      End
      Begin VB.CheckBox chkMinimizeAtStartup 
         Height          =   255
         Left            =   -72600
         TabIndex        =   36
         Top             =   1020
         Width           =   1455
      End
      Begin VB.CheckBox chkStartWithWindows 
         Height          =   255
         Left            =   -72600
         TabIndex        =   35
         Top             =   660
         Width           =   1455
      End
      Begin VB.TextBox txtOutBox 
         Height          =   315
         Left            =   -73440
         TabIndex        =   27
         Top             =   4980
         Width           =   5355
      End
      Begin VB.TextBox txtInbox 
         Height          =   315
         Left            =   -73440
         TabIndex        =   26
         Top             =   4560
         Width           =   5355
      End
      Begin VB.Frame frSMTPServer 
         Caption         =   "SMTP Server"
         Height          =   1695
         Left            =   -74820
         TabIndex        =   13
         Top             =   2700
         Width           =   7275
         Begin VB.TextBox txtSMTPFromAddress 
            Height          =   315
            Left            =   1440
            TabIndex        =   23
            Top             =   1140
            Width           =   4335
         End
         Begin VB.TextBox txtSMTPServerPort 
            Height          =   315
            Left            =   1440
            TabIndex        =   22
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox txtSMTPServer 
            Height          =   315
            Left            =   1440
            TabIndex        =   21
            Top             =   300
            Width           =   3675
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "From Address:"
            Height          =   255
            Index           =   9
            Left            =   180
            TabIndex        =   16
            Top             =   1140
            Width           =   1155
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Server Port:"
            Height          =   255
            Index           =   8
            Left            =   180
            TabIndex        =   15
            Top             =   720
            Width           =   1155
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Server Name:"
            Height          =   255
            Index           =   7
            Left            =   180
            TabIndex        =   14
            Top             =   300
            Width           =   1155
         End
      End
      Begin VB.Frame frPopServer 
         Caption         =   "POP Server"
         Height          =   2055
         Left            =   -74820
         TabIndex        =   8
         Top             =   540
         Width           =   7275
         Begin VB.TextBox txtCheckMailInterval 
            Height          =   315
            Left            =   5640
            TabIndex        =   29
            Top             =   1500
            Width           =   615
         End
         Begin VB.TextBox txtPOPPassword 
            Height          =   315
            Left            =   1440
            TabIndex        =   20
            Top             =   1500
            Width           =   1695
         End
         Begin VB.TextBox txtPOPUserID 
            Height          =   315
            Left            =   1440
            TabIndex        =   19
            Top             =   1080
            Width           =   1695
         End
         Begin VB.TextBox txtPOPServerPort 
            Height          =   315
            Left            =   1440
            TabIndex        =   18
            Top             =   660
            Width           =   615
         End
         Begin VB.TextBox txtPOPServer 
            Height          =   315
            Left            =   1440
            TabIndex        =   17
            Top             =   240
            Width           =   3675
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Check Mail Interval:"
            Height          =   255
            Index           =   12
            Left            =   3840
            TabIndex        =   28
            Top             =   1560
            Width           =   1635
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Password:"
            Height          =   255
            Index           =   6
            Left            =   180
            TabIndex        =   12
            Top             =   1560
            Width           =   1155
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "User Name:"
            Height          =   255
            Index           =   5
            Left            =   180
            TabIndex        =   11
            Top             =   1140
            Width           =   1155
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Server Port:"
            Height          =   255
            Index           =   4
            Left            =   180
            TabIndex        =   10
            Top             =   720
            Width           =   1155
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Server Name:"
            Height          =   255
            Index           =   3
            Left            =   180
            TabIndex        =   9
            Top             =   300
            Width           =   1155
         End
      End
      Begin VB.TextBox txtGamesMasterEMail 
         Height          =   315
         Left            =   2640
         TabIndex        =   6
         Top             =   1560
         Width           =   4335
      End
      Begin VB.TextBox txtGalaxyNGHome 
         Height          =   315
         Left            =   2640
         TabIndex        =   5
         Top             =   1080
         Width           =   4335
      End
      Begin VB.TextBox txtServerName 
         Height          =   315
         Left            =   2640
         TabIndex        =   4
         Top             =   660
         Width           =   1695
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Log Errors:"
         Height          =   255
         Index           =   20
         Left            =   1380
         TabIndex        =   47
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Auto Run Games:"
         Height          =   255
         Index           =   19
         Left            =   -74760
         TabIndex        =   42
         Top             =   2820
         Width           =   1875
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Auto Check Mail:"
         Height          =   255
         Index           =   18
         Left            =   -74760
         TabIndex        =   41
         Top             =   2460
         Width           =   1875
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Show Get Mail:"
         Height          =   255
         Index           =   17
         Left            =   -74760
         TabIndex        =   40
         Top             =   2100
         Width           =   1875
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Show Send Mail:"
         Height          =   255
         Index           =   16
         Left            =   -74760
         TabIndex        =   39
         Top             =   1740
         Width           =   1875
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Show Games:"
         Height          =   255
         Index           =   15
         Left            =   -74760
         TabIndex        =   37
         Top             =   1380
         Width           =   1875
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Minimize At Startup:"
         Height          =   255
         Index           =   14
         Left            =   -74760
         TabIndex        =   34
         Top             =   1020
         Width           =   1875
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Start with Windows:"
         Height          =   255
         Index           =   13
         Left            =   -74760
         TabIndex        =   33
         Top             =   660
         Width           =   1875
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Outbox:"
         Height          =   255
         Index           =   11
         Left            =   -74700
         TabIndex        =   25
         Top             =   5040
         Width           =   1155
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Inbox:"
         Height          =   255
         Index           =   10
         Left            =   -74700
         TabIndex        =   24
         Top             =   4620
         Width           =   1155
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "GalaxyNG Home Folder:"
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   3
         Top             =   1140
         Width           =   2355
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Games Master E-Mail Address:"
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   2
         Top             =   1620
         Width           =   2355
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "GalaxyNG Server Name:"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   1
         Top             =   720
         Width           =   2355
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub LoadOptions()
    txtServerName = Options.ServerName
    txtGalaxyNGHome = Options.GalaxyNGHome
    txtGamesMasterEMail = Options.GamesMasterEMail
    
    chkLogErrors = IIf(Options.LogErrors, vbChecked, vbUnchecked)
    
    txtPOPServer = Options.POPServer
    txtPOPServerPort = Options.POPServerPort
    txtPOPUserID = Options.POPUserID
    txtPOPPassword = Options.POPPassword
    txtCheckMailInterval = Options.CheckMailInterval
    
    txtSMTPServer = Options.SMTPServer
    txtSMTPServerPort = Options.SMTPServerPort
    txtSMTPFromAddress = Options.SMTPFromAddress
    txtInbox = Options.Inbox
    txtOutBox = Options.Outbox
    
    chkStartWithWindows = IIf(Options.StartWithWindows, vbChecked, vbUnchecked)
    chkMinimizeAtStartup = IIf(Options.MinimizeatStartup, vbChecked, vbUnchecked)
    chkShowGames = IIf(Options.ShowGames, vbChecked, vbUnchecked)
    chkShowGetMail = IIf(Options.ShowGetMail, vbChecked, vbUnchecked)
    chkShowSendMail = IIf(Options.ShowSendMail, vbChecked, vbUnchecked)
    chkAutoCheckMail = IIf(Options.AutoCheckMail, vbChecked, vbUnchecked)
    chkAutoRunGames = IIf(Options.AutoRunGames, vbChecked, vbUnchecked)

End Sub

Public Sub SaveOptions()
    Options.ServerName = txtServerName
    Options.GalaxyNGHome = txtGalaxyNGHome
    Options.GamesMasterEMail = txtGamesMasterEMail
    
    Options.LogErrors = (chkLogErrors = vbChecked)
    
    Options.POPServer = txtPOPServer
    Options.POPServerPort = txtPOPServerPort
    Options.POPUserID = txtPOPUserID
    Options.POPPassword = txtPOPPassword
    Options.CheckMailInterval = txtCheckMailInterval
    
    Options.SMTPServer = txtSMTPServer
    Options.SMTPServerPort = txtSMTPServerPort
    Options.SMTPFromAddress = txtSMTPFromAddress
    Options.Inbox = txtInbox
    Options.Outbox = txtOutBox
    
    Options.StartWithWindows = (chkStartWithWindows = vbChecked)
    Options.MinimizeatStartup = (chkMinimizeAtStartup = vbChecked)
    Options.ShowGames = (chkShowGames = vbChecked)
    Options.ShowGetMail = (chkShowGetMail = vbChecked)
    Options.ShowSendMail = (chkShowSendMail = vbChecked)
    Options.AutoCheckMail = (chkAutoCheckMail = vbChecked)
    Options.AutoRunGames = (chkAutoRunGames = vbChecked)
    
    Options.SaveSettings
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdRefresh_Click()
    Call LoadOptions
End Sub

Private Sub cmdSave_Click()
    Call SaveOptions
    If Options.StartWithWindows Then
        Call SaveString(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", App.EXEName, App.Path & "\" & App.EXEName & ".exe")
    Else
        Call DelSetting(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", App.EXEName)
    End If
    
    Unload Me
End Sub

Private Sub cmdViewErrors_Click()
    ShellOpen LogFilename
End Sub

Private Sub Form_Load()
    Me.Icon = MainForm.Icon
    Call LoadOptions
End Sub

