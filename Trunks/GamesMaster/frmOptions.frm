VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   6360
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7800
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   7800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   4920
      TabIndex        =   37
      Top             =   5820
      Width           =   1155
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   36
      Top             =   5820
      Width           =   1155
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   375
      Left            =   1560
      TabIndex        =   35
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
      TabPicture(0)   =   "frmOptions.frx":0CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label(13)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label(14)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtServerName"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtGalaxyNGHome"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtGamesMasterEMail"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "chkLogErrors"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdViewErrors"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "chkSaveEMail"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdReset"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtGamesMasterPassword"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "E-Mail"
      TabPicture(1)   =   "frmOptions.frx":0CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtOutBox"
      Tab(1).Control(1)=   "txtInbox"
      Tab(1).Control(2)=   "frSMTPServer"
      Tab(1).Control(3)=   "frPopServer"
      Tab(1).Control(4)=   "Label(11)"
      Tab(1).Control(5)=   "Label(10)"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Startup"
      TabPicture(2)   =   "frmOptions.frx":0D02
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "chkAutoRunGames"
      Tab(2).Control(1)=   "chkAutoCheckMail"
      Tab(2).Control(2)=   "chkShowGetMail"
      Tab(2).Control(3)=   "chkShowSendMail"
      Tab(2).Control(4)=   "chkShowGames"
      Tab(2).Control(5)=   "chkMinimizeAtStartup"
      Tab(2).Control(6)=   "chkStartWithWindows"
      Tab(2).ControlCount=   7
      Begin VB.TextBox txtGamesMasterPassword 
         Height          =   315
         Left            =   2640
         TabIndex        =   5
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "Reset"
         Height          =   375
         Left            =   2640
         TabIndex        =   11
         Top             =   3120
         Width           =   1155
      End
      Begin VB.CheckBox chkSaveEMail 
         Alignment       =   1  'Right Justify
         Caption         =   "Save EMail:"
         Height          =   255
         Left            =   1620
         TabIndex        =   10
         Top             =   2760
         Width           =   1215
      End
      Begin VB.CommandButton cmdViewErrors 
         Caption         =   "View Errors"
         Height          =   375
         Left            =   3060
         TabIndex        =   9
         Top             =   2340
         Width           =   1095
      End
      Begin VB.CheckBox chkLogErrors 
         Alignment       =   1  'Right Justify
         Caption         =   "Log Errors:"
         Height          =   315
         Left            =   1740
         TabIndex        =   8
         Top             =   2340
         Width           =   1095
      End
      Begin VB.CheckBox chkAutoRunGames 
         Alignment       =   1  'Right Justify
         Caption         =   "Auto Run Games:"
         Height          =   255
         Left            =   -73680
         TabIndex        =   44
         Top             =   2820
         Width           =   1695
      End
      Begin VB.CheckBox chkAutoCheckMail 
         Alignment       =   1  'Right Justify
         Caption         =   "Auto Check Mail:"
         Height          =   255
         Left            =   -73680
         TabIndex        =   43
         Top             =   2460
         Width           =   1695
      End
      Begin VB.CheckBox chkShowGetMail 
         Alignment       =   1  'Right Justify
         Caption         =   "Show Get Mail:"
         Height          =   255
         Left            =   -73500
         TabIndex        =   42
         Top             =   2100
         Width           =   1515
      End
      Begin VB.CheckBox chkShowSendMail 
         Alignment       =   1  'Right Justify
         Caption         =   "Show Send Mail:"
         Height          =   255
         Left            =   -73620
         TabIndex        =   41
         Top             =   1740
         Width           =   1635
      End
      Begin VB.CheckBox chkShowGames 
         Alignment       =   1  'Right Justify
         Caption         =   "Show Games:"
         Height          =   255
         Left            =   -73440
         TabIndex        =   40
         Top             =   1380
         Width           =   1455
      End
      Begin VB.CheckBox chkMinimizeAtStartup 
         Alignment       =   1  'Right Justify
         Caption         =   "Minimize At Startup:"
         Height          =   255
         Left            =   -73800
         TabIndex        =   39
         Top             =   1020
         Width           =   1815
      End
      Begin VB.CheckBox chkStartWithWindows 
         Alignment       =   1  'Right Justify
         Caption         =   "Start with Windows:"
         Height          =   255
         Left            =   -73800
         TabIndex        =   38
         Top             =   660
         Width           =   1815
      End
      Begin VB.TextBox txtOutBox 
         Height          =   315
         Left            =   -73380
         TabIndex        =   32
         Top             =   4980
         Width           =   5355
      End
      Begin VB.TextBox txtInbox 
         Height          =   315
         Left            =   -73380
         TabIndex        =   31
         Top             =   4560
         Width           =   5355
      End
      Begin VB.Frame frSMTPServer 
         Caption         =   "SMTP Server"
         Height          =   1935
         Left            =   -74820
         TabIndex        =   17
         Top             =   2400
         Width           =   7275
         Begin VB.CheckBox chkAttachReports 
            Alignment       =   1  'Right Justify
            Caption         =   "Send Reports as Attachments:"
            Height          =   315
            Left            =   360
            TabIndex        =   28
            Top             =   1500
            Width           =   2475
         End
         Begin VB.TextBox txtSMTPFromAddress 
            Height          =   315
            Left            =   1440
            TabIndex        =   27
            Top             =   1140
            Width           =   4335
         End
         Begin VB.TextBox txtSMTPServerPort 
            Height          =   315
            Left            =   1440
            TabIndex        =   26
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox txtSMTPServer 
            Height          =   315
            Left            =   1440
            TabIndex        =   25
            Top             =   300
            Width           =   3675
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "From Address:"
            Height          =   255
            Index           =   9
            Left            =   180
            TabIndex        =   20
            Top             =   1140
            Width           =   1155
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Server Port:"
            Height          =   255
            Index           =   8
            Left            =   180
            TabIndex        =   19
            Top             =   720
            Width           =   1155
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Server Name:"
            Height          =   255
            Index           =   7
            Left            =   180
            TabIndex        =   18
            Top             =   300
            Width           =   1155
         End
      End
      Begin VB.Frame frPopServer 
         Caption         =   "POP Server"
         Height          =   1935
         Left            =   -74820
         TabIndex        =   12
         Top             =   420
         Width           =   7275
         Begin VB.TextBox txtCheckMailInterval 
            Height          =   315
            Left            =   5640
            TabIndex        =   34
            Top             =   1500
            Width           =   615
         End
         Begin VB.TextBox txtPOPPassword 
            Height          =   315
            Left            =   1440
            TabIndex        =   24
            Top             =   1500
            Width           =   1695
         End
         Begin VB.TextBox txtPOPUserID 
            Height          =   315
            Left            =   1440
            TabIndex        =   23
            Top             =   1080
            Width           =   1695
         End
         Begin VB.TextBox txtPOPServerPort 
            Height          =   315
            Left            =   1440
            TabIndex        =   22
            Top             =   660
            Width           =   615
         End
         Begin VB.TextBox txtPOPServer 
            Height          =   315
            Left            =   1440
            TabIndex        =   21
            Top             =   240
            Width           =   3675
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Check Mail Interval:"
            Height          =   255
            Index           =   12
            Left            =   3840
            TabIndex        =   33
            Top             =   1560
            Width           =   1635
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Password:"
            Height          =   255
            Index           =   6
            Left            =   180
            TabIndex        =   16
            Top             =   1560
            Width           =   1155
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "User Name:"
            Height          =   255
            Index           =   5
            Left            =   180
            TabIndex        =   15
            Top             =   1140
            Width           =   1155
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Server Port:"
            Height          =   255
            Index           =   4
            Left            =   180
            TabIndex        =   14
            Top             =   720
            Width           =   1155
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Server Name:"
            Height          =   255
            Index           =   3
            Left            =   180
            TabIndex        =   13
            Top             =   300
            Width           =   1155
         End
      End
      Begin VB.TextBox txtGamesMasterEMail 
         Height          =   315
         Left            =   2640
         TabIndex        =   7
         Top             =   1920
         Width           =   4335
      End
      Begin VB.TextBox txtGalaxyNGHome 
         Height          =   315
         Left            =   2640
         TabIndex        =   6
         Top             =   1500
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
         Caption         =   "Games Master Password:"
         Height          =   255
         Index           =   14
         Left            =   180
         TabIndex        =   46
         Top             =   1140
         Width           =   2355
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Reset List Column Widths:"
         Height          =   255
         Index           =   13
         Left            =   180
         TabIndex        =   45
         Top             =   3180
         Width           =   2355
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Outbox:"
         Height          =   255
         Index           =   11
         Left            =   -74640
         TabIndex        =   30
         Top             =   5040
         Width           =   1155
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Inbox:"
         Height          =   255
         Index           =   10
         Left            =   -74640
         TabIndex        =   29
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
         Top             =   1560
         Width           =   2355
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Games Master E-Mail Address:"
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   2
         Top             =   1980
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
    txtGamesMasterPassword = Options.GamesMasterPassword
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
    chkAttachReports = IIf(Options.AttachReports, vbChecked, vbUnchecked)
    chkSaveEMail = IIf(Options.SaveEMail, vbChecked, vbUnchecked)
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
    Options.GamesMasterPassword = txtGamesMasterPassword
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
    Options.AttachReports = (chkAttachReports = vbChecked)
    Options.SaveEMail = (chkSaveEMail = vbChecked)
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

Private Sub cmdReset_Click()
    Call DeleteSetting(App.EXEName)
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
    tabOptions.Tab = 0
    Call LoadFormSettings(Me, True)
    Call LoadOptions
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveFormSettings(Me, True)
End Sub

