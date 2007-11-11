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
      TabIndex        =   31
      Top             =   5820
      Width           =   1155
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   30
      Top             =   5820
      Width           =   1155
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   375
      Left            =   1560
      TabIndex        =   29
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
      Tabs            =   2
      TabsPerRow      =   2
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
      Tab(0).Control(3)=   "txtServerName"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtGalaxyNGHome"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtGamesMasterEMail"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "E-Mail"
      TabPicture(1)   =   "frmOptions.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label(10)"
      Tab(1).Control(1)=   "Label(11)"
      Tab(1).Control(2)=   "frPopServer"
      Tab(1).Control(3)=   "frSMTPServer"
      Tab(1).Control(4)=   "txtInbox"
      Tab(1).Control(5)=   "txtOutBox"
      Tab(1).ControlCount=   6
      Begin VB.TextBox txtOutBox 
         Height          =   315
         Left            =   -73440
         TabIndex        =   26
         Top             =   4980
         Width           =   5355
      End
      Begin VB.TextBox txtInbox 
         Height          =   315
         Left            =   -73440
         TabIndex        =   25
         Top             =   4560
         Width           =   5355
      End
      Begin VB.Frame frSMTPServer 
         Caption         =   "SMTP Server"
         Height          =   1695
         Left            =   -74820
         TabIndex        =   12
         Top             =   2700
         Width           =   7275
         Begin VB.TextBox txtSMTPFromAddress 
            Height          =   315
            Left            =   1440
            TabIndex        =   22
            Top             =   1140
            Width           =   4335
         End
         Begin VB.TextBox txtSMTPServerPort 
            Height          =   315
            Left            =   1440
            TabIndex        =   21
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox txtSMTPServer 
            Height          =   315
            Left            =   1440
            TabIndex        =   20
            Top             =   300
            Width           =   3675
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "From Address:"
            Height          =   255
            Index           =   9
            Left            =   180
            TabIndex        =   15
            Top             =   1140
            Width           =   1155
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Server Port:"
            Height          =   255
            Index           =   8
            Left            =   180
            TabIndex        =   14
            Top             =   720
            Width           =   1155
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Server Name:"
            Height          =   255
            Index           =   7
            Left            =   180
            TabIndex        =   13
            Top             =   300
            Width           =   1155
         End
      End
      Begin VB.Frame frPopServer 
         Caption         =   "POP Server"
         Height          =   2055
         Left            =   -74820
         TabIndex        =   7
         Top             =   540
         Width           =   7275
         Begin VB.TextBox txtCheckMailInterval 
            Height          =   315
            Left            =   5640
            TabIndex        =   28
            Top             =   1500
            Width           =   615
         End
         Begin VB.TextBox txtPOPPassword 
            Height          =   315
            Left            =   1440
            TabIndex        =   19
            Top             =   1500
            Width           =   1695
         End
         Begin VB.TextBox txtPOPUserID 
            Height          =   315
            Left            =   1440
            TabIndex        =   18
            Top             =   1080
            Width           =   1695
         End
         Begin VB.TextBox txtPOPServerPort 
            Height          =   315
            Left            =   1440
            TabIndex        =   17
            Top             =   660
            Width           =   615
         End
         Begin VB.TextBox txtPOPServer 
            Height          =   315
            Left            =   1440
            TabIndex        =   16
            Top             =   240
            Width           =   3675
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Check Mail Interval:"
            Height          =   255
            Index           =   12
            Left            =   3840
            TabIndex        =   27
            Top             =   1560
            Width           =   1635
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Password:"
            Height          =   255
            Index           =   6
            Left            =   180
            TabIndex        =   11
            Top             =   1560
            Width           =   1155
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "User Name:"
            Height          =   255
            Index           =   5
            Left            =   180
            TabIndex        =   10
            Top             =   1140
            Width           =   1155
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Server Port:"
            Height          =   255
            Index           =   4
            Left            =   180
            TabIndex        =   9
            Top             =   720
            Width           =   1155
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "Server Name:"
            Height          =   255
            Index           =   3
            Left            =   180
            TabIndex        =   8
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
         Caption         =   "Outbox:"
         Height          =   255
         Index           =   11
         Left            =   -74700
         TabIndex        =   24
         Top             =   5040
         Width           =   1155
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Inbox:"
         Height          =   255
         Index           =   10
         Left            =   -74700
         TabIndex        =   23
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
End Sub

Public Sub SaveOptions()
    Options.ServerName = txtServerName
    Options.GalaxyNGHome = txtGalaxyNGHome
    Options.GamesMasterEMail = txtGamesMasterEMail
    
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
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = MainForm.Icon
    Call LoadOptions
End Sub
