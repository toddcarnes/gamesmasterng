VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"GalaxyNG"
Attribute VB_Ext_KEY = "Member1" ,"INIFile"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Option Compare Text

Private NextTurnFile As String
Private GamesMasterReportFile As String
Private RaceReportFile As String
Private RaceMachineFile As String
Private mcolMessages As Messages

Public TurnConstant As String
Public RaceConstant As String
Public GalaxyNGHomeConstant As String
Public GalaxyngExeConstant As String
Public GamesMasterEMailConstant As String
Public GamesMasterPassword As String
Public ServerNameConstant As String

Public GamesMasterEMail As String
Public GalaxyNGHome As String
Public GalaxyNGData As String
Public GalaxyNGReports As String
Public GalaxyNGOrders As String
Public GalaxyNGNotices As String
Public GalaxyNGStatistics As String
Public GalaxyNGLog As String
Public GalaxyNGexe As String
Public GalaxyNG As GalaxyNG

Public Inbox As String
Public Outbox As String
Public ServerName As String
Public POPServer As String
Public POPServerPort As Long
Public POPUserID  As String
Public POPPassword  As String
Public SMTPServer  As String
Public SMTPServerPort  As Long
Public SMTPFromAddress As String
Public CheckMailInterval As Long
Public AttachReports As Boolean
Public SaveEMail As Boolean

Public StartWithWindows As Boolean
Public MinimizeatStartup As Boolean
Public ShowGames As Boolean
Public ShowGetMail As Boolean
Public ShowSendMail As Boolean
Public AutoCheckMail As Boolean
Public AutoRunGames As Boolean
Public LogErrors As Boolean

Public INIFile As INIFile

Public Property Get Messages() As Messages
    If mcolMessages Is Nothing Then
        Set mcolMessages = New Messages
        Call mcolMessages.Load
    End If
    Set Messages = mcolMessages
End Property

Public Property Set Messages(ByVal colMessages As Messages)
    Set mcolMessages = colMessages
End Property

Public Property Get INIFileName() As String
    INIFileName = INIFile.File
End Property

Public Property Let INIFileName(ByVal strFileName As String)
    INIFile.File = strFileName
End Property

Public Function OrdersFileName() As String
    OrdersFileName = "orders.txt"
End Function

Public Function ForecastFileName() As String
    ForecastFileName = "forecast.txt"
End Function

Public Function ReportFileName() As String
    ReportFileName = "report.txt"
End Function

Public Function CommandFileName() As String
    CommandFileName = App.EXEName & "_execute.cmd"
End Function

Public Function GalaxyNGNextTurn(ByVal Game As String) As String
    GalaxyNGNextTurn = Me.GalaxyNGData & Game & "\" & NextTurnFile
End Function

Public Function GamesMasterReport(ByVal Game As String, ByVal Turn As Long) As String
    Dim strFile As String
    
    strFile = Replace(GamesMasterReportFile, TurnConstant, CStr(Turn))
    GamesMasterReport = Me.GalaxyNGReports & Game & "\" & strFile
End Function

Public Function RaceReport(ByVal Game As String, ByVal Race As String, ByVal Turn As Long) As String
    Dim strFile As String
    
    strFile = Replace(RaceReportFile, RaceConstant, Race)
    strFile = Replace(strFile, TurnConstant, CStr(Turn))
    RaceReport = Me.GalaxyNGReports & Game & "\" & strFile
End Function

Public Function RaceMachineReport(ByVal Game As String, ByVal Race As String, ByVal Turn As Long) As String
    Dim strFile As String
    
    strFile = Replace(RaceMachineFile, RaceConstant, Race)
    strFile = Replace(strFile, TurnConstant, CStr(Turn))
    RaceMachineReport = Me.GalaxyNGReports & Game & "\" & strFile
End Function

Public Sub LoadSettings()
    With INIFile
        TurnConstant = .GetSetting("Constants", "Turn", "[turn]")
        RaceConstant = .GetSetting("Constants", "Race", "[race]")
        GalaxyNGHomeConstant = .GetSetting("Constants", "GalaxyNGHome", "[GalaxyNGHome]")
        GalaxyngExeConstant = .GetSetting("Constants", "GalaxyngExe", "[GalaxyngExe]")
        GamesMasterEMailConstant = .GetSetting("Constants", "GamesMasterEMail", "[GamesMasterEMail]")
        ServerNameConstant = .GetSetting("Constants", "ServerName", "[ServerName]")
        
        GalaxyNGHome = .GetSetting("Folders", "GalaxyNGHome", App.Path & "\")
        GalaxyNGData = GalaxyNGHome & "data\"
        GalaxyNGReports = GalaxyNGHome & "reports\"
        GalaxyNGOrders = GalaxyNGHome & "orders\"
        GalaxyNGNotices = GalaxyNGHome & "notices\"
        GalaxyNGStatistics = GalaxyNGHome & "statistics\"
        GalaxyNGLog = GalaxyNGHome & "log\"
    
        GamesMasterEMail = .GetSetting("EMail", "GamesMasterEMail", "")
        CheckMailInterval = .GetSetting("EMail", "Interval", "5")
        Inbox = .GetSetting("EMail", "Inbox", App.Path & "\Inbox\")
        Outbox = .GetSetting("EMail", "Outbox", App.Path & "\Outbox\")
        ServerName = .GetSetting("EMail", "ServerName", "")
        POPServer = .GetSetting("EMail", "POPServer", "")
        POPServerPort = .GetSetting("EMail", "POPServerPort", "110")
        POPUserID = .GetSetting("EMail", "POPUserID", "")
        POPPassword = .GetSetting("EMail", "POPPassword", "")
        SMTPServer = .GetSetting("EMail", "SMTPServer", "")
        SMTPServerPort = .GetSetting("EMail", "SMTPServerPort", "25")
        SMTPFromAddress = .GetSetting("EMail", "SMTPFromAddress", "")
        AttachReports = .GetSetting("EMail", "AttachReports", True)
        SaveEMail = .GetSetting("EMail", "SaveEMail", True)
    
        NextTurnFile = .GetSetting("FileNames", "NextTurn", "next_turn")
        GamesMasterReportFile = .GetSetting("FileNames", "GamesMasterReport", "NG_GameMaster_" & TurnConstant & ".txt")
        RaceReportFile = .GetSetting("FileNames", "RaceReport", RaceConstant & "_" & TurnConstant & ".txt")
        RaceMachineFile = .GetSetting("FileNames", "RaceMachineReport", RaceConstant & "_" & TurnConstant & ".m")
        GalaxyNGexe = .GetSetting("FileNames", "Executable", GalaxyNGHome & "GalaxyNG.exe")
    
        StartWithWindows = .GetSetting("Startup", "StartWithWindows", False)
        MinimizeatStartup = .GetSetting("Startup", "MinimizeAtStartup", False)
        ShowGames = .GetSetting("Startup", "ShowGames", False)
        ShowGetMail = .GetSetting("Startup", "ShowGetMail", False)
        ShowSendMail = .GetSetting("Startup", "ShowSendMail", False)
        AutoCheckMail = .GetSetting("Startup", "AutoCheckMail", False)
        AutoRunGames = .GetSetting("Startup", "AutoRunGames", False)
        
        LogErrors = .GetSetting("Settings", "LogErrors", False)
        GamesMasterPassword = .GetSetting("Settings", "GamesMasterPassword", Format(Rnd(Now) * 10000, "0000"))
    End With
    
    If Dir(Inbox, vbDirectory) = "" Then
        MkDir Inbox
    End If
    If Dir(Outbox, vbDirectory) = "" Then
        MkDir Outbox
    End If
End Sub

Public Sub SaveSettings()
    With INIFile
        .File = App.Path & "\" & App.EXEName & ".ini"
        Call .SaveSetting("Constants", "Turn", TurnConstant)
        Call .SaveSetting("Constants", "Race", RaceConstant)
        Call .SaveSetting("Constants", "GalaxyNGHome", GalaxyNGHomeConstant)
        Call .SaveSetting("Constants", "GalaxyngExe", GalaxyngExeConstant)
        Call .SaveSetting("Constants", "GamesMasterEMail", GamesMasterEMailConstant)
        Call .SaveSetting("Constants", "ServerName", ServerNameConstant)
        
        Call .SaveSetting("Folders", "GalaxyNGHome", GalaxyNGHome)
    
        Call .SaveSetting("EMail", "GamesMasterEMail", GamesMasterEMail)
        Call .SaveSetting("EMail", "Interval", CheckMailInterval)
        Call .SaveSetting("EMail", "Inbox", Inbox)
        Call .SaveSetting("EMail", "Outbox", Outbox)
        Call .SaveSetting("EMail", "ServerName", ServerName)
        Call .SaveSetting("EMail", "POPServer", POPServer)
        Call .SaveSetting("EMail", "POPServerPort", POPServerPort)
        Call .SaveSetting("EMail", "POPUserID", POPUserID)
        Call .SaveSetting("EMail", "POPPassword", POPPassword)
        Call .SaveSetting("EMail", "SMTPServer", SMTPServer)
        Call .SaveSetting("EMail", "SMTPServerPort", SMTPServerPort)
        Call .SaveSetting("EMail", "SMTPFromAddress", SMTPFromAddress)
        Call .SaveSetting("EMail", "AttachReports", AttachReports)
        Call .SaveSetting("EMail", "SaveEMail", SaveEMail)
        
        Call .SaveSetting("FileNames", "NextTurn", NextTurnFile)
        Call .SaveSetting("FileNames", "GamesMasterReport", GamesMasterReportFile)
        Call .SaveSetting("FileNames", "RaceReport", RaceReportFile)
        Call .SaveSetting("FileNames", "RaceMachineReport", RaceMachineFile)
        Call .SaveSetting("FileNames", "Executable", GalaxyNGexe)
    
        Call .SaveSetting("Startup", "StartWithWindows", StartWithWindows)
        Call .SaveSetting("Startup", "MinimizeAtStartup", MinimizeatStartup)
        Call .SaveSetting("Startup", "ShowGames ", ShowGames)
        Call .SaveSetting("Startup", "ShowGetMail", ShowGetMail)
        Call .SaveSetting("Startup", "ShowSendMail", ShowSendMail)
        Call .SaveSetting("Startup", "AutoCheckMail", AutoCheckMail)
        Call .SaveSetting("Startup", "AutoRunGames", AutoRunGames)
        
        Call .SaveSetting("Settings", "LogErrors", LogErrors)
        Call .SaveSetting("Settings", "GamesMasterPassword", GamesMasterPassword)
    
    End With
End Sub

Public Function GetMessage(ByVal strKey As String, ParamArray Parm() As Variant) As String
    Dim strMessage As String
    Dim i As Long
    Dim objMessage As Message
    
    On Error Resume Next
    Set objMessage = Messages(strKey)
    strMessage = objMessage.Message
    On Error GoTo 0
    
    strMessage = Replace(strMessage, GalaxyNGHomeConstant, GalaxyNGHome)
    strMessage = Replace(strMessage, GalaxyngExeConstant, GalaxyNGexe)
    strMessage = Replace(strMessage, GamesMasterEMailConstant, GamesMasterEMail)
    strMessage = Replace(strMessage, ServerNameConstant, ServerName)
    
    If Not IsEmpty(Parm) Then
        For i = LBound(Parm) To UBound(Parm)
            strMessage = Replace(strMessage, "[" & CStr(i + 1) & "]", Parm(i))
        Next i
    End If
    GetMessage = strMessage
End Function

Public Function DesignDescriptions() As Variant
    DesignDescriptions = Array( _
        "0 - Leave Alone", _
        "1 - On Circle", _
        "2 - On Circle+Middle", _
        "3 - GalaxyNG Random")
End Function

Public Function SeedDescriptions() As Variant
    SeedDescriptions = Array( _
        "0 - No Seeding", _
        "1 - Seed HomeWorld", _
        "2 - Seed Center of Circle", _
        "3 - Seed Home+Center", _
        "4 - Seed MidPoint", _
        "5 - Seed Home+Mid", _
        "6 - Seed Center+Mid", _
        "7 - Seed Home+Center+Mid")
End Function

Private Sub Class_Initialize()
    Set INIFile = New INIFile
    With INIFile
        .File = App.Path & "\" & App.EXEName & ".ini"
    End With
End Sub

