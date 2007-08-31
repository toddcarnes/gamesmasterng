Attribute VB_Name = "modGlobal"
Option Explicit

Private mcTurn As String
Private mcRace As String
Private NextTurnFile As String
Private GamesMasterReportFile As String
Private RaceReportFile As String

Public GalaxyNGHome As String
Public GalaxyNGData As String
Public GalaxyNGReports As String
Public GalaxyNGOrders As String
Public GalaxyNGNotices As String
Public GalaxyNGStatistics As String
Public GalaxyNGLog As String
Public GalaxyNGexe As String
Public GalaxyNG As GalaxyNG

Public Sub Main()
    Dim fMain As frmMain
    Call LoadSettings
    Set GalaxyNG = New GalaxyNG
    Set fMain = New frmMain
    fMain.Show
End Sub

Public Function GalaxyNGNextTurn(ByVal Game As String) As String
    GalaxyNGNextTurn = GalaxyNGData & Game & "\" & NextTurnFile
End Function

Public Function GamesMasterReport(ByVal Game As String, ByVal Turn As Long) As String
    Dim strFile As String
    
    strFile = Replace(GamesMasterReportFile, mcTurn, CStr(Turn))
    GamesMasterReport = GalaxyNGReports & Game & strFile
End Function

Public Function RaceReport(ByVal Game As String, ByVal Race As String, Turn As Long) As String
    Dim strFile As String
    
    strFile = Replace(RaceReportFile, mcRace, Race)
    strFile = Replace(strFile, mcTurn, CStr(Turn))
    RaceReport = GalaxyNGReports & Game & strFile
End Function

Public Function GetFileName(ByVal FilePath As String) As String
    Dim i As Long
    Dim j As Long
    
    i = InStrRev(FilePath, "\")
    j = InStrRev(FilePath, ".")
    If j = 0 Then j = Len(FilePath)
    GetFileName = Mid(FilePath, i + 1, j - i - 1)
End Function

Private Sub LoadSettings()
    Dim objINI As INIFile
    
    Set objINI = New INIFile
    With objINI
        .File = App.Path & "\" & App.EXEName & ".ini"
        mcTurn = .GetSetting("Constants", "Turn", "[turn]")
        mcRace = .GetSetting("Constants", "Race", "[race]")
        
        GalaxyNGHome = .GetSetting("Folders", "GalaxyNGHome", App.Path & "\")
        GalaxyNGData = .GetSetting("Folders", "GalaxyNGData", GalaxyNGHome & "data\")
        GalaxyNGReports = .GetSetting("Folders", "GalaxyNGReports", GalaxyNGHome & "reports\")
        GalaxyNGOrders = .GetSetting("Folders", "GalaxyNGOrders", GalaxyNGHome & "orders\")
        GalaxyNGNotices = .GetSetting("Folders", "GalaxyNGNotices", GalaxyNGHome & "notices\")
        GalaxyNGStatistics = .GetSetting("Folders", "GalaxyNGStatistics", GalaxyNGHome & "statistics\")
        GalaxyNGLog = .GetSetting("Folders", "GalaxyNGLog", GalaxyNGHome & "log\")
    
        NextTurnFile = .GetSetting("FileNames", "NextTurn", "next_turn")
        GamesMasterReportFile = .GetSetting("FileNames", "GamesMasterReport", "NG_GameMaster_" & mcTurn & ".txt")
        RaceReportFile = .GetSetting("FileNames", "RaceReport", mcRace & "_" & mcTurn & ".txt")
        GalaxyNGexe = .GetSetting("FileNames", "Executable", GalaxyNGHome & "GalaxyNG.exe")
    End With
End Sub

Private Sub SaveSettings()
    Dim objINI As INIFile
    
    Set objINI = New INIFile
    With objINI
        .File = App.Path & "\" & App.EXEName & ".ini"
        Call .SaveSetting("Constants", "Turn", mcTurn)
        Call .SaveSetting("Constants", "Race", mcRace)
        
        Call .SaveSetting("Folders", "GalaxyNGHome", GalaxyNGHome)
        Call .SaveSetting("Folders", "GalaxyNGData", GalaxyNGData)
        Call .SaveSetting("Folders", "GalaxyNGReports", GalaxyNGReports)
        Call .SaveSetting("Folders", "GalaxyNGOrders", GalaxyNGOrders)
        Call .SaveSetting("Folders", "GalaxyNGNotices", GalaxyNGNotices)
        Call .SaveSetting("Folders", "GalaxyNGStatistics", GalaxyNGStatistics)
        Call .SaveSetting("Folders", "GalaxyNGLog", GalaxyNGLog)
    
        Call .SaveSetting("FileNames", "NextTurn", NextTurnFile)
        Call .SaveSetting("FileNames", "GamesMasterReport", GamesMasterReportFile)
        Call .SaveSetting("FileNames", "RaceReport", RaceReportFile)
        Call .SaveSetting("FileNames", "Executable", GalaxyNGexe)
    End With
End Sub

Public Function RunGalaxyNG(Optional ByVal strParameters As String) As Boolean
    Dim ret As Long
    Dim intFN As Integer
    Dim strCommandFile As String
    
    ' Write a command file with the commands wanted
    strCommandFile = App.EXEName & "_1.cmd"
    intFN = FreeFile
    Open strCommandFile For Output As intFN
    Print #intFN, "SET GalaxyNGHome=."
    Print #intFN, "CD """ & GalaxyNGHome & """ "
    Print #intFN, """" & GalaxyNGexe & """ " & strParameters
    Close intFN
    
    'Run the command file and wait for completion
    Call ShellWait(strCommandFile, SW_HIDE)
    
    'Delete the command file
    Kill strCommandFile
End Function
