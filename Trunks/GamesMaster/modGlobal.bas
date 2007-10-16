Attribute VB_Name = "modGlobal"
Option Explicit

Private mcTurn As String
Private mcRace As String
Private NextTurnFile As String
Private GamesMasterReportFile As String
Private RaceReportFile As String
Private RaceMachineFile As String
Private mcolMessages As Collection

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

Public Const gcOrdersFileName = "orders.txt"
Public Const gcForecastFileName = "forecast.txt"

Public MainForm As frmMain
Public INIFile As INIFile

Public Function gcCommandFileName() As String
    gcCommandFileName = App.EXEName & "_execute.cmd"
End Function

Public Sub Main()
    Call LoadSettings
    Call SaveSettings
    Set GalaxyNG = New GalaxyNG
    Set MainForm = New frmMain
    MainForm.Show
End Sub

Public Function GalaxyNGNextTurn(ByVal Game As String) As String
    GalaxyNGNextTurn = GalaxyNGData & Game & "\" & NextTurnFile
End Function

Public Function GamesMasterReport(ByVal Game As String, ByVal Turn As Long) As String
    Dim strFile As String
    
    strFile = Replace(GamesMasterReportFile, mcTurn, CStr(Turn))
    GamesMasterReport = GalaxyNGReports & Game & "\" & strFile
End Function

Public Function RaceReport(ByVal Game As String, ByVal Race As String, ByVal Turn As Long) As String
    Dim strFile As String
    
    strFile = Replace(RaceReportFile, mcRace, Race)
    strFile = Replace(strFile, mcTurn, CStr(Turn))
    RaceReport = GalaxyNGReports & Game & "\" & strFile
End Function

Public Function RaceMachineReport(ByVal Game As String, ByVal Race As String, ByVal Turn As Long) As String
    Dim strFile As String
    
    strFile = Replace(RaceMachineFile, mcRace, Race)
    strFile = Replace(strFile, mcTurn, CStr(Turn))
    RaceMachineReport = GalaxyNGReports & Game & "\" & strFile
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
    Set INIFile = New INIFile
    With INIFile
        .File = App.Path & "\" & App.EXEName & ".ini"
        mcTurn = .GetSetting("Constants", "Turn", "[turn]")
        mcRace = .GetSetting("Constants", "Race", "[race]")
        
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
    
        NextTurnFile = .GetSetting("FileNames", "NextTurn", "next_turn")
        GamesMasterReportFile = .GetSetting("FileNames", "GamesMasterReport", "NG_GameMaster_" & mcTurn & ".txt")
        RaceReportFile = .GetSetting("FileNames", "RaceReport", mcRace & "_" & mcTurn & ".txt")
        RaceMachineFile = .GetSetting("FileNames", "RaceMachineReport", mcRace & "_" & mcTurn & ".m")
        GalaxyNGexe = .GetSetting("FileNames", "Executable", GalaxyNGHome & "GalaxyNG.exe")
    End With
    If Dir(Inbox, vbDirectory) = "" Then
        MkDir Inbox
    End If
    If Dir(Outbox, vbDirectory) = "" Then
        MkDir Outbox
    End If
End Sub

Private Sub SaveSettings()
    With INIFile
        .File = App.Path & "\" & App.EXEName & ".ini"
        Call .SaveSetting("Constants", "Turn", mcTurn)
        Call .SaveSetting("Constants", "Race", mcRace)
        
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
        
        Call .SaveSetting("FileNames", "NextTurn", NextTurnFile)
        Call .SaveSetting("FileNames", "GamesMasterReport", GamesMasterReportFile)
        Call .SaveSetting("FileNames", "RaceReport", RaceReportFile)
        Call .SaveSetting("FileNames", "RaceMachineReport", RaceMachineFile)
        Call .SaveSetting("FileNames", "Executable", GalaxyNGexe)
    End With
End Sub

Public Function RunGalaxyNG(Optional ByVal strParameters As String) As Boolean
    Dim strCommand As String
    
    strCommand = _
    "SET GalaxyNGHome=." & vbNewLine & _
    "CD """ & GalaxyNGHome & """ " & vbNewLine & _
    """" & GalaxyNGexe & """ " & strParameters & vbNewLine
    Call RunCommandFile(strCommand)
End Function

Public Function RunCommandFile(ByVal strCommands) As Boolean
    Dim ret As Long
    Dim intFN As Integer
    
    ' Write a command file with the commands wanted
    intFN = FreeFile
    If Dir(gcCommandFileName) <> "" Then
        Kill gcCommandFileName
    End If
    Open gcCommandFileName For Output As intFN
    Print #intFN, strCommands;
    Close intFN
    
    'Run the command file and wait for completion
    Call ShellWait(gcCommandFileName, SW_HIDE)
    
    'Delete the command file
    Kill gcCommandFileName
End Function

Public Function GetAddress(ByVal strEMail As String) As String
    Dim i1 As Long
    Dim i2 As Long
    
    i1 = InStr(1, strEMail, "<")
    If i1 = 0 Then
        GetAddress = Trim(strEMail)
    Else
        i2 = InStr(i1, strEMail, ">")
        GetAddress = Trim(Mid(strEMail, i1 + 1, i2 - i1 - 1))
    End If
    
End Function

Public Function GetFile(ByVal strPath As String) As String
    Dim intFN As Integer
    Dim strBuffer As String
    Dim lngLength As Long
    
    lngLength = FileLen(strPath)
    strBuffer = String(lngLength, " ")
    
    intFN = FreeFile
    Open strPath For Binary As #intFN
    Get intFN, , strBuffer
    Close intFN
    GetFile = strBuffer
End Function

Public Function GetMessage(ByVal strKey As String, ParamArray Parm() As Variant) As String
    Dim strMessage As String
    Dim i As Long
    If mcolMessages Is Nothing Then
        Call LoadMessages
    End If
    On Error Resume Next
    strMessage = mcolMessages(strKey)
    On Error GoTo 0
    
    strMessage = Replace(strMessage, "[galaxynghome]", GalaxyNGHome)
    strMessage = Replace(strMessage, "[galaxyngexe]", GalaxyNGexe)
    strMessage = Replace(strMessage, "[orders]", GalaxyNGOrders)
    strMessage = Replace(strMessage, "[reports]", GalaxyNGReports)
    strMessage = Replace(strMessage, "[data]", GalaxyNGData)
    strMessage = Replace(strMessage, "[log]", GalaxyNGLog)
    strMessage = Replace(strMessage, "[notices]", GalaxyNGNotices)
    strMessage = Replace(strMessage, "[statistics]", GalaxyNGStatistics)
    strMessage = Replace(strMessage, "[gamesmaster]", GamesMasterEMail)
    strMessage = Replace(strMessage, "[server]", ServerName)
    
    If Not IsEmpty(Parm) Then
        For i = LBound(Parm) To UBound(Parm)
            strMessage = Replace(strMessage, "[" & CStr(i + 1) & "]", Parm(i))
        Next i
    End If
    GetMessage = strMessage
End Function

Private Sub LoadMessages()
    Dim varMessage As Variant
    Dim lngNo As Long
    Dim strKey As String
    Dim intFN As Integer
    Dim i As Long
    Dim strLine As String
    Dim blnText As Boolean
    
    Set mcolMessages = New Collection
    intFN = FreeFile
    Open App.Path & "\" & App.EXEName & ".txt" For Input As intFN
    blnText = False
    
    While Not EOF(intFN)
        Line Input #intFN, strLine
        strLine = Trim(strLine)
        If blnText Then
            If strLine = "@" Then
                blnText = False
                mcolMessages.Add varMessage, strKey
            Else
                varMessage = varMessage & strLine & vbNewLine
            End If
        Else
            i = InStr(1, strLine, " ")
            lngNo = Val(Left(strLine, i - 1))
            strKey = Trim(Mid(strLine, i + 1))
            varMessage = ""
            blnText = True
        End If
    Wend
    Close #intFN
End Sub

Public Sub SaveFile(ByVal strFileName As String, ByVal strData As String)
    Dim intFN As Integer
    
    intFN = FreeFile
    Open strFileName For Output As #intFN
    Print #intFN, strData;
    Close #intFN
End Sub

Public Function MarkText(ByVal strSource As String) As String
    MarkText = "> " & Replace(strSource, vbCrLf, vbCrLf & "> ")
End Function

Public Sub NotifyUsers(ByVal strGame As String)
    Dim objGames As Games
    Dim objGame As Game
    Dim objRace As Race
    Dim strRace As String
    Dim strMessage As String
    
    Set objGames = New Games
    objGames.Refresh
    Set objGame = objGames(strGame)
    Call objGame.Refresh

    For Each objRace In objGame.Races
        strRace = objRace.RaceName
        If objRace.Flag(R_DEAD) Then
        ElseIf objGame.FinalOrdersReceived(strRace) Then
        ElseIf objGame.OrdersReceived(strRace) Then
        ElseIf objGame.NotificationSent(strRace) Then
        Else
            strMessage = GetMessage("Header")
            strMessage = strMessage & GetMessage("NotifyUser", "24 hours", strRace)
            strMessage = strMessage & vbNewLine & GetMessage("Footer")
            strMessage = Replace(strMessage, "[turn]", objGame.NextTurn)
            strMessage = Replace(strMessage, "[game]", strGame)
            Call SendEMail(objRace.EMail, _
                    "[GNG] " & objGame.GameName & " turn " & objGame.NextTurn & _
                    " Notification for " & strRace, _
                    strMessage)
        End If
    Next objRace
End Sub

Public Sub RunGame(ByVal strGame As String)
    Dim strCommand As String
    Dim objGames As Games
    Dim objGame As Game
    
    Set objGames = New Games
    objGames.Refresh
    Set objGame = objGames(strGame)
    Call objGame.Refresh
    
    strCommand = GetMessage("run_game")
    strCommand = Replace(strCommand, "[turn]", objGame.NextTurn)
    strCommand = Replace(strCommand, "[game]", strGame)
    
    Call RunCommandFile(strCommand)
    Call SendReports(strGame)
End Sub
