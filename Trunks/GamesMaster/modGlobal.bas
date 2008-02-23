Attribute VB_Name = "modGlobal"
Option Explicit

Private mcolMessages As Collection
Public GalaxyNG As GalaxyNG
Public MainForm As frmMain
Public INIFile As INIFile
Public Options As Options
Private mdtStartTime As Date

Public Const gcStuffMaxSize = 200

Public Sub Main()
    Dim blnShowOptions As Boolean
    Dim strINIFile As String
    Dim fOptions As frmOptions
    
    Call GetCommandLine(blnShowOptions, strINIFile)
    Set Options = New Options
    If strINIFile <> "" Then
        Options.INIFileName = strINIFile
    End If
    Call Options.LoadSettings
    Call Options.SaveSettings
    
    If blnShowOptions Then
        Set fOptions = New frmOptions
        Load fOptions
        fOptions.Show
    Else
        mdtStartTime = Now
        Set GalaxyNG = New GalaxyNG
        Set MainForm = New frmMain
        MainForm.Show
    End If
End Sub

Private Sub GetCommandLine(ByRef blnShowOptions As Boolean, ByRef strINIFile As String)
    Dim i As Long
    Dim strLine As String
    
    ' get the command line
    strLine = LCase(Command$)
    
    ' Shw the Options Dialog Only
    If InStr(1, strLine, "-showoptions") > 0 Then
        blnShowOptions = True
        strLine = Replace(strLine, "-showoptions", "")
    End If
    
    'Clean up the remainder of the line
    strLine = Trim(strLine)
    strLine = Replace(strLine, vbTab, " ")
    
    While InStr(1, strLine, "  ") > 0
        strLine = Replace(strLine, "  ", " ")
    Wend
    
    'Anything left is a path to the INIFile to use
    If strLine <> "" Then
        If InStr(1, strLine, "\") = 0 Then
            strINIFile = App.Path & "\" & strLine
        Else
            strINIFile = strLine
        End If
    End If
    
    
End Sub

Public Function GetFileName(ByVal FilePath As String) As String
    Dim i As Long
    Dim j As Long
    
    i = InStrRev(FilePath, "\")
    j = InStrRev(FilePath, ".")
    If j = 0 Then j = Len(FilePath)
    GetFileName = Mid(FilePath, i + 1, j - i - 1)
End Function

Public Function GetFullFileName(ByVal strPath As String) As String
    Dim i As Long
    
    i = InStrRev(strPath, "\")
    If i = 0 Then
        GetFullFileName = strPath
    Else
        GetFullFileName = Mid(strPath, i + 1)
    End If
End Function

Public Function RunGalaxyNG(Optional ByVal strParameters As String) As Boolean
    Dim strCommand As String
    
    strCommand = _
    "SET GalaxyNGHome=." & vbNewLine & _
    "CD """ & Options.GalaxyNGHome & """ " & vbNewLine & _
    """" & Options.GalaxyNGexe & """ " & strParameters & vbNewLine
    Call RunCommandFile(strCommand)
End Function

Public Function RunCommandFile(ByVal strCommands) As Boolean
    Dim Ret As Long
    Dim intFN As Integer
    
    ' Write a command file with the commands wanted
    intFN = FreeFile
    If Dir(Options.CommandFileName) <> "" Then
        Kill Options.CommandFileName
    End If
    Open Options.CommandFileName For Output As intFN
    Print #intFN, strCommands;
    Close intFN
    
    'Run the command file and wait for completion
    Call ShellWait(Options.CommandFileName, SW_HIDE)
    
    'Delete the command file
    Kill Options.CommandFileName
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
    
    On Error GoTo ErrorTag
    
    If Dir(strPath) = "" Then Exit Function
    lngLength = FileLen(strPath)
    strBuffer = String(lngLength, " ")
    
    intFN = FreeFile
    Open strPath For Binary As #intFN
    Get intFN, , strBuffer
    Close intFN
    GetFile = strBuffer
    Exit Function
    
ErrorTag:
    Call LogError(Err.Number, Err.Description, Err.Source, "modGlobal", "GetFile", "File: " & strPath)
    GetFile = ""
    
End Function

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

Public Sub LogError(ByVal lngError As Long, _
                    ByVal strError As String, _
                    Optional ByVal strSource As String = "", _
                    Optional ByVal strModule As String = "", _
                    Optional ByVal strProcedure As String = "", _
                    Optional ByVal strData As String = "")
    Dim strMessage As String
    
    strMessage = "Error: " & CStr(lngError) & " - " & strError
    If strSource <> "" Then strMessage = strMessage & vbNewLine & _
                                         "    Source: " & strSource
    If strModule <> "" Then strMessage = strMessage & vbNewLine & _
                                        "    Module: " & strModule
    If strProcedure <> "" Then strMessage = strMessage & vbNewLine & _
                                        "    Procedure: " & strProcedure
    If strData <> "" Then strMessage = strMessage & vbNewLine & _
                                        "    Debug Data follows... " & vbNewLine & _
                                        strData
    If Options.LogErrors Then
        Call WriteLogFile(strMessage)
    End If
'    MsgBox strMessage, vbCritical + vbOKOnly, App.Title & " Error"
End Sub

Public Sub WriteLogFile(ByVal strData As String)
    Dim intFN As Integer
    
    intFN = FreeFile
    Open LogFilename For Append As #intFN
    Print #intFN, Format(Now, "hh:nn:ss dd-mmm-yyyy") & ": " & strData
    Close #intFN
End Sub

Public Function LogFilename() As String
    LogFilename = App.Path & "\" & App.EXEName & ".log"
End Function

Public Sub CreateGame(ByVal strTemplate As String)
    Dim objGame As Game
    Dim objTemplate As Template
    
    Set objGame = GalaxyNG.Games(strTemplate)
    Set objTemplate = objGame.Template
    Call ApplyDesign(objTemplate)
    Call objTemplate.Save
    Call RunGalaxyNG("-create """ & objGame.TemplateFile & """ >" & strTemplate & ".txt")
End Sub

Public Sub StartGame(ByVal strGame As String)
    Dim objGame As Game
    Dim strBuffer As String
    
    GalaxyNG.Games.Refresh
    Set objGame = GalaxyNG.Games(strGame)
    objGame.Refresh
    
    Call RunGalaxyNG("-mail0 " & strGame)
    
    ' Change the date on the Next Turn file
    strBuffer = GetFile(Options.GalaxyNGNextTurn(strGame))
    Call SaveFile(Options.GalaxyNGNextTurn(strGame), strBuffer)
    
    Call SendReports(strGame)

End Sub

Public Sub RunGame(ByVal strGame As String)
    Dim strCommand As String
    Dim objGames As Games
    Dim objGame As Game
    
    Set objGames = New Games
    objGames.Refresh
    Set objGame = objGames(strGame)
    Call objGame.Refresh
    
    strCommand = Options.GetMessage("run_game")
    strCommand = Replace(strCommand, "[turn]", objGame.NextTurn)
    strCommand = Replace(strCommand, "[game]", strGame)
    
    Call RunCommandFile(strCommand)
    Call SendReports(strGame)
'    Call MainForm.SendMail.Send
End Sub

Public Sub ResendReports(ByVal strGame As String)
    Dim objGame As Game
    
    Set objGame = GalaxyNG.Games(strGame)
    Call objGame.Refresh
    
    Call SendReports(strGame)
    Call MainForm.SendMail.Send
End Sub

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
        If objRace.flag(R_DEAD) Then
        ElseIf objGame.FinalOrdersReceived(strRace) Then
        ElseIf objGame.OrdersReceived(strRace) Then
        ElseIf objGame.NotificationSent(strRace) Then
        Else
            strMessage = Options.GetMessage("Header")
            strMessage = strMessage & Options.GetMessage("NotifyUser", "24 hours", strRace)
            strMessage = strMessage & vbNewLine & Options.GetMessage("Footer")
            strMessage = Replace(strMessage, "[turn]", objGame.NextTurn)
            strMessage = Replace(strMessage, "[game]", strGame)
            Call SendEMail(objRace.EMail, _
                    "[GNG] " & objGame.GameName & " turn " & objGame.NextTurn & _
                    " Notification for " & strRace, _
                    strMessage)
            Call SaveFile(Options.GalaxyNGOrders & objGame.GameName & "\" & strRace & "_" & objGame.NextTurn & ".notify", _
            "Notified " & Format(Now, "hh:nn:ss dddd, d mmmm yyyy"))
        End If
    Next objRace

End Sub

Public Function QuoteText(ByVal strText As String) As String
    Dim strTemp As String
    strTemp = "> " & Replace(strText, vbNewLine, vbNewLine & "> ")
    If Right(strTemp, 2) = "> " Then
        strTemp = Left(strTemp, Len(strTemp) - 2)
    End If
    QuoteText = strTemp
End Function

Public Function InIDE() As Boolean
' Return whether the program is running live or in the development IDE
    On Error GoTo ErrorTag
    
    Debug.Print 1 / 0
    InIDE = False
    Exit Function

ErrorTag:
    InIDE = True
End Function

Public Function CheckRestart() As Boolean
' Restart the program at Midnight if it has been running for more than 3 hours
    Dim objForm As Form
    
    CheckRestart = False
    If CDate(Now - Int(Now)) > #1:00:00 AM# Then Exit Function
    If CDate(Now - mdtStartTime) < #3:00:00 AM# Then Exit Function
    
    CheckRestart = True
    ' Stop timers
    MainForm.tmrGalaxyNG.Enabled = False
    MainForm.tmrMail = False
    
    ' Start a new instance
    Shell App.Path & "\" & App.EXEName & ".exe", vbNormalNoFocus
    
    ' Shutdown
    Call MainForm.mnuExit_Click
    
End Function

Public Function PI() As Single
    PI = Atn(1) * 4
End Function

Public Function Round(ByVal sngNo As Single, Optional ByVal lngPlaces As Long = 2)
    Dim sngFactor As Single
    
    sngFactor = 10 ^ lngPlaces
    Round = Int(sngNo * sngFactor) / sngFactor

End Function

Public Sub DeleteGame(ByVal strGame As String)
    On Error Resume Next
    Kill Options.GalaxyNGNextTurn(strGame)
    Kill Options.GalaxyNGData & strGame & "\0.New"
    RmDir Options.GalaxyNGData & strGame
    RmDir Options.GalaxyNGOrders & strGame
    RmDir Options.GalaxyNGReports & strGame
    RmDir Options.GalaxyNGStatistics & strGame
End Sub

Public Sub SaveGridSettings(ByVal Grid As MSHFlexGrid, Optional ByVal ID As String = "")
    Dim c As Long
    
    
    With Grid
        ID = ID & "." & Grid.Name
        For c = 0 To .Cols - 1
            Call SaveSetting(App.Title, ID, "Col" & CStr(c), .ColWidth(c))
        Next c
    End With
End Sub

Public Sub LoadGridSettings(ByVal Grid As MSHFlexGrid, Optional ByVal ID As String = "")
    Dim c As Long
    
    With Grid
        ID = ID & "." & Grid.Name
        For c = 0 To .Cols - 1
            .ColWidth(c) = GetSetting(App.Title, ID, "Col" & CStr(c), .ColWidth(c))
        Next c
    End With
End Sub

Public Sub DeleteGridSettings(ByVal Grid As MSHFlexGrid, Optional ByVal ID As String = "")
    Dim c As Long
    
    With Grid
        ID = ID & "." & Grid.Name
        For c = 0 To .Cols - 1
            Call DeleteSetting(App.Title, ID, "Col" & CStr(c))
        Next c
    End With
End Sub

