Attribute VB_Name = "modGlobal"
Option Explicit

Private mcolMessages As Collection
Public GalaxyNG As GalaxyNG
Public MainForm As frmMain
Public INIFile As INIFile
Public Options As Options

Public Sub Main()
    Set Options = New Options
    Options.SaveSettings
    Set GalaxyNG = New GalaxyNG
    Set MainForm = New frmMain
    MainForm.Show
End Sub

Public Function GetFileName(ByVal FilePath As String) As String
    Dim i As Long
    Dim j As Long
    
    i = InStrRev(FilePath, "\")
    j = InStrRev(FilePath, ".")
    If j = 0 Then j = Len(FilePath)
    GetFileName = Mid(FilePath, i + 1, j - i - 1)
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
    Dim ret As Long
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
    
    lngLength = FileLen(strPath)
    strBuffer = String(lngLength, " ")
    
    intFN = FreeFile
    Open strPath For Binary As #intFN
    Get intFN, , strBuffer
    Close intFN
    GetFile = strBuffer
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
            strMessage = Options.GetMessage("Header")
            strMessage = strMessage & Options.GetMessage("NotifyUser", "24 hours", strRace)
            strMessage = strMessage & vbNewLine & Options.GetMessage("Footer")
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
    
    strCommand = Options.GetMessage("run_game")
    strCommand = Replace(strCommand, "[turn]", objGame.NextTurn)
    strCommand = Replace(strCommand, "[game]", strGame)
    
    Call RunCommandFile(strCommand)
    Call SendReports(strGame)
End Sub

Public Sub LogError(ByVal lngError As Long, _
                    ByVal strError As String, _
                    ByVal strSource As String, _
                    ByVal strModule As String, _
                    ByVal strProcedure As String)
    Dim strMessage As String
    
    strMessage = "Error: " & CStr(lngError) & " - " & strError & vbNewLine & _
                 "Source: " & strSource
    If strModule <> "" Then strMessage = strMessage & vbNewLine & _
                                        "Module: " & strModule
    If strProcedure <> "" Then strMessage = strMessage & vbNewLine & _
                                        "Procedure: " & strProcedure
    MsgBox strMessage, vbCritical + vbOKOnly, App.Title & " Error"
End Sub

