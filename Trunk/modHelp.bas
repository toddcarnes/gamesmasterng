Attribute VB_Name = "modHelp"
'********************************************************
'   Copyright 2007,2008 Ian Evans.                      *
'   This program is distributed under the terms of the  *
'       GNU General Public License.                     *
'********************************************************
Option Explicit
Option Compare Text

Private Enum udtHelp
    Unknown = 0
    Help = 1
    HelpGames = 2
    HelpAllGames = 3
    HelpGame = 4
End Enum


' Process Assistance Enquiries received by E-Mail
' Subject: Help
' Subject: Help Games
' Subject: Help All Games
' Subject: Help <game>
Public Sub HelpEmail(ByVal strSubject As String, _
                    ByVal strFrom As String, _
                    ByVal strEMail)
    Dim varSubject As Variant
    Dim strMessage As String
    Dim objGame As Game
    Dim uHelp As udtHelp
    Dim strData As String
    
    strSubject = Trim(Replace(strSubject, vbTab, " "))
    While InStr(1, strSubject, "  ") > 0
        strSubject = Replace(strSubject, "  ", " ")
    Wend
    varSubject = Split(strSubject, " ")
    ReDim Preserve varSubject(3)
    
    ' Identify the Information being requested
    If varSubject(0) = "help" Then
        If varSubject(1) = "" Then
            uHelp = Help
        ElseIf varSubject(1) = "games" Then
            uHelp = HelpGames
        ElseIf varSubject(1) = "all" _
        And varSubject(2) = "games" Then
            uHelp = HelpAllGames
        ElseIf varSubject(2) = "" Then
            uHelp = HelpGame
        End If
    End If
    
    ' Send the requested information
    'Help
    If uHelp = Help Then
        strSubject = "RE: " & strSubject
        strMessage = Options.GetMessage("Help")
    
    'Help games
    ElseIf uHelp = HelpGames _
    Or uHelp = HelpAllGames Then
        strSubject = "RE: " & strSubject
        strMessage = "Listed below is that status of the currently active games." & vbNewLine & _
                    vbNewLine
        strMessage = strMessage & Pad("Game", 15) & "Status" & vbNewLine
        strMessage = strMessage & Pad("---------------", 15) & "-------------------------" & vbNewLine
        For Each objGame In GalaxyNG.Games
            If Not objGame.Template.Finished _
            Or uHelp = HelpAllGames Then
                strMessage = strMessage & Pad(objGame.GameName, 15) & _
                            objGame.Status & vbNewLine
            End If
        Next objGame
        
    'Help <game>
    ElseIf uHelp = HelpGame Then
        Set objGame = GalaxyNG.Games(varSubject(1))
        If objGame Is Nothing Then
            strSubject = "ERROR: " & strSubject
            strMessage = "Game " & varSubject(1) & "does not exist on this server." & vbNewLine
        
        Else
            strSubject = "RE: " & strSubject
            If objGame.Template.Finished Then
                strMessage = "Status: Finished" & vbNewLine
            ElseIf objGame.Started Then
                strMessage = "Status: Active" & vbNewLine
            ElseIf objGame.Created Then
                strMessage = "Status: Pending" & vbNewLine
            ElseIf objGame.Template.OpenForRegistrations Then
                strMessage = "Status: Open for registrations" & vbNewLine
            ElseIf objGame.Template.RegistrationOpen > Now Then
                strMessage = "Status: Active: Open " & _
                            Format(objGame.Template.RegistrationOpen, "d-mmm-yyyy") & vbNewLine
            Else
                strMessage = "Status: Under Development." & vbNewLine
            End If
            
            ' Get the game Description
            strData = objGame.Template.Description
            If strData <> "" Then
                strMessage = strMessage & String(60, "-") & vbNewLine
                strMessage = strMessage & strData & vbNewLine
            End If
            
            ' Get Score is exists
            strData = GetFile(Options.GalaxyNGNotices & objGame.GameName & ".score")
            If strData <> "" Then
                strMessage = strMessage & String(60, "-") & vbNewLine
                strMessage = strMessage & strData
            End If
            
            'Get the Game Details
            strData = objGame.Template.Details
            If strData <> "" Then
                strMessage = strMessage & String(60, "-") & vbNewLine
                strMessage = strMessage & objGame.Template.Details
            End If
        End If
    End If
    
    ' Format and send the e-mail
    strMessage = Options.GetMessage("Header") & _
                strMessage & _
                Options.GetMessage("Footer")
    
    Call SendEMail(strFrom, strSubject, strMessage)

End Sub

Private Function Pad(ByVal vData As Variant, ByVal lngPad As Long, Optional ByVal lngSpaces As Long = 2) As String
    Dim strData As String
    
    strData = CStr(vData)
    strData = Left(vData & String(lngPad, " "), lngPad) & String(lngSpaces, " ")
    Pad = strData
    
End Function

