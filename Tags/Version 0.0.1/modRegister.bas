Attribute VB_Name = "modRegister"
Option Explicit
Option Compare Text

Public Sub JoinGame(ByVal strGame As String, ByVal strFrom As String, ByVal varBody As Variant)
    Dim objGame As Game
    Dim objExisting As Registration
    Dim objRegistration As Registration
    Dim blnValid As Boolean
    Dim strMessage As String
    Dim strAddress As String
    
    Set objGame = GalaxyNG.Games(strGame)
    If objGame Is Nothing Then
        strMessage = Options.GetMessage("NoGame", strGame)
        blnValid = False
    ElseIf objGame.Created Then
        strMessage = Options.GetMessage("GameStarted", strGame)
        blnValid = False
    ElseIf Not objGame.Template.OpenForRegistrations Then
        strMessage = Options.GetMessage("NotOpen", strGame)
        blnValid = False
    Else
        strAddress = GetAddress(strFrom)
        Set objExisting = objGame.Template.Registrations(strAddress)
        If Not objExisting Is Nothing Then
            Set objRegistration = RegisterPlayer(varBody)
            blnValid = True
        ElseIf objGame.Template.Registrations.Count >= objGame.Template.MaxPlayers Then
            strMessage = Options.GetMessage("GameFull", strGame, objGame.Template.MaxPlayers)
            blnValid = False
        Else
            Set objRegistration = RegisterPlayer(varBody)
            objRegistration.EMail = GetAddress(strFrom)
            blnValid = True
        End If
    End If
    
    If blnValid Then
        If objRegistration.HomeWorlds.Count = 0 Then
            Set objRegistration.HomeWorlds = objGame.Template.DefaultHomeWorlds
        ElseIf objRegistration.HomeWorlds.Count > objGame.Template.MaxPlanets Then
            strMessage = Options.GetMessage("TooManyPlanets", strGame, _
                        objRegistration.HomeWorlds.Count, _
                        objGame.Template.MaxPlanets)
            Set objRegistration.HomeWorlds = objGame.Template.DefaultHomeWorlds
            blnValid = True
        ElseIf objRegistration.HomeWorlds.MaxSize > objGame.Template.MaxPlanetSize Then
            strMessage = Options.GetMessage("PlanetTooLarge", strGame, _
            objRegistration.HomeWorlds.MaxSize, objGame.Template.MaxPlanetSize)
            Set objRegistration.HomeWorlds = objGame.Template.DefaultHomeWorlds
            blnValid = True
        ElseIf objRegistration.HomeWorlds.TotalSize <> objGame.Template.TotalPlanetSize Then
            strMessage = Options.GetMessage("TotalPlanets", strGame, _
            objRegistration.HomeWorlds.TotalSize, _
            objGame.Template.TotalPlanetSize)
            Set objRegistration.HomeWorlds = objGame.Template.DefaultHomeWorlds
            blnValid = True
        End If
    End If
    
    If blnValid Then
        If objExisting Is Nothing Then
            objGame.Template.Registrations.Add objRegistration
            strMessage = strMessage & vbNewLine & _
                            Options.GetMessage("RegistrationAccepted", strGame, objRegistration.HomeWorlds.Text)
        Else
            Set objExisting.HomeWorlds = objRegistration.HomeWorlds
            strMessage = strMessage & vbNewLine & _
                            Options.GetMessage("RegistrationUpdated", strGame, objRegistration.HomeWorlds.Text)
        End If
    End If
    
    ' Send Message
    strMessage = Options.GetMessage("Header") & _
                strMessage & _
                Options.GetMessage("Footer", Options.ServerName)
    Call SendEMail(strFrom, "re: Join " & strGame, strMessage)
    
    If blnValid Then
        objGame.Template.Save
    End If

    ' Clean up
    Set objExisting = Nothing
    Set objRegistration = Nothing
    Set objGame = Nothing
End Sub

Public Function RegisterPlayer(ByVal varBody As Variant) As Registration
    Dim i As Long
    Dim j As Long
    Dim strLine As String
    Dim varFields As Variant
    Dim objHomeworld As HomeWorld
    Dim objRegistration As Registration
    
    Set objRegistration = New Registration
    For i = LBound(varBody) To UBound(varBody)
        strLine = Trim(varBody(i))
        If strLine = "" Then
            ' ignore
        Else
            While InStr(1, strLine, "  ") > 0
                strLine = Replace(strLine, "  ", " ")
            Wend
            varFields = Split(strLine, " ")
            If varFields(0) = "#planets" Then
                Set objRegistration.HomeWorlds = New HomeWorlds
                For j = 1 To UBound(varFields)
                    Set objHomeworld = New HomeWorld
                    objHomeworld.Size = varFields(j)
                    objRegistration.HomeWorlds.Add objHomeworld
                Next j
            ElseIf varFields(0) = "#racename" Then
                objRegistration.RaceName = varFields(1)
            End If
        End If
    Next i
    Set RegisterPlayer = objRegistration

End Function


