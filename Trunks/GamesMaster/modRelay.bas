Attribute VB_Name = "modRelay"
Option Explicit
Option Compare Text

Public Sub RelayMessage(ByVal strTo As String, ByVal strFrom As String, ByVal strEMail As String)
    Dim lngStart As Long
    Dim lngEnd As Long
    Dim lngEOL As Long
    Dim strOrders As String
    Dim strHeader As String
    Dim varHeader As Variant
    Dim strGame As String
    Dim strRace As String
    Dim strPassword As String
    Dim objGame As Game
    Dim objRace As Race
    Dim objToRace As Race
    Dim strSubject As String
    Dim strMessage As String
    Dim strFileName As String
    Dim strFileName1 As String
    
    strSubject = "Major Problems Processing your orders email"
    
    ' extract just the orders
    lngStart = InStr(1, strEMail, "#galaxy", vbTextCompare)
    If lngStart = 0 Then
        'Invalid EMail
        strMessage = Options.GetMessage("InvalidRelayEMail", strEMail)
        GoTo Error
    End If
    
    lngEnd = InStr(lngStart, strEMail, "#end", vbTextCompare)
    If lngEnd = 0 Then
        'Invalid EMail
        strMessage = Options.GetMessage("InvalidRelayEMail", strEMail)
        GoTo Error
    End If
    strOrders = Mid(strEMail, lngStart, lngEnd - lngStart)
    
    ' Extract the header
    lngEOL = InStr(1, strOrders, vbCrLf)
    If lngEOL = 0 Then
        strHeader = strOrders
        strOrders = ""
    Else
        strHeader = Left(strOrders, lngEOL - 1)
        strOrders = Mid(strOrders, lngEOL + 2)
    End If
    'Reduce multiple spaces to single spaces
    strHeader = Replace(strHeader, vbTab, " ")
    While InStr(1, strHeader, "  ") > 0
        strHeader = Replace(strHeader, "  ", " ")
    Wend
    
    strSubject = "Major Problems Processing your orders"
    'Split the header by arguements
    varHeader = Split(strHeader, " ")
    If UBound(varHeader) < 3 Then
        'Invalid Header
        strMessage = Options.GetMessage("InvalidRelayHeader", strEMail)
        GoTo Error
    End If
    If varHeader(0) <> "#galaxy" Then
        'Invalid Header
        strMessage = Options.GetMessage("InvalidRelayHeader", strEMail)
        GoTo Error
    End If
    strGame = varHeader(1)
    strRace = varHeader(2)
    strPassword = varHeader(3)
    
    'Validate the Game
    Set objGame = GalaxyNG.Games(strGame)
    
    If objGame Is Nothing Then
        'Invalid Header
        strMessage = Options.GetMessage("InvalidRelayHeader", _
                "An unknown game was specified.", _
                strEMail)
        GoTo Error
    End If
    
    objGame.Refresh
    Set objRace = objGame.Races(strRace)
    If objRace Is Nothing Then
        'Invalid Header
        strMessage = Options.GetMessage("InvalidRelayHeader", _
                "An unknown race was specified.", _
                strEMail)
        GoTo Error
    End If
    
    If objRace.Password <> strPassword Then
        'Invalid Header
        strMessage = Options.GetMessage("InvalidRelayHeader", _
                "An invalid password was specified for the selected race.", _
                strEMail)
        GoTo Error
    End If
    
    Set objToRace = objGame.Races(strTo)
    If objToRace Is Nothing And strTo <> "gm" And strTo <> "gamesmaster" And strTo <> strGame Then
        'Invalid race
        strMessage = Options.GetMessage("InvalidRelayHeader", _
                "An invalid race name was specified to receive the message.", _
                strEMail)
        GoTo Error
    End If
    
    ' Send message to all races
    If strTo = strGame Then
        For Each objToRace In objGame.Races
            If Not objToRace.Flag(R_DEAD) Then
                strSubject = "[GNG]" & strGame & " relay to " & objToRace.RaceName & " from " & strRace
                Call SendEMail(objToRace.EMail, strSubject, strOrders)
            End If
        Next objToRace
        strSubject = "[GNG]" & strGame & " relay to GM from " & strRace
        Call SendEMail(Options.GamesMasterEMail, strSubject, strOrders)
    Else
        strSubject = "[GNG]" & strGame & " message relay " & strRace
        If objRace Is Nothing Then
            strSubject = "[GNG]" & strGame & " relay to GamesMaster from " & strRace
            Call SendEMail(Options.GamesMasterEMail, strSubject, strOrders)
        Else
            strSubject = "[GNG]" & strGame & " relay to " & objToRace.RaceName & " from " & strRace
            Call SendEMail(objToRace.EMail, strSubject, strOrders)
        End If
    End If
    
    strSubject = "[GNG]" & strGame & " relay sent to " & strTo
    strMessage = Options.GetMessage("RelaySent", strTo)


Error:
    strMessage = Options.GetMessage("Header") & _
                    strMessage & _
                    Options.GetMessage("Footer")
Send:
    Call SendEMail(strFrom, strSubject, strMessage)

End Sub


