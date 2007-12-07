Attribute VB_Name = "modReport"
Option Explicit
Option Compare Text

Public Sub SendReport(ByVal strFrom As String, ByVal strEMail As String)
    Dim lngStart As Long
    Dim lngEOL As Long
    Dim strHeader As String
    Dim varHeader As Variant
    Dim strGame As String
    Dim strRace As String
    Dim strPassword As String
    Dim lngTurn As Long
    Dim objGames As Games
    Dim objGame As Game
    Dim objRace As Race
    Dim strSubject As String
    Dim strMessage As String
    
    strSubject = "Major Problems Processing your Report email"
    
    ' extract just the Header Linet
    lngStart = InStr(1, strEMail, "#galaxy", vbTextCompare)
    If lngStart = 0 Then
        'Invalid EMail
        strMessage = Options.GetMessage("InvalidReportEMail", strEMail)
        GoTo Error
    End If
    
    lngEOL = InStr(lngStart, strEMail, vbCrLf)
    If lngEOL = 0 Then
        strHeader = Mid(strEMail, lngStart)
    Else
        strHeader = Mid(strEMail, lngStart, lngEOL - lngStart)
    End If
    
    'Reduce multiple spaces to single spaces
    strHeader = Replace(strHeader, vbTab, " ")
    While InStr(1, strHeader, "  ") > 0
        strHeader = Replace(strHeader, "  ", " ")
    Wend
    
    'Split the header by arguements
    varHeader = Split(strHeader, " ")
    If UBound(varHeader) < 4 Then
        'Invalid Header
        strMessage = Options.GetMessage("InvalidReportHeader", _
                                        "Insufficient parameters Supplied in the #galaxy line", _
                                        QuoteText(strHeader))
        GoTo Error
    End If
    strGame = varHeader(1)
    strRace = varHeader(2)
    strPassword = varHeader(3)
    lngTurn = varHeader(4)
    
    'Validate the Game
    Set objGames = New Games
    objGames.Refresh
    Set objGame = GalaxyNG.Games(strGame)
    
    If objGame Is Nothing Then
        'Invalid Header
        strMessage = Options.GetMessage("InvalidReportHeader", _
                "An unknown game was specified.", _
                QuoteText(strHeader))
        GoTo Error
    End If
    
    objGame.Refresh
    Set objRace = objGame.Races(strRace)
    If objRace Is Nothing Then
        'Invalid Header
        strMessage = Options.GetMessage("InvalidReportHeader", _
                "An unknown race was specified.", _
                QuoteText(strHeader))
        GoTo Error
    End If
    
    If objRace.Password <> strPassword Then
        'Invalid Header
        strMessage = Options.GetMessage("InvalidReportHeader", _
                "An invalid password was specified for the selected race.", _
                QuoteText(strHeader))
        GoTo Error
    End If
    
    If lngTurn > objGame.Turn Then
        'Invalid Header
        strMessage = Options.GetMessage("InvalidReportHeader", _
                "The turn number is for a turn that has not been processed.", _
                QuoteText(strHeader))
        GoTo Error
    End If
    
    ' Get the Report file to send
    strSubject = "[GNG] " & strGame & " turn " & CStr(lngTurn) & _
                    " text report for " & strRace
    Call RunGalaxyNG("-report " & strGame & " " & strRace & " " & CStr(lngTurn) & " >" & Options.ReportFileName)
    strMessage = GetFile(Options.GalaxyNGHome & Options.ReportFileName)
    Kill Options.GalaxyNGHome & Options.ReportFileName

    GoTo Send

Error:
    strMessage = Options.GetMessage("Header") & _
                    strMessage & _
                    Options.GetMessage("Footer")
Send:
    Call SendEMail(strFrom, strSubject, strMessage)
End Sub



