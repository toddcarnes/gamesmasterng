Attribute VB_Name = "modProcessEMail"
Option Explicit
Option Compare Text

Public Sub ProcessEMails()
    Dim varEMails As Variant
    Dim i As Long
    
    varEMails = GetEMails
    If IsEmpty(varEMails) Then Exit Sub
    For i = 0 To UBound(varEMails)
        Call ProcessEMail(Inbox & varEMails(i))
    Next i
End Sub

Private Sub ProcessEMail(ByVal strPath As String)
    Dim strEMail As String
    Dim strFrom As String
    Dim strSubject As String
    Dim varBody As Variant
    Dim varSubject As Variant

    strEMail = GetFile(strPath)
    Call AnalyseEMail(strEMail, strFrom, strSubject, varBody)
    
    While InStr(1, strSubject, "  ") > 0
        strSubject = Replace(strSubject, "  ", " ")
    Wend
    
    varSubject = Split(strSubject, " ")
    Select Case varSubject(0)
    Case "join"
        Call JoinGame(varSubject(1), strFrom, varBody)
    Case "orders", "order"
        Call CheckOrders(strFrom, strEMail)
    Case "relay"
        Call RelayMessage(varSubject(1), strFrom, varBody)
    Case "report"
        Call EMailReport(strFrom, varBody)
    End Select
    
    Name strPath As strPath & ".sav"

End Sub

Private Sub JoinGame(ByVal strGame As String, ByVal strFrom As String, ByVal varBody As Variant)
    Dim objGame As Game
    Dim objExisting As Registration
    Dim objRegistration As Registration
    Dim blnValid As Boolean
    Dim strMessage As String
    Dim strAddress As String
    
    Set objGame = GalaxyNG.Games(strGame)
    If objGame Is Nothing Then
        strMessage = GetMessage("NoGame", strGame)
        blnValid = False
    ElseIf objGame.Created Then
        strMessage = GetMessage("GameStarted", strGame)
        blnValid = False
    ElseIf Not objGame.Template.OpenForRegistrations Then
        strMessage = GetMessage("NotOpen", strGame)
        blnValid = False
    Else
        strAddress = GetAddress(strFrom)
        Set objExisting = objGame.Template.Registrations(strAddress)
        If Not objExisting Is Nothing Then
            Set objRegistration = RegisterPlayer(varBody)
            blnValid = True
        ElseIf objGame.Template.Registrations.Count >= objGame.Template.MaxPlayers Then
            strMessage = GetMessage("GameFull", strGame, objGame.Template.MaxPlayers)
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
            strMessage = GetMessage("TooManyPlanets", strGame, _
                        objRegistration.HomeWorlds.Count, _
                        objGame.Template.MaxPlanets)
            Set objRegistration.HomeWorlds = objGame.Template.DefaultHomeWorlds
            blnValid = True
        ElseIf objRegistration.HomeWorlds.MaxSize > objGame.Template.MaxPlanetSize Then
            strMessage = GetMessage("PlanetTooLarge", strGame, _
            objRegistration.HomeWorlds.MaxSize, objGame.Template.MaxPlanetSize)
            Set objRegistration.HomeWorlds = objGame.Template.DefaultHomeWorlds
            blnValid = True
        ElseIf objRegistration.HomeWorlds.TotalSize <> objGame.Template.TotalPlanetSize Then
            strMessage = GetMessage("TotalPlanets", strGame, _
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
                            GetMessage("RegistrationAccepted", strGame, objRegistration.HomeWorlds.Text)
        Else
            Set objExisting.HomeWorlds = objRegistration.HomeWorlds
            strMessage = strMessage & vbNewLine & _
                            GetMessage("RegistrationUpdated", strGame, objRegistration.HomeWorlds.Text)
        End If
    End If
    
    ' Send Message
    strMessage = GetMessage("Header") & _
                strMessage & _
                GetMessage("Footer", ServerName)
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


Private Sub CheckOrders(ByVal strFrom As String, ByVal strEMail As String)
    Dim lngStart As Long
    Dim lngEnd As Long
    Dim lngEOL As Long
    Dim strOrders As String
    Dim strHeader As String
    Dim varHeader As Variant
    Dim strGame As String
    Dim strRace As String
    Dim strPassword As String
    Dim lngTurn As Long
    Dim blnFinalOrders As Boolean
    Dim objGame As Game
    Dim objRace As Race
    Dim strSubject As String
    Dim strMessage As String
    Dim strFileName As String
    Dim strFileName1 As String
    
    strSubject = "Major Problems Processing your orders email"
    
    ' extract just the orders
    lngStart = InStr(1, strEMail, "#galaxy", vbTextCompare)
    If lngStart = 0 Then
        'Invalid EMail
        strMessage = GetMessage("InvalidOrdersEMail", strEMail)
        GoTo Error
    End If
    
    lngEnd = InStr(lngStart, strEMail, "#end", vbTextCompare)
    If lngEnd = 0 Then
        'Invalid EMail
        strMessage = GetMessage("InvalidOrdersEMail", strEMail)
        GoTo Error
    End If
    lngEOL = InStr(lngEnd, strEMail, vbCrLf)
    If lngEOL = 0 Then
        strOrders = Mid(strEMail, lngStart, lngEnd + 3 - lngStart)
    Else
        strOrders = Mid(strEMail, lngStart, lngEOL + 2 - lngStart)
    End If
    
    ' Extract the header
    lngEOL = InStr(1, strOrders, vbCrLf)
    If lngEOL = 0 Then
        strHeader = strOrders
    Else
        strHeader = Left(strOrders, lngEOL - 1)
    End If
    'Reduce multiple spaces to single spaces
    strHeader = Replace(strHeader, vbTab, " ")
    While InStr(1, strHeader, "  ") > 0
        strHeader = Replace(strHeader, "  ", " ")
    Wend
    
    strSubject = "Major Problems Processing your orders"
    'Split the header by arguements
    varHeader = Split(strHeader, " ")
    If UBound(varHeader) < 4 Then
        'Invalid Header
        strMessage = GetMessage("InvalidOrdersHeader", strEMail)
        GoTo Error
    End If
    If varHeader(0) <> "#galaxy" Then
        'Invalid Header
        strMessage = GetMessage("InvalidOrdersHeader", strEMail)
        GoTo Error
    End If
    strGame = varHeader(1)
    strRace = varHeader(2)
    strPassword = varHeader(3)
    lngTurn = varHeader(4)
    If UBound(varHeader) = 5 Then
        If varHeader(5) = "finalorders" Then
            blnFinalOrders = True
        Else
            'Invalid Header
            strMessage = GetMessage("InvalidOrdersHeader", strEMail)
            GoTo Error
        End If
    End If
    If UBound(varHeader) > 5 Then
        'Invalid Header
        strMessage = GetMessage("InvalidOrdersHeader", _
                "An invalid number of header parameters were specified", _
                strEMail)
        GoTo Error
    End If
    
    'Validate the Game
    Set objGame = GalaxyNG.Games(strGame)
    
    If objGame Is Nothing Then
        'Invalid Header
        strMessage = GetMessage("InvalidOrdersHeader", _
                "An unknown game was specified.", _
                strEMail)
        GoTo Error
    End If
    
    objGame.Refresh
    Set objRace = objGame.Races(strRace)
    If objRace Is Nothing Then
        'Invalid Header
        strMessage = GetMessage("InvalidOrdersHeader", _
                "An unknown race was specified.", _
                strEMail)
        GoTo Error
    End If
    
    If objRace.Password <> strPassword Then
        'Invalid Header
        strMessage = GetMessage("InvalidOrdersHeader", _
                "An invalid password was specified for the selected race.", _
                strEMail)
        GoTo Error
    End If
    
    If lngTurn < objGame.NextTurn Then
        'Invalid Header
        strMessage = GetMessage("InvalidOrdersHeader", _
                "The turn number is for a turn that has already been processed.", _
                strEMail)
        GoTo Error
    End If
    
    If lngTurn > objGame.NextTurn Then
        strSubject = "[GNG] " & objGame.GameName & _
                    " turn " & CStr(lngTurn) & _
                    " orders received for " & objRace.RaceName
        strMessage = GetMessage("Header", ServerName) & _
                    GetMessage("FutureOrders", lngTurn, MarkText(strOrders)) & _
                    GetMessage("Footer", ServerName)
        
    Else 'if lngTurn = objGame.NextTurn Then
        strSubject = "[GNG] " & objGame.GameName & _
                    " turn " & CStr(lngTurn) & _
                    " text forecast for " & objRace.RaceName
        ' check it
        Call SaveFile(GalaxyNGHome & gcOrdersFileName, strOrders)
        Call RunGalaxyNG("-check " & strGame & " " & strRace & "<" & gcOrdersFileName & " >" & gcForecastFileName)
        strMessage = GetFile(GalaxyNGHome & gcForecastFileName)
        Kill GalaxyNGHome & gcOrdersFileName
        Kill GalaxyNGHome & gcForecastFileName
    End If

    'File the orders
    strFileName = GalaxyNGOrders & strGame & "\" & strRace & "." & CStr(lngTurn)
    strFileName1 = GalaxyNGOrders & strGame & "\" & strRace & "_final" & "." & CStr(lngTurn)
    
    If Dir(strFileName) <> "" Then Kill strFileName
    If Dir(strFileName1) <> "" Then Kill strFileName1
    If blnFinalOrders Then
        Call SaveFile(strFileName1, strOrders)
    Else
        Call SaveFile(strFileName, strOrders)
    End If

    GoTo Send

Error:
    strMessage = GetMessage("Header", ServerName) & _
                    strMessage & _
                    GetMessage("Footer", ServerName)
Send:
    Call SendEMail(strFrom, strSubject, strMessage)
End Sub

Private Sub RelayMessage(ByVal strTo As String, ByVal strFrom As String, ByVal varBody As Variant)

End Sub

Private Sub EMailReport(ByVal strFrom As String, ByVal varBody As Variant)

End Sub

Private Sub AnalyseEMail(ByVal strEMail As String, _
                        ByRef strFrom As String, _
                        ByRef strSubject As String, _
                        ByRef varBody As Variant)
    Dim varLines As Variant
    Dim strLine As String
    Dim strWord As String
    Dim blnBody As Boolean
    Dim strText As String
    
    Dim i As Long
    Dim j As Long
    Dim B As Long
    
    B = -1
    varLines = Split(strEMail, vbCrLf)
    For i = LBound(varLines) To UBound(varLines)
        strLine = varLines(i)
        If blnBody Then
            B = B + 1
            If B > UBound(varBody) Then
                ReDim Preserve varBody(B + 100)
            End If
            varBody(B) = strLine
        Else
            j = InStr(1, strLine, " ")
            If j > 0 Then
                strWord = Left(strLine, j - 1)
                strText = Mid(strLine, j + 1)
                Select Case strWord
                Case "from:"
                    strFrom = strText
                Case "subject:"
                    strSubject = strText
                End Select
            ElseIf strLine = "" Then
                blnBody = True
                ReDim varBody(99)
            End If
        End If
    Next i
    
    If B >= 0 Then
        ReDim Preserve varBody(B)
    End If
End Sub

Private Function GetEMails() As Variant
    Dim varFiles As Variant
    Dim i As Long
    Dim strFile As String
    ReDim varFiles(100) As Variant
    i = -1
    
    strFile = Dir(Inbox & "\*.txt")
    While strFile <> ""
        i = i + 1
        If i > UBound(varFiles) Then
            ReDim Preserve varFiles(i + 99)
        End If
        varFiles(i) = strFile
        strFile = Dir()
    Wend
    If i = -1 Then
        GetEMails = Empty
    Else
        ReDim Preserve varFiles(i)
        GetEMails = varFiles
    End If
    
End Function

Public Function SendEMail(ByVal strTo As String, ByVal strSubject As String, ByVal strBody As String)
    Dim intFN As Integer
    Dim i As Long
    Dim strFileName As String
    Do
        strFileName = Outbox & Format(Now, "yyyymmddhhnnss") & "_" & Format(i, "0") & ".txt"
        If Dir(strFileName) = "" Then Exit Do
        i = i + 1
    Loop
    intFN = FreeFile
    Open strFileName For Output As #intFN
    Print #intFN, "To: " & strTo
    Print #intFN, "From: " & SMTPFromAddress
    Print #intFN, "Subject: " & strSubject
    Print #intFN, ""
    Print #intFN, strBody
    Close #intFN
    
End Function

Public Sub SendReports(ByVal strGame As String)
    Dim objGame As Game
    Dim objRace As Race
    Dim strTurn As String
    Dim strBody As String
    
    Set objGame = GalaxyNG.Games(strGame)
    strTurn = objGame.Turn
    
    For Each objRace In objGame.Races
        If Not objRace.Flag(R_DEAD) Then
            strBody = GetFile(GalaxyNGReports & objGame.GameName & "\" & _
                        objRace.RaceName & "_" & strTurn & ".txt")
            Call SendEMail(objRace.EMail, _
                    "[GNG] " & objGame.GameName & _
                    " turn " & strTurn & _
                    " text report for " & objRace.RaceName, _
                    strBody)
        End If
    Next objRace
End Sub
