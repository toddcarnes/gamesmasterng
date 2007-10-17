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
    If UBound(varSubject) >= 0 Then
        Select Case varSubject(0)
        Case "join"
            Call JoinGame(varSubject(1), strFrom, varBody)
        Case "orders", "order"
            Call CheckOrders(strFrom, strEMail)
        Case "relay"
            Call RelayMessage(varSubject(1), strFrom, varBody)
        Case "report"
            Call SendReport(strFrom, strEMail)
        End Select
        Name strPath As strPath & ".sav"
    Else
        Name strPath As strPath & ".err"
    End If

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
    Dim strFileName As String
    
    GalaxyNG.Games.Refresh
    Set objGame = GalaxyNG.Games(strGame)
    objGame.Refresh
    strTurn = objGame.Turn
    
    'Send the Games Master Report
    strFileName = GamesMasterReport(strGame, strTurn)
    If Dir(strFileName) <> "" Then
        strBody = GetFile(strFileName)
        Call SendEMail(GamesMasterEMail, _
                "[GNG] " & objGame.GameName & " turn " & strTurn & _
                " text report for GM", _
                strBody)
    End If
    
    For Each objRace In objGame.Races
        If Not objRace.Flag(R_DEAD) Then
            'Get the Text Report
            strFileName = RaceReport(strGame, objRace.RaceName, strTurn)
            strBody = GetFile(strFileName)
            Call SendEMail(objRace.EMail, _
                    "[GNG] " & objGame.GameName & " turn " & strTurn & _
                    " text report for " & objRace.RaceName, _
                    strBody)
            
            ' Check for the machinery Report
            strFileName = RaceMachineReport(strGame, objRace.RaceName, strTurn)
            If Dir(strFileName) <> "" Then
                strBody = GetFile(strFileName)
                Call SendEMail(objRace.EMail, _
                        "[GNG] " & objGame.GameName & " turn " & strTurn & _
                        " machine report for " & objRace.RaceName, _
                        strBody)
            End If
            
        End If
    Next objRace
End Sub
