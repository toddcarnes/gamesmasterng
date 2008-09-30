Attribute VB_Name = "modProcessEMail"
Option Explicit
Option Compare Text

Public Function ProcessEMails() As Boolean
    Dim varEMails As Variant
    Dim i As Long
    
    ProcessEMails = False
    varEMails = GetEMails
    If IsEmpty(varEMails) Then Exit Function
    For i = 0 To UBound(varEMails)
        Call ProcessEMail(Options.Inbox & varEMails(i))
    Next i
    ProcessEMails = True
End Function

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
            Call RelayMessage(varSubject(1), strFrom, strEMail)
        Case "report"
            Call SendReport(strFrom, strEMail)
        Case "help"
            Call HelpEmail(strSubject, strFrom, strEMail)
        Case "re:"
            ReDim Preserve varSubject(5)
            If varSubject(1) = "[gng]" _
            And varSubject(3) = "message" _
            And varSubject(4) = "relay" Then
                Call RelayMessage(varSubject(5), strFrom, strEMail)
            End If
        End Select
        If Options.SaveEMail Then
            Name strPath As strPath & ".sav"
        Else
            Kill strPath
        End If
    Else
        Name strPath As strPath & ".err"
    End If

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
    
    strFile = Dir(Options.Inbox & "\*.eml")
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
    Dim objTimeZone As CTimeZone
    Dim strTime As String
    
    Set objTimeZone = New CTimeZone
    strTime = objTimeZone.TimeEMail
    Set objTimeZone = Nothing
    
    Do
        strFileName = Options.Outbox & Format(Now, "yyyymmddhhnnss") & "_" & Format(i, "0") & ".eml"
        If Dir(strFileName) = "" Then Exit Do
        i = i + 1
    Loop
    intFN = FreeFile
 
    Open strFileName For Output As #intFN
    Print #intFN, "To: " & strTo
    Print #intFN, "From: " & Options.SMTPFromAddress
    Print #intFN, "Subject: " & strSubject
    Print #intFN, "Date: " & strTime
    Print #intFN, ""
    Print #intFN, strBody
    Close #intFN
    
End Function

Public Function SendNewEMail(ByVal strBody As String)
    Dim intFN As Integer
    Dim i As Long
    Dim strFileName As String
    Dim objTimeZone As CTimeZone
    
    Do
        strFileName = Options.Outbox & Format(Now, "yyyymmddhhnnss") & "_" & Format(i, "0") & ".eml"
        If Dir(strFileName) = "" Then Exit Do
        i = i + 1
    Loop
    intFN = FreeFile
 
    Open strFileName For Output As #intFN
    Print #intFN, strBody;
    Close #intFN
    
End Function

Public Sub SendReports(ByVal strGame As String)
    Dim objGame As Game
    Dim objRace As Race
    Dim strTurn As String
    Dim strBody As String
    Dim strFileName As String
    Dim objNE As NewEMail
    Dim objA As Attachment
    Dim objZip As Zip
    Dim strMessage As String
    Dim blnCompress As Boolean
    Dim blnAttach As Boolean
    Dim blnText As Boolean
    
    GalaxyNG.Games.Refresh
    Set objGame = GalaxyNG.Games(strGame)
    objGame.Refresh
    strTurn = objGame.Turn
    
    'Send the Games Master Report
    strFileName = Options.GamesMasterReport(strGame, strTurn)
    If Dir(strFileName) <> "" Then
        strBody = GetFile(strFileName)
        Call SendEMail(Options.GamesMasterEMail, _
                "[GNG] " & objGame.GameName & " turn " & strTurn & _
                " text report for GM", _
                strBody)
    End If
    
    For Each objRace In objGame.Races
        If Not objRace.flag(R_DEAD) Then
            'Choose which method to send the report/s
            blnCompress = False
            blnAttach = False
            blnText = False
            If objRace.flag(R_COMPRESS) Then
                blnCompress = True
            ElseIf Options.AttachReports Then
                blnAttach = True
            Else
                blnText = True
            End If
            
            ' Send the reports compressed as a zip file
            If blnCompress Then
                Set objNE = New NewEMail
                objNE.ToAddress = objRace.EMail
                objNE.FromAddress = Options.SMTPFromAddress
                objNE.DateSent = Now
                objNE.Subject = "[GNG] " & objGame.GameName & " turn " & strTurn & _
                        " text report for " & objRace.RaceName
                
                'EMail Body
                Set objA = New Attachment
                strMessage = Options.GetMessage("Header") & _
                            Options.GetMessage("GamesMasterMessage") & _
                            objGame.Template.Message & _
                            Options.GetMessage("Footer")
                Call objA.Store(strMessage, uefText)
                objNE.Attachments.Add objA
                
                'EMail Zip File Attachment
                Set objZip = New Zip
                objZip.RootDirectory = Options.GalaxyNGReports & strGame & "\"
                ChDir objZip.RootDirectory
                
                objZip.ZipFileName = Options.GalaxyNGReports & strGame & "\" & objRace.RaceName & "_" & strTurn & ".zip"
                If Dir(objZip.ZipFileName) <> "" Then
                    Kill objZip.ZipFileName
                End If
                
                'Text Report
                strFileName = Options.RaceReport(strGame, objRace.RaceName, strTurn)
                If Dir(strFileName) <> "" Then
                    objZip.AddFile GetFullFileName(strFileName)
                End If
                
                'Machine Report
                strFileName = Options.RaceMachineReport(strGame, objRace.RaceName, strTurn)
                If Dir(strFileName) <> "" Then
                    objZip.AddFile GetFullFileName(strFileName)
                End If
                objZip.MakeZipFile
                ChDir App.Path
                
                ' Attach the Zip file.
                strFileName = objZip.ZipFileName
                If Dir(strFileName) <> "" Then
                    strBody = GetFile(strFileName)
                    Set objA = New Attachment
                    Call objA.Store(strBody, uefBinary, strFileName)
                    objNE.Attachments.Add objA
                    Call SendNewEMail(objNE.EMailData)
                Else ' Problem creating zip file so send as attachments
                    Call LogError(-10001, "Zip File failed to be created", "ZIP", _
                    "modProcessEMail", "SendReports", _
                    "    Game: " & strGame & vbNewLine & _
                    "    Race: " & objRace.RaceName & vbNewLine & _
                    "    File: " & strFileName & vbNewLine & _
                    "     Msg: " & objZip.GetLastMessage)
                    Call SendEMail(Options.GamesMasterEMail, _
                            "[GNG ERROR] " & objGame.GameName & " turn " & strTurn & _
                            " Race " & objRace.RaceName, _
                            "A zip file failed to generate." & vbNewLine & _
                            "Reports were sent to the player as attachments." & vbNewLine & _
                            "    Game: " & strGame & vbNewLine & _
                            "    Race: " & objRace.RaceName & vbNewLine & _
                            "    File: " & strFileName & vbNewLine & _
                            "     Msg: " & objZip.GetLastMessage)
                    blnAttach = True
                End If
            End If
            
            ' Send the reports as attachments
            If blnAttach Then
                Set objNE = New NewEMail
                objNE.ToAddress = objRace.EMail
                objNE.FromAddress = Options.SMTPFromAddress
                objNE.DateSent = Now
                objNE.Subject = "[GNG] " & objGame.GameName & " turn " & strTurn & _
                        " text report for " & objRace.RaceName
                
                'EMail Body
                Set objA = New Attachment
                strMessage = Options.GetMessage("Header")
                If blnCompress Then 'The compress failed
                    strMessage = strMessage & _
                        "**** Trouble was encountered creating the ZIP file requested. " & vbNewLine & _
                        "     The Games Master has been informed of the problem." & vbNewLine & _
                        "**** Your report is attached as a text file." & vbNewLine & vbNewLine
                End If
                strMessage = strMessage & Options.GetMessage("GamesMasterMessage") & _
                            objGame.Template.Message & _
                            Options.GetMessage("Footer")
                Call objA.Store(strMessage, uefText)
                objNE.Attachments.Add objA
                
                'Text Report
                strFileName = Options.RaceReport(strGame, objRace.RaceName, strTurn)
                If Dir(strFileName) <> "" Then
                    Set objA = New Attachment
                    strBody = GetFile(strFileName)
                    Call objA.Store(strBody, uefText, strFileName)
                    objNE.Attachments.Add objA
                End If
                
                'Machine Report
                strFileName = Options.RaceMachineReport(strGame, objRace.RaceName, strTurn)
                If Dir(strFileName) <> "" Then
                    Set objA = New Attachment
                    strBody = GetFile(strFileName)
                    Call objA.Store(strBody, uefText, strFileName)
                    objNE.Attachments.Add objA
                End If
                
                Call SendNewEMail(objNE.EMailData)
            End If
            
            ' Send the reports a pure text e-mails.
            If blnText Then ' EMail reports seperately
                'Text Report
                strFileName = Options.RaceReport(strGame, objRace.RaceName, strTurn)
                If Dir(strFileName) <> "" Then
                    strBody = GetFile(strFileName)
                    Call SendEMail(objRace.EMail, _
                            "[GNG] " & objGame.GameName & " turn " & strTurn & _
                            " text report for " & objRace.RaceName, _
                            strBody)
                End If
                
                'Machine Report
                strFileName = Options.RaceMachineReport(strGame, objRace.RaceName, strTurn)
                If Dir(strFileName) <> "" Then
                    strBody = GetFile(strFileName)
                    Call SendEMail(objRace.EMail, _
                            "[GNG] " & objGame.GameName & " turn " & strTurn & _
                            " machine report for " & objRace.RaceName, _
                            strBody)
                End If
            End If
        End If
    Next objRace
End Sub
