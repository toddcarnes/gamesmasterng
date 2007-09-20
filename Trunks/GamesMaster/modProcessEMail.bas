Attribute VB_Name = "modProcessEMail"
Option Explicit
Option Compare Text

Public Sub ProcessEMails()
    Dim varEMails As Variant
    Dim i As Long
    
    varEMails = GetEMails
    If IsEmpty(varEMails) Then Exit Sub
    For i = 0 To UBound(varEMails)
        Call ProcessEMail(Inbox & "\" & varEMails(i))
    Next i
End Sub

Private Sub ProcessEMail(ByVal strPath As String)
    Dim strEMail As String
    Dim strFrom As String
    Dim strSubject As String
    Dim varBody As String
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
    Case "orders" Or "order"
        Call CheckOrders(strFrom, varBody)
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
    
    Set objGame = GalaxyNG.Games(strGame)
    If objGame Is Nothing Then
        strMessage = "No such Game Exists"
        blnValid = False
    ElseIf objGame.Created Then
        strMessage = "Game has already started"
        blnValid = False
    ElseIf Not objGame.Template.OpenForRegistrations Then
        strMessage = "Not is not open for registrations"
        blnValid = False
    Else
        strFrom = GetAddress(strFrom)
        Set objExisting = objGame.Template.Registrations(strFrom)
        If Not objExisting Is Nothing Then
            Set objRegistration = RegisterPlayer(varBody)
            blnValid = True
        ElseIf objGame.Template.Registrations.Count > objGame.Template.MaxPlayers Then
            strMessage = "Game is full"
            blnValid = False
        Else
            Set objRegistration = RegisterPlayer(varBody)
            objRegistration.EMail = strFrom
            blnValid = True
        End If
    End If
    
    If blnValid Then
        If objRegistration.HomeWorlds.Count > objGame.Template.MaxPlanets Then
            strMessage = "Too many Planets"
            blnValid = False
        ElseIf objRegistration.HomeWorlds.MaxSize > objGame.Template.MaxPlanetSize Then
            strMessage = "Planet too Large"
            blnValid = False
        ElseIf objRegistration.HomeWorlds.TotalSize > objGame.Template.TotalPlanetSize Then
            strMessage = "Total planet size invalid"
            blnValid = False
        End If
    End If
    If blnValid Then
        If objExisting Is Nothing Then
            objGame.Template.Registrations.Add objRegistration
            strMessage = "Registration accepted"
        Else
            Set objExisting.HomeWorlds = objRegistration.HomeWorlds
            strMessage = "Registration updated"
        End If
    End If
    
    ' Send Message
    
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
    Dim objHomeWorld As HomeWorld
    Dim objRegistration As Registration
    
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
                    Set objHomeWorld = New HomeWorld
                    objHomeWorld.Size = varFields(j)
                Next j
            ElseIf varFields(0) = "#racename" Then
                objRegistration.RaceName = varFields(1)
            End If
        End If
    Next i
    

End Function


Private Sub CheckOrders(ByVal strFrom As String, ByVal varBody As Variant)

End Sub

Private Sub RelayMessage(ByVal strTo As String, ByVal strFrom As String, ByVal varBody As Variant)

End Sub

Private Sub EMailReport(ByVal strFrom As String, ByVal varBody As Variant)

End Sub

Private Sub AnalyseEMail(ByVal strEMail As String, _
                        ByRef strFrom As String, _
                        ByVal strSubject As String, _
                        ByVal varBody As Variant)
    Dim varLines As Variant
    Dim strLine As String
    Dim strWord As String
    Dim blnBody As Boolean
    Dim strText As String
    
    Dim i As Long
    Dim j As Long
    Dim B As Long
    
    varLines = Split(strEMail, vbCrLf)
    For i = LBound(varLines) To UBound(varLines)
        strLine = varLines(i)
        If blnBody Then
            B = B + 1
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
                B = -1
            End If
        End If
    Next i
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

