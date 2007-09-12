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

    strEMail = GetEMail(strPath)
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
    Dim blnValid As Boolean
    Dim strMessage As String
    
    Set objGame = GalaxyNG.Games(strGame)
    If objGame Is Nothing Then
        strMessage = ""
    ElseIf objGame.Created Then
        strMessage = ""
    Else
        'if an existing registration then
            'Change the registration
            blnValid = True
        'elseif closed
        'elseif full
        'else
            'Process registration
            blnValid = True
            objGame.Template.Save
        'endif
    End If
    
    strMessage = Replace(strMessage, "[game]", strGame)
    ' Send Message
    Set objGame = Nothing
End Sub

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
    Dim b As Long
    
    varLines = Split(strEMail, vbCrLf)
    For i = LBound(varLines) To UBound(varLines)
        strLine = varLines(i)
        If blnBody Then
            b = b + 1
            varBody(b) = strLine
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
                b = -1
            End If
        End If
    Next i
End Sub

Private Function GetEMail(ByVal strPath As String) As String
    Dim intFN As Integer
    Dim strBuffer As String
    Dim lngLength As Long
    
    lngLength = FileLen(strPath)
    strBuffer = String(lngLength, " ")
    
    intFN = FreeFile
    Open strPath For Binary As #intFN
    Get intFN, , strBuffer
    Close intFN
    GetEMail = strBuffer
End Function

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
