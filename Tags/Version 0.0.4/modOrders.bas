Attribute VB_Name = "modOrders"
'********************************************************
'   Copyright 2007,2008 Ian Evans.                      *
'   This program is distributed under the terms of the  *
'       GNU General Public License.                     *
'********************************************************
Option Explicit
Option Compare Text

Public Sub CheckOrders(ByVal strFrom As String, ByVal strEMail As String)
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
    Dim strFinalOrders As String
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
        strMessage = Options.GetMessage("InvalidOrdersEMail", QuoteText(strEMail))
        GoTo Error
    End If
    
    lngEnd = InStr(lngStart, strEMail, "#end", vbTextCompare)
    If lngEnd = 0 Then
        'Invalid EMail
        strMessage = Options.GetMessage("InvalidOrdersEMail", QuoteText(strEMail))
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
    strHeader = Trim(strHeader)
    strHeader = Replace(strHeader, vbTab, " ")
    While InStr(1, strHeader, "  ") > 0
        strHeader = Replace(strHeader, "  ", " ")
    Wend
    
    strSubject = "Major Problems Processing your orders"
    'Split the header by arguements
    varHeader = Split(strHeader, " ")
    ReDim Preserve varHeader(6)
    If varHeader(0) <> "#galaxy" Then
        'Invalid Header
        strMessage = Options.GetMessage("InvalidOrdersHeader", _
                "An invalid #galaxy line was specified", _
                QuoteText(strOrders))
        GoTo Error
    End If
    strGame = varHeader(1)
    strRace = varHeader(2)
    strPassword = varHeader(3)
    lngTurn = Val(varHeader(4))
    strFinalOrders = varHeader(5)
    
    'Validate the Game
    Set objGame = GalaxyNG.Games(strGame)
    
    If objGame Is Nothing Then
        'Invalid Header
        strMessage = Options.GetMessage("InvalidOrdersHeader", _
                "An unknown game was specified.", _
                QuoteText(strOrders))
        GoTo Error
    End If
    
    objGame.Refresh
    Set objRace = objGame.Races(strRace)
    If objRace Is Nothing Then
        'Invalid Header
        strMessage = Options.GetMessage("InvalidOrdersHeader", _
                "An unknown race was specified.", _
                QuoteText(strOrders))
        GoTo Error
    End If
    
    If objRace.Password <> strPassword Then
        'Invalid Header
        strMessage = Options.GetMessage("InvalidOrdersHeader", _
                "An invalid password was specified for the selected race.", _
                QuoteText(strOrders))
        GoTo Error
    End If
    
    If lngTurn = 0 Then
        'Invalid Header
        strMessage = Options.GetMessage("InvalidOrdersHeader", _
                "The turn number is missing from the #galaxy header line.", _
                QuoteText(strOrders))
        GoTo Error
    End If
    
    If lngTurn < objGame.NextTurn Then
        'Invalid Header
        strMessage = Options.GetMessage("InvalidOrdersHeader", _
                "The turn number is for a turn that has already been processed.", _
                QuoteText(strOrders))
        GoTo Error
    End If
    
    If strFinalOrders = "" Then
        blnFinalOrders = False
    ElseIf strFinalOrders = "finalorders" Then
        blnFinalOrders = True
    Else
        'Invalid Header
        strMessage = Options.GetMessage("InvalidOrdersHeader", _
            "The Finalorders parameter was specified but it is not ""finalorders""", _
            QuoteText(strOrders))
        GoTo Error
    End If
    
    If lngTurn > objGame.NextTurn Then
        strSubject = "[GNG] " & objGame.GameName & _
                    " turn " & CStr(lngTurn) & _
                    IIf(blnFinalOrders, " finalorders", " orders") & _
                    " received for " & objRace.RaceName
        strMessage = Options.GetMessage("Header", Options.ServerName) & _
                    Options.GetMessage("FutureOrders", lngTurn, MarkText(strOrders)) & _
                    Options.GetMessage("Footer", Options.ServerName)
        
    Else 'if lngTurn = objGame.NextTurn Then
        strSubject = "[GNG] " & objGame.GameName & _
                    " turn " & CStr(lngTurn) & _
                    " text" & _
                    IIf(blnFinalOrders, " finalorders", "") & _
                    " forecast for " & objRace.RaceName
        ' check it
        Call SaveFile(Options.GalaxyNGHome & Options.OrdersFileName, strOrders)
        Call RunGalaxyNG("-check " & strGame & " " & strRace & "<" & Options.OrdersFileName & " >" & Options.ForecastFileName)
        strMessage = GetFile(Options.GalaxyNGHome & Options.ForecastFileName)
        Kill Options.GalaxyNGHome & Options.OrdersFileName
        Kill Options.GalaxyNGHome & Options.ForecastFileName
    End If

    'File the orders
    strFileName = Options.GalaxyNGOrders & strGame & "\" & strRace & "." & CStr(lngTurn)
    strFileName1 = Options.GalaxyNGOrders & strGame & "\" & strRace & "_final" & "." & CStr(lngTurn)
    
    If Dir(strFileName) <> "" Then Kill strFileName
    If Dir(strFileName1) <> "" Then Kill strFileName1
    If blnFinalOrders Then
        Call SaveFile(strFileName1, strOrders)
    Else
        Call SaveFile(strFileName, strOrders)
    End If

    GoTo Send

Error:
    strMessage = Options.GetMessage("Header") & _
                    strMessage & _
                    Options.GetMessage("Footer")
Send:
    Call SendEMail(strFrom, strSubject, strMessage)
End Sub



