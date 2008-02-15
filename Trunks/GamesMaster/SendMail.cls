VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SendMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Option Compare Text

Public Event LogData(ByVal strData As String)
Public Event Connecting(ByVal strServer As String)
Public Event Sending(ByVal lngEMail As Long, ByVal lngTotal As Long)
Public Event Closing()
Public Event Disconnected()
Private WithEvents mobjSocket As CSocketMaster
Attribute mobjSocket.VB_VarHelpID = -1

Private muStatus As SM_Command
Public Enum SM_Command
    SM_CLOSE = 0
    SM_OPEN
    SM_HELO
    SM_MAIL
    SM_RCPT
    SM_DATA
    SM_Sent
    SM_QUIT
End Enum
Const c_CLOSE = "CLOSE"
Const c_HELO = "HELO "
Const c_MAIL = "MAIL FROM: "
Const c_RCPT = "RCPT TO: "
Const c_DATA = "DATA"
Const c_QUIT = "QUIT"

Private mlngMessages As Long
Private mlngCurrentMessage As Long
Private mobjEMail As EMail
Private mvarEMail As Variant
Private mlngEMail As Long

Public Property Get Status() As SM_Command
    Status = muStatus
End Property

Public Sub Send()
    Dim strFileName As String
    Dim e As Long
    
    If Status = SM_CLOSE Then
        ' Abort if the outbox is empty
        strFileName = Dir(Options.Outbox & "*.eml")
        If strFileName = "" Then Exit Sub
        ReDim mvarEMail(100)
        mlngEMail = 0
        e = -1
        'Build an array of the emails
        While strFileName <> ""
            e = e + 1
            If e > UBound(mvarEMail) Then
                ReDim Preserve mvarEMail(e + 99)
            End If
            mvarEMail(e) = Options.Outbox & strFileName
            strFileName = Dir()
        Wend
        ReDim Preserve mvarEMail(e)
        mlngMessages = UBound(mvarEMail) + 1
        
        'Connect to the mail server
        If Status <> GM_CLOSE Then
            RaiseEvent LogData("SendMail.Send - Status is NOT GM_CLOSE. Changing." & vbNewLine)
            Call WriteLogFile("SendMail.Send - Status is NOT GM_CLOSE. Changing.")
            muStatus = GM_CLOSE
        End If
        
        Call Connect
    End If
End Sub

Private Sub Connect()
    RaiseEvent Connecting(Options.SMTPServer)
    RaiseEvent LogData("> open " & Options.SMTPServer & ":" & Options.SMTPServerPort & vbNewLine)
    If mobjSocket Is Nothing Then
        Set mobjSocket = New CSocketMaster
    End If
    If mobjSocket.State <> sckClosed Then
        RaiseEvent LogData("SendMail.Send - mobjSocket.State is NOT Closed. Closing" & vbNewLine)
        Call WriteLogFile("SendMail.Send - mobjSocket.State is NOT Closed. Closing")
        mobjSocket.CloseSck
        While mobjSocket.State <> sckClosed
            DoEvents
        Wend
        RaiseEvent LogData("SendMail.Send - mobjSocket.State has Closed. Continuing." & vbNewLine)
        Call WriteLogFile("SendMail.Send - mobjSocket.State has Closed. Continuing.")
    End If
    With mobjSocket
        .Protocol = sckTCPProtocol
        .Connect Options.SMTPServer, Options.SMTPServerPort
    End With
    muStatus = SM_OPEN
End Sub

Private Sub Disconnect()
    RaiseEvent Closing
    RaiseEvent LogData("> closed ")
    With mobjSocket
        .CloseSck
    End With
    muStatus = SM_CLOSE
'    Set mobjSocket = Nothing
    RaiseEvent Disconnected
End Sub

Private Sub ProcessData(ByVal strData As String)
    Dim varFields As Variant
    Dim i As Long
    Dim lngResponse As Long
    Dim lngError As Long
    Dim strError As String
    Dim strSource As String
    Dim lngState As Long
    Dim strFileName As String
    Dim blnNext As Boolean
    Dim blnError As Boolean
    
    On Error GoTo ErrorTag
    
    DoEvents
    RaiseEvent LogData("S: " & strData)
    i = InStr(1, strData, " ")
    lngResponse = Val(Left(strData, i - 1))
    
    Select Case muStatus
    Case SM_Command.SM_OPEN
        If lngResponse = 220 Then
            SendData c_HELO & Options.ServerName
            muStatus = SM_HELO
        Else
            SendData c_QUIT
            muStatus = SM_QUIT
        End If
        
    Case SM_Command.SM_HELO
        If lngResponse = 250 Then
            RaiseEvent Sending(mlngEMail + 1, mlngMessages)
            Set mobjEMail = New EMail
            mobjEMail.Load mvarEMail(mlngEMail)
            SendData c_MAIL & mobjEMail.From
            muStatus = SM_MAIL
        Else
            SendData c_QUIT
            muStatus = SM_QUIT
        End If
        
    Case SM_Command.SM_MAIL
        If lngResponse = 250 Then
            SendData c_RCPT & mobjEMail.SendTo
            muStatus = SM_RCPT
        Else
            SendData c_QUIT
            muStatus = SM_QUIT
        End If
        
    Case SM_Command.SM_RCPT
        If lngResponse = 250 _
        Or lngResponse = 251 Then
            SendData c_DATA
            muStatus = SM_DATA
        ElseIf lngResponse = 550 _
        Or lngResponse = 551 _
        Or lngResponse = 552 _
        Or lngResponse = 553 _
        Or lngResponse = 450 _
        Or lngResponse = 451 _
        Or lngResponse = 452 Then
            blnNext = True
            blnError = True
        Else
            SendData c_QUIT
            muStatus = SM_QUIT
        End If
        
    Case SM_Command.SM_DATA
        If lngResponse = 354 Then
            SendData mobjEMail.Body & vbCrLf & "."
            muStatus = SM_Sent
        Else
            SendData c_QUIT
            muStatus = SM_QUIT
        End If
    
    Case SM_Command.SM_Sent
        If lngResponse = 250 Then
            blnNext = True
        Else
            SendData c_QUIT
            muStatus = SM_QUIT
        End If
    
    Case SM_Command.SM_QUIT
        Call Disconnect
    Case Else
    
    End Select
    
    ' Process the next e-mail
    If blnNext Then
        'Save the Erroreous e-mail
        If blnError Then
            strFileName = mvarEMail(mlngEMail) & ".err"
            i = 0
            While Dir(strFileName) <> ""
                i = i + 1
                strFileName = mvarEMail(mlngEMail) & "_" & CStr(i) & ".err"
            Wend
            Name mvarEMail(mlngEMail) As strFileName
            blnError = False
            
        'Save the e-mail
        ElseIf Options.SaveEMail Then
            strFileName = mvarEMail(mlngEMail) & ".sav"
            i = 0
            While Dir(strFileName) <> ""
                i = i + 1
                strFileName = mvarEMail(mlngEMail) & "_" & CStr(i) & ".sav"
            Wend
            Name mvarEMail(mlngEMail) As strFileName
        
        'Delete the email
        Else
            Kill mvarEMail(mlngEMail)
        End If
        
        mlngEMail = mlngEMail + 1
        ' No more e-mail
        If mlngEMail > UBound(mvarEMail) Then
            SendData c_QUIT
            muStatus = SM_QUIT
        
        'Process the next email
        Else
            RaiseEvent Sending(mlngEMail, mlngMessages)
            Set mobjEMail = New EMail
            mobjEMail.Load mvarEMail(mlngEMail)
            SendData c_MAIL & mobjEMail.From
            muStatus = SM_MAIL
        End If
        blnNext = False
    End If
    
    Exit Sub
    
ErrorTag:
    lngError = Err.Number
    strError = Err.Description
    strSource = Err.Source
    lngState = mobjSocket.State
    Call LogError(lngError, strError, strSource, _
                "SendMail", "Processdata", _
                "    Mail Status: " & CStr(Status) & vbNewLine & _
                "    Socket State: " & CStr(lngState))
    Call Disconnect
End Sub

Private Sub SendData(ByVal strData As String)
    On Error GoTo ErrorTag
    
    RaiseEvent LogData("> " & strData & vbCrLf)
    mobjSocket.SendData strData & vbCrLf
    Exit Sub
    
ErrorTag:
    Err.Raise Err.Number, "SendMail.Senddata." & Err.Source, Err.Description
End Sub

Private Sub Class_Terminate()
    If Status <> SM_CLOSE Then
        mobjSocket.CloseSck
    End If
    muStatus = GM_CLOSE
    Set mobjSocket = Nothing
End Sub

Private Sub mobjSocket_CloseSck()
    RaiseEvent Disconnected
End Sub

Private Sub mobjSocket_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String
    
    Call mobjSocket.GetData(strData, vbString)
    Call ProcessData(strData)
End Sub

Private Sub SaveEMail(ByVal strEMail As String)
    Dim intFN As Integer
    Dim i As Long
    Dim strFileName As String
    
    Do
        strFileName = Options.Inbox & Format(Now, "yyyymmddhhnnss") & " " & Format(i, "00000") & ".txt"
        If Dir(strFileName) = "" Then Exit Do
        i = i + 1
    Loop
        
    intFN = FreeFile
    Open strFileName For Output As #intFN
    Print #intFN, strEMail;
    Close #intFN

End Sub



