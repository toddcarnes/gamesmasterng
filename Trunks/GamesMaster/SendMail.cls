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
        strFileName = Dir(Outbox & "*.txt")
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
            mvarEMail(e) = Outbox & strFileName
            strFileName = Dir()
        Wend
        ReDim Preserve mvarEMail(e)
        
        'Connect to the mail server
        Call Connect
    End If
End Sub

Private Sub Connect()
    RaiseEvent Connecting(SMTPServer)
    RaiseEvent LogData("> open " & SMTPServer & ":" & SMTPServerPort & vbNewLine)
    With mobjSocket
        .Protocol = sckTCPProtocol
        .Connect SMTPServer, SMTPServerPort
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
End Sub

Private Sub ProcessData(ByVal strData As String)
    Dim varFields As Variant
    Dim i As Long
    Dim lngResponse As Long
    
    DoEvents
    RaiseEvent LogData(strData)
    i = InStr(1, strData, " ")
    lngResponse = Val(Left(strData, i - 1))
    
    Select Case muStatus
    Case SM_Command.SM_OPEN
        If lngResponse = 220 Then
            SendData c_HELO & ServerName
            muStatus = SM_HELO
        Else
            SendData c_QUIT
            muStatus = SM_QUIT
        End If
        
    Case SM_Command.SM_HELO
        If lngResponse = 250 Then
            RaiseEvent Sending(mlngEMail, mlngMessages)
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
        If lngResponse = 250 Then
            SendData c_DATA
            muStatus = SM_DATA
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
            Name mvarEMail(mlngEMail) As mvarEMail(mlngEMail) & ".sav"
            mlngEMail = mlngEMail + 1
            If mlngEMail > UBound(mvarEMail) Then
                SendData c_QUIT
                muStatus = SM_QUIT
            Else
                RaiseEvent Sending(mlngEMail, mlngMessages)
                Set mobjEMail = New EMail
                mobjEMail.Load mvarEMail(mlngEMail)
                SendData c_MAIL & " " & mobjEMail.From
                muStatus = SM_MAIL
            End If
        Else
            SendData c_QUIT
            muStatus = SM_QUIT
        End If
    
    Case SM_Command.SM_QUIT
        Call Disconnect
    Case Else
    
    End Select
End Sub

Private Sub SendData(ByVal strData As String)
    RaiseEvent LogData("> " & strData & vbCrLf)
    mobjSocket.SendData strData & vbCrLf
End Sub

Private Sub Class_Initialize()
    Set mobjSocket = New CSocketMaster
End Sub

Private Sub Class_Terminate()
    If Status <> SM_CLOSE Then
        mobjSocket.CloseSck
    End If
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
        strFileName = Inbox & Format(Now, "yyyymmddhhnnss") & " " & Format(i, "00000") & ".txt"
        If Dir(strFileName) = "" Then Exit Do
        i = i + 1
    Loop
        
    intFN = FreeFile
    Open strFileName For Output As #intFN
    Print #intFN, strEMail;
    Close #intFN

End Sub


