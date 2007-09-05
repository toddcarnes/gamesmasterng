VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GetMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Public Event LogData(ByVal strData As String)
Public Event NewConnection()
Private WithEvents mobjSocket As CSocketMaster
Attribute mobjSocket.VB_VarHelpID = -1


Private muStatus As SM_Command
Public Enum SM_Command
    SM_CLOSE = 0
    SM_OPEN
    SM_USER
    SM_PASS
    SM_STAT
    SM_LIST
    SM_RETR
    SM_DELE
    SM_NOOP
    SM_RSET
    SM_QUIT
    SM_APOP
    SM_TOP
    SM_UIDL
End Enum
Const c_CLOSE = "CLOSE"
Const c_OPEN = "OPEN"
Const c_USER = "USER"
Const c_PASS = "PASS"
Const c_STAT = "STAT"
Const c_LIST = "LIST"
Const c_RETR = "RETR"
Const c_DETE = "DELE"
Const c_NOOP = "NOOP"
Const c_RSET = "RSET"
Const c_QUIT = "QUIT"
Const c_APOP = "APOP"
Const c_TOP = "TOP"
Const c_UIDL = "UIDL"
Const c_OK = "+OK"
Const c_ERR = "-ERR"

Private mlngMessages As Long
Private mlngCurrentMessage As Long
Private mstrEMail As String

Public Property Get Status() As SM_Command
    Status = muStatus
End Property

Public Sub GetMail()
    Call Connect
End Sub

Private Sub Connect()
    RaiseEvent NewConnection
    RaiseEvent LogData("> open " & POPServer & ":" & POPServerPort & vbNewLine)
    With mobjSocket
        .Protocol = sckTCPProtocol
        .Connect POPServer, POPServerPort
    End With
    muStatus = SM_OPEN
End Sub

Private Sub Disconnect()
    RaiseEvent LogData("> closed ")
    With mobjSocket
        .CloseSck
    End With
    muStatus = SM_CLOSE
End Sub

Private Sub ProcessData(ByVal strData As String)
    Dim blnOK As Boolean
    Dim blnError As Boolean
    Dim varFields As Variant
    Dim blnEMailOK As Boolean
    
    RaiseEvent LogData(strData)
    
    blnOK = (Left(strData, Len(c_OK)) = c_OK)
    blnError = (Left(strData, Len(c_ERR)) = c_ERR)
    
    Select Case muStatus
    Case SM_Command.SM_OPEN
        If blnOK Then
            SendData c_USER & " " & POPUserID
            muStatus = SM_USER
        Else
            SendData c_QUIT
            muStatus = SM_QUIT
        End If
        
    Case SM_Command.SM_USER
        If blnOK Then
            SendData c_PASS & " " & POPPassword
            muStatus = SM_PASS
        Else
            SendData c_QUIT
            muStatus = SM_QUIT
        End If
        
    Case SM_Command.SM_PASS
        If blnOK Then
            SendData c_STAT
            muStatus = SM_STAT
        Else
            SendData c_QUIT
            muStatus = SM_QUIT
        End If
        
    Case SM_Command.SM_STAT
        If blnOK Then
            varFields = Split(strData, " ")
            mlngMessages = Val(varFields(1))
            If mlngMessages > 0 Then
                mlngCurrentMessage = 1
                mstrEMail = ""
                SendData c_RETR & " " & CStr(mlngCurrentMessage)
                muStatus = SM_RETR
            Else
                SendData c_QUIT
                muStatus = SM_QUIT
            End If
        Else
            SendData c_QUIT
            muStatus = SM_QUIT
        End If
        
    Case SM_Command.SM_LIST
    
    Case SM_Command.SM_RETR
        If blnOK Then
            'ignore this line and receive the message
        ElseIf blnError Then
            SendData c_QUIT
            muStatus = SM_QUIT
        ElseIf Right(strData, 5) = vbCrLf & "." & vbCrLf Then
            mstrEMail = mstrEMail & Left(strData, Len(strData) - 5)
            Call SaveEMail(mstrEMail)
            SendData c_DETE & " " & CStr(mlngCurrentMessage)
            muStatus = SM_DELE
        Else
            mstrEMail = mstrEMail & strData
        End If
    
    Case SM_Command.SM_DELE
        If blnOK Then
            mlngCurrentMessage = mlngCurrentMessage + 1
            If mlngCurrentMessage <= mlngMessages Then
                mstrEMail = ""
                SendData c_RETR & " " & CStr(mlngCurrentMessage)
                muStatus = SM_RETR
            Else
                SendData c_QUIT
                muStatus = SM_QUIT
            End If
        Else
            SendData c_QUIT
            muStatus = SM_QUIT
        End If
    
    Case SM_Command.SM_QUIT
        Call Disconnect
    Case SM_Command.SM_RSET
    
    Case SM_Command.SM_NOOP
    
    Case SM_Command.SM_TOP
    
    Case SM_Command.SM_APOP
    
    Case SM_Command.SM_UIDL
    
    Case Else
    
    End Select
End Sub

Public Sub SendData(ByVal strData As String)
    RaiseEvent LogData("> " & strData & vbCrLf)
    mobjSocket.SendData strData & vbCrLf
End Sub

Private Sub Class_Initialize()
    Set mobjSocket = New CSocketMaster
End Sub

Private Sub Class_Terminate()
    Set mobjSocket = Nothing
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


