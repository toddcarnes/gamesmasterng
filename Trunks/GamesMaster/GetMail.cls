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
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"GetMail"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Option Compare Text

Public Event LogData(ByVal strData As String)
Public Event Connecting(ByVal strServer As String)
Public Event Validating()
Public Event Receiving(ByVal lngEMail As Long, ByVal lngTotal As Long)
Public Event Closing()
Public Event Disconnected()

Private WithEvents mobjSocket As CSocketMaster
Attribute mobjSocket.VB_VarHelpID = -1

Private muStatus As GM_Command
Public Enum GM_Command
    GM_CLOSE = 0
    GM_OPEN
    GM_USER
    GM_PASS
    GM_STAT
    GM_LIST
    GM_RETR
    GM_DELE
    GM_NOOP
    GM_RSET
    GM_QUIT
    GM_APOP
    GM_TOP
    GM_UIDL
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

Public Property Get Status() As GM_Command
    Status = muStatus
End Property

Public Sub GetMail()
    If Status = GM_CLOSE Then
        Call Connect
    End If
End Sub

Private Sub Connect()
    RaiseEvent Connecting(Options.POPServer)
    RaiseEvent LogData("> open " & Options.POPServer & ":" & Options.POPServerPort & vbNewLine)
    If mobjSocket Is Nothing Then
        Set mobjSocket = New CSocketMaster
    End If
    If mobjSocket.State <> sckClosed Then
        mobjSocket.CloseSck
        DoEvents
    End If
    With mobjSocket
        .Protocol = sckTCPProtocol
        .Connect Options.POPServer, Options.POPServerPort
    End With
    muStatus = GM_OPEN
End Sub

Private Sub Disconnect()
    RaiseEvent LogData("> closed ")
    RaiseEvent Closing
    With mobjSocket
        .CloseSck
    End With
'    Set mobjSocket = Nothing
    muStatus = GM_CLOSE
    RaiseEvent Disconnected
End Sub

Private Sub ProcessData(ByVal strData As String)
    Dim blnOK As Boolean
    Dim blnError As Boolean
    Dim varFields As Variant
    Dim blnEMailOK As Boolean
    
    DoEvents
    RaiseEvent LogData("S: " & strData)
    blnOK = (Left(strData, Len(c_OK)) = c_OK)
    blnError = (Left(strData, Len(c_ERR)) = c_ERR)
    
    Select Case muStatus
    Case GM_Command.GM_OPEN
        If blnOK Then
            RaiseEvent Validating
            SendData c_USER & " " & Options.POPUserID
            muStatus = GM_USER
        Else
            SendData c_QUIT
            muStatus = GM_QUIT
        End If
        
    Case GM_Command.GM_USER
        If blnOK Then
            SendData c_PASS & " " & Options.POPPassword
            muStatus = GM_PASS
        Else
            SendData c_QUIT
            muStatus = GM_QUIT
        End If
        
    Case GM_Command.GM_PASS
        If blnOK Then
            SendData c_STAT
            muStatus = GM_STAT
        Else
            SendData c_QUIT
            muStatus = GM_QUIT
        End If
        
    Case GM_Command.GM_STAT
        If blnOK Then
            varFields = Split(strData, " ")
            mlngMessages = Val(varFields(1))
            If mlngMessages > 0 Then
                mlngCurrentMessage = 1
                mstrEMail = ""
                RaiseEvent Receiving(mlngCurrentMessage, mlngMessages)
                SendData c_RETR & " " & CStr(mlngCurrentMessage)
                muStatus = GM_RETR
            Else
                SendData c_QUIT
                muStatus = GM_QUIT
            End If
        Else
            SendData c_QUIT
            muStatus = GM_QUIT
        End If
        
    Case GM_Command.GM_LIST
    
    Case GM_Command.GM_RETR
        If blnOK Then
            'ignore this line and receive the message
        ElseIf blnError Then
            SendData c_QUIT
            muStatus = GM_QUIT
        ElseIf strData = "." & vbCrLf Then
            Call SaveEMail(mstrEMail)
            SendData c_DETE & " " & CStr(mlngCurrentMessage)
            muStatus = GM_DELE
        ElseIf Right(strData, 5) = vbCrLf & "." & vbCrLf Then
            mstrEMail = mstrEMail & Left(strData, Len(strData) - 5)
            Call SaveEMail(mstrEMail)
            SendData c_DETE & " " & CStr(mlngCurrentMessage)
            muStatus = GM_DELE
        Else
            mstrEMail = mstrEMail & strData
        End If
    
    Case GM_Command.GM_DELE
        If blnOK Then
            mlngCurrentMessage = mlngCurrentMessage + 1
            If mlngCurrentMessage <= mlngMessages Then
                mstrEMail = ""
                RaiseEvent Receiving(mlngCurrentMessage, mlngMessages)
                SendData c_RETR & " " & CStr(mlngCurrentMessage)
                muStatus = GM_RETR
            Else
                SendData c_QUIT
                muStatus = GM_QUIT
            End If
        Else
            SendData c_QUIT
            muStatus = GM_QUIT
        End If
    
    Case GM_Command.GM_QUIT
        Call Disconnect
    Case GM_Command.GM_RSET
    
    Case GM_Command.GM_NOOP
    
    Case GM_Command.GM_TOP
    
    Case GM_Command.GM_APOP
    
    Case GM_Command.GM_UIDL
    
    Case Else
    
    End Select
End Sub

Private Sub SendData(ByVal strData As String)
    RaiseEvent LogData("> " & strData & vbCrLf)
    mobjSocket.SendData strData & vbCrLf
End Sub

Private Sub Class_Terminate()
    If Status <> GM_CLOSE Then
        mobjSocket.CloseSck
    End If
    muStatus = GM_CLOSE
    Set mobjSocket = Nothing
End Sub

Private Sub mobjSocket_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String
    
    Call mobjSocket.GetData(strData, vbString)
    Call ProcessData(strData)
End Sub

Private Sub SaveEMail(ByVal strEMail As String)
    Dim i As Long
    Dim strFileName As String
    
    Do
        strFileName = Options.Inbox & Format(Now, "yyyymmddhhnnss") & "_" & Format(i, "0") & ".txt"
        If Dir(strFileName) = "" Then Exit Do
        i = i + 1
    Loop
        
    Call SaveFile(strFileName, strEMail)
End Sub


