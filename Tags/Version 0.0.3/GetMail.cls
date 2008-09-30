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
    If Status <> GM_CLOSE Then
        RaiseEvent LogData("GetMail.GetMail - Status is NOT GM_CLOSE. Changing." & vbNewLine)
        Call WriteLogFile("GetMail.GetMail - Status is NOT GM_CLOSE. Changing.")
        muStatus = GM_CLOSE
    End If
    
    Call Connect
End Sub

Private Sub Connect()
    RaiseEvent Connecting(Options.POPServer)
    RaiseEvent LogData("> open " & Options.POPServer & ":" & Options.POPServerPort & vbNewLine)
    If mobjSocket Is Nothing Then
        Set mobjSocket = New CSocketMaster
    End If
    If mobjSocket.State <> sckClosed Then
        RaiseEvent LogData("GetMail.GetMail - mobjSocket.State is NOT Closed. Closing" & vbNewLine)
        Call WriteLogFile("GetMail.GetMail - mobjSocket.State is NOT Closed. Closing")
        mobjSocket.CloseSck
        While mobjSocket.State <> sckClosed
            DoEvents
        Wend
        RaiseEvent LogData("GetMail.GetMail - mobjSocket.State has Closed. Continuing." & vbNewLine)
        Call WriteLogFile("GetMail.GetMail - mobjSocket.State has Closed. Continuing.")
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
    Dim blnNextEMail As Boolean
    Dim blnSaveEMail As Boolean
    Dim blnAbort As Boolean
    Dim blnSendLogon As Boolean
    Dim blnSendPassword As Boolean
    Dim blnGetEMailCount As Boolean
    Dim lngError As Long
    Dim strError As String
    Dim strSource As String
    Dim lngState As Long
    
    On Error GoTo ErrorTag
    DoEvents
    RaiseEvent LogData("S: " & strData)
    blnOK = (Left(strData, Len(c_OK)) = c_OK)
    blnError = (Left(strData, Len(c_ERR)) = c_ERR)
    
    Select Case muStatus
    Case GM_Command.GM_OPEN
        If blnOK Then
            blnSendLogon = True
        Else
            blnAbort = True
        End If
        
    Case GM_Command.GM_USER
        If blnOK Then
            blnSendPassword = True
        Else
            blnAbort = True
        End If
        
    Case GM_Command.GM_PASS
        If blnOK Then
            blnGetEMailCount = True
        Else
            blnAbort = True
        End If
        
    Case GM_Command.GM_STAT
        If blnOK Then
            varFields = Split(strData, " ")
            mlngMessages = Val(varFields(1))
            If mlngMessages > 0 Then
                mlngCurrentMessage = 0
                blnNextEMail = True
            Else
                blnAbort = True
            End If
        Else
            blnAbort = True
        End If
        
    Case GM_Command.GM_LIST
    
    Case GM_Command.GM_RETR
        If blnOK Then
            'ignore this line and receive the message
        ElseIf blnError Then
            blnAbort = True
        ElseIf strData = "." & vbCrLf Then
            blnSaveEMail = True
        ElseIf Right(strData, 5) = vbCrLf & "." & vbCrLf Then
            mstrEMail = mstrEMail & Left(strData, Len(strData) - 5)
            blnSaveEMail = True
        Else
            mstrEMail = mstrEMail & strData
        End If
    
    Case GM_Command.GM_DELE
        If blnOK Then
            blnNextEMail = True
        Else
            blnAbort = True
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
    
    If blnSendLogon Then
        RaiseEvent Validating
        SendData c_USER & " " & Options.POPUserID
        muStatus = GM_USER
        blnSendLogon = False
    End If
    
    If blnSendPassword Then
        SendData c_PASS & " " & Options.POPPassword
        muStatus = GM_PASS
        blnSendPassword = False
    End If
    
    If blnSaveEMail Then
        ' Check Email
        If Not ValidEMail(mstrEMail) Then
            Call SaveEMail(mstrEMail, True)
            blnNextEMail = True
        Else
            Call SaveEMail(mstrEMail)
            SendData c_DETE & " " & CStr(mlngCurrentMessage)
            muStatus = GM_DELE
        End If
        blnSaveEMail = False
    End If
    
    If blnGetEMailCount Then
        SendData c_STAT
        muStatus = GM_STAT
        blnGetEMailCount = False
    End If
    
    If blnNextEMail Then
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
        blnNextEMail = False
    End If
    
    If blnAbort Then
        SendData c_QUIT
        muStatus = GM_QUIT
        blnAbort = False
    End If
    
    Exit Sub

ErrorTag:
    lngError = Err.Number
    strError = Err.Description
    strSource = Err.Source
    lngState = mobjSocket.State
    Call LogError(lngError, strError, strSource, _
                "GetMail", "Processdata", _
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
    Err.Raise Err.Number, "GetMail.Senddata." & Err.Source, Err.Description
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

Private Sub SaveEMail(ByVal strEMail As String, Optional ByVal blnError As Boolean = False)
    Dim i As Long
    Dim strFileName As String
    
    Do
        If blnError Then
            strFileName = Options.Inbox & Format(Now, "yyyymmddhhnnss") & "_invalid_" & Format(i, "0") & ".txt"
        Else
            strFileName = Options.Inbox & Format(Now, "yyyymmddhhnnss") & "_" & Format(i, "0") & ".eml"
        End If
        If Dir(strFileName) = "" Then Exit Do
        i = i + 1
    Loop
        
    Call SaveFile(strFileName, strEMail)
End Sub

Private Function ValidEMail(ByVal strEMail As String) As Boolean
    Dim lngHeaderEnd As Long
    Dim strHeader As String
    
    'Validate the EMail
    lngHeaderEnd = InStr(1, strEMail, vbCrLf & vbCrLf)
    strHeader = Left(strEMail, lngHeaderEnd)
    
    If InStr(1, strEMail, vbCrLf & "subject:", vbTextCompare) = 0 Then Exit Function
    If InStr(1, strEMail, vbCrLf & "from:", vbTextCompare) = 0 Then Exit Function
    If InStr(1, strEMail, vbCrLf & "to:", vbTextCompare) = 0 Then Exit Function
    ValidEMail = True
    
End Function

