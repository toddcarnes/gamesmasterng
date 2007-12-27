VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NewEMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Attachments"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mstrToAddress As String
Private mstrFromAddress As String
Private mstrSubject As String
Private mdtDateSent As Date
Private mstrEMail As String
Private mstrHeader As String
Private mstrBody As String

Private mcolAttachments As Attachments

Public Property Get EMailData() As String
    Call MakeEMail
    EMailData = mstrEMail
End Property

Public Sub MakeEMail()
    Dim strSeperator As String
    Dim strEMail As String
    Dim objTimeZone As CTimeZone
    Dim objAttachment As Attachment
    
    Set objTimeZone = New CTimeZone
    
    mstrEMail = "To: " & mstrToAddress & vbNewLine & _
                "From: " & mstrFromAddress & vbNewLine & _
                "Date: " & objTimeZone.TimeEMail(mdtDateSent) & vbNewLine & _
                "Subject: " & mstrSubject & vbNewLine
    
    strSeperator = NewSeperator
    mstrEMail = mstrEMail & _
                "MIME-version: 1.0" & vbNewLine & _
                "Content-type: multipart/mixed; boundary=""" & strSeperator & """" & vbNewLine & _
                vbNewLine
    
    For Each objAttachment In mcolAttachments
        mstrEMail = mstrEMail & _
                    "--" & strSeperator & vbNewLine & _
                    objAttachment.EMailData & vbNewLine
    Next objAttachment
End Sub

Public Property Get Attachments() As Attachments
    If mcolAttachments Is Nothing Then
        Set mcolAttachments = New Attachments
    End If

    Set Attachments = mcolAttachments
End Property

Public Property Set Attachments(vData As Attachments)
    Set mcolAttachments = vData
End Property

Public Property Let Subject(ByVal strSubject As String)
    mstrSubject = strSubject
End Property

Public Property Get Subject() As String
    Subject = mstrSubject
End Property

Public Property Let ToAddress(ByVal strToAddress As String)
    mstrToAddress = strToAddress
End Property

Public Property Get ToAddress() As String
    ToAddress = mstrToAddress
End Property

Public Property Let FromAddress(ByVal strFromAddress As String)
    mstrFromAddress = strFromAddress
End Property

Public Property Get FromAddress() As String
    FromAddress = mstrFromAddress
End Property

Public Property Let DateSent(ByVal dtDateSent As Date)
    mdtDateSent = dtDateSent
End Property

Public Property Get DateSent() As Date
    DateSent = mdtDateSent
End Property

Private Sub Class_Terminate()
    Set mcolAttachments = Nothing
End Sub

Public Function NewSeperator() As String
    Dim strSeperator
    Dim blnFound As Boolean
    Dim objAttachment As Attachment
    
    ' Generate a seperator
    Do
        ' Generate a seperator
        blnFound = False
        strSeperator = MakeSeperator
        
        ' Check to see if the seperator is unique
        For Each objAttachment In Attachments
            If InStr(1, objAttachment.EMailData, strSeperator) > 0 Then
                blnFound = True
                Exit For
            End If
        Next objAttachment
        
        ' Use it if it is unique
        If Not blnFound Then Exit Do
    Loop
    
    NewSeperator = strSeperator
End Function

Public Function MakeSeperator() As String
    Dim bData() As Byte
    Dim i As Long
    Dim objBase64 As Base64
    Const cLen = 18
    
    Randomize
    ReDim bData(cLen - 1)
    For i = 0 To cLen - 1
        bData(i) = Rnd() * 255
    Next i
    Set objBase64 = New Base64
    MakeSeperator = objBase64.EncodeArr(bData)
    Set objBase64 = Nothing
End Function


