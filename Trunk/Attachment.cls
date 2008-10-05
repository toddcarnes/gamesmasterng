VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Attachment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
'********************************************************
'   Copyright 2007,2008 Ian Evans.                      *
'   This program is distributed under the terms of the  *
'       GNU General Public License.                     *
'********************************************************
Option Explicit

Public Key As String
Private mstrSource As String
Private mstrData As String
Private mstrEMailData As String
Private muFormat As uefFormat

Public Enum uefFormat
    uefText = 1
    uefBinary = 2
    uefHTML = 3
End Enum

Public Sub Store(ByVal strData As String, ByVal uFormat As uefFormat, Optional ByVal strSource As String = "")
    Dim objBase64 As Base64
    
    mstrData = strData
    mstrSource = strSource
    muFormat = uFormat
    
    ' Test attachment
    If uFormat = uefText Then
        mstrEMailData = "Content-Type: text/plain"
        If strSource <> "" Then
            mstrEMailData = mstrEMailData & _
                            "; name=""" & GetFullFileName(mstrSource) & """"
        End If
        mstrEMailData = mstrEMailData & vbNewLine & _
                        "Content-Transfer-Encoding: 7bit"
        If strSource <> "" Then
            mstrEMailData = mstrEMailData & vbNewLine & _
                        "Content-Disposition: inline; filename=""" & _
                        GetFullFileName(mstrSource) & """"
        End If
        mstrEMailData = mstrEMailData & vbNewLine & _
                    vbNewLine & _
                    strData
    
    'treat everything else as a binary attachment
    Else
        mstrEMailData = "Content-Type: application/octet-stream"
        If strSource <> "" Then
            mstrEMailData = mstrEMailData & _
                            "; name=""" & GetFullFileName(mstrSource) & """"
        End If
        mstrEMailData = mstrEMailData & vbNewLine & _
                        "Content-Transfer-Encoding: base64"
        If strSource <> "" Then
            mstrEMailData = mstrEMailData & vbNewLine & _
                        "Content-Disposition: inline; filename=""" & _
                        GetFullFileName(mstrSource) & """"
        End If
        Set objBase64 = New Base64
        mstrEMailData = mstrEMailData & vbNewLine & _
                        vbNewLine & _
                        objBase64.Encode(strData)
        Set objBase64 = Nothing
    End If
End Sub

Public Property Let Format(ByVal uFormat As uefFormat)
    muFormat = uFormat
End Property

Public Property Get Format() As uefFormat
    Format = muFormat
End Property

Public Property Let Data(ByVal vData As String)
    mstrData = vData
End Property

Public Property Get Data() As String
    Data = mstrData
End Property

Public Property Let Source(ByVal vData As String)
    mstrSource = vData
End Property

Public Property Get Source() As String
    Source = mstrSource
End Property

Public Property Get EMailData() As String
    EMailData = mstrEMailData
End Property

Public Property Let EMailData(strEMailData As String)
    mstrEMailData = strEMailData
End Property



