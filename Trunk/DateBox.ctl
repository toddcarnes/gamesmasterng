VERSION 5.00
Begin VB.UserControl DateBox 
   ClientHeight    =   945
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1860
   ScaleHeight     =   945
   ScaleWidth      =   1860
   ToolboxBitmap   =   "DateBox.ctx":0000
   Begin VB.TextBox txtDate 
      Height          =   315
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "DateBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'********************************************************
'   Copyright 2007,2008 Ian Evans.                      *
'   This program is distributed under the terms of the  *
'       GNU General Public License.                     *
'********************************************************
Option Explicit
'Default Property Values:
'Const m_def_TimeStamp = 0
Const m_def_Date = 0
Const m_def_DateFormat = "Short Date"
Const m_def_TimeFormat = "Short Time"
'Property Variables:
Dim m_TimeStamp As Date
'Dim m_TimeStamp As Variant
Dim m_DateFormat As String
Dim m_TimeFormat As String
Dim mNullDate As Date
'Event Declarations:
Event Change()
Attribute Change.VB_Description = "Occurs when the contents of a control have changed."
'Event Validate(Cancel As Boolean) 'MappingInfo=txtDate,txtDate,-1,Validate
Event Click() 'MappingInfo=txtDate,txtDate,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=txtDate,txtDate,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=txtDate,txtDate,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=txtDate,txtDate,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=txtDate,txtDate,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=txtDate,txtDate,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=txtDate,txtDate,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=txtDate,txtDate,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."

Private Sub txtDate_GotFocus()
    With txtDate
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtDate_Validate(Cancel As Boolean)
    Dim S As Integer
    Dim dtNew As Date
    Dim strInput As String
    Dim strDate As String
    Dim strTime As String
    Dim intLen As Integer
    Dim dtLast As Date
    
    dtLast = m_TimeStamp

    If txtDate = "" Then
        m_TimeStamp = mNullDate
    Else
        If m_DateFormat = "" Then
            strDate = Format(CDate(0), "dd/mm/yyyy")
            strTime = txtDate
        ElseIf m_TimeFormat = "" Then
            strDate = txtDate
            strTime = "00:00:00"
        Else
            strInput = txtDate
            S = InStr(1, strInput, " ")
            If S > 0 Then
                strDate = Trim(Left(strInput, S - 1))
                strTime = Trim(Mid(strInput, S + 1))
            Else
                strDate = strInput
                strTime = ""
            End If
        End If
    
        If InStr(1, strDate, "/") = 0 Then
            intLen = Len(strDate)
            If intLen Mod 2 = 1 Then
                strDate = "0" & strDate
                intLen = intLen + 1
            End If
            If intLen = 0 Then
                strDate = strDate & Format(Now, "ddmmyyyy")
            ElseIf intLen = 2 Then
                strDate = strDate & Format(Now, "mmyyyy")
            ElseIf intLen = 4 Then
                strDate = strDate & Format(Now, "yyyy")
            End If
            strDate = Mid(strDate, 1, 2) & "/" & _
                    Mid(strDate, 3, 2) & "/" & _
                    Mid(strDate, 5)
        End If
        If strTime = "" Then
            strTime = "00:00:01"
        End If
        If InStr(1, strTime, ":") = 0 Then
            If Len(strTime) = 1 _
            Or Len(strTime) = 3 Then
                strTime = "0" & strTime
            End If
            If Len(strTime) = 2 Then
                strTime = strTime & "00"
            End If
            strTime = Mid(strTime, 1, 2) & ":" & _
                    Mid(strTime, 3)
        End If
        If IsDate(strDate & " " & strTime) Then
            If m_TimeFormat = "" Then
                m_TimeStamp = CDate(strDate)
            ElseIf m_DateFormat = "" Then
                m_TimeStamp = CDate(strTime)
            Else
                m_TimeStamp = CDate(strDate & " " & strTime)
            End If
            DisplayDate
        Else
            Beep
            Cancel = True
        End If
    End If
    If Cancel = 0 Then
        If dtLast <> m_TimeStamp Then
            RaiseEvent Change
        End If
    End If
End Sub


Private Sub UserControl_Resize()
    Dim blnResizing As Boolean
    
    If blnResizing Then Exit Sub
    blnResizing = True
    With txtDate
        .Top = 0
        .Left = 0
        .Width = UserControl.ScaleWidth
        .Height = UserControl.ScaleHeight
        UserControl.Height = .Height
    End With
    blnResizing = False
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtDate,txtDate,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = txtDate.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    txtDate.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtDate,txtDate,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = txtDate.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    txtDate.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtDate,txtDate,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = txtDate.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set txtDate.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackStyle
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=UserControl,UserControl,-1,BorderStyle
'Public Property Get BorderStyle() As Integer
'    BorderStyle = UserControl.BorderStyle
'End Property
'
'Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
'    UserControl.BorderStyle() = New_BorderStyle
'    PropertyChanged "BorderStyle"
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtDate,txtDate,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    txtDate.Refresh
End Sub

Private Sub txtDate_Click()
    RaiseEvent Click
End Sub

Private Sub txtDate_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub txtDate_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    Dim c As String
    Dim a As Integer
    Dim blncancel As Boolean
    Dim dtLast As Date
    Dim dtActual As Date
    
    If txtDate.Locked Then Exit Sub
    dtLast = m_TimeStamp
    dtActual = Format(dtLast, "dd-mmm-yyyy hh:nn")
    RaiseEvent KeyPress(KeyAscii)
    a = 0
    c = Chr(KeyAscii)
    Select Case c
    Case Chr$(vbKeyReturn)
        txtDate_Validate blncancel
    Case "="
        m_TimeStamp = Format(Now, "dd-mmm-yyyy hh:nn")
    Case "+"
        m_TimeStamp = DateAdd("d", 1, m_TimeStamp)
    Case "-"
        m_TimeStamp = DateAdd("d", -1, m_TimeStamp)
    Case "d"
        m_TimeStamp = DateAdd("d", 1, m_TimeStamp)
    Case "D"
        m_TimeStamp = DateAdd("d", -1, m_TimeStamp)
    Case "e"
        m_TimeStamp = DateAdd("d", 1, m_TimeStamp)
        m_TimeStamp = DateAdd("m", 1, m_TimeStamp)
        m_TimeStamp = CDate("01-" & Format(m_TimeStamp, "mmm-yyyy hh:nn:ss"))
        m_TimeStamp = DateAdd("d", -1, m_TimeStamp)
    Case "E"
        m_TimeStamp = CDate("01-" & Format(m_TimeStamp, "mmm-yyyy hh:nn:ss"))
        m_TimeStamp = DateAdd("d", -1, m_TimeStamp)
    Case "w"
        m_TimeStamp = DateAdd("d", 7, m_TimeStamp)
    Case "W"
        m_TimeStamp = DateAdd("d", -7, m_TimeStamp)
    Case "f"
        m_TimeStamp = DateAdd("d", 14, m_TimeStamp)
    Case "F"
        m_TimeStamp = DateAdd("d", -14, m_TimeStamp)
    Case "m"
        m_TimeStamp = DateAdd("m", 1, m_TimeStamp)
    Case "M"
        m_TimeStamp = DateAdd("m", -1, m_TimeStamp)
    Case "q"
        m_TimeStamp = DateAdd("m", 3, m_TimeStamp)
    Case "Q"
        m_TimeStamp = DateAdd("m", -3, m_TimeStamp)
    Case "s"
        m_TimeStamp = DateAdd("m", 1, m_TimeStamp)
        m_TimeStamp = CDate("01-" & Format(m_TimeStamp, "mmm-yyyy hh:nn:ss"))
    Case "S"
        m_TimeStamp = DateAdd("d", -1, m_TimeStamp)
        m_TimeStamp = CDate("01-" & Format(m_TimeStamp, "mmm-yyyy hh:nn:ss"))
    Case "y"
        m_TimeStamp = DateAdd("m", 12, m_TimeStamp)
    Case "Y"
        m_TimeStamp = DateAdd("m", -12, m_TimeStamp)
    Case "h"
        m_TimeStamp = DateAdd("h", 1, dtActual)
    Case "H"
        m_TimeStamp = DateAdd("h", -1, dtActual)
    Case "t"
        m_TimeStamp = DateAdd("n", 10, dtActual)
    Case "T"
        m_TimeStamp = DateAdd("n", -10, dtActual)
    Case "n"
        m_TimeStamp = DateAdd("n", 1, dtActual)
    Case "N"
        m_TimeStamp = DateAdd("n", -1, dtActual)
    Case Else
        If InStr(1, "0123456789 /:" & Chr$(vbKeyBack), c) > 0 Then
            a = KeyAscii
        End If
    End Select
    
    If KeyAscii <> a Then
        If Not blncancel Then
            DisplayDate
            With txtDate
                .SelStart = 0
                .SelLength = Len(txtDate)
            End With
        End If
    End If
    KeyAscii = a
    If Trim(m_DateFormat) = "" Then
        m_TimeStamp = Format(m_TimeStamp, "hh:nn:ss")
    End If
    If Trim(m_TimeFormat) = "" Then
        m_TimeStamp = Format(m_TimeStamp, "dd-mmm-yyyy")
    End If
    If dtLast <> m_TimeStamp Then
        RaiseEvent Change
    End If
End Sub


Private Sub txtDate_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub txtDate_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub txtDate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub txtDate_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,Short Date
Public Property Get DateFormat() As String
Attribute DateFormat.VB_ProcData.VB_Invoke_Property = ";Data"
    DateFormat = m_DateFormat
End Property

Public Property Let DateFormat(ByVal New_DateFormat As String)
    m_DateFormat = New_DateFormat
    PropertyChanged "DateFormat"
    DisplayDate
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,Short Time
Public Property Get TimeFormat() As String
Attribute TimeFormat.VB_ProcData.VB_Invoke_Property = ";Data"
    TimeFormat = m_TimeFormat
End Property

Public Property Let TimeFormat(ByVal New_TimeFormat As String)
    m_TimeFormat = New_TimeFormat
    PropertyChanged "TimeFormat"
    DisplayDate
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_DateFormat = m_def_DateFormat
    m_TimeFormat = m_def_TimeFormat
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    txtDate.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    txtDate.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set txtDate.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    m_DateFormat = PropBag.ReadProperty("DateFormat", m_def_DateFormat)
    m_TimeFormat = PropBag.ReadProperty("TimeFormat", m_def_TimeFormat)
    txtDate.Locked = PropBag.ReadProperty("Locked", False)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", txtDate.BackColor, &H80000005)
    Call PropBag.WriteProperty("ForeColor", txtDate.ForeColor, &H80000008)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", txtDate.Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("DateFormat", m_DateFormat, m_def_DateFormat)
    Call PropBag.WriteProperty("TimeFormat", m_TimeFormat, m_def_TimeFormat)
    Call PropBag.WriteProperty("Locked", txtDate.Locked, False)
End Sub

Private Sub DisplayDate()
    Dim strDate As String
    
    If Trim(m_DateFormat) <> "" Then
        If CDate(Format(m_TimeStamp, "dd-mmm-yyyy")) = mNullDate Then
            txtDate = ""
            Exit Sub
        End If
        strDate = Format(m_TimeStamp, m_DateFormat)
    End If
    
    
    If m_TimeFormat <> "" And m_DateFormat <> "" Then
        strDate = strDate & " "
    End If
    
    If m_TimeFormat <> "" Then
        If CDate(Format(m_TimeStamp, m_TimeFormat)) <> mNullDate Then
            strDate = strDate & Format(m_TimeStamp, m_TimeFormat)
        End If
    End If
    txtDate = strDate
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtDate,txtDate,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = txtDate.BorderStyle
End Property
'
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=3,0,2,0
Public Property Get TimeStamp() As Date
Attribute TimeStamp.VB_ProcData.VB_Invoke_Property = ";Data"
Attribute TimeStamp.VB_UserMemId = 0
Attribute TimeStamp.VB_MemberFlags = "400"
    TimeStamp = m_TimeStamp
End Property

Public Property Let TimeStamp(ByVal New_TimeStamp As Date)
    If Ambient.UserMode = False Then Err.Raise 387
    m_TimeStamp = New_TimeStamp
    PropertyChanged "TimeStamp"
    DisplayDate
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtDate,txtDate,-1,Locked
Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Determines whether a control can be edited."
    Locked = txtDate.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    txtDate.Locked() = New_Locked
    PropertyChanged "Locked"
End Property

