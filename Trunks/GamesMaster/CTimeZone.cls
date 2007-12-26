VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTimeZone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' *************************************************************
'  Copyright ©1998-2004, Karl E. Peterson
'  http://www.mvps.org/vb/
' *************************************************************
'  Author grants royalty-free rights to use this code within
'  compiled applications. Selling or otherwise distributing
'  this source code is not allowed without author's express
'  permission.
' *************************************************************
Option Explicit

Private Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long

Private Type SYSTEMTIME
   wYear As Integer
   wMonth As Integer
   wDayOfWeek As Integer
   wDay As Integer
   wHour As Integer
   wMinute As Integer
   wSecond As Integer
   wMilliseconds As Integer
End Type

Private Type TIME_ZONE_INFORMATION
   Bias As Long
   StandardName(0 To 63) As Byte
   StandardDate As SYSTEMTIME
   StandardBias As Long
   DaylightName(0 To 63) As Byte
   DaylightDate As SYSTEMTIME
   DaylightBias As Long
End Type

Private Const TIME_ZONE_ID_INVALID As Long = &HFFFFFFFF
Private Const TIME_ZONE_ID_UNKNOWN As Long = 0
Private Const TIME_ZONE_ID_STANDARD As Long = 1
Private Const TIME_ZONE_ID_DAYLIGHT As Long = 2

' Member variables
Private m_TZI As TIME_ZONE_INFORMATION
Private m_rng As tzCurrentRanges
Private m_err As Long

' Enumerations
Public Enum tzCurrentRanges
   tzUnknown = TIME_ZONE_ID_UNKNOWN
   tzStandard = TIME_ZONE_ID_STANDARD
   tzDaylight = TIME_ZONE_ID_DAYLIGHT
   tzInvalid = TIME_ZONE_ID_INVALID
End Enum

' ************************************************
'  Initialization and Termination
' ************************************************
Private Sub Class_Initialize()
   Me.Refresh
End Sub

' ************************************************
'  Properties (read-only)
' ************************************************
Public Property Get Bias() As Long
   ' All the bias logic wrapped into one property.
   Select Case m_rng
      Case tzUnknown
         Bias = m_TZI.Bias
      Case tzDaylight
         Bias = m_TZI.Bias + m_TZI.DaylightBias
      Case tzStandard
         Bias = m_TZI.Bias + m_TZI.StandardBias
   End Select
End Property

Public Property Get BiasDaylight() As Long
   ' This value is added to the value of the Bias member
   ' to form the bias used during daylight saving time.
   ' In most time zones, the value of this member is –60.
   If m_rng <> tzInvalid Then
      BiasDaylight = m_TZI.DaylightBias
   End If
End Property

Public Property Get BiasStandard() As Long
   ' This value is added to the value of the Bias member
   ' to form the bias used during standard time. In most
   ' time zones, the value of this member is zero.
   If m_rng <> tzInvalid Then
      BiasStandard = m_TZI.StandardBias
   End If
End Property

Public Property Get BiasUtc() As Long
   ' All translations between UTC time and local time
   ' are based on the following formula:
   '    UTC = local time + bias
   ' The bias is the difference, in minutes, between
   ' UTC time and local time.
   If m_rng <> tzInvalid Then
      BiasUtc = m_TZI.Bias
   End If
End Property

Public Property Get CurrentRange() As tzCurrentRanges
   ' Determined with call to GetTimeZoneInformation
   Me.Refresh
   CurrentRange = m_rng
End Property

Public Property Get CurrentRangeName() As String
   Select Case m_rng
      Case tzUnknown
         CurrentRangeName = "The system cannot determine the current time zone."
      Case tzStandard
         CurrentRangeName = "Standard Time"
      Case tzDaylight
         CurrentRangeName = "Daylight Savings Time"
      Case tzInvalid
         CurrentRangeName = "System Error (" & m_err & ")."
   End Select
End Property

Public Property Get DateDaylight() As Date
   If m_rng <> tzInvalid Then
      DateDaylight = tzDate(m_TZI.DaylightDate)
   End If
End Property

Public Property Get DateStandard() As Date
   If m_rng <> tzInvalid Then
      DateStandard = tzDate(m_TZI.StandardDate)
   End If
End Property

Public Property Get TimeLocal() As Date
   TimeLocal = Now
End Property

Public Property Get TimeEMail(Optional ByVal dtDate = 0) As String
    If dtDate = 0 Then
        dtDate = Now
    End If
    TimeEMail = Format(dtDate, "ddd, dd mmm yyyy hh:nn:ss") & _
                " " & Me.Offset
End Property

Public Property Get TimeUtc() As Date
   If m_rng <> tzInvalid Then
      TimeUtc = DateAdd("n", Me.Bias, Now)
   End If
End Property

Public Property Get NameDaylight() As String
   If m_rng <> tzInvalid Then
      NameDaylight = TrimNull(m_TZI.DaylightName)
   End If
End Property

Public Property Get Offset() As String
   Dim sRet As String
   Dim diff As Long
   ' Provide offset from UTC in format typically used
   ' in email and newsgroup post headers.
   If m_rng <> tzInvalid Then
      diff = Me.Bias
      If diff > 0 Then
         sRet = "-"
      Else
         sRet = "+"
         diff = -diff  'Avoid Format$() including sign.
      End If
      sRet = sRet & Format$(diff \ 60, "00")
      sRet = sRet & Format$(diff Mod 60, "00")
   End If
   Offset = sRet
End Property

Public Property Get NameStandard() As String
   If m_rng <> tzInvalid Then
      NameStandard = TrimNull(m_TZI.StandardName)
   End If
End Property

' ************************************************
'  Public Methods
' ************************************************
Public Sub Refresh()
   Dim tz As TIME_ZONE_INFORMATION
   ' Clear existing info, and collect new.
   m_TZI = tz
   m_rng = GetTimeZoneInformation(m_TZI)
   If m_rng = TIME_ZONE_ID_INVALID Then
      ' This shouldn't happen, but could be useful
      ' in debugging on oddball systems.
      m_err = Err.LastDllError
   End If
End Sub

' ************************************************
'  Private Methods
' ************************************************
Private Function tzDate(st As SYSTEMTIME) As Date
   Dim i As Long
   Dim n As Long
   Dim d1 As Long
   Dim d2 As Long
   
   ' This member supports two date formats. Absolute format
   ' specifies an exact date and time when standard time
   ' begins. In this form, the wYear, wMonth, wDay, wHour,
   ' wMinute, wSecond, and wMilliseconds members of the
   ' SYSTEMTIME structure are used to specify an exact date.
   If st.wYear Then
      tzDate = _
         DateSerial(st.wYear, st.wMonth, st.wDay) + _
         TimeSerial(st.wHour, st.wMinute, st.wSecond)
   
   ' Day-in-month format is specified by setting the wYear
   ' member to zero, setting the wDayOfWeek member to an
   ' appropriate weekday, and using a wDay value in the
   ' range 1 through 5 to select the correct day in the
   ' month. Using this notation, the first Sunday in April
   ' can be specified, as can the last Thursday in October
   ' (5 is equal to "the last").
   Else
      ' Get first day of month
      d1 = DateSerial(Year(Now), st.wMonth, 1)
      ' Get last day of month
      d2 = DateSerial(Year(d1), st.wMonth + 1, 0)
      
      ' Match weekday with appropriate week...
      If st.wDay = 5 Then
         ' Work backwards
         For i = d2 To d1 Step -1
            If Weekday(i) = (st.wDayOfWeek + 1) Then
               Exit For
            End If
         Next i
      Else
         ' Start at 1st and work forward
         For i = d1 To d2
            If Weekday(i) = (st.wDayOfWeek + 1) Then
               n = n + 1  'incr week value
               If n = st.wDay Then
                  Exit For
               End If
            End If
         Next i
      End If
      
      ' Got the serial date!  Just format it and
      ' add in the appropriate time.
      tzDate = i + _
         TimeSerial(st.wHour, st.wMinute, st.wSecond)
   End If
End Function

Private Function TrimNull(ByVal StrIn As String) As String
   Dim nul As Long
   ' Truncate input string at first null.
   ' If no nulls, perform ordinary Trim.
   nul = InStr(StrIn, vbNullChar)
   Select Case nul
      Case Is > 1
         TrimNull = Left(StrIn, nul - 1)
      Case 1
         TrimNull = ""
      Case 0
         TrimNull = Trim(StrIn)
   End Select
End Function
