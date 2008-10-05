VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMap 
   Caption         =   "Map"
   ClientHeight    =   6330
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8730
   LinkTopic       =   "Form1"
   ScaleHeight     =   422
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   582
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar sbMap 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   4
      Top             =   6030
      Width           =   8730
      _ExtentX        =   15399
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7885
            MinWidth        =   1323
            Key             =   "General"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Text            =   "planet name"
            TextSave        =   "planet name"
            Key             =   "Planet"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1323
            MinWidth        =   1323
            Text            =   "Size: 999"
            TextSave        =   "Size: 999"
            Key             =   "Size"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2223
            MinWidth        =   1323
            Text            =   "Resources: 9.99"
            TextSave        =   "Resources: 9.99"
            Key             =   "Resources"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   794
            MinWidth        =   794
            Text            =   "(x,y)"
            TextSave        =   "(x,y)"
            Key             =   "Position"
         EndProperty
      EndProperty
   End
   Begin VB.HScrollBar HScroll 
      Height          =   255
      Left            =   420
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
   Begin VB.VScrollBar vScroll 
      Height          =   1035
      Left            =   3060
      TabIndex        =   2
      Top             =   660
      Width           =   255
   End
   Begin VB.PictureBox picOuter 
      Height          =   2055
      Left            =   180
      ScaleHeight     =   133
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   165
      TabIndex        =   1
      Top             =   180
      Width           =   2535
      Begin VB.PictureBox picInner 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1275
         Left            =   240
         ScaleHeight     =   85
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   109
         TabIndex        =   0
         ToolTipText     =   "Zoom In=Shift-Click Out=Ctrl-Click"
         Top             =   240
         Width           =   1635
      End
   End
End
Attribute VB_Name = "frmMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************
'   Copyright 2007,2008 Ian Evans.                      *
'   This program is distributed under the terms of the  *
'       GNU General Public License.                     *
'********************************************************
Option Explicit

Private msngGalaxySize As Single
Private msngZoom As Single
Private mcolPlanets As Collection
Private maPlanets() As utPlanet
Private msngXDown As Single
Private msngYDown As Single
Private mblnResizing As Boolean

Private Type utPlanet
    Name As String
    Race As String
    Index As Long
    X As Single
    Y As Single
    Size As Single
    Resources As Single
End Type


Private Sub Form_Load()
    Dim H As Single, W As Single, S As Single
    
    Set Me.Icon = MainForm.Icon
    msngZoom = 1
    msngGalaxySize = 100
    H = Me.ScaleWidth - vScroll.Width - (picOuter.Width - picOuter.ScaleWidth)
    W = Me.ScaleHeight - HScroll.Height - (picOuter.Height - picOuter.ScaleHeight) - sbMap.Height
    If H < GalaxySize Then H = GalaxySize
    If W < GalaxySize Then W = GalaxySize
    If H < W Then
        S = H
    Else
        S = W
    End If
    picInner.Move 0, 0, S, S
    vScroll.Min = 0
    vScroll.Max = 1000
    vScroll.Value = 0
    vScroll.SmallChange = 10
    HScroll.Min = 0
    HScroll.Max = 1000
    HScroll.Value = 0
    HScroll.SmallChange = 100
    picInner.Line (0, 0)-(S, S)
    picInner.Line (S, 0)-(0, S)
    Call LoadFormSettings(Me)
End Sub

Public Property Get Planets() As Collection
    Set Planets = mcolPlanets
End Property

Public Property Set Planets(ByVal colPlanets As Collection)
    Set mcolPlanets = colPlanets
    Call LoadPlanets
    Call DrawPlanets
End Property

Public Property Get GalaxySize() As Single
    GalaxySize = msngGalaxySize
End Property

Public Property Let GalaxySize(sngGalaxySize As Single)
    msngGalaxySize = sngGalaxySize
End Property

Public Property Get Zoom() As Single
    Zoom = msngZoom
End Property

Public Property Let Zoom(sngZoom As Single)
    msngZoom = sngZoom
End Property

Private Sub Form_Resize()
    Dim X As Single, Y As Single, W As Single, H As Single
    Dim S As Single, L As Single, T As Single
        
    mblnResizing = True
    'Position the outer box
    W = Me.ScaleWidth - vScroll.Width
    H = Me.ScaleHeight - HScroll.Height - sbMap.Height
    If W < 0 Then W = 0
    If H < 0 Then H = 0
    picOuter.Move 0, 0, W, H
    
    If Zoom = 1 Then
        If picOuter.ScaleWidth < picOuter.ScaleHeight Then
            S = picOuter.ScaleWidth
        Else
            S = picOuter.ScaleHeight
        End If
    Else
        S = picInner.Width
    End If
    
    'Check the Inner Box position incase of Resize
    L = picInner.Left
    If S < picOuter.ScaleWidth Then
        L = 0
    ElseIf L + S < picOuter.ScaleWidth Then
        L = picOuter.ScaleWidth - S
    End If
    
    T = picInner.Top
    If S < picOuter.ScaleHeight Then
        T = 0
    ElseIf T + S < picOuter.ScaleHeight Then
        T = picOuter.ScaleHeight - S
    End If
    picInner.Move L, T, S, S
    
    On Error Resume Next
    'Position the Verticle Scroll Bar
    vScroll.Visible = False
    vScroll.Move W, 0, vScroll.Width, H
    Y = picInner.Height - picOuter.ScaleHeight
    If picInner.Top = 0 And Y <= 0 Then
        vScroll.Enabled = False
    Else
        vScroll.LargeChange = picOuter.ScaleHeight / Y * vScroll.Max
        vScroll.Enabled = True
        vScroll.Value = -T / Y * vScroll.Max
        vScroll.Visible = True
    End If
    
    'Position the Hosizontal scroll bar
    HScroll.Visible = False
    HScroll.Move 0, H, W
    X = picInner.Width - picOuter.ScaleWidth
    If picInner.Left = 0 And X <= 0 Then
        HScroll.Enabled = False
    Else
        HScroll.LargeChange = picOuter.ScaleWidth / X * HScroll.Max
        HScroll.Enabled = True
        HScroll.Value = -L / X * HScroll.Max
        HScroll.Visible = True
    End If
    On Error GoTo 0
    
    mblnResizing = False
    On Error Resume Next
    Call DrawPlanets
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveFormSettings(Me)
End Sub

Private Sub HScroll_Change()
    If mblnResizing Then Exit Sub
    With picInner
        .Move -HScroll.Value / HScroll.Max * (picInner.Width - picOuter.ScaleWidth)
    End With
End Sub

Private Sub HScroll_Scroll()
    Call HScroll_Change
End Sub

Private Sub picInner_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Shift = 0 Then
        picInner.MousePointer = vbSizeAll
        msngXDown = X
        msngYDown = Y
    Else
        msngXDown = -1
        msngYDown = -1
    End If
End Sub

Private Sub picInner_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim L As Single, T As Single
    
    If msngXDown > 0 Then
        ' Scroll
        L = picInner.Left + X - msngXDown
        T = picInner.Top + Y - msngYDown
        If L + picInner.Width < picOuter.ScaleWidth Then
            L = picOuter.ScaleWidth - picInner.Width
        End If
        If T + picInner.Height < picOuter.ScaleHeight Then
            T = picOuter.ScaleHeight - picInner.Height
        End If
        If L > 0 Then L = 0
        If T > 0 Then T = 0
        picInner.Move L, T
        
        If vScroll.Enabled Then
            vScroll.Value = -T / (picInner.Height - picOuter.ScaleHeight) * vScroll.Max
        End If
        If HScroll.Enabled Then
            HScroll.Value = -L / (picInner.Width - picOuter.ScaleWidth) * HScroll.Max
        End If
    End If
    
    'Set the Status bar information
    Call SetPosition(X, Y)
End Sub

Private Sub picInner_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim L As Single, T As Single, S As Single, Z1 As Long
    Dim X1 As Single, Y1 As Single
    Dim f As Single
    
    With picInner
        If msngXDown > 0 Then
            picInner.MousePointer = vbDefault
            'Cancel Scrolling
            msngXDown = -1
            msngYDown = -1
            Exit Sub
        End If
        
        If Button = vbLeftButton Then
            X1 = X + .Left
            Y1 = Y + .Top
            If (Shift And vbShiftMask) = vbShiftMask Then
                Z1 = Zoom * 2
                f = Z1 / Zoom
            ElseIf (Shift And vbCtrlMask) = vbCtrlMask Then
                Z1 = Zoom / 2
                If Z1 < 1 Then Z1 = 1
                f = Z1 / Zoom
            Else
                Z1 = Zoom
                f = 1
            End If
            
            'Calculate the Zoom
            L = -X * f + X1
            T = -Y * f + Y1
            If L > 0 Then L = 0
            If T > 0 Then T = 0
            S = .Width * f
            Zoom = Z1
            
            'Check the new Inner box position
            If S < picOuter.ScaleWidth Then
                L = 0
            ElseIf L + S < picOuter.ScaleWidth Then
                L = picOuter.ScaleWidth - S
            End If
           
            If S < picOuter.ScaleHeight Then
                T = 0
            ElseIf T + S < picOuter.ScaleHeight Then
                T = picOuter.ScaleHeight - S
            End If
            
            .Move L, T, S, S
            Call Form_Resize
        End If
    End With
End Sub

Private Sub picInner_Paint()
    Call DrawPlanets
End Sub

Private Sub vScroll_Change()
    If mblnResizing Then Exit Sub
    With picInner
        .Move .Left, -vScroll.Value / vScroll.Max * (picInner.Height - picOuter.ScaleHeight)
    End With
End Sub

Private Sub vScroll_Scroll()
    Call vScroll_Change
End Sub

Private Sub LoadPlanets()
    Dim objHP As Object
    Dim p As Long
    Dim uPlanet As utPlanet
    
    ReDim maPlanets(100)
    For Each objHP In Planets
        p = p + 1
        uPlanet.Index = p
        With objHP
            uPlanet.Name = .Planet
            If TypeName(objHP) = "Planet" Then
                If .Owner Is Nothing Then
                    uPlanet.Race = ""
                Else
                    uPlanet.Race = .Owner.RaceName
                End If
            Else
                uPlanet.Race = .Owner
            End If
            uPlanet.X = .X
            uPlanet.Y = .Y
            uPlanet.Size = .Size
            uPlanet.Resources = .Resources
        End With
        
        If p > UBound(maPlanets) Then
            ReDim Preserve maPlanets(p + 99)
        End If
        maPlanets(p) = uPlanet
    Next objHP
        
    ReDim Preserve maPlanets(p)
    Call SortPlanets
End Sub

Private Sub DrawPlanets()
    Dim p As Long
    Dim S As Single
    Dim X As Single, Y As Single, R As Single
    Dim X1 As Single, Y1 As Single, H As Single, W As Single
    
    With picInner
        .Cls
        .FillStyle = vbFSSolid
        .BackColor = vbBlack
    End With
    For p = 1 To UBound(maPlanets)
        With maPlanets(p)
            If .Race = "" Then
                picInner.ForeColor = vbWhite
                picInner.FillColor = vbWhite
            Else
                picInner.ForeColor = vbYellow
                picInner.FillColor = vbYellow
            End If
            X = Xpos(.X)
            Y = Ypos(.Y)
            R = Rsize(.Size)
            picInner.Circle (X, Y), R, vbBlack
            With picInner
                .FontSize = 8
                W = .TextWidth(maPlanets(p).Name)
                H = .TextHeight(maPlanets(p).Name)
                Y1 = Y + R + 1
                If Y1 + H > .ScaleHeight Then
                    Y1 = Y + R - H - 1
                End If
                X1 = X - (W / 2)
                If X1 < 0 Then
                    X1 = 0
                End If
                If X1 + W > .ScaleWidth Then
                    X1 = .ScaleWidth - W
                End If
                .CurrentX = X1
                .CurrentY = Y1
                .ForeColor = vbCyan
                picInner.Print maPlanets(p).Name
            End With
        End With
    Next p
    
End Sub

Private Function Xpos(ByVal X As Single) As Single
    Xpos = Fix(X / GalaxySize * picInner.ScaleWidth)
End Function

Private Function Ypos(ByVal Y As Single) As Single
    Ypos = Fix(Y / GalaxySize * picInner.ScaleHeight)
End Function

Private Function Rsize(ByVal R As Single) As Single
    Select Case R
    Case Is < 251
        R = 1
    Case Is < 351
        R = 2
    Case Is < 600
        R = 4
    Case Is < 1000
        R = 6
    Case Is < 1500
        R = 8
    Case Else
        R = 10
    End Select
        
    Rsize = R
End Function

Private Sub SortPlanets()
    Dim uHPlanet As utPlanet
    Dim T As Long
    Dim p As Long
    Dim P1 As Long
    
    
    For T = 2 To UBound(maPlanets)
        p = T
        uHPlanet = maPlanets(p)
        For p = T To 2 Step -1
            P1 = p - 1
            If uHPlanet.Size > maPlanets(P1).Size Then
                maPlanets(p) = maPlanets(P1)
            Else
                Exit For
            End If
        Next p
        maPlanets(p) = uHPlanet
    Next T
End Sub

Private Sub SetPosition(ByVal X As Single, ByVal Y As Single)
    Dim i As Long
    Dim objPlanet As Planet
    

    sbMap.Panels("Position") = "(" & _
                            Format(X / picInner.ScaleWidth * GalaxySize, "0.00") & _
                            "," & _
                            Format(Y / picInner.ScaleHeight * GalaxySize, "0.00") & _
                            ")"
    i = PlanetIndex(X, Y)
    If i = 0 Then
        sbMap.Panels("Planet") = ""
        sbMap.Panels("Size") = ""
        sbMap.Panels("Resources") = ""
        sbMap.Panels("General") = ""
    Else
        sbMap.Panels("Planet") = maPlanets(i).Name
        sbMap.Panels("Size") = "Size: " & Format(maPlanets(i).Size, "0")
        sbMap.Panels("Resources") = "Resources: " & Format(maPlanets(i).Resources, "0.0")
        On Error Resume Next
        Set objPlanet = Planets(maPlanets(i).Index)
        If objPlanet Is Nothing Then
            sbMap.Panels("General") = maPlanets(i).Race
        ElseIf objPlanet.Pop > 0 Then
            sbMap.Panels("General") = maPlanets(i).Race & _
                                    ": P=" & Format(objPlanet.Pop, "0") & _
                                    ", I=" & Format(objPlanet.Ind, "0")
        Else
            sbMap.Panels("General") = ""
        End If
    End If
End Sub

Private Function PlanetIndex(ByVal X As Single, Y As Single) As Long
    Dim i As Long
    Dim i1 As Long, d1 As Single
    Dim dx As Single, dy As Single
    Dim d As Single
    i1 = 0
    
    For i = LBound(maPlanets) + 1 To UBound(maPlanets)
        dx = Abs(Xpos(maPlanets(i).X) - X)
        dy = Abs(Xpos(maPlanets(i).Y) - Y)
        If dx <= 10 _
        And dy <= 10 Then
            d = dx * dx + dy * dy
            If i1 = 0 _
            Or d < d1 Then
                i1 = i
                d1 = d
            End If
        End If
    Next i
    PlanetIndex = i1
End Function
