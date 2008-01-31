VERSION 5.00
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
   StartUpPosition =   1  'CenterOwner
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
      TabIndex        =   0
      Top             =   180
      Width           =   2535
      Begin VB.PictureBox picInner 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   1275
         Left            =   240
         ScaleHeight     =   85
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   109
         TabIndex        =   1
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
Option Explicit

Private msngGalaxySize As Single
Private msngZoom As Single
Private mcolPlanets As Collection
Private maPlanets() As utPlanet
Private msngXDown As Single
Private msngYDown As Single

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
    Dim h As Single, W As Single, S As Single
    
    Me.Icon = MainForm.Icon
    msngZoom = 1
    msngGalaxySize = 100
    h = Me.ScaleWidth - vScroll.Width - (picOuter.Width - picOuter.ScaleWidth)
    W = Me.ScaleHeight - HScroll.Height - (picOuter.Height - picOuter.ScaleHeight)
    If h < GalaxySize Then h = GalaxySize
    If W < GalaxySize Then W = GalaxySize
    If h < W Then
        S = h
    Else
        S = W
    End If
    picInner.Move 0, 0, S, S
    vScroll.Min = 0
    vScroll.Value = 0
'    vScroll.SmallChange = 10
    HScroll.Min = 0
    HScroll.Value = 0
    HScroll.SmallChange = 10
    picInner.Line (0, 0)-(S, S)
    picInner.Line (S, 0)-(0, S)
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
    Dim W As Single
    Dim h As Single
    Dim S As Single
    
    W = Me.ScaleWidth - vScroll.Width
    h = Me.ScaleHeight - HScroll.Height
    If W < 0 Then W = 0
    If h < 0 Then h = 0
    picOuter.Move 0, 0, W, h
    
    vScroll.Move W, 0, vScroll.Width, h
    S = picInner.Height - picOuter.ScaleHeight
    If picInner.Top = 0 And S <= 0 Then
        S = 0
        vScroll.Enabled = False
    Else
        vScroll.Max = picInner.ScaleHeight - picOuter.ScaleHeight
        vScroll.LargeChange = picOuter.ScaleHeight
        vScroll.Enabled = True
        vScroll.Value = -picInner.Top
    End If
    
    HScroll.Move 0, h, W
    S = picInner.Width - picOuter.ScaleWidth
    If picInner.Left = 0 And S <= 0 Then
        S = 0
        HScroll.Enabled = False
    Else
        HScroll.Max = picInner.ScaleWidth - picOuter.ScaleWidth
        HScroll.LargeChange = picOuter.ScaleWidth
        HScroll.Enabled = True
        HScroll.Value = -picInner.Left
    End If
    
    With picInner
        If .Left < 0 _
        And .Left + .Width < picOuter.ScaleWidth Then
            S = picOuter.ScaleWidth - .Width
            If S > 0 Then S = 0
            HScroll.Value = -S
        End If
        If .Top < 0 _
        And .Top + .Height < picOuter.ScaleHeight Then
            S = picOuter.ScaleHeight - .Height
            If S > 0 Then S = 0
            vScroll.Value = -S
        End If
    End With
    Call DrawPlanets
End Sub

Private Sub HScroll_Change()
    With picInner
        .Move -HScroll.Value
    End With
End Sub

Private Sub HScroll_Scroll()
    With picInner
        .Move -HScroll.Value
    End With
End Sub

Private Sub picInner_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Shift = 0 Then
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
        vScroll.Value = -T
        HScroll.Value = -L
    End If
End Sub

Private Sub picInner_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim L As Single, T As Single, S As Single, Z1 As Long
    Dim X1 As Single, Y1 As Single
    Dim f As Single
    
    With picInner
        If msngXDown > 0 Then
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
            L = -X * f + X1
            T = -Y * f + Y1
            If L > 0 Then L = 0
            If T > 0 Then T = 0
            S = .Width * f
            Zoom = Z1
            picInner.Visible = False
            .Move L, T, S, S
            Call Form_Resize
            picInner.Visible = True
        End If
    End With
End Sub

Private Sub vScroll_Change()
    With picInner
        .Move .Left, -vScroll.Value
    End With
End Sub

Private Sub vScroll_Scroll()
    With picInner
        .Move .Left, -vScroll.Value
    End With
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
End Sub

Private Sub DrawPlanets()
    Dim p As Long
    Dim S As Single
    
    With picInner
        .Cls
        .ForeColor = vbWhite
        .FillColor = vbWhite
        .FillStyle = vbFSSolid
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
            picInner.Circle (Xpos(.X), Ypos(.Y)), Rsize(.Size)
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
        
    R = R / 100 + 1
    If R > 10 Then R = 10
    Rsize = R
End Function

