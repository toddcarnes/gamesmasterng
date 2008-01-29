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
Private Type utPlanet
    Name As String
    Race As String
    Index As Long
    x As Single
    Y As Single
    Size As Single
    Resources As Single
End Type


Private Sub Form_Load()
    Dim h As Single, W As Single, s As Single
    
    Me.Icon = MainForm.Icon
    msngZoom = 1
    msngGalaxySize = 100
    h = Me.ScaleWidth - vScroll.Width - (picOuter.Width - picOuter.ScaleWidth)
    W = Me.ScaleHeight - HScroll.Height - (picOuter.Height - picOuter.ScaleHeight)
    If h < GalaxySize Then h = GalaxySize
    If W < GalaxySize Then W = GalaxySize
    If h < W Then
        s = h
    Else
        s = W
    End If
    picInner.Move 0, 0, s, s
    vScroll.Min = 0
    vScroll.Value = 0
'    vScroll.SmallChange = 10
    HScroll.Min = 0
    HScroll.Value = 0
    HScroll.SmallChange = 10
    picInner.Line (0, 0)-(s, s)
    picInner.Line (s, 0)-(0, s)
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
    sngZoom = Zoom
End Property

Private Sub Form_Resize()
    Dim W As Single
    Dim h As Single
    Dim s As Single
    
    W = Me.ScaleWidth - vScroll.Width
    h = Me.ScaleHeight - HScroll.Height
    If W < 0 Then W = 0
    If h < 0 Then h = 0
    picOuter.Move 0, 0, W, h
    
    vScroll.Move W, 0, vScroll.Width, h
    s = picInner.Height - picOuter.ScaleHeight
    If s <= 0 Then
        s = 0
        vScroll.Enabled = False
    Else
        vScroll.Max = s
        vScroll.LargeChange = picOuter.ScaleHeight
        vScroll.Enabled = True
    End If
    
    HScroll.Move 0, h, W
    s = picInner.Width - picOuter.ScaleWidth
    If s <= 0 Then
        s = 0
        HScroll.Enabled = False
    Else
        HScroll.Max = s
        HScroll.LargeChange = picOuter.ScaleWidth
        HScroll.Enabled = True
    End If
    
    With picInner
        If .Left < 0 _
        And .Left + .Width < picOuter.ScaleWidth Then
            s = picOuter.ScaleWidth - .Width
            If s > 0 Then s = 0
            HScroll.Value = -s
        End If
        If .Top < 0 _
        And .Top + .Height < picOuter.ScaleHeight Then
            s = picOuter.ScaleHeight - .Height
            If s > 0 Then s = 0
            vScroll.Value = -s
        End If
    End With
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
    Dim P As Long
    Dim uPlanet As utPlanet
    
    ReDim maPlanets(100)
    For Each objHP In Planets
        P = P + 1
        uPlanet.Index = P
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
            uPlanet.x = .x
            uPlanet.Y = .Y
            uPlanet.Size = .Size
            uPlanet.Resources = .Resources
        End With
        
        If P > UBound(maPlanets) Then
            ReDim Preserve maPlanets(P + 99)
        End If
        maPlanets(P) = uPlanet
    Next objHP
        
    ReDim Preserve maPlanets(P)
End Sub

Private Sub DrawPlanets()
    Dim P As Long
    Dim i As Long
    
    picInner.Cls
    picInner.ForeColor = vbWhite
    For P = 1 To UBound(maPlanets)
        With maPlanets(P)
            For i = 0 To Rsize(.Size)
                picInner.Circle (Xpos(.x), Ypos(.Y)), i
            Next i
        End With
    Next P
    
End Sub

Private Function Xpos(ByVal x As Single) As Single
    Xpos = Fix(x / GalaxySize * picInner.ScaleWidth)
End Function

Private Function Ypos(ByVal Y As Single) As Single
    Ypos = picInner.ScaleHeight - Fix(Y / GalaxySize * picInner.ScaleHeight)
End Function

Private Function Rsize(ByVal R As Single) As Single
        
    R = R / 200 + 1
    If R > 10 Then R = 5
    Rsize = R
End Function

