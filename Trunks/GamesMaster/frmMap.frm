VERSION 5.00
Begin VB.Form frmMap 
   Caption         =   "Form1"
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
Private mcolPlanets As Planets
Private msngZoom As Single

Private Sub Form_Load()
    Dim H As Single, W As Single, S As Single
    
    msngZoom = 1
    msngGalaxySize = 100
    H = Me.ScaleWidth - vScroll.Width - (picOuter.Width - picOuter.ScaleWidth)
    W = Me.ScaleHeight - HScroll.Height - (picOuter.Height - picOuter.ScaleHeight)
    If H < GalaxySize Then H = GalaxySize
    If W < GalaxySize Then W = GalaxySize
    If H < W Then
        S = H
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

Public Property Get GalaxySize() As Single
    GalaxySize = msngGalaxySize
End Property

Public Property Let GalaxySize(sngGalaxySize As Single)
    sngGalaxySize = GalaxySize
End Property

Public Property Get Planets() As Planets
    Set Planets = mcolPlanets
End Property

Public Property Set Planets(colPlanets As Planets)
    Set mcolPlanets = colPlanets
End Property

Public Property Get Zoom() As Single
    Zoom = msngZoom
End Property

Public Property Let Zoom(sngZoom As Single)
    sngZoom = Zoom
End Property

Private Sub Form_Resize()
    Dim W As Single
    Dim H As Single
    Dim S As Single
    
    W = Me.ScaleWidth - vScroll.Width
    H = Me.ScaleHeight - HScroll.Height
    If W < 0 Then W = 0
    If H < 0 Then H = 0
    picOuter.Move 0, 0, W, H
    
    vScroll.Move W, 0, vScroll.Width, H
    S = picInner.Height - picOuter.ScaleHeight
    If S <= 0 Then
        S = 0
        vScroll.Enabled = False
    Else
        vScroll.Max = S
        vScroll.LargeChange = picOuter.ScaleHeight
        vScroll.Enabled = True
    End If
    
    HScroll.Move 0, H, W
    S = picInner.Width - picOuter.ScaleWidth
    If S <= 0 Then
        S = 0
        HScroll.Enabled = False
    Else
        HScroll.Max = S
        HScroll.LargeChange = picOuter.ScaleWidth
        HScroll.Enabled = True
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
