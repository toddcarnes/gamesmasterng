VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmTemplates 
   Caption         =   "Templates"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdTemplates 
      Height          =   1455
      Left            =   540
      TabIndex        =   0
      Top             =   780
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   2566
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frmTemplates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private mobjTemplates As Templates

Public Property Get Templates() As Templates
    Set Templates = mobjTemplates
End Property

Public Property Set Templates(ByVal objTemplates As Templates)
    Set mobjTemplates = objTemplates
End Property

Public Sub Load()
    Dim lngRow As Long
    Dim objTemplate As Template
    
    With grdTemplates
        .Clear
        .FixedCols = 1
        .FixedRows = 1
        .Rows = 2
        .Cols = 4
        .FocusRect = flexFocusNone
        .SelectionMode = flexSelectionByRow
    
        .ColWidth(0) = 16 * Screen.TwipsPerPixelX
        .ColWidth(1) = 2000
        .ColWidth(2) = 800
        .ColAlignment(2) = flexAlignCenterTop
        .ColWidth(3) = 800
        .ColAlignment(3) = flexAlignCenterTop
        .TextMatrix(0, 1) = "Name"
        .TextMatrix(0, 2) = "Turn"
        .TextMatrix(0, 3) = "Players"
    
        lngRow = 0
        For Each objTemplate In Templates
            lngRow = lngRow + 1
            If lngRow > .Rows Then .Rows = lngRow
            .TextMatrix(lngRow, 1) = objTemplate.TemplateName
        Next objTemplate
    End With
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    With grdTemplates
        .Top = 0
        .Left = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight
    End With
End Sub

