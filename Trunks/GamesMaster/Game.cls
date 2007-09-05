VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Game"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"Races"
Attribute VB_Ext_KEY = "Member1" ,"Template"
Option Explicit

Public GameName As String
Public TemplateFile As String
Public VersionString As String
Public VersionNumber As Single
Public VersionDate As String
Public Turn As Long
Public Flags As Long
Public GalacticPeace As Long
Public InitialDrive As Single
Public InitialWeapons As Single
Public InitialShield As Single
Public InitialCargo As Single
Public GalaxySize As Single
Public Created As Boolean
Private mlngNextTurn As Long
Private mvarRaces As Races
Private mobjTemplate As Template

'Game Options Flags
Public Enum GameFlag
    G_NONGBOMBING = 1       'planet is bombed completely and does not change owners
    G_KEEPPRODUCTION = 2    'production already spent is not lost after a switch of production type
    G_CREATECIRCLE = 4      'Create systems along the edge of a circle.
    G_NODROP = 8            'Do not drop idle players
    G_SAVECOPY = 16         'Save a copy of the turn report in /report when running a turn
    G_SPHERICALGALAXY = 32  'Galaxy is mapped on a sphere
End Enum

Public Property Get Flag(ByVal uFlag As GameFlag) As Boolean
    Flag = ((Flags And uFlag) = uFlag)
End Property

Public Property Get Template() As Template
    If mobjTemplate Is Nothing _
    And TemplateFile <> "" Then
        Set mobjTemplate = New Template
        Call mobjTemplate.Load(TemplateFile)
    End If
    Set Template = mobjTemplate
End Property

Public Property Get Races() As Races
    Set Races = mvarRaces
End Property

Public Function GameFile(Optional ByVal Turn As Long = -1)
    If Turn < 0 Then Turn = NextTurn - 1
    GameFile = GalaxyNGData & GameName & "\" & CStr(Turn) & ".New"
End Function

Public Property Get NextTurn() As Long
    Dim intFN As Integer
    If Created Then
        If mlngNextTurn < 0 Then
            intFN = FreeFile
            Open GalaxyNGNextTurn(GameName) For Input As intFN
            Input #intFN, mlngNextTurn
            Close intFN
        End If
    End If
    NextTurn = mlngNextTurn
End Property

Public Sub Refresh()
    Dim intFN As Integer
    Dim lngSize As Long
    Dim strBuffer As String
    Dim vLines As Variant
    Dim l As Long
    Dim objRace As Race
    Dim i As Long
    
    Set Races = New Races
    
    If Not Created Then Exit Sub
    '------------------------------------------------------------
    'Load the game file into an array
    lngSize = FileLen(GameFile)
    strBuffer = String(lngSize, ".")
    
    intFN = FreeFile
    Open GameFile For Binary As #intFN
    Get #intFN, , strBuffer
    Close #intFN
    
    vLines = Split(strBuffer, vbCrLf)
    l = 0
    
    ' GalaxyNG Version
    VersionString = vLines(l)
    i = InStr(1, VersionString, "release-")
    VersionNumber = Replace(Mid(VersionString, i + 8, 3), "-", ".")
    i = InStr(1, VersionString, ", ")
    VersionDate = Mid(VersionString, i + 2)
    
    ' Turn Number
    l = l + 1
    Turn = vLines(l)
    
    ' Galaxy Size
    l = l + 1
    GalaxySize = vLines(l)
    
    '------------------------------------------------------------
    'Find The start of the game options
    Do
        If Left(vLines(l), 12) = "@GameOptions" Then Exit Do
        l = l + 1
    Loop
    
    ' Game Options
    l = l + 1
    Flags = 0 & vLines(l)
    
    ' Galactic Peace
    l = l + 1
    GalacticPeace = 0 & vLines(l)
    
    ' Initial Drive
    l = l + 1
    InitialDrive = vLines(l)
    
    ' Initial Weapons
    l = l + 1
    InitialWeapons = vLines(l)
    
    ' Initial Shields
    l = l + 1
    InitialShield = vLines(l)
    
    ' Initial Cargo
    l = l + 1
    InitialCargo = vLines(l)
    
    '------------------------------------------------------------
    'Find The start of the Players section
    Do
        If Left(vLines(l), 8) = "@Players" Then Exit Do
        l = l + 1
    Loop
    Do
        '------------------------------------------------------------
        ' Get the Player Details
        l = l + 1
        If Left(vLines(l), 9) = "@EPlayers" Then Exit Do
        Set objRace = New Race
        objRace.RaceName = vLines(l)
        l = l + 1
        objRace.EMail = vLines(l)
        l = l + 1
        objRace.Password = vLines(l)
        l = l + 1
        objRace.Drive = vLines(l)
        l = l + 1
        objRace.Weapons = vLines(l)
        l = l + 1
        objRace.Shields = vLines(l)
        l = l + 1
        objRace.Cargo = vLines(l)
        l = l + 1
        objRace.mX = 0 & vLines(l)
        l = l + 1
        objRace.mY = 0 & vLines(l)
        l = l + 1
        objRace.mSize = 0 & vLines(l)
        l = l + 1
        objRace.PlayerName = vLines(l)
        l = l + 1
        objRace.TeamName = vLines(l)
        l = l + 1
        objRace.Unused3 = vLines(l)
        l = l + 1
        objRace.Unused4 = vLines(l)
        l = l + 1
        objRace.Unused5 = vLines(l)
        l = l + 1
        objRace.MassLost = vLines(l)
        l = l + 1
        objRace.MassProduced = vLines(l)
        l = l + 1
        objRace.LastOrders = vLines(l)
        l = l + 1
        objRace.Flags = vLines(l)
        
        '------------------------------------------------------------
        ' Find the Player Ship types
        Do
            If Left(vLines(l), 10) = "@ShipTypes" Then Exit Do
            l = l + 1
        Loop
        'Ship Type Name
        'drive
        'attacks
        'weapons
        'shields
        'cargo
        
        '------------------------------------------------------------
        ' Find the Player Fleets
        Do
            If Left(vLines(l), 7) = "@Fleets" Then Exit Do
            l = l + 1
        Loop
        'Fleet Name
        
        '------------------------------------------------------------
        ' Find the last line for the Player
        Do
            If Left(vLines(l), 8) = "@EFleets" Then Exit Do
            l = l + 1
        Loop
        Races.Add objRace, objRace.RaceName
    Loop
    
End Sub

Public Property Get ActivePlayers() As Long
    Dim objRace As Race
    
    ActivePlayers = 0
    For Each objRace In mvarRaces
        If Not objRace.Flag(R_DEAD) Then ActivePlayers = ActivePlayers + 1
    Next objRace
End Property

Public Property Set Races(vData As Races)
    Set mvarRaces = vData
End Property

Private Sub Class_Initialize()
    mlngNextTurn = -1
    Turn = -1
End Sub

Private Sub Class_Terminate()
    Set mvarRaces = Nothing
    Set mobjTemplate = Nothing
End Sub


