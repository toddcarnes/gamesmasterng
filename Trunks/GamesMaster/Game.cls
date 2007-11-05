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
Attribute VB_Ext_KEY = "Member0" ,"Planets"
Attribute VB_Ext_KEY = "Member1" ,"Races"
Attribute VB_Ext_KEY = "Member2" ,"Template"
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
Private mcolRaces As Races
Private mobjTemplate As Template
Private mcolPlanets As Planets


'Game Options Flags
Public Enum GameFlag
    G_NONGBOMBING = 1       'planet is bombed completely and does not change owners
    G_KEEPPRODUCTION = 2    'production already spent is not lost after a switch of production type
    G_CREATECIRCLE = 4      'Create systems along the edge of a circle.
    G_NODROP = 8            'Do not drop idle players
    G_SAVECOPY = 16         'Save a copy of the turn report in /report when running a turn
    G_SPHERICALGALAXY = 32  'Galaxy is mapped on a sphere
End Enum

Public Property Get Planets() As Planets
    If mcolPlanets Is Nothing Then
        Set mcolPlanets = New Planets
    End If
    Set Planets = mcolPlanets
End Property

Public Property Set Planets(vData As Planets)
    Set mcolPlanets = vData
End Property

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
    Set Races = mcolRaces
End Property

Public Function GameFile(Optional ByVal Turn As Long = -1)
    If Turn < 0 Then Turn = NextTurn - 1
    GameFile = Options.GalaxyNGData & GameName & "\" & CStr(Turn) & ".New"
End Function

Public Property Get NextTurn() As Long
    Dim intFN As Integer
    If Created Then
        If mlngNextTurn < 0 Then
            intFN = FreeFile
            Open Options.GalaxyNGNextTurn(GameName) For Input As intFN
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
    Dim i As Long
    Dim objRace As Race
    Dim objShipType As ShipType
    Dim objPlanet As Planet
    Dim objGroup As Group
    
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
    If vLines(l) = "" Then
        Turn = 0
    Else
        Turn = vLines(l)
    End If
    
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
    ' Process the Player Details
    Do
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
        objRace.LastOrders = Val(vLines(l))
        l = l + 1
        objRace.Flags = vLines(l)
        
        '------------------------------------------------------------
        ' Find the Player Ship types
        Do
            l = l + 1
            If vLines(l) = "@ShipTypes" Then Exit Do
        Loop
        ' Process the Ship Types
        Do
            l = l + 1
            If vLines(l) = "@EShipTypes" Then Exit Do
            objShipType.ShipType = vLines(l)
            l = l + 1
            objShipType.Drive = vLines(l)
            l = l + 1
            objShipType.Attacks = vLines(l)
            l = l + 1
            objShipType.Weapons = vLines(l)
            l = l + 1
            objShipType.Shields = vLines(l)
            l = l + 1
            objShipType.Cargo = vLines(l)
            objRace.ShipTypes.Add objShipType
        Loop
        
        '------------------------------------------------------------
        ' Find the Player Fleets
        Do
            l = l + 1
            If vLines(l) = "@Fleets" Then Exit Do
        Loop
        ' Process the Player Fleets
        Do
            l = l + 1
            If vLines(l) = "@EFleets" Then Exit Do
            'Fleet Name
        Loop
        
        Races.Add objRace, objRace.RaceName
    Loop
    
    '------------------------------------------------------------
    ' Find the Allies
    Do
        l = l + 1
        If vLines(l) = "@Allies" Then Exit Do
    Loop
    ' Process the Allies
    Do
        l = l + 1
        If vLines(l) = "@EAllies" Then Exit Do
        'Alliance Details
    Loop
    
    '------------------------------------------------------------
    ' Find the Planets
    Do
        l = l + 1
        If vLines(l) = "@Planets" Then Exit Do
    Loop
    ' Process the Planets
    Do
        Set objPlanet = New Planet
        l = l + 1
        If vLines(l) = "@EPlanets" Then Exit Do
        objPlanet.Planet = vLines(l)
        l = l + 1
        objPlanet.OwnerID = vLines(l)
        Set objPlanet.Owner = Races(objPlanet.OwnerID)
        l = l + 1
        objPlanet.X = vLines(l)
        l = l + 1
        objPlanet.Y = vLines(l)
        l = l + 1
        objPlanet.Size = vLines(l)
        l = l + 1
        objPlanet.Resources = vLines(l)
        l = l + 1
        objPlanet.Pop = vLines(l)
        l = l + 1
        objPlanet.Ind = vLines(l)
        l = l + 1
        objPlanet.Producing = vLines(l)
        l = l + 1
        objPlanet.Cap = vLines(l)
        l = l + 1
        objPlanet.Mat = vLines(l)
        l = l + 1
        objPlanet.Col = vLines(l)
        l = l + 1
        objPlanet.InProgress = vLines(l)
        l = l + 1
        objPlanet.Spent = vLines(l)
        l = l + 1
        objPlanet.Flags = vLines(l)
        
        Planets.Add objPlanet
    Loop
    
    '------------------------------------------------------------
    ' Find the Routes
    Do
        l = l + 1
        If vLines(l) = "@Routes" Then Exit Do
    Loop
    
    ' Process the Routes
    Do
        l = l + 1
        If vLines(l) = "@ERoutes" Then Exit Do
        'Route Details
    Loop
    
    '------------------------------------------------------------
    ' Find the Player groups
    Do
        l = l + 1
        If vLines(l) = "@PGroups" Then Exit Do
    Loop
    
    ' Process the Player Groups
    Do
        Set objGroup = New Group
        l = l + 1
        If vLines(l) = "@EPGroups" Then Exit Do
        objGroup.ShipTypeID = vLines(l)
        l = l + 1
        objGroup.Number = vLines(l)
        l = l + 1
        objGroup.Drive = vLines(l)
        l = l + 1
        objGroup.Weapons = vLines(l)
        l = l + 1
        objGroup.Shields = vLines(l)
        l = l + 1
        objGroup.Cargo = vLines(l)
        l = l + 1
        objGroup.LoadType = vLines(l)
        l = l + 1
        objGroup.Load = vLines(l)
        l = l + 1
        objGroup.FromPlanetID = vLines(l)
        Set objGroup.FromPlanet = Planets(objGroup.FromPlanetID)
        l = l + 1
        objGroup.WherePlanetID = vLines(l)
        Set objGroup.WherePlanet = Planets(objGroup.WherePlanetID)
        l = l + 1
        objGroup.Dist = vLines(l)
        l = l + 1
        objGroup.Ships = vLines(l)
        l = l + 1
        objGroup.FleetID = vLines(l)
        objRace.Groups.Add objGroup
    Loop
    
    Do
        l = l + 1
        If l > UBound(vLines) Then Exit Do
        ' Save data at the end of the file
    Loop
        
End Sub

Public Property Get ActivePlayers() As Long
    Dim objRace As Race
    
    ActivePlayers = 0
    For Each objRace In mcolRaces
        If Not objRace.Flag(R_DEAD) Then ActivePlayers = ActivePlayers + 1
    Next objRace
End Property

Public Property Get PlayersReady() As Long
    Dim objRace As Race
    
    PlayersReady = 0
    For Each objRace In mcolRaces
        If Not objRace.Flag(R_DEAD) Then
            If OrdersReceived(objRace.RaceName) Then
                PlayersReady = PlayersReady + 1
            End If
        End If
    Next objRace
End Property

Public Property Set Races(vData As Races)
    Set mcolRaces = vData
End Property

Private Sub Class_Initialize()
    mlngNextTurn = -1
    Turn = -1
End Sub

Private Sub Class_Terminate()
    Set mcolPlanets = Nothing
    Set mcolRaces = Nothing
    Set mobjTemplate = Nothing
End Sub

Public Property Get Started() As Boolean
    Started = (Dir(Options.GalaxyNGReports & GameName & "\*.*") <> "")
End Property

Public Function ReadyToRun() As Boolean
    Dim dtNext As Date
    Dim objRace As Race
    Dim blnAllFinalOrdersReceived As Boolean
    
    dtNext = NextRunDate
    If dtNext <> 0 And dtNext < Now Then
        ReadyToRun = True
    
    ElseIf dtNext <> 0 And Template.FinalOrders Then
        blnAllFinalOrdersReceived = True
        For Each objRace In Races
            If Not objRace.Flag(R_DEAD) Then
                If Not FinalOrdersReceived(objRace.RaceName) Then
                    blnAllFinalOrdersReceived = False
                    Exit For
                End If
            End If
        Next objRace
        ReadyToRun = blnAllFinalOrdersReceived
    Else
        ReadyToRun = True
    End If
End Function

Public Function NotifyUsers() As Boolean
    Dim objRace As Race
    Dim strRace As String
    Dim dtNext As Date
    
    dtNext = NextRunDate
    If dtNext = 0 Then Exit Function
    
    ' Notify User 24 hours before turn
    dtNext = DateAdd("h", -24, dtNext)
    ' If not due to notify
    If dtNext > Now Then Exit Function
    
    ' Only check for 15 minutes
    If DateDiff("n", dtNext, Now) > 15 Then Exit Function
    
    For Each objRace In Races
        If Not objRace.Flag(R_DEAD) Then
            strRace = objRace.RaceName
            If OrdersReceived(strRace) Then
            ElseIf NotificationSent(strRace) Then
            Else
                NotifyUsers = True
                Exit For
            End If
        End If
    Next objRace
End Function

Public Function NotificationSent(ByVal strRace As String) As Boolean
    Dim strFileName As String
    
    strFileName = Options.GalaxyNGOrders & GameName & "\" & strRace & "_" & NextTurn & ".notify"
    If Dir(strFileName) <> "" Then
        NotificationSent = True
    End If
End Function

Public Function FinalOrdersReceived(ByVal strRace As String) As Boolean
    Dim strFileName As String
    
    strFileName = Options.GalaxyNGOrders & GameName & "\" & strRace & "_finalorders." & NextTurn
    
    If Dir(strFileName) <> "" Then
        FinalOrdersReceived = True
    End If
    
End Function

Public Function OrdersReceived(ByVal strRace As String) As Boolean
    Dim strFileName As String
    
    strFileName = Options.GalaxyNGOrders & GameName & "\" & strRace & "." & NextTurn
    If Dir(strFileName) <> "" Then
        OrdersReceived = True
    Else
        strFileName = Options.GalaxyNGOrders & GameName & "\" & strRace & "_final." & NextTurn
        If Dir(strFileName) <> "" Then
            OrdersReceived = True
        End If
    End If
End Function

Public Function NextRunDate() As Date
    Dim dtLast As Date
    Dim dtNext As Date
    
    dtLast = LastRunDate
    If dtLast <> 0 Then
        If Template.ScheduleActive Then
            dtNext = DateAdd("d", Template.ScheduleDays, dtLast)
            dtNext = CDate(Format(dtNext, "dd-mmm-yyyy") & " " & Format(Template.RunTime, "hh:nn:ss"))
            NextRunDate = dtNext
        Else
            NextRunDate = 0
        End If
    Else
        NextRunDate = 0
    End If
    NextRunDate = dtNext
End Function


Public Function LastRunDate() As Date
    If Created Then
        If Started Then
            LastRunDate = FileDateTime(Options.GalaxyNGNextTurn(GameName))
        End If
    End If
End Function
