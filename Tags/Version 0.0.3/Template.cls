VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Template"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Design Template for a Galaxy NG Game"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"Registrations"
Attribute VB_Ext_KEY = "Member1" ,"Message"
Attribute VB_Ext_KEY = "Member2" ,"Planets"
Option Explicit

Public Filename As String
Public TemplateName As String
Public Size As Single
Public race_spacing As Single
Public core_sizes As Variant
Public empty_planets As Long
Public empty_radius As Single
Public stuff_planets As Long
Public InitialTechLevels As Variant
Public FullBombing As Boolean
Public Peace As Long
Public KeepProduction As Boolean
Public DontDropDead As Boolean
Public sphericalgalaxy As Boolean
Public ScheduleActive As Boolean
Public MaxPlayers As Long
Public MinPlayers As Long
Public RegistrationOpen As Date
Public RegistrationClose As Date
Public RunTime As Date
Public StartDate As Date
Public ScheduleDays As Long
Public FinalOrders As Boolean
Public MaxPlanets As Long
Public MaxPlanetSize As Long
Public TotalPlanetSize As Long
Public Description As String
Public Message As String
Public Finished As Boolean

Public DesignType As uDesignType
Public OrbitPlanets As Boolean
Public OrbitDistance As Single
Public SeedType As uSeedType

Private mcolRegistrations As Registrations
Private mcolPlanets As Planets

Public IsNew As Boolean

Public Enum uDesignType
    LeaveAlone = 0
    OnCircle = 1
    OnCircleMiddle = 2
    GalaxyNGRandom = 3
End Enum

Public Enum uSeedType
    NoSeeding = 0
    SeedHome = 1
    SeedCenter = 2
    SeedWaypoint = 4
End Enum

Public Enum Tech
    Drive = 0
    Weapons = 1
    Shields = 2
    Cargo = 3
    Count = 3
End Enum

Public Function Seed(ByVal uSeed As uSeedType) As Boolean
    Seed = ((SeedType And uSeed) = uSeed)
End Function

Public Property Set Planets(ByVal colPlanets As Planets)
    Set mcolPlanets = colPlanets
End Property

Public Property Get Planets() As Planets
    If mcolPlanets Is Nothing Then
        Set mcolPlanets = New Planets
    End If
    Set Planets = mcolPlanets
End Property

Public Function OpenForRegistrations() As Boolean
    ' Not Open
    If RegistrationOpen = 0 _
    Or RegistrationOpen + RunTime > Now Then
        OpenForRegistrations = False
    
    ' Closed
    ElseIf RegistrationClose <> 0 _
    And RegistrationClose + RunTime <= Now Then
        OpenForRegistrations = False
    
    ' Open and not closed
    Else
        OpenForRegistrations = True
    End If
    
End Function

Public Property Get Registrations() As Registrations
    If mcolRegistrations Is Nothing Then
        Set mcolRegistrations = New Registrations
    End If
    Set Registrations = mcolRegistrations
End Property

Public Property Set Registrations(ByVal objRegistrations As Registrations)
    Set mcolRegistrations = objRegistrations
End Property

Public Property Get DefaultHomeWorlds() As HomeWorlds
    Dim objHomeWorlds As HomeWorlds
    Dim objHomeworld As HomeWorld
    Dim i As Long
    
    Set objHomeWorlds = New HomeWorlds
    For i = 0 To UBound(core_sizes)
        Set objHomeworld = New HomeWorld
        objHomeworld.Size = core_sizes(i)
        objHomeWorlds.Add objHomeworld
    Next i
    Set DefaultHomeWorlds = objHomeWorlds
End Property

Public Sub Load(Optional ByVal strFileName As String)
    Dim intFN As Integer
    Dim strLine As String
    Dim strFields() As String
    Dim i As Long
    Dim objRegistration As Registration
    Dim objHomeworld As HomeWorld
    Dim blnPlayersFound As Boolean
    Dim objPlanet As Planet
    Dim blnPlanets As Boolean
    
    Set Registrations = New Registrations
    Description = ""
    Message = ""
    If strFileName <> "" Then
        Filename = strFileName
    End If

    intFN = FreeFile
    Open Filename For Input As #intFN
    While Not EOF(intFN)
        Line Input #intFN, strLine
        strLine = Trim(strLine)
        If strLine <> "" Then
            strFields = Split(strLine, " ")
            Select Case LCase(strFields(0))
            Case ";scheduleactive"
                ScheduleActive = strFields(1)
            Case ";maxplanets"
                MaxPlanets = strFields(1)
            Case ";maxplanetsize"
                MaxPlanetSize = strFields(1)
            Case ";totalplanetsize"
                TotalPlanetSize = strFields(1)
            Case ";maxplayers" 'First comment line
                MaxPlayers = strFields(1)
                blnPlayersFound = True
            Case ";minplayers"
                MinPlayers = strFields(1)
            Case ";registrationopen"
                RegistrationOpen = CDate(strFields(1))
            Case ";registrationclose"
                RegistrationClose = CDate(strFields(1))
            Case ";runtime"
                RunTime = CDate(strFields(1))
            Case ";startdate"
                StartDate = CDate(strFields(1))
            Case ";scheduledays"
                ScheduleDays = strFields(1)
            Case ";finalorders"
                FinalOrders = strFields(1)
            Case ";description"
                If Description <> "" Then Description = Description & vbNewLine
                Description = Description & Mid(strLine, 14)
            Case ";message"
                If Message <> "" Then Message = Message & vbNewLine
                Message = Message & Mid(strLine, 10)
            Case ";finished"
                Finished = strFields(1)
            Case ";designtype"
                DesignType = Val(strFields(1))
            Case ";orbitplanets"
                OrbitPlanets = strFields(1)
            Case ";orbitdistance"
                OrbitDistance = Val(strFields(1))
            Case ";seedtype"
                SeedType = Val(strFields(1))
            
            Case "name"
                TemplateName = strFields(1)
            Case "size"
                Size = strFields(1)
            Case "race_spacing"
                race_spacing = strFields(1)
            Case "core_sizes"
                ReDim core_sizes(UBound(strFields) - 1)
                For i = 1 To UBound(strFields)
                    core_sizes(i - 1) = strFields(i)
                Next i
            Case "empty_planets"
                empty_planets = strFields(1)
            Case "empty_radius"
                empty_radius = strFields(1)
            Case "stuff_planets"
                stuff_planets = strFields(1)
            Case "initialtechlevels"
                InitialTechLevels(Tech.Drive) = strFields(1)
                InitialTechLevels(Tech.Weapons) = strFields(2)
                InitialTechLevels(Tech.Shields) = strFields(3)
                InitialTechLevels(Tech.Cargo) = strFields(4)
            Case "fullbombing"
                FullBombing = True
            Case "peace"
                Peace = strFields(1)
            Case "keepproduction"
                KeepProduction = True
            Case "dontdropdead"
                DontDropDead = True
            Case "sphericalgalaxy"
                sphericalgalaxy = True
            Case "start_player"
                Set objRegistration = New Registration
                If Not blnPlayersFound Then
                    MaxPlayers = MaxPlayers + 1
                End If
            Case "email"
                objRegistration.EMail = Mid(strLine, 7)
            Case "homeworld"
                If Not objRegistration Is Nothing Then
                    Set objHomeworld = New HomeWorld
                    objRegistration.HomeWorlds.Add objHomeworld
                    With objHomeworld
                        .Size = strFields(1)
                        If UBound(strFields) = 3 Then
                            .X = strFields(2)
                            .Y = strFields(3)
                        End If
                    End With
                End If
            Case "end_player"
                If InStr(1, objRegistration.EMail, "@their.address.tld") = 0 Then
                    Registrations.Add objRegistration
                    Set objHomeworld = Nothing
                End If
            Case "start_planets"
                blnPlanets = True
            Case "end_planets"
                blnPlanets = False
            Case Else
                If blnPlanets Then
                    Set objPlanet = New Planet
                    objPlanet.X = Val(strFields(0))
                    objPlanet.Y = Val(strFields(1))
                    If UBound(strFields) < 2 Then
                        objPlanet.Size = -1
                    Else
                        objPlanet.Size = Val(strFields(2))
                    End If
                    If UBound(strFields) < 3 Then
                        objPlanet.Resources = -1
                    Else
                        objPlanet.Resources = Val(strFields(3))
                    End If
                    Planets.Add objPlanet
                End If
            End Select
        End If
    Wend
    Close #intFN
    IsNew = False
    
End Sub

Public Function Clone() As Template
    Dim i As Long
    Dim strNo As String
    
    Set Clone = New Template
    With Clone
'        .IsNew = .IsNew
'        .Filename = filename
'        .ScheduleActive = ScheduleActive
'        .FinalOrders = Finished
        .MaxPlanets = MaxPlanets
        .MaxPlanetSize = MaxPlanetSize
        .TotalPlanetSize = TotalPlanetSize
        .MinPlayers = MinPlayers
        .MaxPlayers = MaxPlayers
'        .RegistrationOpen = RegistrationOpen
'        .RegistrationClose = RegistrationClose
        .RunTime = RunTime
'        .StartDate = StartDate
        .ScheduleDays = ScheduleDays
        .FinalOrders = FinalOrders
        .MaxPlanets = MaxPlanets
        .MaxPlanetSize = MaxPlanetSize
        .TotalPlanetSize = TotalPlanetSize
        .Description = Description
'        .Message = Message
        
        i = Len(TemplateName)
        Do
            If Not IsNumeric(Mid(TemplateName, i)) Then
                Exit Do
            End If
            strNo = Mid(TemplateName, i)
            i = i - 1
        Loop
        If Val(strNo) = 0 Then strNo = "1"
        .TemplateName = Left(TemplateName, i) & Val(strNo) + 1
        .Size = Size
        .race_spacing = race_spacing
        .core_sizes = core_sizes
        .empty_planets = empty_planets
        .empty_radius = empty_radius
        .InitialTechLevels = InitialTechLevels
        .FullBombing = FullBombing
        .Peace = Peace
        .KeepProduction = KeepProduction
        .DontDropDead = DontDropDead
        .sphericalgalaxy = sphericalgalaxy
    End With
End Function

Public Sub Save()
    Dim intFN As Integer
    Dim i As Long
    Dim objRegistration As Registration
    Dim objHomeworld As HomeWorld
    Dim vData As Variant
    Dim objPlanet As Planet
    
    If Filename = vbNullString Then
        Filename = Options.GalaxyNGHome & TemplateName & ".glx"
    End If
    
    If Dir(Filename) <> "" Then
        If Dir(Filename & ".bak") <> "" Then
            Kill Filename & ".bak"
        End If
        Name Filename As Filename & ".bak"
    End If
    
    intFN = FreeFile
    Open Filename For Output As intFN
    Print #intFN, ";ScheduleActive "; CStr(ScheduleActive)
    Print #intFN, ";MaxPlayers "; CStr(MaxPlayers)
    Print #intFN, ";MinPlayers "; CStr(MinPlayers)
    Print #intFN, ";MaxPlanets "; CStr(MaxPlanets)
    Print #intFN, ";MaxPlanetSize "; CStr(MaxPlanetSize)
    Print #intFN, ";TotalPlanetSize "; CStr(TotalPlanetSize)
    Print #intFN, ";RegistrationOpen "; Format(RegistrationOpen, "dd-mmm-yyyy")
    Print #intFN, ";RegistrationClose "; Format(RegistrationClose, "dd-mmm-yyyy")
    Print #intFN, ";RunTime "; Format(RunTime, "hh:nn:ss")
    Print #intFN, ";Startdate "; Format(StartDate, "dd-mmm-yyyy")
    Print #intFN, ";ScheduleDays "; CStr(ScheduleDays)
    Print #intFN, ";FinalOrders "; CStr(FinalOrders)
    Print #intFN, ";DesignType "; CStr(DesignType)
    Print #intFN, ";OrbitPlanets "; CStr(OrbitPlanets)
    Print #intFN, ";OrbitDistance "; CStr(OrbitDistance)
    Print #intFN, ";SeedType "; CStr(SeedType)
    
    If Description <> "" Then
        vData = Split(Description, vbNewLine)
        For i = LBound(vData) To UBound(vData)
            Print #intFN, ";Description "; vData(i)
        Next i
    End If
    If Message <> "" Then
        vData = Split(Message, vbNewLine)
        For i = LBound(vData) To UBound(vData)
            Print #intFN, ";Message "; vData(i)
        Next i
    End If
    Print #intFN, ";Finished "; CStr(Finished)
    
    Print #intFN, "Name "; TemplateName
    
    Print #intFN, "Size "; CStr(Size)
    Print #intFN, "race_spacing "; CStr(race_spacing)
    Print #intFN, "Core_Sizes";
    For i = 0 To UBound(core_sizes)
        Print #intFN, " "; CStr(core_sizes(i));
    Next i
    Print #intFN, ""
    Print #intFN, "empty_planets "; CStr(empty_planets)
    Print #intFN, "empty_radius "; CStr(empty_radius)
    Print #intFN, "stuff_planets "; CStr(stuff_planets)
    Print #intFN, "InitialTechLevels "; CStr(InitialTechLevels(Tech.Drive)) & _
                                    " "; CStr(InitialTechLevels(Tech.Weapons)) & _
                                    " "; CStr(InitialTechLevels(Tech.Shields)) & _
                                    " "; CStr(InitialTechLevels(Tech.Cargo))
    If Not FullBombing Then Print #intFN, ";";
    Print #intFN, "FullBombing"
    If Peace = 0 Then Print #intFN, ";";
    Print #intFN, "Peace "; CStr(Peace)
    If Not KeepProduction Then Print #intFN, ";";
    Print #intFN, "KeepProduction"
    If Not DontDropDead Then Print #intFN, ";";
    Print #intFN, "DontDropDead"
    If Not sphericalgalaxy Then Print #intFN, ";";
    Print #intFN, "SphericalGalaxy"
    
    If Planets.Count > 0 Then
        Print #intFN, "Start_Planets"
        For Each objPlanet In Planets
            Print #intFN, CStr(objPlanet.X); " "; CStr(objPlanet.Y);
            If objPlanet.Size >= 0 Then
                Print #intFN, " "; CStr(objPlanet.Size);
                If objPlanet.Resources >= 0 Then
                    Print #intFN, " "; CStr(objPlanet.Resources);
                End If
            End If
            Print #intFN,
        Next objPlanet
        Print #intFN, "End_Planets"
    End If
    
    For Each objRegistration In Registrations
        Print #intFN, "start_player"
        Print #intFN, "email "; objRegistration.EMail
        For Each objHomeworld In objRegistration.HomeWorlds
            Print #intFN, "homeworld "; CStr(objHomeworld.Size);
            If objHomeworld.X <> 0 Or objHomeworld.Y <> 0 Then
                Print #intFN, " "; CStr(objHomeworld.X);
                Print #intFN, " "; CStr(objHomeworld.Y);
            End If
            Print #intFN, ""
        Next objHomeworld
        Print #intFN, "end_player"
    Next objRegistration
    Close #intFN
    IsNew = False

End Sub

Public Function Details() As String
    Dim strDetails As String
    Dim i As Long
    
    strDetails = strDetails & "Name: " & TemplateName & vbNewLine
    strDetails = strDetails & "Active: " & CStr(ScheduleActive) & vbNewLine
    strDetails = strDetails & "Game Ended: " & CStr(Finished) & vbNewLine
    
    strDetails = strDetails & vbNewLine
    strDetails = strDetails & "Registrations" & vbNewLine
    strDetails = strDetails & "-------------" & vbNewLine
    strDetails = strDetails & "Open: " & Format(RegistrationOpen, "dd-mmm-yyyy") & vbNewLine
    strDetails = strDetails & "Close: " & Format(RegistrationClose, "dd-mmm-yyyy") & vbNewLine
    strDetails = strDetails & "Maximum Players: " & CStr(MaxPlayers) & vbNewLine
    strDetails = strDetails & "Minimum Players: " & CStr(MinPlayers) & vbNewLine
    strDetails = strDetails & "Maximum Planets: " & CStr(MaxPlanets) & vbNewLine
    strDetails = strDetails & "Maximum Planet Size: " & CStr(MaxPlanetSize) & vbNewLine
    strDetails = strDetails & "Total Planet Size: " & CStr(TotalPlanetSize) & vbNewLine
    strDetails = strDetails & "Core Sizes:"
    For i = 0 To UBound(core_sizes)
        strDetails = strDetails & " " & CStr(core_sizes(i))
    Next i
    strDetails = strDetails & vbNewLine
    
    strDetails = strDetails & vbNewLine
    strDetails = strDetails & "Running Details" & vbNewLine
    strDetails = strDetails & "---------------" & vbNewLine
    strDetails = strDetails & "Run Time: " & Format(RunTime, "hh:nn:ss") & vbNewLine
    strDetails = strDetails & "Start Date: " & Format(StartDate, "dd-mmm-yyyy") & vbNewLine
    strDetails = strDetails & "Schedule Days: " & CStr(ScheduleDays) & vbNewLine
    strDetails = strDetails & "Final Orders: " & CStr(FinalOrders) & vbNewLine
    
    strDetails = strDetails & vbNewLine
    strDetails = strDetails & "Galaxy Design Details" & vbNewLine
    strDetails = strDetails & "---------------------" & vbNewLine
    strDetails = strDetails & "Size: " & CStr(Size) & vbNewLine
    strDetails = strDetails & "Race Spacing: " & CStr(race_spacing) & vbNewLine
    strDetails = strDetails & "Empty Planets: " & CStr(empty_planets) & vbNewLine
    strDetails = strDetails & "Empty Radius: " & CStr(empty_radius) & vbNewLine
    strDetails = strDetails & "Stuff Planets: " & CStr(stuff_planets) & vbNewLine
    
    strDetails = strDetails & "Design Type: " & Options.DesignDescriptions(DesignType) & vbNewLine
    strDetails = strDetails & "Orbit Planets: " & CStr(OrbitPlanets) & vbNewLine
    strDetails = strDetails & "Orbit Distance: " & CStr(OrbitDistance) & vbNewLine
    strDetails = strDetails & "Seeding Type: " & Options.SeedDescriptions(SeedType) & vbNewLine
    
    strDetails = strDetails & vbNewLine
    strDetails = strDetails & "Game Details" & vbNewLine
    strDetails = strDetails & "------------" & vbNewLine
    strDetails = strDetails & "Initial Tech Levels: Drive:   " & CStr(InitialTechLevels(Tech.Drive)) & vbNewLine
    strDetails = strDetails & "                     Weapons: " & CStr(InitialTechLevels(Tech.Weapons)) & vbNewLine
    strDetails = strDetails & "                     Shields: " & CStr(InitialTechLevels(Tech.Shields)) & vbNewLine
    strDetails = strDetails & "                     Cargo:   " & CStr(InitialTechLevels(Tech.Cargo)) & vbNewLine
    
    strDetails = strDetails & "Full Bombing: " & CStr(FullBombing) & vbNewLine
    strDetails = strDetails & "Enforced Peace: "
    If Peace = 0 Then
        strDetails = strDetails & "False" & vbNewLine
    Else
        strDetails = strDetails & CStr(Peace) & " Turns" & vbNewLine
    End If
    strDetails = strDetails & "Keep Production: " & CStr(KeepProduction) & vbNewLine
    strDetails = strDetails & "Dont Drop Dead: " & CStr(DontDropDead) & vbNewLine
    strDetails = strDetails & "Spherical Galaxy: " & CStr(sphericalgalaxy) & vbNewLine
    Details = strDetails
End Function

Private Sub Class_Initialize()
    ReDim core_sizes(0)
    core_sizes(0) = 1000
    
    ReDim InitialTechLevels(Tech.Count)
    InitialTechLevels(Tech.Drive) = 1
    InitialTechLevels(Tech.Weapons) = 1
    InitialTechLevels(Tech.Shields) = 1
    InitialTechLevels(Tech.Cargo) = 1
    FullBombing = False
    KeepProduction = False
    DontDropDead = False
    sphericalgalaxy = False
    IsNew = True
End Sub

Private Sub Class_Terminate()
    Set Planets = Nothing
    Set Registrations = Nothing
End Sub
