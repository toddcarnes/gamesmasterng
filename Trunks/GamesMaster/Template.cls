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


Private mobjRegistrations As Registrations

Public Enum Tech
    Drive = 0
    Weapons = 1
    Shields = 2
    Cargo = 3
    Count = 3
End Enum

Public Function OpenForRegistrations() As Boolean
    ' Not Open
    If RegistrationOpen = 0 _
    Or RegistrationOpen > Now Then
        OpenForRegistrations = False
    
    ' Closed
    ElseIf RegistrationClose <> 0 _
    And RegistrationClose <= Now Then
        OpenForRegistrations = False
    
    ' Open and not closed
    Else
        OpenForRegistrations = True
    End If
    
End Function


Public Property Get Registrations() As Registrations
    If mobjRegistrations Is Nothing Then
        Set mobjRegistrations = New Registrations
    End If
    Set Registrations = mobjRegistrations
End Property

Public Property Set Registrations(ByVal objRegistrations As Registrations)
    Set mobjRegistrations = objRegistrations
End Property

Public Sub Load(Optional ByVal strFileName As String)
    Dim intFN As Integer
    Dim strLine As String
    Dim strFields() As String
    Dim i As Long
    Dim objRegistration As Registration
    Dim objHomeWorld As HomeWorld
    Dim blnPlayersFound As Boolean
    
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
                RegistrationOpen = CDate(strFields(1) & " " & strFields(2))
            Case ";registrationclose"
                RegistrationClose = CDate(strFields(1) & " " & strFields(2))
            Case ";runtime"
                RunTime = CDate(strFields(1) & " " & strFields(2))
            Case ";startdate"
                StartDate = CDate(strFields(1) & " " & strFields(2))
            Case ";scheduledays"
                ScheduleDays = strFields(1)
            Case ";finalorders"
                FinalOrders = strFields(1)
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
                    Set objHomeWorld = New HomeWorld
                    objRegistration.HomeWorlds.Add objHomeWorld
                    With objHomeWorld
                        .Size = strFields(1)
                        If UBound(strFields) = 3 Then
                            .x = strFields(2)
                            .y = strFields(3)
                        End If
                    End With
                End If
            Case "end_player"
                If InStr(1, objRegistration.EMail, "@their.address.tld") = 0 Then
                    Registrations.Add objRegistration
                    Set objHomeWorld = Nothing
                End If
            End Select
        End If
    Wend
    Close #intFN

End Sub

Public Function Clone(ByVal strTemplateName As String) As Template
    Set Clone = New Template
    With Clone
        .Filename = GalaxyNGHome & strTemplateName & ".glx"
        .MaxPlanets = MaxPlanets
        .MaxPlanetSize = MaxPlanetSize
        .TotalPlanetSize = TotalPlanetSize
        .MaxPlayers = MaxPlayers
        .MinPlayers = MinPlayers
        .RegistrationOpen = RegistrationOpen
        .RegistrationClose = RegistrationClose
        
        .TemplateName = strTemplateName
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
        .RunTime = RunTime
        .ScheduleDays = ScheduleDays
        .FinalOrders = FinalOrders
    End With
End Function

Public Sub Save(Optional ByVal strFileName As String)
    Dim intFN As Integer
    Dim i As Long
    Dim objRegistration As Registration
    Dim objHomeWorld As HomeWorld
    
    If strFileName = vbNullString Then
        strFileName = Filename
    End If
    
    If Dir(strFileName) <> "" Then
        If Dir(strFileName & ".bak") <> "" Then
            Kill strFileName & ".bak"
        End If
        Name strFileName As strFileName & ".bak"
    End If
    
    intFN = FreeFile
    Open strFileName For Output As intFN
    Print #intFN, ";MaxPlayers "; CStr(MaxPlayers)
    Print #intFN, ";MinPlayers "; CStr(MinPlayers)
    Print #intFN, ";MaxPlanets "; CStr(MaxPlanets)
    Print #intFN, ";MaxPlanetSize "; CStr(MaxPlanetSize)
    Print #intFN, ";TotalPlanetSize "; CStr(TotalPlanetSize)
    Print #intFN, ";RegistrationOpen "; Format(RegistrationOpen, "dd-mmm-yyyy hh:nn:ss")
    Print #intFN, ";RegistrationClose "; Format(RegistrationClose, "dd-mmm-yyyy hh:nn:ss")
    Print #intFN, ";RunTime "; Format(RunTime, "dd-mmm-yyyy hh:nn:ss")
    Print #intFN, ";Startdate "; Format(StartDate, "dd-mmm-yyyy hh:nn:ss")
    Print #intFN, ";ScheduleDays "; CStr(ScheduleDays)
    Print #intFN, ";FinalOrders "; CStr(FinalOrders)
    
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
    
    For Each objRegistration In Registrations
        Print #intFN, "start_player"
        Print #intFN, "email "; objRegistration.EMail
        For Each objHomeWorld In objRegistration.HomeWorlds
            Print #intFN, "homeworld "; CStr(objHomeWorld.Size);
            If objHomeWorld.x <> 0 Or objHomeWorld.y <> 0 Then
                Print #intFN, " "; CStr(objHomeWorld.x);
                Print #intFN, " "; CStr(objHomeWorld.y);
            End If
            Print #intFN, ""
        Next objHomeWorld
        Print #intFN, "end_player"
    Next objRegistration
    Close #intFN

End Sub

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
End Sub
