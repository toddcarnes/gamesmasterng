Attribute VB_Name = "modDesign"
Option Explicit
Option Compare Text

Public Sub ApplyDesign(ByVal objTemplate As Template)

    Randomize
    If objTemplate.DesignType = LeaveAlone Then
        '
    ElseIf objTemplate.SeedType = NoSeeding Then
        If objTemplate.DesignType = OnCircle Then
            Call DesignCircle(objTemplate)
        ElseIf objTemplate.DesignType = OnCircleMiddle Then
            Call DesignCircleMiddle(objTemplate)
        ElseIf objTemplate.DesignType = GalaxyNGRandom Then
            Call DesignDefault(objTemplate)
        End If
    Else
        If objTemplate.DesignType = OnCircle Then
            Call SeedCircle(objTemplate)
        End If
    End If
End Sub

Private Sub DesignDefault(ByVal objTemplate As Template)
    Dim objRego As Registration
    Dim objWorld As HomeWorld
    
    Set objTemplate.Planets = Nothing
    For Each objRego In objTemplate.Registrations
        For Each objWorld In objRego.HomeWorlds
            objWorld.X = 0
            objWorld.Y = 0
        Next objWorld
    Next objRego
    
End Sub

Private Sub DesignCircle(ByVal objTemplate As Template)
' Populate the edge of the circle only
    Dim objRego As Registration
    Dim objWorld As HomeWorld
    Dim rc As Long
    Dim a As Single
    Dim ao As Single
    Dim R As Single
    Dim ro As Single
    Dim S As Single
    Dim Px As Single
    Dim Py As Single
    Dim W As Long
    Dim wa As Single
    Dim wao As Single
    Dim H As Long
    
    
    'calculate the radius of the circle
    R = CalcRadius(objTemplate.Registrations.Count, objTemplate.race_spacing)
    
    'calculate the galaxy Size
    S = Int(R + objTemplate.empty_radius)
    If objTemplate.Size < S Then
        objTemplate.Size = S
    End If
    
    'Calculate the center of the circle
    ro = objTemplate.Size / 2
    
    ' calculate the angle between races
    a = 2 * PI / objTemplate.Registrations.Count
    
    ' calculate a random offset for the races
    ao = -a * Rnd()

    'Create the Homeworlds where needed
    For Each objRego In objTemplate.Registrations
        If objRego.HomeWorlds.Count = 0 Then
            For H = 1 To objTemplate.DefaultHomeWorlds.Count
                Set objWorld = New HomeWorld
                objWorld.Size = objTemplate.DefaultHomeWorlds(H)
                objRego.HomeWorlds.Add objWorld
            Next H
        End If
    Next objRego

    For rc = 1 To objTemplate.Registrations.Count
        Set objRego = objTemplate.Registrations(rc)
        If objRego.HomeWorlds.Count = 0 Then
'
        End If
        Set objWorld = objRego.HomeWorlds(1)
        ao = ao + a
        Px = Round(R * Cos(ao) + ro)
        Py = Round(R * Sin(ao) + ro)
        objWorld.X = Px
        objWorld.Y = Py
    
        wa = 2 * PI / (objRego.HomeWorlds.Count - 1)
        wao = -wa * Rnd()
        For W = 2 To objRego.HomeWorlds.Count
            Set objWorld = objRego.HomeWorlds(W)
            If objTemplate.OrbitPlanets Then
                wao = wao + wa
                objWorld.X = Round(objTemplate.OrbitDistance * Cos(wao) + Px)
                objWorld.Y = Round(objTemplate.OrbitDistance * Sin(wao) + Py)
            Else
                objWorld.X = 0
                objWorld.Y = 0
            End If
        Next W
    Next rc
    
End Sub

Private Sub DesignCircleMiddle(ByVal objTemplate As Template)
' Populate the edge of the circle and put the last in the middle
    Dim objRego As Registration
    Dim objWorld As HomeWorld
    Dim rc As Long
    Dim a As Single
    Dim ao As Single
    Dim R As Single
    Dim ro As Single
    Dim S As Single
    Dim Px As Single
    Dim Py As Single
    Dim W As Long
    Dim wa As Single
    Dim wao As Single
    Dim H As Long
    
    'calculate the radius of the circle
    R = CalcRadius(objTemplate.Registrations.Count - 1, objTemplate.race_spacing)
    
    'calculate the galaxy Size
    S = Int(R + objTemplate.empty_radius)
    If objTemplate.Size < S Then
        objTemplate.Size = S
    End If
    
    'Calculate the center of the circle
    ro = objTemplate.Size / 2
    
    ' calculate the angle between races
    a = 2 * PI / objTemplate.Registrations.Count
    
    ' calculate a random offset for the races
    ao = -a * Rnd()
    
    'Create the Homeworlds where needed
    For Each objRego In objTemplate.Registrations
        If objRego.HomeWorlds.Count = 0 Then
            For H = 1 To objTemplate.DefaultHomeWorlds.Count
                Set objWorld = New HomeWorld
                objWorld.Size = objTemplate.DefaultHomeWorlds(H)
                objRego.HomeWorlds.Add objWorld
            Next H
        End If
    Next objRego

    For rc = 1 To objTemplate.Registrations.Count
        Set objRego = objTemplate.Registrations(rc)
        Set objWorld = objRego.HomeWorlds(1)
        If rc < objTemplate.Registrations.Count Then
            ao = ao + a
            Px = Round(R * Cos(ao) + ro)
            Py = Round(R * Sin(ao) + ro)
        Else
            Px = Round(objTemplate.Size / 2)
            Py = Round(objTemplate.Size / 2)
        End If
        objWorld.X = Px
        objWorld.Y = Py
    
        wa = 2 * PI / (objRego.HomeWorlds.Count - 1)
        wao = -wa * Rnd()
        For W = 2 To objRego.HomeWorlds.Count
            Set objWorld = objRego.HomeWorlds(W)
            If objTemplate.OrbitPlanets Then
                wao = wao + wa
                objWorld.X = Round(objTemplate.OrbitDistance * Cos(wao) + Px)
                objWorld.Y = Round(objTemplate.OrbitDistance * Sin(wao) + Py)
            Else
                objWorld.X = 0
                objWorld.Y = 0
            End If
        Next W
    Next rc
    
End Sub

Private Sub SeedCircle(ByVal objTemplate As Template)
' Populate the edge of the circle only
    Dim objRego As Registration
    Dim objWorld As HomeWorld
    Dim objPlanet As Planet
    Dim rc As Long
    Dim rc1 As Long
    Dim a As Single
    Dim ao As Single
    Dim R As Single
    Dim ro As Single
    Dim S As Single
    Dim Px As Single
    Dim Py As Single
    Dim Px1 As Single
    Dim Py1 As Single
    Dim W As Long
    Dim wa As Single
    Dim wao As Single
    Dim ws As Long
    Dim i As Long
    Dim H As Long
    
    'Initialise
    Set objTemplate.Planets = Nothing
    
    'calculate the radius of the circle
    R = CalcRadius(objTemplate.Registrations.Count, objTemplate.race_spacing)
    
    'calculate the galaxy Size
    S = Int(R + objTemplate.empty_radius)
    If objTemplate.Size < S Then
        objTemplate.Size = S
    End If
    
    'Calculate the center of the circle
    ro = objTemplate.Size / 2
    
    ' calculate the angle between races
    a = 2 * PI / objTemplate.Registrations.Count
    
    ' calculate a random offset for the races
    ao = -a * Rnd()

    'Create the Homeworlds where needed
    For Each objRego In objTemplate.Registrations
        If objRego.HomeWorlds.Count = 0 Then
            For H = 1 To objTemplate.DefaultHomeWorlds.Count
                Set objWorld = New HomeWorld
                objWorld.Size = objTemplate.DefaultHomeWorlds(H)
                objRego.HomeWorlds.Add objWorld
            Next H
        End If
    Next objRego
    
    'Position the Home worlds
    For rc = 1 To objTemplate.Registrations.Count
        Set objRego = objTemplate.Registrations(rc)
        ' Default the homeworlds if no nominated
        ao = ao + a
        Px = Round(R * Cos(ao) + ro)
        Py = Round(R * Sin(ao) + ro)
        If objTemplate.Seed(SeedHome) Then
            Set objPlanet = New Planet
            objPlanet.X = Px
            objPlanet.Y = Py
            objPlanet.Size = objTemplate.MaxPlanetSize * 4
            objPlanet.Resources = 10
            objTemplate.Planets.Add objPlanet
            wa = 2 * PI / (objRego.HomeWorlds.Count)
            ws = 1
        Else
            Set objWorld = objRego.HomeWorlds(1)
            objWorld.X = Px
            objWorld.Y = Py
            wa = 2 * PI / (objRego.HomeWorlds.Count - 1)
            ws = 2
        End If

        wao = -wa * Rnd()
        For W = ws To objRego.HomeWorlds.Count
            Set objWorld = objRego.HomeWorlds(W)
            If objTemplate.OrbitPlanets Then
                wao = wao + wa
                objWorld.X = Round(objTemplate.OrbitDistance * Cos(wao) + Px)
                objWorld.Y = Round(objTemplate.OrbitDistance * Sin(wao) + Py)
            Else
                objWorld.X = 0
                objWorld.Y = 0
            End If
        Next W
    Next rc
    
    ' Seed planets
    For rc = 1 To objTemplate.Registrations.Count
        If rc = 1 Then
            rc1 = objTemplate.Registrations.Count
        Else
            rc1 = rc - 1
        End If
        
        ' get homeworld and next homeworld position
        If objTemplate.Seed(SeedHome) Then
            With objTemplate.Planets(rc)
                Px = .X
                Py = .Y
            End With
            With objTemplate.Planets(rc1)
                Px1 = .X
                Py1 = .Y
            End With
        Else
            With objTemplate.Registrations(rc).HomeWorlds(1)
                Px = .X
                Py = .Y
            End With
            With objTemplate.Registrations(rc1).HomeWorlds(1)
                Px1 = .X
                Py1 = .Y
            End With
        End If
        'Seed a waypoint
        If objTemplate.Seed(SeedWaypoint) Then
            Set objPlanet = New Planet
            With objPlanet
                .X = (Px - Px1) / 2 + Px1
                .Y = (Py - Py1) / 2 + Py1
                .Size = Round(Rnd() * gcStuffMaxSize, 0)
                .Resources = Round(Rnd() * 10)
            End With
            objTemplate.Planets.Add objPlanet
        End If
        
        'Seed empty planets
        For i = 1 To objTemplate.empty_planets
            Set objPlanet = New Planet
            With objPlanet
                a = 2 * PI * Rnd()
                R = objTemplate.empty_radius * Rnd()
                .X = Cos(a) * R + Px
                .Y = Sin(a) * R + Py
                .Size = Round(Rnd() * (objTemplate.MaxPlanetSize - gcStuffMaxSize) + gcStuffMaxSize, 0)
                .Resources = Round(Rnd() * 10)
            End With
            objTemplate.Planets.Add objPlanet
        Next i
    Next rc
        
    For rc = 1 To objTemplate.Registrations.Count
        'Seed stuff planets around the galaxy
        For i = 1 To objTemplate.stuff_planets
            Set objPlanet = New Planet
            With objPlanet
                .X = objTemplate.Size * Rnd()
                .Y = objTemplate.Size * Rnd()
                .Size = Round(Rnd() * gcStuffMaxSize, 0)
                .Resources = Round(Rnd() * 10)
            End With
            objTemplate.Planets.Add objPlanet
        Next i
    Next rc
    
    'Seed the center of the galaxy
    If objTemplate.Seed(SeedCenter) Then
        Set objPlanet = New Planet
        With objPlanet
            .X = objTemplate.Size / 2
            .Y = objTemplate.Size / 2
            .Size = Round(Rnd() * gcStuffMaxSize, 0)
            .Resources = Round(Rnd() * 10)
        End With
        objTemplate.Planets.Add objPlanet
    End If
        
End Sub


Private Function CalcRadius(ByVal lngPlayers As Long, ByVal sngSpacing As Single) As Single
    'Angle Between Players
    'a = (2 * PI) / lngPlayers
    '
    'x = sngspacing / 2
    'Radius = x / sin(a/2)
    CalcRadius = (sngSpacing / 2) / Sin(PI / lngPlayers)
End Function


