Attribute VB_Name = "modDesign"
'********************************************************
'   Copyright 2007,2008 Ian Evans.                      *
'   This program is distributed under the terms of the  *
'       GNU General Public License.                     *
'********************************************************
Option Explicit
Option Compare Text

Private mobjTemplate As Template

Public Sub ApplyDesign(ByVal objTemplate As Template)

    If objTemplate.Registrations.Count < 1 Then Exit Sub
    
    Set mobjTemplate = objTemplate
    Randomize
    
    If mobjTemplate.DesignType = LeaveAlone Then
        ' do nothing
    ElseIf mobjTemplate.DesignType = GalaxyNGRandom Then
        Call DesignDefault
    Else
        Call BuildGalaxy
    End If
    
'    ElseIf mobjTemplate.SeedType = NoSeeding Then
'        If mobjTemplate.DesignType = OnCircle Then
'            Call DesignCircle
'        ElseIf mobjTemplate.DesignType = OnCircleMiddle Then
'            Call DesignCircleMiddle
'        ElseIf mobjTemplate.DesignType = GalaxyNGRandom Then
'            Call DesignDefault
'        End If
'    ElseIf mobjTemplate.DesignType = OnCircle Then
'        Call SeedCircle
'    ElseIf mobjTemplate.DesignType = GenerateRandom Then
'        'Call SeedGalaxy
'    End If
    
    Set mobjTemplate = Nothing
End Sub

Private Sub DesignDefault()
    Dim objRego As Registration
    Dim objWorld As HomeWorld
    
    Set mobjTemplate.Planets = Nothing
    For Each objRego In mobjTemplate.Registrations
        For Each objWorld In objRego.HomeWorlds
            objWorld.X = 0
            objWorld.Y = 0
        Next objWorld
    Next objRego
    
End Sub

Private Sub DesignCircle()
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
    R = CalcRadius(mobjTemplate.Registrations.Count, mobjTemplate.race_spacing)
    
    'calculate the galaxy Size
    S = Int(R + mobjTemplate.empty_radius)
    If mobjTemplate.Size < S Then
        mobjTemplate.Size = S
    End If
    
    'Calculate the center of the circle
    ro = mobjTemplate.Size / 2
    
    ' calculate the angle between races
    a = 2 * PI / mobjTemplate.Registrations.Count
    
    ' calculate a random offset for the races
    ao = -a * Rnd()

    'Create the Homeworlds where needed
    For Each objRego In mobjTemplate.Registrations
        If objRego.HomeWorlds.Count = 0 Then
            For H = 1 To mobjTemplate.DefaultHomeWorlds.Count
                Set objWorld = New HomeWorld
                objWorld.Size = mobjTemplate.DefaultHomeWorlds(H)
                objRego.HomeWorlds.Add objWorld
            Next H
        End If
    Next objRego

    For rc = 1 To mobjTemplate.Registrations.Count
        Set objRego = mobjTemplate.Registrations(rc)
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
            If mobjTemplate.OrbitPlanets Then
                wao = wao + wa
                objWorld.X = Round(mobjTemplate.OrbitDistance * Cos(wao) + Px)
                objWorld.Y = Round(mobjTemplate.OrbitDistance * Sin(wao) + Py)
            Else
                objWorld.X = 0
                objWorld.Y = 0
            End If
        Next W
    Next rc
    
End Sub

Private Sub DesignCircleMiddle()
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
    R = CalcRadius(mobjTemplate.Registrations.Count - 1, mobjTemplate.race_spacing)
    
    'calculate the galaxy Size
    S = Int(R + mobjTemplate.empty_radius)
    If mobjTemplate.Size < S Then
        mobjTemplate.Size = S
    End If
    
    'Calculate the center of the circle
    ro = mobjTemplate.Size / 2
    
    ' calculate the angle between races
    a = 2 * PI / mobjTemplate.Registrations.Count
    
    ' calculate a random offset for the races
    ao = -a * Rnd()
    
    'Create the Homeworlds where needed
    For Each objRego In mobjTemplate.Registrations
        If objRego.HomeWorlds.Count = 0 Then
            For H = 1 To mobjTemplate.DefaultHomeWorlds.Count
                Set objWorld = New HomeWorld
                objWorld.Size = mobjTemplate.DefaultHomeWorlds(H)
                objRego.HomeWorlds.Add objWorld
            Next H
        End If
    Next objRego

    For rc = 1 To mobjTemplate.Registrations.Count
        Set objRego = mobjTemplate.Registrations(rc)
        Set objWorld = objRego.HomeWorlds(1)
        If rc < mobjTemplate.Registrations.Count Then
            ao = ao + a
            Px = Round(R * Cos(ao) + ro)
            Py = Round(R * Sin(ao) + ro)
        Else
            Px = Round(mobjTemplate.Size / 2)
            Py = Round(mobjTemplate.Size / 2)
        End If
        objWorld.X = Px
        objWorld.Y = Py
    
        wa = 2 * PI / (objRego.HomeWorlds.Count - 1)
        wao = -wa * Rnd()
        For W = 2 To objRego.HomeWorlds.Count
            Set objWorld = objRego.HomeWorlds(W)
            If mobjTemplate.OrbitPlanets Then
                wao = wao + wa
                objWorld.X = Round(mobjTemplate.OrbitDistance * Cos(wao) + Px)
                objWorld.Y = Round(mobjTemplate.OrbitDistance * Sin(wao) + Py)
            Else
                objWorld.X = 0
                objWorld.Y = 0
            End If
        Next W
    Next rc
    
End Sub

Private Sub SeedCircle()
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
    Set mobjTemplate.Planets = Nothing
    
    'calculate the radius of the circle
    R = CalcRadius(mobjTemplate.Registrations.Count, mobjTemplate.race_spacing)
    
    'calculate the galaxy Size
    S = Int(R + mobjTemplate.empty_radius)
    If mobjTemplate.Size < S Then
        mobjTemplate.Size = S
    End If
    
    'Calculate the center of the circle
    ro = mobjTemplate.Size / 2
    
    ' calculate the angle between races
    a = 2 * PI / mobjTemplate.Registrations.Count
    
    ' calculate a random offset for the races
    ao = -a * Rnd()

    'Create the Homeworlds where needed
    For Each objRego In mobjTemplate.Registrations
        If objRego.HomeWorlds.Count = 0 Then
            For H = 1 To mobjTemplate.DefaultHomeWorlds.Count
                Set objWorld = New HomeWorld
                objWorld.Size = mobjTemplate.DefaultHomeWorlds(H)
                objRego.HomeWorlds.Add objWorld
            Next H
        End If
    Next objRego
    
    'Position the Home worlds
    For rc = 1 To mobjTemplate.Registrations.Count
        Set objRego = mobjTemplate.Registrations(rc)
        ' Default the homeworlds if no nominated
        ao = ao + a
        Px = Round(R * Cos(ao) + ro)
        Py = Round(R * Sin(ao) + ro)
        If mobjTemplate.Seed(SeedHome) Then
            Set objPlanet = New Planet
            objPlanet.X = Px
            objPlanet.Y = Py
            objPlanet.Size = mobjTemplate.MaxPlanetSize * 4
            objPlanet.Resources = 10
            mobjTemplate.Planets.Add objPlanet
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
            If mobjTemplate.OrbitPlanets Then
                wao = wao + wa
                objWorld.X = Round(mobjTemplate.OrbitDistance * Cos(wao) + Px)
                objWorld.Y = Round(mobjTemplate.OrbitDistance * Sin(wao) + Py)
            Else
                objWorld.X = 0
                objWorld.Y = 0
            End If
        Next W
    Next rc
    
    ' Seed planets
    For rc = 1 To mobjTemplate.Registrations.Count
        If rc = 1 Then
            rc1 = mobjTemplate.Registrations.Count
        Else
            rc1 = rc - 1
        End If
        
        ' get homeworld and next homeworld position
        If mobjTemplate.Seed(SeedHome) Then
            With mobjTemplate.Planets(rc)
                Px = .X
                Py = .Y
            End With
            With mobjTemplate.Planets(rc1)
                Px1 = .X
                Py1 = .Y
            End With
        Else
            With mobjTemplate.Registrations(rc).HomeWorlds(1)
                Px = .X
                Py = .Y
            End With
            With mobjTemplate.Registrations(rc1).HomeWorlds(1)
                Px1 = .X
                Py1 = .Y
            End With
        End If
        'Seed a waypoint
        If mobjTemplate.Seed(SeedWaypoint) Then
            Set objPlanet = New Planet
            With objPlanet
                .X = (Px - Px1) / 2 + Px1
                .Y = (Py - Py1) / 2 + Py1
                .Size = Round(Rnd() * gcStuffMaxSize, 0)
                .Resources = Round(Rnd() * 10)
            End With
            mobjTemplate.Planets.Add objPlanet
        End If
        
        'Seed empty planets
        For i = 1 To mobjTemplate.empty_planets
            Set objPlanet = New Planet
            With objPlanet
                a = 2 * PI * Rnd()
                R = mobjTemplate.empty_radius * Rnd()
                .X = Cos(a) * R + Px
                .Y = Sin(a) * R + Py
                .Size = Round(Rnd() * (mobjTemplate.MaxPlanetSize - gcStuffMaxSize) + gcStuffMaxSize, 0)
                .Resources = Round(Rnd() * 10)
            End With
            mobjTemplate.Planets.Add objPlanet
        Next i
    Next rc
        
    For rc = 1 To mobjTemplate.Registrations.Count
        'Seed stuff planets around the galaxy
        For i = 1 To mobjTemplate.stuff_planets
            Set objPlanet = New Planet
            With objPlanet
                .X = mobjTemplate.Size * Rnd()
                .Y = mobjTemplate.Size * Rnd()
                .Size = Round(Rnd() * gcStuffMaxSize, 0)
                .Resources = Round(Rnd() * 10)
            End With
            mobjTemplate.Planets.Add objPlanet
        Next i
    Next rc
    
    'Seed the center of the galaxy
    If mobjTemplate.Seed(SeedCenter) Then
        Set objPlanet = New Planet
        With objPlanet
            .X = mobjTemplate.Size / 2
            .Y = mobjTemplate.Size / 2
            .Size = Round(Rnd() * gcStuffMaxSize, 0)
            .Resources = Round(Rnd() * 10)
        End With
        mobjTemplate.Planets.Add objPlanet
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

Private Sub BuildGalaxy()
' This is a universal build Module for the games
    Dim objRego As Registration

    With mobjTemplate
        Set .Planets = New Planets
        
        ' Set the minimum Galaxy Size
        Call SetGalaxySize
        
        'Set the Player Primary Positions
        While Not SetPlayerPositions()
            .Size = .Size + .race_spacing / 4
        Wend
        
        'Setup the Player positions
        For Each objRego In .Registrations
            Call SetPrimaryWorld(objRego)
            Call SetSecondaryWorlds(objRego)
        Next objRego
        
        'Put in the additional design planets
        Call SetAdditionalSeedings
        
        'Setup the Player positions
        For Each objRego In .Registrations
            Call SetEmptyWorlds(objRego)
        Next objRego
        
        'Put in the stuff planets
        Call SetStuffWorlds
        
        
    End With
End Sub

Private Sub SetGalaxySize()
    Dim R As Single
    Dim S As Long
    
    With mobjTemplate
        If .DesignType = OnCircle Then
            R = CalcRadius(.Registrations.Count, .race_spacing)
            S = Int(2 * (R + .empty_radius))
            .sphericalgalaxy = False
        ElseIf .DesignType = OnCircleMiddle Then
            R = CalcRadius(.Registrations.Count - 1, .race_spacing)
            S = Int(2 * (R + .empty_radius))
            .sphericalgalaxy = False
        
        ElseIf .DesignType = GenerateRandom Then
            S = Int(Sqr(.Registrations.Count))
            S = S * .race_spacing
        End If
        
        If S > .Size Then
            .Size = S
        End If
    End With
End Sub

Private Function SetPlayerPositions() As Boolean
    Dim ro As Single
    Dim a As Single
    Dim ao As Single
    Dim Rego As Registration
    Dim Rego2 As Registration
    Dim rc As Long
    Dim R As Single
    Dim i As Long
    Dim d As Single
    Dim minD As Single
    Dim blnBad As Boolean
    Dim lngCount1 As Long
    Dim lngCount2 As Long
    Dim X As Single
    Dim X1 As Single
    Dim Y As Single
    Dim Y1 As Single
    
    With mobjTemplate
        
        If .DesignType = OnCircle Then
            ro = .Size / 2  'Calculate the center of the circle
            a = 2 * PI / .Registrations.Count ' calculate the angle between races
            ao = -a * Rnd() ' calculate a random offset for the races
            R = CalcRadius(.Registrations.Count, .race_spacing)
            'Position the Players
            For rc = 1 To .Registrations.Count
                Set Rego = .Registrations(rc)
                ao = ao + a
                Rego.X = Round(R * Cos(ao) + ro)
                Rego.Y = Round(R * Sin(ao) + ro)
            Next rc
            SetPlayerPositions = True
        
        ElseIf .DesignType = OnCircleMiddle Then
            ro = .Size / 2  'Calculate the center of the circle
            a = 2 * PI / (.Registrations.Count - 1) ' calculate the angle between races
            ao = -a * Rnd() ' calculate a random offset for the races
            R = CalcRadius(.Registrations.Count - 1, .race_spacing)
            'Position the Players
            For rc = 1 To .Registrations.Count - 1
                Set Rego = .Registrations(rc)
                ao = ao + a
                Rego.X = Round(R * Cos(ao) + ro)
                Rego.Y = Round(R * Sin(ao) + ro)
            Next rc
            Set Rego = .Registrations(.Registrations.Count)
            Rego.X = ro
            Rego.Y = ro
            SetPlayerPositions = True
        
        Else 'Place players randomly in the galaxy
            minD = .race_spacing ^ 2
            lngCount2 = gcMaxTries
            Do
                For rc = 1 To .Registrations.Count
                    Set Rego = .Registrations(rc)
                    lngCount1 = gcMaxTries
                    Do
                        'Generate a random position
                        Rego.X = Round(Rnd() * .Size)
                        Rego.Y = Round(Rnd() * .Size)
                        blnBad = False
                        
                        ' Too close to edge of galaxy
                        If (Rego.X < .empty_radius _
                        Or Rego.Y < .empty_radius _
                        Or Rego.X > (.Size - .empty_radius) _
                        Or Rego.Y > (.Size - .empty_radius)) Then
                            blnBad = True
                        
                        Else
                            'Check to see if it is too close to other players
                            For i = 1 To rc - 1
                                Set Rego2 = .Registrations(i)
                                'Calculate the distance between players
                                X = Abs(Rego.X - Rego2.X)
                                Y = Abs(Rego.Y - Rego2.Y)
                                If .sphericalgalaxy Then 'See if closer wrapping
                                    X1 = .Size - X
                                    If X1 < X Then X = X1
                                    Y1 = .Size - Y
                                    If Y1 < Y Then Y = Y1
                                End If
                                d = X * X + Y * Y
                                
                                If d < minD Then
                                    blnBad = True
                                    Exit For
                                End If
                            Next i
                        End If
                        
                        If Not blnBad Then Exit Do 'Position is OK
                        
                        lngCount1 = lngCount1 - 1 'Stop infinite looping
                        If lngCount1 < 0 Then
                            blnBad = True
                            Exit Do
                        End If
                    Loop
                    If blnBad Then 'Can't space the players
                        Exit For
                    End If
                Next rc
            
                If Not blnBad Then Exit Do
            
                'Try all players again
                lngCount2 = lngCount2 - 1
                If lngCount2 < 0 Then
                    Exit Do
                End If
            Loop
            
            If Not blnBad Then 'Successfully spaced players
                SetPlayerPositions = True
            End If
        End If
    End With
End Function

Private Sub SetPrimaryWorld(ByVal objRego As Registration)
    Dim objWorld As HomeWorld
    Dim objPlanet As Planet
    Dim i As Long
    
    With objRego
        'Generate default Home World Sizes
        If .HomeWorlds.Count = 0 Then
            For i = 0 To UBound(mobjTemplate.core_sizes)
                Set objWorld = New HomeWorld
                objWorld.Size = mobjTemplate.core_sizes(i)
                .HomeWorlds.Add objWorld
            Next i
        End If
        'Reset World Locations
        For Each objWorld In .HomeWorlds
            objWorld.X = 0
            objWorld.Y = 0
        Next objWorld
        
        ' The Primary world will be a seeded planet
        If mobjTemplate.Seed(SeedHome) Then
            Set objPlanet = New Planet
            objPlanet.X = .X
            objPlanet.Y = .Y
            objPlanet.Size = mobjTemplate.SeedSize(SeedHome)
            If objPlanet.Size = 0 Then
                objPlanet.Size = gcSeededHomeSize
            End If
            objPlanet.Resources = gcSeededHomeResources
            mobjTemplate.Planets.Add objPlanet
            
        'Position the First World as the primary
        Else
            Set objWorld = .HomeWorlds(1)
            objWorld.X = .X
            objWorld.Y = .Y
        End If
    End With
End Sub

Private Sub SetSecondaryWorlds(ByVal objRego As Registration)
    Dim wa As Single
    Dim ws As Single
    Dim wao As Single
    Dim W As Long
    Dim objWorld As HomeWorld
    Dim objPlanet As Planet
    
    With objRego
        If mobjTemplate.Seed(SeedHome) Then
            wa = 2 * PI / (.HomeWorlds.Count)
            ws = 1
        Else
            If .HomeWorlds.Count > 1 Then
                wa = 2 * PI / (.HomeWorlds.Count - 1)
            Else
                wa = 0
            End If
            ws = 2
        End If
        
        If mobjTemplate.OrbitPlanets Then
            wao = -wa * Rnd()
            For W = ws To objRego.HomeWorlds.Count
                Set objWorld = objRego.HomeWorlds(W)
                wao = wao + wa
                objWorld.X = Round(mobjTemplate.OrbitDistance * Cos(wao) + .X)
                objWorld.Y = Round(mobjTemplate.OrbitDistance * Sin(wao) + .Y)
            Next W
        
        Else
            For W = ws To .HomeWorlds.Count
                Set objWorld = objRego.HomeWorlds(W)
                Set objPlanet = EmptyWorld(objRego.X, objRego.Y, gcMaximumSecondaryRadius)
                objWorld.X = objPlanet.X
                objWorld.Y = objPlanet.Y
            Next W
        End If
    End With
End Sub

Private Sub SetEmptyWorlds(ByVal objRego As Registration)
    Dim objPlanet As Planet
    Dim p As Long
    
    With objRego
        'Seed empty planets
        For p = 1 To mobjTemplate.empty_planets
            Set objPlanet = EmptyWorld(objRego.X, objRego.Y)
            mobjTemplate.Planets.Add objPlanet
        Next p
    End With
End Sub

Private Sub SetStuffWorlds()
    Dim objRego As Registration
    Dim objPlanet As Planet
    Dim p As Long
    
    For Each objRego In mobjTemplate.Registrations
        'Seed empty planets
        For p = 1 To mobjTemplate.stuff_planets
            Set objPlanet = StuffWorld()
            mobjTemplate.Planets.Add objPlanet
        Next p
    Next objRego
End Sub

Private Sub SetAdditionalSeedings()
    Dim objPlanet As Planet
    Dim objRego As Registration
    Dim objRego1 As Registration
    Dim rl As Long
    Dim R As Long
    
    With mobjTemplate
        'Seed the centre of the galaxy
        If .DesignType = OnCircle _
        And .Seed(SeedCenter) Then
            Set objPlanet = StuffWorld
            objPlanet.X = Round(.Size / 2)
            objPlanet.Y = Round(.Size / 2)
            If mobjTemplate.SeedSize(SeedCenter) > 0 Then
                objPlanet.Size = mobjTemplate.SeedSize(SeedCenter)
            End If
            .Planets.Add objPlanet
        End If
        
        'Seed the way points for players on the circle
        If (.DesignType = OnCircle _
        Or .DesignType = OnCircleMiddle) _
        And .Seed(SeedWaypoint) Then
            If .DesignType = OnCircle Then
                rl = .Registrations.Count
            Else
                rl = .Registrations.Count - 1
            End If
            Set objRego1 = .Registrations(rl)
            
            For R = 1 To rl
                Set objRego = .Registrations(R)
                Set objPlanet = StuffWorld
                objPlanet.X = Round(objRego.X + (objRego1.X - objRego.X) / 2)
                objPlanet.Y = Round(objRego.Y + (objRego1.Y - objRego.Y) / 2)
                If mobjTemplate.SeedSize(SeedWaypoint) > 0 Then
                    objPlanet.Size = mobjTemplate.SeedSize(SeedWaypoint)
                End If
                .Planets.Add objPlanet
                
                Set objRego1 = objRego
            Next R
        End If
    End With
End Sub

Private Function EmptyWorld(ByVal X As Single, ByVal Y As Single, Optional ByVal lngDistance = -1) As Planet
    Dim objPlanet As Planet
    Dim a As Single
    Dim R As Single
    Dim c As Long
    
    
    If lngDistance <= 0 Then lngDistance = mobjTemplate.empty_radius
    
    Set objPlanet = New Planet
    With objPlanet
        c = gcMaxTries
        Do
            a = 2 * PI * Rnd()
            R = lngDistance * Rnd()
            .X = Round(Cos(a) * R + X)
            .Y = Round(Sin(a) * R + Y)
            If NearestPlanet(.X, .Y) >= gcMinimumPlanetDistance _
            Or c <= 0 Then
                .Size = Round(Rnd() * (gcSecondaryPlanetMaxSize - gcStuffMaxSize) + gcStuffMaxSize, 0)
                .Resources = Round(Rnd() * 10)
                Set EmptyWorld = objPlanet
                Exit Do
            End If
            c = c - 1
        Loop
    End With
End Function

Private Function StuffWorld() As Planet
    Dim objPlanet As Planet
    Dim c As Long
    
    Set objPlanet = New Planet
    With objPlanet
        c = gcMaxTries
        Do
            .X = Round(mobjTemplate.Size * Rnd())
            .Y = Round(mobjTemplate.Size * Rnd())
            If NearestPlanet(.X, .Y) >= gcMinimumPlanetDistance _
            Or c <= 0 Then
                .Size = Round(Rnd() * gcStuffMaxSize, 0)
                .Resources = Round(Rnd() * 10)
                Set StuffWorld = objPlanet
                Exit Do
            End If
            c = c - 1
        Loop
    End With
End Function

Private Function NearestPlanet(ByVal X As Single, ByVal Y As Single) As Single
' Locate the minimum Distance Squared and return the minimum distance
    Dim objRego As Registration
    Dim objHomeworld As HomeWorld
    Dim objPlanet As Planet
    Dim d2 As Single
    Dim D2Min As Single
    
    D2Min = 99999
    
    ' Check Home Worlds
    For Each objRego In mobjTemplate.Registrations
        For Each objHomeworld In objRego.HomeWorlds
            With objHomeworld
                If .X <> 0 And .Y <> 0 Then
                    d2 = (.X - X) ^ 2 + (.Y - Y) ^ 2
                    If d2 < D2Min Then D2Min = d2
                End If
            End With
        Next objHomeworld
    Next objRego
    
    ' Check Planets
    For Each objPlanet In mobjTemplate.Planets
        With objPlanet
            d2 = (.X - X) ^ 2 + (.Y - Y) ^ 2
            If d2 < D2Min Then D2Min = d2
        End With
    Next objPlanet
    
    NearestPlanet = Sqr(D2Min)
End Function
