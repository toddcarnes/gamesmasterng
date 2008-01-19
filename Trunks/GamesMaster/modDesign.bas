Attribute VB_Name = "modDesign"
Option Explicit
Option Compare Text

Public Sub ApplyDesign(ByVal objTemplate As Template)

    If objTemplate.SeedType = NoSeeding Then
        If objTemplate.DesignType = DefaultDesign Then
            Call DesignDefault(objTemplate)
        ElseIf objTemplate.DesignType = OnCircle Then
            Call DesignCircle(objTemplate)
        ElseIf objTemplate.DesignType = OnCircleMiddle Then
            Call DesignCircleMiddle(objTemplate)
        End If
    Else
    End If
End Sub

Private Sub DesignDefault(ByVal objTemplate As Template)
    Dim objRego As Registration
    Dim objWorld As HomeWorld
    
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
    Dim r As Single
    Dim ro As Single
    Dim s As Single
    Dim Px As Single
    Dim Py As Single
    Dim w As Long
    Dim wa As Single
    Dim wao As Single
    
    
    'calculate the radius of the circle
    r = CalcRadius(objTemplate.Registrations.Count, objTemplate.race_spacing)
    
    'calculate the galaxy Size
    s = Int(r + objTemplate.empty_radius)
    If objTemplate.Size < s Then
        objTemplate.Size = s
    End If
    
    'Calculate the center of the circle
    ro = objTemplate.Size / 2
    
    ' calculate the angle between races
    a = 2 * PI / objTemplate.Registrations.Count
    
    ' calculate a random offset for the races
    ao = -a * Rnd(0)

    For rc = 1 To objTemplate.Registrations.Count
        Set objRego = objTemplate.Registrations(rc)
        Set objWorld = objRego.HomeWorlds(1)
        ao = ao + a
        Px = Round(r * Cos(ao) + ro)
        Py = Round(r * Sin(ao) + ro)
        objWorld.X = Px
        objWorld.Y = Py
    
        wa = 2 * PI / (objRego.HomeWorlds.Count - 1)
        wao = -wa * Rnd(0)
        For w = 2 To objRego.HomeWorlds.Count
            Set objWorld = objRego.HomeWorlds(w)
            If objTemplate.OrbitPlanets Then
                wao = wao + wa
                objWorld.X = Round(objTemplate.OrbitDistance * Cos(wao) + Px)
                objWorld.Y = Round(objTemplate.OrbitDistance * Sin(wao) + Py)
            Else
                objWorld.X = 0
                objWorld.Y = 0
            End If
        Next w
    Next rc
    
End Sub

Private Sub DesignCircleMiddle(ByVal objTemplate As Template)
' Populate the edge of the circle and put the last in the middle
    Dim objRego As Registration
    Dim objWorld As HomeWorld
    Dim rc As Long
    Dim a As Single
    Dim ao As Single
    Dim r As Single
    Dim ro As Single
    Dim s As Single
    Dim Px As Single
    Dim Py As Single
    Dim w As Long
    Dim wa As Single
    Dim wao As Single
    
    
    'calculate the radius of the circle
    r = CalcRadius(objTemplate.Registrations.Count - 1, objTemplate.race_spacing)
    
    'calculate the galaxy Size
    s = Int(r + objTemplate.empty_radius)
    If objTemplate.Size < s Then
        objTemplate.Size = s
    End If
    
    'Calculate the center of the circle
    ro = objTemplate.Size / 2
    
    ' calculate the angle between races
    a = 2 * PI / objTemplate.Registrations.Count
    
    ' calculate a random offset for the races
    ao = -a * Rnd(0)

    For rc = 1 To objTemplate.Registrations.Count
        Set objRego = objTemplate.Registrations(rc)
        Set objWorld = objRego.HomeWorlds(1)
        If rc < objTemplate.Registrations.Count Then
            ao = ao + a
            Px = Round(r * Cos(ao) + ro)
            Py = Round(r * Sin(ao) + ro)
        Else
            Px = Round(objTemplate.Size / 2)
            Py = Round(objTemplate.Size / 2)
        End If
        objWorld.X = Px
        objWorld.Y = Py
    
        wa = 2 * PI / (objRego.HomeWorlds.Count - 1)
        wao = -wa * Rnd(0)
        For w = 2 To objRego.HomeWorlds.Count
            Set objWorld = objRego.HomeWorlds(w)
            If objTemplate.OrbitPlanets Then
                wao = wao + wa
                objWorld.X = Round(objTemplate.OrbitDistance * Cos(wao) + Px)
                objWorld.Y = Round(objTemplate.OrbitDistance * Sin(wao) + Py)
            Else
                objWorld.X = 0
                objWorld.Y = 0
            End If
        Next w
    Next rc
    
End Sub

Private Function CalcRadius(ByVal lngPlayers As Long, ByVal sngSpacing As Single) As Single
    'Angle Between Players
    'a = (2 * PI) / lngPlayers
    '
    'x = sngspacing / 2
    'Radius = x / sin(a/2)
    CalcRadius = (sngSpacing / 2) / Sin(PI / lngPlayers)
End Function


