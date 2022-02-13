Attribute VB_Name = "UnitCode"
'Unit behavior is as follows: Each tick, each unit is checked to see if it's currently engaged in
'combat or moving. If so, 'ProcessMove' is invoked for that unit. This checks each enemy unit to
'see if it's visible. If so, the sighting is processed, attack and follow routines are
'executed, and the unit is marked as 'engaged'.
'
'
'
'
'


Public Sub AddOrder(unit As unitstruct, cmd As String, n1 As Single, n2 As Single)

    If unit.ocount < 10 Then
        unit.orders(unit.ocount).command = cmd
        unit.orders(unit.ocount).n1 = n1
        unit.orders(unit.ocount).n2 = n2
        unit.ocount = unit.ocount + 1
    End If

End Sub
Public Sub AddSighting(unit As unitstruct)

    Static msg As String
    Static i As Integer

    If unit.side = side Then
        Exit Sub
    End If

    msg = Format$(GlobalTime, "00\:00\:00 ")
    msg = msg + Format$(specs(unit.type).name, "!@@@@@@@@@@@") + "at "
    msg = msg + Format$(unit.x, "00") + ", "
    msg = msg + Format$(unit.y, "00")

    ' If there are any sightings of the same type at the same spot,
    ' remove them.

    For i = 0 To MapForm!SightList.ListCount - 1
        If Right$(msg, 20) = Right$(MapForm!SightList.List(i), 20) Then
            MapForm!SightList.RemoveItem (i)
            Exit For
        End If
    Next i

    MapForm!SightList.AddItem msg
    MapForm!SightList.Text = msg
    
    If side = us Then
        AddDisplayItem "S", Int(unit.x), Int(unit.y), DSQUARE, RED, WHITE, Left$(specs(army(1, unit.Index).type).name, 1)
    Else
        AddDisplayItem "S", Int(unit.x), Int(unit.y), DSQUARE, RED, WHITE, Left$(specs(army(0, unit.Index).type).name, 1)
    End If

End Sub
Public Sub CheckBattle(a As unitstruct, b As unitstruct)

    Static chance As Single
    Static impact As Single
    Static xdist As Single
    Dim dx As Single
    Dim dy As Single
    Dim dist As Single

    dx = a.x - b.x
    dy = a.y - b.y
    dist = Sqr(dx * dx + dy * dy)                   ' units are squares

    xdist = dist * 10                       ' distance in tenths

    If xdist <= 0 Then xdist = 1            ' avoid pesky divide by 0

    ' a attacks b. wrange is in tenths...
    If xdist < specs(a.type).wrange * a.attack / 100 Then
        a.engaged = True
        b.engaged = True
        
        chance = specs(a.type).accuracy / 100
        chance = (xdist - (xdist * chance)) / specs(a.type).wrange + chance

        If chance > Rnd Then
            hit a, b
        End If

        ' b survives, attacks a. If remote, other computer will do it.

        If (remote = 0) And (b.health > 0) And (xdist < specs(b.type).wrange * b.attack / 100) Then

            chance = specs(b.type).accuracy / 100
            chance = (xdist - (xdist * chance)) / specs(b.type).wrange + chance

            If chance > Rnd Then
                hit b, a
            Else
                WriteCCC "Unit " + Str$(b.Index) + " missed"
            End If
        End If
    End If

End Sub

Public Sub CheckFollow(u As unitstruct, t As unitstruct)

    Dim dx As Single
    Dim dy As Single

    If (u.health <= 0) Or (t.health <= 0) Then
        Exit Sub
    End If

    ' 'u' is us (the unit that sighted the enemy) and 't' is them. 't' may not be able to see 'u'
    
    ' if our desire to follow is greater than the superiority of their armor over our weapons..

    If u.follow And (specs(t.type).air = 0) And (u.follow > (specs(t.type).armor - specs(u.type).wstrength + 50)) Then
        If u.orders(0).command <> "X" Then
            InsertOrder u, "X", t.x, t.y
        Else
            u.orders(0).n1 = t.x
            u.orders(0).n2 = t.y
        End If
    End If

    ' if our desire to retreat is greater than the superiority of our armor over their weapons,
    ' or if our desire to retreat is greater than our remaining health...

    If u.Retreat And (u.Retreat > (specs(u.type).armor - specs(t.type).wstrength + 50)) Or (u.health < u.Retreat) Then
        dx = u.x - (t.x - u.x)
        dy = u.y - (t.y - u.y)
        If dx < 0 Then dx = 0
        If dy < 0 Then dy = 0
        If dx > axis - 1 Then dx = axis - 1
        If dy > axis - 1 Then dy = axis - 1
        If u.orders(0).command <> "X" Then
            InsertOrder u, "X", dx, dy
        Else
            u.orders(0).n1 = dx
            u.orders(0).n2 = dy
        End If
'       WriteCCC "Unit " + Str$(u.index) + " of army " + Str$(u.side) + " retreating from " + specs(t.type).name + Str$(t.index)
    End If

End Sub

Public Sub DestroyUnit(unit As unitstruct)

    Static x As Integer
    Static y As Integer
    Static i As Integer

    x = Int(unit.x)
    y = Int(unit.y)

    If unit.side = side Then
        WriteCCC "Our " + specs(unit.type).name + " " + Str$(unit.Index) + " at " + Str$(x) + ", " + Str$(y) + " destroyed"
        specs(unit.type).count = specs(unit.type).count - 1
        UpdateDispBoxes
    Else
        WriteCCC "Enemy " + specs(unit.type).name + " " + Str$(unit.Index) + " at " + Str$(x) + ", " + Str$(y) + " destroyed"
    End If
    
    For i = 0 To dlcount - 1            ' do we have a dl item for him?
        If (unit.side = side) And (unit.Index = dl(i).unit) Then
            RemoveDisplayItem (i)
        End If
    Next i
    
    unit.health = 0
    unit.x = 0
    unit.y = 0

End Sub

Public Sub hit(a As unitstruct, b As unitstruct)
        
    Static impact As Single
    
    ' a has hit b

    impact = specs(a.type).wstrength / specs(b.type).armor
        
    ' impact is ratio of weapon superiority over armor strength. A ratio of 2.5
    ' results in destruction, and .5 means no damage.

    impact = (impact - 0.5) * 100

    ' fuzz it up a bit

    impact = impact + (Rnd * impact) - (impact * 0.5)

    If impact <= 0 Then
        Exit Sub
    End If
    
    If remote <> 0 And b.side = THEM Then
        SendHitInfo b, impact
        b.changed = False
    Else
        b.health = b.health - impact
        b.changed = True
        WriteCCC "Unit " + Str$(b.Index) + " suffered " + Format$(impact, "####") + " damage"
    End If

    If b.health <= 0 Then
        b.health = 0
        DestroyUnit b
    End If

End Sub

Public Sub InsertOrder(unit As unitstruct, cmd As String, n1 As Single, n2 As Single)

    Dim i As Integer

    If unit.ocount = 10 Then unit.ocount = 9
    i = unit.ocount
    While i > 0
        unit.orders(i).command = unit.orders(i - 1).command
        unit.orders(i).n1 = unit.orders(i - 1).n1
        unit.orders(i).n2 = unit.orders(i - 1).n2
        i = i - 1
    Wend
    unit.orders(0).command = cmd
    unit.orders(0).n1 = n1
    unit.orders(0).n2 = n2
    unit.ocount = unit.ocount + 1

End Sub
Function MoveUnit(unit As unitstruct) As Integer

    ' returns TRUE if unit actually moves, FALSE otherwise

    Static speed As Single
    Static dx As Single
    Static dy As Single
    Static dist As Single
    Static xs As Single
    Static ys As Single
    Static altitude As Single

    xs = unit.x
    ys = unit.y
        
    ' derate speed based on altitude
    ' calculate multiplier...

    MoveUnit = True

    If unit.health < 10 Then
        MoveUnit = False
        Exit Function
    End If

    ' speed at this point is a multiplier, with 1 = 100%

    speed = 1 - terrain(unit.x, unit.y).d / 10
    ' if we're an airplane, altitude doesn't matter
    If specs(unit.type).air = 1 Then
        speed = 1
    End If

    If speed < 0.05 Then speed = 0.05
    
    speed = speed * specs(unit.type).speed
        
    ' derate speed if health is less than 70%
    If unit.health < 70 Then
        speed = speed * (unit.health - 10) / 60
    End If
        
    
    unit.speed = speed
    unit.changed = True

    speed = speed / 1000
    dx = (unit.dx - xs)
    dy = (unit.dy - ys)
    dist = Sqr(dx * dx + dy * dy)
    If speed > dist Then
        unit.x = unit.dx
        unit.y = unit.dy
        unit.speed = 0
    Else
        unit.x = xs + (speed / dist) * dx
        unit.y = ys + (speed / dist) * dy
    End If
    SendUnitInfo unit

End Function
Public Sub ProcessMove(unit As unitstruct)

    ' unit moved. Can he now see any enemy units, and can enemy units see him?
    
    Static u As Integer
    Static v As Integer
    Static sight As Integer
    Dim we As Integer
    Dim they As Integer
    
    we = unit.side
    If we = 0 Then
        they = 1
    Else
        they = 0
    End If
    
    unit.engaged = False
    For u = 0 To asize - 1
        If PeekAtUnit(unit, army(they, u)) Then
            AddSighting army(they, u)
            CheckBattle unit, army(they, u)
            CheckFollow unit, army(they, u)
        End If
    
        ' now, give enemy a chance to see us...
        If PeekAtUnit(army(they, u), unit) Then
            AddSighting unit
            CheckBattle army(they, u), unit
            CheckFollow army(they, u), unit
        End If
    
    Next u

End Sub
' Can unit A see unit B?
Public Function PeekAtUnit(a As unitstruct, b As unitstruct) As Boolean

Dim dx As Single
Dim dy As Single
Dim dist As Single
Dim v As Integer
Dim Height As Single


    ' is sighting possible? check to twice nominal range
    v = specs(a.type).vision / 5                    ' units are squares
    PeekAtUnit = False
    If a.health > 0 And b.health > 0 And Abs(b.x - a.x) < v And Abs(b.y - a.y) < v Then
        v = v / 2                                       ' correct to units
        dx = a.x - b.x
        dy = a.y - b.y
        dist = Sqr(dx * dx + dy * dy)                   ' units are squares
        ' factor in stealth of enemy
        v = v * (100 - specs(b.type).stealth) / 100
        If (specs(b.type).air = 0) And (specs(a.type).air = 0) Then
            ' height confers an advantage
            Height = terrain(a.x, a.y).a
            Height = Height - terrain(b.x, b.y).a
            v = (v / 2) + (Height / 20)
        End If
        If dist < v Then
            PeekAtUnit = True
        End If
    End If

End Function

Public Sub OrderProcess(unit As unitstruct)

    ' if he's moving, see if he's at destination

    If unit.health = 0 Then
        Exit Sub
    End If

    If (unit.orders(0).command = "M") Or (unit.orders(0).command = "X") Then
        If (unit.x = unit.orders(0).n1) And (unit.y = unit.orders(0).n2) Then
            RemoveOrder unit
        End If
    End If

    If unit.ocount > 0 Then
        Select Case unit.orders(0).command
            Case "M"
                unit.dx = unit.orders(0).n1
                unit.dy = unit.orders(0).n2
            Case "X"
                unit.dx = unit.orders(0).n1
                unit.dy = unit.orders(0).n2
            Case "W"
                If GlobalCommand = unit.orders(0).n1 Then
                    RemoveOrder unit
                End If
            Case "A"
                unit.attack = unit.orders(0).n1
                RemoveOrder unit
            Case "R"
                unit.Retreat = unit.orders(0).n1
                RemoveOrder unit
            Case "F"
                unit.follow = unit.orders(0).n1
                RemoveOrder unit
            Case "C"
                unit.camo = unit.orders(0).n1
                RemoveOrder unit
            Case "A"
                unit.camo = 1
                RemoveOrder unit
            Case "S"
                GlobalCommand = unit.orders(0).n1
                RemoveOrder unit
        End Select
    End If

End Sub
Public Sub RemoveOrder(unit As unitstruct)
        
    ' collapse a unit's orders by removing the first one
    Static i As Integer
    
    i = 0
        
    While i < unit.ocount - 1
        unit.orders(i) = unit.orders(i + 1)
        i = i + 1
    Wend
    unit.ocount = unit.ocount - 1

End Sub

