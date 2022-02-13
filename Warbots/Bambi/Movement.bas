Attribute VB_Name = "Movement"
Option Explicit

Private Type loc
    alive As Boolean
    score As Integer
    x As Single
    y As Single
End Type

'
' Plan next ten seconds movement. Use following strategy:
' 1) Try to have exactly one opponent in range
' 2) Try to have in-range opponent(s) about 600M away
' To accomplish this, project each opponent's position in ten seconds.
' Compute eight possible destinations for ourselves.
' Score each destination based on criteria above.
'
Sub PlanMove1()

Dim badguys(1 To 4) As loc
Dim dest(0 To 7) As loc
Dim i As Integer
Dim j As Integer
Dim dist As Single
Dim inrange As Single
Dim bestscore As Integer
Dim bestdest As Integer

' Don't run too often. We're called from many places.
If MyBot.Time < nextcourse Then Exit Sub
nextcourse = MyBot.Time + 5
' Predict bad guy locations. For now, ignore movement.
'MyBot.pause
i = i
For i = 1 To 4
    If enemies(i).alive Then
        badguys(i).x = enemies(i).lastx ' + enemies(i).vx * (Time - enemies(i).lastseen + 2)
        badguys(i).y = enemies(i).lasty ' + enemies(i).vy * 10
        'Call MyBot.Mark(Int(badguys(i).x), Int(badguys(i).y), RGB(0, 0, 0))
        'badguys(i).x = badguys(i).x + enemies(i).vx * (Time - enemies(i).lastseen + 2)
        'badguys(i).y = badguys(i).y + enemies(i).vy * (Time - enemies(i).lastseen + 2)
        'Call MyBot.Mark(Int(badguys(i).x), Int(badguys(i).y), RGB(255, 255, 255))
        If badguys(i).x > 999 Then badguys(i).x = 999
        If badguys(i).x < 0 Then badguys(i).x = 0
        If badguys(i).y > 999 Then badguys(i).y = 999
        If badguys(i).y < 0 Then badguys(i).y = 0
    End If
    badguys(i).alive = enemies(i).alive
Next i

' determine possible destinations
For i = 0 To 7
    dest(i).alive = True
    dest(i).score = 2000
    dest(i).x = MyBot.x + 350 * Cos(i * 45 / 57.3)
    dest(i).y = MyBot.y + 350 * Sin(i * 45 / 57.3)
    If dest(i).x < 0 Or dest(i).x > 999 Then dest(i).score = 0
    If dest(i).y < 0 Or dest(i).y > 999 Then dest(i).score = 0
Next i

' Deduct points for enemies too close or too many or too few
bestscore = 0
For i = 0 To 7
    inrange = 0
    If dest(i).score > 0 Then
    'Call MyBot.Mark(Int(dest(i).x), Int(dest(i).y), RGB(255, 255, 0))
    For j = 1 To 4
        If badguys(j).alive Then
            dist = distance(dest(i).x, dest(i).y, badguys(j).x, badguys(j).y)
            If dist < 700 Then
                'MyBot.pause
                inrange = inrange + 1
            End If
            ' If dist > 700 Then dest(i).score = dest(i).score - (dist - 700) / 10
            If dist < 500 Then dest(i).score = dest(i).score - 25
            If dist < 350 Then dest(i).score = dest(i).score - 25
            If dist < 250 Then dest(i).score = dest(i).score - 25
        End If
    Next j
    ' Deduct points for too many or none in range
    If inrange = 3 Then dest(i).score = dest(i).score - 600
    If inrange = 2 Then dest(i).score = dest(i).score - 400
    If inrange = 0 Then dest(i).score = dest(i).score - 300
    If dest(i).score > bestscore Then
        bestscore = dest(i).score
        bestdest = i
    End If
    End If
Next i

If bestscore > 0 Then
    dir = bestdest * 45
End If

'MyBot.post "Best dir: " & Str(dir) & " score: " & Str(bestscore) & " " & Str(inrange)
End Sub
'
' Plan next five seconds movement. Use following strategy:
' 1) Try to have exactly one opponent in range
' 2) Try to have in-range opponent(s) about 600M away
' 3) Try to move near perpendicular to path from nearest enemy (missed shots
'    typically go long or short)
' 4) Try not to make abrupt course changes
' 25 possible destinations: 5x5 grid on arena. Discard those requiring more
' than 90 degree turn.
'
Sub PlanMove()

Dim badguys(1 To 4) As loc
Dim dest(0 To 9, 0 To 9) As loc
Dim x As Single
Dim y As Single
Dim x1 As Integer
Dim y1 As Integer
Dim x2 As Integer
Dim y2 As Integer
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim dist As Single
Dim inrange As Single
Dim living As Integer
Dim bestscore As Integer
Dim bestdest(2) As Integer
Dim bearing As Single
Dim myself As String
Dim nearest As Single
Dim choice1 As Single
Dim choice2 As Single

' Don't run too often. We're called from many places.
If MyBot.Time < nextcourse Then Exit Sub
nextcourse = MyBot.Time + 3

' Print map to file
'MyBot.pause

x = Int(MyBot.x / 100)
y = Int(MyBot.y / 100)
For j = 9 To 0 Step -1
    For i = 0 To 9
        inrange = 0
        myself = "-"
        For k = 1 To 4
            If enemies(k).alive Then
                If i = Int(enemies(k).lastx / 100) Then
                    If j = Int(enemies(k).lasty / 100) Then
                        inrange = inrange + 1
                    End If
                End If
            End If
        Next k
        If x = i And y = j Then myself = "X"
        'Print #1, " " & myself & Format(inrange, "##");
    Next i
    'Print #1,
Next j
'Print #1,

' Predict bad guy locations. For now, ignore movement.
'MyBot.pause
living = 0
For i = 1 To 4
    If enemies(i).alive And enemies(i).lastx <> 0 Then
        badguys(i).x = enemies(i).lastx ' + enemies(i).vx * (Time - enemies(i).lastseen + 2)
        badguys(i).y = enemies(i).lasty ' + enemies(i).vy * 10
        'Call MyBot.Mark(Int(badguys(i).x), Int(badguys(i).y), RGB(0, 0, 0))
        'badguys(i).x = badguys(i).x + enemies(i).vx * (Time - enemies(i).lastseen + 2)
        'badguys(i).y = badguys(i).y + enemies(i).vy * (Time - enemies(i).lastseen + 2)
        'Call MyBot.Mark(Int(badguys(i).x), Int(badguys(i).y), RGB(255, 255, 255))
        If badguys(i).x > 999 Then badguys(i).x = 999
        If badguys(i).x < 0 Then badguys(i).x = 0
        If badguys(i).y > 999 Then badguys(i).y = 999
        If badguys(i).y < 0 Then badguys(i).y = 0
    End If
    If enemies(i).alive Then living = living + 1
    badguys(i).alive = enemies(i).alive
Next i

' determine possible destinations
For j = 9 To 0 Step -1
    For i = 0 To 9
        dest(i, j).alive = True
        dest(i, j).score = 2000
        dest(i, j).x = i * 100 + 50
        dest(i, j).y = j * 100 + 50
        ' get delta bearing to cell
        bearing = Abs(MyBot.direction - GetBearing(Int(dest(i, j).x), Int(dest(i, j).y)))
        If bearing > 180 Then bearing = 360 - bearing
        ' Penalize abrupt course changes
        dest(i, j).score = dest(i, j).score - bearing
    Next i
Next j

' current cell is not valid
dest(Int(MyBot.x / 100), Int(MyBot.y / 100)).score = 0

' Deduct points for enemies too close or too many or too few
bestscore = 0
nearest = 1000

' Score every cell
x = MyBot.x
y = MyBot.y
For j = 9 To 0 Step -1
    For i = 0 To 9
        ' Don't look farther than 350 meters
        dist = distance(x, y, dest(i, j).x, dest(i, j).y)
        If dist > 350 Or dist < 125 Then dest(i, j).score = 0
        inrange = 0
        If dest(i, j).score > 0 Then
            'Call MyBot.Mark(Int(dest(i, j).x), Int(dest(i, j).y), RGB(255, 255, 0))
            For k = 1 To 4
                If badguys(k).alive Then
                    dist = distance(dest(i, j).x, dest(i, j).y, badguys(k).x, badguys(k).y)
                    If dist < 700 Then
                        'MyBot.pause
                        inrange = inrange + 1
                    End If
                    If dist < nearest Then
                        nearest = dist
                        bearing = GetBearing(Int(badguys(k).x), Int(badguys(k).y))
                    End If
                    If dist > 700 Then dest(i, j).score = dest(i, j).score - (dist - 700) / 10
                    If dist < 500 Then dest(i, j).score = dest(i, j).score - 25
                    If dist < 350 Then dest(i, j).score = dest(i, j).score - 25
                    If dist < 250 Then dest(i, j).score = dest(i, j).score - 25
                End If
            Next k
            ' Deduct points for too many or none in range
            If inrange = 3 Then dest(i, j).score = dest(i, j).score - 600
            If inrange = 2 Then dest(i, j).score = dest(i, j).score - 400
            If living = 1 Then
                If inrange = 0 Then dest(i, j).score = dest(i, j).score - 300
            Else
                If inrange = 1 Then dest(i, j).score = dest(i, j).score - 300
            End If
            If dest(i, j).score > bestscore Then
                bestscore = dest(i, j).score
                bestdest(0) = i
                bestdest(1) = j
            End If
        End If
        'Print #1, "  " & Str(dest(i, j).score) & Str(inrange);
    Next i
    'Print #1,
Next j

' if not too near edge, move perpendicular to nearest opponent

x = MyBot.x
y = MyBot.y

If x > 100 And x < 900 And y > 100 And y < 900 And nearest > 600 And living = 1 Then
    If nearest > 650 Then
        choice1 = (bearing + 70) Mod 360
        choice2 = (bearing + 290) Mod 360
    Else
        choice1 = (bearing + 110) Mod 360
        choice2 = (bearing + 250) Mod 360
    End If
    
    ' Choice based on best destination
    dir = GetBearing(Int(dest(bestdest(0), bestdest(1)).x), Int(dest(bestdest(0), bestdest(1)).y))
    
    x1 = MyBot.x + 200 * Sin(choice1 / 57.3)
    x2 = MyBot.x + 200 * Sin(choice2 / 57.3)
    y1 = MyBot.y + 200 * Cos(choice1 / 57.3)
    y2 = MyBot.y + 200 * Cos(choice2 / 57.3)
    
    If x1 > 0 And x1 < 999 And y1 > 0 And y1 < 999 Then
        dir = choice1
'MyBot.post "angle: " & Str(dest(bestdest(0), bestdest(1)).x) & ", " & Str(dest(bestdest(0), bestdest(1)).y) & " score: " & Str(bestscore) & " " & Str(dir)
'Print #1, "angle: " & Str(dest(bestdest(0), bestdest(1)).x) & ", " & Str(dest(bestdest(0), bestdest(1)).y) & " score: " & Str(bestscore) & " " & Str(dir)
    End If
    
    If x2 > 0 And x2 < 999 And y2 > 0 And y2 < 999 Then
        dir = choice2
'MyBot.post "angle: " & Str(dest(bestdest(0), bestdest(1)).x) & ", " & Str(dest(bestdest(0), bestdest(1)).y) & " score: " & Str(bestscore) & " " & Str(dir)
'Print #1, "angle: " & Str(dest(bestdest(0), bestdest(1)).x) & ", " & Str(dest(bestdest(0), bestdest(1)).y) & " score: " & Str(bestscore) & " " & Str(dir)
    End If
Else
    dir = GetBearing(Int(dest(bestdest(0), bestdest(1)).x), Int(dest(bestdest(0), bestdest(1)).y))
'MyBot.post "location: " & Str(dest(bestdest(0), bestdest(1)).x) & ", " & Str(dest(bestdest(0), bestdest(1)).y) & " score: " & Str(bestscore) & " " & Str(dir)
'Print #1, "location: " & Str(dest(bestdest(0), bestdest(1)).x) & ", " & Str(dest(bestdest(0), bestdest(1)).y) & " score: " & Str(bestscore) & " " & Str(dir)
End If

Print #1,

End Sub

Function distance(x As Single, y As Single, x1 As Single, y1 As Single) As Single

distance = Sqr((x - x1) ^ 2 + (y - y1) ^ 2)

End Function
