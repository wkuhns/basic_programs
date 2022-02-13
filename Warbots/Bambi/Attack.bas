Attribute VB_Name = "AttackCode"
Option Explicit

' We've been given an enemy to attack. We look where he last was,
' then scan outwards. We then calculate the tangent error based
' on the range and our scan resolution. If it's more than 15 meters,
' set scan resolution to a tighter number and exit. Leave 'attacking'
' flag set so we'll return right away.
' If we can't find him, call off attack and exit.
'
Sub Attack(enemy As Integer)
    
Dim myscan As Single
Dim myres As Integer
Dim startscan As Single
Dim endscan As Single
Dim range As Single
Dim delta As Single
'Dim t As Integer
Dim edir As Single
Dim edir2 As Single
Dim result As Integer

    ' set global status flag
    attacking = enemy
    edir = enemies(enemy).lastdir
    range = Scan(edir, Int(enemies(enemy).scanres))
    
    ' If we missed him, scan outwards
    
    If range = 0 Then
        edir2 = (edir + (enemies(enemy).scanres * 1.5) + 360) Mod 360
        range = Scan(edir2, Int(enemies(enemy).scanres))
    End If
    If range = 0 Then
        edir2 = (edir - (enemies(enemy).scanres * 1.5) + 360) Mod 360
        range = Scan(edir2, Int(enemies(enemy).scanres))
    End If
    If range = 0 Then
        edir2 = (edir + (enemies(enemy).scanres * 3) + 360) Mod 360
        range = Scan(edir2, Int(enemies(enemy).scanres))
    End If
    If range = 0 Then
        edir2 = (edir - (enemies(enemy).scanres * 3) + 360) Mod 360
        range = Scan(edir2, Int(enemies(enemy).scanres))
    End If
    
    ' All done scanning. Did we find him?
    If range > 0 And range < 700 Then
        ' if he's close enough, shoot
        If (range < (10 / Tan((enemies(enemy).scanres / 57.3)))) Or (enemies(enemy).scanres = 1) Then
            result = Shoot(enemy, edir, Int(range))
            'MyBot.post ("Close enough at " & Str(range) & ", " & Str(scanres) & ", " & Str(edir))
            If result = -1 Then
                ' If we're now reloading, call off attack
                attacking = 0
                enemies(enemy).scanres = 8
            End If
            If result = 1 Then
                ' If we fired a shot, call off attack
                attacking = 0
                enemies(enemy).scanres = 8
            End If
            Exit Sub
        Else
            ' We see him, but need more resolution
            scandir = edir2
            enemies(enemy).scanres = enemies(enemy).scanres / 2
        End If
    Else
        ' We lost him. Call off attack
        attacking = 0
        enemies(enemy).scanres = 8
        hunting = True
        closest = 0
        standoff = 2000
        MyBot.post "Attack: Hunting"
    End If
    
    ' Catch possible scanres=0 bug
    If enemies(enemy).scanres < 1 Then
        'MyBot.pause
        enemies(enemy).scanres = 1
    End If

End Sub

Function Shoot(e As Integer, b As Single, r As Integer) As Integer
' shoot enemy e. return -1 for survey, 0 for click, 1 for shot

Dim dx As Long
Dim dy As Long
Dim tx As Single
Dim ty As Single
Dim tof As Single
Dim hsq As Long
Dim mytime As Single
Dim b2 As Single

    mytime = MyBot.Time
    ' if we have two sightings that are recent enough, plot path.
    If enemies(e).depth > 1 And enemies(e).lastseen > (mytime - 6) Then
        tof = r / 200
        MyBot.Mark Int(enemies(e).s(1).x), Int(enemies(e).s(1).y), RGB(255, 255, 255)
        MyBot.Mark Int(enemies(e).s(0).x), Int(enemies(e).s(0).y), RGB(168, 168, 168)
        
        tx = enemies(e).s(0).x + enemies(e).vx * (mytime + tof - enemies(e).s(0).t)
        ty = enemies(e).s(0).y + enemies(e).vy * (mytime + tof - enemies(e).s(0).t)
        
        b = GetBearing(Int(tx), Int(ty))
        dx = tx - MyBot.x
        dy = ty - MyBot.y
        hsq = dx * dx + dy * dy
        r = Sqr(hsq)
        MyBot.Mark Int(tx), Int(ty), RGB(0, 0, 0)
    End If
        
    Shoot = Fire(Int(b), r)

End Function
Function Fire(b As Integer, r As Integer) As Integer
' fire cannon at b,r. return -1 for survey, 0 for click, 1 for shot

Dim stat As Integer
Dim mytime As Single
Dim i As Integer
Dim prevdir As Integer

    ' Catch negative range bug
    If r < 0 Then
        'MyBot.pause
        r = r
    End If
    
    mytime = MyBot.Time
    ' Call EZPlay("P:\Warbots\Bambi\gunshot.wav", ssFile)
    If MyBot.cannon(b, r) = 0 Then
        Fire = 1
        'MyBot.post "Fired " & Str(b) & ", " & Str(r)
        ' Update barrel temp calcs
        If btemp > 0 Then
            btemp = btemp - (mytime - lastbtime) * 2
            If btemp < 0 Then btemp = 0
        End If
        btemp = btemp + 20
        lastbtime = mytime
        ' If we're too hot, cool down
        While btemp > 35
            prevdir = dir
            If Timer > flight And standoff > 400 Then speed = 35
            Call PlanMove
            If dir <> prevdir Or MyBot.speed <> speed Then
                Call MyBot.Drive(dir, speed)
            End If
            mytime = MyBot.Time
            btemp = btemp - (mytime - lastbtime) * 2
            lastbtime = mytime
        Wend
        shells = shells - 1
        If shells = 0 Then
            shells = 4
            MyBot.post "Reloading"
            reloading = mytime + 12
            Fire = -1
        Else
            ' didn't empty clip - just 4 second reload
            reloading = mytime + 4
        End If
    Else
        ' Shouldn't happen....
        MyBot.post "Click"
        Fire = 0
    End If

End Function


