Attribute VB_Name = "ControlCode"
Option Explicit
' This is our main control loop. Each pass through, we want to do a
' list of useful things:
'   - Check whether we're running from a hostile ping (flight)
'   - Reassess our direction of movement and adjust if needed
'   - Depending on whether we can shoot or not, scan for enemies
'     or attack.
' By design, we should run through this list no less than once every
' two seconds or so.
'
Public Sub UserFrame()

Dim prevdir As Integer
Dim i As Integer
Dim alive As Integer
Dim known As Integer

    ' Check if direction or speed need to be changed.
    prevdir = dir
    
    ' Are we still running?
    speed = 75
    If MyBot.heat > 100 Then speed = 35
    If standoff < 400 Then speed = 100
    If Timer < flight Then speed = 100
    
    ' Do we need to change direction?
    If MyBot.Time >= nextcourse Then
        Call PlanMove
    End If
    
    If dir <> prevdir Or speed <> MyBot.speed Then
        Call MyBot.Drive(dir, speed)
    End If

    ' Count how many enemies are alive and how many are known
    
    alive = 0
    known = 0
    For i = 1 To 4
        If enemies(i).alive Then
            alive = alive + 1
            If enemies(i).verified = True Then
                known = known + 1
            End If
        End If
    Next i
    
    ' If we've got time and we don't know where they are,
    ' try and locate more of the enemies.
    ' Otherwise, just keep track of them until we can shoot.
    If reloading > (MyBot.Time + 3) Then
        ' We're in a long reload cycle
        'MyBot.post "Long: " & Str(alive) & Str(known)
        If known = alive Then
            ' If we know where they all are, just verify
            Call verify(Int(scanres))
        Else
            ' If we don't know, look around a bit
            Call surveytick
        End If
        Exit Sub
    End If
    
    ' We don't have much time. Verify known enemies and prepare
    ' to attack. If we don't know of any, continue survey.
    If reloading > (MyBot.Time) Then
        ' we're less than two seconds from being able to fire.
        'MyBot.post "Short: " & Str(alive) & Str(known)
        If known > 0 Then
            Call verify(Int(scanres))
        Else
            Call surveytick
        End If
        Exit Sub
    End If
        
    ' If we're in the midst of an attack, continue.
    'MyBot.post "Ready: " & Str(alive) & Str(known)
    If attacking > 0 Then
        ' We're in the middle of an attack
        'MyBot.post "Re-attack " & Str(attacking)
        Attack (attacking)
    Else
        ' hunting means we need to find a target
        If hunting = True Then
            Call checkscan
        Else
            If closest <> 0 Then
                ' We have a target. Attack closest enemy
                Attack (closest)
            Else
                hunting = True
                standoff = 2000
                MyBot.post "Control: Hunting"
            End If
        End If
    End If

End Sub
Function PredictBearing(e As Integer, t As Single) As Single
' predict where enemy 'e' will be at time 't'

Dim tx As Integer
Dim ty As Integer
Dim dt As Single
    
    dt = t - enemies(e).s(0).t
    
    tx = enemies(e).s(0).x + enemies(e).vx * dt
    ty = enemies(e).s(0).y + enemies(e).vy * dt
    
    PredictBearing = GetBearing(tx, ty)
       
End Function

'
' We've been pinged by someone. Mark them as alive as of this moment.
' If we know where they are and they're more than 700 meters away, we don't
' need to run.
'
Public Sub Ping(m As Integer)

    ' check if we know where he is. If too far, don't run.
        
    enemies(Val(m)).lastseen = MyBot.Time
    enemies(Val(m)).alive = True
    If enemies(Val(m)).depth > 0 Then
        If enemies(m).lastrange < 750 Then
            speed = 100
            flight = Timer + 4
            MyBot.Drive dir, speed
        Else
            ' He can't hurt us.
        End If
    Else
        ' We don't know where he is, so run
        speed = 100
        flight = Timer + 4
        MyBot.Drive dir, speed
    End If
End Sub


Sub UserInit()

Dim i As Integer

MyBot.SetName ("Bambi")
flight = Timer

Open "p:\programming\warbots\bambi\strategery.txt" For Output As #1

For i = 1 To 4
    enemies(i).alive = True
    enemies(i).lastseen = MyBot.Time - 15
    enemies(i).scanres = 8
    enemies(i).verified = False
Next i

' Don't look for myself
enemies(MyBot.Index).alive = False

speed = 35
scanres = 8
scandir = 340

shells = 4
closest = 0
standoff = 2000
hunting = True
nextcourse = MyBot.Time + 5

Call MyBot.Drive(180, 100)


End Sub
Function GetBearing(x As Integer, y As Integer) As Single
        
' Calculate bearing from ourselves to x,y
Dim dx As Single
Dim dy As Single
Dim b As Single
    
    dx = x - MyBot.x
    If dx = 0 Then dx = 1
    dy = y - MyBot.y
    b = Atn(dy / dx) * 57.3 + 360
    If b > 360 Then b = b - 360
    If dx < 0 Then b = b + 180
    If b > 360 Then b = b - 360
    GetBearing = b
    
End Function

