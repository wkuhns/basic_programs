Attribute VB_Name = "Module1"
Option Explicit

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Mybot As RobotLink
' There are three subroutines that are automatically run by
' the robot, which may contain user code:
'    UserInit - executed once at startup.
'    Ping - executed by the server when this robot is scanned
'       by another robot
'    UserFrame - run continuously as long as the robot is
'       alive.
' Any other user created subroutines must be called from one
' of these.
'
' This is where you put anything that you want your robot to
' do continuously. You may create your own subroutines and
' call them from here. Depending on the complexity of your
' code, this routine is run about four times per second as
' long as your robot is alive.
'
' User defined variables:
' These are not required except as used by your application
'
Dim speed As Integer
Global dir As Integer
Dim scanres As Integer
Dim scandir As Single
Dim flight As Long
Dim reverse As Long
Dim ccw As Integer

' Information about a sighting of an enemy 'bot
Private Type sighting
    x As Single             ' Reported x,y
    y As Single
    x2 As Single            ' alternate x,y based on barrel heat
    y2 As Single
    t As Single
End Type

Private Type history
    distance As Integer
    depth As Integer
    vx As Single
    vy As Single
    s(3) As sighting
End Type

Dim enemies(4) As history
Dim closest As Integer          ' Enemy who is closest to us
Dim standoff As Single          ' Distance to nearest enemy
Dim btemp As Single             ' barrel temp
Dim lastbtime As Integer        ' time of last barrel temp calculation

Dim shells As Integer           ' remaining shells in this clip

Global newdir As Integer           ' The direction we want to go

Sub CheckDrive()
    
Dim x As Single
Dim y As Single
Dim heat As Integer

    ' change direction every 120 seconds. Hug walls.
    ' update x, y
    
    If reverse < Timer And flight < Timer Then
        reverse = Timer + 120
        ccw = ccw * -1
        newdir = (dir + 180) Mod 360
        MyForm.Timer2.Enabled = True
        Mybot.post "changed direction to " + Str(ccw)
    End If
    
    x = Mybot.x
    y = Mybot.y

    ' ccw alternates between 1 and -1. It's used as a multiplier
    ' later on.
        
    ' Check if we've reached first wall after start
    If ccw = 0 Then
        If x > 875 Or x < 125 Then ccw = -1
        If y > 875 Or y < 125 Then ccw = -1
        If ccw = -1 Then reverse = Timer + 120
    End If
       
    If ccw = -1 Then
        If (x > 875 And dir = 0) Then newdir = 90: scandir = 270: MyForm.Timer2.Enabled = True
        If (x < 125 And dir = 180) Then newdir = 270: scandir = 90: MyForm.Timer2.Enabled = True
        If (y > 875 And dir = 90) Then newdir = 180: scandir = 0: MyForm.Timer2.Enabled = True
        If (y < 125 And dir = 270) Then newdir = 0: scandir = 180: MyForm.Timer2.Enabled = True
    End If
    
    If ccw = 1 Then
        If (x > 875 And dir = 0) Then newdir = 270: scandir = 90: MyForm.Timer2.Enabled = True
        If (x < 125 And dir = 180) Then newdir = 90: scandir = 270: MyForm.Timer2.Enabled = True
        If (y > 875 And dir = 90) Then newdir = 0: scandir = 180: MyForm.Timer2.Enabled = True
        If (y < 125 And dir = 270) Then newdir = 180: scandir = 0: MyForm.Timer2.Enabled = True
    End If
    
    ' flight tells us when to stop running from a ping.
    
    If flight < Timer And speed = 100 Then
        speed = 35
    End If
    
    ' check motor heat and adjust speed
    
    heat = Mybot.heat
    
    If (heat > 150 And flight < Timer) Then speed = 25
    
    If (heat < 10 And speed < 35) Then speed = 35
    
    Call Mybot.Drive(dir, speed)

End Sub


Function Fire(b As Integer, r As Integer) As Integer
' fire cannon at b,r. return -1 for survey, 0 for click, 1 for shot

Dim stat As Integer
Dim mytime As Integer
Dim i As Integer

    mytime = Mybot.Time
    If Mybot.cannon(b Mod 360, r) = 0 Then
        Fire = 1
        Mybot.post "Bang"
        ' Update barrel temp calcs
        If btemp > 0 Then
            btemp = btemp - (mytime - lastbtime) * 0.2
            If btemp < 0 Then btemp = 0
        End If
        btemp = btemp + 20
        lastbtime = mytime
        ' If we're too hot, cool down
        If btemp > 35 Then
            Mybot.post "Cooling...."
            For i = 1 To (btemp - 35)
                DoEvents
                Sleep (400)
                CheckDrive
            Next i
            mytime = Mybot.Time
            btemp = btemp - (mytime - lastbtime) * 0.2
            lastbtime = mytime
        End If
        shells = shells - 1
        If shells = 0 Then
            shells = 4
            Mybot.post "Reloading"
            survey
            Fire = -1
        End If
    Else
        Mybot.post "Click"
        Fire = 0
    End If

End Function

Function GetBearing(x As Integer, y As Integer) As Single
        
' Calculate bearing from ourselves to x,y
Dim dx As Single
Dim dy As Single
Dim b As Single
    
    dx = x - Mybot.x
    If dx = 0 Then dx = 1
    dy = y - Mybot.y
    b = Atn(dy / dx) * 57.3 + 360
    If b > 360 Then b = b - 360
    If dx < 0 Then b = b + 180
    If b > 360 Then b = b - 360
    GetBearing = b
    
End Function

Function scan(b As Single, res As Integer) As Integer
    
Dim myrange As Integer
Dim enemy As Integer
Dim where As Long
Dim x As Integer
Dim y As Integer

    ' Look for enemy
    
    myrange = Mybot.scan(b, res)
    
    ' Keep track of who is near
    If myrange > 0 Then
        enemy = Mybot.dsp
        where = Mybot.WhereIs(enemy)
        y = where Mod 1000
        x = (where - y) / 1000
        enemies(enemy).distance = myrange
        ' Is this guy closest to us?
        If enemy = closest Or myrange < standoff Then
            standoff = myrange
            closest = enemy
        End If
        If res < 5 Then
'            mybot.pause
            PushEnemy enemy, myrange, b, 0
        Else
            enemies(0).distance = myrange
        End If
    End If

    scan = myrange
    
End Function

Function PredictBearing(e As Integer, t As Integer) As Single
' predict where enemy 'e' will be at time 't'

Dim tx As Integer
Dim ty As Integer
Dim dt As Single
    
    dt = t - enemies(e).s(0).t
    
    tx = enemies(e).s(0).x + enemies(e).vx * dt
    ty = enemies(e).s(0).y + enemies(e).vy * dt
    
    PredictBearing = GetBearing(tx, ty)
       
End Function

' add an enemy sighting of enemy e at range r
' and bearing b
Sub PushEnemy(e As Integer, r As Integer, b As Single, Shoot As Integer)

Dim i As Integer
Dim dx As Long
Dim dy As Long
Dim dx2 As Long
Dim dy2 As Long
Dim dt As Integer
Dim tx As Single
Dim ty As Single
Dim tx2 As Single
Dim ty2 As Single
Dim hsq As Long
Dim x As Integer
Dim y As Integer
Dim xs As Single
Dim mytime As Integer

    mytime = Mybot.Time
    
    tx = Mybot.x + r * Cos(b / 57.3)
    ty = Mybot.y + r * Sin(b / 57.3)
    
    If enemies(e).depth > 0 Then
    
        dt = mytime - enemies(e).s(0).t
        
        ' check barrel temperature
        
        btemp = btemp - (mytime - lastbtime) * 0.2
        If btemp < 0 Then btemp = 0
        
        ' calculate target x,y. If barrel is hot, there's an error.
        ' try to pick the right possibility.
        If btemp > 0 Then
            tx2 = Mybot.x + (r + btemp) * Cos(b / 57.3)
            ty2 = Mybot.y + (r + btemp) * Sin(b / 57.3)
            ' make a mark on the arena
            Mybot.Mark Int(tx2), Int(ty2), RGB(0, 255, 255)
            tx = Mybot.x + (r - btemp) * Cos(b / 57.3)
            ty = Mybot.y + (r - btemp) * Sin(b / 57.3)
            ' make a mark on the arena
            Mybot.Mark Int(tx), Int(ty), RGB(255, 255, 0)
            dx = tx - enemies(e).s(0).x
            dy = ty - enemies(e).s(0).y
            dx2 = tx2 - enemies(e).s(0).x
            dy2 = ty2 - enemies(e).s(0).y
            'if velocity based on tx2/ty2 is lower and velocity based
            'on tx/ty is too high, then use tx2/ty2 instead.
            If ((Abs(dx) > Abs(dx2)) And (Abs(dx) > 2 * dt)) Or _
            (Abs(dy) > Abs(dy2)) And (Abs(dy) > 2 * dt) Then
                tx = tx2
                ty = ty2
            End If
        End If
    
        If dt <= 2 Then
            dx = tx - enemies(e).s(0).x
            dy = ty - enemies(e).s(0).y
            ' Impossibly fast? average them out.
            If dx > 2 * dt Then
                xs = dx - 2 * dt
                tx = tx - xs / 2
                enemies(e).s(0).x = enemies(e).s(0).x + xs / 2
            End If
            If dx < -2 * dt Then
                xs = dx + 2 * dt
                tx = tx - xs / 2
                enemies(e).s(0).x = enemies(e).s(0).x + xs / 2
            End If
            If dy > 2 * dt Then
                xs = dy - 2 * dt
                ty = ty - xs / 2
                enemies(e).s(0).y = enemies(e).s(0).y + xs / 2
            End If
            If dy < -2 * dt Then
                xs = dy + 2 * dt
                ty = ty - xs / 2
                enemies(e).s(0).y = enemies(e).s(0).y + xs / 2
            End If
            enemies(e).s(0).t = enemies(e).s(0).t + dt / 2
            Mybot.Mark Int(tx), Int(ty), RGB(255, 255, 255)
        
        Else    ' Scans weren't too close together
            ' Push down stack
            For i = enemies(e).depth To 1 Step -1
                enemies(e).s(i) = enemies(e).s(i - 1)
            Next i
            dx = tx - enemies(e).s(1).x
            dy = ty - enemies(e).s(1).y
            dt = mytime - enemies(e).s(1).t
            ' Impossibly fast? average them out.
            If dx > 2 * dt Then
                xs = dx - 2 * dt
                tx = tx - xs / 2
                enemies(e).s(1).x = enemies(e).s(1).x + xs / 2
            End If
            If dx < -2 * dt Then
                xs = dx + 2 * dt
                tx = tx - xs / 2
                enemies(e).s(1).x = enemies(e).s(1).x + xs / 2
            End If
            If dy > 2 * dt Then
                xs = dy - 2 * dt
                ty = ty - xs / 2
                enemies(e).s(1).y = enemies(e).s(1).y + xs / 2
            End If
            If dy < -2 * dt Then
                xs = dy + 2 * dt
                ty = ty - xs / 2
                enemies(e).s(1).y = enemies(e).s(1).y + xs / 2
            End If
            Mybot.Mark Int(tx), Int(ty), RGB(255, 255, 255)
            enemies(e).s(0).t = mytime
            ' increment depth
            If enemies(e).depth < 3 Then
                enemies(e).depth = enemies(e).depth + 1
            End If
        End If  ' end if scans were / weren't too close
        enemies(e).s(0).x = tx
        enemies(e).s(0).y = ty
        dx = enemies(e).s(0).x - enemies(e).s(1).x
        dy = enemies(e).s(0).y - enemies(e).s(1).y
        dt = enemies(e).s(0).t - enemies(e).s(1).t
        If dt = 0 Then Stop
        enemies(e).vx = dx / dt
        enemies(e).vy = dy / dt
        ' we *should* check and average as above...
        If enemies(e).vx > 1.2 Then enemies(e).vx = 1.2
        If enemies(e).vx < -1.2 Then enemies(e).vx = -1.2
        If enemies(e).vy > 1.2 Then enemies(e).vy = 1.2
        If enemies(e).vy < -1.2 Then enemies(e).vy = -1.2
    Else
        ' depth was 0
        enemies(e).s(0).t = mytime
        enemies(e).depth = enemies(e).depth + 1
        enemies(e).s(0).x = tx
        enemies(e).s(0).y = ty
        enemies(e).vx = 0
        enemies(e).vy = 0
    End If  ' end if depth > 0
    enemies(e).distance = r
    Mybot.Mark Int(tx), Int(ty), RGB(255, 255, 255)
        
' reset depth to ignore data over 80 ticks old
'    For i = 1 To enemies(e).depth - 1
'        If (t - enemies(e).s(i).t) > 80 Then
'            Str (enemies(e).s(i).t)
'            enemies(e).depth = i
'            Exit For
'        End If
'    Next i
    
End Sub
Function Shoot(e As Integer, b As Single, r As Integer) As Integer
' shoot enemy e. return -1 for survey, 0 for click, 1 for shot

Dim dx As Long
Dim dy As Long
Dim tx As Single
Dim ty As Single
Dim tof As Integer
Dim hsq As Long
Dim mytime As Integer

    ' time is in 100ms ticks.
    mytime = Mybot.Time
    If enemies(e).depth > 1 Then
        tof = r / 20
        tx = enemies(e).s(0).x + enemies(e).vx * (mytime + tof - enemies(e).s(0).t)
        ty = enemies(e).s(0).y + enemies(e).vy * (mytime + tof - enemies(e).s(0).t)
        b = GetBearing(Int(tx), Int(ty))
        dx = tx - Mybot.x
        dy = ty - Mybot.y
        hsq = dx * dx + dy * dy
        r = Sqr(hsq)
        Mybot.Mark Int(tx), Int(ty), RGB(0, 0, 0)
    End If
    Shoot = Fire(Int(b), r)

End Function

Sub survey()

    Dim myscandir As Single
    Dim mytime As Integer
    Dim range As Integer
    Dim enemy As Integer
    
    mytime = Mybot.Time + 110
    If closest <> 0 Then
        myscandir = (GetBearing(Int(enemies(closest).s(0).x), _
            Int(enemies(closest).s(0).y)) + 330) Mod 360
    Else
        myscandir = (dir + 180) Mod 360
    End If
    
    standoff = 2000
    
    While Mybot.Time < mytime
        range = scan(myscandir, 10)
        myscandir = (myscandir + 360 + 15) Mod 360
        CheckDrive
    Wend
    Mybot.post "Closest is " + Str(closest) + " at " + Str(standoff) + " meters"
    
    scandir = PredictBearing(closest, Mybot.Time)
    
End Sub

' Verify location of enemy e
Function tickle(e As Integer) As Integer

Dim b As Single
Dim range As Integer
Dim t As Integer

        b = PredictBearing(e, Mybot.Time)
 
        range = scan(b, 10)
        
        ' did we see anyone?
        If range > 0 Then
            t = Mybot.dsp
            ' was it who we expected?
            If e = t Then
                tickle = 1
            Else
                tickle = 0
            End If
        Else
            tickle = 0
        End If

End Function

'
' This is where you put anything that you want your robot to
' do continuously. You may create your own subroutines and
' call them from here. Depending on the complexity of your
' code, this routine is run about four times per second as
' long as your robot is alive.
'
Public Sub UserFrame()

' Perform a cycle of calculations for our robot.

Dim range As Integer
Dim enemy As Integer

    CheckDrive
    
    range = scan(scandir, 10)
    
    If range > 40 And range < 700 Then
        ' If he's really close, just shoot
        If range < 150 Then
            Call Fire(Int(scandir), range)
        End If
        ' In any event, go into attack mode
        attack (scandir)
        scandir = (scandir + 360 - (ccw * 15)) Mod 360
    End If
    
    ' Check if we've scanned past our left shoulder
    If ccw <> 0 Then
        If Abs(scandir - dir) < 11 Then
            ' reset scan to behind us
            scandir = (dir + 180) Mod 360
        Else
            scandir = (scandir + 360 + (ccw * 15)) Mod 360
        End If
    Else
        scandir = (scandir + 360 + (15)) Mod 360
    End If
        
End Sub

Sub attack(targetdir As Single)
    
Dim myscan As Single
Dim myres As Integer
Dim startscan As Single
Dim endscan As Single
Dim range As Single
Dim delta As Single
Dim t As Integer
Dim enemy As Integer
Dim edir As Single

    enemy = Mybot.dsp
    edir = targetdir
    myscan = edir
    myres = 5
    delta = 0.1
    Mybot.post "Attack at " + Str(myscan)
    While delta < 15
        ' scan at 4 degrees
        myscan = edir + delta
        If myscan > 360 Then myscan = myscan - 360
        range = scan(myscan, 4)
        ' if we have him, zero in
        If (range > 40 And range < 700) Then
            delta = 0.1
            myres = 2
            edir = myscan
            While delta < 8
                myscan = edir + delta
                If myscan > 360 Then myscan = myscan - 360
                range = scan(myscan, 1)
                If (range > 40 And range < 700) Then
                    If Shoot(enemy, myscan, Int(range)) = -1 Then
                        Mybot.post "Calling off attack"
                        Exit Sub
                    End If
                    delta = 0.1
                    edir = myscan
                Else
                    delta = (Abs(delta) + 1.5) * -1 * Sgn(delta)
                End If
'                edir = PredictBearing(enemy, mybot.Time)
                CheckDrive
            Wend
        Else
            delta = (Abs(delta) + 3) * -1 * Sgn(delta)
        End If
'        edir = PredictBearing(enemy, mybot.Time)
        CheckDrive
    Wend
    speed = 35

End Sub
'
' This subroutine MUST exist with EXACTLY this name and
' this argument list. The contents are up to the user.
' This subroutine is invoked by the server when this robot
' is scanned by another robot.
'
Public Sub Ping(m As Integer)

    ' check if we know where he is. If too far, don't run.
        
    If enemies(m).depth > 0 Then
        If enemies(m).distance < 750 Then
            speed = 100
            flight = Timer + 4
            Mybot.Drive dir, speed
        Else
            ' He can't hurt us.
        End If
    Else
        ' We don't know where he is, so run
        speed = 100
        flight = Timer + 4
        Mybot.Drive dir, speed
    End If
    
End Sub



Sub UserInit()

Dim closewall As Integer
Dim t As Integer
Dim r As Integer
Dim x As Integer
Dim y As Integer

While x = 0 Or y = 0
    x = Mybot.x
    y = Mybot.y
    DoEvents
Wend

If x < 500 Then
    r = 0
    closewall = x
Else
    r = 1
    closewall = 999 - x
End If
If y < 500 Then
    ' We're in the bottom half.
    If y < closewall Then
        ' We're closer to the bottom than sides
        newdir = 270
    Else
        If r Then
            newdir = 0
        Else
            newdir = 180
        End If
    End If
Else
    ' We're in the top half
    If (999 - y) < closewall Then
        ' We're closer to the top than sides
        newdir = 90
    Else
        If r Then
            newdir = 0
        Else
            newdir = 180
        End If
    End If
End If

Mybot.post Str(x) + "," + Str(y) + " " + Str(newdir) + " " + Str(closewall)
speed = 35
Mybot.SetName ("Killer")
scanres = 10
ccw = 0
shells = 4
reverse = Timer + 120
closest = 0
standoff = 2000

End Sub




