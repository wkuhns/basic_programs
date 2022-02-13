Attribute VB_Name = "ScanCode"
Option Explicit
' Hunting mode. Perform intelligent scan increment based
' on location.
'
Sub checkscan()
Dim quad(4) As Integer
Dim i As Integer
Dim x As Integer
Dim y As Integer

x = MyBot.x
y = MyBot.y

scandir = (scandir + 15) Mod 360

For i = 0 To 3
    quad(i) = 1
Next i
    
' Left edge
If x < 100 Then
   quad(1) = 0
   quad(2) = 0
End If

' right
If x > 900 Then
   quad(0) = 0
   quad(3) = 0
End If

' top
If y > 900 Then
   quad(0) = 0
   quad(1) = 0
End If

' bottom
If y < 100 Then
   quad(2) = 0
   quad(3) = 0
End If

If quad(Int(scandir / 90)) = 0 Then
    While quad(Int(scandir / 90)) = 0
        scandir = (scandir + 10) Mod 360
    Wend
    scandir = (scandir + 350) Mod 360
End If

range = Scan(scandir, 8)

End Sub

' Process scan request. Update enemy location and status.
' If tangent error is low enough, update tracking data.
'
Function Scan(b As Single, res As Integer) As Integer
    
Dim myrange As Integer
Dim enemy As Integer
Dim where As Long
Dim x As Integer
Dim y As Integer

    ' Look for enemy
    
    myrange = MyBot.Scan(b, res)
    
    ' Keep track of who is near
    If myrange > 0 Then
        enemy = MyBot.dsp
        enemies(enemy).lastseen = MyBot.Time
        enemies(enemy).alive = True
        enemies(enemy).lastrange = myrange
        enemies(enemy).lastdir = b
        enemies(enemy).lastx = MyBot.x + myrange * Cos(b / 57.3)
        enemies(enemy).lasty = MyBot.y + myrange * Sin(b / 57.3)
        enemies(enemy).verified = True
        enemies(enemy).scanres = res
        ' Debugging code
        'where = MyBot.WhereIs(enemy)
        'y = where Mod 1000
        'x = (where - y) / 1000
        
        ' Enemy is within range. We don't need to keep hunting
        If myrange < 700 Then hunting = False
        
        ' Is this guy closest to us?
        If enemy = closest Or myrange < standoff Then
            standoff = myrange
            closest = enemy
        End If
        
        ' If max tangent error is less than 15 meters,
        ' keep sighting in array for future use
        'If myrange < (15 / Tan((res / 57.3))) Then
        If myrange < (12 / Tan((res / 57.3))) Then
            PushEnemy enemy, myrange, b
        End If
    End If

    Scan = myrange
    
End Function
' add an enemy sighting of enemy e at range r
' and bearing b
Sub PushEnemy(e As Integer, r As Integer, b As Single)

Dim i As Integer
Dim dx As Single
Dim dy As Single
Dim v11 As Single
Dim v12 As Single
Dim v21 As Single
Dim v22 As Single
Dim s11 As Single
Dim s12 As Single
Dim s21 As Single
Dim s22 As Single
Dim dt As Single
Dim tx1 As Single
Dim ty1 As Single
Dim tx2 As Single
Dim ty2 As Single
Dim hsq As Single
Dim best As Single
'Dim y As Integer
Dim xs As Single
Dim mytime As Single
Dim where As Long
Dim x As Integer
Dim y As Integer
Dim mspeed As Integer
Dim test As Single

    mytime = MyBot.Time
    
    ' Barrel temp introduces an error - look at both
    ' possible positions
    btemp = btemp - (mytime - lastbtime) * 2
    lastbtime = mytime
    If btemp < 0 Then btemp = 0
    
    tx1 = MyBot.x + (r + btemp) * Cos(b / 57.3)
    ty1 = MyBot.y + (r + btemp) * Sin(b / 57.3)
    
    tx2 = MyBot.x + (r - btemp) * Cos(b / 57.3)
    ty2 = MyBot.y + (r - btemp) * Sin(b / 57.3)
    
    ' reset depth to ignore data over 6 seconds old
    For i = enemies(e).depth - 1 To 0 Step -1
        If (enemies(e).s(i).t + 6) < mytime Then
            'MyBot.pause
            enemies(e).depth = i
        End If
    Next i

    If enemies(e).depth > 0 Then
        ' error trap for debugging
        If Abs(btemp - MyBot.bheat) > 3 Then
            'MyBot.pause
            dt = MyBot.bheat
        End If
        dt = mytime - enemies(e).s(0).t
        
        ' Four possible solutions - each of
        ' two current locations and two previous.
        ' Calculate resultant velocities
        
        dx = tx1 - enemies(e).s(0).x1
        dy = ty1 - enemies(e).s(0).y1
        v11 = Sqr(dx * dx + dy * dy) / dt
        
        dx = tx1 - enemies(e).s(0).x2
        dy = ty1 - enemies(e).s(0).y2
        v12 = Sqr(dx * dx + dy * dy) / dt
        
        dx = tx2 - enemies(e).s(0).x1
        dy = ty2 - enemies(e).s(0).y1
        v21 = Sqr(dx * dx + dy * dy) / dt
        
        dx = tx2 - enemies(e).s(0).x2
        dy = ty2 - enemies(e).s(0).y2
        v22 = Sqr(dx * dx + dy * dy) / dt
        
        ' Push down stack
        For i = enemies(e).depth To 1 Step -1
            enemies(e).s(i) = enemies(e).s(i - 1)
        Next i
        
        enemies(e).s(0).x1 = tx1
        enemies(e).s(0).y1 = ty1
        enemies(e).s(0).x2 = tx2
        enemies(e).s(0).y2 = ty2
    
        ' typical speed is 7-10m/sec, max is 20
        ' Find most likely solution.
        ' Start with solution 1-1
        mspeed = 12      ' was 12
        
        ' find most likely trajectory:
        ' speed must be less than 22 (20 plus measurement error)
        ' likely parallel to wall
        ' likely above 6
        
        s11 = 0
        s12 = 0
        s22 = 0
        s21 = 0
        
        ' Eliminate impossible velocities
        
        If v11 > 22 Then
            s11 = -5 * v11
        Else
            s11 = 20 - Abs(mspeed - v11)
        End If
        
        If v12 > 22 Then
            s12 = -5 * v12
        Else
            s12 = 20 - Abs(mspeed - v12)
        End If
        
        If v22 > 22 Then
            s22 = -5 * v22
        Else
            s22 = 20 - Abs(mspeed - v22)
        End If
        
        If v21 > 22 Then
            s21 = -5 * v21
        Else
            s21 = 20 - Abs(mspeed - v21)
        End If
        
        ' Look for wall following (dx or dy very small, but velocity is good)
        ' penalize for deviation from expected speed
        
        dx = Abs((tx1 - enemies(e).s(1).x1))
        dy = Abs((ty1 - enemies(e).s(1).y1))
        'If (dx / dt < 1) And (dy / dt > 6) And ((enemies(e).s(1).x1 > 900) Or (enemies(e).s(1).x1 < 100)) Then
        If (dx < 1) And ((enemies(e).s(1).x1 > 900) Or (enemies(e).s(1).x1 < 100)) Then
            s11 = s11 + 30 - Abs(dy / dt - mspeed)
            MyBot.post ("Wall x")
        End If
        If (dy < 1) And ((enemies(e).s(1).y1 > 900) Or (enemies(e).s(1).y1 < 100)) Then
            s11 = s11 + 30 - Abs(dx / dt - mspeed)
            MyBot.post ("Wall y")
        End If
        
        dx = Abs((tx1 - enemies(e).s(1).x2))
        dy = Abs((ty1 - enemies(e).s(1).y2))
        If (dx < 1) And ((enemies(e).s(1).x2 > 900) Or (enemies(e).s(1).x2 < 100)) Then
            s12 = s12 + 30
            MyBot.post ("Wall x")
        End If
        If (dy < 1) And ((enemies(e).s(1).y2 > 900) Or (enemies(e).s(1).y2 < 100)) Then
            s12 = s12 + 30 - Abs(dx / dt - mspeed)
            MyBot.post ("Wall y")
        End If
        
        dx = Abs((tx2 - enemies(e).s(1).x2))
        dy = Abs((ty2 - enemies(e).s(1).y2))
        If (dx < 1) And ((enemies(e).s(1).x2 > 900) Or (enemies(e).s(1).x2 < 100)) Then
            s22 = s22 + 30
            MyBot.post ("Wall x")
        End If
        If (dy < 1) And ((enemies(e).s(1).y2 > 900) Or (enemies(e).s(1).y2 < 100)) Then
            s22 = s22 + 30 - Abs(dx / dt - mspeed)
            MyBot.post ("Wall y")
        End If
        
        dx = Abs((tx2 - enemies(e).s(1).x1))
        dy = Abs((ty2 - enemies(e).s(1).y1))
        If (dx < 1) And ((enemies(e).s(1).x1 > 900) Or (enemies(e).s(1).x1 < 100)) Then
            s21 = s21 + 30
            MyBot.post ("Wall x")
        End If
        If (dy < 1) And ((enemies(e).s(1).y1 > 900) Or (enemies(e).s(1).y1 < 100)) Then
            s21 = s21 + 30 - Abs(dx / dt - mspeed)
            MyBot.post ("Wall y")
        End If
        
        best = s11
        
        enemies(e).s(0).x = enemies(e).s(0).x1
        enemies(e).s(0).y = enemies(e).s(0).y1
        enemies(e).s(1).x = enemies(e).s(1).x1
        enemies(e).s(1).y = enemies(e).s(1).y1
         
        If s12 > best Then
            best = s12
            enemies(e).s(0).x = enemies(e).s(0).x1
            enemies(e).s(0).y = enemies(e).s(0).y1
            enemies(e).s(1).x = enemies(e).s(1).x2
            enemies(e).s(1).y = enemies(e).s(1).y2
        End If
        
        If s21 > best Then
            best = s21
            enemies(e).s(0).x = enemies(e).s(0).x2
            enemies(e).s(0).y = enemies(e).s(0).y2
            enemies(e).s(1).x = enemies(e).s(1).x1
            enemies(e).s(1).y = enemies(e).s(1).y1
        End If
        
        If s22 > best Then
            best = s22
            enemies(e).s(0).x = enemies(e).s(0).x2
            enemies(e).s(0).y = enemies(e).s(0).y2
            enemies(e).s(1).x = enemies(e).s(1).x2
            enemies(e).s(1).y = enemies(e).s(1).y2
        End If
        
        
        enemies(e).lastx = enemies(e).s(0).x
        enemies(e).lasty = enemies(e).s(0).y
        
        ' error trap for debugging
        where = MyBot.WhereIs(e)
        y = where Mod 1000
        x = (where - y) / 1000
        
        Print #1, Format(v11, " ")
        
        If Abs(x - enemies(e).lastx) > 5 Or Abs(y - enemies(e).lasty) > 5 Then
            ' Debug: White marks on arena for prev. and current
            ' estimated positions
            MyBot.Mark Int(enemies(e).s(1).x1), Int(enemies(e).s(1).y1), RGB(0, 255, 0)
            MyBot.Mark Int(enemies(e).s(1).x2), Int(enemies(e).s(1).y2), RGB(0, 255, 0)
            MyBot.Mark Int(tx1), Int(ty1), RGB(255, 0, 0)
            MyBot.Mark Int(tx2), Int(ty2), RGB(255, 0, 0)
            MyBot.Mark Int(enemies(e).s(0).x), Int(enemies(e).s(0).y), RGB(0, 0, 255)
            MyBot.Mark Int(enemies(e).s(0).x), Int(enemies(e).s(0).y), RGB(0, 0, 255)
            MyBot.pause
        End If
        
        enemies(e).s(0).t = mytime
          
        ' increment depth
        If enemies(e).depth < 3 Then
            enemies(e).depth = enemies(e).depth + 1
        End If
        
        
        dt = enemies(e).s(0).t - enemies(e).s(1).t
        If dt = 0 Then Exit Sub
        
        ' If points are too close in time, combine
        If dt < 1 Then
            enemies(e).s(0).x = (enemies(e).s(0).x + enemies(e).s(1).x) / 2
            enemies(e).s(0).y = (enemies(e).s(0).y + enemies(e).s(1).y) / 2
            enemies(e).s(0).x1 = (enemies(e).s(0).x1 + enemies(e).s(1).x1) / 2
            enemies(e).s(0).y1 = (enemies(e).s(0).y1 + enemies(e).s(1).y1) / 2
            enemies(e).s(0).x2 = (enemies(e).s(0).x2 + enemies(e).s(1).x2) / 2
            enemies(e).s(0).y2 = (enemies(e).s(0).y2 + enemies(e).s(1).y2) / 2
            enemies(e).s(0).t = (enemies(e).s(0).t + enemies(e).s(1).t) / 2
            enemies(e).depth = enemies(e).depth - 1
        Else
            dx = -enemies(e).s(1).x
            dy = enemies(e).s(0).y - enemies(e).s(1).y
            enemies(e).vx = dx / dt
            enemies(e).vy = dy / dt
            
            ' Compute resulting x an y velocities
            ' we should readjust coordinates if velocity is
            ' unreasonable. Now we just limit velocity
            If enemies(e).vx > 18 Then enemies(e).vx = 18
            If enemies(e).vx < -18 Then enemies(e).vx = -18
            If enemies(e).vy > 18 Then enemies(e).vy = 18
            If enemies(e).vy < -18 Then enemies(e).vy = -18
        End If
    Else
        ' depth was 0
        enemies(e).s(0).t = mytime
        enemies(e).depth = 1
        enemies(e).s(0).x = tx1
        enemies(e).s(0).y = ty1
        enemies(e).s(0).x1 = tx1
        enemies(e).s(0).y1 = ty1
        enemies(e).s(0).x2 = tx2
        enemies(e).s(0).y2 = ty2
        enemies(e).vx = 0
        enemies(e).vy = 0
    End If  ' end if depth > 0
    
    enemies(e).lastrange = r
        
    
End Sub

' We're reloading. Take next scan in 360 degree sweep.
' We have about 8 seconds to do 360 degrees. We can do at most
' 40 scans in that time. We choose to do 12 degree cones
' at 10 degree increments
'
Sub surveytick()

    Dim range As Integer
    
    ' Do next coarse scan
        
    scandir = (scandir + 12) Mod 360
    
    range = Scan(scandir, 8)
    
End Sub

' Update location of each enemy.
' If verified on last pass, narrow scan resolution
' If not verified on last pass, widen scan resolution

Function verify(r As Integer) As Integer
Dim i As Integer
Dim icu As Integer
Dim ecount As Integer
Dim dstr As String
Dim narrowest As Integer

closest = 0
standoff = 2000
ecount = 0

dstr = ""
narrowest = 8

For i = 1 To 4
    If enemies(i).alive Then
        If enemies(i).verified Then
            ' Verified? Zoom in.
            enemies(i).scanres = enemies(i).scanres / 2
            If enemies(i).scanres < 1 Then
                enemies(i).scanres = 1
            End If
        End If
        ' Try to find. Might find wrong one...
        icu = tickle(i, Int(enemies(i).scanres))
        If icu > 0 Then
            enemies(icu).verified = True
            enemies(icu).scanres = enemies(i).scanres
            ' If we found the right one...
            If i = icu Then ecount = ecount + 1
        Else
            ' Couldn't find. Open up scanres.
            enemies(i).verified = False
            enemies(i).scanres = enemies(i).scanres * 2
            If enemies(i).scanres > 8 Then
                enemies(i).scanres = 8
            End If
            ' No tickle. Dead?
            If enemies(i).lastseen < (MyBot.Time - 15) Then
                enemies(i).alive = False
                MyBot.post ("Bot " & Str(i) & " is dead")
            End If
        End If
        If enemies(i).scanres < narrowest Then narrowest = enemies(i).scanres
        dstr = dstr & Str(enemies(i).scanres)
    Else
        dstr = dstr & "X"
    End If

Next i

If closest = 0 And narrowest = 8 Then
    ' Never found anyone
    hunting = True
    MyBot.post "Verify: Hunting - " & dstr
Else
    hunting = False
    MyBot.post "Verify: Attack - " & dstr
End If

verify = ecount

End Function

' Verify location of enemy e with scan of resolution r

Function tickle(e As Integer, r As Integer) As Integer

Dim b As Single
Dim b2 As Single
Dim range As Integer
Dim t As Integer

    If enemies(e).lastseen + 15 < MyBot.Time Then
        tickle = 0
        Exit Function
    End If
    
    If enemies(e).depth > 0 Then
        b = PredictBearing(e, MyBot.Time)
    Else
        b = enemies(e).lastdir
    End If
    range = Scan(b, r)
    If range = 0 Then
        'MyBot.post ("Tickle: " & Str(b) & " > " & Str(enemies(e).lastdir))
        b2 = enemies(e).lastdir
        range = Scan(b2, r)
        If range > 0 Then
            'MyBot.post ("...worked")
            b = b2
        End If
    End If
    
    ' Try +/- 15 degrees if no one there
    If range = 0 Then range = Scan((b + (r * 1.5) + 360) Mod 360, r)
    If range = 0 Then range = Scan((b + 360 - (r * 1.5) + 360) Mod 360, r)
    
    ' did we see anyone?
    If range > 0 Then
        tickle = MyBot.dsp
    Else
        tickle = 0
    End If

End Function

