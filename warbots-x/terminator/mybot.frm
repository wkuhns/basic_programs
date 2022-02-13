VERSION 5.00
Begin VB.Form MyForm 
   Caption         =   "Form1"
   ClientHeight    =   600
   ClientLeft      =   1728
   ClientTop       =   2244
   ClientWidth     =   1344
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   600
   ScaleWidth      =   1344
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   333
      Left            =   240
      Top             =   120
   End
End
Attribute VB_Name = "MyForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' MyForm is the startup object for this application.
' This form (MyForm) is not visible at run-time. It is only required
' because of object referencing limitations in Visual Basic
' Client/Server applications.

' There are a small number of required subroutines. Some should
' not be altered:
'    Form_Load - establishes the linkage to the robot server
'    Die - Used to clean up and exit

' There are three subroutines that are automatically run by
' the robot, which may contain user code:
'    UserInit - executed once at startup.
'    Ping - executed by the server when this robot is scanned
'       by another robot
'    UserFrame - run continuously as long as the robot is
'       alive.
' Any other user created subroutines must be called from one
' of these.

' There is one required global object: MyRobot.
' This is the robot object which provides the interface to
' the robot server. In VB5, you can use the Object Browser
' (F2 Key) to view the methods available.

Dim MyRobot As Robot

' User defined globals:
' These are 'global' to this form. Use these or add your own.
' They are not required except as used by your application

Dim speed As Integer
Dim dir As Integer
Dim scanres As Integer
Dim scandir As Single
Dim flight As Long
Dim reverse As Long
Dim ccw As Integer

' Information about a sighting of an enemy 'bot
Private Type sighting
    distance As Integer
    bearing As Single
    x As Single             ' Reported x,y
    y As Single
    x2 As Single            ' alternate x,y based on barrel heat
    y2 As Single
    t As Single
    vx As Single
    vy As Single
End Type

Dim enemies(4, 4) As sighting
Dim latest(4) As sighting

Private Type celltype
    dir(8) As Boolean
    destx(8) As Integer
    desty(8) As Integer
    safety As Integer
End Type

Dim cells(5, 5) As celltype

Dim cellx As Integer            ' x,y coords of current 200m square cell
Dim celly As Integer

Dim bestx As Integer            ' x,y coords of safest 200m square cell
Dim besty As Integer

Dim depth(4) As Integer
Dim closest As Integer          ' Enemy who is closest to us
Dim standoff As Single          ' Distance to nearest enemy
Dim btemp As Single             ' barrel temp
Dim lastbtime As Integer        ' time of last barrel temp calculation

Dim shells As Integer           ' remaining shells in this clip

Dim newdir As Integer           ' The direction we want to go
Dim bestdir As Integer          ' our preferred direction
Dim nextdir As Integer          ' our next preferred direction
Dim forward As Long
Sub CheckDrive()
    
Dim x As Single
Dim y As Single
Dim heat As Integer

    ' Try to go towards safest cell.
    
    MyRobot.ShowStatus
    
    If (cellx <> Int(MyRobot.x / 200) Or celly <> Int(MyRobot.y / 200)) And _
    forward < Timer Then
        cellx = Int(MyRobot.x / 200)
        celly = Int(MyRobot.y / 200)
        SetSafety
        forward = Timer + 10
    End If
    
    dir = bestdir

    If (heat > 150 And flight < Timer) Then speed = 25
    
    If (heat < 10 And speed < 35) Then speed = 35
    
    Call MyRobot.Drive(dir, speed)

End Sub

'
' Our life is over. Release robot's soul and die. Do not
' change this subroutine
'
Sub Die()

Dim finis As Long

Timer1.Enabled = False

MyRobot.post "Dying..."

finis = Timer + 2

While finis > Timer
    DoEvents
Wend

Set MyRobot = Nothing

End

End Sub

Function Fire(b As Integer, r As Integer) As Integer
' fire cannon at b,r. return -1 for survey, 0 for click, 1 for shot,
' 2 for overheated barrel

Dim stat As Integer
Dim mytime As Integer
Dim i As Integer

    mytime = MyRobot.Time
    If MyRobot.cannon(b Mod 360, r) = 0 Then
        Fire = 1
        MyRobot.post "Bang"
        ' Update barrel temp calcs
        If btemp > 0 Then
            btemp = btemp - (mytime - lastbtime) * 0.2
            If btemp < 0 Then btemp = 0
        End If
        btemp = btemp + 20
        lastbtime = mytime
        ' If we're too hot, let caller know
        If btemp > 35 Then
            Fire = 2
            MyRobot.post "Hot Barrel...."
        End If
        ' Check for reloading. Overrides hot barrel.
        shells = shells - 1
        If shells = 0 Then
            shells = 4
            MyRobot.post "Reloading"
            Fire = -1
        End If
    Else
        MyRobot.post "Click: " + Str(b) + "," + Str(r)
        Fire = 0
    End If

End Function

Function GetBearing(x As Integer, y As Integer) As Single
        
' Calculate bearing from ourselves to x,y
Dim dx As Single
Dim dy As Single
Dim b As Single
    
    dx = x - MyRobot.x
    If dx = 0 Then dx = 1
    dy = y - MyRobot.y
    b = Atn(dy / dx) * 57.3 + 360
    If b > 360 Then b = b - 360
    If dx < 0 Then b = b + 180
    If b > 360 Then b = b - 360
    GetBearing = b
    
End Function

Sub InitCells()

Dim i As Integer
Dim j As Integer
Dim k As Integer

For i = 0 To 4
    For j = 0 To 4
        cells(i, j).safety = 2000
        For k = 0 To 7
            cells(i, j).dir(k) = True
            If k = 7 Or k = 0 Or k = 1 Then cells(i, j).destx(k) = i + 1
            If k > 2 And k < 6 Then cells(i, j).destx(k) = i - 1
            If k = 2 Or k = 6 Then cells(i, j).destx(k) = i
            If k > 0 And k < 4 Then cells(i, j).desty(k) = j + 1
            If k > 4 Then cells(i, j).desty(k) = j - 1
            If k = 0 Or k = 4 Then cells(i, j).desty(k) = j
        Next k
        If i = 0 Then
            cells(i, j).dir(3) = False
            cells(i, j).dir(4) = False
            cells(i, j).dir(5) = False
        End If
        If i = 4 Then
            cells(i, j).dir(0) = False
            cells(i, j).dir(1) = False
            cells(i, j).dir(7) = False
        End If
        If j = 0 Then
            cells(i, j).dir(5) = False
            cells(i, j).dir(6) = False
            cells(i, j).dir(7) = False
        End If
        If j = 4 Then
            cells(i, j).dir(1) = False
            cells(i, j).dir(2) = False
            cells(i, j).dir(3) = False
        End If
    Next j
Next i
End Sub
Function scan(b As Single, res As Integer) As Integer
    
Dim myrange As Integer
Dim enemy As Integer
Dim where As Long
Dim x As Integer
Dim y As Integer

    ' Look for enemy
    
    myrange = MyRobot.scan(b, res)
    
    ' Keep track of who is near
    If myrange > 0 Then
        enemy = MyRobot.dsp
'        where = MyRobot.WhereIs(enemy)
'        y = where Mod 1000
'        x = (where - y) / 1000
        latest(enemy).distance = myrange
        latest(enemy).bearing = b
        latest(enemy).t = MyRobot.Time
        latest(enemy).x = MyRobot.x + myrange * Cos(b / 57.3)
        latest(enemy).y = MyRobot.y + myrange * Sin(b / 57.3)
        
        ' Is this guy closest to us?
        If enemy = closest Or myrange < standoff Then
            standoff = myrange
            closest = enemy
        End If
        If res < 3 Then
'            If Abs(x - (MyRobot.x + myrange * Cos(b / 57.3))) > 5 Then
'                MyRobot.pause
'                Stop
'            End If
'            If Abs(y - (MyRobot.y + myrange * Sin(b / 57.3))) > 5 Then
'                MyRobot.pause
'                Stop
'            End If
            PushEnemy enemy, myrange, b, 0
        End If
    End If

    scan = myrange
    
End Function

Function PredictBearing(e As Integer, t As Integer) As Single
' predict where enemy 'e' will be at time 't'

Dim tx As Integer
Dim ty As Integer
Dim dt As Single
    
    If depth(e) > 0 Then
        dt = t - enemies(e, 0).t
        
        tx = enemies(e, 0).x + enemies(e, 0).vx * dt
        ty = enemies(e, 0).y + enemies(e, 0).vy * dt
        
        PredictBearing = GetBearing(tx, ty)
   Else
        PredictBearing = GetBearing(Int(latest(e).x), Int(latest(e).y))
    End If
    
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

    mytime = MyRobot.Time
    
    tx = latest(e).x
    ty = latest(e).y
    
    If depth(e) > 0 Then
    
        dt = mytime - enemies(e, 0).t
        
        ' check barrel temperature
        
        btemp = btemp - (mytime - lastbtime) * 0.2
        lastbtime = mytime
        If btemp < 0 Then btemp = 0
        
        ' calculate target x,y. If barrel is hot, there's an error.
        ' try to pick the right possibility.
        If btemp > 0 Then
            tx2 = MyRobot.x + (r + btemp) * Cos(b / 57.3)
            ty2 = MyRobot.y + (r + btemp) * Sin(b / 57.3)
            ' make a mark on the arena
            MyRobot.Mark Int(tx2), Int(ty2), RGB(0, 255, 255)
            tx = MyRobot.x + (r - btemp) * Cos(b / 57.3)
            ty = MyRobot.y + (r - btemp) * Sin(b / 57.3)
            ' make a mark on the arena
            MyRobot.Mark Int(tx), Int(ty), RGB(255, 255, 0)
            dx = tx - enemies(e, 0).x
            dy = ty - enemies(e, 0).y
            dx2 = tx2 - enemies(e, 0).x
            dy2 = ty2 - enemies(e, 0).y
            'if velocity based on tx2/ty2 is lower and velocity based
            'on tx/ty is too high, then use tx2/ty2 instead.
            If ((Abs(dx) > Abs(dx2)) And (Abs(dx) > 2 * dt)) Or _
            (Abs(dy) > Abs(dy2)) And (Abs(dy) > 2 * dt) Then
                tx = tx2
                ty = ty2
            End If
        End If
    
        ' If scans are closer in time than 400 ms, average them
        If dt <= 4 Then
            dx = tx - enemies(e, 0).x
            dy = ty - enemies(e, 0).y
            ' Impossibly fast? average them out.
            If dx > 2 * dt Then
                xs = dx - 2 * dt
                tx = tx - xs / 2
                enemies(e, 0).x = enemies(e, 0).x + xs / 2
            End If
            If dx < -2 * dt Then
                xs = dx + 2 * dt
                tx = tx - xs / 2
                enemies(e, 0).x = enemies(e, 0).x + xs / 2
            End If
            If dy > 2 * dt Then
                xs = dy - 2 * dt
                ty = ty - xs / 2
                enemies(e, 0).y = enemies(e, 0).y + xs / 2
            End If
            If dy < -2 * dt Then
                xs = dy + 2 * dt
                ty = ty - xs / 2
                enemies(e, 0).y = enemies(e, 0).y + xs / 2
            End If
            enemies(e, 0).t = enemies(e, 0).t + dt / 2
            MyRobot.Mark Int(tx), Int(ty), RGB(255, 255, 255)
        
        Else    ' Scans weren't too close together
            ' Push down stack
            For i = depth(e) To 1 Step -1
                enemies(e, i) = enemies(e, i - 1)
            Next i
           
            dx = tx - enemies(e, 1).x
            dy = ty - enemies(e, 1).y
            dt = mytime - enemies(e, 1).t
            ' Impossibly fast? average them out.
            If dx > 2 * dt Then
                xs = dx - 2 * dt
                tx = tx - xs / 2
                enemies(e, 1).x = enemies(e, 1).x + xs / 2
            End If
            If dx < -2 * dt Then
                xs = dx + 2 * dt
                tx = tx - xs / 2
                enemies(e, 1).x = enemies(e, 1).x + xs / 2
            End If
            If dy > 2 * dt Then
                xs = dy - 2 * dt
                ty = ty - xs / 2
                enemies(e, 1).y = enemies(e, 1).y + xs / 2
            End If
            If dy < -2 * dt Then
                xs = dy + 2 * dt
                ty = ty - xs / 2
                enemies(e, 1).y = enemies(e, 1).y + xs / 2
            End If
            MyRobot.Mark Int(tx), Int(ty), RGB(255, 255, 255)
            enemies(e, 0).t = mytime
            ' increment depth
            If depth(e) < 3 Then
                depth(e) = depth(e) + 1
            End If
        End If  ' end if scans were / weren't too close
        enemies(e, 0).x = tx
        enemies(e, 0).y = ty
        dx = enemies(e, 0).x - enemies(e, 1).x
        dy = enemies(e, 0).y - enemies(e, 1).y
        dt = enemies(e, 0).t - enemies(e, 1).t
        If dt = 0 Then Stop
        enemies(e, 0).vx = dx / dt
        enemies(e, 0).vy = dy / dt
        ' we *should* check and average as above...
        If enemies(e, 0).vx > 1.2 Then enemies(e, 0).vx = 1.2
        If enemies(e, 0).vx < -1.2 Then enemies(e, 0).vx = -1.2
        If enemies(e, 0).vy > 1.2 Then enemies(e, 0).vy = 1.2
        If enemies(e, 0).vy < -1.2 Then enemies(e, 0).vy = -1.2
            
        ' ditch any older than 6 seconds
        For i = 1 To depth(e)
            If enemies(e, i).t < (enemies(e, 0).t - 60) Then
                depth(e) = i - 1
                Exit For
            End If
        Next i
    Else
        ' depth was 0
        enemies(e, 0).t = mytime
        depth(e) = depth(e) + 1
        enemies(e, 0).x = tx
        enemies(e, 0).y = ty
        enemies(e, 0).vx = 0
        enemies(e, 0).vy = 0
    End If  ' end if depth > 0
    enemies(e, 0).distance = r
    MyRobot.Mark Int(tx), Int(ty), RGB(255, 255, 255)
            
End Sub
Sub SetSafety()
' Assess safety of adjacent cells. Set bestdir and nextdir

Dim i As Integer
Dim safest As Integer
Dim j As Integer
Dim e As Integer
Dim cx As Integer
Dim cy As Integer
Dim dx As Double
Dim dy As Double
Dim hsq As Double
Dim ax As Integer
Dim ay As Integer

cx = Int(MyRobot.x / 200)
cy = Int(MyRobot.y / 200)
'MyRobot.pause
safest = 0
For i = 0 To 4
    For j = 0 To 4
        cells(i, j).safety = 2000
        For e = 1 To 4
            If latest(e).distance > 0 Then
                dx = latest(e).x - (i * 200 + 100)
                dy = latest(e).y - (j * 200 + 100)
                hsq = Sqr(dx * dx + dy * dy)
                If hsq < cells(i, j).safety Then cells(i, j).safety = hsq
            End If
        Next e
        If cells(i, j).safety > safest Then
            safest = cells(i, j).safety
            bestx = i
            besty = j
        End If
    Next j
Next i

bestdir = GetBearing(bestx * 200 + 100, besty * 200 + 100)
' round to nearest 45 degrees
bestdir = bestdir / 45
bestdir = bestdir * 45
MyRobot.post "Aiming for " + Str(bestx) + "," + Str(besty) + " at " + Str(bestdir)

' Now pick best ( lowest danger ) direction. Has to be +/- 45 degrees
' from bestdir

safest = 0

i = bestdir / 45                            ' bestdir
' if you can move in this direction
If cells(cx, cy).dir(i) Then
    ax = cells(cx, cy).destx(i)
    ay = cells(cx, cy).desty(i)
    If safest < cells(ax, ay).safety Then
        safest = cells(ax, ay).safety
        nextdir = i * 45
    End If
End If

i = ((bestdir / 45) + 7) Mod 8              ' bestdir - 1
' if you can move in this direction
If cells(cx, cy).dir(i) Then
    ax = cells(cx, cy).destx(i)
    ay = cells(cx, cy).desty(i)
    If safest < cells(ax, ay).safety Then
        safest = cells(ax, ay).safety
        nextdir = i * 45
    End If
End If

i = ((bestdir / 45) + 1) Mod 8              ' bestdir +1
' if you can move in this direction
If cells(cx, cy).dir(i) Then
    ax = cells(cx, cy).destx(i)
    ay = cells(cx, cy).desty(i)
    If safest < cells(ax, ay).safety Then
        safest = cells(ax, ay).safety
        nextdir = i * 45
    End If
End If

' if we didn't find any, we *could* be heading into a corner diagonally.
If safest = 0 Then
    bestdir = (bestdir + 135) Mod 360
Else
    bestdir = nextdir
End If

' now look +/- 90 degrees from best to get next best

safest = 0

i = ((bestdir / 45) + 6) Mod 8              ' bestdir - 2
' if you can move in this direction
If cells(cx, cy).dir(i) Then
    ax = cells(cx, cy).destx(i)
    ay = cells(cx, cy).desty(i)
    If safest < cells(ax, ay).safety Then
        safest = cells(ax, ay).safety
        nextdir = i * 45
    End If
End If

i = ((bestdir / 45) + 7) Mod 8              ' bestdir - 1
' if you can move in this direction
If cells(cx, cy).dir(i) Then
    ax = cells(cx, cy).destx(i)
    ay = cells(cx, cy).desty(i)
    If safest < cells(ax, ay).safety Then
        safest = cells(ax, ay).safety
        nextdir = i * 45
    End If
End If

i = ((bestdir / 45) + 1) Mod 8              ' bestdir +1
' if you can move in this direction
If cells(cx, cy).dir(i) Then
    ax = cells(cx, cy).destx(i)
    ay = cells(cx, cy).desty(i)
    If safest < cells(ax, ay).safety Then
        safest = cells(ax, ay).safety
        nextdir = i * 45
    End If
End If

i = ((bestdir / 45) + 2) Mod 8              ' bestdir - 2
' if you can move in this direction
If cells(cx, cy).dir(i) Then
    ax = cells(cx, cy).destx(i)
    ay = cells(cx, cy).desty(i)
    If safest < cells(ax, ay).safety Then
        safest = cells(ax, ay).safety
        nextdir = i * 45
    End If
End If

MyRobot.post "bestdir = " + Str(bestdir) + ". Next = " + Str(nextdir)

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
    mytime = MyRobot.Time
    If depth(e) > 1 Then
        tof = r / 20
        tx = enemies(e, 0).x + enemies(e, 0).vx * (mytime + tof - enemies(e, 0).t)
        ty = enemies(e, 0).y + enemies(e, 0).vy * (mytime + tof - enemies(e, 0).t)
        b = GetBearing(Int(tx), Int(ty))
        dx = tx - MyRobot.x
        dy = ty - MyRobot.y
        hsq = dx * dx + dy * dy
        r = Sqr(hsq)
        MyRobot.Mark Int(tx), Int(ty), RGB(0, 0, 0)
    End If
    Shoot = Fire(Int(b), r)

End Function

Sub survey(initial As Boolean)

    Dim myscandir As Single
    Dim mytime As Integer
    Dim range As Integer
    Dim enemy As Integer
    Dim known As Integer
    Dim scans As Integer
    
    myscandir = 0
    known = 0
    standoff = 2000
    
    While known < 3 And scans < 21
        range = scan(myscandir, 10)
        If range > 0 Then
            If MyRobot.dsp <> enemy Then
                known = known + 1
                enemy = MyRobot.dsp
            End If
            If range < standoff Then
                closest = enemy
                standoff = range
            End If
            If initial And range > 40 And range < 400 Then attack myscandir, False
        End If
        scans = scans + 1
        myscandir = (myscandir + 18) Mod 360
        If Not initial Then CheckDrive
    Wend
    MyRobot.post "Found " + Str(known) + ". Closest is " + Str(closest) + " at " + Str(standoff) + " meters"
    
    scandir = PredictBearing(closest, MyRobot.Time)
    
End Sub

' Verify location of enemy e
Function tickle(e As Integer) As Integer

Dim b As Single
Dim range As Integer
Dim t As Integer

        b = PredictBearing(e, MyRobot.Time)
 
        range = scan(b, 10)
        
        ' did we see anyone?
        If range > 0 Then
            t = MyRobot.dsp
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
Dim i As Integer
Dim mytime As Integer

    CheckDrive
    
    range = scan(scandir, 10)
    
    If range > 40 And range < 700 Then
        ' If he's really close, just shoot
        If range < 150 Then
            Call Fire(Int(scandir), range)
        End If
        ' In any event, go into attack mode
        If attack(scandir, True) = -1 Then                ' Reloading
            If btemp > 35 Then
                For i = 1 To (btemp - 35)
                    DoEvents
                    Sleep (400)
                    CheckDrive
                Next i
                mytime = MyRobot.Time
                btemp = btemp - (mytime - lastbtime) * 0.2
                lastbtime = mytime
            End If
            survey False                                 ' Find everyone
            scandir = PredictBearing(closest, MyRobot.Time)
        Else
            scandir = (scandir + 360 - (ccw * 15)) Mod 360
        End If
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

Function attack(targetdir As Single, vicious As Boolean) As Integer
    
Dim myscan As Single
Dim myres As Integer
Dim startscan As Single
Dim endscan As Single
Dim range As Single
Dim delta As Single
Dim t As Integer
Dim enemy As Integer
Dim edir As Single
Dim result As Integer

    enemy = MyRobot.dsp
    'back off 12 degrees
    edir = targetdir
    myscan = edir
    myres = 5
    delta = -12
    MyRobot.post "Attack at " + Str(myscan)
    While delta < 15
        ' scan at 3 degrees
        myscan = (edir + delta + 360) Mod 360
        If myscan > 360 Then myscan = myscan - 360
        range = scan(myscan, 3)
        enemy = MyRobot.dsp
        edir = GetBearing(Int(latest(enemy).x), Int(latest(enemy).y))
        ' if we have him, zero in
        If (range > 40 And range < 700) Then
            delta = -2
            edir = myscan
            While delta < 8
                myscan = (edir + delta + 360) Mod 360
                If myscan > 360 Then myscan = myscan - 360
                range = scan(myscan, 1)
                If (range > 40 And range < 700) Then
                    result = Shoot(enemy, myscan, Int(range))
                    If result = -1 Then
                        attack = -1
                        Exit Function
                    End If
                    ' If we're not vicious, exit after first shot
                    If Not vicious Then
                        attack = 1
                        Exit Function
                    End If
                    delta = -1
                    edir = GetBearing(Int(latest(enemy).x), Int(latest(enemy).y))
                Else
                    enemy = MyRobot.dsp
                    edir = GetBearing(Int(latest(enemy).x), Int(latest(enemy).y))
                    delta = delta + 1.5
                End If
'                edir = PredictBearing(enemy, MyRobot.Time)
                CheckDrive
            Wend
        Else
            delta = delta + 5
        End If
'        edir = PredictBearing(enemy, MyRobot.Time)
        CheckDrive
    Wend
    speed = 35

End Function
'
' This subroutine MUST exist with EXACTLY this name and
' this argument list. The contents are up to the user.
' This subroutine is invoked by the server when this robot
' is scanned by another robot.
'
Public Sub Ping(m As Integer)

    ' check if we know where he is. If too far, don't run.
        
    If depth(m) > 0 Then
        If latest(m).distance < 750 Then
            speed = 100
            flight = Timer + latest(m).distance / 200
            MyRobot.Drive dir, speed
        Else
            ' He can't hurt us.
        End If
    Else
        ' We don't know where he is, so run
        speed = 100
        flight = Timer + 3
        MyRobot.Drive dir, speed
    End If
    
End Sub



Sub UserInit()

Dim closewall As Integer
Dim t As Integer
Dim r As Integer
Dim x As Integer
Dim y As Integer

While x = 0 Or y = 0
    x = MyRobot.x
    y = MyRobot.y
    MyRobot.ShowStatus
Wend

cellx = Int(x / 200)
celly = Int(y / 200)

If x < 500 Then
    dir = 0
Else
    dir = 180
End If
MyRobot.SetName ("Terminator")

speed = MyRobot.speed
While speed = 0
    MyRobot.Drive dir, 35
    MyRobot.ShowStatus
    speed = MyRobot.speed
Wend

InitCells

speed = 35
shells = 4
closest = 0
standoff = 2000
survey True
SetSafety
dir = bestdir

End Sub

'
' Don't change this at all. This code creates the linkage
' to the robot server process.
'
Private Sub Form_Load()

' Create robot object

Set MyRobot = CreateObject("RobotDLL.Robot")

' Register 'Ping' procedure with server.

Call MyRobot.RegisterAlert(MyForm)

' Do user's initialization.

MyRobot.ShowStatus

UserInit

' Don't change this - User specific stuff is in DoFrame.

While True

    ' Do the user's cyclic stuff.
    UserFrame
    
Wend

End Sub

Private Sub Form_Unload(Cancel As Integer)

MsgBox ("AppForm unloading...")

Die

End Sub





Private Sub Timer1_Timer()

Dim delta As Integer

delta = newdir - dir

If delta <> 0 Then
    If delta > 180 Then delta = delta - 360
    If delta < -180 Then delta = delta + 360

    If delta > 9 Then delta = 9

    If delta < -9 Then delta = -9

    dir = (dir + delta + 360) Mod 360
Else
    Timer1.Enabled = False
End If

End Sub


