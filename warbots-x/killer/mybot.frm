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
Dim depth(4) As Integer
Dim closest As Integer          ' Enemy who is closest to us
Dim standoff As Single          ' Distance to nearest enemy
Dim btemp As Single             ' barrel temp
Dim lastbtime As Integer        ' time of last barrel temp calculation

Dim shells As Integer           ' remaining shells in this clip

Dim newdir As Integer           ' The direction we want to go

Sub CheckDrive()
    
Dim x As Single
Dim y As Single
Dim heat As Integer

    ' change direction every 120 seconds. Hug walls.
    ' update x, y
    MyRobot.ShowStatus
    
    If reverse < Timer And flight < Timer Then
        reverse = Timer + 120
        ccw = ccw * -1
        newdir = (dir + 180) Mod 360
        Timer1.Enabled = True
        MyRobot.post "changed direction to " + Str(ccw)
    End If
    
    x = MyRobot.x
    y = MyRobot.y

    ' ccw alternates between 1 and -1. It's used as a multiplier
    ' later on.
        
    ' Check if we've reached first wall after start
    If ccw = 0 Then
        If x > 875 Or x < 125 Then ccw = -1
        If y > 875 Or y < 125 Then ccw = -1
        If ccw = -1 Then reverse = Timer + 120
    End If
       
    If ccw = -1 Then
        If (x > 875 And dir = 0) Then newdir = 90: scandir = 270: Timer1.Enabled = True
        If (x < 125 And dir = 180) Then newdir = 270: scandir = 90: Timer1.Enabled = True
        If (y > 875 And dir = 90) Then newdir = 180: scandir = 0: Timer1.Enabled = True
        If (y < 125 And dir = 270) Then newdir = 0: scandir = 180: Timer1.Enabled = True
    End If
    
    If ccw = 1 Then
        If (x > 875 And dir = 0) Then newdir = 270: scandir = 90: Timer1.Enabled = True
        If (x < 125 And dir = 180) Then newdir = 90: scandir = 270: Timer1.Enabled = True
        If (y > 875 And dir = 90) Then newdir = 0: scandir = 180: Timer1.Enabled = True
        If (y < 125 And dir = 270) Then newdir = 180: scandir = 0: Timer1.Enabled = True
    End If
    
    ' flight tells us when to stop running from a ping.
    
    If flight < Timer And speed = 100 Then
        speed = 35
    End If
    
    ' check motor heat and adjust speed
    
    heat = MyRobot.heat
    
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
    
        If dt <= 2 Then
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
        
' reset depth to ignore data over 80 ticks old
'    For i = 1 To enemies(e).depth - 1
'        If (t - enemies(e, i).t) > 80 Then
'            Str (enemies(e, i).t)
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

Sub survey()

    Dim myscandir As Single
    Dim mytime As Integer
    Dim range As Integer
    Dim enemy As Integer
    
    mytime = MyRobot.Time + 110
    If closest <> 0 Then
        myscandir = (GetBearing(Int(enemies(closest, 0).x), _
            Int(enemies(closest, 0).y)) + 330) Mod 360
    Else
        myscandir = (dir + 180) Mod 360
    End If
    
    standoff = 2000
    
    While MyRobot.Time < mytime
        range = scan(myscandir, 10)
        myscandir = (myscandir + 360 + 15) Mod 360
        CheckDrive
    Wend
    MyRobot.post "Closest is " + Str(closest) + " at " + Str(standoff) + " meters"
    
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
        If attack(scandir) = -1 Then                ' Reloading
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
            survey                                  ' Find everyone
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

Function attack(targetdir As Single) As Integer
    
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
    edir = targetdir
    myscan = edir
    myres = 5
    delta = 0.1
    MyRobot.post "Attack at " + Str(myscan)
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
                    result = Shoot(enemy, myscan, Int(range))
                    If result = -1 Then
                        attack = -1
                        Exit Function
                    End If
                    delta = 0.1
                    edir = myscan
                Else
                    delta = (Abs(delta) + 1.5) * -1 * Sgn(delta)
                End If
'                edir = PredictBearing(enemy, MyRobot.Time)
                CheckDrive
            Wend
        Else
            delta = (Abs(delta) + 3) * -1 * Sgn(delta)
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
            flight = Timer + 4
            MyRobot.Drive dir, speed
        Else
            ' He can't hurt us.
        End If
    Else
        ' We don't know where he is, so run
        speed = 100
        flight = Timer + 4
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

MyRobot.post Str(x) + "," + Str(y) + " " + Str(newdir) + " " + Str(closewall)
speed = 35
MyRobot.SetName ("Killer")
scanres = 10
ccw = 0
shells = 4
reverse = Timer + 120
closest = 0
standoff = 2000

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


