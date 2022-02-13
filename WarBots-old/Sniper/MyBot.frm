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

Dim MyRobot As RobotDLL.Robot

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

Private Type sighting
    x As Integer
    y As Integer
    t As Integer
End Type

Private Type history
    depth As Integer
    s(3) As sighting
End Type

Dim enemies(4) As history

Dim shells As Integer

Sub CheckDrive()
    
Dim x, y As Single

    ' change direction every 120 seconds. Hug walls.
    
    If reverse < Timer And flight < Timer Then
        reverse = Timer + 120
        ccw = ccw * -1
        dir = (dir + 180) Mod 360
        MyRobot.post "changed direction to " + Str(ccw)
    End If
    
    x = MyRobot.x
    y = MyRobot.y

    ' ccw alternates between 1 and -1. It's used as a multiplier
    ' later on.
    
    If ccw = -1 Then
        If (x > 900 And dir = 0) Then dir = 90: scandir = 270
        If (x < 100 And dir = 180) Then dir = 270: scandir = 90
        If (y > 900 And dir = 90) Then dir = 180: scandir = 0
        If (y < 100 And dir = 270) Then dir = 0: scandir = 180
    Else
        If (x > 900 And dir = 0) Then dir = 270: scandir = 90
        If (x < 100 And dir = 180) Then dir = 90: scandir = 270
        If (y > 900 And dir = 90) Then dir = 0: scandir = 180
        If (y < 100 And dir = 270) Then dir = 180: scandir = 0
    End If
    
    ' flight tells us when to stop running from a ping.
    
    If flight < Timer And speed = 100 Then
        speed = 35
    End If
    
    ' check motor heat and adjust speed
    
    If (MyRobot.heat > 150 And flight < Timer) Then speed = 25
    
    If (MyRobot.heat < 10 And speed < 35) Then speed = 35
    
    Call MyRobot.Drive(dir, speed)

End Sub

'
' Our life is over. Release robot's soul and die. Do not
' change this subroutine
'
Sub Die()

Set MyRobot = Nothing

End

End Sub

Sub Fire(b As Integer, r As Integer)

Dim stat As Integer

    If MyRobot.cannon(b, r) = 0 Then
        MyRobot.post "Bang"
        shells = shells - 1
        If shells = 0 Then
            shells = 4
            MyRobot.post "Reloading"
        End If
    Else
        MyRobot.post "Click"
    End If

End Sub

' add an enemy sighting of enemy e at time t at range r
' and bearing b
Sub PushEnemy(e As Integer, r As Integer, b As Single, t As Integer)

Dim i As Integer
Dim dx As Long
Dim dy As Long
Dim dt As Integer
Dim vx As Single
Dim vy As Single
Dim tx As Single
Dim ty As Single
Dim tof As Integer
Dim hsq As Long
Dim x As Integer
Dim y As Integer

    If enemies(e).depth > 0 Then
        For i = enemies(e).depth To 1 Step -1
            enemies(e).s(i) = enemies(e).s(i - 1)
        Next i
    End If
    
    If enemies(e).depth < 3 Then
        enemies(e).depth = enemies(e).depth + 1
    End If
    
    ' reset depth to ignore data over 80 ticks old
    For i = 1 To enemies(e).depth - 1
        If (t - enemies(e).s(i).t) > 80 Then
            MyRobot.post "Now: " + Str(t) + " Then: " + _
            Str(enemies(e).s(i).t)
            enemies(e).depth = i
            Exit For
        End If
    Next i
    
    enemies(e).s(0).t = t
    x = MyRobot.x
    y = MyRobot.y
    enemies(e).s(0).x = x + r * Cos(b / 57.3)
    enemies(e).s(0).y = y + r * Sin(b / 57.3)
    MyRobot.post "Enemy " + Str(e) + ": " + Str(enemies(e).s(0).x) + ", " _
    + Str(enemies(e).s(0).y) + " at " + Str(t)

    ' time is in 100ms ticks.
    MyRobot.post "Raw: " + Str(b) + ", " + Str(r)
    If enemies(e).depth > 1 Then
        dx = enemies(e).s(0).x - enemies(e).s(1).x
        dy = enemies(e).s(0).y - enemies(e).s(1).y
        dt = enemies(e).s(0).t - enemies(e).s(1).t
        vx = dx / dt
        vy = dy / dt
        tof = r / 20
        tx = enemies(e).s(0).x + vx * tof
        ty = enemies(e).s(0).y + vy * tof
        If tx < 0 Or tx > 999 Or ty < 0 Or ty > 999 Then
            Stop
        End If
        dx = tx - MyRobot.x
        If dx = 0 Then dx = 1
        dy = ty - MyRobot.y
        b = Atn(dy / dx) * 57.3 + 360 Mod 360
        hsq = dx * dx + dy * dy
        r = Sqr(hsq)
        MyRobot.post "Corrected: " + Str(b) + ", " + Str(r)
    End If
    Call Fire(Int(b), r)

End Sub

Sub SmartFire(e As Integer, r As Integer, b As Single, t As Integer)

Dim i As Integer
Dim dx As Integer
Dim dy As Integer
Dim dt As Integer
Dim vx As Single
Dim vy As Single
Dim tx As Single
Dim ty As Single
Dim tof As Integer

    ' time is in 100ms ticks.
    If enemies(e).depth > 1 Then
        dx = enemies(e).s(0).x - enemies(e).s(1).x
        dy = enemies(e).s(0).y - enemies(e).s(1).y
        dt = enemies(e).s(0).t - enemies(e).s(1).t
        vx = dx / dt
        vy = dy / dt
        tof = r / 20
        tx = enemies(e).s(0).x + vx * tof
        ty = enemies(e).s(0).y + vy * tof
        dx = tx - MyRobot.x
        dy = ty - MyRobot.y
        b = Atn(dy / dx) * 57.3
        r = Int(Sqr(dx * dx + dy * dy))
    End If
    Call Fire(Int(b), r)

End Sub

'
' This is where you put anything that you want your robot to
' do continuously. You may create your own subroutines and
' call them from here. Depending on the complexity of your
' code, this routine is run about four times per second as
' long as your robot is alive.
'
Public Sub UserFrame()

' Perform a cycle of calculations for our robot.

Dim x As Integer
Dim y As Integer
Dim range As Integer
Dim startscan As Single
Dim endscan As Single
Dim oldscan As Single
Dim stat As Integer

    CheckDrive
    
    ' Look for enemy
    
    range = MyRobot.scan(scandir, scanres)
    
    If range > 40 And range < 700 Then
        If range < 150 Then
            Call Fire(Int(scandir), range)
            scandir = (scandir + 360 - (ccw * scanres)) Mod 360
        End If
        attack (scandir)
    End If
    
    ' Check if we've scanned past our left shoulder
    If scandir = dir Then
        ' reset scan to behind us
        scandir = (dir + 180) Mod 360
    End If

    scandir = (scandir + 360 + (ccw * scanres * 2)) Mod 360
    
End Sub

Sub attack(targetdir As Single)
    
Dim myscan As Single
Dim startscan As Single
Dim endscan As Single
Dim range As Single
Dim delta As Single
Dim t As Integer

    myscan = targetdir
    speed = 50
    delta = 1.5
    While delta < 12
        range = MyRobot.scan(myscan, 1)
        If (range > 40 And range < 700) Then
            t = MyRobot.Time
'            Call Fire(Int(myscan), Int(range))
            Call PushEnemy(MyRobot.dsp, Int(range), myscan, t)
'            Call SmartFire(MyRobot.dsp, Int(range), scandir, t)
            delta = 0
        Else
            If delta > 0 Then
                delta = delta * -1
            Else
                delta = (delta * -1) + 2
            End If
        End If
        myscan = (targetdir + delta) Mod 360
'       MyRobot.ShowStatus
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
Public Sub Ping(m As String)

    speed = 100
    flight = Timer + 4
    MyRobot.Drive dir, speed
    
End Sub





Sub UserInit()

dir = 0
speed = 35
MyRobot.SetName ("Sniper")
scanres = 10
ccw = 1
shells = 4

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

UserInit

' Don't change this - User specific stuff is in DoFrame.

While True

    ' Check to see if we're dead. You can't cheat death
    ' by changing this - all that will happen is that
    ' you'll have dead processes cluttering up your
    ' system.
    If MyRobot.status = "K" Then
        Die
        Exit Sub
    End If
    
    ' ShowStatus MUST be called periodically.
    MyRobot.ShowStatus
    
    ' Do the user's cyclic stuff.
    UserFrame
    
Wend

End Sub

Private Sub Form_Unload(Cancel As Integer)

MsgBox ("AppForm unloading...")

Die

End Sub





