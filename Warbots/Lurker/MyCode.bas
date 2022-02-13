Attribute VB_Name = "MyCode"
Option Explicit
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
Dim dir As Integer
Dim scanres As Integer
Dim scandir As Single
Dim flight As Long
Dim reverse As Long
Dim ccw As Integer
Dim shells As Integer

Private Type sighting
    x As Integer
    y As Integer
    t As Integer
End Type

Private Type enemy
    s(3) As sighting
End Type

Dim enemies(4) As enemy

Sub CheckDrive()
    
Dim x, y As Single

    x = Mybot.x
    y = Mybot.y

    ' We're not running. Take time to turn around if we need to.
    If flight < Timer Then
        If (x > 200) Then dir = 180
        If (x < 100) Then dir = 0
        speed = 35
    End If
   
    Call Mybot.Drive(dir, speed)

End Sub


Sub GoHome()

Dim x As Single
Dim y As Single

    x = Mybot.x
    y = Mybot.y
    
    While y > 100
        Call Mybot.Drive(270, 100)
        y = Mybot.y
    Wend
    Call Mybot.Drive(180, 35)
    While x > 100
        Call Mybot.Drive(180, 100)
        x = Mybot.x
    Wend
    Call Mybot.Drive(0, 0)
    
End Sub

Sub survey()

Dim closest As Integer
Dim range As Integer

Mybot.post ("Doing survey")

closest = 2000
For scandir = 0 To 120 Step 7.5
    range = Mybot.scan(scandir, 5)
    If range > 0 And range < closest Then
        closest = range
        enemies(0).s(0).x = range * Sin(scandir / 57.3)
        enemies(0).s(0).y = range * Cos(scandir / 57.3)
        Mybot.post "Bogey at " + Str(enemies(0).s(0).x) + ", " + _
        Str(enemies(0).s(0).y)
    End If
    CheckDrive
Next scandir


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
    
    range = Mybot.scan(scandir, scanres)
    
    If range > 40 And range < 700 Then
        If range < 200 Then
            stat = Mybot.cannon(Int(scandir), range)
            If stat = 0 Then
                shells = shells - 1
                If shells = 0 Then
                    survey
                    shells = 4
                End If
            End If
        End If
        attack (scandir)
    End If
    
    ' Check if we've scanned past 90 degrees
    If scandir >= 90 Then
        ' reset scan to east
        scandir = 10
    Else
        scandir = scandir + 20
    End If

    
End Sub

Sub attack(targetdir As Single)
    
Dim stat As Integer
Dim myscan As Single
Dim startscan As Single
Dim endscan As Single
Dim range As Single
Dim delta As Single
        
    myscan = targetdir
    speed = 50
    stat = -1
    delta = 1.5
    While delta < 12
        range = Mybot.scan(myscan, 1)
        If (range > 0 And range < 700) Then
            stat = Mybot.cannon(Int(myscan), Int(range))
            If stat = 0 Then
                shells = shells - 1
                If shells = 0 Then
                    survey
                    shells = 4
                End If
            End If
            delta = 0
        Else
            If delta > 0 Then
                delta = delta * -1
            Else
                delta = (delta * -1) + 2
            End If
        End If
        myscan = (targetdir + delta) Mod 360
'       mybot.ShowStatus
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

    If Mybot.x < 150 And dir = 180 Then
        dir = 0
    End If
    speed = 100
    flight = Timer + 4
    Mybot.Drive dir, speed
    
End Sub


Sub UserInit()

dir = 0
speed = 35
Mybot.SetName ("Lurker")
scanres = 10
shells = 4
ccw = 1
GoHome

End Sub







