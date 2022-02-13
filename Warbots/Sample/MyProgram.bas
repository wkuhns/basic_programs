Attribute VB_Name = "MyProgram"
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

' User defined variables:
' These are not required except as used by your application
'
Dim speed As Integer        ' How fast do we want to go?
Dim scandir As Single       ' What direction do we want point our scanner?
Dim dir As Integer          ' What direction do we want to go?
'
' This is where you put anything that you want your robot to
' do continuously. You may create your own subroutines and
' call them from here. Depending on the complexity of your
' code, this routine is run about four times per second as
' long as your robot is alive.
'
Public Sub UserFrame()

' Perform a cycle of calculations for our robot.

Dim x As Integer        ' Store horizontal position of our robot
Dim y As Integer        ' Store vertical position of our robot
Dim range As Integer    ' Store distance to enemy from scanner
    
    ' Check his location and avoid walls. This guy circles
    ' counter-clockwise around the perimeter.
    
    x = Mybot.x
    y = Mybot.y

    If (x > 900 And dir = 0) Then dir = 90: scandir = 90
    If (x < 100 And dir = 180) Then dir = 270: scandir = 270
    If (y > 900 And dir = 90) Then dir = 180: scandir = 180
    If (y < 100 And dir = 270) Then dir = 0: scandir = 0

    ' Check motor temp. Adjust speed to cool motors if necessary
    If Mybot.heat > 190 Then speed = 35
    If Mybot.heat < 50 Then speed = 100
    Call Mybot.Drive(dir, speed)
    
    ' Look for enemy. Use +/-5 degree scan
    
    range = Mybot.scan(scandir, 5)
    
    ' If we see an enemy, take a shot
    
    If (range > 0) Then
        Call Mybot.cannon(Int(scandir), Int(range))
    End If
    
    ' Check if we've scanned past our left shoulder.
    ' If so, we want to reset our scanner to just right of center
    If scandir = (dir + 180) Mod 360 Then
        scandir = (dir + 340) Mod 360
    End If

    ' Move scanner 10 degrees counterclockwise
    scandir = (scandir + 10) Mod 360
    
End Sub

' This routine is executed once, when the robot is 'created' -
' (placed in the arena). Set initial values here, and set
' the robot's name.

Sub UserInit()

' Start going to the right at top speed
dir = 0
speed = 100
' Tell server who we are
Mybot.SetName ("Sample")

End Sub

' Respond to being scanned by another robot. In this case we
' just display a message.

Public Sub ping(m As Integer)
    
    Call Mybot.post("Pinged by: " + Str(m))
    
End Sub
