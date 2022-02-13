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
Dim scandir As Single
Dim dir As Integer
Dim scanres As Integer

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
    
    ' Check his location and avoid walls. This guy circles
    ' counter-clockwise around the perimeter at constant
    ' speed.
    
    x = Mybot.x
    y = Mybot.y

    If (x > 900 And dir = 0) Then dir = 270: scandir = 270
    If (x > 900 And y > 900) Then dir = 270: scandir = 270
    If (y > 900 And x < 100) Then dir = 270: scandir = 90
    If (x < 500 And dir = 270 And y < 100) Then dir = 45: scandir = 45
    If (y < 100 And dir = 270 And x > 500) Then dir = 135: scandir = 135

    If Mybot.heat > 190 Then speed = 35
    If Mybot.heat < 50 Then speed = 100
    Call Mybot.Drive(dir, speed)
    
    ' Look for enemy
    
    range = Mybot.scan(scandir, 5)
    
    ' If we see an enemy, take a shot
    
    If (range > 0) Then
        Call Mybot.cannon(Int(scandir), Int(range))
    End If
    
    ' set scan going south on right side, at end of scan reset
    If (dir = 270 And x > 500 And scandir = 90) Then
        scandir = 270
    End If
    
    ' set scan going south on left side, at end of scan reset
    If (dir = 270 And x < 500 And scandir = 270) Then
        scandir = 90
    End If
    
    ' always scan clockwise, increment scan angle
    scandir = (scandir + 350) Mod 360
    
End Sub

'
' This subroutine MUST exist with EXACTLY this name and
' this argument list. The contents are up to the user.
' This subroutine is invoked by the server when this robot
' is scanned by another robot.
'
Public Sub Ping(m As Integer)
speed = 100
    Call Mybot.post("Pinged by: " + Str(m))

End Sub

Sub UserInit()

dir = 0
speed = 100
Mybot.SetName ("Jason")

End Sub

