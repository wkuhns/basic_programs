VERSION 1.0 CLASS
BEGIN
  MultiUse = 0   'False
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Client"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
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

' User defined globals:
' These are 'global' to this form. Use these or add your own.
' They are not required except as used by your application

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
    ' counter-clockwise around the perimeter at constant
    ' speed.
    
    x = Mybot.x
    y = Mybot.y

    If (x > 900 And dir = 0) Then dir = 90: scandir = 90
    If (x < 100 And dir = 180) Then dir = 270: scandir = 270
    If (y > 900 And dir = 90) Then dir = 180: scandir = 180
    If (y < 100 And dir = 270) Then dir = 0: scandir = 0

    If Mybot.heat > 190 Then speed = 35
    If Mybot.heat < 50 Then speed = 100
    Call Mybot.Drive(dir, 50)
    
    ' Look for enemy
    
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
' (placed in the arena). Set initial values here, and set the robot's name.
'
Sub UserInit()

dir = 0
speed = 100
Mybot.SetName ("Sample")

End Sub

'
' We spend our whole life in this loop. Because robot functions sleep,
' we don't tie up all CPU time.
'
Public Sub LoopForever()

Dim mystatus As String

While mystatus <> "K"

    mystatus = Mybot.Status
    If mystatus <> "P" Then
        ' Do the user's cyclic stuff.
        UserFrame
    End If
    
Wend

Die

End Sub
'
' Our life is over. Release robot's soul and die. Do not
' change this subroutine
'
Public Sub Die()

Set Mybot = Nothing

End

End Sub

