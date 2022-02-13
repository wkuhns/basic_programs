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
Dim scandir As Single
Dim dir As Integer
Dim scanres As Integer

'
' Our life is over. Release robot's soul and die. Do not
' change this subroutine
'
Sub Die()

Set MyRobot = Nothing

End

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
    
    ' Check his location and avoid walls. This guy circles
    ' counter-clockwise around the perimeter at constant
    ' speed.
    
    x = MyRobot.x
    y = MyRobot.y

    If (x > 900 And dir = 0) Then dir = 90: scandir = 90
    If (x < 100 And dir = 180) Then dir = 270: scandir = 270
    If (y > 900 And dir = 90) Then dir = 180: scandir = 180
    If (y < 100 And dir = 270) Then dir = 0: scandir = 0

    If MyRobot.heat > 190 Then speed = 35
    If MyRobot.heat < 50 Then speed = 100
    Call MyRobot.Drive(dir, speed)
    
    ' Look for enemy
    
    range = MyRobot.scan(scandir, 10)
    
    ' If we see an enemy, take a shot
    
    If (range > 0) Then
    range = MyRobot.scan(scandir, 3)
    End If
    If (range > 0) Then
        Call MyRobot.cannon(Int(scandir), Int(range))
    End If
    
    ' Check if we've scanned past our left shoulder
    If scandir = (dir + 120) Mod 360 Then
        ' reset scan to just right of center
        scandir = (dir - 15) Mod 360
    End If

    scandir = (scandir + 5) Mod 360
    
End Sub

'
' This subroutine MUST exist with EXACTLY this name and
' this argument list. The contents are up to the user.
' This subroutine is invoked by the server when this robot
' is scanned by another robot.
'
Public Sub Ping(m As Integer)
    speed = 100
    MyRobot.post "Pinged by: " + Str(m)
    MyRobot.Drive dir, speed
End Sub





Sub UserInit()

dir = 0
speed = 35
MyRobot.SetName ("Bloody")

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





