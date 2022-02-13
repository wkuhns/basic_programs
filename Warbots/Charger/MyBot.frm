VERSION 5.00
Begin VB.Form MyForm 
   Caption         =   "Form1"
   ClientHeight    =   605
   ClientLeft      =   1727
   ClientTop       =   2244
   ClientWidth     =   1518
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   605
   ScaleWidth      =   1518
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Left            =   242
      Top             =   242
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

Dim MyRobot As RobotLink

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
Static pursue As Long

    ' Check his location and avoid walls. This guy charges
    ' any enemy.
    
    x = MyRobot.x
    y = MyRobot.y

    If (x > 900) Then dir = 180
    If (x < 100) Then dir = 0
    If (y > 900) Then dir = 270
    If (y < 100) Then dir = 90

    If MyRobot.heat > 150 Then speed = 20
    Call MyRobot.Drive(dir, speed)
    
    ' Look for enemy
    
    range = MyRobot.scan(scandir, scanres)
    
    ' If we see an enemy, zero in and shoot
    
    If (range > 0 And range < 700) Then
        If scanres = 1 Then
            Call MyRobot.cannon(Int(scandir), Int(range))
            scandir = (scandir + 357) Mod 360
            pursue = pursue + 1
        Else
            scandir = (scandir + 350) Mod 360
            scanres = 1
            pursue = Timer + 5
        End If
        dir = scandir
        speed = 100
    End If
    
    If pursue < Timer Then
        ' Done chasing
        speed = 35
        scanres = 10
    End If
    
    scandir = (scandir + scanres * 2) Mod 360
    
End Sub

'
' This subroutine MUST exist with EXACTLY this name and
' this argument list. The contents are up to the user.
' This subroutine is invoked by the server when this robot
' is scanned by another robot.
'
Public Sub Ping(m As Integer)

    Call MyRobot.post("Pinged by: " + Str(m))

End Sub





Sub UserInit()

dir = 0
speed = 50
scanres = 10
MyRobot.SetName ("Charger")

End Sub

'
' Don't change this at all. This code creates the linkage
' to the robot server process.
'
Private Sub Form_Load()

' Create robot object

Set MyRobot = CreateObject("RobotAPI.RobotLink")

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
    If MyRobot.Status = "K" Then
        Die
        Exit Sub
    End If
    
    ' Do the user's cyclic stuff.
    UserFrame
    
Wend

End Sub

Private Sub Form_Unload(Cancel As Integer)

MsgBox ("AppForm unloading...")

Die

End Sub
Private Sub Timer1_Timer()

Dim e As Integer
    e = MyRobot.pinged
    If e <> 0 Then
        Ping (e)
    End If

End Sub





