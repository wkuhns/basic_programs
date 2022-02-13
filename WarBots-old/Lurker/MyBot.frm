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

    x = MyRobot.x
    y = MyRobot.y

    ' We're not running. Take time to turn around if we need to.
    If flight < Timer Then
        If (x > 200) Then dir = 180
        If (x < 100) Then dir = 0
        speed = 35
    End If
   
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

Sub GoHome()

Dim x As Single
Dim y As Single

    x = MyRobot.x
    y = MyRobot.y
    
    While y > 100
        Call MyRobot.Drive(270, 100)
        y = MyRobot.y
    Wend
    Call MyRobot.Drive(180, 35)
    While x > 100
        Call MyRobot.Drive(180, 100)
        x = MyRobot.x
    Wend
    Call MyRobot.Drive(0, 0)
    
End Sub

Sub survey()

Dim closest As Integer
Dim range As Integer

MyRobot.post ("Doing survey")

closest = 2000
For scandir = 0 To 120 Step 7.5
    range = MyRobot.scan(scandir, 5)
    If range > 0 And range < closest Then
        closest = range
        enemies(0).s(0).x = range * Sin(scandir / 57.3)
        enemies(0).s(0).y = range * Cos(scandir / 57.3)
        MyRobot.post "Bogey at " + Str(enemies(0).s(0).x) + ", " + _
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
    
    range = MyRobot.scan(scandir, scanres)
    
    If range > 40 And range < 700 Then
        If range < 200 Then
            stat = MyRobot.cannon(Int(scandir), range)
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
        range = MyRobot.scan(myscan, 1)
        If (range > 0 And range < 700) Then
            stat = MyRobot.cannon(Int(myscan), Int(range))
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

    If MyRobot.x < 150 And dir = 180 Then
        dir = 0
    End If
    speed = 100
    flight = Timer + 4
    MyRobot.Drive dir, speed
    
End Sub


Sub UserInit()

dir = 0
speed = 35
MyRobot.SetName ("Lurker")
scanres = 10
shells = 4
ccw = 1
GoHome

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





