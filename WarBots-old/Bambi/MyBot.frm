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

Dim MyRobot As RobotDLL.RobotLink

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
Dim Range As Single
Dim Goal As Integer

'
' Our life is over. Release robot's soul and die. Do not
' change this subroutine
'
Sub Die()

Set MyRobot = Nothing

End

End Sub
' Go to top right
Sub Home0()
       If Timer > flight Then speed = 35
       While MyRobot.y < 900
          dir = 90
          If Timer > flight Then speed = 35
          Call MyRobot.Drive(dir, speed)
          Call ScandirCommander
          Call MyRobot.ShowStatus
          If MyRobot.Status = "K" Then Die
       Wend
       While MyRobot.x < 900
          dir = 0
          If Timer > flight Then speed = 35
          Call MyRobot.Drive(dir, speed)
          Call ScandirCommander
          Call MyRobot.ShowStatus
          If MyRobot.Status = "K" Then Die
          If Goal <> 1 Then Exit Sub
       Wend
       speed = 0
       Call MyRobot.Drive(dir, speed)
       
End Sub

' Top Left
Sub Home1()
       If Timer > flight Then speed = 35
       While MyRobot.y < 900
          dir = 90
          If Timer > flight Then speed = 35
          Call MyRobot.Drive(dir, speed)
          Call ScandirCommander
          Call MyRobot.ShowStatus
          If MyRobot.Status = "K" Then Die
       Wend
       While MyRobot.x > 100
          dir = 180
          If Timer > flight Then speed = 35
          Call MyRobot.Drive(dir, speed)
          Call ScandirCommander
          Call MyRobot.ShowStatus
          If MyRobot.Status = "K" Then Die
          If Goal <> 1 Then Exit Sub
       Wend
       speed = 0
       Call MyRobot.Drive(dir, speed)
End Sub
'bottom left
Sub Home2()
       If Timer > flight Then speed = 35
       While MyRobot.y > 100
          dir = 270
          If Timer > flight Then speed = 35
          Call MyRobot.Drive(dir, speed)
          Call ScandirCommander
          Call MyRobot.ShowStatus
          If MyRobot.Status = "K" Then Die
       Wend
       While MyRobot.x > 100
          dir = 180
          If Timer > flight Then speed = 35
          Call MyRobot.Drive(dir, speed)
          Call ScandirCommander
          Call MyRobot.ShowStatus
          If MyRobot.Status = "K" Then Die
          If Goal <> 1 Then Exit Sub
       Wend
       speed = 0
       Call MyRobot.Drive(dir, speed)
   
End Sub
' Bottom right
Sub Home3()
       If Timer > flight Then speed = 35
       While MyRobot.y > 100
          dir = 270
          If Timer > flight Then speed = 35
          Call MyRobot.Drive(dir, speed)
          Call ScandirCommander
          Call MyRobot.ShowStatus
          If MyRobot.Status = "K" Then Die
       Wend
       While MyRobot.x < 900
          dir = 0
          If Timer > flight Then speed = 35
          Call MyRobot.Drive(dir, speed)
          Call ScandirCommander
          Call MyRobot.ShowStatus
          If MyRobot.Status = "K" Then Die
          If Goal <> 1 Then Exit Sub
       Wend
       speed = 0
       Call MyRobot.Drive(dir, speed)
End Sub


Sub ScandirCommander()
Dim quad(4) As Integer
Dim i As Integer
Dim x As Integer
Dim y As Integer

x = MyRobot.x
y = MyRobot.y

scandir = (scandir + (scanres * 2) - 1) Mod 360

For i = 0 To 3
    quad(i) = 1
Next i
    
' Left edge
If x < 150 Then
   quad(1) = 0
   quad(2) = 0
End If

' right
If x > 850 Then
   quad(0) = 0
   quad(3) = 0
End If

' top
If y > 850 Then
   quad(0) = 0
   quad(1) = 0
End If

' bottom
If y < 150 Then
   quad(2) = 0
   quad(3) = 0
End If

While quad(Int(scandir / 90)) = 0
    MyRobot.post "index " & Int(scandir / 90)
    scandir = (scandir + 30) Mod 360
Wend
'MyRobot.post "Scandir is " & scandir

Range = MyRobot.scan(scandir, scanres)

If Range > 0 And Range < 700 Then Call attack

End Sub

'
' This is where you put anything that you want your robot to
' do continuously. You may create your own subroutines and
' call them from here. Depending on the complexity of your
' code, this routine is run about four times per second as
' long as your robot is alive.
'
Public Sub UserFrame()
Dim st As Long
      
   If Timer > flight Then speed = 35
  
   Call MyRobot.Drive(dir, speed)
         
   If Goal = 0 Then Call Home0
   
   If Goal = 1 Then Call Home1
   
   If Goal = 2 Then Call Home2
   
   If Goal = 3 Then Call Home3
   
    
End Sub
Sub attack()
    
Dim stat As Integer
Dim myscan As Single
Dim startscan As Single
Dim endscan As Single
Dim Find As Integer
Dim scans As Integer
    
    scans = 10
    Find = 0
    
MyRobot.ShowStatus
  
' If he's really close, just shoot.
If Range > 0 And Range < 150 Then
    scanres = 10
    Call MyRobot.cannon(Int(scandir), Int(Range))
Else
    ' If he's farther, zoom in
    While scans > 0
        If scanres > 6 Then scanres = 6: scandir = (scandir + 360 - scanres) Mod 360
        Range = MyRobot.scan(scandir, scanres)
        If Range = 0 Then
            scandir = (scandir + 360 + scanres * 2 - 1) Mod 360
            scans = scans - 1
        Else
            If (scanres > 1.5) Then
                scanres = scanres / 2
                scandir = (scandir + 360 - scanres) Mod 360
                MyRobot.post "Zoom in to " & scanres
            Else
                Call MyRobot.cannon(Int(scandir), Int(Range * 0.95))
                scans = 0
            End If
        End If
    Wend
    scanres = 10
End If
   
End Sub
'
' This subroutine MUST exist with EXACTLY this name and
' this argument list. The contents are up to the user.
' This subroutine is invoked by the server when this robot
' is scanned by another robot.
'
Public Sub Ping(m As String)

    speed = 100
    flight = Timer + 2
    MyRobot.Drive dir, speed
'    MyRobot.post "Run Away!"
    
   If MyRobot.x < 150 And MyRobot.y < 150 Then Goal = 3
   
   If MyRobot.x < 150 And MyRobot.y > 850 Then Goal = 2
   
   If MyRobot.x > 850 And MyRobot.y > 850 Then Goal = 1
   
   If MyRobot.x > 850 And MyRobot.y < 150 Then Goal = 0
   
End Sub

Sub UserInit()

Dim xmin As Integer
Dim ymin As Integer

MyRobot.SetName ("Bambi")
flight = Timer

If MyRobot.x < (1000 - MyRobot.x) Then
    ' left
    xmin = MyRobot.x
Else
    ' right
    xmin = MyRobot.x - 1000
End If

If MyRobot.y < (1000 - MyRobot.y) Then
    ' bottom
    ymin = MyRobot.y
Else
    ' top
    ymin = MyRobot.y - 1000
End If

If Abs(xmin) < Abs(ymin) Then
    ' closer to left/right than top/bottom
    If xmin > 0 Then
        dir = 180
        MyRobot.post "left"
    Else
        dir = 0
        MyRobot.post "right"
    End If
Else
    If ymin > 0 Then
        dir = 270
        MyRobot.post "botom"
    Else
        dir = 90
        MyRobot.post "top"
    End If
End If

speed = 35
scanres = 10
scandir = 340
Goal = 1


End Sub

'
' Don't change this at all. This code creates the linkage
' to the robot server process.
'
Private Sub Form_Load()

' Create robot object

Set MyRobot = CreateObject("RobotDLL.RobotLink")

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





