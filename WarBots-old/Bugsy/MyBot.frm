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

Sub Home1()
    If Timer > flight Then speed = 35
    If MyRobot.x < 100 And MyRobot.y > 100 Then dir = 270
    If MyRobot.x < 100 And MyRobot.y < 100 Then
        dir = 0
        speed = 0
    End If
    If MyRobot.x > 100 And MyRobot.y < 100 Then dir = 180
    If MyRobot.x > 100 And MyRobot.y > 100 Then dir = 225
         
       
End Sub
Sub Home2()
    If Timer > flight Then speed = 35
    If MyRobot.x < 100 And MyRobot.y < 900 Then dir = 90
    If MyRobot.x < 100 And MyRobot.y > 900 Then
        dir = 0
        speed = 0
    End If
    If MyRobot.x < 100 And MyRobot.y > 900 Then dir = 270
    If MyRobot.x > 100 And MyRobot.y > 100 Then dir = 180
    

End Sub
Sub Home3()
    If Timer > flight Then speed = 35
    If MyRobot.x < 900 And MyRobot.y > 900 Then dir = 0
    If MyRobot.x > 900 And MyRobot.y > 900 Then
        dir = 180
        speed = 0
    End If
    If MyRobot.x < 900 And MyRobot.y < 900 Then dir = 45
    If MyRobot.x > 900 And MyRobot.y < 900 Then dir = 90
       
   
End Sub

Sub Home4()
    If Timer > flight Then speed = 35
    If MyRobot.x < 900 And MyRobot.y < 100 Then dir = 0
    If MyRobot.x < 100 And MyRobot.y > 900 Then
        dir = 90
        speed = 0
    End If
    If MyRobot.x > 900 And MyRobot.y > 100 Then dir = 270
    If MyRobot.x < 900 And MyRobot.y > 100 Then dir = 315
    

End Sub


Sub Realtor()
Select Case Goal
    Case 1: Home1
    Case 2: Home2
    Case 3: Home3
    Case 4: Home4
    Case Else: MyRobot.post "Homeless"
End Select


End Sub

Sub ScandirCommander()

scanres = 10
scandir = (scandir + 19) Mod 360
Range = MyRobot.scan(scandir, scanres)

If Range > 0 Then Call attack
' Bottom Left
If MyRobot.x < 100 And MyRobot.y < 100 Then
   If scandir > 100 And scandir < 120 Then scandir = (scandir + 240) Mod 360
End If
' Left
If MyRobot.x < 100 And MyRobot.y > 100 And MyRobot.y < 900 Then
   If scandir > 100 And scandir < 120 Then scandir = (scandir + 140) Mod 360
End If
' Top Left
If MyRobot.x < 100 And MyRobot.y > 900 Then
   If scandir > 20 And scandir < 50 Then scandir = (scandir + 240) Mod 360
End If
' Top
If MyRobot.y > 900 And MyRobot.x > 100 And MyRobot.x < 900 Then
   If scandir > 20 And scandir < 50 Then scandir = (scandir + 140) Mod 360
End If
' Top Right
If MyRobot.x > 900 And MyRobot.y > 900 Then
   If scandir > 290 Then scandir = (scandir + 240) Mod 360
End If
' Right
If MyRobot.x > 900 And MyRobot.y > 100 And MyRobot.y < 900 Then
   If scandir > 290 Then scandir = (scandir + 140) Mod 360
End If
' Bottom Right
If MyRobot.x > 900 And MyRobot.y < 100 Then
   If scandir > 190 Then scandir = (scandir + 240) Mod 360
End If
' Bottom
If MyRobot.y < 100 And MyRobot.x > 100 And MyRobot.x < 900 Then
   If scandir > 200 And scandir < 220 Then scandir = (scandir + 140) Mod 360
End If

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
    Realtor
    Call MyRobot.Drive(dir, speed)
    Call ScandirCommander
    Call MyRobot.ShowStatus
    If MyRobot.status = "K" Then Die
      
 
    
End Sub
Sub attack()
    
Dim stat As Integer
Dim myscan As Single
Dim startscan As Single
Dim endscan As Single
Dim Find As Integer
Dim Clos As Integer
    Clos = 0
    Find = 0
    ' range = MyRobot.scan(scandir, scanres)
    MyRobot.ShowStatus
  
  While Range > 0 And Range < 150
        scanres = 10
        Call MyRobot.cannon(Int(scandir), Int(Range))
        MyRobot.post "Fire 1!"
        Range = MyRobot.scan(scandir, scanres)
    
        If Range = 0 Then
           scandir = (scandir + 351) Mod 360
           Range = MyRobot.scan(scandir, scanres)
        End If
        Call MyRobot.ShowStatus
        Realtor
  Wend
    
  While Range > 150 And Range < 700
      scanres = 6
      While scanres = 6
        scandir = (scandir + 355) Mod 360
        Range = MyRobot.scan(scandir, scanres)
    
        If Range = 0 Then
           scandir = (scandir + 10) Mod 360
           Range = MyRobot.scan(scandir, scanres)
        End If
     
        If Range > 0 And Range < 700 Then
           scanres = 3
           scandir = scandir + 357.5
           If scandir >= 360 Then scandir = scandir - 360
           Range = MyRobot.scan(scandir, scanres)
        End If
        Call MyRobot.ShowStatus
        Realtor
        If Range = 0 Then
            scandir = scandir + 5
            Range = MyRobot.scan(scandir, scanres)
        End If
        If Range = 0 Then scanres = 10
      Wend
      
        While scanres = 3
          MyRobot.pause
            While Range > 0 And Range < 700
               scanres = 1.5
               scandir = scandir + 358.5
               If scandir >= 360 Then scandir = scandir - 360
               Range = MyRobot.scan(scandir, scanres)
        
               If Range = 0 Then
                  scandir = (scandir + 3) Mod 360
                  Range = MyRobot.scan(scandir, scanres)
               End If
         
               If Range > 0 And Range < 700 Then
                   Call MyRobot.cannon(Int(scandir), Int(Range))
                   MyRobot.post "Fire 2!"
                   Range = MyRobot.scan(scandir, scanres)
               End If
               Call MyRobot.ShowStatus
               Realtor
             Wend
            Realtor
        Wend
        scanres = 10
        Range = MyRobot.scan(scandir, scanres)
        Realtor
      Wend
   
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
    MyRobot.post "Run Away!"
    
   If MyRobot.x < 150 And MyRobot.y < 150 Then Goal = 2
   
   If MyRobot.x < 150 And MyRobot.y > 850 Then Goal = 3
   
   If MyRobot.x > 850 And MyRobot.y > 850 Then Goal = 4
   
   If MyRobot.x > 850 And MyRobot.y < 150 Then Goal = 1
   
End Sub





Sub UserInit()
MyRobot.SetName ("Bugsy")
flight = Timer
dir = 0
speed = 35
scanres = 5
scandir = 340
If MyRobot.x < 500 And MyRobot.y < 500 Then Goal = 1
If MyRobot.x < 500 And MyRobot.y > 500 Then Goal = 2
If MyRobot.x > 500 And MyRobot.y > 500 Then Goal = 3
If MyRobot.x > 500 And MyRobot.y < 500 Then Goal = 4


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





