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

Sub Home1()
    If Timer > flight Then speed = 35
    If Mybot.x < 100 And Mybot.y > 100 Then dir = 270
    If Mybot.x < 100 And Mybot.y < 100 Then
        dir = 0
        speed = 0
    End If
    If Mybot.x > 100 And Mybot.y < 100 Then dir = 180
    If Mybot.x > 100 And Mybot.y > 100 Then dir = 225
         
       
End Sub
Sub Home2()
    If Timer > flight Then speed = 35
    If Mybot.x < 100 And Mybot.y < 900 Then dir = 90
    If Mybot.x < 100 And Mybot.y > 900 Then
        dir = 0
        speed = 0
    End If
    If Mybot.x < 100 And Mybot.y > 900 Then dir = 270
    If Mybot.x > 100 And Mybot.y > 100 Then dir = 180
    

End Sub
Sub Home3()
    If Timer > flight Then speed = 35
    If Mybot.x < 900 And Mybot.y > 900 Then dir = 0
    If Mybot.x > 900 And Mybot.y > 900 Then
        dir = 180
        speed = 0
    End If
    If Mybot.x < 900 And Mybot.y < 900 Then dir = 45
    If Mybot.x > 900 And Mybot.y < 900 Then dir = 90
       
   
End Sub

Sub Home4()
    If Timer > flight Then speed = 35
    If Mybot.x < 900 And Mybot.y < 100 Then dir = 0
    If Mybot.x < 100 And Mybot.y > 900 Then
        dir = 90
        speed = 0
    End If
    If Mybot.x > 900 And Mybot.y > 100 Then dir = 270
    If Mybot.x < 900 And Mybot.y > 100 Then dir = 315
    

End Sub


Sub Realtor()
Select Case Goal
    Case 1: Home1
    Case 2: Home2
    Case 3: Home3
    Case 4: Home4
    Case Else: Mybot.post "Homeless"
End Select


End Sub

Sub ScandirCommander()

scanres = 10
scandir = (scandir + 19) Mod 360
Range = Mybot.scan(scandir, scanres)

If Range > 0 Then Call attack
' Bottom Left
If Mybot.x < 100 And Mybot.y < 100 Then
   If scandir > 100 And scandir < 120 Then scandir = (scandir + 240) Mod 360
End If
' Left
If Mybot.x < 100 And Mybot.y > 100 And Mybot.y < 900 Then
   If scandir > 100 And scandir < 120 Then scandir = (scandir + 140) Mod 360
End If
' Top Left
If Mybot.x < 100 And Mybot.y > 900 Then
   If scandir > 20 And scandir < 50 Then scandir = (scandir + 240) Mod 360
End If
' Top
If Mybot.y > 900 And Mybot.x > 100 And Mybot.x < 900 Then
   If scandir > 20 And scandir < 50 Then scandir = (scandir + 140) Mod 360
End If
' Top Right
If Mybot.x > 900 And Mybot.y > 900 Then
   If scandir > 290 Then scandir = (scandir + 240) Mod 360
End If
' Right
If Mybot.x > 900 And Mybot.y > 100 And Mybot.y < 900 Then
   If scandir > 290 Then scandir = (scandir + 140) Mod 360
End If
' Bottom Right
If Mybot.x > 900 And Mybot.y < 100 Then
   If scandir > 190 Then scandir = (scandir + 240) Mod 360
End If
' Bottom
If Mybot.y < 100 And Mybot.x > 100 And Mybot.x < 900 Then
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
    Call Mybot.Drive(dir, speed)
    Call ScandirCommander
         
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
    ' range = mybot.scan(scandir, scanres)
  
  While Range > 0 And Range < 150
        scanres = 10
        Call Mybot.cannon(Int(scandir), Int(Range))
        Mybot.post "Fire 1!"
        Range = Mybot.scan(scandir, scanres)
    
        If Range = 0 Then
           scandir = (scandir + 351) Mod 360
           Range = Mybot.scan(scandir, scanres)
        End If
        Realtor
  Wend
    
  While Range > 150 And Range < 700
      scanres = 6
      While scanres = 6
        scandir = (scandir + 355) Mod 360
        Range = Mybot.scan(scandir, scanres)
    
        If Range = 0 Then
           scandir = (scandir + 10) Mod 360
           Range = Mybot.scan(scandir, scanres)
        End If
     
        If Range > 0 And Range < 700 Then
           scanres = 3
           scandir = scandir + 357.5
           If scandir >= 360 Then scandir = scandir - 360
           Range = Mybot.scan(scandir, scanres)
        End If
        Realtor
        If Range = 0 Then
            scandir = scandir + 5
            Range = Mybot.scan(scandir, scanres)
        End If
        If Range = 0 Then scanres = 10
      Wend
      
        While scanres = 3
          Mybot.pause
            While Range > 0 And Range < 700
               scanres = 1.5
               scandir = scandir + 358.5
               If scandir >= 360 Then scandir = scandir - 360
               Range = Mybot.scan(scandir, scanres)
        
               If Range = 0 Then
                  scandir = (scandir + 3) Mod 360
                  Range = Mybot.scan(scandir, scanres)
               End If
         
               If Range > 0 And Range < 700 Then
                   Call Mybot.cannon(Int(scandir), Int(Range))
                   Mybot.post "Fire 2!"
                   Range = Mybot.scan(scandir, scanres)
               End If
               Realtor
             Wend
            Realtor
        Wend
        scanres = 10
        Range = Mybot.scan(scandir, scanres)
        Realtor
      Wend
   
End Sub
'
' This subroutine MUST exist with EXACTLY this name and
' this argument list. The contents are up to the user.
' This subroutine is invoked by the server when this robot
' is scanned by another robot.
'
Public Sub Ping(m As Integer)

    speed = 100
    flight = Timer + 2
    Mybot.Drive dir, speed
    Mybot.post "Run Away!"
    
   If Mybot.x < 150 And Mybot.y < 150 Then Goal = 2
   
   If Mybot.x < 150 And Mybot.y > 850 Then Goal = 3
   
   If Mybot.x > 850 And Mybot.y > 850 Then Goal = 4
   
   If Mybot.x > 850 And Mybot.y < 150 Then Goal = 1
   
End Sub

Sub UserInit()
Mybot.SetName ("Bugsy")
flight = Timer
dir = 0
speed = 35
scanres = 5
scandir = 340
If Mybot.x < 500 And Mybot.y < 500 Then Goal = 1
If Mybot.x < 500 And Mybot.y > 500 Then Goal = 2
If Mybot.x > 500 And Mybot.y > 500 Then Goal = 3
If Mybot.x > 500 And Mybot.y < 500 Then Goal = 4


End Sub

