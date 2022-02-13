VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Robot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
 
Public Function direction() As Integer

direction = mydirection

End Function

Public Function dsp() As Integer

dsp = Myrobot.dsp

End Function
Public Sub Mark(x As Integer, y As Integer, c As Long)

Myrobot.Mark x, y, c

End Sub

Public Sub pause()

Myrobot.pause

End Sub

Public Sub post(m As String)

If Myrobot.health > 0 Then
    MyForm.PostBox.Text = MyForm.PostBox.Text + vbCrLf + m
    MyForm.PostBox.SelStart = 10000
End If

End Sub

Public Function RegisterAlert(ByRef obj As Object) As Integer

Call Myrobot.alert(obj)

RegisterAlert = 1

End Function

Public Sub run()

Myrobot.run

End Sub

Public Function scan(dir As Single, res As Integer) As Integer

Sleep (200)
DoEvents

scan = Myrobot.scan(dir, res)
myx = Myrobot.x
myy = Myrobot.y

End Function

Public Function cannon(dir As Integer, range As Integer) As Integer

cannon = Myrobot.cannon(dir, range)

If cannon = -1 Then
    Sleep (100)
Else
    Sleep (200)
End If

'DoEvents

End Function

Public Function drive(dir As Integer, speed As Integer) As Integer

drive = Myrobot.drive(dir, speed)

Sleep (100)
    
'DoEvents

End Function

Public Sub SetName(name As String)

MyForm.Caption = name + " (" + Str(Myrobot.Index) + ")"

End Sub

Public Sub ShowStatus()
    
Dim color As Long

    Sleep (100)
    
    'DoEvents
    
    ' If we're being called too often, just return
'    If Timer < lastaccess + 0.2 Then
'        Sleep (100)
'        DoEvents
'        Exit Sub
'    End If
'    lastaccess = Timer
    
    If stat <> "W" And Myrobot.status = "W" Then
        post ("*** WINNER ***")
    End If
    
    On Error GoTo deadbot
    
    stat = Myrobot.status
    
    myhealth = Myrobot.health
    myx = Myrobot.x
    myy = Myrobot.y
    myheat = Myrobot.heat
    mydirection = Myrobot.direction
    myspeed = Myrobot.speed
    
    MyForm.SpeedBox = myspeed
    MyForm.HealthBox = myhealth
    MyForm.DirBox = mydirection
    MyForm.HeatBox = myheat
        
    color = vbGreen
    If health < 25 Then
        color = vbYellow
    End If
    If health < 10 Then
        color = &H80FF&
    End If
    If health < 5 Then
        color = vbRed
    End If
            
    If health = 0 Then
        color = vbBlack
    End If
            
    MyForm.HealthBox.BackColor = color
    
    On Error GoTo 0
    
    Exit Sub

deadbot:
    post "Problem getting status"
    
End Sub

Public Function status() As String

status = stat

End Function

Public Function time() As Integer

time = Myrobot.time

End Function

Public Function WhereIs(Bot As Integer) As Long

WhereIs = Myrobot.WhereIs(Bot)

End Function
Public Function x() As Integer

x = myx

End Function

Public Function color() As Long

color = Myrobot.color

End Function

Public Function health() As Integer

health = myhealth

End Function


Public Function y() As Integer

y = myy

End Function
Public Function speed() As Integer

speed = myspeed

End Function
Public Function heat() As Integer

heat = myheat

End Function
Public Function bHeat() As Integer

bHeat = Myrobot.bHeat

End Function




Private Sub Class_Initialize()

Set MyForm = New StatForm
Set Myrobot = CreateObject("RobotServer.ServerRobot")

MyForm.Top = (MyForm.Height + 10) * (Myrobot.Index - 1)
MyForm.Shape1.FillColor = Myrobot.color
MyForm.Enabled = True
MyForm.Visible = True
' Place default caption in title bar
MyForm.Caption = "Robot " + Str(Myrobot.Index)

End Sub


Private Sub Class_Terminate()

'stat = "K"

' MsgBox ("Robot terminating...")

Set Myrobot = Nothing
Set MyForm = Nothing

End Sub


