VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RobotLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Public Event ping(m As String)
'Private client As Object

Public Function direction() As Integer

direction = Myrobot.direction

End Function

Public Function dsp() As Integer

dsp = Myrobot.dsp

End Function
Public Function index() As Integer

index = Myrobot.index

End Function
Public Sub Mark(x As Integer, y As Integer, c As Long)

Myrobot.Mark x, y, c

End Sub
Public Sub pause()

Myrobot.pause

End Sub

Public Function pinged() As Integer

pinged = Myrobot.pinged

End Function
Public Sub pinglink(m As String)

'RaiseEvent ping(m)
Myrobot.post ("Ping")

End Sub
Public Sub post(m As String)

Myrobot.post (m)

End Sub

Public Function RegisterAlert(ByRef obj As Object) As Integer

Call Myrobot.RegisterAlert(obj)
'client = obj
RegisterAlert = 1

End Function

Public Sub run()

Myrobot.run

End Sub

Public Function scan(dir As Single, res As Integer) As Integer

Sleep (200)

scan = Myrobot.scan(dir, res)

End Function

Public Function cannon(dir As Integer, range As Integer) As Integer

cannon = Myrobot.cannon(dir, range)

If cannon = -1 Then
    Sleep (100)
Else
    Sleep (200)
End If

DoEvents

End Function

Public Function drive(dir As Integer, speed As Integer) As Integer

drive = Myrobot.drive(dir, speed)

Sleep (100)
    
DoEvents

End Function

Public Sub SetName(name As String)

Myrobot.SetName (name)

End Sub


Public Function status() As String

status = Myrobot.status

End Function

Public Function time() As Single

time = Myrobot.time

End Function

Public Function WhereIs(Bot As Integer) As Long

WhereIs = Myrobot.WhereIs(Bot)

End Function
Public Function x() As Integer

x = Myrobot.x

End Function

Public Function color() As Long

color = Myrobot.color

End Function

Public Function health() As Integer

health = Myrobot.health

End Function


Public Function y() As Integer

y = Myrobot.y

End Function
Public Function speed() As Integer

speed = Myrobot.speed

End Function
Public Function heat() As Integer

heat = Myrobot.heat

End Function
Public Function bheat() As Integer

bheat = Myrobot.bheat

End Function
Public Function server_status() As String

server_status = Myrobot.server_status

End Function
Private Sub Class_Initialize()

Set Myrobot = CreateObject("RobotServer.Robot")

End Sub

Private Sub Class_Terminate()

Set Myrobot = Nothing

End Sub


