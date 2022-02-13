VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "First"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' speed1 is not visible. Each client gets access to
' private copy via class methods

Private speed1 As Integer

' Each client gets a copy, set to the index of their
' bot in the hidden array.

Private myindex As Integer

' speed4 is visible to clients. Each gets a
' private copy

Public speed4 As Integer

Public Function cannon(dir As Integer, range As Integer) As Integer

Dim dx As Single
Dim dy As Single
Dim tof As Single       ' Time of flight

If (dir < 0) Or (dir >= 360) Or (Server.Timer1.Enabled = False) Or (range > 700) Or (range < 40) Or (Bots(myindex).reload > 0) Or (Bots(myindex).health = 0) Then
    cannon = -1
    Server.StatBox(myindex - 1).Text = "Click"
    Exit Function
End If

dx = range * Cos(dir / 57.3)
dy = range * Sin(dir / 57.3)

Bots(myindex).tx = Bots(myindex).x + dx
If Bots(myindex).tx > 999 Then Bots(myindex).tx = 999
If Bots(myindex).tx < 0 Then Bots(myindex).tx = 0

Bots(myindex).ty = Bots(myindex).y + dy
If Bots(myindex).ty > 999 Then Bots(myindex).ty = 999
If Bots(myindex).ty < 0 Then Bots(myindex).ty = 0

tof = range / 20
' Add two ticks to show explosion
Bots(myindex).fire = tof + 2
Bots(myindex).reload = 40
Bots(myindex).dx = dx / tof
Bots(myindex).dy = dy / tof

' calculate reload time
Bots(myindex).shells = Bots(myindex).shells - 1
If Bots(myindex).shells = 0 Then
    Bots(myindex).shells = 4
    Bots(myindex).reload = 120
Else
    Bots(myindex).reload = 40
End If

' heat barrel
Bots(myindex).bheat = Bots(myindex).bheat + 20
' Let display know there's a new shot to handle
Bots(myindex).newshot = 1

Server.StatBox(myindex - 1).Text = "Bang"

End Function
Public Function drive(dir As Integer, speed As Integer) As Integer

' Range checking

If (Bots(myindex).health = 0 Or dir < 0 Or dir >= 360 Or speed < 0 Or speed > 100) Then
    drive = -1
    Exit Function
End If

drive = pvtdrive(myindex, dir, speed)

End Function

Public Function color() As Long

color = Bots(myindex).color

End Function








Public Function health() As Integer

health = Bots(myindex).health

End Function

Public Function heat() As Integer

heat = Int(Bots(myindex).mheat)

End Function


Public Function status() As String

status = Bots(myindex).status

End Function

Public Function time() As Integer

time = tick

End Function

Public Function dsp() As Integer

dsp = Bots(myindex).lastscanned

End Function

Public Function x() As Integer

x = Bots(myindex).x

End Function

Public Function Index() As Integer

Index = myindex

End Function

Public Function direction() As Integer

direction = Bots(myindex).dir

End Function


Public Function y() As Integer

y = Bots(myindex).y

End Function
Public Function scan(dir As Single, res As Integer) As Integer

Dim i As Integer
Dim dx As Single
Dim dy As Single
Dim td As Single
Dim range As Single
Dim closest As Single
Dim dt As Single

' A little range checking

If (Server.Timer1.Enabled = False Or Bots(myindex).health = 0 Or dir < 0 Or dir >= 360 Or res < 1 Or res > 10) Then
    scan = -1
    Exit Function
    Bots(myindex).scan = 0         ' Let world know we're scanning
End If

Bots(myindex).scan = 2          ' Let world know we're scanning

Bots(myindex).sdir = dir
Bots(myindex).sres = res

closest = 2000

For i = 1 To LastBot
    If i <> myindex And Bots(i).health > 0 Then         ' Don't check myself
        dx = Bots(i).x - Bots(myindex).x
        dy = Bots(i).y - Bots(myindex).y
        ' avoid div by zero
        ' Big problem: need to distinguish 45/225 degrees
        If (dx = 0) Then dx = 0.00001
        td = Atn(dy / dx) * 57.3 + 360    ' target direction
        If td > 360 Then td = td - 360
        ' If he's below us, theta to him is > 180
        If dx < 0 Then td = td + 180
        If td > 360 Then td = td - 360
        dt = Abs(td - dir) + 360         ' delta theta
        If dt > 360 Then dt = dt - 360
        If (dt <= res) Then  ' we can see him
            range = Sqr(dx ^ 2 + dy ^ 2)
            On Error GoTo DeadBot
            Call Bots(i).proc.ping(Str(myindex))
            On Error GoTo 0
            If range < closest Then
                closest = range
                Bots(myindex).lastscanned = i
            End If
       End If
    End If
Next i

' add scan errors here (heat etc)

If closest < 2000 And Bots(myindex).bheat < 35 Then
    closest = closest + ((Rnd + Rnd) - 1) * res
    If Rnd > 0.5 Then
        closest = closest + Bots(myindex).bheat
    Else
        closest = closest - Bots(myindex).bheat
    End If
    scan = closest
Else
    scan = 0
End If

' Normal exit
Exit Function

' Error handler
DeadBot:
'    Bots(i).health = 0
    Server.StatBox(i - 1).Text = "Ping failed"
    Bots(i).status = "K"
    Resume Next

End Function

Public Sub alert(ByRef proc As Object)

Set Bots(myindex).proc = proc

End Sub

Public Function speed() As Integer

speed = Bots(myindex).speed

End Function

Private Sub Class_Initialize()

Server.Visible = True

myindex = setindex()

If myindex > 0 Then
    PlaceBot (myindex)
    Server.StatBox(myindex - 1).BackColor = Bots(myindex).color
    Server.StatBox(myindex - 1).Text = "Registered"
End If

End Sub


