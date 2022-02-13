VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ServerRobot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Each client gets a copy of myindex, set to the index
' of their bot in the hidden array.

Private myindex As Integer


Public Function cannon(ByVal dir As Integer, ByVal range As Integer) As Integer

Dim dx As Single
Dim dy As Single
Dim tof As Single       ' Time of flight

' Check for valid command. Exit if not valid.
If (dir < 0) Or (dir >= 360) Or _
    (Server.Timer1.Enabled = False) Or _
    (range < 40) Or _
    (Bots(myindex).reload > 0) Or _
    (Bots(myindex).health = 0) Then
    cannon = -1
    Exit Function
End If

If range > 700 Then range = 700

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
Public Function direction() As Integer

direction = Bots(myindex).dir

End Function

Public Function drive(ByVal dir As Integer, ByVal speed As Integer) As Integer

' Range checking

If (Bots(myindex).health = 0 Or dir < 0 Or _
    dir >= 360 Or speed < 0 Or speed > 100) Then
    drive = -1
    Exit Function
End If

' Call hidden drive function
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


Property Get HeatHealth() As Long

HeatHealth = Int(Bots(myindex).mheat) * 1000 + Bots(myindex).health

End Property

Public Sub Pause()

Dim m As String

If DebugState Then
    m = Server.Text2.Text
    Server.PauseBtn_Click
    Server.Text2.Text = "USER PAUSED: " + m
End If

End Sub

Public Sub Place(r As Integer, x As Integer, y As Integer)

If DebugState Then
    Bots(r).x = x
    Bots(r).y = y
End If

End Sub

Public Sub Run()

If DebugState Then
    Server.RunBtn_Click
End If

End Sub


Property Get SpeedDirection() As Long

SpeedDirection = Int(Bots(myindex).speed) * 1000 + Bots(myindex).dir


End Property


Public Function status() As String

status = Bots(myindex).status

End Function

Property Get time() As Integer

time = tick

End Property

Public Sub mark(x As Integer, y As Integer, c As Long)

    Arena.DrawMark x, y, c

End Sub

' Report who was last scanned
Public Function dsp() As Integer

dsp = Bots(myindex).lastscanned

End Function

Public Function WhereIs(Bot As Integer) As Long

If DebugState Then
    WhereIs = Int(Bots(Bot).x) * 1000 + Int(Bots(Bot).y)
Else
    WhereIs = 0
End If

End Function
Public Function bheat() As Integer

If DebugState Then
    bheat = Bots(myindex).bheat
Else
    bheat = 0
End If

End Function

Property Get x() As Integer

x = Bots(myindex).x

End Property
' Report who we are
Public Function Index() As Integer

Index = myindex

End Function

Property Get xy() As Long

xy = Int(Bots(myindex).x) * 1000 + Bots(myindex).y

End Property
Property Get y() As Integer

y = Bots(myindex).y

End Property
Public Function scan(ByVal dir As Single, ByVal res As Integer) As Integer

Dim i As Integer
Dim dx As Single
Dim dy As Single
Dim td As Single
Dim range As Single
Dim closest As Single
Dim dt As Single

' A little range checking

If (Server.Timer1.Enabled = False Or _
    Bots(myindex).health = 0 Or dir < 0 Or dir >= 360 _
    Or res < 1 Or res > 10) Then
    scan = -1
    Bots(myindex).scan = 0
    Exit Function
End If

Bots(myindex).scan = 2          ' Let display know we're scanning

Bots(myindex).sdir = dir
Bots(myindex).sres = res

closest = 2000

' Check each bot to see if we can see him
For i = 1 To LastBot
    ' Ignore ourselves and dead 'bots
    If i <> myindex And Bots(i).health > 0 Then         ' Don't check myself
        dx = Bots(i).x - Bots(myindex).x
        dy = Bots(i).y - Bots(myindex).y
        ' Rect-to-polar conversion - avoid div by zero
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
            Call Bots(i).proc.ping(myindex)
            On Error GoTo 0
            If range < closest Then
                closest = range
                Bots(myindex).lastscanned = i
            End If
       End If
    End If
Next i

' scan errors here (heat etc)

If closest < 2000 And Bots(myindex).bheat < 35 Then
    closest = closest + ((Rnd + Rnd) - 1) * res
    If Rnd > 0.5 Then
        closest = closest + Bots(myindex).bheat
    Else
        closest = closest - Bots(myindex).bheat
    End If
    scan = closest
    If DebugState Then
        Server.Text1 = Str(time) + ": " + Str(myindex) + ">" + _
        Str(Bots(myindex).lastscanned) + " r= " + Str(closest) + _
        " loc= " + Str(Bots(Bots(myindex).lastscanned).x) + "," + _
        Str(Bots(Bots(myindex).lastscanned).y) + " bh= " + Str(Bots(myindex).bheat)
    End If
Else
    scan = 0
End If

' Normal exit
Exit Function

' Error handler. Robot process probably died.
DeadBot:
'    Bots(i).health = 0
    Server.StatBox(i - 1).Text = "Ping failed"
    If Not DebugState Then
        Bots(i).status = "K"
    End If
    Resume Next

End Function
' Register the user's form object, which MUST contain
' a public subroutine called Ping.
'
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


