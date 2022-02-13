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
' This file contains the properties and methods for the robot object.
' Each user robot instantiates one of these, and their control program
' has access to the public methods.

' Bot data is kept in a public array. Each client gets a copy of
' myindex, set to the index of their bot in the public array.

Private myindex As Integer

' Each instance gets a status form.
Private MyForm As StatForm

'
Public Property Get rServer() As rServer
    Set rServer = grServer
End Property

' One of many public 'property gets' - allow user to read direction
Property Get direction() As Integer

direction = Bots(myindex).dir

End Property
'
Property Get color() As Long

color = Bots(myindex).color

End Property
'
Property Get heat() As Integer

heat = Int(Bots(myindex).mHeat)

End Property
'
Property Get status() As String

status = Bots(myindex).status

End Property
'
Property Get server_status() As String

server_status = s_status

End Property
'
Property Get x() As Integer

x = Bots(myindex).x

End Property
'
Property Get y() As Integer

y = Bots(myindex).y

End Property
'
Property Get health() As Integer

health = Bots(myindex).health

End Property
'
Property Get speed() As Integer

speed = Bots(myindex).speed

End Property
' Report who we are
Public Function Index() As Integer

Index = myindex

End Function
' Time
Property Get time() As Single

time = tick / quanta

End Property
' Fire our cannon
Public Function cannon(ByVal dir As Integer, ByVal range As Integer) As Integer

Dim dx As Single
Dim dy As Single
Dim tof As Single       ' Time of flight

' Check for valid command. Exit if not valid.
If (dir < 0) Or (dir >= 360) Or _
    (serverbox.Timer1.Enabled = False) Or _
    (range < 40) Or _
    (Bots(myindex).reload > 0) Or _
    (Bots(myindex).health = 0) Or range < 0 Then
    cannon = -1
    Exit Function
End If

' Cannon only has 700M range
If range > 700 Then range = 700

dx = range * Cos(dir / 57.3)
dy = range * Sin(dir / 57.3)

' Set target x and target y
Bots(myindex).tx = Bots(myindex).x + dx
If Bots(myindex).tx > 999 Then Bots(myindex).tx = 999
If Bots(myindex).tx < 0 Then Bots(myindex).tx = 0

Bots(myindex).ty = Bots(myindex).y + dy
If Bots(myindex).ty > 999 Then Bots(myindex).ty = 999
If Bots(myindex).ty < 0 Then Bots(myindex).ty = 0

' Time of flight in seconds
tof = range / 200#
' We'll calculate fire in the arena code.
Bots(myindex).fire = range
Bots(myindex).reload = 4
' dx and dy are distance-per-second at this point
Bots(myindex).dx = dx / tof
Bots(myindex).dy = dy / tof

' calculate reload time
Bots(myindex).shells = Bots(myindex).shells - 1
If Bots(myindex).shells = 0 Then
    Bots(myindex).shells = 4
    Bots(myindex).reload = 12
Else
    Bots(myindex).reload = 4
End If

' heat barrel
Bots(myindex).bHeat = Bots(myindex).bHeat + 20
' Let display server know there's a new shot to handle
Bots(myindex).newshot = 1

If ServerForm.SoundBox.Value = 1 Then
    Call EZPlay(bang(myindex), ssFile)
End If

End Function
' send command to drive system
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
' Use radar
Public Function scan(ByVal dir As Single, ByVal res As Integer) As Integer

Dim i As Integer
Dim dx As Single
Dim dy As Single
Dim td As Single
Dim range As Single
Dim closest As Single
Dim dt As Single

' A little range checking

If (serverbox.Timer1.Enabled = False Or _
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
            On Error GoTo deadbot
            'Call Bots(i).proc.pinglink(Str(myindex))
            ' Set var to let bot know I pinged him
            Bots(i).pingnotify = myindex
            On Error GoTo 0
            If range < closest Then
                closest = range
                ' Set var so I can see who I pinged
                Bots(myindex).lastscanned = i
            End If
       End If
    End If
Next i

' scan errors here (heat etc)

If closest < 2000 And Bots(myindex).bHeat < 35 Then
    closest = closest + ((Rnd + Rnd) - 1) * res
    If Rnd > 0.5 Then
        closest = closest + Bots(myindex).bHeat
    Else
        closest = closest - Bots(myindex).bHeat
    End If
    scan = closest
Else
    scan = 0
End If

' Normal exit
Exit Function

' Error handler. Robot process probably died.
deadbot:
'    Bots(i).health = 0
    serverbox.StatBox(i - 1).Text = "Ping failed"
    If Not DebugState Then
        Bots(i).status = "K"
    End If
    Resume Next

End Function
' Report who was last scanned
Public Function dsp() As Integer

dsp = Bots(myindex).lastscanned

End Function
' Display message on our status window
Public Sub post(m As String)

If Bots(myindex).health > 0 Then
    MyForm.PostBox.Text = MyForm.PostBox.Text + vbCrLf + m
    MyForm.PostBox.SelStart = 10000
End If

End Sub
' Set our robot's name
Public Sub SetName(name As String)

MyForm.Caption = name + " (" + Str(myindex) + ")"
serverbox.StatBox(myindex - 1).Text = name
Bots(myindex).name = name
End Sub
' One of many 'debug' control functions - pause server
Public Sub pause()

If DebugState Then
    serverbox.PauseBtn_Click
End If

End Sub
' Polled ping notification - ID of bot who pinged me
Public Function pinged() As Integer
    pinged = Bots(myindex).pingnotify
    Bots(myindex).pingnotify = 0
End Function
' 'debug' control functions
Public Sub run()

If DebugState Then
    serverbox.RunBtn_Click
End If

End Sub
' Allow user to draw mark on arena
Public Sub Mark(x As Integer, y As Integer, c As Long)

    arena.DrawMark x, y, c

End Sub

Public Function WhereIs(Bot As Integer) As Long

If DebugState Then
    WhereIs = Int(Bots(Bot).x) * 1000 + Int(Bots(Bot).y)
Else
    WhereIs = 0
End If

End Function
Public Function bHeat() As Integer

If DebugState Then
    bHeat = Bots(myindex).bHeat
Else
    bHeat = 0
End If

End Function
' Register the user's form object, which MUST contain
' a public subroutine called Ping.
'
Public Sub RegisterAlert(ByRef proc As Object)

Set Bots(myindex).proc = proc
'RegisterAlert = 1

End Sub
Friend Sub UpdateStatus()

Dim color As Long
Dim health As Integer
    
    If Bots(myindex).status = "W" Then
        post ("*** WINNER ***")
        Bots(myindex).status = "K"
    End If
    
    On Error GoTo deadbot
    
    health = Bots(myindex).health
    
    MyForm.SpeedBox = Bots(myindex).speed
    MyForm.HealthBox = health
    MyForm.DirBox = Bots(myindex).dir
    MyForm.HeatBox = Bots(myindex).mHeat
        
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
Private Sub Class_Initialize()

If grServer Is Nothing Then
    Set grServer = New rServer
End If
gUseCount = gUseCount + 1

myindex = setindex()

If myindex > 0 Then
    PlaceBot (myindex)
    serverbox.StatBox(myindex - 1).BackColor = Bots(myindex).color
    serverbox.StatBox(myindex - 1).Text = "Registered:"
End If

Set MyForm = New StatForm

MyForm.Timer1.Interval = 500 + myindex

MyForm.Top = (MyForm.Height + 10) * (myindex - 1)
MyForm.Shape1.FillColor = Bots(myindex).color
MyForm.Enabled = True
MyForm.Visible = True
' Place default caption in title bar
MyForm.Caption = "Robot " + Str(myindex)

End Sub

Private Sub Class_Terminate()

gUseCount = gUseCount - 1
If gUseCount = 0 Then
    Set grServer = Nothing
End If

Set MyForm = Nothing

End Sub
