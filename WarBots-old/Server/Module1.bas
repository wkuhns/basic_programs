Attribute VB_Name = "Srvmod"
Option Explicit

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public tick As Long

Const maxspeed = 20
Const maxscan = 10

' rStruct is the structure containing all robot
' data. This structure is hidden from the client.

Type rStruct
    x As Single
    y As Single
    deltax As Single    ' x motion per tick
    deltay As Single
    sinfactor As Single ' sin(dir) in radians divided by tick rate
    cosfactor As Single
    status As String * 1
    icon As Integer
    color As Long
    proc As Object
    dir As Single
    dirgoal As Single
    dirdelta As Single  ' +/- increment to turn
    speed As Single
    speedgoal As Single
    shells As Integer
    mheat As Single     ' motor heat
    bheat As Single     ' barrel heat
    reload As Integer   ' ticks till ready
    fire As Integer     ' tick when fired
    newshot As Integer  ' flag for new shell. Cleared by display
    tx As Single
    ty As Single
    dx As Single        ' shell x, y
    dy As Single
    scan As Integer     ' scanning flag - at present,
                        ' set by scan method and cleared
                        ' when display accesses scan data
    sdir As Single
    sres As Integer
    lastscanned As Integer
    health As Integer
End Type

Public Bots(4) As rStruct

' LastBot is not visible. Clients can only
' access it indirectly through class methods
' that call methods in this module. All
' Clients use a single (global) copy.

Public LastBot As Integer       ' Last robot used

Public status As String * 1     ' Server status

Public speed2 As Integer

Private speed3 As Integer

Private Sub kaboom(x As Single, y As Single)

Dim i As Integer
Dim dist As Single
Dim dx As Single
Dim dy As Single
Dim damage As Integer

For i = 1 To LastBot
    dx = Bots(i).x - x
    dy = Bots(i).y - y
    dist = Sqr(dx * dx + dy * dy)
    damage = 0
    If (dist < 40) Then damage = 3
    If (dist < 20) Then damage = 7
    If (dist < 10) Then damage = 12
    If (dist < 6) Then damage = 25
    Call wound(i, damage)
Next i
    
End Sub

Sub Main()

Randomize

Bots(1).color = RGB(255, 0, 0)
Bots(2).color = RGB(0, 255, 0)
Bots(3).color = RGB(0, 0, 255)
Bots(4).color = RGB(255, 0, 255)

End Sub

Public Sub MoveBots()

Dim i As Integer
Dim delta As Single
Dim rate As Single
Dim newx As Single
Dim newy As Single
Dim heat As Single

For i = 1 To LastBot
    If Bots(i).status = "A" Then
        
        ' adjust heading if we're turning
        If Bots(i).dirdelta <> 0 Then
            ' calculate turning rate
            Select Case Int(Bots(i).speed / 25)
                Case 4: rate = 3
                Case 3: rate = 3
                Case 2: rate = 4
                Case 1: rate = 6
                Case 0: rate = 9
            End Select
            delta = Bots(i).dirdelta
            If (delta > rate) Then delta = rate
            If (delta < rate * -1) Then delta = rate * -1
            Bots(i).dir = (Bots(i).dir + delta) Mod 360
            Bots(i).dirdelta = Bots(i).dirdelta - delta
            ' Cos(dir) in radians
            ' Precalculated to avoid doing it every frame
            Bots(i).cosfactor = Cos(Bots(i).dir / 57.3)
            Bots(i).sinfactor = Sin(Bots(i).dir / 57.3)
            Bots(i).deltax = Bots(i).speed * Bots(i).cosfactor / 50
            Bots(i).deltay = Bots(i).speed * Bots(i).sinfactor / 50
        End If
        
        ' barrel cooling
        Bots(i).bheat = Bots(i).bheat - 0.2
        If Bots(i).bheat < 0 Then Bots(i).bheat = 0
        
        ' motor heating
        heat = (Bots(i).speed - 35) / 50
        ' accelerate cooling
        If heat <= 0 Then heat = heat * 3 - 2
        Bots(i).mheat = Bots(i).mheat + heat
        If Bots(i).mheat >= 200 Then
            Bots(i).mheat = 200
            Bots(i).speed = 35
            Bots(i).speedgoal = 35
            Bots(i).deltax = Bots(i).speed * Bots(i).cosfactor / 50
            Bots(i).deltay = Bots(i).speed * Bots(i).sinfactor / 50
        End If
        
        If Bots(i).mheat < 0 Then Bots(i).mheat = 0
        
        ' adjust speed
        delta = Bots(i).speedgoal - Bots(i).speed
        If delta <> 0 Then
            If (delta > 2) Then delta = 2
            If (delta < -2) Then delta = -2
        
            Bots(i).speed = Bots(i).speed + delta
            ' calculate per tick movement
            Bots(i).deltax = Bots(i).speed * Bots(i).cosfactor / 50
            Bots(i).deltay = Bots(i).speed * Bots(i).sinfactor / 50
        End If
        
        ' Move the puppy
        newx = Bots(i).x + Bots(i).deltax
        If newx > 999 Or newx < 0 Then
            Bots(i).speed = 0
            Bots(i).speedgoal = 0
            Bots(i).deltax = 0
            Bots(i).deltay = 0
            Call wound(i, 5)
        Else
            Bots(i).x = newx
        End If
        
        newy = Bots(i).y + Bots(i).deltay
        If newy > 999 Or newy < 0 Then
            Bots(i).speed = 0
            Bots(i).speedgoal = 0
            Bots(i).deltax = 0
            Bots(i).deltay = 0
            Call wound(i, 5)
        Else
            Bots(i).y = newy
        End If
        
        ' Decrement counters
        ' shell reaches target at +2 to allow time for explosion
        
        If Bots(i).fire > 0 Then
            Bots(i).fire = Bots(i).fire - 1
            If Bots(i).fire = 2 Then    'ka-boom
                Call kaboom(Bots(i).tx, Bots(i).ty)
            End If
        End If
        
        If Bots(i).reload > 0 Then Bots(i).reload = Bots(i).reload - 1
    End If  ' end if status = A
Next i

End Sub

Public Function pvtdrive(i As Integer, dir As Integer, speed As Integer) As Integer

Dim delta As Single

Bots(i).dirgoal = dir Mod 360
delta = dir - Bots(i).dir
If delta > 180 Then delta = delta - 360
If delta <= -180 Then delta = delta + 360
Bots(i).dirdelta = delta

Bots(i).speedgoal = speed

delta = Abs(delta)
If delta > 25 And Bots(i).speed > 50 Then
    Bots(i).speedgoal = 0
End If
If delta > 50 And Bots(i).speed > 30 Then
    Bots(i).speedgoal = 0
End If
If delta > 75 And Bots(i).speed > 20 Then
    Bots(i).speedgoal = 0
End If

End Function



Sub PlaceBot(i As Integer)

Dim x As Integer
Dim y As Integer

    Select Case (i)
    Case 1:
        x = 100: y = 100
    Case 2:
        x = 100: y = 600
    Case 3:
        x = 600: y = 100
    Case 4:
        x = 600: y = 600
    End Select

    Bots(i).x = 300 * Rnd() + x
    Bots(i).y = 300 * Rnd() + y
    Bots(i).status = 360 * Rnd()
    Bots(i).speed = 0
    Bots(i).speedgoal = 0
    Bots(i).dir = 360 * Rnd()
    Bots(i).sinfactor = Sin(Bots(i).dir / 57.3)
    Bots(i).cosfactor = Cos(Bots(i).dir / 57.3)
    Bots(i).deltax = 0
    Bots(i).deltay = 0
    Bots(i).dirgoal = Bots(i).dir
    Bots(i).dirdelta = 0
    Bots(i).health = 100
    Bots(i).shells = 4
    Bots(i).status = "A"
    Bots(i).mheat = 0
    Bots(i).bheat = 0
    Bots(i).reload = 0
    Bots(i).fire = 0
    Bots(i).scan = 0
    Server.HealthBar(i - 1) = 100
End Sub

Public Function setindex() As Integer

If LastBot < 4 Then
    LastBot = LastBot + 1
    setindex = LastBot
Else
    setindex = 0
End If

End Function


Sub wound(i As Integer, d As Integer)

' wound robot i by damage d. Check to see if
' battle is over.

Dim x As Integer
Dim living As Integer
Dim winner As Integer

Bots(i).health = Bots(i).health - d

If Bots(i).health <= 0 Then
    Bots(i).health = 0
    Bots(i).speed = 0
    Bots(i).deltax = 0
    Bots(i).deltay = 0
    Bots(i).scan = 0
    Bots(i).speedgoal = 0
    Bots(i).status = "D"
    Server.StatBox(i - 1).Text = "Dead"
End If

living = 0

For x = 1 To LastBot
    If Bots(x).health > 0 Then
        living = living + 1
        winner = x
    End If
Next x

If living = 1 Then
    Server.StatBox(winner - 1) = "WINNER"
    Bots(winner).status = "W"
    Server.PauseBtn_Click
    Server.Text2.Text = Server.Text2.Text + ": Winner is #" + Str(winner)
End If
Server.HealthBar(i - 1) = Bots(i).health
End Sub


