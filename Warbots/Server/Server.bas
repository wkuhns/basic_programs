Attribute VB_Name = "Srvmod"
Option Explicit

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public tick As Long

Public serverbox As ServerForm
Public arena As ArenaForm

Const maxspeed = 20
Const maxscan = 10
Public Const quanta = 50   ' ticks per second

' rStruct is the structure containing all robot
' data. This structure is hidden from the client.

Type rStruct
    x As Single
    y As Single
    name As String
    quadrant As Integer ' starting quadrant
    deltax As Single    ' x,y motion per tick
    deltay As Single
    sinfactor As Single ' sin(dir) in radians divided by tick rate
    cosfactor As Single
    status As String * 1
    icon As Integer     ' not used
    color As Long
    pingnotify As Integer   ' ID of bot who scanned me
    proc As Object      ' pointer to object in user's program space
                        ' containing ping and die methods
    dir As Single
    dirgoal As Single
    dirdelta As Single  ' +/- increment to turn
    speed As Single
    speedgoal As Single
    shells As Integer
    mHeat As Single     ' motor heat
    bHeat As Single     ' barrel heat
    reload As Single    ' ticks till ready
    fire As Single      ' ticks till shell hits
    newshot As Integer  ' flag for new shell. Cleared by display
    tx As Single        ' target x,y
    ty As Single
    dx As Single        ' shell x, y
    dy As Single
    scan As Integer     ' scanning flag - at present,
                        ' set by scan method and cleared
                        ' when display accesses scan data
    sdir As Single
    sres As Integer
    lastscanned As Integer  ' ID of last 'bot who I scanned
    health As Integer
End Type

Type stats
    bdate As Date
    btime As Date
    bfinish As Integer
    nopponents As Integer
    nshots As Integer
    trange As Integer
    nhits As Integer
    tdamage As Integer
    nwounds As Integer
    tinjury As Integer
End Type
    
Private boom(4) As String

Public bang(4) As String

Public Bots(4) As rStruct      ' Moved to Robot Object

Public BotStats(4) As stats

Public LastBot As Integer       ' Last robot used

Public DebugState As Boolean

Public s_status As String * 1     ' Server status

' Quadrant data for intial 'bot placement
Type quad
    used As Integer
    x As Integer
    y As Integer
End Type
Public quads(4) As quad

' Floating point mod operator - %^$##@&** Visual Basic
' Mod returns truncated integer.

Public Function Fmod(x As Single, y As Single) As Single

While x < 0
    x = x + y
Wend

While x > y
    x = x - y
Wend

Fmod = x

End Function

' shell has exploded at x,y.
Private Sub kaboom(Bot As Integer, x As Single, y As Single)

Dim i As Integer
Dim dist As Single
Dim dx As Single
Dim dy As Single
Dim damage As Integer
    
    ' Credit shooter with taking a shot
    BotStats(Bot).nshots = BotStats(Bot).nshots + 1
    
    ' explode ordnance at x,y. Damage everyone nearby.
    For i = 1 To LastBot
        dx = Bots(i).x - x
        dy = Bots(i).y - y
        dist = Sqr(dx * dx + dy * dy)
        damage = 0
        If (dist < 40) Then damage = 3
        If (dist < 20) Then damage = 7
        If (dist < 10) Then damage = 12
        'If (dist < 6) Then damage = 25
        If damage > 0 Then
            ' Update shooter's stats
            BotStats(Bot).nhits = BotStats(Bot).nhits + 1
            BotStats(Bot).tdamage = BotStats(Bot).tdamage + damage
            ' Update victims's stats
            BotStats(i).nwounds = BotStats(i).nwounds + 1
            BotStats(i).tinjury = BotStats(i).tinjury + damage
            Call wound(i, damage)
        End If
    Next i
    
End Sub

Public Sub KillBot(i As Integer)
    
    Dim x As Integer
    Dim living As Integer
    
    living = 0
    
    For x = 1 To LastBot
        If Bots(x).health > 0 Then
            living = living + 1
        End If
    Next x
    
    BotStats(i).bfinish = living + 1
    
    ' Reset everything - he's dead. We don't kill process
    ' yet, though.
    Bots(i).health = 0
    Bots(i).speed = 0
    Bots(i).deltax = 0
    Bots(i).deltay = 0
    Bots(i).scan = 0
    Bots(i).speedgoal = 0
    Bots(i).status = "D"
    Bots(i).x = -1000
    Bots(i).y = 1000
    serverbox.StatBox(i - 1).Text = "Rank: " + Str(living + 1)
    quads(Bots(i).quadrant).used = 0

End Sub

' Move all robots, do all recurrent processing. Eat
' lots of CPU cycles.
'
Public Sub MoveBots()

Dim i As Integer
Dim delta As Single
Dim rate As Single
Dim newx As Single
Dim newy As Single
Dim heat As Single

For i = 1 To LastBot
    ' Don't process dead ones...
    If Bots(i).status = "A" Then
        
        ' adjust heading if we're turning
        If Bots(i).dirdelta <> 0 Then
            ' calculate turning rate
            Select Case Int(Bots(i).speed / 25)
                Case 4: rate = 30 / quanta
                Case 3: rate = 30 / quanta
                Case 2: rate = 40 / quanta
                Case 1: rate = 60 / quanta
                Case 0: rate = 90 / quanta
            End Select
            delta = Bots(i).dirdelta
            If (delta > rate) Then delta = rate
            If (delta < rate * -1) Then delta = rate * -1
            Bots(i).dir = Fmod((Bots(i).dir + delta), 360)
            Bots(i).dirdelta = Bots(i).dirdelta - delta
            ' Cos(dir) in radians
            ' Precalculated to avoid doing it every frame
            Bots(i).cosfactor = Cos(Bots(i).dir / 57.3)
            Bots(i).sinfactor = Sin(Bots(i).dir / 57.3)
            Bots(i).deltax = Bots(i).speed * Bots(i).cosfactor / (5 * quanta)
            Bots(i).deltay = Bots(i).speed * Bots(i).sinfactor / (5 * quanta)
        End If
        
        ' barrel cooling
        Bots(i).bHeat = Bots(i).bHeat - (2 / quanta)
        If Bots(i).bHeat < 0 Then Bots(i).bHeat = 0
        
        ' motor heating
        heat = (Bots(i).speed - 35) / (5 * quanta)
        ' accelerate cooling
        If heat <= 0 Then heat = heat * (30 / quanta) - 2
        Bots(i).mHeat = Bots(i).mHeat + heat
        If Bots(i).mHeat >= 200 Then
            Bots(i).mHeat = 200
            Bots(i).speed = 35
            Bots(i).speedgoal = 35
            Bots(i).deltax = Bots(i).speed * Bots(i).cosfactor / (5 * quanta)
            Bots(i).deltay = Bots(i).speed * Bots(i).sinfactor / (5 * quanta)
        End If
        
        If Bots(i).mHeat < 0 Then Bots(i).mHeat = 0
        
        ' adjust speed
        delta = Bots(i).speedgoal - Bots(i).speed
        If delta <> 0 Then
            If (delta > (20 / quanta)) Then delta = (20 / quanta)
            If (delta < (-20 / quanta)) Then delta = (-20 / quanta)
        
            Bots(i).speed = Bots(i).speed + delta
            ' calculate per tick movement
            Bots(i).deltax = Bots(i).speed * Bots(i).cosfactor / (5 * quanta)
            Bots(i).deltay = Bots(i).speed * Bots(i).sinfactor / (5 * quanta)
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
        
        If Bots(i).reload > 0 Then Bots(i).reload = Bots(i).reload - 1 / quanta
    
    End If  ' end if status = A
        
    ' Process in-flight shells, even if sniper is dead
    ' shell reaches target at +2 to allow time for explosion
    
    If Bots(i).fire > 0 Then
        Bots(i).fire = Bots(i).fire - 1
        If Bots(i).fire = 2 Then    'ka-boom
            Call kaboom(i, Bots(i).tx, Bots(i).ty)
            If ServerForm.SoundBox.Value = 1 Then
                Call EZPlay(boom(i), ssFile)
            End If
            'serverbox.StatBox(i - 1) = ""
        End If
    End If
        
Next i

End Sub
Public Function pvtdrive(i As Integer, dir As Integer, speed As Integer) As Integer
' Process user command to travel in dir at speed

Dim delta As Single

Bots(i).dirgoal = dir Mod 360
delta = dir - Bots(i).dir
If delta > 180 Then delta = delta - 360
If delta <= -180 Then delta = delta + 360

' Are we going too fast? Coast to stop in current dir.
If Abs(delta) > 75 And Bots(i).speed > 20 Then
    speed = 0
    Bots(i).dirgoal = Bots(i).dir
End If
If Abs(delta) > 50 And Bots(i).speed > 30 Then
    speed = 0
    Bots(i).dirgoal = Bots(i).dir
End If
If Abs(delta) > 25 And Bots(i).speed > 50 Then
    speed = 0
    Bots(i).dirgoal = Bots(i).dir
End If

Bots(i).dirdelta = delta
Bots(i).speedgoal = speed
pvtdrive = speed

End Function

Sub PlaceBot(i As Integer)

Dim x As Integer
Dim y As Integer
Dim q As Integer

    ' Place new 'bot somewhere in arena. Set up all required
    ' inital values.

    ' Check that there is an open quad...
    x = 0
    For q = 1 To 4
        If quads(q).used = 0 Then x = 1
    Next q
    
    If x = 0 Then
        MsgBox "Quadrant allocation error. Start over."
        quads(1).used = 0
    End If
    
    ' Pick a random quadrant
    q = Int(Rnd * 4 + 1)
    While quads(q).used = 1
        q = Int(Rnd * 4 + 1)
    Wend
    quads(q).used = 1
    
    Bots(i).name = "Temp"
    Bots(i).quadrant = q
    Bots(i).x = 300 * Rnd() + quads(q).x
    Bots(i).y = 300 * Rnd() + quads(q).y
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
    Bots(i).mHeat = 0
    Bots(i).bHeat = 0
    Bots(i).reload = 0
    Bots(i).fire = 0
    Bots(i).scan = 0
    ServerForm.HealthBar(i - 1) = 100
End Sub

Public Function setindex() As Integer
' Get index for new 'bot

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

' Check if we just died. If we did, see if game is over
If Bots(i).health <= 0 Then
    KillBot (i)
    living = 0
    
    For x = 1 To LastBot
        If Bots(x).health > 0 Then
            living = living + 1
            winner = x
        End If
    Next x
    ' There's only one left, must be the winner
    If living = 1 Then
        serverbox.StatBox(winner - 1) = "WINNER"
        Bots(winner).status = "W"
        serverbox.PauseBtn_Click
        serverbox.Text2.Text = serverbox.Text2.Text + ": Winner is #" + Str(winner)
        BotStats(winner).bfinish = 1
        Call writestats
    End If
End If

serverbox.HealthBar(i - 1) = Bots(i).health
End Sub
Sub writestats()

    Dim i As Integer
    Dim fname As String
    Dim fdata As String
    Dim junk
    
    fname = "p:\programming\warbots\stats.csv"
    Open fname For Append As #1
    
    For i = 1 To LastBot
        fdata = Bots(i).name + "," _
        + Str(BotStats(i).bfinish) + "," _
        + Str(LastBot) + "," _
        + Str(BotStats(i).nshots) + "," _
        + Str(BotStats(i).trange) + "," _
        + Str(BotStats(i).nhits) + "," _
        + Str(BotStats(i).tdamage) + "," _
        + Str(BotStats(i).nwounds) + "," _
        + Str(BotStats(i).tinjury)
        Print #1, fdata
    Next i
    Close #1
    
End Sub
Public Sub rsinitialize()

boom(1) = "P:\programming\Warbots\server\boom1.wav"
boom(2) = "P:\programming\Warbots\server\boom2.wav"
boom(3) = "P:\programming\Warbots\server\boom3.wav"
boom(4) = "P:\programming\Warbots\server\boom4.wav"
bang(1) = "P:\programming\Warbots\server\bang1.wav"
bang(2) = "P:\programming\Warbots\server\bang2.wav"
bang(3) = "P:\programming\Warbots\server\bang3.wav"
bang(4) = "P:\programming\Warbots\server\bang4.wav"

s_status = "P"

Bots(1).color = RGB(255, 0, 0)
Bots(2).color = RGB(0, 255, 0)
Bots(3).color = RGB(0, 0, 255)
Bots(4).color = RGB(255, 0, 255)
Randomize

'arena.Visible = True

End Sub
Public Sub pause()

serverbox.Timer1.Enabled = False
serverbox.Text2 = "Paused"
s_status = "P"

End Sub

Public Sub reset()

Dim i As Integer

serverbox.Timer1.Enabled = False

For i = 1 To LastBot
    ' Clean up structures
    KillBot (i)
    serverbox.StatBox(i - 1).BackColor = serverbox.Text2.BackColor
    serverbox.StatBox(i - 1).Text = "Null"
    ' This next step generates an error, since
    ' it kills the client process, which then fails to
    ' complete the OLE handshake.
    On Error Resume Next
    Call Bots(i).proc.die
Next i

arena.Form_Load

serverbox.Text2 = "Reset, Paused"
 
LastBot = 0

End Sub

Public Sub restart()

Dim i As Integer

serverbox.PauseBtn_Click
arena.Form_Load

For i = 1 To LastBot
    KillBot (i)
Next i

For i = 1 To LastBot
    PlaceBot (i)
    serverbox.StatBox(i - 1).Text = "Ready"
Next i

serverbox.Text2.Text = "Restarted, Paused"

End Sub
Public Sub run()

serverbox.Timer1.Enabled = True
serverbox.Text2 = "Running"
s_status = "R"

End Sub
Public Sub rstick()

On Error Resume Next
MoveBots
On Error Resume Next
arena.DrawFrame

tick = tick + 1

End Sub

Public Sub CleanUpAndDie()

Dim i As Integer

For i = 1 To LastBot
    'On Error Resume Next
    'Call Bots(i).proc.die
    KillBot (i)
    On Error Resume Next
    Call Bots(i).proc.die
Next i

Sleep 1500

End

End Sub

