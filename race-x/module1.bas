Attribute VB_Name = "Module1"
Option Explicit

Dim combuff As String

Enum statusvals
    inactive
    waiting
    ready
    done
    finished
End Enum
    
Type carstruct
    status As statusvals        ' Status: 0=inactive 1=waiting 2=ready 3=done 4=finished
    local As Boolean
    grid As Integer             ' Grid position
    xstart As Integer           ' Calculated from grid position
    ystart As Integer
    xmove As Integer            ' current move coordinates
    ymove As Integer
    x As Integer
    y As Integer
    speed As Integer
    xvel As Integer
    yvel As Integer
    penalty As Integer
    name As String
    color As Long
    Number As Integer
End Type

Type hofstruct
    rank As Integer
    name As String
    score As Single
End Type

Global cars(5) As carstruct

Global hof(10) As hofstruct
Global hofcount

Global LocalPlayer As Integer

Global movecount As Integer
Global MyTurn As Integer

Global trackfile As String
Global bmpfile As String
Global track_setup As Integer

Global roadspeed As Integer
Global wetspeed As Integer
Global icespeed As Integer

Global roadcolor As Long
Global wetcolor As Long
Global icecolor As Long
Global wallcolor As Long

Global Slave As Integer
Global Master As Integer
Global localwait As Integer

Global sl_xs, sl_ys, sl_xf, sl_yf As Integer

Global Const trkfilter = "Track (*.trk)|*.trk"
Global Const bmpfilter = "Bitmap (*.bmp)|*.bmp"

Sub Enroll(pname As String, grid As Integer, plocal As Boolean)
' If plocal, user has identified 'pname' as a local player. If we're networked,
' let everyone know, else just add to list.
' Also may be invoked by incoming network message.

    Dim msg As String
    Dim i As Integer
    
    msg = "Player:" + Str(grid) + " " + pname + "|"
    
    ' We're a networked slave. Let master know.
    If plocal And Slave Then
        RaceForm!Netsocket(0).SendData (msg)
    End If
    
    ' Leave if we're already registered. Set grid position in case
    ' we're being assigned one by the master
    For i = 0 To 4
        If cars(i).status <> inactive And cars(i).name = pname Then
            cars(i).grid = grid
            Exit Sub
        End If
    Next i
    
    ' we've got a new player. If we're a networked slave, our 'grid'
    ' value will be replaced by the master later on.
    
    For i = 0 To 4
        If cars(i).status = inactive Then
            cars(i).status = waiting
            cars(i).name = pname
            cars(i).grid = i
            cars(i).local = plocal
            RaceForm!PlayerBox(i) = cars(i).name + Str(cars(i).local)
            RaceForm!PlayerBox(i).ForeColor = cars(i).color
            Exit For
        End If
    Next i
    
    ' Tell the world
    msg = "Player:" + Str(cars(i).grid) + " " + pname + "|"
    
    If Master Then
        RaceForm.netcast (msg)
    End If
    
End Sub
Sub AddHOFF(player As Integer, dist As Single)
                
    Dim hofentry, hofname As String
    Dim i, j As Integer
    Dim score As Single

'    hofname = InputBox("Enter your name:", "Hall Of Fame")
    hofname = cars(player).name
    i = hofcount
    j = hofcount
    score = movecount - dist
    While (i > 0)
        If (score < hof(i - 1).score) Then
            hof(i).name = hof(i - 1).name
            hof(i).score = hof(i - 1).score
            j = i - 1
        End If
        i = i - 1
    Wend
    hof(j).name = hofname
    hof(j).score = score
    
    If hofcount < 10 Then
        hofcount = hofcount + 1
    End If
    
    HOFForm!HOFBox.Clear
    
    For i = 0 To hofcount - 1
        hofentry = hof(i).name + Space$(30 - Len(hof(i).name)) + Format$(hof(i).score, "###.00")
        HOFForm!HOFBox.AddItem hofentry
    Next i
    HOFForm.Show

    track_save

End Sub

Sub crash(player)
        
        RaceForm!MessageBox.Text = cars(player).name + " crashed."
        cars(player).penalty = 2
        cars(player).xvel = 0
        cars(player).yvel = 0

End Sub

Sub DrawCar(x)
    RaceForm!Track.Line (cars(x).x - 2, cars(x).y - 2)-Step(5, 5), cars(x).color, BF
End Sub

Function GetNextPlayer() As Integer
' Get next local player if any

    Dim i, finished As Integer
        
    finished = True

    For i = 0 To 4
        If cars(i).status = waiting And cars(i).local Then
            LocalPlayer = i
            finished = False
            Exit For
        End If
    Next

    GetNextPlayer = Not finished

End Function

Sub GetSetupData()
    

End Sub


'   This function does all the real work. The vehicle
'   trajectory is calculated and each point is checked
'   for collision. The car is moved and repainted.
'
Sub processMove(player As Integer, x, y As Integer)
    
    Dim pcolor As Long
    Dim dist, odist As Single
    Dim distance As Single
    Dim buff As String
    Dim i, crashed, finished, cycles As Integer
    Dim xs, xf, ys, yf, xv, yv, xnom, ynom, xdiff, ydiff As Single
    Dim xo, yo, oldx, oldy, deltax, deltay As Single
    Dim roadcount, wetcount, icecount As Integer

    ' If we need to send move data to remote, do it
    If (Master Or Slave) And cars(player).local Then
        'sendcom Str$(x)
        'sendcom Str$(y)
    End If
    
    xnom = cars(player).x + cars(player).xvel
    ynom = cars(player).y + cars(player).yvel
    
    ' If click outside circle
    dist = Sqr((x - xnom) * (x - xnom) + (y - ynom) * (y - ynom))
    If dist > cars(player).speed Then
        x = xnom + (x - xnom) * (cars(player).speed / dist)
        y = ynom + (y - ynom) * (cars(player).speed / dist)
    End If
    xv = x - xnom
    yv = y - ynom
    cars(player).xvel = cars(player).xvel + xv
    cars(player).yvel = cars(player).yvel + yv
    
    xs = cars(player).x
    xf = xs + cars(player).xvel
    ys = cars(player).y
    yf = ys + cars(player).yvel
    xdiff = xf - xs
    ydiff = yf - ys
    crashed = 0
    finished = 0
    
    If Int(xdiff) = 0 And Int(ydiff = 0) Then
        Exit Sub
    End If

    oldx = xs
    oldy = ys
    If Abs(xdiff) > Abs(ydiff) Then
        deltax = xdiff / Abs(Int(xdiff))
        deltay = ydiff / Abs(Int(xdiff))
        cycles = Abs(Int(xdiff))
    Else
        deltax = xdiff / Abs(Int(ydiff))
        deltay = ydiff / Abs(Int(ydiff))
        cycles = Abs(Int(ydiff))
    End If

    i = 1
    roadcount = 1
    wetcount = 0
    icecount = 0

    '   Check each point in path
    While (i <= cycles And crashed = 0)
        x = xs + deltax * i
        y = ys + deltay * i

        pcolor = RaceForm!Track.Point(x, y)
        Select Case pcolor
            Case wallcolor
                crash (player)
                crashed = 1
                xf = oldx
                yf = oldy
            Case roadcolor, cars(0).color, cars(1).color, cars(2).color, cars(3).color, cars(4).color
                roadcount = roadcount + 1
            Case wetcolor
                wetcount = wetcount + 1
            Case icecolor
                icecount = icecount + 1
            Case RGB(255, 255, 255)     ' start/finish
                If movecount > 2 Then
                    dist = Sqr(xdiff * xdiff + ydiff * ydiff)
                    xo = xf - x
                    yo = yf - y
                    odist = Sqr(xo * xo + yo * yo)
                    dist = odist / dist
                    finished = 1
                End If
            Case Else
'                Debug.Print pcolor, cars(0).color, cars(1).color, cars(2).color, cars(3).color, cars(4).color
        End Select
        oldx = x
        oldy = y
        i = i + 1
    Wend
    
    If finished = 1 Then
        ScoreBoard.Show
        cars(player).status = finished
        If hofcount < 10 Or movecount - dist < hof(9).score Then
            distance = dist             ' Why do I have to do this?
            Call AddHOFF(player, distance)
        End If
        buff = cars(player).name + Space(20 - Len(cars(player).name)) + Format$(movecount - dist, "##.00")
        ScoreBoard!List1.AddItem buff

    End If

    If crashed = 0 Then
        RaceForm!MessageBox.Text = " "
        oldx = x
        oldy = y
        cars(player).speed = roadcount * roadspeed + wetcount * wetspeed + icecount * icespeed
        cars(player).speed = cars(player).speed / cycles
    End If
    
    RaceForm!Track.Line (xs, ys)-(xf, yf), cars(player).color
    cars(player).x = xf
    cars(player).y = yf
    DrawCar (player)
    
    If cars(player).status = ready Then
        cars(player).status = done
    End If
    
End Sub


Sub SetUpMove()

    Dim x, y As Integer

    ' Draw / redraw start/finish line
    RaceForm!Track.DrawWidth = 2
    RaceForm!Track.Line (sl_xs, sl_ys)-(sl_xf, sl_yf), RGB(255, 255, 255)
    RaceForm!Track.DrawWidth = 1
    
    RaceForm!Track.AutoRedraw = False
    RaceForm!Track.Cls
    
    RaceForm!NameBox.Text = cars(LocalPlayer).name
    RaceForm!NameBox.ForeColor = cars(LocalPlayer).color
    
    ' Draw circle for move
    x = cars(LocalPlayer).x + cars(LocalPlayer).xvel
    y = cars(LocalPlayer).y + cars(LocalPlayer).yvel
    RaceForm!Track.Circle (x, y), cars(LocalPlayer).speed, RGB(255, 255, 255)
    RaceForm!Track.AutoRedraw = True
    
End Sub

Sub StartGame()
    
    Dim i As Integer
    Dim dx, dy As Single

    RaceForm!Track.Cls
    
    dx = (sl_xf - sl_xs) / 6
    dy = (sl_yf - sl_ys) / 6

    For i = 0 To 4
        If cars(i).status <> inactive Then
            cars(i).xstart = sl_xs + dx * (cars(i).grid + 1)
            cars(i).ystart = sl_ys + dy * (cars(i).grid + 1)
            cars(i).speed = roadspeed
            cars(i).xvel = 0
            cars(i).yvel = 0
            cars(i).x = cars(i).xstart
            cars(i).y = cars(i).ystart
            DrawCar (i)
        End If
    Next i
    
    ScoreBoard!List1.Clear
    ScoreBoard.Hide
    
    movecount = 0
    
    GetNextPlayer
    SetUpMove

End Sub

Sub Track_load()
    
    Dim i As Integer
    Dim hofentry As String
    Dim dx, dy As Single

    On Error Resume Next
    Open trackfile For Input As #1
    
    If Err > 0 Then
        On Error GoTo 0
        RaceForm!TrackDialog.Filter = trkfilter
        RaceForm!TrackDialog.Action = 1
        trackfile = RaceForm!TrackDialog.filename
        Open trackfile For Input As #1
    End If
    On Error GoTo 0
    
    Input #1, bmpfile
    Input #1, sl_xs, sl_ys, sl_xf, sl_yf
    Input #1, roadcolor, roadspeed
    Input #1, wetcolor, wetspeed
    Input #1, icecolor, icespeed
    Input #1, wallcolor
    Input #1, hofcount
    
    HOFForm!HOFBox.Clear
    For i = 0 To hofcount - 1
        Input #1, hof(i).name, hof(i).score
        hofentry = hof(i).name + Space$(30 - Len(hof(i).name)) + Format$(hof(i).score, "###.00")
        HOFForm!HOFBox.AddItem hofentry
    Next i

    Close #1

    RaceForm!Track.Picture = LoadPicture(bmpfile)
    
End Sub

Sub track_save()
    
    Dim i As Integer
    
    Open trackfile For Output As #1
    Print #1, """"; bmpfile; """"
    Print #1, sl_xs, sl_ys, sl_xf, sl_yf
    Print #1, roadcolor, roadspeed
    Print #1, wetcolor, wetspeed
    Print #1, icecolor, icespeed
    Print #1, wallcolor
    Print #1, hofcount
    For i = 0 To hofcount - 1
        Print #1, """"; hof(i).name; """"; hof(i).score
    Next i
    Close #1


End Sub

Sub TrackClick(x, y As Integer)
    
    Dim msg As String
    
    If cars(LocalPlayer).status = waiting Then
        Debug.Print "Local Move:", x, y
        cars(LocalPlayer).xmove = x
        cars(LocalPlayer).ymove = y
        cars(LocalPlayer).status = ready
        If Master Or Slave Then
            msg = "Move:" + Str(cars(LocalPlayer).grid) + "," + Str(x) + "," + Str(y) + "|"
            MsgBox msg
            RaceForm.netcast msg
        End If
        processMove LocalPlayer, x, y
        If GetNextPlayer() Then       ' if another local move...
            SetUpMove
        End If
    End If
    
End Sub

