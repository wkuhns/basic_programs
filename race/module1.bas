Option Explicit

Dim combuff As String

Type carstruct
    active As Integer
    local As Integer
    xstart As Integer
    ystart As Integer
    X As Integer
    Y As Integer
    speed As Integer
    xvel As Integer
    yvel As Integer
    penalty As Integer
    name As String
    color As Long
    number As Integer
End Type

Type hofstruct
    name As String
    score As Single
End Type

Global cars(5) As carstruct
Global actives(5) As Integer

Global hof(10) As hofstruct
Global hofcount

Global player As Integer

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

Global remotestatus As Integer
Global master As Integer
Global localwait As Integer

Global CRLF As String

Global sl_xs, sl_ys, sl_xf, sl_yf As Integer

Global Const trkfilter = "Track (*.trk)|*.trk"
Global Const bmpfilter = "Bitmap (*.bmp)|*.bmp"

Sub AddHOFF (dist As Single)
		
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

Sub crash ()
	
	RaceForm!MessageBox.Text = cars(player).name + " crashed."
	cars(player).penalty = 2
	cars(player).xvel = 0
	cars(player).yvel = 0

End Sub

Sub DrawCar (X)
    RaceForm!Track.Line (cars(X).X - 2, cars(X).Y - 2)-Step(5, 5), cars(X).color, BF
End Sub

Function getcom () As String

    If RaceForm!Com2.InBufferCount > 0 Then
	combuff = combuff + RaceForm!Com2.Input
    End If
    
    While InStr(combuff, Chr$(10)) = 0
	DoEvents
	If RaceForm!Com2.InBufferCount > 0 Then
	    combuff = combuff + RaceForm!Com2.Input
	End If
    Wend

    getcom = Left$(combuff, InStr(combuff, Chr$(13)) - 1)
    combuff = Right$(combuff, Len(combuff) - InStr(combuff, Chr$(10)))

End Function

Function GetNextPlayer () As Integer
	
    Dim i, finished As Integer
	
    finished = True

    For i = 0 To 4
	If cars(i).active Then
	    finished = False
	End If
    Next

newplayer:

    If Not finished Then
	player = player + 1
	If player = 5 Then                  ' back to 0
	    movecount = movecount + 1
	    player = 0
	End If
      
	' skip inactive players
	If cars(player).active = 0 Then
	    GoTo newplayer
	End If

	If cars(player).penalty > 0 Then
	    cars(player).penalty = cars(player).penalty - 1
	    GoTo newplayer
	End If
    
    End If

    RaceForm!Text1.Text = Str$(movecount)
    
    GetNextPlayer = Not finished

End Function

Sub GetRemoteMove ()

    Dim k As String
    Dim X, Y As Single


    X = Val(getcom())
    Y = Val(getcom())
    Debug.Print "Remote Move: ", X, Y
    Call processMove(Int(X), Int(Y))
    If GetNextPlayer() Then       ' if game not over
	SetUpMove
'        If Not cars(player).local Then
'            GetRemoteMove
'        End If
    Else
	RaceForm!Text1.Text = "Race over!"
    End If

End Sub

Sub GetSetupData ()
    
    Dim buff As String
    Dim i As Integer
    Dim j As Integer

    trackfile = getcom()
    Track_load
    
    For i = 0 To 4
	buff = getcom()
	cars(i).name = buff
	buff = getcom()
	cars(i).active = Val(buff)
	buff = getcom()
	cars(i).local = Val(buff)
    Next i
    
'    Debug.Print buff

End Sub

Sub OpenCom ()

    
    RaceForm!Com2.CommPort = 2
    RaceForm!Com2.InputLen = 0
    RaceForm!Com2.PortOpen = True

    sendcom "READY"
    If getcom() <> "READY" Then
	While getcom() <> "READY"
	    DoEvents
	Wend
    End If


End Sub

'   This function does all the real work. The vehicle
'   trajectory is calculated and each point is checked
'   for collision. The car is moved and repainted.
'
Sub processMove (X, Y As Integer)
    
    Dim pcolor As Long
    Dim dist, odist As Single
    Dim distance As Single
    Dim buff As String
    Dim i, crashed, finished, cycles As Integer
    Dim xs, xf, ys, yf, xv, yv, xnom, ynom, xdiff, ydiff As Single
    Dim xo, yo, oldx, oldy, deltax, deltay As Single
    Dim roadcount, wetcount, icecount As Integer

    ' If we need to send move data to remote, do it
    If remotestatus <> 0 And cars(player).local Then
	sendcom Str$(X)
	sendcom Str$(Y)
    End If
    
    xnom = cars(player).X + cars(player).xvel
    ynom = cars(player).Y + cars(player).yvel
    
    ' If click outside circle
    dist = Sqr((X - xnom) * (X - xnom) + (Y - ynom) * (Y - ynom))
    If dist > cars(player).speed Then
	X = xnom + (X - xnom) * (cars(player).speed / dist)
	Y = ynom + (Y - ynom) * (cars(player).speed / dist)
    End If
    xv = X - xnom
    yv = Y - ynom
    cars(player).xvel = cars(player).xvel + xv
    cars(player).yvel = cars(player).yvel + yv
    
    xs = cars(player).X
    xf = xs + cars(player).xvel
    ys = cars(player).Y
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
	X = xs + deltax * i
	Y = ys + deltay * i

	pcolor = RaceForm!Track.Point(X, Y)
	Select Case pcolor
	    Case wallcolor
		crash
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
		    xo = xf - X
		    yo = yf - Y
		    odist = Sqr(xo * xo + yo * yo)
		    dist = odist / dist
		    finished = 1
		End If
	    Case Else
'                Debug.Print pcolor, cars(0).color, cars(1).color, cars(2).color, cars(3).color, cars(4).color
	End Select
	oldx = X
	oldy = Y
	i = i + 1
    Wend
    
    If finished = 1 Then
	scoreboard.Show
	cars(player).active = False
	If hofcount < 10 Or movecount - dist < hof(9).score Then
	    distance = dist             ' Why do I have to do this?
	    Call AddHOFF(distance)
	End If
	buff = cars(player).name + Space(20 - Len(cars(player).name)) + Format$(movecount - dist, "##.00")
	scoreboard!List1.AddItem buff

    End If

    If crashed = 0 Then
	RaceForm!MessageBox.Text = " "
	oldx = X
	oldy = Y
	cars(player).speed = roadcount * roadspeed + wetcount * wetspeed + icecount * icespeed
	cars(player).speed = cars(player).speed / cycles
    End If
    
    RaceForm!Track.Line (xs, ys)-(xf, yf), cars(player).color
    cars(player).X = xf
    cars(player).Y = yf
    DrawCar (player)
    
End Sub

Sub sendcom (buffer As String)

    RaceForm!Com2.InputLen = 0
    RaceForm!Com2.Output = buffer + Chr$(13) + Chr$(10)

End Sub

Sub SendSetupData ()

    Dim buff As String
    Dim i As Integer

    sendcom trackfile
    For i = 0 To 4
	sendcom cars(i).name
	sendcom Str$(cars(i).active)
	sendcom Str$(Not cars(i).local)
    Next i

End Sub

Sub SetUpMove ()

    Dim X, Y As Integer

    ' Draw / redraw start/finish line
    RaceForm!Track.DrawWidth = 2
    RaceForm!Track.Line (sl_xs, sl_ys)-(sl_xf, sl_yf), RGB(255, 255, 255)
    RaceForm!Track.DrawWidth = 1
    
    RaceForm!Track.AutoRedraw = False
    RaceForm!Track.Cls
    
    RaceForm!NameBox.Text = cars(player).name
    RaceForm!NameBox.ForeColor = cars(player).color
    
    ' Draw circle for move, or wait for remote
    If cars(player).local Then
	X = cars(player).X + cars(player).xvel
	Y = cars(player).Y + cars(player).yvel
	RaceForm!Track.Circle (X, Y), cars(player).speed, RGB(255, 255, 255)
	RaceForm!Track.AutoRedraw = True
    Else
	RaceForm!Track.AutoRedraw = True
	RaceForm!MessageBox.Text = "Waiting for remote player"
	GetRemoteMove
    End If

    
End Sub

Sub StartGame ()
    
    Dim i As Integer
    Dim dx, dy As Single

    If remotestatus = True Then
	If master = True Then
	    SendSetupData
	Else
	    GetSetupData
	End If
    End If

    RaceForm!Track.Cls
    
    dx = (sl_xf - sl_xs) / 6
    dy = (sl_yf - sl_ys) / 6

    For i = 0 To 4
	cars(i).xstart = sl_xs + dx * (i + 1)
	cars(i).ystart = sl_ys + dy * (i + 1)
	cars(i).speed = roadspeed
	cars(i).active = actives(i)
    Next i
    
    scoreboard!List1.Clear
    scoreboard.Hide

    player = 4              ' roll over to 0+
    i = GetNextPlayer()
    movecount = 0

    For i = 0 To 4
	If cars(i).active Then
	    cars(i).xvel = 0
	    cars(i).yvel = 0
	    cars(i).X = cars(i).xstart
	    cars(i).Y = cars(i).ystart
	    DrawCar (i)
	End If
    Next i
    
    SetUpMove

End Sub

Sub Track_load ()
    
    Dim i As Integer
    Dim hofentry As String
    Dim dx, dy As Single

    On Error Resume Next
    Open trackfile For Input As #1
    
    If Err > 0 Then
	On Error GoTo 0
	RaceForm!TrackDialog.Filter = trkfilter
	RaceForm!TrackDialog.Action = 1
	trackfile = RaceForm!TrackDialog.Filename
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

Sub track_save ()
    
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

Sub TrackClick (X, Y As Integer)
    
    If cars(player).local Then
	Debug.Print "Local Move:", X, Y
	processMove X, Y
	If GetNextPlayer() Then       ' if game not over
	    SetUpMove
	    If cars(player).local = False Then
		GetRemoteMove
	    End If
	Else
	    RaceForm!Text1.Text = "Race over!"
	End If
    End If
	
End Sub

